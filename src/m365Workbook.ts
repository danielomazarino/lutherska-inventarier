import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
} from '@azure/msal-browser'
import { assetTagError, normalizeAssetTag, normalizeCategoryPrefix } from './inventoryIntegrity'

export type Option = { id: string; name: string; color?: string; prefix?: string }
export type Item = {
  id: string
  assetTag: string
  name: string
  categoryId: string
  primaryGroupId: string
  secondaryGroupIds: string[]
  locationId: string
  quantity: number
  notes: string
}
export type Loan = {
  id: string
  itemId: string
  borrower: string
  borrowerGroupId: string
  recordedBy: string
  lentAt: string
  dueAt: string
  returnedAt?: string
}
export type InventoryData = {
  categories: Option[]
  groups: Option[]
  locations: Option[]
  items: Item[]
  loans: Loan[]
}
export type M365Config = {
  tenantId: string
  clientId: string
  workbookUrl: string
}

type TableName = 'Categories' | 'Groups' | 'Locations' | 'Inventory' | 'Loans'
type GraphRows = { value: { index: number; values: unknown[][] }[] }
type DriveItem = { id: string; name: string; parentReference: { driveId: string } }

const scopes = ['Files.ReadWrite']
const graphRoot = 'https://graph.microsoft.com/v1.0'
const retryableStatuses = new Set([429, 503, 504])

const encodeSharingUrl = (url: string) => {
  const encoded = btoa(unescape(encodeURIComponent(url)))
    .replace(/=+$/, '')
    .replace(/\//g, '_')
    .replace(/\+/g, '-')
  return `u!${encoded}`
}

const text = (value: unknown) => String(value ?? '')

export class M365Workbook {
  private readonly auth: PublicClientApplication
  private account?: AccountInfo
  private driveId = ''
  private itemId = ''
  private sessionId = ''
  private rowIndexes: Record<TableName, Map<string, number>> = {
    Categories: new Map(),
    Groups: new Map(),
    Locations: new Map(),
    Inventory: new Map(),
    Loans: new Map(),
  }

  constructor(private readonly config: M365Config) {
    this.auth = new PublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        redirectUri: new URL(import.meta.env.BASE_URL, window.location.href).href,
      },
      cache: { cacheLocation: 'localStorage' },
    })
  }

  get userName() {
    return this.account?.name ?? this.account?.username ?? ''
  }

  async connect(interactive = true) {
    await this.auth.initialize()
    this.account = this.auth.getAllAccounts()[0]
    if (!this.account && interactive) {
      const login = await this.auth.loginPopup({ scopes, prompt: 'select_account' })
      this.account = login.account
    }
    if (!this.account) return false

    const shareId = encodeSharingUrl(this.config.workbookUrl)
    const item = await this.request<DriveItem>(`/shares/${shareId}/driveItem?$select=id,name,parentReference`)
    this.driveId = item.parentReference.driveId
    this.itemId = item.id
    await this.createSession()
    return true
  }

  async disconnect() {
    if (this.account) await this.auth.logoutPopup({ account: this.account })
  }

  async load(): Promise<InventoryData> {
    const [categories, groups, locations, items, loans] = await Promise.all([
      this.readTable('Categories'),
      this.readTable('Groups'),
      this.readTable('Locations'),
      this.readTable('Inventory'),
      this.readTable('Loans'),
    ])
    return {
      categories: categories.map((row) => ({ id: text(row[0]), name: text(row[1]), color: text(row[2]) || undefined, prefix: text(row[3]) || undefined })),
      groups: groups.map((row) => ({ id: text(row[0]), name: text(row[1]) })),
      locations: locations.map((row) => ({ id: text(row[0]), name: text(row[1]) })),
      items: items.map((row) => ({
        id: text(row[0]),
        assetTag: text(row[1]),
        name: text(row[2]),
        categoryId: text(row[3]),
        primaryGroupId: text(row[4]),
        secondaryGroupIds: text(row[5]).split(';').filter(Boolean),
        locationId: text(row[6]),
        quantity: Number(row[7]) || 1,
        notes: text(row[8]),
      })),
      loans: loans.map((row) => ({
        id: text(row[0]),
        itemId: text(row[1]),
        borrower: text(row[2]),
        borrowerGroupId: text(row[3]),
        recordedBy: text(row[4]),
        lentAt: text(row[5]),
        dueAt: text(row[6]),
        returnedAt: text(row[7]) || undefined,
      })),
    }
  }

  async addItem(item: Item) {
    await this.assertAssetTagAvailable(item)
    await this.addRow('Inventory', [item.id, item.assetTag, item.name, item.categoryId, item.primaryGroupId, item.secondaryGroupIds.join(';'), item.locationId, item.quantity, item.notes])
    const rows = await this.readTable('Inventory')
    const matchingIds = rows.filter((row) => normalizeAssetTag(text(row[1])) === item.assetTag).map((row) => text(row[0]))
    if (matchingIds.length > 1) {
      const currentIndex = this.rowIndexes.Inventory.get(item.id)
      const earliestIndex = Math.min(...matchingIds.map((id) => this.rowIndexes.Inventory.get(id) ?? Number.MAX_SAFE_INTEGER))
      if (currentIndex !== undefined && currentIndex > earliestIndex) {
        await this.deleteRow('Inventory', item.id)
        throw new Error(`Inventarienumret ${item.assetTag} hann registreras av en annan användare. Din dubblettrad sparades inte.`)
      }
    }
  }

  async updateItem(item: Item) {
    await this.assertAssetTagAvailable(item)
    return this.updateRow('Inventory', item.id, [item.id, item.assetTag, item.name, item.categoryId, item.primaryGroupId, item.secondaryGroupIds.join(';'), item.locationId, item.quantity, item.notes])
  }

  async deleteItem(itemId: string) {
    const loanRows = await this.readTable('Loans')
    if (loanRows.some((row) => text(row[1]) === itemId)) throw new Error('Föremålet kan inte tas bort eftersom det finns i lånehistoriken.')
    await this.readTable('Inventory')
    return this.deleteRow('Inventory', itemId)
  }

  addLoan(loan: Loan) {
    return this.addRow('Loans', [loan.id, loan.itemId, loan.borrower, loan.borrowerGroupId, loan.recordedBy, loan.lentAt, loan.dueAt, loan.returnedAt ?? ''])
  }

  updateLoan(loan: Loan) {
    return this.updateRow('Loans', loan.id, [loan.id, loan.itemId, loan.borrower, loan.borrowerGroupId, loan.recordedBy, loan.lentAt, loan.dueAt, loan.returnedAt ?? ''])
  }

  returnLoan(loan: Loan) {
    return this.updateLoan(loan)
  }

  async addOption(kind: 'categories' | 'groups' | 'locations', option: Option) {
    const table = this.optionTable(kind)
    if (table === 'Categories') await this.assertCategoryPrefixAvailable(option)
    return this.addRow(table, table === 'Categories' ? [option.id, option.name, option.color ?? '', option.prefix ?? ''] : [option.id, option.name])
  }

  async renameOption(kind: 'categories' | 'groups' | 'locations', option: Option) {
    const table = this.optionTable(kind)
    if (table === 'Categories') await this.assertCategoryPrefixAvailable(option)
    return this.updateRow(table, option.id, table === 'Categories' ? [option.id, option.name, option.color ?? '', option.prefix ?? ''] : [option.id, option.name])
  }

  private optionTable(kind: 'categories' | 'groups' | 'locations'): TableName {
    return kind === 'categories' ? 'Categories' : kind === 'groups' ? 'Groups' : 'Locations'
  }

  private async assertAssetTagAvailable(item: Item) {
    const rows = await this.readTable('Inventory')
    const categoryRows = await this.readTable('Categories')
    const items = rows.map((row) => ({ id: text(row[0]), assetTag: text(row[1]), categoryId: text(row[3]) }))
    const categories = categoryRows.map((row) => ({ id: text(row[0]), name: text(row[1]), prefix: text(row[3]) || undefined }))
    const error = assetTagError(item.assetTag, items, categories, item.id, item.categoryId)
    if (error) throw new Error(error)
    item.assetTag = normalizeAssetTag(item.assetTag)
  }

  private async assertCategoryPrefixAvailable(option: Option) {
    const prefix = normalizeCategoryPrefix(option.prefix ?? '')
    if (!/^[A-Z]{3}$/.test(prefix)) throw new Error('Prefixet ska bestå av exakt tre bokstäver.')
    const rows = await this.readTable('Categories')
    const duplicate = rows.find((row) => text(row[0]) !== option.id && normalizeCategoryPrefix(text(row[3])) === prefix)
    if (duplicate) throw new Error(`Prefixet ${prefix} används redan av kategorin ${text(duplicate[1])}.`)
    option.prefix = prefix
  }

  private async createSession() {
    try {
      const session = await this.request<{ id: string }>('/workbook/createSession', {
        method: 'POST',
        body: JSON.stringify({ persistChanges: true }),
      })
      this.sessionId = session.id
    } catch (error) {
      console.warn('Workbook session unavailable; Graph calls will still persist.', error)
    }
  }

  private async readTable(name: TableName) {
    const result = await this.request<GraphRows>(`/workbook/tables/${name}/rows?$top=5000`)
    const indexes = new Map<string, number>()
    const rows = result.value.map((entry) => {
      const values = entry.values[0] ?? []
      indexes.set(text(values[0]), entry.index)
      return values
    })
    this.rowIndexes[name] = indexes
    return rows
  }

  private async addRow(name: TableName, values: unknown[]) {
    const row = await this.request<{ index: number }>(`/workbook/tables/${name}/rows/add`, {
      method: 'POST',
      body: JSON.stringify({ values: [values] }),
    })
    const id = text(values[0])
    if (id) this.rowIndexes[name].set(id, row.index)
  }

  private async updateRow(name: TableName, id: string, values: unknown[]) {
    const index = this.rowIndexes[name].get(id)
    if (index === undefined) throw new Error(`${name} row ${id} was not found. Refresh and try again.`)
    await this.request(`/workbook/tables/${name}/rows/itemAt(index=${index})`, {
      method: 'PATCH',
      body: JSON.stringify({ values: [values] }),
    })
  }

  private async deleteRow(name: TableName, id: string) {
    const index = this.rowIndexes[name].get(id)
    if (index === undefined) return
    await this.request(`/workbook/tables/${name}/rows/itemAt(index=${index})`, { method: 'DELETE' })
    await this.readTable(name)
  }

  private async token() {
    if (!this.account) throw new Error('Sign in to Microsoft 365 first.')
    try {
      return (await this.auth.acquireTokenSilent({ account: this.account, scopes })).accessToken
    } catch (error) {
      if (!(error instanceof InteractionRequiredAuthError)) throw error
      return (await this.auth.acquireTokenPopup({ account: this.account, scopes })).accessToken
    }
  }

  private async request<T = void>(path: string, init: RequestInit = {}, attempt = 0): Promise<T> {
    const workbookPath = this.driveId && this.itemId
      ? `/drives/${encodeURIComponent(this.driveId)}/items/${encodeURIComponent(this.itemId)}`
      : ''
    const isShareLookup = path.startsWith('/shares/')
    const url = `${graphRoot}${isShareLookup ? path : `${workbookPath}${path}`}`
    const response = await fetch(url, {
      ...init,
      headers: {
        Authorization: `Bearer ${await this.token()}`,
        'Content-Type': 'application/json',
        ...(this.sessionId && !isShareLookup ? { 'workbook-session-id': this.sessionId } : {}),
        ...init.headers,
      },
    })
    if (retryableStatuses.has(response.status) && attempt < 3) {
      const retryAfter = Number(response.headers.get('Retry-After')) || 2 ** attempt
      await new Promise((resolve) => setTimeout(resolve, retryAfter * 1000))
      return this.request<T>(path, init, attempt + 1)
    }
    if (!response.ok) {
      const detail = await response.text()
      throw new Error(`Microsoft Graph ${response.status}: ${detail || response.statusText}`)
    }
    return response.status === 204 ? (undefined as T) : response.json() as Promise<T>
  }
}