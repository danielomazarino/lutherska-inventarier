import { FormEvent, useEffect, useMemo, useRef, useState } from 'react'
import {
  Boxes,
  BookOpen,
  Check,
  ChevronDown,
  CircleAlert,
  Cloud,
  CloudOff,
  ClipboardList,
  Handshake,
  LayoutDashboard,
  LogIn,
  MapPin,
  Menu,
  PackagePlus,
  Plus,
  RefreshCw,
  RotateCcw,
  Search,
  Settings,
  SlidersHorizontal,
  Tag,
  Trash2,
  Users,
  X,
} from 'lucide-react'
import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import './App.css'
import readme from '../README.md?raw'
import { assetTagError, BUILT_IN_CATEGORY_PREFIXES, categoryPrefixFor, inventoryIntegrityIssues, migrateSampleInventory, normalizeAssetTag, normalizeCategoryPrefix } from './inventoryIntegrity'
import {
  M365Workbook,
  type InventoryData,
  type Item,
  type Loan,
  type M365Config,
  type Option,
} from './m365Workbook'

type View = 'dashboard' | 'inventory' | 'loans' | 'administration' | 'documentation' | 'setup'

const STORAGE_KEY = 'church-inventory-v4'
const LEGACY_STORAGE_KEY = 'church-inventory-v3'
const M365_CONFIG_KEY = 'church-inventory-m365-config'
const deployedM365Config: M365Config = {
  tenantId: import.meta.env.VITE_M365_TENANT_ID ?? '',
  clientId: import.meta.env.VITE_M365_CLIENT_ID ?? '',
  workbookUrl: import.meta.env.VITE_M365_WORKBOOK_URL ?? '',
}
const palette = ['#386641', '#bc6c25', '#3d5a80', '#8f5d5d', '#5b4b8a', '#277da1']
const isoDate = (daysFromToday = 0) => {
  const date = new Date()
  date.setDate(date.getDate() + daysFromToday)
  return date.toISOString().slice(0, 10)
}

const seedData: InventoryData = {
  categories: [
    { id: 'cat-furniture', name: 'Möbler', color: '#6f8657', prefix: 'MOB' },
    { id: 'cat-lighting', name: 'Belysning', color: '#f59b45', prefix: 'BEL' },
    { id: 'cat-sound', name: 'Ljudutrustning', color: '#3688df', prefix: 'LJU' },
    { id: 'cat-kitchen', name: 'Kök & fika', color: '#c73f3b', prefix: 'KOK' },
  ],
  groups: [
    { id: 'grp-board-mission', name: 'Styrelsen Lutherska missionsföreningen' },
    { id: 'grp-board-hagaparken', name: 'Styrelsen Hagaparken Ekonomisk förening' },
    { id: 'grp-premises-change', name: 'Arbetsgrupp för lokalförändringar' },
    { id: 'grp-sound', name: 'Ljudgruppen' },
    { id: 'grp-service-visuals', name: 'Bildvisning i gudstjänst' },
    { id: 'grp-kitchen-cleaning-purchases', name: 'Inköp för kök och städ' },
    { id: 'grp-allergy-purchases', name: 'Allergiinköp' },
    { id: 'grp-household', name: 'Husmor/far' },
    { id: 'grp-holiday-decoration', name: 'Utsmyckning storhelger' },
    { id: 'grp-music-committee', name: 'Musikutskottet' },
    { id: 'grp-service-committee', name: 'Gudstjänstutskottet' },
    { id: 'grp-children-youth-committee', name: 'Barn- och ungdomsutskottet' },
    { id: 'grp-orchestra', name: 'Lutherska Missionskyrkans orkester' },
    { id: 'grp-choir', name: 'Lutherska Missionskyrkans kör' },
    { id: 'grp-choir-council', name: 'Lutherska Missionskyrkans körråd' },
    { id: 'grp-sunday-school', name: 'Söndagsskolan' },
    { id: 'grp-forest-school', name: 'Skogsskolan' },
    { id: 'grp-baby-rhythmics', name: 'Babyrytmik' },
    { id: 'grp-clap-and-sound', name: 'Klapp och klang' },
    { id: 'grp-popkidz', name: 'Popkidz' },
    { id: 'grp-tweenies', name: 'Tweenies' },
    { id: 'grp-lux', name: 'LUX' },
    { id: 'grp-gamla-barn', name: 'Gamla Barn' },
  ],
  locations: [
    { id: 'loc-sanctuary', name: 'Kyrksalen' },
    { id: 'loc-stage', name: 'Scenförrådet' },
    { id: 'loc-kitchen', name: 'Köksförrådet' },
    { id: 'loc-basement', name: 'Källarförrådet' },
  ],
  items: [
    { id: 'item-1', assetTag: 'LJU-001', name: 'Trådlösa mikrofoner', categoryId: 'cat-sound', primaryGroupId: 'grp-sound', secondaryGroupIds: ['grp-lux'], locationId: 'loc-stage', quantity: 2, notes: 'Två handmikrofoner och mottagare.' },
    { id: 'item-2', assetTag: 'MOB-014', name: 'Fällbord', categoryId: 'cat-furniture', primaryGroupId: 'grp-premises-change', secondaryGroupIds: ['grp-household'], locationId: 'loc-basement', quantity: 8, notes: '' },
    { id: 'item-3', assetTag: 'KOK-008', name: 'Stor kaffebryggare', categoryId: 'cat-kitchen', primaryGroupId: 'grp-household', secondaryGroupIds: ['grp-kitchen-cleaning-purchases'], locationId: 'loc-kitchen', quantity: 1, notes: 'Rymmer 10 liter.' },
    { id: 'item-4', assetTag: 'BEL-004', name: 'Portabla LED-lampor', categoryId: 'cat-lighting', primaryGroupId: 'grp-service-visuals', secondaryGroupIds: ['grp-lux'], locationId: 'loc-stage', quantity: 6, notes: 'Förvaras i vadderad transportväska.' },
  ],
  loans: [
    { id: 'loan-1', itemId: 'item-4', borrower: 'Anna Lind', borrowerGroupId: 'grp-lux', recordedBy: 'Erik Nilsson', lentAt: isoDate(-6), dueAt: isoDate(-1) },
    { id: 'loan-2', itemId: 'item-2', borrower: 'Jonas Berg', borrowerGroupId: 'grp-household', recordedBy: 'Sara Holm', lentAt: isoDate(-1), dueAt: isoDate(4) },
  ],
}

const newId = (prefix: string) => `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`

const nextInventorySequence = (data: InventoryData, categoryId: string, currentItemId?: string) => {
  const prefix = categoryPrefixFor(categoryId, data.categories)
  const usedNumbers = data.items
    .filter((item) => item.id !== currentItemId && normalizeAssetTag(item.assetTag).startsWith(`${prefix}-`))
    .map((item) => Number(normalizeAssetTag(item.assetTag).slice(-3)))
    .filter((number) => Number.isInteger(number) && number >= 1 && number <= 999)
  const highestNumber = Math.max(0, ...usedNumbers)
  if (highestNumber < 999) return String(highestNumber + 1).padStart(3, '0')
  const used = new Set(usedNumbers)
  for (let number = 1; number <= 999; number += 1) {
    if (!used.has(number)) return String(number).padStart(3, '0')
  }
  return ''
}

function App() {
  const [data, setData] = useState<InventoryData>(() => {
    const stored = localStorage.getItem(STORAGE_KEY) ?? localStorage.getItem(LEGACY_STORAGE_KEY)
    const migrated = migrateSampleInventory(stored ? JSON.parse(stored) as InventoryData : seedData)
    localStorage.setItem(STORAGE_KEY, JSON.stringify(migrated))
    return migrated
  })
  const [view, setView] = useState<View>('dashboard')
  const [search, setSearch] = useState('')
  const [categoryFilter, setCategoryFilter] = useState('all')
  const [modal, setModal] = useState<'item' | 'edit-item' | 'loan' | 'edit-loan' | null>(null)
  const [selectedItemId, setSelectedItemId] = useState('')
  const [selectedLoanId, setSelectedLoanId] = useState('')
  const [formError, setFormError] = useState('')
  const [mobileNavOpen, setMobileNavOpen] = useState(false)
  const edgeSwipeStart = useRef<{ x: number; y: number }>()
  const [m365Config, setM365Config] = useState<M365Config>(() => {
    const stored = localStorage.getItem(M365_CONFIG_KEY)
    return stored ? JSON.parse(stored) as M365Config : deployedM365Config
  })
  const [workbook, setWorkbook] = useState<M365Workbook>()
  const [syncStatus, setSyncStatus] = useState<'sample' | 'connecting' | 'synced' | 'saving' | 'error'>('sample')
  const [syncError, setSyncError] = useState('')
  const [connectedUser, setConnectedUser] = useState('')

  useEffect(() => {
    if ('scrollRestoration' in history) history.scrollRestoration = 'manual'

    const resetMobileScroll = () => {
      if (!window.matchMedia('(max-width: 900px)').matches) return
      window.requestAnimationFrame(() => window.requestAnimationFrame(() => window.scrollTo(0, 0)))
    }

    resetMobileScroll()
    window.addEventListener('pageshow', resetMobileScroll)
    return () => window.removeEventListener('pageshow', resetMobileScroll)
  }, [])

  useEffect(() => {
    if (!workbook) localStorage.setItem(STORAGE_KEY, JSON.stringify(data))
  }, [data, workbook])

  const refreshWorkbook = async (client = workbook) => {
    if (!client) return undefined
    setSyncStatus('connecting')
    try {
      const latestData = await client.load()
      setData(latestData)
      setConnectedUser(client.userName)
      setSyncError('')
      setSyncStatus('synced')
      return latestData
    } catch (error) {
      setSyncError(error instanceof Error ? error.message : 'Excel could not be refreshed.')
      setSyncStatus('error')
      return undefined
    }
  }

  const connectWorkbook = async (config: M365Config, interactive = true) => {
    setSyncStatus('connecting')
    setSyncError('')
    try {
      const client = new M365Workbook(config)
      if (!await client.connect(interactive)) {
        setSyncStatus('sample')
        return
      }
      setM365Config(config)
      localStorage.setItem(M365_CONFIG_KEY, JSON.stringify(config))
      setWorkbook(client)
      await refreshWorkbook(client)
    } catch (error) {
      setSyncError(error instanceof Error ? error.message : 'Microsoft 365 connection failed.')
      setSyncStatus('error')
    }
  }

  useEffect(() => {
    if (m365Config.tenantId && m365Config.clientId && m365Config.workbookUrl) {
      void connectWorkbook(m365Config, false)
    }
    // The saved connection is restored once when the application starts.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => {
    if (!workbook) return
    const interval = window.setInterval(() => void refreshWorkbook(workbook), 60000)
    return () => window.clearInterval(interval)
    // refreshWorkbook is intentionally scoped to the active workbook instance.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [workbook])

  const activeLoans = data.loans.filter((loan) => !loan.returnedAt)
  const today = isoDate()
  const selectedLoan = data.loans.find((loan) => loan.id === selectedLoanId)
  const overdueLoans = activeLoans.filter((loan) => loan.dueAt < today)
  const loanedItemIds = new Set(activeLoans.map((loan) => loan.itemId))
  const filteredItems = useMemo(() => {
    const query = search.trim().toLowerCase()
    return data.items.filter((item) => {
      const category = data.categories.find((option) => option.id === item.categoryId)?.name ?? ''
      const group = data.groups.find((option) => option.id === item.primaryGroupId)?.name ?? ''
      const matchesSearch = [item.name, item.assetTag, category, group].join(' ').toLowerCase().includes(query)
      return matchesSearch && (categoryFilter === 'all' || item.categoryId === categoryFilter)
    })
  }, [categoryFilter, data.categories, data.groups, data.items, search])

  const optionName = (options: Option[], id: string) => options.find((option) => option.id === id)?.name ?? 'Ej angivet'
  const openItem = async (itemId = '') => {
    setFormError('')
    const latestData = workbook ? await refreshWorkbook(workbook) : data
    if (workbook && !latestData) return
    if (itemId && !latestData?.items.some((item) => item.id === itemId)) {
      setSyncError('Föremålet finns inte längre i Excel. Listan har uppdaterats.')
      setSyncStatus('error')
      return
    }
    setSelectedItemId(itemId)
    setModal(itemId ? 'edit-item' : 'item')
  }
  const openLoan = (itemId = '') => { setFormError(''); setSelectedItemId(itemId); setModal('loan') }
  const openLoanDetails = (loanId: string) => { setFormError(''); setSelectedLoanId(loanId); setModal('edit-loan') }

  const suggestInventorySequence = async (categoryId: string, currentItemId?: string) => {
    const latestData = workbook ? await refreshWorkbook(workbook) : data
    return nextInventorySequence(latestData ?? data, categoryId, currentItemId)
  }

  const saveRemote = async (operation: (client: M365Workbook) => Promise<unknown>) => {
    if (!workbook) return true
    setSyncStatus('saving')
    try {
      await operation(workbook)
      setSyncError('')
      setSyncStatus('synced')
      return true
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Excel could not be updated.'
      const validationConflict = /^(Inventarienumret|Kategorin|Prefixet|Föremålet)/.test(message)
      if (validationConflict) {
        setFormError(message)
        setSyncError('')
        setSyncStatus('synced')
      } else {
        setSyncError(message)
        setSyncStatus('error')
      }
      return false
    }
  }

  const addItem = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    const form = new FormData(event.currentTarget)
    const assetTag = normalizeAssetTag(String(form.get('assetTag')))
    const categoryId = String(form.get('categoryId'))
    const error = assetTagError(assetTag, data.items, data.categories, undefined, categoryId)
    if (error) { setFormError(error); return }
    const item: Item = {
      id: newId('item'), assetTag, name: String(form.get('name')),
      categoryId, primaryGroupId: String(form.get('primaryGroupId')),
      secondaryGroupIds: form.getAll('secondaryGroupIds').map(String), locationId: String(form.get('locationId')),
      quantity: Number(form.get('quantity')), notes: String(form.get('notes')),
    }
    if (await saveRemote((client) => client.addItem(item))) {
      setData((current) => ({ ...current, items: [...current.items, item] }))
      setModal(null)
    }
  }

  const updateItem = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    const currentItem = data.items.find((item) => item.id === selectedItemId)
    if (!currentItem) return
    const form = new FormData(event.currentTarget)
    const assetTag = normalizeAssetTag(String(form.get('assetTag')))
    const categoryId = String(form.get('categoryId'))
    const error = assetTagError(assetTag, data.items, data.categories, currentItem.id, categoryId)
    if (error) { setFormError(error); return }
    const item: Item = {
      ...currentItem, assetTag, name: String(form.get('name')),
      categoryId, primaryGroupId: String(form.get('primaryGroupId')),
      secondaryGroupIds: form.getAll('secondaryGroupIds').map(String), locationId: String(form.get('locationId')),
      quantity: Number(form.get('quantity')), notes: String(form.get('notes')),
    }
    if (await saveRemote((client) => client.updateItem(item))) {
      setData((current) => ({ ...current, items: current.items.map((entry) => entry.id === item.id ? item : entry) }))
      setModal(null)
    }
  }

  const deleteItem = async (item: Item) => {
    if (data.loans.some((loan) => loan.itemId === item.id)) {
      setFormError('Föremålet kan inte tas bort eftersom det finns i lånehistoriken.')
      return
    }
    if (await saveRemote((client) => client.deleteItem(item.id))) {
      setData((current) => ({ ...current, items: current.items.filter((entry) => entry.id !== item.id) }))
      setModal(null)
    }
  }

  const addLoan = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    const form = new FormData(event.currentTarget)
    const loan: Loan = {
      id: newId('loan'), itemId: String(form.get('itemId')), borrower: String(form.get('borrower')),
      borrowerGroupId: String(form.get('borrowerGroupId')), recordedBy: String(form.get('recordedBy')),
      lentAt: String(form.get('lentAt')), dueAt: String(form.get('dueAt')),
    }
    if (await saveRemote((client) => client.addLoan(loan))) {
      setData((current) => ({ ...current, loans: [loan, ...current.loans] }))
      setModal(null)
    }
  }

  const updateLoan = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (!selectedLoan) return
    const form = new FormData(event.currentTarget)
    const loan: Loan = {
      ...selectedLoan, itemId: String(form.get('itemId')), borrower: String(form.get('borrower')),
      borrowerGroupId: String(form.get('borrowerGroupId')), recordedBy: String(form.get('recordedBy')),
      lentAt: String(form.get('lentAt')), dueAt: String(form.get('dueAt')),
      returnedAt: form.has('returnedAt') ? String(form.get('returnedAt')) : selectedLoan.returnedAt,
    }
    if (await saveRemote((client) => client.updateLoan(loan))) {
      setData((current) => ({ ...current, loans: current.loans.map((entry) => entry.id === loan.id ? loan : entry) }))
      setModal(null)
    }
  }

  const returnLoan = async (loanId: string) => {
    const loan = data.loans.find((entry) => entry.id === loanId)
    if (!loan) return
    const returnedLoan = { ...loan, returnedAt: isoDate() }
    if (await saveRemote((client) => client.returnLoan(returnedLoan))) {
      setData((current) => ({ ...current, loans: current.loans.map((entry) => entry.id === loanId ? returnedLoan : entry) }))
    }
  }

  const addOption = async (kind: 'categories' | 'groups' | 'locations', name: string, prefix?: string) => {
    if (!name.trim()) return
    const option = { id: newId(kind.slice(0, 3)), name: name.trim(), ...(kind === 'categories' && { color: palette[data.categories.length % palette.length], prefix: normalizeCategoryPrefix(prefix ?? '') }) }
    if (await saveRemote((client) => client.addOption(kind, option))) {
      setData((current) => ({ ...current, [kind]: [...current[kind], option] }))
    }
  }
  const renameOption = async (kind: 'categories' | 'groups' | 'locations', id: string, name: string, prefix?: string) => {
    const currentOption = data[kind].find((option) => option.id === id)
    if (!currentOption || !name.trim()) return
    const renamed = { ...currentOption, name: name.trim(), ...(kind === 'categories' && { prefix: normalizeCategoryPrefix(prefix ?? currentOption.prefix ?? '') }) }
    if (currentOption.name === renamed.name && currentOption.prefix === renamed.prefix) return
    if (await saveRemote((client) => client.renameOption(kind, renamed))) {
      setData((current) => ({ ...current, [kind]: current[kind].map((option) => option.id === id ? renamed : option) }))
    }
  }

  const navItems: { id: View; label: string; icon: typeof LayoutDashboard; count?: number }[] = [
    { id: 'dashboard', label: 'Översikt', icon: LayoutDashboard },
    { id: 'inventory', label: 'Inventarier', icon: Boxes, count: data.items.length },
    { id: 'loans', label: 'Utlåning', icon: Handshake, count: activeLoans.length },
    { id: 'administration', label: 'Administration', icon: SlidersHorizontal },
    { id: 'documentation', label: 'Dokumentation', icon: BookOpen },
    { id: 'setup', label: 'Inställningar', icon: Settings },
  ]
  const navigate = (nextView: View) => { setView(nextView); setMobileNavOpen(false) }
  const startEdgeSwipe = (event: React.TouchEvent<HTMLDivElement>) => {
    const touch = event.touches[0]
    if (!touch || mobileNavOpen || window.innerWidth > 900 || touch.clientX > 32) return
    edgeSwipeStart.current = { x: touch.clientX, y: touch.clientY }
  }
  const finishEdgeSwipe = (event: React.TouchEvent<HTMLDivElement>) => {
    const start = edgeSwipeStart.current
    edgeSwipeStart.current = undefined
    const touch = event.changedTouches[0]
    if (!start || !touch) return

    const horizontalDistance = touch.clientX - start.x
    const verticalDistance = Math.abs(touch.clientY - start.y)
    if (horizontalDistance >= 72 && verticalDistance <= 48) setMobileNavOpen(true)
  }

  return (
    <div className="app-shell" onTouchStart={startEdgeSwipe} onTouchEnd={finishEdgeSwipe} onTouchCancel={() => { edgeSwipeStart.current = undefined }}>
      <aside className={`sidebar ${mobileNavOpen ? 'sidebar-open' : ''}`}>
        <div className="brand"><img src="https://lutherska.nu/wp-content/uploads/2025/08/logo_Lutherska_vit.png" alt="Lutherska Missionskyrkan" /><span>Inventarieregister</span></div>
        <nav aria-label="Main navigation">
          {navItems.map(({ id, label, icon: Icon, count }) => <button className={view === id ? 'nav-item active' : 'nav-item'} key={id} onClick={() => navigate(id)}><Icon size={19} /><span>{label}</span>{count !== undefined && <span className="nav-count">{count}</span>}</button>)}
        </nav>
        <div className={`sidebar-note sync-${syncStatus}`}>{workbook ? <Cloud size={15} /> : <CloudOff size={15} />}<div><strong>{syncStatus === 'synced' ? 'Synkad med Excel' : syncStatus === 'saving' ? 'Sparar i Excel...' : syncStatus === 'connecting' ? 'Ansluter...' : syncStatus === 'error' ? 'Synkfel' : 'Sampleläge'}</strong><small>{connectedUser || (workbook ? 'OneDrive-arbetsbok' : 'Anslut Microsoft 365 under Inställningar')}</small></div></div>
      </aside>
      {mobileNavOpen && <button className="nav-backdrop" aria-label="Close menu" onClick={() => setMobileNavOpen(false)} />}
      <main>
        <header className="topbar">
          <button className="icon-button mobile-menu" aria-label="Open menu" onClick={() => setMobileNavOpen(true)}><Menu /></button>
          <div><p className="eyebrow">Lutherska Missionskyrkan</p><h1>{navItems.find((item) => item.id === view)?.label}</h1></div>
          <div className="header-actions">
            {view !== 'setup' && (workbook ? <button className="button secondary" disabled={syncStatus === 'connecting'} onClick={() => void refreshWorkbook()}><RefreshCw size={17} className={syncStatus === 'connecting' ? 'spin' : ''} /> Uppdatera</button> : <button className="button secondary" onClick={() => navigate('setup')}><LogIn size={17} /> Anslut Microsoft 365</button>)}
            {(view === 'dashboard' || view === 'inventory' || view === 'loans') && <button className="button primary" onClick={() => view === 'loans' ? openLoan() : openItem()}><Plus size={18} /> {view === 'loans' ? 'Registrera lån' : 'Lägg till föremål'}</button>}
          </div>
        </header>
        <div className="content">
          {view === 'dashboard' && <Dashboard data={data} activeLoans={activeLoans} overdueLoans={overdueLoans} optionName={optionName} onNavigate={navigate} onEdit={openItem} onReturn={returnLoan} />}
          {view === 'inventory' && <Inventory data={data} items={filteredItems} search={search} categoryFilter={categoryFilter} loanedItemIds={loanedItemIds} optionName={optionName} onSearch={setSearch} onCategoryFilter={setCategoryFilter} onEdit={openItem} onLoan={openLoan} />}
          {view === 'loans' && <Loans data={data} optionName={optionName} today={today} onEdit={openLoanDetails} onReturn={returnLoan} />}
          {view === 'administration' && <Administration data={data} onAdd={addOption} onRename={renameOption} />}
          {view === 'setup' && <Setup m365Config={m365Config} syncStatus={syncStatus} syncError={syncError} connectedUser={connectedUser} onConnect={connectWorkbook} onRefresh={async () => { await refreshWorkbook() }} />}
          {view === 'documentation' && <Documentation />}
        </div>
      </main>
      {(modal === 'item' || modal === 'edit-item') && <ItemModal data={data} item={modal === 'edit-item' ? data.items.find((entry) => entry.id === selectedItemId) : undefined} error={formError} hasLoanHistory={modal === 'edit-item' && data.loans.some((loan) => loan.itemId === selectedItemId)} onSuggestSequence={suggestInventorySequence} onDelete={deleteItem} onErrorClear={() => setFormError('')} onClose={() => setModal(null)} onSubmit={modal === 'edit-item' ? updateItem : addItem} />}
      {modal === 'loan' && <LoanModal data={data} availableItems={data.items.filter((item) => !loanedItemIds.has(item.id))} selectedItemId={selectedItemId} onClose={() => setModal(null)} onSubmit={addLoan} />}
      {modal === 'edit-loan' && <LoanModal data={data} availableItems={data.items.filter((item) => Boolean(selectedLoan?.returnedAt) || item.id === selectedLoan?.itemId || !loanedItemIds.has(item.id))} loan={selectedLoan} onClose={() => setModal(null)} onSubmit={updateLoan} />}
    </div>
  )
}

function Documentation() {
  return <article className="documentation panel">
    <ReactMarkdown
      remarkPlugins={[remarkGfm]}
      components={{
        img: ({ src, alt }) => <img src={new URL((src ?? '').replace(/^public\//, ''), document.baseURI).href} alt={alt ?? ''} />,
        a: ({ href, children }) => <a href={href} target={href?.startsWith('http') ? '_blank' : undefined} rel="noreferrer">{children}</a>,
      }}
    >{readme}</ReactMarkdown>
  </article>
}

type SharedProps = { data: InventoryData; optionName: (options: Option[], id: string) => string }

function Dashboard({ data, activeLoans, overdueLoans, optionName, onNavigate, onEdit, onReturn }: SharedProps & { activeLoans: Loan[]; overdueLoans: Loan[]; onNavigate: (view: View) => void; onEdit: (itemId: string) => void; onReturn: (loanId: string) => void }) {
  return <div className="page-stack">
    <section className="metrics-grid" aria-label="Inventory summary">
      <Metric icon={Boxes} label="Föremål" value={data.items.length} detail={`${data.categories.length} kategorier`} tone="green" />
      <Metric icon={Handshake} label="Utlånade nu" value={activeLoans.length} detail="Alla grupper" tone="blue" />
      <Metric icon={CircleAlert} label="Försenade" value={overdueLoans.length} detail={overdueLoans.length ? 'Behöver följas upp' : 'Inget att följa upp'} tone="red" />
      <Metric icon={MapPin} label="Platser" value={data.locations.length} detail="Rum och förråd" tone="gold" />
    </section>
    {overdueLoans.length > 0 && <section className="overdue-panel">
      <div className="section-heading"><div><span className="section-kicker danger">Åtgärd behövs</span><h2>Försenade återlämningar</h2></div><button className="text-button" onClick={() => onNavigate('loans')}>Visa all utlåning</button></div>
      <div className="loan-list">{overdueLoans.map((loan) => {
        const item = data.items.find((entry) => entry.id === loan.itemId)
        const days = Math.ceil((Date.now() - new Date(loan.dueAt).getTime()) / 86400000)
        return <div className="loan-row overdue" key={loan.id}><div className="item-symbol"><CircleAlert size={20} /></div><div className="loan-main"><strong>{item?.name}</strong><span>{item?.assetTag} · {loan.borrower}</span></div><div className="loan-date"><span>Skulle åter {formatDate(loan.dueAt)}</span><strong>{days} dag{days === 1 ? '' : 'ar'} sen</strong></div><button className="button return" onClick={() => onReturn(loan.id)}><RotateCcw size={16} /> Markera återlämnad</button></div>
      })}</div>
    </section>}
    <div className="two-column">
      <section className="panel"><div className="section-heading"><div><span className="section-kicker">Register</span><h2>Senast tillagda</h2></div><button className="text-button" onClick={() => onNavigate('inventory')}>Hela inventariet</button></div><div className="compact-list">{data.items.slice(-3).reverse().map((item) => <div className="compact-row inventory-summary-row" tabIndex={0} role="button" aria-label={`Visa och redigera ${item.name}`} key={item.id} onClick={() => onEdit(item.id)} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); onEdit(item.id) } }}><span className="category-swatch" style={{ background: data.categories.find((category) => category.id === item.categoryId)?.color }} /><div><strong>{item.name}</strong><span>{item.assetTag} · {optionName(data.locations, item.locationId)}</span></div><span className="quiet-label">{optionName(data.categories, item.categoryId)}</span></div>)}</div></section>
      <section className="panel"><div className="section-heading"><div><span className="section-kicker">Ansvar</span><h2>Föremål per grupp</h2></div></div><div className="group-bars">{data.groups.map((group) => {
        const count = data.items.filter((item) => item.primaryGroupId === group.id).length
        const width = data.items.length ? `${Math.max((count / data.items.length) * 100, 4)}%` : '0%'
        return <div className="group-bar" key={group.id}><div><span>{group.name}</span><strong>{count}</strong></div><div className="bar-track"><span style={{ width }} /></div></div>
      })}</div></section>
    </div>
  </div>
}

function Metric({ icon: Icon, label, value, detail, tone }: { icon: typeof Boxes; label: string; value: number; detail: string; tone: string }) {
  return <article className="metric"><div className={`metric-icon ${tone}`}><Icon size={21} /></div><div><span>{label}</span><strong>{value}</strong><small>{detail}</small></div></article>
}

function Inventory({ data, items, search, categoryFilter, loanedItemIds, optionName, onSearch, onCategoryFilter, onEdit, onLoan }: SharedProps & { items: Item[]; search: string; categoryFilter: string; loanedItemIds: Set<string>; onSearch: (value: string) => void; onCategoryFilter: (value: string) => void; onEdit: (itemId: string) => void; onLoan: (itemId: string) => void }) {
  return <section className="panel table-panel">
    <div className="filters"><label className="search-box"><Search size={18} /><input aria-label="Sök i inventariet" placeholder="Sök namn, inventarienummer eller grupp" value={search} onChange={(event) => onSearch(event.target.value)} /></label><label className="select-box"><Tag size={17} /><select aria-label="Filtrera kategori" value={categoryFilter} onChange={(event) => onCategoryFilter(event.target.value)}><option value="all">Alla kategorier</option>{data.categories.map((category) => <option value={category.id} key={category.id}>{category.name}</option>)}</select><ChevronDown size={16} /></label><span className="result-count">{items.length} föremål</span></div>
    <div className="table-wrap"><table><thead><tr><th>Föremål</th><th>Kategori</th><th>Ansvarig</th><th>Ordinarie plats</th><th>Antal</th><th>Status</th><th><span className="sr-only">Åtgärder</span></th></tr></thead><tbody>{items.map((item) => {
      const loaned = loanedItemIds.has(item.id)
      return <tr className="inventory-row" tabIndex={0} aria-label={`Visa och redigera ${item.name}`} key={item.id} onClick={() => onEdit(item.id)} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); onEdit(item.id) } }}><td><div className="item-cell"><div className="item-symbol"><Boxes size={19} /></div><div><strong>{item.name}</strong><span>{item.assetTag}</span></div></div></td><td><span className="category-label"><i style={{ background: data.categories.find((category) => category.id === item.categoryId)?.color }} />{optionName(data.categories, item.categoryId)}</span></td><td><strong className="table-strong">{optionName(data.groups, item.primaryGroupId)}</strong>{item.secondaryGroupIds.length > 0 && <small>+ {item.secondaryGroupIds.length} sekundär</small>}</td><td><span className="with-icon"><MapPin size={15} />{optionName(data.locations, item.locationId)}</span></td><td>{item.quantity}</td><td><span className={`status ${loaned ? 'status-loan' : 'status-home'}`}>{loaned ? 'Utlånad' : 'På plats'}</span></td><td><button className="icon-text-button" disabled={loaned} onClick={(event) => { event.stopPropagation(); onLoan(item.id) }}><Handshake size={16} /> Låna ut</button></td></tr>
    })}</tbody></table>{!items.length && <div className="empty-state"><Search size={28} /><strong>Inga föremål hittades</strong><span>Prova en annan sökning eller kategori.</span></div>}</div>
  </section>
}

function Loans({ data, optionName, today, onEdit, onReturn }: SharedProps & { today: string; onEdit: (loanId: string) => void; onReturn: (loanId: string) => void }) {
  const sorted = [...data.loans].sort((a, b) => Number(Boolean(a.returnedAt)) - Number(Boolean(b.returnedAt)))
  return <section className="panel table-panel"><div className="panel-intro"><div><span className="section-kicker">Rörelselogg</span><h2>Alla lån och återlämningar</h2><p>Historik över vem som lånade och vem som registrerade händelsen.</p></div></div><div className="table-wrap"><table><thead><tr><th>Föremål</th><th>Låntagare</th><th>Registrerat av</th><th>Utlånat</th><th>Planerad retur</th><th>Status</th><th><span className="sr-only">Åtgärder</span></th></tr></thead><tbody>{sorted.map((loan) => {
    const item = data.items.find((entry) => entry.id === loan.itemId)
    const overdue = !loan.returnedAt && loan.dueAt < today
    return <tr className={`loan-table-row${overdue ? ' overdue-table-row' : ''}`} tabIndex={0} aria-label={`Visa och redigera lån för ${item?.name}`} key={loan.id} onClick={() => onEdit(loan.id)} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); onEdit(loan.id) } }}><td><div className="item-cell"><div className="item-symbol"><ClipboardList size={19} /></div><div><strong>{item?.name}</strong><span>{item?.assetTag}</span></div></div></td><td><div className="borrower-cell"><strong>{loan.borrower}</strong><span>{optionName(data.groups, loan.borrowerGroupId)}</span></div></td><td><span className="recorded-by">{loan.recordedBy}</span></td><td>{formatDate(loan.lentAt)}</td><td>{formatDate(loan.dueAt)}</td><td><span className={`status ${loan.returnedAt ? 'status-returned' : overdue ? 'status-overdue' : 'status-loan'}`}>{loan.returnedAt ? `Åter ${formatDate(loan.returnedAt)}` : overdue ? 'Försenad' : 'Utlånad'}</span></td><td>{!loan.returnedAt && <button className="icon-text-button" onClick={(event) => { event.stopPropagation(); onReturn(loan.id) }}><RotateCcw size={16} /> Återlämna</button>}</td></tr>
  })}</tbody></table></div></section>
}

function Administration({ data, onAdd, onRename }: {
  data: InventoryData
  onAdd: (kind: 'categories' | 'groups' | 'locations', name: string, prefix?: string) => void
  onRename: (kind: 'categories' | 'groups' | 'locations', id: string, name: string, prefix?: string) => void
}) {
  const integrityIssues = inventoryIntegrityIssues(data.items, data.categories)
  return <div className="administration-page">
    <div className="page-intro"><span className="section-kicker">Registervård</span><h2>Kategorier, grupper och platser</h2><p>Hantera de gemensamma val som används när föremål och lån registreras.</p></div>
    <section className={`panel integrity-panel ${integrityIssues.length ? 'has-issues' : ''}`}>
      <div className="integrity-copy"><CircleAlert size={22} /><div><span className="section-kicker">Inventarienummer</span><h2>{integrityIssues.length ? `${integrityIssues.length} problem behöver rättas` : 'Kontrollen är godkänd'}</h2><p>Använd tre svenska kategoribokstäver och tre siffror: <strong>MOB-001</strong>, <strong>BEL-001</strong>, <strong>LJU-001</strong> eller <strong>KOK-001</strong>. För en ny kategori väljs en unik, lättbegriplig trebokstavskod när kategorin skapas.</p></div></div>
      {integrityIssues.length > 0 && <ul>{integrityIssues.map((issue) => <li key={issue}>{issue}</li>)}</ul>}
    </section>
    <div className="setup-grid"><SetupList icon={Tag} title="Kategorier" description="Olika typer av föremål" kind="categories" options={data.categories} onAdd={onAdd} onRename={onRename} /><SetupList icon={Users} title="Verksamhetsgrupper" description="Primärt och sekundärt ansvar" kind="groups" options={data.groups} onAdd={onAdd} onRename={onRename} /><SetupList icon={MapPin} title="Platser" description="Fasta placeringar och förråd" kind="locations" options={data.locations} onAdd={onAdd} onRename={onRename} /></div>
  </div>
}

function Setup({ m365Config, syncStatus, syncError, connectedUser, onConnect, onRefresh }: {
  m365Config: M365Config
  syncStatus: 'sample' | 'connecting' | 'synced' | 'saving' | 'error'
  syncError: string
  connectedUser: string
  onConnect: (config: M365Config) => Promise<void>
  onRefresh: () => Promise<void>
}) {
  const connect = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    const form = new FormData(event.currentTarget)
    void onConnect({
      tenantId: String(form.get('tenantId')).trim(),
      clientId: String(form.get('clientId')).trim(),
      workbookUrl: String(form.get('workbookUrl')).trim(),
    })
  }
  return <div className="setup-page">
    <section className="connection-panel">
      <div className="connection-copy"><div className="connection-icon"><Cloud size={24} /></div><div><span className="section-kicker">Microsoft 365</span><h2>Excel på OneDrive</h2><p>Arbetsboken är gemensam datakälla. Alla ändringar i appen sparas direkt i Excel.</p></div></div>
      <form className="connection-form" onSubmit={connect}>
        <label>Tenant ID<input name="tenantId" required defaultValue={m365Config.tenantId} placeholder="00000000-0000-0000-0000-000000000000" /></label>
        <label>Application (client) ID<input name="clientId" required defaultValue={m365Config.clientId} placeholder="00000000-0000-0000-0000-000000000000" /></label>
        <label className="full">Delningslänk till Excel-arbetsboken<input name="workbookUrl" type="url" required defaultValue={m365Config.workbookUrl} placeholder="https://...sharepoint.com/:x:/..." /></label>
        <div className="connection-actions full">
          <span className={`connection-status status-${syncStatus}`}>{syncStatus === 'synced' ? `Ansluten som ${connectedUser}` : syncStatus === 'connecting' ? 'Ansluter...' : syncStatus === 'saving' ? 'Sparar...' : syncStatus === 'error' ? 'Anslutningen behöver åtgärdas' : 'Inte ansluten'}</span>
          {syncStatus === 'synced' && <button type="button" className="button secondary" onClick={() => void onRefresh()}><RefreshCw size={16} /> Uppdatera</button>}
          <button className="button primary" disabled={syncStatus === 'connecting'}><LogIn size={16} /> {syncStatus === 'synced' ? 'Byt anslutning' : 'Logga in och anslut'}</button>
        </div>
        {syncError && <p className="connection-error full"><CircleAlert size={15} /> {syncError}</p>}
      </form>
    </section>
    <section className="panel maintenance-panel">
      <div className="panel-intro"><span className="section-kicker">Löpande administration</span><h2>Skötsel av anslutning och åtkomst</h2><p>Det här behöver bara kontrolleras vid ändrade behörigheter eller anslutningsproblem.</p></div>
      <ul className="maintenance-list">
        <li>Kontrollera att synkstatus visar <strong>Ansluten</strong> innan viktiga ändringar görs.</li>
        <li>Hantera användarnas behörighet till Excel-filen i OneDrive.</li>
        <li>Använd OneDrives versionshistorik om en felaktig ändring behöver återställas.</li>
        <li>Flytta inte Excel-filen utan att även uppdatera arbetsbokslänken ovan.</li>
      </ul>
    </section>
  </div>
}

function SetupList({ icon: Icon, title, description, kind, options, onAdd, onRename }: { icon: typeof Tag; title: string; description: string; kind: 'categories' | 'groups' | 'locations'; options: Option[]; onAdd: (kind: 'categories' | 'groups' | 'locations', name: string, prefix?: string) => void; onRename: (kind: 'categories' | 'groups' | 'locations', id: string, name: string, prefix?: string) => void }) {
  const [name, setName] = useState('')
  const [prefix, setPrefix] = useState('')
  const [error, setError] = useState('')
  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    const normalizedPrefix = normalizeCategoryPrefix(prefix)
    if (kind === 'categories') {
      if (!/^[A-Z]{3}$/.test(normalizedPrefix)) { setError('Ange en kod med exakt tre bokstäver.'); return }
      if (options.some((option) => categoryPrefixFor(option.id, options) === normalizedPrefix)) { setError(`Koden ${normalizedPrefix} används redan.`); return }
    }
    onAdd(kind, name, normalizedPrefix)
    setName('')
    setPrefix('')
    setError('')
  }
  return <section className="panel setup-panel"><div className="setup-title"><div className="metric-icon green"><Icon size={20} /></div><div><h2>{title}</h2><p>{description}</p></div></div><div className="editable-list">{options.map((option) => <SetupOptionRow key={`${option.id}-${option.name}`} kind={kind} option={option} options={options} onRename={onRename} />)}</div><form className="quick-add" onSubmit={submit}><input aria-label={`New ${title.toLowerCase()}`} required placeholder={`Lägg till ${title.toLowerCase()}`} value={name} onChange={(event) => setName(event.target.value)} />{kind === 'categories' && <input className="prefix-input" aria-label="Kod för ny kategori" required pattern="[A-Za-z]{3}" maxLength={3} placeholder="KOD" value={prefix} onChange={(event) => { setPrefix(event.target.value.toUpperCase()); setError('') }} />}<button className="icon-button" aria-label={`Add ${title.toLowerCase()}`}><Plus size={18} /></button>{error && <small className="quick-add-error">{error}</small>}</form></section>
}

function SetupOptionRow({ kind, option, options, onRename }: { kind: 'categories' | 'groups' | 'locations'; option: Option; options: Option[]; onRename: (kind: 'categories' | 'groups' | 'locations', id: string, name: string, prefix?: string) => void }) {
  const [name, setName] = useState(option.name)
  const [prefix, setPrefix] = useState(categoryPrefixFor(option.id, options))
  const [error, setError] = useState('')
  const save = () => {
    if (kind === 'categories') {
      const normalizedPrefix = normalizeCategoryPrefix(prefix)
      if (!/^[A-Z]{3}$/.test(normalizedPrefix)) { setError('Koden ska ha tre bokstäver.'); return }
      if (options.some((candidate) => candidate.id !== option.id && categoryPrefixFor(candidate.id, options) === normalizedPrefix)) { setError(`${normalizedPrefix} används redan.`); return }
      setPrefix(normalizedPrefix)
      setError('')
      onRename(kind, option.id, name, normalizedPrefix)
      return
    }
    onRename(kind, option.id, name)
  }
  const fixedPrefix = Boolean(BUILT_IN_CATEGORY_PREFIXES[option.id])
  return <div className="editable-option-row">{option.color && <i style={{ background: option.color }} />}<input aria-label={`Rename ${option.name}`} value={name} onChange={(event) => setName(event.target.value)} onBlur={save} />{kind === 'categories' && <input className="prefix-input" aria-label={`Kod för ${option.name}`} pattern="[A-Za-z]{3}" maxLength={3} value={prefix} readOnly={fixedPrefix} title={fixedPrefix ? 'Fast kategorikod' : 'Trebokstavskod'} onChange={(event) => { setPrefix(event.target.value.toUpperCase()); setError('') }} onBlur={save} />}{error && <small>{error}</small>}</div>
}

function Modal({ title, subtitle, onClose, children }: { title: string; subtitle: string; onClose: () => void; children: React.ReactNode }) {
  return <div className="modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><section className="modal" role="dialog" aria-modal="true" aria-label={title}><div className="modal-header"><div><h2>{title}</h2><p>{subtitle}</p></div><button className="icon-button" aria-label="Close" onClick={onClose}><X size={19} /></button></div>{children}</section></div>
}

function ItemModal({ data, item, error, hasLoanHistory, onSuggestSequence, onDelete, onErrorClear, onClose, onSubmit }: { data: InventoryData; item?: Item; error: string; hasLoanHistory: boolean; onSuggestSequence: (categoryId: string, currentItemId?: string) => Promise<string>; onDelete: (item: Item) => Promise<void>; onErrorClear: () => void; onClose: () => void; onSubmit: (event: FormEvent<HTMLFormElement>) => void }) {
  const initialCategoryId = item?.categoryId ?? data.categories[0]?.id ?? ''
  const initialSequence = item?.assetTag.match(/(\d{3})$/)?.[1] ?? nextInventorySequence(data, initialCategoryId, item?.id)
  const [categoryId, setCategoryId] = useState(initialCategoryId)
  const [sequence, setSequence] = useState(initialSequence)
  const [sequenceLoading, setSequenceLoading] = useState(false)
  const [confirmDelete, setConfirmDelete] = useState(false)
  const prefix = categoryPrefixFor(categoryId, data.categories)
  const assetTag = `${prefix}-${sequence}`
  const changeCategory = async (nextCategoryId: string) => {
    setCategoryId(nextCategoryId)
    setSequenceLoading(true)
    setSequence(await onSuggestSequence(nextCategoryId, item?.id))
    setSequenceLoading(false)
    onErrorClear()
  }
  return <Modal title={item ? 'Visa och redigera föremål' : 'Lägg till föremål'} subtitle={item ? 'Kontrollera eller ändra föremålets uppgifter.' : 'Registrera var föremålet hör hemma och vem som ansvarar.'} onClose={onClose}>
    <form className="form-grid" onSubmit={onSubmit}>
      <label className="full">Namn<input name="name" required autoFocus defaultValue={item?.name} placeholder="t.ex. Portabel projektor" /></label>
      <label>Kategori<select name="categoryId" required value={categoryId} disabled={sequenceLoading} onChange={(event) => void changeCategory(event.target.value)}>{data.categories.map((option) => <option key={option.id} value={option.id}>{option.name}</option>)}</select></label>
      <label>Antal<input name="quantity" required type="number" min="1" defaultValue={item?.quantity ?? 1} /></label>
      <label>Inventarienummer<div className="asset-tag-control"><span aria-label={`Kategorikod ${prefix}`}>{prefix}-</span><input aria-label="Löpnummer" inputMode="numeric" required pattern="[0-9]{3}" maxLength={3} disabled={sequenceLoading} value={sequence} onChange={(event) => { setSequence(event.target.value.replace(/\D/g, '').slice(0, 3)); onErrorClear() }} /></div><input name="assetTag" type="hidden" value={assetTag} /><small>{sequenceLoading ? 'Kontrollerar nästa lediga nummer...' : 'Nästa lediga nummer är ifyllt. Det kan ändras.'}</small></label>
      <label>Ansvarig grupp<select name="primaryGroupId" required defaultValue={item?.primaryGroupId}>{data.groups.map((option) => <option key={option.id} value={option.id}>{option.name}</option>)}</select></label>
      <label className="full">Sekundära grupper<select name="secondaryGroupIds" multiple size={Math.min(data.groups.length, 4)} defaultValue={item?.secondaryGroupIds}>{data.groups.map((option) => <option key={option.id} value={option.id}>{option.name}</option>)}</select><small>Håll Ctrl för att välja flera.</small></label>
      <label className="full">Ordinarie plats<select name="locationId" required defaultValue={item?.locationId}>{data.locations.map((option) => <option key={option.id} value={option.id}>{option.name}</option>)}</select></label>
      <label className="full">Anteckningar<textarea name="notes" rows={3} defaultValue={item?.notes} placeholder="Skick, tillbehör eller andra användbara uppgifter" /></label>
      {error && <p className="form-error full"><CircleAlert size={16} /> {error}</p>}
      {item && <div className="delete-item-area full">{confirmDelete ? <><span>Ta bort {item.name} permanent?</span><button type="button" className="button danger" onClick={() => void onDelete(item)}><Trash2 size={16} /> Bekräfta borttagning</button><button type="button" className="text-button" onClick={() => setConfirmDelete(false)}>Behåll</button></> : <><button type="button" className="button danger" disabled={hasLoanHistory} onClick={() => setConfirmDelete(true)}><Trash2 size={16} /> Ta bort föremål</button>{hasLoanHistory && <span>Kan inte tas bort eftersom föremålet finns i lånehistoriken.</span>}</>}</div>}
      <div className="form-actions full"><button type="button" className="button secondary" onClick={onClose}>Avbryt</button><button className="button primary" disabled={sequenceLoading}>{item ? <Check size={17} /> : <PackagePlus size={17} />} {item ? 'Spara ändringar' : 'Lägg till föremål'}</button></div>
    </form>
  </Modal>
}

function LoanModal({ data, availableItems, selectedItemId = '', loan, onClose, onSubmit }: { data: InventoryData; availableItems: Item[]; selectedItemId?: string; loan?: Loan; onClose: () => void; onSubmit: (event: FormEvent<HTMLFormElement>) => void }) {
  const itemId = loan?.itemId ?? selectedItemId
  return <Modal title={loan ? 'Visa och redigera lån' : 'Registrera lån'} subtitle={loan ? 'Kontrollera eller ändra uppgifterna för lånet.' : 'Inget godkännande behövs. Registrera låntagare och planerad retur.'} onClose={onClose}><form className="form-grid" onSubmit={onSubmit}><label className="full">Föremål<select name="itemId" required defaultValue={itemId}>{!itemId && <option value="">Välj ett tillgängligt föremål</option>}{availableItems.map((item) => <option key={item.id} value={item.id}>{item.assetTag} · {item.name}</option>)}</select></label><label>Låntagare<input name="borrower" required autoFocus defaultValue={loan?.borrower} placeholder="För- och efternamn" /></label><label>Låntagarens grupp<select name="borrowerGroupId" required defaultValue={loan?.borrowerGroupId}>{data.groups.map((option) => <option key={option.id} value={option.id}>{option.name}</option>)}</select></label><label>Utlånad den<input name="lentAt" type="date" required defaultValue={loan?.lentAt ?? isoDate()} /></label><label>Planerad retur<input name="dueAt" type="date" required min={loan ? undefined : isoDate()} defaultValue={loan?.dueAt ?? isoDate(7)} /></label>{loan?.returnedAt && <label>Återlämnad den<input name="returnedAt" type="date" required defaultValue={loan.returnedAt} /></label>}<label className={loan?.returnedAt ? '' : 'full'}>Registrerat av<input name="recordedBy" required defaultValue={loan?.recordedBy} placeholder="Ditt namn" /></label><div className="form-actions full"><button type="button" className="button secondary" onClick={onClose}>Avbryt</button><button className="button primary"><Check size={17} /> {loan ? 'Spara ändringar' : 'Registrera lån'}</button></div></form></Modal>
}

function formatDate(value: string) {
  return new Intl.DateTimeFormat('sv-SE', { day: 'numeric', month: 'short', year: 'numeric' }).format(new Date(`${value}T12:00:00`))
}

export default App
