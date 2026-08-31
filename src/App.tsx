import { FormEvent, useEffect, useMemo, useState } from 'react'
import {
  Boxes,
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
  Tag,
  Users,
  X,
} from 'lucide-react'
import './App.css'
import {
  M365Workbook,
  type InventoryData,
  type Item,
  type Loan,
  type M365Config,
  type Option,
} from './m365Workbook'

type View = 'dashboard' | 'inventory' | 'loans' | 'setup'

const STORAGE_KEY = 'church-inventory-v3'
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
    { id: 'cat-furniture', name: 'Möbler', color: '#6f8657' },
    { id: 'cat-lighting', name: 'Belysning', color: '#f59b45' },
    { id: 'cat-sound', name: 'Ljudutrustning', color: '#3688df' },
    { id: 'cat-kitchen', name: 'Kök & fika', color: '#c73f3b' },
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
    { id: 'item-1', assetTag: 'SND-001', name: 'Trådlösa mikrofoner', categoryId: 'cat-sound', primaryGroupId: 'grp-sound', secondaryGroupIds: ['grp-lux'], locationId: 'loc-stage', quantity: 2, notes: 'Två handmikrofoner och mottagare.' },
    { id: 'item-2', assetTag: 'FUR-014', name: 'Fällbord', categoryId: 'cat-furniture', primaryGroupId: 'grp-premises-change', secondaryGroupIds: ['grp-household'], locationId: 'loc-basement', quantity: 8, notes: '' },
    { id: 'item-3', assetTag: 'KIT-008', name: 'Stor kaffebryggare', categoryId: 'cat-kitchen', primaryGroupId: 'grp-household', secondaryGroupIds: ['grp-kitchen-cleaning-purchases'], locationId: 'loc-kitchen', quantity: 1, notes: 'Rymmer 10 liter.' },
    { id: 'item-4', assetTag: 'LGT-004', name: 'Portabla LED-lampor', categoryId: 'cat-lighting', primaryGroupId: 'grp-service-visuals', secondaryGroupIds: ['grp-lux'], locationId: 'loc-stage', quantity: 6, notes: 'Förvaras i vadderad transportväska.' },
  ],
  loans: [
    { id: 'loan-1', itemId: 'item-4', borrower: 'Anna Lind', borrowerGroupId: 'grp-lux', recordedBy: 'Erik Nilsson', lentAt: isoDate(-6), dueAt: isoDate(-1) },
    { id: 'loan-2', itemId: 'item-2', borrower: 'Jonas Berg', borrowerGroupId: 'grp-household', recordedBy: 'Sara Holm', lentAt: isoDate(-1), dueAt: isoDate(4) },
  ],
}

const newId = (prefix: string) => `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`

function App() {
  const [data, setData] = useState<InventoryData>(() => {
    const stored = localStorage.getItem(STORAGE_KEY)
    return stored ? (JSON.parse(stored) as InventoryData) : seedData
  })
  const [view, setView] = useState<View>('dashboard')
  const [search, setSearch] = useState('')
  const [categoryFilter, setCategoryFilter] = useState('all')
  const [modal, setModal] = useState<'item' | 'loan' | null>(null)
  const [selectedItemId, setSelectedItemId] = useState('')
  const [mobileNavOpen, setMobileNavOpen] = useState(false)
  const [m365Config, setM365Config] = useState<M365Config>(() => {
    const stored = localStorage.getItem(M365_CONFIG_KEY)
    return stored ? JSON.parse(stored) as M365Config : deployedM365Config
  })
  const [workbook, setWorkbook] = useState<M365Workbook>()
  const [syncStatus, setSyncStatus] = useState<'sample' | 'connecting' | 'synced' | 'saving' | 'error'>('sample')
  const [syncError, setSyncError] = useState('')
  const [connectedUser, setConnectedUser] = useState('')

  useEffect(() => {
    if (!workbook) localStorage.setItem(STORAGE_KEY, JSON.stringify(data))
  }, [data, workbook])

  const refreshWorkbook = async (client = workbook) => {
    if (!client) return
    setSyncStatus('connecting')
    try {
      setData(await client.load())
      setConnectedUser(client.userName)
      setSyncError('')
      setSyncStatus('synced')
    } catch (error) {
      setSyncError(error instanceof Error ? error.message : 'Excel could not be refreshed.')
      setSyncStatus('error')
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
  const openLoan = (itemId = '') => { setSelectedItemId(itemId); setModal('loan') }

  const saveRemote = async (operation: (client: M365Workbook) => Promise<unknown>) => {
    if (!workbook) return true
    setSyncStatus('saving')
    try {
      await operation(workbook)
      setSyncError('')
      setSyncStatus('synced')
      return true
    } catch (error) {
      setSyncError(error instanceof Error ? error.message : 'Excel could not be updated.')
      setSyncStatus('error')
      return false
    }
  }

  const addItem = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    const form = new FormData(event.currentTarget)
    const item: Item = {
      id: newId('item'), assetTag: String(form.get('assetTag')), name: String(form.get('name')),
      categoryId: String(form.get('categoryId')), primaryGroupId: String(form.get('primaryGroupId')),
      secondaryGroupIds: form.getAll('secondaryGroupIds').map(String), locationId: String(form.get('locationId')),
      quantity: Number(form.get('quantity')), notes: String(form.get('notes')),
    }
    if (await saveRemote((client) => client.addItem(item))) {
      setData((current) => ({ ...current, items: [...current.items, item] }))
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

  const returnLoan = async (loanId: string) => {
    const loan = data.loans.find((entry) => entry.id === loanId)
    if (!loan) return
    const returnedLoan = { ...loan, returnedAt: isoDate() }
    if (await saveRemote((client) => client.returnLoan(returnedLoan))) {
      setData((current) => ({ ...current, loans: current.loans.map((entry) => entry.id === loanId ? returnedLoan : entry) }))
    }
  }

  const addOption = async (kind: 'categories' | 'groups' | 'locations', name: string) => {
    if (!name.trim()) return
    const option = { id: newId(kind.slice(0, 3)), name: name.trim(), ...(kind === 'categories' && { color: palette[data.categories.length % palette.length] }) }
    if (await saveRemote((client) => client.addOption(kind, option))) {
      setData((current) => ({ ...current, [kind]: [...current[kind], option] }))
    }
  }
  const renameOption = async (kind: 'categories' | 'groups' | 'locations', id: string, name: string) => {
    const currentOption = data[kind].find((option) => option.id === id)
    if (!currentOption || currentOption.name === name || !name.trim()) return
    const renamed = { ...currentOption, name: name.trim() }
    if (await saveRemote((client) => client.renameOption(kind, renamed))) {
      setData((current) => ({ ...current, [kind]: current[kind].map((option) => option.id === id ? renamed : option) }))
    }
  }

  const navItems: { id: View; label: string; icon: typeof LayoutDashboard; count?: number }[] = [
    { id: 'dashboard', label: 'Översikt', icon: LayoutDashboard },
    { id: 'inventory', label: 'Inventarier', icon: Boxes, count: data.items.length },
    { id: 'loans', label: 'Utlåning', icon: Handshake, count: activeLoans.length },
    { id: 'setup', label: 'Inställningar', icon: Settings },
  ]
  const navigate = (nextView: View) => { setView(nextView); setMobileNavOpen(false) }

  return (
    <div className="app-shell">
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
            {workbook ? <button className="button secondary" disabled={syncStatus === 'connecting'} onClick={() => void refreshWorkbook()}><RefreshCw size={17} className={syncStatus === 'connecting' ? 'spin' : ''} /> Uppdatera</button> : <button className="button secondary" onClick={() => navigate('setup')}><LogIn size={17} /> Anslut Microsoft 365</button>}
            <button className="button primary" onClick={() => setModal(view === 'loans' ? 'loan' : 'item')}><Plus size={18} /> {view === 'loans' ? 'Registrera lån' : 'Lägg till föremål'}</button>
          </div>
        </header>
        <div className="content">
          {view === 'dashboard' && <Dashboard data={data} activeLoans={activeLoans} overdueLoans={overdueLoans} optionName={optionName} onNavigate={navigate} onReturn={returnLoan} />}
          {view === 'inventory' && <Inventory data={data} items={filteredItems} search={search} categoryFilter={categoryFilter} loanedItemIds={loanedItemIds} optionName={optionName} onSearch={setSearch} onCategoryFilter={setCategoryFilter} onLoan={openLoan} />}
          {view === 'loans' && <Loans data={data} optionName={optionName} today={today} onReturn={returnLoan} />}
          {view === 'setup' && <Setup data={data} m365Config={m365Config} syncStatus={syncStatus} syncError={syncError} connectedUser={connectedUser} onConnect={connectWorkbook} onRefresh={() => refreshWorkbook()} onAdd={addOption} onRename={renameOption} />}
        </div>
      </main>
      {modal === 'item' && <ItemModal data={data} onClose={() => setModal(null)} onSubmit={addItem} />}
      {modal === 'loan' && <LoanModal data={data} availableItems={data.items.filter((item) => !loanedItemIds.has(item.id))} selectedItemId={selectedItemId} onClose={() => setModal(null)} onSubmit={addLoan} />}
    </div>
  )
}

type SharedProps = { data: InventoryData; optionName: (options: Option[], id: string) => string }

function Dashboard({ data, activeLoans, overdueLoans, optionName, onNavigate, onReturn }: SharedProps & { activeLoans: Loan[]; overdueLoans: Loan[]; onNavigate: (view: View) => void; onReturn: (loanId: string) => void }) {
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
      <section className="panel"><div className="section-heading"><div><span className="section-kicker">Register</span><h2>Senast tillagda</h2></div><button className="text-button" onClick={() => onNavigate('inventory')}>Hela inventariet</button></div><div className="compact-list">{data.items.slice(-3).reverse().map((item) => <div className="compact-row" key={item.id}><span className="category-swatch" style={{ background: data.categories.find((category) => category.id === item.categoryId)?.color }} /><div><strong>{item.name}</strong><span>{item.assetTag} · {optionName(data.locations, item.locationId)}</span></div><span className="quiet-label">{optionName(data.categories, item.categoryId)}</span></div>)}</div></section>
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

function Inventory({ data, items, search, categoryFilter, loanedItemIds, optionName, onSearch, onCategoryFilter, onLoan }: SharedProps & { items: Item[]; search: string; categoryFilter: string; loanedItemIds: Set<string>; onSearch: (value: string) => void; onCategoryFilter: (value: string) => void; onLoan: (itemId: string) => void }) {
  return <section className="panel table-panel">
    <div className="filters"><label className="search-box"><Search size={18} /><input aria-label="Sök i inventariet" placeholder="Sök namn, märkning eller grupp" value={search} onChange={(event) => onSearch(event.target.value)} /></label><label className="select-box"><Tag size={17} /><select aria-label="Filtrera kategori" value={categoryFilter} onChange={(event) => onCategoryFilter(event.target.value)}><option value="all">Alla kategorier</option>{data.categories.map((category) => <option value={category.id} key={category.id}>{category.name}</option>)}</select><ChevronDown size={16} /></label><span className="result-count">{items.length} föremål</span></div>
    <div className="table-wrap"><table><thead><tr><th>Föremål</th><th>Kategori</th><th>Ansvarig</th><th>Ordinarie plats</th><th>Antal</th><th>Status</th><th><span className="sr-only">Åtgärder</span></th></tr></thead><tbody>{items.map((item) => {
      const loaned = loanedItemIds.has(item.id)
      return <tr key={item.id}><td><div className="item-cell"><div className="item-symbol"><Boxes size={19} /></div><div><strong>{item.name}</strong><span>{item.assetTag}</span></div></div></td><td><span className="category-label"><i style={{ background: data.categories.find((category) => category.id === item.categoryId)?.color }} />{optionName(data.categories, item.categoryId)}</span></td><td><strong className="table-strong">{optionName(data.groups, item.primaryGroupId)}</strong>{item.secondaryGroupIds.length > 0 && <small>+ {item.secondaryGroupIds.length} sekundär</small>}</td><td><span className="with-icon"><MapPin size={15} />{optionName(data.locations, item.locationId)}</span></td><td>{item.quantity}</td><td><span className={`status ${loaned ? 'status-loan' : 'status-home'}`}>{loaned ? 'Utlånad' : 'På plats'}</span></td><td><button className="icon-text-button" disabled={loaned} onClick={() => onLoan(item.id)}><Handshake size={16} /> Låna ut</button></td></tr>
    })}</tbody></table>{!items.length && <div className="empty-state"><Search size={28} /><strong>Inga föremål hittades</strong><span>Prova en annan sökning eller kategori.</span></div>}</div>
  </section>
}

function Loans({ data, optionName, today, onReturn }: SharedProps & { today: string; onReturn: (loanId: string) => void }) {
  const sorted = [...data.loans].sort((a, b) => Number(Boolean(a.returnedAt)) - Number(Boolean(b.returnedAt)))
  return <section className="panel table-panel"><div className="panel-intro"><div><span className="section-kicker">Rörelselogg</span><h2>Alla lån och återlämningar</h2><p>Historik över vem som lånade och vem som registrerade händelsen.</p></div></div><div className="table-wrap"><table><thead><tr><th>Föremål</th><th>Låntagare</th><th>Registrerat av</th><th>Utlånat</th><th>Planerad retur</th><th>Status</th><th><span className="sr-only">Åtgärder</span></th></tr></thead><tbody>{sorted.map((loan) => {
    const item = data.items.find((entry) => entry.id === loan.itemId)
    const overdue = !loan.returnedAt && loan.dueAt < today
    return <tr className={overdue ? 'overdue-table-row' : ''} key={loan.id}><td><div className="item-cell"><div className="item-symbol"><ClipboardList size={19} /></div><div><strong>{item?.name}</strong><span>{item?.assetTag}</span></div></div></td><td><strong className="table-strong">{loan.borrower}</strong><small>{optionName(data.groups, loan.borrowerGroupId)}</small></td><td>{loan.recordedBy}</td><td>{formatDate(loan.lentAt)}</td><td>{formatDate(loan.dueAt)}</td><td><span className={`status ${loan.returnedAt ? 'status-returned' : overdue ? 'status-overdue' : 'status-loan'}`}>{loan.returnedAt ? `Åter ${formatDate(loan.returnedAt)}` : overdue ? 'Försenad' : 'Utlånad'}</span></td><td>{!loan.returnedAt && <button className="icon-text-button" onClick={() => onReturn(loan.id)}><RotateCcw size={16} /> Återlämna</button>}</td></tr>
  })}</tbody></table></div></section>
}

function Setup({ data, m365Config, syncStatus, syncError, connectedUser, onConnect, onRefresh, onAdd, onRename }: {
  data: InventoryData
  m365Config: M365Config
  syncStatus: 'sample' | 'connecting' | 'synced' | 'saving' | 'error'
  syncError: string
  connectedUser: string
  onConnect: (config: M365Config) => Promise<void>
  onRefresh: () => Promise<void>
  onAdd: (kind: 'categories' | 'groups' | 'locations', name: string) => void
  onRename: (kind: 'categories' | 'groups' | 'locations', id: string, name: string) => void
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
    <div className="setup-grid"><SetupList icon={Tag} title="Kategorier" description="Olika typer av föremål" kind="categories" options={data.categories} onAdd={onAdd} onRename={onRename} /><SetupList icon={Users} title="Verksamhetsgrupper" description="Primärt och sekundärt ansvar" kind="groups" options={data.groups} onAdd={onAdd} onRename={onRename} /><SetupList icon={MapPin} title="Platser" description="Fasta placeringar och förråd" kind="locations" options={data.locations} onAdd={onAdd} onRename={onRename} /></div>
  </div>
}

function SetupList({ icon: Icon, title, description, kind, options, onAdd, onRename }: { icon: typeof Tag; title: string; description: string; kind: 'categories' | 'groups' | 'locations'; options: Option[]; onAdd: (kind: 'categories' | 'groups' | 'locations', name: string) => void; onRename: (kind: 'categories' | 'groups' | 'locations', id: string, name: string) => void }) {
  const [name, setName] = useState('')
  return <section className="panel setup-panel"><div className="setup-title"><div className="metric-icon green"><Icon size={20} /></div><div><h2>{title}</h2><p>{description}</p></div></div><div className="editable-list">{options.map((option) => <label key={`${option.id}-${option.name}`}>{option.color && <i style={{ background: option.color }} />}<input aria-label={`Rename ${option.name}`} defaultValue={option.name} onBlur={(event) => onRename(kind, option.id, event.target.value)} /></label>)}</div><form className="quick-add" onSubmit={(event) => { event.preventDefault(); onAdd(kind, name); setName('') }}><input aria-label={`New ${title.toLowerCase()}`} placeholder={`Add ${title.toLowerCase()}`} value={name} onChange={(event) => setName(event.target.value)} /><button className="icon-button" aria-label={`Add ${title.toLowerCase()}`}><Plus size={18} /></button></form></section>
}

function Modal({ title, subtitle, onClose, children }: { title: string; subtitle: string; onClose: () => void; children: React.ReactNode }) {
  return <div className="modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><section className="modal" role="dialog" aria-modal="true" aria-label={title}><div className="modal-header"><div><h2>{title}</h2><p>{subtitle}</p></div><button className="icon-button" aria-label="Close" onClick={onClose}><X size={19} /></button></div>{children}</section></div>
}

function ItemModal({ data, onClose, onSubmit }: { data: InventoryData; onClose: () => void; onSubmit: (event: FormEvent<HTMLFormElement>) => void }) {
  return <Modal title="Lägg till föremål" subtitle="Registrera var föremålet hör hemma och vem som ansvarar." onClose={onClose}><form className="form-grid" onSubmit={onSubmit}><label className="full">Namn<input name="name" required autoFocus placeholder="t.ex. Portabel projektor" /></label><label>Märkning<input name="assetTag" required placeholder="t.ex. AV-012" /></label><label>Antal<input name="quantity" required type="number" min="1" defaultValue="1" /></label><label>Kategori<select name="categoryId" required>{data.categories.map((option) => <option key={option.id} value={option.id}>{option.name}</option>)}</select></label><label>Ansvarig grupp<select name="primaryGroupId" required>{data.groups.map((option) => <option key={option.id} value={option.id}>{option.name}</option>)}</select></label><label className="full">Sekundära grupper<select name="secondaryGroupIds" multiple size={Math.min(data.groups.length, 4)}>{data.groups.map((option) => <option key={option.id} value={option.id}>{option.name}</option>)}</select><small>Håll Ctrl för att välja flera.</small></label><label className="full">Ordinarie plats<select name="locationId" required>{data.locations.map((option) => <option key={option.id} value={option.id}>{option.name}</option>)}</select></label><label className="full">Anteckningar<textarea name="notes" rows={3} placeholder="Skick, tillbehör eller andra användbara uppgifter" /></label><div className="form-actions full"><button type="button" className="button secondary" onClick={onClose}>Avbryt</button><button className="button primary"><PackagePlus size={17} /> Lägg till föremål</button></div></form></Modal>
}

function LoanModal({ data, availableItems, selectedItemId, onClose, onSubmit }: { data: InventoryData; availableItems: Item[]; selectedItemId: string; onClose: () => void; onSubmit: (event: FormEvent<HTMLFormElement>) => void }) {
  return <Modal title="Registrera lån" subtitle="Inget godkännande behövs. Registrera låntagare och planerad retur." onClose={onClose}><form className="form-grid" onSubmit={onSubmit}><label className="full">Föremål<select name="itemId" required defaultValue={selectedItemId}>{!selectedItemId && <option value="">Välj ett tillgängligt föremål</option>}{availableItems.map((item) => <option key={item.id} value={item.id}>{item.assetTag} · {item.name}</option>)}</select></label><label>Låntagare<input name="borrower" required autoFocus placeholder="För- och efternamn" /></label><label>Låntagarens grupp<select name="borrowerGroupId" required>{data.groups.map((option) => <option key={option.id} value={option.id}>{option.name}</option>)}</select></label><label>Utlånad den<input name="lentAt" type="date" required defaultValue={isoDate()} /></label><label>Planerad retur<input name="dueAt" type="date" required min={isoDate()} defaultValue={isoDate(7)} /></label><label className="full">Registrerat av<input name="recordedBy" required placeholder="Ditt namn" /></label><div className="form-actions full"><button type="button" className="button secondary" onClick={onClose}>Avbryt</button><button className="button primary"><Check size={17} /> Registrera lån</button></div></form></Modal>
}

function formatDate(value: string) {
  return new Intl.DateTimeFormat('sv-SE', { day: 'numeric', month: 'short', year: 'numeric' }).format(new Date(`${value}T12:00:00`))
}

export default App
