import { useCallback, useEffect, useRef, useState } from 'react';
import { Grid2X2, Plus, FolderOpen, Clock3, Pin, Search, ArrowUpRight, Download, Check, X, Sun, Moon, Monitor, History, ShieldCheck, Settings2, FileSpreadsheet, ChevronDown, BarChart3, Printer, BookOpen, Undo2, Redo2, ArrowLeft, Copy, Trash2 } from 'lucide-react';
import { blankBook, BookRecord, usedExtent, parseNative } from './model';
import { fromTemplate, templates } from './templates';
import { store } from './store';
import { download, isNative, openNative, writeNative } from './files';
import { Modal } from './Modal';
import type { SheetEngine } from './engine';

type Appearance = 'light' | 'dark' | 'system';
const titleFile = (title: string) => title.replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').trim() || 'Workbook';
const escapeHtml = (value: string) => value.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]!));

export default function App() {
  const [book, setBook] = useState(blankBook);
  const current = useRef(book); current.current = book;
  const [home, setHome] = useState(true);
  const [library, setLibrary] = useState<BookRecord[]>([]);
  const [filter, setFilter] = useState('All workbooks');
  const [search, setSearch] = useState('');
  const [modal, setModal] = useState('');
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [status, setStatus] = useState<'saved' | 'saving' | 'dirty'>('saved');
  const [loaded, setLoaded] = useState(false);
  const [session, setSession] = useState(0);
  const [engineReady, setEngineReady] = useState(false);
  const [address, setAddress] = useState('A1');
  const [raw, setRaw] = useState('');
  const [displayed, setDisplayed] = useState('');
  const [versionName, setVersionName] = useState('');
  const [chartData, setChartData] = useState<{ label: string; value: number }[]>([]);
  const [appearance, setAppearance] = useState<Appearance>(() => (localStorage.getItem('sheetcloud-appearance') as Appearance) || 'system');
  const [pinned, setPinned] = useState<string[]>(() => { try { return JSON.parse(localStorage.getItem('sheetcloud-pinned') || '[]'); } catch { return []; } });
  const host = useRef<HTMLDivElement>(null);
  const engine = useRef<SheetEngine | null>(null);
  const fileInput = useRef<HTMLInputElement>(null);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const sequence = useRef(0);
  const baseline = useRef<{ path: string; bytes: Uint8Array } | undefined>(undefined);
  const dark = appearance === 'dark' || (appearance === 'system' && matchMedia('(prefers-color-scheme: dark)').matches);

  const refreshLibrary = useCallback(async () => { setLibrary((await store.list()).sort((a, b) => b.updatedAt.localeCompare(a.updatedAt))); }, []);
  useEffect(() => { refreshLibrary().catch(e => setError(String(e))).finally(() => setLoaded(true)); }, [refreshLibrary]);
  useEffect(() => { if (home && loaded) refreshLibrary().catch(e => setError(String(e))); }, [home, loaded, refreshLibrary]);
  useEffect(() => { if (!notice) return; const timeout = setTimeout(() => setNotice(''), 6000); return () => clearTimeout(timeout); }, [notice]);
  useEffect(() => {
    const media = matchMedia('(prefers-color-scheme: dark)');
    const apply = () => { const value = appearance === 'dark' || (appearance === 'system' && media.matches); document.documentElement.classList.toggle('dark', value); engine.current?.api.toggleDarkMode(value); };
    apply(); media.addEventListener('change', apply); localStorage.setItem('sheetcloud-appearance', appearance);
    return () => media.removeEventListener('change', apply);
  }, [appearance]);

  const snapshot = useCallback((): BookRecord => ({ ...current.current, updatedAt: new Date().toISOString(), workbook: engine.current?.save() || current.current.workbook }), []);
  const persist = useCallback(async () => {
    if (timer.current) { clearTimeout(timer.current); timer.current = null; }
    const savedSequence = sequence.current;
    const next = snapshot();
    setStatus('saving'); await store.save(next);
    if (savedSequence === sequence.current) { setStatus('saved'); current.current = next; setBook(next); }
    return next;
  }, [snapshot]);
  const changed = useCallback(() => {
    if (!engine.current) return;
    sequence.current++; setStatus('dirty');
    if (timer.current) clearTimeout(timer.current);
    timer.current = setTimeout(() => { persist().catch(e => { setStatus('dirty'); setError(`Autosave failed. Export a native backup: ${String(e)}`); }); }, 800);
  }, [persist]);

  useEffect(() => {
    if (home || !host.current) return;
    let disposed = false;
    let controller: SheetEngine | undefined;
    setEngineReady(false);
    import('./engine').then(({ createEngine }) => {
      if (disposed) return;
      controller = createEngine(host.current!, current.current, changed, (a, value, display) => {
        setDisplayed(display);
        if (document.activeElement?.getAttribute('aria-label') !== 'Cell address') setAddress(a);
        if (document.activeElement?.getAttribute('aria-label') !== 'Cell value or formula') setRaw(value);
      }, dark);
      engine.current = controller; setEngineReady(true);
    }).catch(e => setError(`Could not start the spreadsheet engine: ${String(e)}`));
    return () => { disposed = true; if (timer.current) clearTimeout(timer.current); if (engine.current === controller) engine.current = null; controller?.dispose(); };
  }, [home, session, changed]);

  useEffect(() => { const before = (e: BeforeUnloadEvent) => { if (status !== 'saved') { e.preventDefault(); e.returnValue = ''; } }; addEventListener('beforeunload', before); return () => removeEventListener('beforeunload', before); }, [status]);

  async function load(next: BookRecord) {
    try { if (!home && engine.current) await persist(); baseline.current = undefined; current.current = next; setBook(next); sequence.current++; setStatus('dirty'); setHome(false); setSession(n => n + 1); setModal(''); await store.save(next); setStatus('saved'); }
    catch (e) { setError(`Could not open workbook: ${String(e)}`); }
  }
  async function goHome() { try { await persist(); setHome(true); } catch (e) { setError(String(e)); } }

  async function importFile(name: string, bytes: Uint8Array, path?: string) {
    try {
      if (bytes.byteLength > 50 * 1024 * 1024) throw new Error('Choose a file smaller than 50 MB.');
      const ext = name.split('.').pop()?.toLowerCase();
      const { importXlsx, importCsv } = await import('./exchange');
      let next: BookRecord;
      if (ext === 'scloud') {
        next = parseNative(new TextDecoder().decode(bytes));
        const local = await store.get(next.id);
        if (local && JSON.stringify(local.workbook) !== JSON.stringify(next.workbook)) {
          next.id = crypto.randomUUID(); next.workbook.id = next.id; next.title += ' (imported copy)'; next.workbook.name = next.title;
          setNotice('Opened an imported copy to preserve the different workbook already on this device.');
        }
      }
      else if (ext === 'xlsx') {
        next = await importXlsx(bytes, name);
        await store.original(next.id, name, bytes);
        setNotice('XLSX imported. Macros, charts, pivots, and advanced Excel features may not transfer. The original is retained locally.');
      } else if (ext === 'csv') { next = importCsv(new TextDecoder().decode(bytes), name); setNotice('CSV imported as data. Formula-like text stays literal.'); }
      else throw new Error('Open .scloud, .xlsx, or .csv workbooks. Legacy .xls and macro-enabled .xlsm are not supported.');
      await load(next);
      if (path && ext === 'scloud') baseline.current = { path, bytes };
    } catch (e) { setError(`Import failed: ${String(e)}`); }
  }
  async function open() { try { if (!isNative()) fileInput.current?.click(); else { const file = await openNative(); if (file) await importFile(file.name, file.bytes, file.path); } } catch (e) { setError(String(e)); } }

  async function exportFile(format: 'scloud' | 'xlsx' | 'csv', saveAs = false) {
    try {
      const saved = await persist();
      const { exportXlsx, exportCsv } = await import('./exchange');
      let data: Uint8Array | string;
      if (format === 'scloud') data = JSON.stringify(saved);
      else if (format === 'xlsx') {
        if (!confirm('XLSX export preserves basic cells, formulas, styles, merges, dimensions, and freeze panes. Advanced validation, conditional formatting, charts, pivots, and Excel-specific features may be lost. Your full native workbook stays saved. Continue?')) return;
        data = await exportXlsx(saved);
      } else {
        const active = engine.current?.api.getActiveWorkbook()?.getActiveSheet().getSheetId() || saved.workbook.sheetOrder[0];
        data = exportCsv(saved.workbook.sheets[active]); setNotice('CSV exports the active sheet only. Use .scloud to preserve formulas, formatting, and other sheets.');
      }
      const filename = `${titleFile(saved.title)}.${format}`;
      if (isNative()) { const result = await writeNative(data, filename, format === 'scloud' && !saveAs ? baseline.current : undefined); if (!result) return; if (format === 'scloud') baseline.current = result; }
      else download(data, filename);
      setModal(''); setNotice(`Exported ${filename}`);
    } catch (e) { setError(`Export failed: ${String(e)}`); }
  }

  function goToCell() {
    try {
      const value = address.trim().toUpperCase();
      if (!/^\$?[A-Z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Z]{1,3}\$?[1-9]\d{0,6})?$/.test(value)) throw new Error('Enter a cell or range such as B4 or A1:D10.');
      const sheet = engine.current?.api.getActiveWorkbook()?.getActiveSheet();
      if (!sheet) return;
      const range = sheet.getRange(value); const bounds = range.getRange();
      if (bounds.endRow >= sheet.getMaxRows() || bounds.endColumn >= sheet.getMaxColumns()) throw new Error('That address is outside this sheet. Insert more rows or columns first.');
      range.activate(); setRaw(range.getFormula() || String(range.getValue() ?? '')); setDisplayed(range.getDisplayValue()); setAddress(value);
    } catch (e) { setError(String(e)); }
  }
  function applyCell() { try { const range = engine.current?.api.getActiveWorkbook()?.getActiveSheet().getActiveCell(); if (!range) return; range.setValue(raw); setDisplayed(range.getDisplayValue()); } catch (e) { setError(String(e)); } }
  function rename(title: string) { const next = { ...current.current, title }; current.current = next; setBook(next); engine.current?.api.getActiveWorkbook()?.setName(title || 'Untitled workbook'); changed(); }

  function previewChart() {
    const range = engine.current?.api.getActiveWorkbook()?.getActiveRange(); if (!range) return;
    const bounds = range.getRange();
    if ((bounds.endRow - bounds.startRow + 1) * (bounds.endColumn - bounds.startColumn + 1) > 10000) { setError('Select a chart range smaller than 10,000 cells.'); return; }
    const values = range.getRawValues();
    const points = values.map((row, i) => ({ label: String(row.length > 1 ? row[0] ?? i + 1 : i + 1), value: row[row.length - 1] })).filter((point): point is { label: string; value: number } => typeof point.value === 'number' && Number.isFinite(point.value));
    if (!points.length) { setError('Select a range containing numbers. The first column provides labels and the last column provides values.'); return; }
    setChartData(points.slice(0, 30)); setModal('chart');
  }
  async function print() {
    try {
      const sheet = engine.current?.api.getActiveWorkbook()?.getActiveSheet(); if (!sheet) return;
      const data = snapshot().workbook.sheets[sheet.getSheetId()]; const extent = usedExtent(data);
      if (extent.rows > 300 || extent.columns > 30) throw new Error('Print currently supports up to 300 used rows and 30 columns. Export XLSX for larger print jobs.');
      const rows = sheet.getRange(0, 0, Math.max(1, extent.rows), Math.max(1, extent.columns)).getDisplayValues();
      const iframe = document.createElement('iframe'); iframe.title = 'Sheet Cloud print layout'; iframe.style.cssText = 'position:fixed;width:0;height:0;border:0'; document.body.appendChild(iframe);
      const target = iframe.contentDocument!; target.open(); target.write(`<!doctype html><html><head><title>${escapeHtml(book.title)}</title><style>@page{size:A4 landscape;margin:12mm}body{font:10pt Arial;color:#243830}h1{font-size:18pt}table{border-collapse:collapse;width:100%}td{border:1px solid #d8e0dc;padding:6px;white-space:pre-wrap;overflow-wrap:anywhere}tr{break-inside:avoid}tr:first-child{font-weight:bold;background:#edf5f1}</style></head><body><h1>${escapeHtml(book.title)}</h1><p>${escapeHtml(sheet.getSheetName())}</p><table>${rows.map(row => `<tr>${row.map(value => `<td>${escapeHtml(value)}</td>`).join('')}</tr>`).join('')}</table></body></html>`); target.close(); await target.fonts.ready; iframe.contentWindow!.onafterprint = () => iframe.remove(); iframe.contentWindow!.print();
    } catch (e) { setError(String(e)); }
  }
  useEffect(() => {
    const handler = (event: KeyboardEvent) => { if (!(event.ctrlKey || event.metaKey)) return; const key = event.key.toLowerCase(); if (!['s', 'o', 'n', 'p'].includes(key)) return; event.preventDefault(); if (key === 's') void exportFile('scloud', event.shiftKey); if (key === 'o') void open(); if (key === 'n') void load(blankBook()); if (key === 'p' && !home) void print(); };
    addEventListener('keydown', handler, true); return () => removeEventListener('keydown', handler, true);
  });

  const visible = library.filter(item => item.title.toLowerCase().includes(search.toLowerCase()) && (filter !== 'Pinned' || pinned.includes(item.id)));
  const maxChart = Math.max(1, ...chartData.map(point => Math.abs(point.value)));
  return <div className="sheet-app">
    <input type="file" ref={fileInput} hidden accept=".scloud,.xlsx,.csv" onChange={async e => { const file = e.target.files?.[0]; if (file) await importFile(file.name, new Uint8Array(await file.arrayBuffer())); e.target.value = ''; }} />
    {home ? <><aside className="sidebar"><div className="brand"><span className="brand-mark"><img src="/logo.svg" alt="" /></span><span>sheet cloud<small>BY BLUECLOUD</small></span></div><button className="primary new" onClick={() => load(blankBook())}><Plus size={18} />New workbook</button><div className="sidebar-label">WORKSPACE</div>{['All workbooks', 'Recent', 'Pinned', 'Templates'].map((item, i) => { const Icon = [FileSpreadsheet, Clock3, Pin, BookOpen][i]; return <button key={item} className={`nav-item ${filter === item ? 'selected' : ''}`} onClick={() => setFilter(item)}><Icon size={17} />{item}{i === 0 && <span className="push">{library.length}</span>}</button>; })}<button className="nav-item" onClick={open}><FolderOpen size={17} />Open from computer<ArrowUpRight size={13} className="push" /></button><div className="sidebar-bottom"><div className="privacy-note"><ShieldCheck size={18} /><div><strong>Your numbers. Your space.</strong><p>Local first. Completely free.</p></div></div><button className="nav-item" onClick={() => setModal('settings')}><Settings2 size={17} />Settings & appearance</button><small className="version">Sheet Cloud · Development edition</small></div></aside>
      <main className="library"><div className="topbar"><span>Your workspace</span><div className="search"><Search size={16} /><input aria-label="Search workbooks" placeholder="Search your workbooks" value={search} onChange={e => setSearch(e.target.value)} /></div><button className="icon" aria-label="Appearance settings" onClick={() => setModal('settings')}><Sun size={19} /></button><div className="avatar">Y</div></div><div className="library-content"><div className="welcome"><div><div className="eyebrow">A FRESH PERSPECTIVE ON YOUR DATA</div><h1>A clearer view of your numbers.</h1><p>From everyday plans to your next big idea. Bring it all together.</p></div><div className="welcome-art"><BarChart3 size={58} strokeWidth={1.2} /></div></div><div className="section-heading"><h2>Start with a little structure</h2><button className="text-button" onClick={() => setFilter(filter === 'Templates' ? 'All workbooks' : 'Templates')}>{filter === 'Templates' ? 'Back to workbooks' : 'Explore templates'}<ArrowUpRight size={14} /></button></div><div className="templates">{templates.map((template, i) => <button key={template.id} className="template" onClick={() => load(fromTemplate(template.id))}><div className={`template-preview ${template.color}`}>{i === 0 ? <Plus size={30} strokeWidth={1} /> : <div className="mini-sheet"><span>{template.title.toUpperCase()}</span><div className="mini-subtitle" /><div className="mini-table">{Array.from({ length: 24 }, (_, n) => <i key={n} className={n < 4 ? 'table-head' : ''}>{n >= 4 && n % 4 !== 0 ? [1200, 365, 82, 1500][n % 4] : ''}</i>)}</div><div className="mini-bars">{[25, 38, 29, 48, 57, 68, 53].map((h, n) => <i key={n} style={{ height: `${h}%` }} />)}</div></div>}</div><strong>{template.title}</strong><small>{template.subtitle}</small></button>)}</div>
        {filter !== 'Templates' && <><div className="section-heading workbook-heading"><h2>{filter === 'Pinned' ? 'Pinned workbooks' : 'Your workbooks'}<span className="badge">{visible.length}</span></h2><span className="muted">Last edited<ChevronDown size={13} /></span></div>{visible.length ? <div className="book-list"><div className="book-list-heading"><span>NAME</span><span>LAST EDITED</span><span>SHEETS</span><span /></div>{(filter === 'Recent' ? visible.slice(0, 12) : visible).map(item => <div className="book-row" key={item.id}><button className="book-name" onClick={() => load(item)}><span className="file-icon"><FileSpreadsheet size={22} /></span><span><strong>{item.title || 'Untitled workbook'}</strong><small>Local workbook · .scloud</small></span></button><span className="muted">{new Date(item.updatedAt).toLocaleDateString(undefined, { month: 'short', day: 'numeric' })}</span><span className="muted">{item.workbook.sheetOrder.length}</span><div className="row-actions"><button aria-label={`Pin ${item.title}`} className={`icon ${pinned.includes(item.id) ? 'green-text' : ''}`} onClick={() => { const next = pinned.includes(item.id) ? pinned.filter(id => id !== item.id) : [...pinned, item.id]; setPinned(next); localStorage.setItem('sheetcloud-pinned', JSON.stringify(next)); }}><Pin size={15} /></button><button aria-label={`Delete ${item.title}`} className="icon" onClick={async () => { if (confirm(`Delete “${item.title}” from this device? Export a backup first if needed.`)) { await store.remove(item.id); await refreshLibrary(); } }}><Trash2 size={15} /></button></div></div>)}</div> : <div className="empty"><span className="empty-icon"><Grid2X2 size={26} /></span><h3>{search ? 'No matching workbooks' : filter === 'Pinned' ? 'Keep the important things close' : 'Every great plan starts with a cell.'}</h3><p>{search ? 'Try searching for another workbook name.' : 'Create a workbook or bring an existing one from your computer.'}</p><button className="secondary" onClick={open}><FolderOpen size={15} />Open a workbook</button></div>}</>}
        <footer className="workspace-footer"><ShieldCheck size={14} />No accounts. No subscriptions. Your data stays with you.<span className="push">Space to think. Room to grow.</span></footer></div></main></> : <>
      <header className="editor-header"><button className="editor-brand" aria-label="Back to workbooks" onClick={goHome}><span className="brand-mark"><img src="/logo.svg" alt="" /></span><span>sheet cloud</span></button><span className="divider" /><div className="title-block"><input aria-label="Workbook title" value={book.title} onChange={e => rename(e.target.value)} /><span><Check size={11} />{status === 'saved' ? 'Saved on this device' : status === 'saving' ? 'Saving…' : 'Unsaved changes'}</span></div><div className="editor-actions"><button className="icon" title="Undo" onClick={() => engine.current?.api.undo()}><Undo2 size={17} /></button><button className="icon" title="Redo" onClick={() => engine.current?.api.redo()}><Redo2 size={17} /></button><button className="secondary" onClick={previewChart}><BarChart3 size={16} />Chart preview</button><button className="secondary" onClick={() => setModal('history')}><History size={16} />History</button><button className="icon" title="Print active sheet" onClick={print}><Printer size={17} /></button><button className="icon" title="Settings and help" onClick={() => setModal('settings')}><Settings2 size={17} /></button><button className="primary" onClick={() => setModal('file')}><Download size={15} />File<ChevronDown size={13} /></button></div></header>
      <div className="cell-bar"><input aria-label="Cell address" className="cell-address" value={address} onChange={e => setAddress(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); goToCell(); } }} /><button className="icon" title="Go to cell or range" onClick={goToCell}><ArrowUpRight size={15} /></button><span className="formula-symbol">ƒx</span><input aria-label="Cell value or formula" placeholder="Enter a value or formula, then press Enter" value={raw} onChange={e => setRaw(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); applyCell(); } if (e.key === 'Escape') { const range = engine.current?.api.getActiveWorkbook()?.getActiveRange(); setRaw(range?.getFormula() || String(range?.getValue() ?? '')); } }} /><button className="icon" title="Apply cell value" onMouseDown={e => e.preventDefault()} onClick={applyCell}><Check size={17} /></button><span className="cell-result" aria-label="Calculated cell value">{displayed}</span></div>
      <div className="grid-host" ref={host} aria-label="Spreadsheet workspace" />{!engineReady && <div className="engine-loading"><Grid2X2 size={24} /><span>Preparing your workbook…</span></div>}
      <div className="editor-footer"><ShieldCheck size={13} /><span>Local workspace</span><button onClick={() => setModal('shortcuts')}>Keyboard shortcuts</button><span className="push">Sheet Cloud · Free development edition</span></div>
    </>}
    {modal === 'file' && <Modal title="Keep your work moving" close={() => setModal('')}><p className="note">Workbooks autosave locally. Export a native copy for the complete supported workbook state.</p><div className="file-quick"><button className="secondary" onClick={() => load(blankBook())}><Plus size={15} />New</button><button className="secondary" onClick={open}><FolderOpen size={15} />Open</button><button className="secondary" onClick={async () => { const copy = structuredClone(snapshot()); copy.id = crypto.randomUUID(); copy.workbook.id = copy.id; copy.title += ' (copy)'; copy.workbook.name = copy.title; await load(copy); }}><Copy size={15} />Duplicate</button></div><div className="export-list">{[['scloud', 'Sheet Cloud workbook', 'Full supported workbook · recommended'], ['xlsx', 'Excel workbook', 'Basic cells, formulas, styles, and sheets'], ['csv', 'CSV data', 'Active sheet values only']].map(([format, title, subtitle]) => <button key={format} onClick={() => exportFile(format as any, true)}><span className="file-icon"><FileSpreadsheet size={21} /></span><span><strong>{title}</strong><small>{subtitle}</small></span><span className="push extension">.{format}</span><Download size={16} /></button>)}<button onClick={async () => { try { const original = await store.getOriginal(book.id); if (!original) { setNotice('No imported XLSX original is stored for this workbook.'); return; } if (isNative()) await writeNative(original.data, original.name); else download(original.data, original.name); } catch (e) { setError(String(e)); } }}><span className="file-icon"><ShieldCheck size={21} /></span><span><strong>Retrieve imported original</strong><small>Unmodified XLSX backup, when available</small></span><Download size={16} className="push" /></button></div></Modal>}
    {modal === 'settings' && <Modal title="A workspace that feels like yours." close={() => setModal('')}><p className="note">A little less clutter. A little more clarity.</p><div className="eyebrow">APPEARANCE</div><div className="appearance-options">{(['light', 'dark', 'system'] as Appearance[]).map((value, i) => { const Icon = [Sun, Moon, Monitor][i]; return <button key={value} className={appearance === value ? 'active' : ''} onClick={() => setAppearance(value)}><Icon size={25} />{value[0].toUpperCase() + value.slice(1)}</button>; })}</div><div className="settings-note"><ShieldCheck size={20} /><div><strong>Your data stays on this device.</strong><p>No mandatory accounts, cloud upload, telemetry, or paid APIs. Export backups before clearing app data.</p></div></div><p className="note">Development edition: basic XLSX conversion, no macros, pivots, secure collaboration, or complete Excel fidelity. Chart preview is a local view of your selection; it is not embedded in XLSX exports.</p><button className="text-button" onClick={() => setModal('shortcuts')}>Keyboard shortcuts & formulas<ArrowUpRight size={14} /></button></Modal>}
    {modal === 'history' && <Modal title="Local version history" close={() => setModal('')}><p className="note">Save named workbook snapshots on this device.</p><div className="version-input"><input aria-label="Version name" placeholder="Name this version" value={versionName} onChange={e => setVersionName(e.target.value)} /><button className="primary" onClick={async () => { const saved = snapshot(); saved.versions = [...saved.versions, { id: crypto.randomUUID(), name: versionName.trim() || `Version ${saved.versions.length + 1}`, at: new Date().toISOString(), workbook: structuredClone(saved.workbook) }]; current.current = saved; sequence.current++; setBook(saved); await persist(); setVersionName(''); setNotice('Version saved locally.'); }}>Save version</button></div>{book.versions.length === 0 && <p className="note">No named versions yet. Your current workbook still autosaves.</p>}{book.versions.slice().reverse().map(version => <div className="version-row" key={version.id}><div><strong>{version.name}</strong><small>{new Date(version.at).toLocaleString()}</small></div><button className="secondary" onClick={async () => { if (!confirm(`Restore “${version.name}”? Your current workbook will be retained as a backup.`)) return; const next = snapshot(); next.versions.push({ id: crypto.randomUUID(), name: 'Before restore', at: new Date().toISOString(), workbook: structuredClone(next.workbook) }); next.workbook = structuredClone(version.workbook); await load(next); }}>Restore</button></div>)}</Modal>}
    {modal === 'chart' && <Modal title="A different view of your selection" close={() => setModal('')}><p className="note">First column: labels. Last column: values. Up to 30 rows. This preview is not embedded in the workbook.</p><svg viewBox="0 0 420 260" role="img" aria-label="Selected data bar chart"><line x1="40" y1="220" x2="410" y2="220" stroke="#ccd7d2" />{chartData.map((point, i) => { const width = 350 / chartData.length; const height = Math.abs(point.value) / maxChart * 170; return <g key={i}><rect x={45 + i * width} y={220 - height} width={Math.max(2, width - 8)} height={height} rx="3" fill={point.value < 0 ? '#cf6c62' : '#29967b'} /><text x={45 + i * width + (width - 8) / 2} y={210 - height} textAnchor="middle" fontSize="9" fill="currentColor">{point.value.toLocaleString()}</text><text x={45 + i * width} y="240" fontSize="8" fill="currentColor">{point.label.slice(0, 8)}</text></g>; })}</svg></Modal>}
    {modal === 'shortcuts' && <Modal title="From a cell to a bigger idea." close={() => setModal('')}><p className="note">Use Ctrl on Windows/Linux or Command on macOS. Grid shortcuts are supplied by the spreadsheet engine.</p><div className="shortcuts">{[['Ctrl / ⌘ N', 'New workbook'], ['Ctrl / ⌘ O', 'Open a workbook'], ['Ctrl / ⌘ S', 'Save / export native workbook'], ['Ctrl / ⌘ Shift S', 'Save As'], ['Ctrl / ⌘ P', 'Print active sheet'], ['Ctrl / ⌘ Z', 'Undo'], ['Ctrl / ⌘ C / V', 'Copy / paste cells'], ['Enter / F2', 'Edit cells in the grid']].map(([keys, title]) => <div key={keys}><span>{title}</span><kbd>{keys}</kbd></div>)}</div><p className="note formula-help">Try <code>=SUM(B2:B10)</code>, <code>=AVERAGE(C2:C10)</code>, <code>=IF(A1&gt;100,"High","Low")</code>, or <code>=B2*C2</code>. Formula behavior varies for unsupported Excel functions.</p></Modal>}
    {error && <div className="error" role="alert"><span>{error}</span><button aria-label="Dismiss error" onClick={() => setError('')}><X size={17} /></button></div>}
    {notice && <div className="toast" role="status"><Check size={16} /><span>{notice}</span><button aria-label="Dismiss notification" onClick={() => setNotice('')}><X size={15} /></button></div>}
  </div>;
}
