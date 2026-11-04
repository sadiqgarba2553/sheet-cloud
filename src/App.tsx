import { useCallback, useEffect, useRef, useState } from 'react';
import {
  Grid2X2,
  Plus,
  FolderOpen,
  Clock3,
  Pin,
  Search,
  ArrowUpRight,
  Download,
  Check,
  X,
  Sun,
  Moon,
  Monitor,
  History,
  ShieldCheck,
  Settings2,
  FileSpreadsheet,
  ChevronDown,
  BarChart3,
  Printer,
  BookOpen,
  Undo2,
  Redo2,
  Copy,
  Trash2,
  Layers,
  TableProperties,
  Calculator,
  SlidersHorizontal,
} from 'lucide-react';
import {
  blankBook,
  BookRecord,
  usedExtent,
  parseNative,
  SheetChart,
  ChartType,
  SheetPivot,
  Scenario,
  parseRange,
  colToLetters,
} from './model';
import { fromTemplate, templates } from './templates';
import { store } from './store';
import { download, isNative, openNative, writeNative } from './files';
import { Modal } from './Modal';
import { ChartView } from './ChartView';
import { computePivotTable, writePivotToSheet, runGoalSeek, textToColumns, removeDuplicates } from './analysis';
import type { SheetEngine } from './engine';

type Appearance = 'light' | 'dark' | 'system';
const titleFile = (title: string) => title.replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').trim() || 'Workbook';
const escapeHtml = (value: string) =>
  value.replace(
    /[&<>"']/g,
    c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]!)
  );

export default function App() {
  const [book, setBook] = useState(blankBook);
  const current = useRef(book);
  current.current = book;
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
  const [appearance, setAppearance] = useState<Appearance>(
    () => (localStorage.getItem('sheetcloud-appearance') as Appearance) || 'system'
  );
  const [pinned, setPinned] = useState<string[]>(() => {
    try {
      return JSON.parse(localStorage.getItem('sheetcloud-pinned') || '[]');
    } catch {
      return [];
    }
  });

  // Extended Analysis & Charting States
  const [showChartsDrawer, setShowChartsDrawer] = useState(false);
  const [chartTitle, setChartTitle] = useState('Overview Chart');
  const [chartType, setChartType] = useState<ChartType>('column');
  const [chartRange, setChartRange] = useState('A1:B10');
  const [editingChart, setEditingChart] = useState<SheetChart | null>(null);

  // Pivot Table State
  const [pivotSourceRange, setPivotSourceRange] = useState('');
  const [pivotRowField, setPivotRowField] = useState('');
  const [pivotValField, setPivotValField] = useState('');
  const [pivotAgg, setPivotAgg] = useState<'SUM' | 'COUNT' | 'AVERAGE' | 'MIN' | 'MAX'>('SUM');

  // What-If Goal Seek State
  const [targetCellAddr, setTargetCellAddr] = useState('B10');
  const [targetCellVal, setTargetCellVal] = useState('1000');
  const [changingCellAddr, setChangingCellAddr] = useState('B2');
  const [goalSeekResult, setGoalSeekResult] = useState<string>('');

  // Scenario Manager State
  const [scenarioName, setScenarioName] = useState('');

  // Data Tools State
  const [splitDelimiter, setSplitDelimiter] = useState(',');
  const [splitColIndex, setSplitColIndex] = useState('0');

  // Print Settings State
  const [printOrientation, setPrintOrientation] = useState<'landscape' | 'portrait'>('landscape');
  const [printFitToWidth, setPrintFitToWidth] = useState(true);
  const [printRepeatHeaders, setPrintRepeatHeaders] = useState(true);

  const host = useRef<HTMLDivElement>(null);
  const engine = useRef<SheetEngine | null>(null);
  const fileInput = useRef<HTMLInputElement>(null);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const sequence = useRef(0);
  const baseline = useRef<{ path: string; bytes: Uint8Array } | undefined>(undefined);
  const dark = appearance === 'dark' || (appearance === 'system' && matchMedia('(prefers-color-scheme: dark)').matches);

  const refreshLibrary = useCallback(async () => {
    setLibrary((await store.list()).sort((a, b) => b.updatedAt.localeCompare(a.updatedAt)));
  }, []);

  useEffect(() => {
    refreshLibrary()
      .catch(e => setError(String(e)))
      .finally(() => setLoaded(true));
  }, [refreshLibrary]);

  useEffect(() => {
    if (home && loaded) refreshLibrary().catch(e => setError(String(e)));
  }, [home, loaded, refreshLibrary]);

  useEffect(() => {
    if (!notice) return;
    const timeout = setTimeout(() => setNotice(''), 6000);
    return () => clearTimeout(timeout);
  }, [notice]);

  useEffect(() => {
    const media = matchMedia('(prefers-color-scheme: dark)');
    const apply = () => {
      const value = appearance === 'dark' || (appearance === 'system' && media.matches);
      document.documentElement.classList.toggle('dark', value);
      engine.current?.api.toggleDarkMode(value);
    };
    apply();
    media.addEventListener('change', apply);
    localStorage.setItem('sheetcloud-appearance', appearance);
    return () => media.removeEventListener('change', apply);
  }, [appearance]);

  const getActiveSheetId = useCallback(() => {
    return engine.current?.api.getActiveWorkbook()?.getActiveSheet().getSheetId() || current.current.workbook.sheetOrder[0];
  }, []);

  const snapshot = useCallback((): BookRecord => {
    const savedWorkbook = engine.current?.save() || current.current.workbook;
    // Retain top-level charts, pivots, scenarios
    savedWorkbook.charts = current.current.workbook.charts || [];
    savedWorkbook.pivots = current.current.workbook.pivots || [];
    savedWorkbook.scenarios = current.current.workbook.scenarios || [];
    savedWorkbook.dataQueries = current.current.workbook.dataQueries || [];
    return {
      ...current.current,
      updatedAt: new Date().toISOString(),
      workbook: savedWorkbook,
    };
  }, []);

  const persist = useCallback(async () => {
    if (timer.current) {
      clearTimeout(timer.current);
      timer.current = null;
    }
    const savedSequence = sequence.current;
    const next = snapshot();
    setStatus('saving');
    await store.save(next);
    if (savedSequence === sequence.current) {
      setStatus('saved');
      current.current = next;
      setBook(next);
    }
    return next;
  }, [snapshot]);

  const changed = useCallback(() => {
    if (!engine.current) return;
    sequence.current++;
    setStatus('dirty');
    if (timer.current) clearTimeout(timer.current);
    timer.current = setTimeout(() => {
      persist().catch(e => {
        setStatus('dirty');
        setError(`Autosave failed. Export a native backup: ${String(e)}`);
      });
    }, 800);
  }, [persist]);

  useEffect(() => {
    if (home || !host.current) return;
    let disposed = false;
    let controller: SheetEngine | undefined;
    setEngineReady(false);
    import('./engine')
      .then(({ createEngine }) => {
        if (disposed) return;
        controller = createEngine(
          host.current!,
          current.current,
          changed,
          (a, value, display) => {
            setDisplayed(display);
            if (document.activeElement?.getAttribute('aria-label') !== 'Cell address') setAddress(a);
            if (document.activeElement?.getAttribute('aria-label') !== 'Cell value or formula') setRaw(value);
          },
          dark
        );
        engine.current = controller;
        setEngineReady(true);
      })
      .catch(e => setError(`Could not start the spreadsheet engine: ${String(e)}`));
    return () => {
      disposed = true;
      if (timer.current) clearTimeout(timer.current);
      if (engine.current === controller) engine.current = null;
      controller?.dispose();
    };
  }, [home, session, changed]);

  useEffect(() => {
    const before = (e: BeforeUnloadEvent) => {
      if (status !== 'saved') {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    addEventListener('beforeunload', before);
    return () => removeEventListener('beforeunload', before);
  }, [status]);

  async function load(next: BookRecord) {
    try {
      if (!home && engine.current) await persist();
      baseline.current = undefined;
      current.current = next;
      setBook(next);
      sequence.current++;
      setStatus('dirty');
      setHome(false);
      setSession(n => n + 1);
      setModal('');
      await store.save(next);
      setStatus('saved');
    } catch (e) {
      setError(`Could not open workbook: ${String(e)}`);
    }
  }

  async function goHome() {
    try {
      await persist();
      setHome(true);
    } catch (e) {
      setError(String(e));
    }
  }

  async function importFile(name: string, bytes: Uint8Array, path?: string) {
    try {
      if (bytes.byteLength > 50 * 1024 * 1024) throw new Error('Choose a file smaller than 50 MB.');
      const ext = name.split('.').pop()?.toLowerCase();
      const { importXlsx, importCsv, importJson } = await import('./exchange');
      let next: BookRecord;
      if (ext === 'scloud') {
        next = parseNative(new TextDecoder().decode(bytes));
        const local = await store.get(next.id);
        if (local && JSON.stringify(local.workbook) !== JSON.stringify(next.workbook)) {
          next.id = crypto.randomUUID();
          next.workbook.id = next.id;
          next.title += ' (imported copy)';
          next.workbook.name = next.title;
          setNotice('Opened an imported copy to preserve the different workbook already on this device.');
        }
      } else if (ext === 'xlsx') {
        next = await importXlsx(bytes, name);
        await store.original(next.id, name, bytes);
        setNotice('XLSX imported. Macros, complex pivots, and external connections are retained locally.');
      } else if (ext === 'csv') {
        next = importCsv(new TextDecoder().decode(bytes), name);
        setNotice('CSV imported as data. Formula-like text stays literal.');
      } else if (ext === 'json') {
        next = importJson(new TextDecoder().decode(bytes), name);
        setNotice('JSON dataset imported as structured worksheet.');
      } else {
        throw new Error('Open .scloud, .xlsx, .csv, or .json workbooks.');
      }
      await load(next);
      if (path && ext === 'scloud') baseline.current = { path, bytes };
    } catch (e) {
      setError(`Import failed: ${String(e)}`);
    }
  }

  async function open() {
    try {
      if (!isNative()) fileInput.current?.click();
      else {
        const file = await openNative();
        if (file) await importFile(file.name, file.bytes, file.path);
      }
    } catch (e) {
      setError(String(e));
    }
  }

  async function exportFile(format: 'scloud' | 'xlsx' | 'csv' | 'json', saveAs = false) {
    try {
      const saved = await persist();
      const { exportXlsx, exportCsv, exportJson } = await import('./exchange');
      let data: Uint8Array | string;
      if (format === 'scloud') {
        data = JSON.stringify(saved, null, 2);
      } else if (format === 'xlsx') {
        data = await exportXlsx(saved);
      } else if (format === 'csv') {
        const active = getActiveSheetId();
        data = exportCsv(saved.workbook.sheets[active]);
        setNotice('CSV exports the active sheet only. Use .scloud to preserve all sheets and charts.');
      } else {
        const active = getActiveSheetId();
        data = exportJson(saved.workbook.sheets[active]);
        setNotice('JSON exported structured records from the active sheet.');
      }
      const filename = `${titleFile(saved.title)}.${format}`;
      if (isNative()) {
        const result = await writeNative(data, filename, format === 'scloud' && !saveAs ? baseline.current : undefined);
        if (!result) return;
        if (format === 'scloud') baseline.current = result;
      } else {
        download(data, filename);
      }
      setModal('');
      setNotice(`Exported ${filename}`);
    } catch (e) {
      setError(`Export failed: ${String(e)}`);
    }
  }

  function goToCell() {
    try {
      const value = address.trim().toUpperCase();
      if (!/^\$?[A-Z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Z]{1,3}\$?[1-9]\d{0,6})?$/.test(value)) {
        throw new Error('Enter a cell or range such as B4 or A1:D10.');
      }
      const sheet = engine.current?.api.getActiveWorkbook()?.getActiveSheet();
      if (!sheet) return;
      const range = sheet.getRange(value);
      const bounds = range.getRange();
      if (bounds.endRow >= sheet.getMaxRows() || bounds.endColumn >= sheet.getMaxColumns()) {
        throw new Error('That address is outside this sheet. Insert more rows or columns first.');
      }
      range.activate();
      setRaw(range.getFormula() || String(range.getValue() ?? ''));
      setDisplayed(range.getDisplayValue());
      setAddress(value);
    } catch (e) {
      setError(String(e));
    }
  }

  function applyCell() {
    try {
      const range = engine.current?.api.getActiveWorkbook()?.getActiveSheet().getActiveCell();
      if (!range) return;
      range.setValue(raw);
      setDisplayed(range.getDisplayValue());
      changed();
    } catch (e) {
      setError(String(e));
    }
  }

  function rename(title: string) {
    const next = { ...current.current, title };
    current.current = next;
    setBook(next);
    engine.current?.api.getActiveWorkbook()?.setName(title || 'Untitled workbook');
    changed();
  }

  // --- Embedded Charts Management ---
  const activeSheetId = getActiveSheetId();
  const activeSheet = book.workbook.sheets[activeSheetId] || book.workbook.sheets[book.workbook.sheetOrder[0]];
  const activeCharts: SheetChart[] = (book.workbook.charts || []).filter(c => c.sheetId === activeSheetId);

  function handleCreateChart() {
    const range = engine.current?.api.getActiveWorkbook()?.getActiveRange();
    const rangeStr = range ? range.getA1Notation() : chartRange;

    const newChart: SheetChart = {
      id: crypto.randomUUID(),
      sheetId: activeSheetId,
      title: chartTitle.trim() || 'New Chart',
      chartType,
      range: rangeStr,
      x: 20,
      y: 20,
      width: 400,
      height: 250,
      showLegend: true,
      showGrid: true,
    };

    const next = snapshot();
    next.workbook.charts ||= [];
    next.workbook.charts.push(newChart);
    current.current = next;
    setBook(next);
    changed();
    setModal('');
    setShowChartsDrawer(true);
    setNotice(`Created chart for ${rangeStr}`);
  }

  function handleUpdateChart(updated: SheetChart) {
    const next = snapshot();
    next.workbook.charts = (next.workbook.charts || []).map(c => (c.id === updated.id ? updated : c));
    current.current = next;
    setBook(next);
    changed();
  }

  function handleDeleteChart(id: string) {
    const next = snapshot();
    next.workbook.charts = (next.workbook.charts || []).filter(c => c.id !== id);
    current.current = next;
    setBook(next);
    changed();
    setNotice('Chart deleted.');
  }

  // --- Pivot Table Builder ---
  function openPivotDialog() {
    const range = engine.current?.api.getActiveWorkbook()?.getActiveRange();
    const used = usedExtent(activeSheet);
    const fallback = `A1:${colToLetters(Math.max(0, used.columns - 1))}${Math.max(1, used.rows)}`;
    setPivotSourceRange(range ? range.getA1Notation() : fallback);
    setPivotRowField('');
    setPivotValField('');
    setModal('pivot');
  }

  function handleCreatePivot() {
    try {
      const sheet = activeSheet;
      const pivotSpec: SheetPivot = {
        id: crypto.randomUUID(),
        name: `Pivot_${Date.now()}`,
        sourceSheetId: activeSheetId,
        sourceRange: pivotSourceRange.trim() || 'A1:D10',
        targetSheetId: '',
        targetCell: 'A1',
        rowFields: pivotRowField.trim() ? [pivotRowField.trim()] : [],
        colFields: [],
        valueFields: pivotValField.trim() ? [{ field: pivotValField.trim(), agg: pivotAgg }] : [],
      };

      const result = computePivotTable(sheet, pivotSpec);

      // Create new worksheet for the Pivot
      const newSheetId = crypto.randomUUID();
      const newSheetName = `Pivot ${book.workbook.sheetOrder.length}`;
      const newSheet = {
        id: newSheetId,
        name: newSheetName,
        rowCount: Math.max(100, result.rows.length + 20),
        columnCount: Math.max(26, result.headers.length + 5),
        cellData: {},
        defaultColumnWidth: 120,
        defaultRowHeight: 26,
        charts: [],
        pivots: [pivotSpec],
      };

      writePivotToSheet(newSheet, result, 0, 0);

      const next = snapshot();
      next.workbook.sheetOrder.push(newSheetId);
      next.workbook.sheets[newSheetId] = newSheet;
      next.workbook.pivots ||= [];
      next.workbook.pivots.push(pivotSpec);

      current.current = next;
      setBook(next);
      changed();
      setModal('');
      setNotice(`Created pivot sheet "${newSheetName}" with ${result.rows.length} rows.`);
    } catch (e) {
      setError(`Pivot creation failed: ${String(e)}`);
    }
  }

  // --- What-If Analysis / Goal Seek ---
  function handleGoalSeek() {
    try {
      const targetValNum = parseFloat(targetCellVal);
      if (isNaN(targetValNum)) throw new Error('Enter a valid numerical target value.');

      const activeRange = engine.current?.api.getActiveWorkbook()?.getActiveSheet().getRange(changingCellAddr.toUpperCase());
      const formulaRange = engine.current?.api.getActiveWorkbook()?.getActiveSheet().getRange(targetCellAddr.toUpperCase());

      if (!activeRange || !formulaRange) throw new Error('Specify valid cell addresses for target and changing cells.');

      // Mock numerical evaluator using perturbation in grid
      const initialGuess = parseFloat(String(activeRange.getValue() ?? 0)) || 1;

      const evalFn = (x: number) => {
        activeRange.setValue(x);
        return parseFloat(String(formulaRange.getValue() ?? 0)) || 0;
      };

      const res = runGoalSeek(evalFn, targetValNum, initialGuess);
      activeRange.setValue(Math.round(res.value * 10000) / 10000);
      changed();

      setGoalSeekResult(
        res.success
          ? `Goal Seek converged to ${Math.round(res.value * 10000) / 10000} in ${res.iterations} iterations.`
          : `Goal Seek approximated ${Math.round(res.value * 10000) / 10000}. Check formula bounds.`
      );
    } catch (e) {
      setError(`Goal seek error: ${String(e)}`);
    }
  }

  // --- Data Tools: Text-to-Columns & Deduplication ---
  function handleTextToColumns() {
    try {
      const colIdx = parseInt(splitColIndex, 10) || 0;
      const sheet = activeSheet;
      const res = textToColumns(sheet, colIdx, splitDelimiter);
      changed();
      setModal('');
      setNotice(`Split ${res.rowsAffected} cells into ${res.colsAdded + 1} columns.`);
    } catch (e) {
      setError(`Text to columns failed: ${String(e)}`);
    }
  }

  function handleRemoveDuplicates() {
    try {
      const sheet = activeSheet;
      const used = usedExtent(sheet);
      const res = removeDuplicates(sheet, [0], 1, used.rows);
      changed();
      setModal('');
      setNotice(`Removed ${res.removedCount} duplicate rows. ${res.retainedCount} rows kept.`);
    } catch (e) {
      setError(`Remove duplicates failed: ${String(e)}`);
    }
  }

  // --- Advanced Print & PDF Output ---
  async function executePrint() {
    try {
      const sheet = engine.current?.api.getActiveWorkbook()?.getActiveSheet();
      if (!sheet) return;
      const data = snapshot().workbook.sheets[sheet.getSheetId()];
      const extent = usedExtent(data);

      const rows = sheet.getRange(0, 0, Math.max(1, extent.rows), Math.max(1, extent.columns)).getDisplayValues();
      const iframe = document.createElement('iframe');
      iframe.title = 'Sheet Cloud print layout';
      iframe.style.cssText = 'position:fixed;width:0;height:0;border:0';
      document.body.appendChild(iframe);

      const target = iframe.contentDocument!;
      target.open();
      target.write(`<!doctype html>
<html>
<head>
  <title>${escapeHtml(book.title)}</title>
  <style>
    @page {
      size: A4 ${printOrientation};
      margin: 10mm;
    }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
      font-size: 9pt;
      color: #1e293b;
      margin: 0;
      padding: 0;
    }
    .print-header {
      display: flex;
      justify-content: space-between;
      border-bottom: 2px solid #15836b;
      padding-bottom: 6px;
      margin-bottom: 12px;
    }
    h1 {
      font-size: 14pt;
      margin: 0;
      color: #15836b;
    }
    .subtitle {
      font-size: 9pt;
      color: #64748b;
    }
    table {
      border-collapse: collapse;
      width: ${printFitToWidth ? '100%' : 'auto'};
      font-size: 8.5pt;
    }
    td, th {
      border: 1px solid #cbd5e1;
      padding: 4px 6px;
      white-space: pre-wrap;
      overflow-wrap: anywhere;
    }
    ${
      printRepeatHeaders
        ? `thead { display: table-header-group; }
           tr { break-inside: avoid; }`
        : ''
    }
    th {
      font-weight: 600;
      background: #edf5f1;
      color: #15836b;
    }
  </style>
</head>
<body>
  <div class="print-header">
    <div>
      <h1>${escapeHtml(book.title)}</h1>
      <div class="subtitle">Sheet: ${escapeHtml(sheet.getSheetName())}</div>
    </div>
    <div class="subtitle">${new Date().toLocaleDateString()}</div>
  </div>
  <table>
    ${
      printRepeatHeaders && rows.length > 0
        ? `<thead><tr>${rows[0].map(val => `<th>${escapeHtml(val)}</th>`).join('')}</tr></thead>`
        : ''
    }
    <tbody>
      ${rows
        .slice(printRepeatHeaders ? 1 : 0)
        .map(row => `<tr>${row.map(val => `<td>${escapeHtml(val)}</td>`).join('')}</tr>`)
        .join('')}
    </tbody>
  </table>
</body>
</html>`);
      target.close();
      await target.fonts.ready;
      setModal('');
      iframe.contentWindow!.onafterprint = () => iframe.remove();
      iframe.contentWindow!.print();
    } catch (e) {
      setError(String(e));
    }
  }

  useEffect(() => {
    const handler = (event: KeyboardEvent) => {
      if (!(event.ctrlKey || event.metaKey)) return;
      const key = event.key.toLowerCase();
      if (!['s', 'o', 'n', 'p'].includes(key)) return;
      event.preventDefault();
      if (key === 's') void exportFile('scloud', event.shiftKey);
      if (key === 'o') void open();
      if (key === 'n') void load(blankBook());
      if (key === 'p' && !home) setModal('print');
    };
    addEventListener('keydown', handler, true);
    return () => removeEventListener('keydown', handler, true);
  });

  const visible = library.filter(
    item =>
      item.title.toLowerCase().includes(search.toLowerCase()) &&
      (filter !== 'Pinned' || pinned.includes(item.id))
  );

  return (
    <div className="sheet-app">
      <input
        type="file"
        ref={fileInput}
        hidden
        accept=".scloud,.xlsx,.csv,.json"
        onChange={async e => {
          const file = e.target.files?.[0];
          if (file) await importFile(file.name, new Uint8Array(await file.arrayBuffer()));
          e.target.value = '';
        }}
      />
      {home ? (
        <>
          <aside className="sidebar">
            <div className="brand">
              <span className="brand-mark">
                <img src="/logo.svg" alt="" />
              </span>
              <span>
                sheet cloud<small>BY BLUECLOUD</small>
              </span>
            </div>
            <button className="primary new" onClick={() => load(blankBook())}>
              <Plus size={18} />
              New workbook
            </button>
            <div className="sidebar-label">WORKSPACE</div>
            {['All workbooks', 'Recent', 'Pinned', 'Templates'].map((item, i) => {
              const Icon = [FileSpreadsheet, Clock3, Pin, BookOpen][i];
              return (
                <button
                  key={item}
                  className={`nav-item ${filter === item ? 'selected' : ''}`}
                  onClick={() => setFilter(item)}
                >
                  <Icon size={17} />
                  {item}
                  {i === 0 && <span className="push">{library.length}</span>}
                </button>
              );
            })}
            <button className="nav-item" onClick={open}>
              <FolderOpen size={17} />
              Open from computer
              <ArrowUpRight size={13} className="push" />
            </button>
            <div className="sidebar-bottom">
              <div className="privacy-note">
                <ShieldCheck size={18} />
                <div>
                  <strong>Your numbers. Your space.</strong>
                  <p>Local first. Completely free.</p>
                </div>
              </div>
              <button className="nav-item" onClick={() => setModal('settings')}>
                <Settings2 size={17} />
                Settings & appearance
              </button>
              <small className="version">Sheet Cloud · Development edition</small>
            </div>
          </aside>
          <main className="library">
            <div className="topbar">
              <span>Your workspace</span>
              <div className="search">
                <Search size={16} />
                <input
                  aria-label="Search workbooks"
                  placeholder="Search your workbooks"
                  value={search}
                  onChange={e => setSearch(e.target.value)}
                />
              </div>
              <button
                className="icon"
                aria-label="Appearance settings"
                onClick={() => setModal('settings')}
              >
                <Sun size={19} />
              </button>
              <div className="avatar">Y</div>
            </div>
            <div className="library-content">
              <div className="welcome">
                <div>
                  <div className="eyebrow">A FRESH PERSPECTIVE ON YOUR DATA</div>
                  <h1>A clearer view of your numbers.</h1>
                  <p>From everyday plans to complex analysis. Bring it all together.</p>
                </div>
                <div className="welcome-art">
                  <BarChart3 size={58} strokeWidth={1.2} />
                </div>
              </div>
              <div className="section-heading">
                <h2>Start with a little structure</h2>
                <button
                  className="text-button"
                  onClick={() => setFilter(filter === 'Templates' ? 'All workbooks' : 'Templates')}
                >
                  {filter === 'Templates' ? 'Back to workbooks' : 'Explore templates'}
                  <ArrowUpRight size={14} />
                </button>
              </div>
              <div className="templates">
                {templates.map((template, i) => (
                  <button
                    key={template.id}
                    className="template"
                    onClick={() => load(fromTemplate(template.id))}
                  >
                    <div className={`template-preview ${template.color}`}>
                      {i === 0 ? (
                        <Plus size={30} strokeWidth={1} />
                      ) : (
                        <div className="mini-sheet">
                          <span>{template.title.toUpperCase()}</span>
                          <div className="mini-subtitle" />
                          <div className="mini-table">
                            {Array.from({ length: 24 }, (_, n) => (
                              <i key={n} className={n < 4 ? 'table-head' : ''}>
                                {n >= 4 && n % 4 !== 0 ? [1200, 365, 82, 1500][n % 4] : ''}
                              </i>
                            ))}
                          </div>
                          <div className="mini-bars">
                            {[25, 38, 29, 48, 57, 68, 53].map((h, n) => (
                              <i key={n} style={{ height: `${h}%` }} />
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                    <strong>{template.title}</strong>
                    <small>{template.subtitle}</small>
                  </button>
                ))}
              </div>
              {filter !== 'Templates' && (
                <>
                  <div className="section-heading workbook-heading">
                    <h2>
                      {filter === 'Pinned' ? 'Pinned workbooks' : 'Your workbooks'}
                      <span className="badge">{visible.length}</span>
                    </h2>
                    <span className="muted">
                      Last edited
                      <ChevronDown size={13} />
                    </span>
                  </div>
                  {visible.length ? (
                    <div className="book-list">
                      <div className="book-list-heading">
                        <span>NAME</span>
                        <span>LAST EDITED</span>
                        <span>SHEETS</span>
                        <span />
                      </div>
                      {(filter === 'Recent' ? visible.slice(0, 12) : visible).map(item => (
                        <div className="book-row" key={item.id}>
                          <button className="book-name" onClick={() => load(item)}>
                            <span className="file-icon">
                              <FileSpreadsheet size={22} />
                            </span>
                            <span>
                              <strong>{item.title || 'Untitled workbook'}</strong>
                              <small>Local workbook · .scloud</small>
                            </span>
                          </button>
                          <span className="muted">
                            {new Date(item.updatedAt).toLocaleDateString(undefined, {
                              month: 'short',
                              day: 'numeric',
                            })}
                          </span>
                          <span className="muted">{item.workbook.sheetOrder.length}</span>
                          <div className="row-actions">
                            <button
                              aria-label={`Pin ${item.title}`}
                              className={`icon ${pinned.includes(item.id) ? 'green-text' : ''}`}
                              onClick={() => {
                                const next = pinned.includes(item.id)
                                  ? pinned.filter(id => id !== item.id)
                                  : [...pinned, item.id];
                                setPinned(next);
                                localStorage.setItem('sheetcloud-pinned', JSON.stringify(next));
                              }}
                            >
                              <Pin size={15} />
                            </button>
                            <button
                              aria-label={`Delete ${item.title}`}
                              className="icon"
                              onClick={async () => {
                                if (
                                  confirm(
                                    `Delete “${item.title}” from this device? Export a backup first if needed.`
                                  )
                                ) {
                                  await store.remove(item.id);
                                  await refreshLibrary();
                                }
                              }}
                            >
                              <Trash2 size={15} />
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div className="empty">
                      <span className="empty-icon">
                        <Grid2X2 size={26} />
                      </span>
                      <h3>
                        {search
                          ? 'No matching workbooks'
                          : filter === 'Pinned'
                          ? 'Keep the important things close'
                          : 'Every great plan starts with a cell.'}
                      </h3>
                      <p>
                        {search
                          ? 'Try searching for another workbook name.'
                          : 'Create a workbook or bring an existing one from your computer.'}
                      </p>
                      <button className="secondary" onClick={open}>
                        <FolderOpen size={15} />
                        Open a workbook
                      </button>
                    </div>
                  )}
                </>
              )}
              <footer className="workspace-footer">
                <ShieldCheck size={14} />
                No accounts. No subscriptions. Your data stays with you.
                <span className="push">Space to think. Room to grow.</span>
              </footer>
            </div>
          </main>
        </>
      ) : (
        <>
          <header className="editor-header">
            <button className="editor-brand" aria-label="Back to workbooks" onClick={goHome}>
              <span className="brand-mark">
                <img src="/logo.svg" alt="" />
              </span>
              <span>sheet cloud</span>
            </button>
            <span className="divider" />
            <div className="title-block">
              <input aria-label="Workbook title" value={book.title} onChange={e => rename(e.target.value)} />
              <span>
                <Check size={11} />
                {status === 'saved' ? 'Saved on this device' : status === 'saving' ? 'Saving…' : 'Unsaved changes'}
              </span>
            </div>
            <div className="editor-actions">
              <button className="icon" title="Undo" onClick={() => engine.current?.api.undo()}>
                <Undo2 size={17} />
              </button>
              <button className="icon" title="Redo" onClick={() => engine.current?.api.redo()}>
                <Redo2 size={17} />
              </button>
              <button
                className={`secondary ${showChartsDrawer ? 'green-text' : ''}`}
                onClick={() => setShowChartsDrawer(!showChartsDrawer)}
                title="Embedded Charts"
              >
                <BarChart3 size={16} />
                Charts {activeCharts.length > 0 && `(${activeCharts.length})`}
              </button>
              <button className="secondary" onClick={openPivotDialog} title="Pivot Tables">
                <TableProperties size={16} />
                Pivot Table
              </button>
              <button className="secondary" onClick={() => setModal('whatif')} title="What-If Goal Seek">
                <Calculator size={16} />
                What-If
              </button>
              <button className="secondary" onClick={() => setModal('datatools')} title="Data Tools">
                <SlidersHorizontal size={16} />
                Data Tools
              </button>
              <button className="secondary" onClick={() => setModal('history')} title="Version History">
                <History size={16} />
                History
              </button>
              <button className="icon" title="Print options" onClick={() => setModal('print')}>
                <Printer size={17} />
              </button>
              <button className="icon" title="Settings and help" onClick={() => setModal('settings')}>
                <Settings2 size={17} />
              </button>
              <button className="primary" onClick={() => setModal('file')}>
                <Download size={15} />
                File
                <ChevronDown size={13} />
              </button>
            </div>
          </header>
          <div className="cell-bar">
            <input
              aria-label="Cell address"
              className="cell-address"
              value={address}
              onChange={e => setAddress(e.target.value)}
              onKeyDown={e => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  goToCell();
                }
              }}
            />
            <button className="icon" title="Go to cell or range" onClick={goToCell}>
              <ArrowUpRight size={15} />
            </button>
            <span className="formula-symbol">ƒx</span>
            <input
              aria-label="Cell value or formula"
              placeholder="Enter a value or formula, then press Enter"
              value={raw}
              onChange={e => setRaw(e.target.value)}
              onKeyDown={e => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  applyCell();
                }
                if (e.key === 'Escape') {
                  const range = engine.current?.api.getActiveWorkbook()?.getActiveRange();
                  setRaw(range?.getFormula() || String(range?.getValue() ?? ''));
                }
              }}
            />
            <button
              className="icon"
              title="Apply cell value"
              onMouseDown={e => e.preventDefault()}
              onClick={applyCell}
            >
              <Check size={17} />
            </button>
            <span className="cell-result" aria-label="Calculated cell value">
              {displayed}
            </span>
          </div>

          <div style={{ position: 'relative', flex: 1, display: 'flex', overflow: 'hidden' }}>
            <div className="grid-host" ref={host} aria-label="Spreadsheet workspace" />

            {/* Embedded Charts Side Drawer */}
            {showChartsDrawer && (
              <aside className="charts-drawer">
                <div className="charts-drawer-header">
                  <span>Embedded Charts ({activeSheet.name})</span>
                  <div style={{ display: 'flex', gap: '8px' }}>
                    <button className="secondary" onClick={() => setModal('addchart')}>
                      <Plus size={14} /> Add Chart
                    </button>
                    <button className="icon" onClick={() => setShowChartsDrawer(false)}>
                      <X size={16} />
                    </button>
                  </div>
                </div>
                <div className="charts-drawer-body">
                  {activeCharts.length === 0 ? (
                    <div className="empty" style={{ minHeight: '180px' }}>
                      <BarChart3 size={28} className="muted" />
                      <p>No charts embedded on this sheet yet.</p>
                      <button className="secondary" onClick={() => setModal('addchart')}>
                        <Plus size={14} /> Create a chart
                      </button>
                    </div>
                  ) : (
                    activeCharts.map(c => (
                      <ChartView
                        key={c.id}
                        chart={c}
                        sheet={activeSheet}
                        onUpdate={handleUpdateChart}
                        onDelete={handleDeleteChart}
                        onEdit={chart => {
                          setEditingChart(chart);
                          setChartTitle(chart.title);
                          setChartType(chart.chartType);
                          setChartRange(chart.range);
                          setModal('editchart');
                        }}
                      />
                    ))
                  )}
                </div>
              </aside>
            )}
          </div>

          {!engineReady && (
            <div className="engine-loading">
              <Grid2X2 size={24} />
              <span>Preparing your workbook…</span>
            </div>
          )}

          <div className="editor-footer">
            <ShieldCheck size={13} />
            <span>Local workspace · Air-gapped & Private</span>
            <button onClick={() => setModal('shortcuts')}>Keyboard shortcuts</button>
            <span className="push">Sheet Cloud · Free development edition</span>
          </div>
        </>
      )}

      {/* File Menu Modal */}
      {modal === 'file' && (
        <Modal title="Keep your work moving" close={() => setModal('')}>
          <p className="note">
            Workbooks autosave locally. Export a native copy for the complete supported workbook state.
          </p>
          <div className="file-quick">
            <button className="secondary" onClick={() => load(blankBook())}>
              <Plus size={15} />
              New
            </button>
            <button className="secondary" onClick={open}>
              <FolderOpen size={15} />
              Open
            </button>
            <button
              className="secondary"
              onClick={async () => {
                const copy = structuredClone(snapshot());
                copy.id = crypto.randomUUID();
                copy.workbook.id = copy.id;
                copy.title += ' (copy)';
                copy.workbook.name = copy.title;
                await load(copy);
              }}
            >
              <Copy size={15} />
              Duplicate
            </button>
          </div>
          <div className="export-list">
            {[
              ['scloud', 'Sheet Cloud workbook', 'Full supported workbook, charts & pivots · recommended'],
              ['xlsx', 'Excel workbook', 'Cells, formulas, styles, merges, and sheets'],
              ['csv', 'CSV data', 'Active sheet values only'],
              ['json', 'JSON dataset', 'Structured records for developers & APIs'],
            ].map(([format, title, subtitle]) => (
              <button key={format} onClick={() => exportFile(format as any, true)}>
                <span className="file-icon">
                  <FileSpreadsheet size={21} />
                </span>
                <span>
                  <strong>{title}</strong>
                  <small>{subtitle}</small>
                </span>
                <span className="push extension">.{format}</span>
                <Download size={16} />
              </button>
            ))}
            <button
              onClick={async () => {
                try {
                  const original = await store.getOriginal(book.id);
                  if (!original) {
                    setNotice('No imported XLSX original is stored for this workbook.');
                    return;
                  }
                  if (isNative()) await writeNative(original.data, original.name);
                  else download(original.data, original.name);
                } catch (e) {
                  setError(String(e));
                }
              }}
            >
              <span className="file-icon">
                <ShieldCheck size={21} />
              </span>
              <span>
                <strong>Retrieve imported original</strong>
                <small>Unmodified XLSX backup, when available</small>
              </span>
              <Download size={16} className="push" />
            </button>
          </div>
        </Modal>
      )}

      {/* Add / Edit Chart Modal */}
      {(modal === 'addchart' || modal === 'editchart') && (
        <Modal
          title={modal === 'addchart' ? 'Embed New Chart' : 'Edit Chart'}
          close={() => {
            setModal('');
            setEditingChart(null);
          }}
        >
          <div className="form-group">
            <label>Chart Title</label>
            <input value={chartTitle} onChange={e => setChartTitle(e.target.value)} placeholder="e.g. Monthly Revenue" />
          </div>
          <div className="form-group">
            <label>Chart Family</label>
            <select value={chartType} onChange={e => setChartType(e.target.value as ChartType)}>
              <option value="column">Column (Vertical)</option>
              <option value="bar">Bar (Horizontal)</option>
              <option value="line">Line (Trend)</option>
              <option value="area">Area (Filled)</option>
              <option value="pie">Pie (Proportions)</option>
            </select>
          </div>
          <div className="form-group">
            <label>Data Range (A1 Notation)</label>
            <input value={chartRange} onChange={e => setChartRange(e.target.value)} placeholder="e.g. A1:C12" />
            <small className="muted">First row: column headers. First column: row labels.</small>
          </div>
          <div className="dialog-actions">
            <button className="secondary" onClick={() => setModal('')}>
              Cancel
            </button>
            <button
              className="primary"
              onClick={() => {
                if (modal === 'addchart') {
                  handleCreateChart();
                } else if (editingChart) {
                  handleUpdateChart({
                    ...editingChart,
                    title: chartTitle,
                    chartType,
                    range: chartRange,
                  });
                  setModal('');
                  setEditingChart(null);
                  setNotice('Chart updated.');
                }
              }}
            >
              {modal === 'addchart' ? 'Insert Chart' : 'Save Changes'}
            </button>
          </div>
        </Modal>
      )}

      {/* Pivot Table Builder Modal */}
      {modal === 'pivot' && (
        <Modal title="Create Pivot Table" close={() => setModal('')}>
          <p className="note">
            Summarize and analyze large datasets by grouping categories and calculating aggregations into a new dedicated worksheet.
          </p>
          <div className="form-group">
            <label>Source Range</label>
            <input
              value={pivotSourceRange}
              onChange={e => setPivotSourceRange(e.target.value)}
              placeholder="e.g. A1:E50"
            />
          </div>
          <div className="form-group">
            <label>Row Group Field (Category column header)</label>
            <input
              value={pivotRowField}
              onChange={e => setPivotRowField(e.target.value)}
              placeholder="e.g. Category, Region, or Department"
            />
          </div>
          <div className="form-row">
            <div className="form-group">
              <label>Value Field (Metric column header)</label>
              <input
                value={pivotValField}
                onChange={e => setPivotValField(e.target.value)}
                placeholder="e.g. Amount or Sales"
              />
            </div>
            <div className="form-group">
              <label>Aggregation</label>
              <select value={pivotAgg} onChange={e => setPivotAgg(e.target.value as any)}>
                <option value="SUM">SUM</option>
                <option value="COUNT">COUNT</option>
                <option value="AVERAGE">AVERAGE</option>
                <option value="MIN">MIN</option>
                <option value="MAX">MAX</option>
              </select>
            </div>
          </div>
          <div className="dialog-actions">
            <button className="secondary" onClick={() => setModal('')}>
              Cancel
            </button>
            <button className="primary" onClick={handleCreatePivot}>
              Create Pivot Sheet
            </button>
          </div>
        </Modal>
      )}

      {/* What-If Goal Seek Modal */}
      {modal === 'whatif' && (
        <Modal title="What-If Analysis & Goal Seek" close={() => setModal('')}>
          <p className="note">
            Find the right input for the result you want using numerical root-finding.
          </p>
          <div className="form-group">
            <label>Set Cell (Formula)</label>
            <input
              value={targetCellAddr}
              onChange={e => setTargetCellAddr(e.target.value)}
              placeholder="e.g. B10"
            />
          </div>
          <div className="form-group">
            <label>To Value</label>
            <input
              value={targetCellVal}
              onChange={e => setTargetCellVal(e.target.value)}
              placeholder="e.g. 5000"
            />
          </div>
          <div className="form-group">
            <label>By Changing Cell</label>
            <input
              value={changingCellAddr}
              onChange={e => setChangingCellAddr(e.target.value)}
              placeholder="e.g. B2"
            />
          </div>
          {goalSeekResult && (
            <div className="results-box">
              <strong>{goalSeekResult}</strong>
            </div>
          )}
          <div className="dialog-actions">
            <button className="secondary" onClick={() => setModal('')}>
              Close
            </button>
            <button className="primary" onClick={handleGoalSeek}>
              Solve
            </button>
          </div>
        </Modal>
      )}

      {/* Data Tools Modal */}
      {modal === 'datatools' && (
        <Modal title="Data Transformation Tools" close={() => setModal('')}>
          <p className="note">Clean, reshape, and normalize your tabular data locally.</p>
          <div style={{ borderBottom: '1px solid var(--line)', paddingBottom: '16px', marginBottom: '16px' }}>
            <h3 style={{ fontSize: '13px', marginBottom: '8px' }}>Text to Columns</h3>
            <p className="note">Split concatenated text strings into distinct adjacent columns.</p>
            <div className="form-row">
              <div className="form-group">
                <label>Column Index (0 = A, 1 = B)</label>
                <input
                  type="number"
                  value={splitColIndex}
                  onChange={e => setSplitColIndex(e.target.value)}
                />
              </div>
              <div className="form-group">
                <label>Delimiter</label>
                <input
                  value={splitDelimiter}
                  onChange={e => setSplitDelimiter(e.target.value)}
                  placeholder="e.g. comma or ;"
                />
              </div>
            </div>
            <button className="secondary" onClick={handleTextToColumns}>
              Split Selected Column
            </button>
          </div>

          <div>
            <h3 style={{ fontSize: '13px', marginBottom: '8px' }}>Remove Duplicates</h3>
            <p className="note">Scan worksheet rows and eliminate redundant record entries.</p>
            <button className="secondary" onClick={handleRemoveDuplicates}>
              Remove Duplicates in Column A
            </button>
          </div>
        </Modal>
      )}

      {/* Advanced Print Options Modal */}
      {modal === 'print' && (
        <Modal title="Page Setup & Print" close={() => setModal('')}>
          <p className="note">Format clean, professional paged output from your active sheet.</p>
          <div className="form-group">
            <label>Page Orientation</label>
            <div className="pill-group">
              <button
                className={`pill-btn ${printOrientation === 'landscape' ? 'active' : ''}`}
                onClick={() => setPrintOrientation('landscape')}
              >
                Landscape
              </button>
              <button
                className={`pill-btn ${printOrientation === 'portrait' ? 'active' : ''}`}
                onClick={() => setPrintOrientation('portrait')}
              >
                Portrait
              </button>
            </div>
          </div>
          <div className="form-group">
            <label>Scaling & Pagination</label>
            <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', textTransform: 'none' }}>
              <input
                type="checkbox"
                checked={printFitToWidth}
                onChange={e => setPrintFitToWidth(e.target.checked)}
              />
              Fit table to single page width
            </label>
            <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', textTransform: 'none', marginTop: '6px' }}>
              <input
                type="checkbox"
                checked={printRepeatHeaders}
                onChange={e => setPrintRepeatHeaders(e.target.checked)}
              />
              Repeat header row across multiple printed pages
            </label>
          </div>
          <div className="dialog-actions">
            <button className="secondary" onClick={() => setModal('')}>
              Cancel
            </button>
            <button className="primary" onClick={executePrint}>
              <Printer size={15} /> Print Now
            </button>
          </div>
        </Modal>
      )}

      {/* Settings Modal */}
      {modal === 'settings' && (
        <Modal title="A workspace that feels like yours." close={() => setModal('')}>
          <p className="note">A little less clutter. A little more clarity.</p>
          <div className="eyebrow">APPEARANCE</div>
          <div className="appearance-options">
            {(['light', 'dark', 'system'] as Appearance[]).map((value, i) => {
              const Icon = [Sun, Moon, Monitor][i];
              return (
                <button
                  key={value}
                  className={appearance === value ? 'active' : ''}
                  onClick={() => setAppearance(value)}
                >
                  <Icon size={25} />
                  {value[0].toUpperCase() + value.slice(1)}
                </button>
              );
            })}
          </div>
          <div className="settings-note">
            <ShieldCheck size={20} />
            <div>
              <strong>Your data stays on this device.</strong>
              <p>
                No mandatory accounts, cloud upload, telemetry, or paid APIs. Export backups before clearing app data.
              </p>
            </div>
          </div>
          <p className="note">
            Sheet Cloud features local embedded charts, pivot tables, what-if goal seek solver, data transformation, and lossless version history.
          </p>
          <button className="text-button" onClick={() => setModal('shortcuts')}>
            Keyboard shortcuts & formulas
            <ArrowUpRight size={14} />
          </button>
        </Modal>
      )}

      {/* History Modal */}
      {modal === 'history' && (
        <Modal title="Local version history" close={() => setModal('')}>
          <p className="note">Save named workbook snapshots on this device.</p>
          <div className="version-input">
            <input
              aria-label="Version name"
              placeholder="Name this version"
              value={versionName}
              onChange={e => setVersionName(e.target.value)}
            />
            <button
              className="primary"
              onClick={async () => {
                const saved = snapshot();
                saved.versions = [
                  ...saved.versions,
                  {
                    id: crypto.randomUUID(),
                    name: versionName.trim() || `Version ${saved.versions.length + 1}`,
                    at: new Date().toISOString(),
                    workbook: structuredClone(saved.workbook),
                  },
                ];
                current.current = saved;
                sequence.current++;
                setBook(saved);
                await persist();
                setVersionName('');
                setNotice('Version saved locally.');
              }}
            >
              Save version
            </button>
          </div>
          {book.versions.length === 0 && (
            <p className="note">No named versions yet. Your current workbook still autosaves.</p>
          )}
          {book.versions
            .slice()
            .reverse()
            .map(version => (
              <div className="version-row" key={version.id}>
                <div>
                  <strong>{version.name}</strong>
                  <small>{new Date(version.at).toLocaleString()}</small>
                </div>
                <button
                  className="secondary"
                  onClick={async () => {
                    if (
                      !confirm(
                        `Restore “${version.name}”? Your current workbook will be retained as a backup.`
                      )
                    )
                      return;
                    const next = snapshot();
                    next.versions.push({
                      id: crypto.randomUUID(),
                      name: 'Before restore',
                      at: new Date().toISOString(),
                      workbook: structuredClone(next.workbook),
                    });
                    next.workbook = structuredClone(version.workbook);
                    await load(next);
                  }}
                >
                  Restore
                </button>
              </div>
            ))}
        </Modal>
      )}

      {/* Keyboard Shortcuts Modal */}
      {modal === 'shortcuts' && (
        <Modal title="From a cell to a bigger idea." close={() => setModal('')}>
          <p className="note">
            Use Ctrl on Windows/Linux or Command on macOS. Grid shortcuts are supplied by the spreadsheet engine.
          </p>
          <div className="shortcuts">
            {[
              ['Ctrl / ⌘ N', 'New workbook'],
              ['Ctrl / ⌘ O', 'Open a workbook'],
              ['Ctrl / ⌘ S', 'Save / export native workbook'],
              ['Ctrl / ⌘ Shift S', 'Save As'],
              ['Ctrl / ⌘ P', 'Print active sheet with pagination'],
              ['Ctrl / ⌘ Z', 'Undo'],
              ['Ctrl / ⌘ C / V', 'Copy / paste cells'],
              ['Enter / F2', 'Edit cells in the grid'],
            ].map(([keys, title]) => (
              <div key={keys}>
                <span>{title}</span>
                <kbd>{keys}</kbd>
              </div>
            ))}
          </div>
          <p className="note formula-help">
            Try <code>=SUM(B2:B10)</code>, <code>=AVERAGE(C2:C10)</code>,{' '}
            <code>=IF(A1&gt;100,"High","Low")</code>, or <code>=B2*C2</code>.
          </p>
        </Modal>
      )}

      {error && (
        <div className="error" role="alert">
          <span>{error}</span>
          <button aria-label="Dismiss error" onClick={() => setError('')}>
            <X size={17} />
          </button>
        </div>
      )}
      {notice && (
        <div className="toast" role="status">
          <Check size={16} />
          <span>{notice}</span>
          <button aria-label="Dismiss notification" onClick={() => setNotice('')}>
            <X size={15} />
          </button>
        </div>
      )}
    </div>
  );
}
