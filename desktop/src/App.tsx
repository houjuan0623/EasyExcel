// Copyright (c) 2026 AbeLin · MIT License
import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import {
  ArrowRight, ArrowUpRight, Check, CheckCheck, ChevronDown, ChevronRight,
  CircleCheck, CircleHelp, Clock3, Columns3, FileSpreadsheet, Folder, FolderOpen,
  GitBranch, History, LayoutGrid, ListFilter, LoaderCircle, LockKeyhole, Maximize2,
  Merge, Minus, Play, RefreshCw, Search, Settings2, ShieldCheck,
  Sparkles, Square, Table2, Upload, X, AlertCircle, Eye, Zap, Code2,
} from 'lucide-react';
import type { GroupPreview, HistoryItem, JobResult, Mode, Preview, SheetChoice, SheetInfo, SourceInfo } from '../shared/types';
import SheetPreview from './SheetPreview';
import CodeWorkspace from './CodeWorkspace';
import appIcon from '../assets/app.png';

const api = window.desktop;
const keyOf = (path: string, sheet: string) => JSON.stringify([path, sheet]);
const number = (value: number) => new Intl.NumberFormat('zh-CN').format(value);
const message = (error: unknown) => error instanceof Error ? error.message.replace(/^Error invoking remote method '[^']+': (?:Error: )?/, '') : String(error);
const size = (bytes: number) => bytes > 1048576 ? `${(bytes / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(bytes / 1024))} KB`;

function IconButton({ title, children, onClick, disabled = false, className = '' }: {
  title: string; children: ReactNode; onClick: () => void; disabled?: boolean; className?: string;
}) {
  return <button type="button" className={`icon-button ${className}`} aria-label={title} title={title} onClick={onClick} disabled={disabled}>{children}</button>;
}

function Toggle({ label, detail, checked, onChange, disabled = false }: {
  label: string; detail?: string; checked: boolean; onChange: (value: boolean) => void; disabled?: boolean;
}) {
  return <label className={`toggle-row ${disabled ? 'muted' : ''}`}>
    <span><strong>{label}</strong>{detail && <small>{detail}</small>}</span>
    <input className="toggle" type="checkbox" role="switch" checked={checked} onChange={(e) => onChange(e.target.checked)} disabled={disabled} />
  </label>;
}

function CardTitle({ step, title, children }: { step?: string; title: string; children?: ReactNode }) {
  return <div className="card-title"><h2>{step && <span className="step-number">{step}</span>}{title}</h2>{children}</div>;
}

function ResultBanner({ result, onOpen }: { result: JobResult; onOpen: () => void }) {
  const issues = Number(result.summary.failed_files || 0) + Number(result.summary.failed_saves || 0) + Number(result.summary.skipped_sheets || 0) + Number(result.summary.skipped || 0);
  const empty = Number(result.summary.files ?? 1) === 0;
  const attention = result.stopped || issues > 0 || empty;
  return <div className={`result-banner ${attention ? 'attention' : ''}`} role="status">
    {attention ? <AlertCircle size={24} /> : <CircleCheck size={24} />}
    <div><strong>{result.stopped ? '任务已停止' : empty ? '未生成拆分文件，请查看处理详情' : issues ? '处理已完成，有部分内容需要核对' : '处理完成，结果已保存'}</strong>
      <p>{result.stopped ? '已生成的部分结果可在输出目录查看。' : `${number(Number(result.summary.rows || 0))} 行数据已处理${result.summary.groups != null ? ` · ${result.summary.groups} 个分组` : ''}${issues ? ` · ${issues} 项跳过或失败` : ''}`}</p>
      {!!result.summary.bundle_file && <p>总工作簿已生成 · {Number(result.summary.bundle_sheets)} 个 Sheet · 位于结果目录的「总工作簿」文件夹</p>}
      {result.summary.mode === 'sheet_bundle' && !result.stopped && <p>已整理为 1 个工作簿 · {Number(result.summary.sheets)} 个独立 Sheet</p>}
    </div>{result.output && <button className="button small" onClick={onOpen}><FolderOpen size={15} />查看结果<ArrowUpRight size={14} /></button>}
  </div>;
}

export default function App() {
  const [page, setPage] = useState<'split' | 'merge' | 'code' | 'history'>('split');
  const mode: Mode = page === 'merge' ? 'merge' : 'split';
  const [help, setHelp] = useState(false);
  const [engine, setEngine] = useState<'checking' | 'ready' | 'error'>('checking');
  const [engineError, setEngineError] = useState('');
  const [inputPath, setInputPath] = useState('');
  const [source, setSource] = useState<SourceInfo | null>(null);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [activeSheet, setActiveSheet] = useState<{ path: string; name: string } | null>(null);
  const [preview, setPreview] = useState<Preview | null>(null);
  const [previewBusy, setPreviewBusy] = useState(false);
  const [previewError, setPreviewError] = useState('');
  const [reading, setReading] = useState(false);
  const [refresh, setRefresh] = useState(0);
  const [search, setSearch] = useState('');
  const [dragging, setDragging] = useState(false);
  const [headerRow, setHeaderRow] = useState<number | null>(null);
  const [column, setColumn] = useState('');
  const [groupData, setGroupData] = useState<GroupPreview | null>(null);
  const [groupError, setGroupError] = useState('');
  const [groupKey, setGroupKey] = useState('');
  const [chosenValues, setChosenValues] = useState<string[] | null>(null);
  const [groupSearch, setGroupSearch] = useState('');
  const [output, setOutput] = useState('');
  const [preserveFormat, setPreserveFormat] = useState(true);
  const [keepFormulas, setKeepFormulas] = useState(false);
  const [makeZip, setMakeZip] = useState(true);
  const [mergeAcrossFiles, setMergeAcrossFiles] = useState(false);
  const [bundleSheets, setBundleSheets] = useState(false);
  const [mergeMode, setMergeMode] = useState<'rows' | 'sheets'>('rows');
  const [advanced, setAdvanced] = useState(false);
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const [codeBusy, setCodeBusy] = useState(false);
  const [cancelling, setCancelling] = useState(false);
  const [progress, setProgress] = useState(0);
  const [logs, setLogs] = useState<string[]>([]);
  const [showLogs, setShowLogs] = useState(false);
  const [result, setResult] = useState<JobResult | null>(null);
  const [history, setHistory] = useState<HistoryItem[]>([]);
  const jobId = useRef<string | null>(null);
  const cancelRequested = useRef(false);
  const previousSource = useRef('');
  const logEnd = useRef<HTMLDivElement>(null);

  const checkEngine = useCallback(() => {
    if (!api) { setEngine('error'); setEngineError('请使用 Electron 桌面窗口打开，以选择和处理本地文件。'); return; }
    setEngine('checking');
    api.health().then(() => { setEngine('ready'); setEngineError(''); }).catch((e) => { setEngine('error'); setEngineError(message(e)); });
  }, []);
  useEffect(checkEngine, [checkEngine]);
  useEffect(() => {
    if (!api) return;
    void api.history().then(setHistory).catch(() => {});
    return api.onJob((event) => {
      if (jobId.current && event.jobId !== jobId.current) return;
      if (event.type === 'log') setLogs((old) => [...old.slice(-499), event.message || '']);
      if (event.type === 'progress') setProgress(event.value || 0);
      if (event.type === 'result' || event.type === 'error') {
        setBusy(false); setCancelling(false); jobId.current = null;
        if (event.type === 'result' && event.result) {
          setResult(event.result);
          if (!event.result.stopped) setProgress(1);
          void api.history().then(setHistory);
        } else { setError(event.message || '处理失败'); setShowLogs(true); }
      }
    });
  }, []);
  useEffect(() => { if (showLogs) logEnd.current?.scrollIntoView({ block: 'nearest' }); }, [logs, showLogs]);

  useEffect(() => {
    if (!inputPath || !api) return;
    let current = true;
    setReading(true); setError(''); setResult(null);
    const timer = setTimeout(() => {
      api.inspect(inputPath, headerRow).then((data) => {
        if (!current) return;
        const changed = previousSource.current !== data.path;
        const eligible = data.workbooks.flatMap((book) => book.sheets.filter((sheet) => sheet.selectable).map((sheet) => keyOf(book.path, sheet.name)));
        setSelected((old) => changed ? new Set(data.workbooks.flatMap((book) => {
          const first = book.sheets.find((sheet) => sheet.selectable);
          return first ? [keyOf(book.path, first.name)] : [];
        })) : new Set([...old].filter((key) => eligible.includes(key))));
        const visible = data.workbooks.flatMap((book) => book.sheets.filter((sheet) => sheet.state === 'visible').map((sheet) => keyOf(book.path, sheet.name)));
        const book = data.workbooks.find((b) => b.sheets.some((sheet) => sheet.selectable)) || data.workbooks.find((b) => b.sheets.some((sheet) => sheet.state === 'visible'));
        const sheet = book?.sheets.find((s) => s.selectable) || book?.sheets.find((s) => s.state === 'visible');
        setActiveSheet((old) => !changed && old && visible.includes(keyOf(old.path, old.name)) ? old : book && sheet ? { path: book.path, name: sheet.name } : null);
        if (changed) { setOutput(data.suggestedOutput); setSearch(''); setColumn(''); }
        previousSource.current = data.path;
        setSource(data);
      }).catch((e) => { if (current) { setError(message(e)); setSource(null); setSelected(new Set()); setActiveSheet(null); } })
        .finally(() => { if (current) setReading(false); });
    }, 160);
    return () => { current = false; clearTimeout(timer); };
  }, [inputPath, headerRow, refresh]);

  const allSheets = useMemo(() => source?.workbooks.flatMap((book) => book.sheets.map((sheet) => ({ ...sheet, path: book.path, bookName: book.name }))) || [], [source]);
  const selectedSheets = useMemo(() => allSheets.filter((sheet) => selected.has(keyOf(sheet.path, sheet.name))), [allSheets, selected]);
  const selection: SheetChoice[] = useMemo(() => source?.workbooks.map((book) => ({ path: book.path, sheets: book.sheets.filter((sheet) => selected.has(keyOf(book.path, sheet.name))).map((sheet) => sheet.name) })).filter((choice) => choice.sheets.length) || [], [source, selected]);
  const columns = useMemo(() => [...new Set(selectedSheets[0]?.columns.filter((name) => name && selectedSheets.every((sheet) => sheet.columns.includes(name))) || [])], [selectedSheets]);
  const queryKey = JSON.stringify({ inputPath, selection, column, headerRow, refresh });
  const groupsBusy = page === 'split' && !!column && selection.length > 0 && queryKey !== groupKey;
  const locked = busy || codeBusy || reading;
  const eligibleCount = allSheets.filter((sheet) => sheet.selectable).length;

  useEffect(() => {
    setColumn((old) => columns.includes(old) ? old : columns.find((name) => /部门|区域|网格|班级|门店/.test(name)) || columns[0] || '');
  }, [columns]);

  useEffect(() => {
    if (!activeSheet || !api || reading) { setPreview(null); return; }
    let current = true;
    setPreviewBusy(true); setPreviewError(''); setPreview(null);
    api.preview(activeSheet.path, activeSheet.name, headerRow).then((data) => { if (current) setPreview(data); })
      .catch((e) => { if (current) setPreviewError(message(e)); })
      .finally(() => { if (current) setPreviewBusy(false); });
    return () => { current = false; };
  }, [activeSheet, headerRow, reading, refresh]);

  useEffect(() => {
    setChosenValues(null); setGroupSearch(''); setGroupError(''); setGroupData(null);
    if (!api || page !== 'split' || !column || !selection.length || reading) return;
    let current = true;
    const timer = setTimeout(() => {
      api.groups({ inputPath, selection, headerRow, column }).then((data) => {
        if (current) { setGroupData(data); setGroupKey(queryKey); }
      }).catch((e) => { if (current) { setGroupError(message(e)); setGroupKey(queryKey); } });
    }, 250);
    return () => { current = false; clearTimeout(timer); };
  }, [queryKey, page, reading]); // queryKey 包含所有 Python 查询参数，过期结果不回填。

  const chooseSource = useCallback(async (kind: 'file' | 'directory') => {
    if (!api || locked) return;
    try {
      const path = await api.chooseSource(kind);
      if (path) { setInputPath(path); setRefresh((n) => n + 1); }
    } catch (e) { setError(message(e)); }
  }, [locked]);
  useEffect(() => {
    const keydown = (event: KeyboardEvent) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'o') { event.preventDefault(); void chooseSource('file'); }
      if (event.key === 'Escape') setHelp(false);
    };
    window.addEventListener('keydown', keydown);
    return () => window.removeEventListener('keydown', keydown);
  }, [chooseSource]);

  function toggleSheet(path: string, sheet: SheetInfo) {
    const key = keyOf(path, sheet.name);
    setSelected((old) => { const next = new Set(old); if (next.has(key)) next.delete(key); else next.add(key); return next; });
    setActiveSheet({ path, name: sheet.name });
    setResult(null);
  }
  function switchPage(next: typeof page) {
    if (busy || codeBusy) return;
    setPage(next); setError(''); setResult(null);
  }
  async function openResult(path: string) {
    try { await api?.openOutput(path); } catch (e) { setError(message(e)); }
  }
  const canRun = engine === 'ready' && !!source && !locked && selection.length > 0 && !!output &&
    (mode === 'merge' || (!!column && !groupsBusy && !!groupData?.values.length && !groupData.missing.length && !groupError && chosenValues?.length !== 0));

  async function start() {
    if (!api || !canRun) return;
    cancelRequested.current = false;
    setBusy(true); setProgress(0); setLogs([]); setResult(null); setError(''); setCancelling(false);
    try {
      jobId.current = await api.start({ inputPath, selection, headerRow, mode, outputPath: output, column,
        values: chosenValues, preserveFormat, keepFormulas, makeZip, mergeAcrossFiles, bundleSheets, mergeMode });
      if (cancelRequested.current) await api.cancel(jobId.current);
    } catch (e) { setBusy(false); setError(message(e)); }
  }
  async function cancel() {
    if (!api) return;
    cancelRequested.current = true;
    setCancelling(true);
    if (!jobId.current) return;
    try { await api.cancel(jobId.current); } catch (e) { setCancelling(false); setError(message(e)); }
  }

  return <div className="app-shell">
    <aside className="sidebar">
      <div className="brand"><img className="brand-mark" src={appIcon} alt="EasyExcel 图标" /><div>EasyExcel<small>表格工作台</small></div></div>
      <div className="nav-label">工作空间</div>
      <nav aria-label="主导航">
        <button className={page === 'split' ? 'nav-item active' : 'nav-item'} onClick={() => switchPage('split')} disabled={busy || codeBusy}><GitBranch size={19} /><span>Excel 拆分</span>{page === 'split' && <span className="nav-dot" />}</button>
        <button className={page === 'merge' ? 'nav-item active' : 'nav-item'} onClick={() => switchPage('merge')} disabled={busy || codeBusy}><Merge size={19} /><span>Excel 汇总</span>{page === 'merge' && <span className="nav-dot" />}</button>
        <button className={page === 'code' ? 'nav-item active' : 'nav-item'} onClick={() => switchPage('code')} disabled={busy || codeBusy}><Code2 size={19} /><span>Python 工作台</span>{page === 'code' && <span className="nav-dot" />}</button>
        <div className="nav-divider" />
        <button className={page === 'history' ? 'nav-item active' : 'nav-item'} onClick={() => switchPage('history')} disabled={busy || codeBusy}><History size={19} /><span>最近任务</span>{history.length > 0 && <span className="nav-count">{history.length}</span>}</button>
      </nav>
      <div className="sidebar-bottom">
        <div className="privacy-note"><span className="privacy-icon"><ShieldCheck size={20} /></span><strong>文件，留在你的电脑</strong><p>所有表格均在本地处理<br />无需上传，安心使用</p><span className="offline-pill"><span />100% 本地运行</span></div>
        <button className="nav-item help-button" onClick={() => setHelp(true)}><CircleHelp size={18} /><span>使用帮助</span><ArrowUpRight size={15} /></button>
        <div className="sidebar-footer"><span>EasyExcel Desktop</span><span className="version">0.5.0</span></div>
      </div>
    </aside>

    <div className="main-shell">
      <header className="topbar">
        <div className="breadcrumb"><LayoutGrid size={16} /><span>工作空间</span><ChevronRight size={13} /><strong>{page === 'history' ? '最近任务' : page === 'code' ? 'Python 工作台' : mode === 'split' ? 'Excel 拆分' : 'Excel 汇总'}</strong></div>
        <div className="topbar-right"><span className="local-label"><LockKeyhole size={13} />本地工作空间</span><div className="window-actions">
          <IconButton title="最小化" onClick={() => api?.window('minimize')}><Minus size={15} /></IconButton>
          <IconButton title="最大化或还原" onClick={() => api?.window('maximize')}><Maximize2 size={13} /></IconButton>
          <IconButton title="关闭" className="close-window" onClick={() => api?.window('close')}><X size={16} /></IconButton>
        </div></div>
      </header>

      <main>
        <div className="page-heading"><div><div className="eyebrow">YOUR DATA, WELL ORGANIZED</div><h1>{page === 'history' ? '每一次处理，都有迹可循。' : page === 'code' ? '用代码，延伸表格的可能。' : mode === 'split' ? '让每份数据，各就其位。' : '把分散的表格，汇在一起。'}</h1><p>{page === 'history' ? '快速找到最近生成的结果，接着完成你的工作。' : page === 'code' ? '导入 AI 编写的脚本，在离线电脑上完成分析、绘图和表格处理。' : mode === 'split' ? '选择工作表，按字段拆分。复杂的表格，也能井井有条。' : '选择需要整理的工作表，可按列名合并，也可分别保存到独立 Sheet。'}</p></div><button className="button guide-button" onClick={() => setHelp(true)}><CircleHelp size={16} />操作指南<ArrowUpRight size={14} /></button></div>

        {engine === 'error' && <div className="notice warning" role="alert"><AlertCircle size={18} /><span>{engineError}</span><button className="text-button" onClick={checkEngine}>重新检查</button></div>}
        {error && <div className="notice error" role="alert"><AlertCircle size={18} /><span>{error}</span><IconButton title="关闭提示" onClick={() => setError('')}><X size={15} /></IconButton></div>}

        {page === 'history' ? <section className="card history-card"><CardTitle title="最近处理记录"><span className="subtle">保留最近 20 次 · 仅存于本机</span></CardTitle>{!history.length ? <div className="empty-state"><Clock3 size={40} /><h3>还没有处理记录</h3><p>完成一次拆分或汇总后，结果会出现在这里。</p><button className="button primary" onClick={() => switchPage('split')}>开始第一个任务<ArrowRight size={16} /></button></div> : <div className="history-list">{history.map((item) => <div className="history-row" key={item.id}><span className="file-icon"><FileSpreadsheet size={22} /></span><div><strong>{item.name}</strong><p>{item.mode === 'code' ? 'Python 脚本' : item.mode === 'split' ? 'Excel 拆分' : 'Excel 汇总'} · {new Date(item.date).toLocaleString('zh-CN')} · {item.stopped ? '已停止' : '处理结束'}</p></div><button className="button small" onClick={() => void openResult(item.output)}><FolderOpen size={15} />打开结果</button></div>)}</div>}</section> : <>
          <div className="workflow-strip" aria-label="处理步骤">
            {[['选择数据源', '导入文件或文件夹'], ['选择工作表', '只处理你需要的 Sheet'], [page === 'code' ? '编写并运行代码' : mode === 'split' ? '配置并拆分' : '合并并导出', '结果保存到本地']].map(([title, detail], index) => <div className={`workflow-step ${index === 0 || (index === 1 && source) || (index === 2 && selectedSheets.length) ? 'current' : ''}`} key={title}><span className="workflow-number">{index === 0 && source ? <Check size={15} /> : `0${index + 1}`}</span><div><strong>{title}</strong><small>{detail}</small></div>{index < 2 && <ChevronRight className="step-arrow" size={18} />}</div>)}
          </div>
          {result && <ResultBanner result={result} onOpen={() => void openResult(result.output)} />}
          <div className={`workspace-grid ${page === 'code' ? 'code-mode' : ''}`}>
            <div className="data-column">
              <section className="card source-card">
                <CardTitle step="01" title="数据源"><span className="tag neutral">.xlsx / .xls</span></CardTitle>
                <div className={`drop-zone ${dragging ? 'dragging' : ''} ${source ? 'has-source' : ''}`} onDragOver={(event) => { event.preventDefault(); if (!locked) setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={(event) => {
                  event.preventDefault(); setDragging(false); if (locked || !api) return;
                  if (event.dataTransfer.files.length !== 1) { setError('请一次拖入一个工作簿，或将多个文件放入文件夹后拖入。'); return; }
                  const path = api.pathForFile(event.dataTransfer.files[0]); if (path) { setInputPath(path); setRefresh((n) => n + 1); }
                }}>
                  <div className="source-description"><span className={`upload-icon ${source ? 'loaded' : ''}`}>{reading ? <LoaderCircle size={25} className="spin" /> : source ? source.isDirectory ? <FolderOpen size={27} /> : <FileSpreadsheet size={27} /> : <Upload size={25} />}</span><div><strong>{reading ? '正在读取工作簿…' : source ? source.name : '将 Excel 文件拖到这里'}</strong><p title={source?.path}>{reading ? '正在识别工作表与表头，请稍候' : source ? `${source.workbooks.length} 个文件 · ${allSheets.length} 个 Sheet · ${size(source.workbooks.reduce((total, book) => total + book.size, 0))}` : '支持单个工作簿，也支持整个文件夹'}</p></div>{source && !reading && <CircleCheck size={19} className="source-check" />}</div>
                  <div className="source-buttons"><button className="button primary" disabled={locked || !api} onClick={() => void chooseSource('file')}><FolderOpen size={16} />{source ? '更换文件' : '选择文件'}</button><button className="button" disabled={locked || !api} onClick={() => void chooseSource('directory')}><Folder size={16} />选择文件夹</button></div>
                </div>
                {source ? <div className="source-path"><Folder size={13} /><span title={source.path}>{source.path}</span><button className="text-button" disabled={locked} onClick={() => setRefresh((n) => n + 1)}><RefreshCw size={12} />重新读取</button></div> : <div className="source-tip"><ShieldCheck size={13} />{page === 'code' ? '选择表格后，脚本可读取完整数据' : '源文件保持原样，结果单独保存'}<span>Ctrl + O 选择文件</span></div>}
              </section>

              <section className="card sheets-card">
                <CardTitle step="02" title="选择工作表"><span className="tag">已选 {selectedSheets.length} / {allSheets.length} 个 Sheet</span></CardTitle>
                <div className="sheet-toolbar"><label className="search-field"><Search size={15} /><input aria-label="搜索工作表" placeholder="搜索工作表名称…" value={search} onChange={(e) => setSearch(e.target.value)} disabled={!source || locked} /></label><button className="text-button" disabled={!eligibleCount || locked} onClick={() => setSelected(new Set(allSheets.filter((sheet) => sheet.selectable).map((sheet) => keyOf(sheet.path, sheet.name))))}><CheckCheck size={14} />全选</button><span className="toolbar-divider" /><button className="text-button subtle" disabled={!selected.size || locked} onClick={() => setSelected(new Set())}>清空</button></div>
                {!source ? <div className="empty-sheets"><div className="sheet-stack"><Table2 size={31} /><span /><span /></div><strong>工作表将在这里展开</strong><p>先选择一个 Excel，再勾选需要处理的 Sheet</p></div> : <div className="sheet-list" aria-busy={reading}>
                  {source.workbooks.map((book) => <div className="workbook-group" key={book.path}>
                    {source.isDirectory && <div className="workbook-label"><FileSpreadsheet size={13} /><span title={book.path}>{book.name}</span><span>{book.sheets.length} Sheets</span></div>}
                    {book.error && <p className="inline-error">{book.name}：{book.error}</p>}
                    {book.sheets.filter((sheet) => sheet.name.toLowerCase().includes(search.toLowerCase())).map((sheet) => {
                      const checked = selected.has(keyOf(book.path, sheet.name));
                      const active = activeSheet?.path === book.path && activeSheet.name === sheet.name;
                      return <div className={`sheet-row ${checked ? 'selected' : ''} ${!sheet.selectable ? 'unavailable' : ''}`} key={sheet.name}>
                        <label className="sheet-choice"><input type="checkbox" aria-label={`选择 ${book.name} / ${sheet.name}`} checked={checked} disabled={locked || !sheet.selectable} onChange={() => toggleSheet(book.path, sheet)} /><span className="sheet-symbol"><Table2 size={19} /></span><span className="sheet-name"><strong>{sheet.name}</strong><small>{sheet.reason || `${sheet.estimatedRows == null ? '行数未知' : `约 ${number(sheet.estimatedRows)} 行`} · ${sheet.columns.filter(Boolean).length} 列`}</small></span></label>
                        {sheet.state !== 'visible' ? <span className="tag neutral"><LockKeyhole size={11} />隐藏</span> : <button className={`preview-button ${active ? 'active' : ''}`} title={`预览 ${sheet.name}`} aria-label={`预览 ${sheet.name}`} disabled={locked} onClick={() => setActiveSheet({ path: book.path, name: sheet.name })}><Eye size={14} />{active ? '预览中' : '预览'}</button>}
                      </div>;
                    })}
                  </div>)}
                  {allSheets.length > 0 && !allSheets.some((sheet) => sheet.name.toLowerCase().includes(search.toLowerCase())) && <p className="list-empty">没有找到匹配的 Sheet</p>}
                </div>}
                <div className="card-footnote"><span className="tiny-dot" />支持多选，只处理勾选的工作表；隐藏 Sheet 不参与处理。</div>
              </section>

              <section className="card preview-card">
                <CardTitle title="数据预览"><span className="subtle">前 100 行 · 原表样式与图表</span></CardTitle>
                <div aria-busy={previewBusy}>
                  {previewBusy ? <div className="preview-placeholder"><LoaderCircle size={23} className="spin" /><span>正在读取工作表样式…</span></div> : previewError ? <div className="preview-placeholder"><AlertCircle size={23} /><span>{previewError}</span></div> : preview ? <SheetPreview key={keyOf(activeSheet!.path, activeSheet!.name)} preview={preview} name={activeSheet!.name} /> : <div className="preview-placeholder"><Columns3 size={28} strokeWidth={1.4} /><span>选择工作表，即可在这里预览数据</span></div>}
                </div><div className="card-footnote"><Eye size={12} />包含原表标题与表头；实际导出以勾选的 Sheet 为准。</div>
              </section>
            </div>

            {page === 'code' ? <CodeWorkspace request={{ inputPath, selection, headerRow }} outputPath={output} onOutputChange={setOutput}
              headerRow={headerRow} onHeaderChange={setHeaderRow} ready={engine === 'ready' && !reading} onBusy={setCodeBusy}
              onFinished={() => { void api?.history().then(setHistory); }} sheets={selectedSheets} /> : <aside className="config-column"><section className="card config-card">
              <CardTitle step="03" title={mode === 'split' ? '拆分设置' : '汇总设置'}><Settings2 size={17} className="subtle" /></CardTitle>
              <div className="config-body">
                {mode === 'split' ? <>
                  <label className="field-label" htmlFor="split-column">按哪个字段拆分 <span className="required">*</span></label>
                  <div className="select-wrap"><GitBranch size={16} /><select id="split-column" value={column} onChange={(e) => { setColumn(e.target.value); setResult(null); }} disabled={!columns.length || locked}><option value="" disabled>{selectedSheets.length ? '没有共同字段' : '先选择工作表'}</option>{columns.map((name) => <option key={name}>{name}</option>)}</select><ChevronDown size={15} /></div>
                  <p className="field-hint">每个不同的取值，将生成一份独立文件。</p>
                  {selectedSheets.length > 0 && !columns.length && <p className="inline-error">所选 Sheet 没有共同字段，请调整选择。</p>}
                  <div className="group-box"><div className="group-heading"><span><ListFilter size={14} />分组预览</span>{groupsBusy ? <LoaderCircle size={14} className="spin" /> : <span className="tag neutral">{groupData?.values.length || 0} 组</span>}</div>
                    {groupsBusy ? <p className="group-placeholder">正在扫描所选工作表的全部分组…</p> : groupError ? <p className="inline-error">{groupError}</p> : groupData?.values.length ? <>
                      {groupData.values.length > 8 && <input className="group-search" aria-label="搜索分组" placeholder="搜索分组…" value={groupSearch} onChange={(e) => setGroupSearch(e.target.value)} />}
                      <div className="group-values">{groupData.values.filter((v) => v.toLowerCase().includes(groupSearch.toLowerCase())).slice(0, 100).map((value) => <label className="group-chip" key={value}><input type="checkbox" checked={chosenValues === null || chosenValues.includes(value)} disabled={locked} onChange={(e) => setChosenValues((old) => e.target.checked ? [...(old || []), value] : (old || groupData.values).filter((v) => v !== value))} /><span title={value}>{value}</span></label>)}</div>
                      <div className="group-actions"><span>已选 {chosenValues?.length ?? groupData.values.length} 组</span><button className="text-button" disabled={locked} onClick={() => setChosenValues(null)}>全选</button><button className="text-button subtle" disabled={locked} onClick={() => setChosenValues([])}>清空</button></div>
                      {groupData.values.length > 100 && <p className="field-hint">列表最多显示 100 组，可搜索；全选包含全部 {groupData.values.length} 组。</p>}
                    </> : <p className="group-placeholder">选择拆分字段后，自动识别所有分组</p>}
                    {!!groupData?.missing.length && <p className="inline-error">部分 Sheet 缺少该字段：{groupData.missing.join('、')}</p>}
                  </div>
                  <div className="config-divider" />
                  <span className="field-label">处理方式</span>
                  <div className="speed-options"><button className={preserveFormat ? 'speed-option selected' : 'speed-option'} disabled={locked} onClick={() => setPreserveFormat(true)}><span><Table2 size={15} />标准模式{preserveFormat && <CircleCheck size={13} />}</span><small>保留原表格式</small></button><button className={!preserveFormat ? 'speed-option selected' : 'speed-option'} disabled={locked} onClick={() => { setPreserveFormat(false); setKeepFormulas(false); }}><span><Zap size={15} />极速模式{!preserveFormat && <CircleCheck size={13} />}</span><small>仅保留数据值</small></button></div>
                  <Toggle label="同时生成总工作簿" detail="每份分表作为独立 Sheet，原分表也会保留" checked={bundleSheets} onChange={(value) => { setBundleSheets(value); setResult(null); }} disabled={locked} />
                  {bundleSheets && <p className="field-hint">包含多个 Sheet 的分表会逐张收入总工作簿；名称过长或重名时自动调整。</p>}
                </> : <>
                  <label className="field-label" htmlFor="merge-mode">汇总方式</label>
                  <select id="merge-mode" className="plain-select" value={mergeMode} disabled={locked} onChange={(e) => { setMergeMode(e.target.value as 'rows' | 'sheets'); setResult(null); }}><option value="rows">按列名合并数据</option><option value="sheets">每个表独立一个 Sheet</option></select>
                  <div className="merge-note"><span className="merge-symbol"><Merge size={24} /></span><strong>{mergeMode === 'sheets' ? '一本工作簿，多张独立表' : '按列名，自动对齐'}</strong><p>{mergeMode === 'sheets' ? '将已拆分的文件放入一个文件夹，选择并勾选需要的表，整理成一个多 Sheet 工作簿。' : '将勾选的同模板 Sheet 合并为一张表，保留格式，并添加「来源」列。'}</p><small>{mergeMode === 'sheets' ? '保留各表的单元格样式、合并区域与行高列宽；公式使用已保存的计算值。' : '结构差异较大的表会跳过，并在汇总说明中列出。'}</small></div>
                </>}

                <div className="config-divider" /><label className="field-label" htmlFor="output-path">结果保存位置</label>
                <div className="output-field"><Folder size={16} /><input id="output-path" aria-label="结果保存位置" value={output} placeholder="默认保存在源文件旁" title={output} readOnly /><IconButton title="选择输出文件夹" disabled={locked || !api} onClick={async () => { try { const path = await api?.chooseOutput(); if (path) setOutput(path); } catch (e) { setError(message(e)); } }}><FolderOpen size={16} /></IconButton></div>
                <p className="field-hint">每次生成独立结果，不覆盖源文件。</p>
                {mode === 'split' && <Toggle label="按分组打包 ZIP" detail="适用于文件夹批量拆分，方便分发" checked={makeZip} onChange={setMakeZip} disabled={locked || !source?.isDirectory} />}
                <button className="advanced-toggle" aria-expanded={advanced} onClick={() => setAdvanced(!advanced)}><Settings2 size={14} />高级设置<ChevronDown size={14} className={advanced ? 'rotated' : ''} /></button>
                {advanced && <div className="advanced-body"><label className="field-label" htmlFor="header-mode">表头识别</label><select id="header-mode" className="plain-select" disabled={locked} value={headerRow === null ? 'auto' : 'row'} onChange={(e) => setHeaderRow(e.target.value === 'auto' ? null : 1)}><option value="auto">自动识别表头</option><option value="row">手动指定表头行</option></select>{headerRow !== null && <label className="header-row-field">表头在第 <input type="number" aria-label="表头行号" min={1} max={1000} value={headerRow} disabled={locked} onChange={(e) => { const n = Number(e.target.value); if (Number.isInteger(n) && n >= 1 && n <= 1000) setHeaderRow(n); }} /> 行</label>}{mode === 'split' && <><Toggle label="保留同行公式" detail="仅标准模式可用；跨行公式转为值" checked={keepFormulas} onChange={setKeepFormulas} disabled={locked || !preserveFormat} />{source?.isDirectory && <Toggle label="跨文件合并同一分组" checked={mergeAcrossFiles} onChange={setMergeAcrossFiles} disabled={locked} />}</>}</div>}
              </div>
            </section>
              <div className="help-tip"><Sparkles size={17} /><p><strong>一个工作簿，也可以有多种选择</strong>只勾选需要的 Sheet，说明页、旧数据无需一起处理。</p></div>
            </aside>}
          </div>

          {page !== 'code' && (logs.length > 0 || busy) && <section className="card log-card"><button className="log-toggle" onClick={() => setShowLogs(!showLogs)}><span><Columns3 size={16} />处理详情 <span className="tag neutral">{logs.length}</span></span><ChevronDown size={16} className={showLogs ? 'rotated' : ''} /></button>{showLogs && <div className="log-body" role="log">{logs.map((line, i) => <div key={i}>{line}</div>)}<div ref={logEnd} /></div>}</section>}
        </>}
        <footer className="workspace-footer"><span className={`engine-status ${engine}`}><span />{engine === 'ready' ? '本地引擎已就绪' : engine === 'checking' ? '正在连接本地引擎…' : '本地引擎未连接'}</span><span>简单处理，专注工作。</span></footer>
      </main>
      {page !== 'history' && page !== 'code' && <div className="action-dock">
        {busy ? <><div className="dock-progress"><div className="run-summary"><strong>{cancelling ? '正在停止…' : '正在处理数据…'}</strong><span>{Math.round(progress * 100)}%</span></div><progress max={1} value={progress} aria-label="处理进度" /></div><p className="run-detail">{cancelling ? '等待当前操作安全结束，已生成结果将保留。' : logs.filter(Boolean).at(-1) || '正在准备本地处理任务…'}</p><button className="button stop-button" onClick={() => void cancel()} disabled={cancelling}><Square size={13} />停止任务</button></> : <><div className="dock-summary"><span className="dock-icon"><FileSpreadsheet size={20} /></span><div><strong>已选择 <b>{selectedSheets.length}</b> 个 Sheet<span className="summary-separator">/</span>{selection.length} 个文件</strong><p>仅处理勾选的工作表，源文件保持原样</p></div></div><span className="dock-local"><ShieldCheck size={14} />全程本地处理</span><button className="button primary run-button" disabled={!canRun} onClick={() => void start()}><Play size={14} fill="currentColor" />{mode === 'split' ? '开始拆分' : '开始汇总'}<ArrowRight size={17} /></button></>}
      </div>}
    </div>
    {help && <div className="modal-backdrop" onClick={() => setHelp(false)}><section className="help-modal" role="dialog" aria-modal="true" aria-labelledby="help-title" onClick={(event) => event.stopPropagation()}><div className="card-title"><h2 id="help-title">三步，整理好你的表格</h2><IconButton title="关闭帮助" onClick={() => setHelp(false)}><X size={19} /></IconButton></div><div className="help-steps"><p><span>01</span><strong>选择数据源</strong>拖入一个 Excel 或一个文件夹，也可以点击选择按钮。支持 .xlsx 和 .xls。</p><p><span>02</span><strong>勾选需要的 Sheet</strong>默认选中每个文件的首个可处理 Sheet。支持多选、全选、搜索和独立预览；隐藏 Sheet 不处理。</p><p><span>03</span><strong>设置并开始</strong>拆分时选择共同字段，可开启「同时生成总工作簿」；汇总可按列名合并，或选择「每个表独立一个 Sheet」。结束后点击「查看结果」。</p></div><div className="notice"><ShieldCheck size={18} /><span>Excel 读取、预览、拆分和汇总全部由本机 Python 完成。Electron 界面不会上传文件。</span></div><p className="help-footnote">旧版 Python 界面继续支持查收、方案和加密分发，可使用项目根目录的 main.py 启动。新版提供拆分、普通汇总和离线 Python 工作台。Python 页面可导入脚本、导出 AI 编程说明，或使用内置示例。</p><button className="button primary" onClick={() => setHelp(false)}>开始使用<ArrowRight size={16} /></button></section></div>}
  </div>;
}
