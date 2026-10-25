// Copyright (c) 2026 AbeLin · MIT License
import { useEffect, useRef, useState } from 'react';
import { ArrowDownToLine, BookMarked, Check, Code2, Copy, FileCode2, FolderOpen, LoaderCircle, Play, Save, Square, Terminal, X } from 'lucide-react';
import type { CodeEnvironment, CodeOutput, CodeResult, ScriptRecord, SelectionRequest } from '../shared/types';
import PythonEditor from './PythonEditor';
import ScriptLibrary from './ScriptLibrary';
import { aiInstructions, templates } from './codeTemplates';
import './code-workspace.css';

const api = window.desktop;
const draftKey = 'excelrouter.python.draft.v1';
const message = (error: unknown) => error instanceof Error ? error.message.replace(/^Error invoking remote method '[^']+': (?:Error: )?/, '') : String(error);
function initialCode() { try { return localStorage.getItem(draftKey) ?? templates[0].code; } catch { return templates[0].code; } }

export default function CodeWorkspace({ request, outputPath, onOutputChange, headerRow, onHeaderChange, ready, onBusy, onFinished, sheets }: {
  request: SelectionRequest; outputPath: string; onOutputChange: (path: string) => void;
  headerRow: number | null; onHeaderChange: (row: number | null) => void;
  ready: boolean; onBusy: (busy: boolean) => void; onFinished: () => void;
  sheets: { name: string; bookName: string; columns: string[] }[];
}) {
  const [code, setCode] = useState(initialCode);
  const [filename, setFilename] = useState('分析脚本.py');
  const [environment, setEnvironment] = useState<CodeEnvironment | null>(null);
  const [environmentError, setEnvironmentError] = useState('');
  const [running, setRunning] = useState(false);
  const [stopping, setStopping] = useState(false);
  const [timeout, setTimeoutValue] = useState(120);
  const [logs, setLogs] = useState<string[]>([]);
  const [outputs, setOutputs] = useState<CodeOutput[]>([]);
  const [result, setResult] = useState<CodeResult | null>(null);
  const [error, setError] = useState('');
  const [info, setInfo] = useState('');
  const [guide, setGuide] = useState(false);
  const [library, setLibrary] = useState(false);
  const [libraryId, setLibraryId] = useState('');
  const [savingLibrary, setSavingLibrary] = useState(false);
  const [tab, setTab] = useState<'results' | 'logs'>('results');
  const [pendingTemplate, setPendingTemplate] = useState<string | null>(null);
  const id = useRef<string | null>(null);
  const active = useRef(false);
  const cancelRequested = useRef(false);
  const latest = useRef({ onBusy, onFinished });
  latest.current = { onBusy, onFinished };
  const prompt = aiInstructions(environment, sheets);
  const canRun = ready && !!environment && !running && !!code.trim() && request.selection.length > 0 && !!outputPath;

  useEffect(() => { try { localStorage.setItem(draftKey, code); } catch { /* Export remains available. */ } }, [code]);
  useEffect(() => {
    let current = true;
    api?.codeEnvironment().then((value) => { if (current) setEnvironment(value); }).catch((e) => { if (current) setEnvironmentError(message(e)); });
    return () => { current = false; };
  }, []);
  useEffect(() => api?.onCode((event) => {
    if (!active.current || (id.current && event.jobId !== id.current)) return;
    if (event.type === 'log') setLogs((old) => [...old.slice(-499), event.message || '']);
    if (event.type === 'output' && event.data) setOutputs((old) => [...old, event.data!].slice(0, 40));
    if (event.type === 'result' && event.result) {
      active.current = false; id.current = null;
      setRunning(false); setStopping(false); latest.current.onBusy(false); latest.current.onFinished();
      setResult(event.result);
      if (event.result.status === 'error') { setTab('logs'); setError(event.result.error || '运行失败'); }
    }
  }), []);

  async function run(sourceCode = code, scriptName = filename) {
    if (!api || !ready || !environment || !sourceCode.trim() || !request.selection.length || !outputPath || active.current) return;
    active.current = true; cancelRequested.current = false; id.current = null;
    setRunning(true); onBusy(true); setStopping(false); setLogs([]); setOutputs([]); setError(''); setInfo(''); setResult(null); setTab('results');
    try {
      const job = await api.startCode({ ...request, outputPath, code: sourceCode, scriptName, timeoutSeconds: timeout });
      if (active.current) { id.current = job; if (cancelRequested.current) await api.cancelCode(job); }
    } catch (e) { active.current = false; setRunning(false); onBusy(false); setError(message(e)); }
  }
  async function stop() {
    cancelRequested.current = true; setStopping(true);
    try { if (id.current) await api?.cancelCode(id.current); } catch (e) { setStopping(false); setError(message(e)); }
  }
  async function importScript() {
    try { const file = await api?.importScript(); if (file) { setCode(file.code); setFilename(file.name); setInfo('脚本已导入，可检查代码后运行。'); setError(''); } } catch (e) { setError(message(e)); }
  }
  async function save(kind: 'script' | 'instructions') {
    try { const path = await api?.saveText(kind, kind === 'script' ? code : prompt); if (path) setInfo(`已保存：${path}`); } catch (e) { setError(message(e)); }
  }
  async function copy(text: string) {
    try { await api?.copyText(text); setInfo('已复制到剪贴板。'); } catch (e) { setError(message(e)); }
  }
  function copyLogs() {
    const text = logs.join('\n');
    const detail = result?.error || error;
    void copy((detail && !text.includes(detail) ? `${text}\n\n${detail}` : text).slice(-200000));
  }
  function loadTemplate(templateId: string) {
    const item = templates.find((t) => t.id === templateId);
    if (item) { setCode(item.code); setFilename(`${item.title}.py`); }
    setPendingTemplate(null);
  }

  async function collectScript() {
    if (!api) return;
    setSavingLibrary(true);
    try {
      const saved = await api.saveScript({ code, name: filename.replace(/\.py$/i, ''), rememberOnly: true });
      if (saved.archived) await api.updateScript(saved.id, { archived: false });
      setLibraryId(saved.id); setLibrary(true);
    } catch (e) { setError(message(e)); } finally { setSavingLibrary(false); }
  }
  async function loadScript(script: ScriptRecord, execute: boolean) {
    if (!api) return;
    if (code.trim() && code !== script.code && !templates.some((template) => template.code === code)) {
      await api.saveScript({ code, name: filename.replace(/\.py$/i, ''), category: '草稿', rememberOnly: true });
    }
    setCode(script.code); setFilename(`${script.name}.py`); setLibrary(false); setLibraryId(script.id);
    setError(''); setInfo(`已载入：${script.name}。使用当前选择的数据源和 Sheet。`);
    if (execute) await run(script.code, `${script.name}.py`);
  }

  return <div className="code-workspace">
    {library && <ScriptLibrary initialId={libraryId} onClose={() => setLibrary(false)} onLoad={loadScript} canRun={ready && !!environment && !!request.selection.length && !!outputPath && !running} />}
    <section className="card code-card">
      <div className="card-title"><h2><span className="step-number">03</span>Python 脚本</h2><span className="tag"><span className="tiny-dot" />离线运行</span></div>
      <div className="code-library-bar"><button className="button small" disabled={running || savingLibrary || !api} onClick={() => setLibrary(true)}><BookMarked size={15} />我的脚本库</button><span>运行自动收录 · 下次一键调用</span><button className="text-button" disabled={running || savingLibrary || !code.trim() || !api} onClick={() => void collectScript()}>{savingLibrary ? <LoaderCircle className="spin" size={13} /> : <Save size={13} />}存入脚本库</button></div>
      <div className="code-toolbar">
        <button className="button small" disabled={running || !api} onClick={() => void importScript()}><FileCode2 size={14} />导入 .py</button>
        <button className="button small" disabled={running || !api} onClick={() => void save('script')}><Save size={14} />保存脚本</button>
        <button className="text-button code-ai-button" disabled={!environment} onClick={() => setGuide(true)}><Copy size={14} />给 AI 的说明</button>
      </div>
      <div className="code-template-bar"><span>从示例开始</span>{templates.map((t) => <button key={t.id} disabled={running} onClick={() => code.trim() && !templates.some((item) => item.code === code) ? setPendingTemplate(t.id) : loadTemplate(t.id)}>{t.title}</button>)}</div>
      {pendingTemplate && <div className="code-inline-prompt">替换当前编辑内容？可先保存脚本。<button className="text-button" onClick={() => loadTemplate(pendingTemplate)}>替换</button><button className="text-button subtle" onClick={() => setPendingTemplate(null)}>取消</button></div>}
      <div className="code-filename"><span><Code2 size={13} />{filename}</span><span>本机自动保存草稿 · Ctrl + Enter 运行</span></div>
      <PythonEditor value={code} disabled={running} onChange={setCode} onRun={() => void run()} />
      <div className="code-context"><code>df</code> 第一张表 <span>·</span> <code>tables</code> 全部勾选的表 <span>·</span> <strong>运行读取完整数据</strong></div>
      <details className="code-settings"><summary>运行设置与内置环境</summary>
        <label>表头行<select aria-label="Python 表头识别" value={headerRow === null ? 'auto' : 'manual'} disabled={running || !ready} onChange={(e) => onHeaderChange(e.target.value === 'auto' ? null : 1)}><option value="auto">自动识别</option><option value="manual">手动指定</option></select>{headerRow !== null && <input type="number" aria-label="Python 表头行号" min={1} max={1000} value={headerRow} disabled={running} onChange={(e) => { const n = Number(e.target.value); if (Number.isInteger(n) && n >= 1 && n <= 1000) onHeaderChange(n); }} />}</label>
        <label>运行上限<select aria-label="运行时间上限" value={timeout} disabled={running} onChange={(e) => setTimeoutValue(Number(e.target.value))}><option value={30}>30 秒</option><option value={120}>2 分钟</option><option value={600}>10 分钟</option><option value={3600}>1 小时</option></select></label>
        <p className="code-environment">{environment ? `Python ${environment.python} · ${Object.entries(environment.libraries).map(([name, version]) => `${name} ${version}`).join(' · ')}` : environmentError || '正在检查内置环境…'}</p>
        {environment?.capabilities && <details className="finance-libraries"><summary>查看内置库的用途与导入方式</summary><div>{environment.capabilities.map((item) => <article key={item.name}><strong>{item.name} <small>{item.version}</small></strong><p>{item.purpose}</p><code>{item.example}</code></article>)}</div></details>}
        <p>每次运行是新进程。隐藏列不载入，隐藏行和合计行保留；导出为数据值。代码以当前用户权限在本机运行。</p>
      </details>
      <div className="code-output-path"><FolderOpen size={14} /><span title={outputPath}>{outputPath || '选择表格后设置结果位置'}</span><button className="text-button" disabled={running || !api} onClick={async () => { try { const path = await api?.chooseOutput(); if (path) onOutputChange(path); } catch (e) { setError(message(e)); } }}>更改</button></div>
      <div className="code-runbar"><span>{running ? '正在本机执行，可随时停止' : `已选 ${sheets.length} 个 Sheet · 每次保存独立结果`}</span>{running ? <button className="button" disabled={stopping} onClick={() => void stop()}><Square size={14} />{stopping ? '正在停止…' : '停止代码'}</button> : <button className="button primary" disabled={!canRun} onClick={() => void run()}><Play size={14} />运行代码</button>}</div>
    </section>
    {environmentError && <div className="notice error">{environmentError}</div>}
    {info && <div className="code-info" role="status"><Check size={14} /><span>{info}</span><button className="icon-button" aria-label="关闭脚本提示" onClick={() => setInfo('')}><X size={13} /></button></div>}
    <section className="card code-results">
      <div className="code-result-tabs"><button className={tab === 'results' ? 'active' : ''} onClick={() => setTab('results')}>结果与图表 <span>{outputs.length}</span></button><button className={tab === 'logs' ? 'active' : ''} onClick={() => setTab('logs')}><Terminal size={13} />运行日志</button>{running && <LoaderCircle size={14} className="spin" />}{(!!logs.length || !!error) && <button className="copy-log" aria-label="复制运行日志" onClick={copyLogs}><Copy size={13} /></button>}</div>
      {result && <div className={`code-result-status ${result.status}`} role="status"><strong>{({ success: '脚本运行完成', error: '脚本运行失败', stopped: '代码已停止', timeout: '已到运行时间上限' })[result.status]}</strong><span>{result.duration.toFixed(1)} 秒{result.status === 'success' ? ` · ${result.rows.toLocaleString()} 行输入数据` : ' · 已写出的文件会保留，请检查是否完整'}</span>{result.output && <button className="text-button" onClick={async () => { try { await api?.openOutput(result.output); } catch (e) { setError(message(e)); } }}><FolderOpen size={13} />打开结果目录</button>}</div>}
      {error && <pre className="code-error" role="alert">{error}</pre>}
      {tab === 'logs' ? <pre className="code-logs" role="log">{logs.join('\n') || '运行后显示 print 输出和错误信息。'}</pre> : <div className="code-output-list">
        {!outputs.length && <div className="code-empty"><Code2 size={27} /><strong>{running ? '正在处理选中的表格…' : '让代码接着处理你的数据'}</strong><p>用 display() 展示分析表，用 plt.show() 或 save_figure() 展示图表。</p></div>}
        {outputs.map((item, i) => <figure className="code-output" key={i}><figcaption>{item.title}</figcaption>{item.kind === 'image' ? <img src={item.src} alt={item.title} /> : <><div className="code-table-scroll"><table><thead><tr><th>#</th>{item.columns.map((name, col) => <th key={col}>{name}</th>)}</tr></thead><tbody>{item.rows.map((row, r) => <tr key={r}><th>{item.index[r]}</th>{row.map((cell, c) => <td key={c} title={cell}>{cell}</td>)}</tr>)}</tbody></table></div><p>共 {item.totalRows.toLocaleString()} 行 × {item.totalColumns} 列，预览最多 100 行、40 列；单格最多 1,000 字符。</p></>}</figure>)}
        {result && result.artifacts.length > 0 && <details className="code-artifacts"><summary>本次文件 · {result.artifactCount} 份（含脚本与运行记录）</summary>{result.artifacts.map((name) => <div key={name}>{name}</div>)}</details>}
      </div>}
    </section>
    {guide && <div className="modal-backdrop" onClick={() => setGuide(false)}><section className="code-guide" role="dialog" aria-modal="true" aria-label="给 AI 的编程说明" onClick={(e) => e.stopPropagation()}><div className="card-title"><h2>把这份说明交给联网电脑上的 AI</h2><button className="icon-button" aria-label="关闭编程说明" onClick={() => setGuide(false)}><X size={18} /></button></div><p>这份普通说明包含原文件名、Sheet 名和表头，不属于严格脱敏包。需要匿名结构与虚构样例时，请关闭此窗口，使用工作表下方的「脱敏导出」。</p><pre>{prompt}</pre><div className="code-guide-actions"><button className="button" onClick={() => void save('instructions')}><ArrowDownToLine size={15} />导出说明 .txt</button><button className="button primary" onClick={() => void copy(prompt)}><Copy size={15} />复制说明</button></div></section></div>}
  </div>;
}
