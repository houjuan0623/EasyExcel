// Copyright (c) 2026 AbeLin · MIT License
import { useEffect, useMemo, useRef, useState } from 'react';
import { BarChart3, LineChart, PieChart, ScatterChart, Box, Grid2X2, Hash, ChartColumn, Plus, Trash2, Save, Upload, Download, RefreshCw, MousePointer2, Square, LoaderCircle, FolderOpen } from 'lucide-react';
import type { SourceInfo } from '../shared/types';
import { parseAnalysisPlan, type AnalysisBlock, type AnalysisCard, type AnalysisField, type AnalysisPlan, type AnalysisRequest, type AnalysisResult, type AnalysisSchema, type AnalysisTable, type ChartKind } from '../shared/analysis';
import AnalysisRangePicker, { columnLabel } from './AnalysisRangePicker';
import './analysis.css';
import DashboardCanvas from './DashboardCanvas';
import { DASHBOARD_CANVAS, nextLayout } from '../shared/dashboard-layout';

const api = window.desktop;
const message = (e: unknown) => e instanceof Error ? e.message.replace(/^Error invoking remote method '[^']+': (?:Error: )?/, '') : String(e);
const id = () => crypto.randomUUID();
const display = (v: unknown) => v === null || v === undefined ? '—' : typeof v === 'number' ? new Intl.NumberFormat('zh-CN', { maximumFractionDigits: 4 }).format(v) : String(v);
const storageKey = 'easyexcel.dashboard.v1';
const templates = [
  { kind: 'metric', name: '指标卡', icon: Hash, purpose: '快速查看一个总量或平均水平', help: '选择一个数值字段，再选合计、平均值或有效值计数。单位只作标注，不进行金额或百分比换算。', method: '对全部有效数值计算一个指标；缺失项不会补成零。' },
  { kind: 'bar', name: '柱形图', icon: BarChart3, purpose: '比较部门、类别或项目', help: '分类轴选择部门、项目等字段；数值字段选择金额、数量等。可选多个数值字段比较预算与实际。', method: '先按分类分组，再计算选定的合计、平均值等指标。同名类别会合并。' },
  { kind: 'line', name: '折线图', icon: LineChart, purpose: '观察时间或有序序列的变化', help: '横轴选择日期、月份或有序序列，纵轴选择数值。日期可按日、月、年汇总；文字月份请核对排序。', method: '相同时间/分类先汇总，再连线。没有记录的日期不自动补零。' },
  { kind: 'pie', name: '饼图', icon: PieChart, purpose: '查看非负总量的组成占比', help: '选择一个分类和一个非负数值字段。支持合计或计数；负数、总量为零、分类超限时会提示，避免错误占比。', method: '每类汇总值 ÷ 全部分类汇总值。所有分类都须进入图表，不能只截取部分当作整体。' },
  { kind: 'histogram', name: '直方图', icon: ChartColumn, purpose: '查看金额、时长等的分布', help: '只选一个数值字段，不需要分类轴。设置组数后，观察数值集中区间和长尾。', method: '使用全部有效数值划分等宽区间，统计每个区间的频数。' },
  { kind: 'scatter', name: '散点图', icon: ScatterChart, purpose: '查看两个数值之间的关系', help: 'X、Y 都必须是数值。同一行表示一对观测；跨 Sheet 请按唯一共同字段关联后再选两列。', method: '每条有效配对记录绘制一个点；Pearson 相关性基于全部有效配对，相关不代表因果。' },
  { kind: 'box', name: '箱线图', icon: Box, purpose: '比较分布、离散程度和异常点', help: '选择 1～6 个数值字段，各字段分别绘制。不同单位的字段请分开成图。异常点只标注，不自动删除。', method: '显示中位数、25%/75% 分位数和 1.5×四分位距范围；须外观测标为候选异常点。' },
  { kind: 'correlation', name: '相关性热力图', icon: Grid2X2, purpose: '比较多个指标的线性相关性', help: '选择 2～6 个数值字段；同一行须代表同一对象。跨 Sheet 先按唯一键关联，不按行号猜测对应。', method: '逐对计算 Pearson r（至少 3 对有效值）。常量或样本不足显示空值；相关不代表因果。' },
] as const;
type Card = AnalysisCard & { result?: AnalysisResult };
function initialPlan(): AnalysisPlan {
  try { const stored = localStorage.getItem(storageKey); if (stored) return parseAnalysisPlan(JSON.parse(stored)); } catch { /* 旧版或损坏草稿不执行，用户仍可导入已保存的方案。 */ }
  return { version: 2, canvas: { ...DASHBOARD_CANVAS }, title: '我的统计分析看板', cards: [] };
}
function blankDraft(inputPath = '', blocks: AnalysisBlock[] = []): AnalysisRequest {
  return { inputPath, blocks, token: '', kind: 'bar', title: '分类对比', unit: '', note: '', combine: 'single', joinKeys: {}, joinType: 'inner', x: '', y: [], aggregation: 'sum', sort: 'source', timeBucket: 'none', bins: 12, limit: 20, fieldLabels: {} };
}
function fieldsFor(schema: AnalysisSchema | null, mode: AnalysisRequest['combine']): AnalysisField[] {
  if (!schema) return [];
  return mode === 'join' ? schema.blocks.flatMap((b) => b.fields.map((f) => ({ ...f, name: `${b.name} / ${f.name}` }))) : schema.blocks[0]?.fields || [];
}
function Table({ table }: { table: AnalysisTable }) {
  return <div className="analysis-table-scroll"><table><thead><tr>{table.columns.map((c, i) => <th key={i}>{c}</th>)}</tr></thead><tbody>{table.rows.map((row, r) => <tr key={r}>{row.map((v, c) => <td key={c}>{display(v)}</td>)}</tr>)}</tbody></table></div>;
}
function ResultView({ result, compact = false }: { result: AnalysisResult; compact?: boolean }) {
  return <div className="analysis-result">
    {result.image ? <img src={result.image} alt={result.title} /> : <div className="analysis-metric"><strong>{display(result.metric)}</strong><span>{result.unit}</span><small>{result.table.rows[0]?.[1]} · {result.table.rows[0]?.[3]} 个有效数值</small></div>}
    <p className="analysis-method">{result.method}</p>{result.note && <p className="analysis-note">{result.note}</p>}
    <div className="analysis-provenance">{result.sources.map((s, i) => <span key={i}>{s.file} / {s.sheet}!{s.range} · {s.rows} 条记录</span>)}<span>组合后 {result.counts.inputRows} 条 · {new Date(result.generatedAt).toLocaleString('zh-CN')}</span></div>
    {!!result.warnings.length && <details className="analysis-warnings" open={!compact}><summary>{result.warnings.length} 条数据处理说明</summary><ul>{result.warnings.map((w, i) => <li key={i}>{w}</li>)}</ul></details>}
    <details className="analysis-details"><summary>查看图表数据与描述统计</summary>{result.kind !== 'box' && <Table table={result.table} />}<p>描述统计基于所选数值字段的全部输入记录，使用有效数值；标准差为样本标准差，缺失不补零。</p><Table table={result.statistics} /></details>
  </div>;
}

export default function AnalysisWorkspace({ source, ready, reading, onChoose, onSourceChange, onBusy }: {
  source: SourceInfo | null; ready: boolean; reading: boolean; onChoose: (kind: 'file' | 'directory') => void; onSourceChange: (path: string) => void; onBusy: (busy: boolean) => void;
}) {
  const [saved] = useState(initialPlan);
  const [title, setTitle] = useState(saved.title), [cards, setCards] = useState<Card[]>(saved.cards);
  const [draft, setDraft] = useState<AnalysisRequest>(() => blankDraft());
  const [schema, setSchema] = useState<AnalysisSchema | null>(null), [schemaKey, setSchemaKey] = useState('');
  const [preview, setPreview] = useState<AnalysisResult | null>(null), [editingId, setEditingId] = useState<string | null>(null);
  const [busy, setBusy] = useState(''), [error, setError] = useState(''), [notice, setNotice] = useState('');
  const [picker, setPicker] = useState<string | null>(null), [cardErrors, setCardErrors] = useState<Record<string, string>>({});
  const [output, setOutput] = useState('');
  const [boardEpoch, setBoardEpoch] = useState(0);
  const editor = useRef<HTMLDivElement>(null), cancelled = useRef(false);
  const choices = useMemo(() => source?.workbooks.flatMap((b) => b.sheets.filter((s) => s.state === 'visible').map((s) => ({ ...s, path: b.path, file: b.name }))) || [], [source]);
  const signature = JSON.stringify({ inputPath: draft.inputPath, blocks: draft.blocks });
  const fresh = schema !== null && schemaKey === signature;
  const fields = fieldsFor(fresh ? schema : null, draft.combine);
  const template = templates.find((t) => t.kind === draft.kind)!;
  const singleY = ['metric', 'pie', 'histogram', 'scatter'].includes(draft.kind);
  const needsX = ['bar', 'line', 'pie', 'scatter'].includes(draft.kind);
  const plan = (): AnalysisPlan => ({ version: 2, canvas: { ...DASHBOARD_CANVAS }, title, cards: cards.map(({ id, config, layout }) => ({ id, config, layout })) });
  useEffect(() => { onBusy(!!busy); return () => onBusy(false); }, [busy, onBusy]);
  useEffect(() => { try { localStorage.setItem(storageKey, JSON.stringify({ version: 2, canvas: { ...DASHBOARD_CANVAS }, title, cards: cards.map(({ id, config, layout }) => ({ id, config, layout })) })); } catch { /* 可通过保存方案按钮落盘。 */ } }, [title, cards]);
  function newBlock(index = 0): AnalysisBlock {
    const sheet = choices[0];
    const header = Math.max(1, sheet?.headerRow || 1), columns = Math.max(1, sheet?.columns.length || 1);
    return { id: id(), name: `数据块 ${index + 1}`, path: sheet?.path || '', sheet: sheet?.name || '',
      range: sheet?.estimatedRows == null ? '' : `A${header}:${columnLabel(columns)}${header + Math.max(1, sheet.estimatedRows)}`,
      headerRows: 1, orientation: 'columns', includeHidden: false, skipTotals: true };
  }
  useEffect(() => {
    if (source && source.path !== draft.inputPath) {
      setDraft(blankDraft(source.path, choices.length ? [newBlock()] : [])); setSchema(null); setPreview(null); setEditingId(null); setError('');
    }
  }, [source?.path, choices]); // 数据源更换只影响编辑区，已有看板卡片保留自己的来源。
  function patch(changes: Partial<AnalysisRequest>) { setDraft((d) => ({ ...d, ...changes })); setPreview(null); setError(''); }
  function changeBlock(blockId: string, changes: Partial<AnalysisBlock>) { patch({ blocks: draft.blocks.map((b) => b.id === blockId ? { ...b, ...changes } : b) }); }
  function startNew(kind: ChartKind = 'bar') {
    const t = templates.find((t) => t.kind === kind)!;
    setDraft({ ...blankDraft(source?.path, choices.length ? [newBlock()] : []), kind, title: t.name });
    setSchema(null); setPreview(null); setEditingId(null); setError(''); editor.current?.scrollIntoView({ behavior: 'smooth' });
  }
  async function suggested(block: AnalysisBlock) {
    if (!api) return; setBusy('识别区域'); setError('');
    try {
      const found = choices.find((s) => s.path === block.path && s.name === block.sheet);
      const header = Math.max(1, found?.headerRow || 1);
      const value = await api.analysisGrid({ inputPath: draft.inputPath, path: block.path, sheet: block.sheet, startRow: 1, startColumn: 1, headerRow: header });
      changeBlock(block.id, { range: value.suggestedRange, headerRows: value.suggestedHeaderRows });
      setNotice('已按检测到的列名和合并表头建议区域。并排表和合计行请在选区中调整，避免重复统计。');
    } catch (e) { setError(message(e)); } finally { setBusy(''); }
  }
  async function inspect() {
    if (!api) return; setBusy('读取数据'); setError(''); setPreview(null);
    try {
      const value = await api.analysisInspect({ inputPath: draft.inputPath, blocks: draft.blocks });
      setSchema(value); setSchemaKey(signature);
      const all = fieldsFor(value, draft.combine), names = Object.fromEntries(all.map((f) => [f.id, f.name]));
      const changed = Object.entries(draft.fieldLabels).some(([key, value]) => key in names && names[key] !== value);
      const ys = draft.y.filter((c) => c in names);
      const suggestedY = all.find((f) => f.kind === 'number')?.id || '';
      const suggestedX = all.find((f) => f.kind === 'date' || f.kind === 'text')?.id || all[0]?.id || '';
      setDraft((d) => ({ ...d, token: value.token, fieldLabels: names, x: d.x in names && !changed ? d.x : suggestedX,
        y: changed ? [] : ys.length ? ys : suggestedY ? [suggestedY] : [],
        joinKeys: Object.fromEntries(value.blocks.map((b) => [b.id, d.joinKeys[b.id] || b.fields[0]?.id || ''])) }));
      if (changed) setNotice('检测到字段名称变化，已清空数值选择，请重新核对字段。');
    } catch (e) { setError(message(e)); setSchema(null); } finally { setBusy(''); }
  }
  async function generate() {
    if (!api || !fresh) return; setBusy('生成图表'); setError(''); cancelled.current = false;
    try { const value = await api.analysisGenerate({ ...draft, fieldLabels: Object.fromEntries(fields.map((f) => [f.id, f.name])) }); if (!cancelled.current) setPreview(value); }
    catch (e) { setError(cancelled.current ? '已停止生成' : message(e)); } finally { setBusy(''); }
  }
  function addToBoard() {
    if (!preview) return;
    const config = { ...draft, fieldLabels: Object.fromEntries(fields.map((f) => [f.id, f.name])) };
    if (editingId) {
      setCards((old) => old.map((c) => c.id === editingId ? { ...c, config, result: preview } : c));
      setCardErrors((old) => { const next = { ...old }; delete next[editingId]; return next; });
    }
    else setCards((old) => [...old, { id: id(), config, result: preview, layout: nextLayout(old.map((c) => c.layout), config.kind === 'metric') }]);
    setNotice(editingId ? '看板图表已更新' : '图表已加入看板'); setEditingId(null); setPreview(null);
  }
  function editCard(card: Card) {
    setDraft(card.config); setEditingId(card.id); setSchema(null); setPreview(null); setError('');
    if (source?.path !== card.config.inputPath) onSourceChange(card.config.inputPath);
    editor.current?.scrollIntoView({ behavior: 'smooth' });
  }
  async function refreshCards(targets: Card[]) {
    if (!api) return; setBusy('刷新看板'); setError(''); cancelled.current = false;
    for (const card of targets) {
      if (cancelled.current) break;
      try {
        const info = await api.analysisInspect({ inputPath: card.config.inputPath, blocks: card.config.blocks });
        if (cancelled.current) break;
        const config = { ...card.config, token: info.token };
        const result = await api.analysisGenerate(config);
        if (!cancelled.current) { setCards((old) => old.map((c) => c.id === card.id ? { ...c, config, result } : c)); setCardErrors((old) => { const next = { ...old }; delete next[card.id]; return next; }); }
      } catch (e) { setCardErrors((old) => ({ ...old, [card.id]: cancelled.current ? '已停止刷新' : message(e) })); }
    }
    setBusy('');
  }
  async function importPlan() {
    if (!api) return; setError('');
    try {
      const value = await api.analysisImportPlan(); if (!value) return;
      setTitle(value.title); setCards(value.cards); setCardErrors({}); setOutput(''); setBoardEpoch((old) => old + 1);
      setNotice('已导入配置。点击“刷新全部图表”读取本机数据；不会自动执行脚本。');
    } catch (e) { setError(message(e)); }
  }
  async function exportReport() {
    if (!api) return; setBusy('导出报告'); setError(''); cancelled.current = false;
    try { const value = await api.analysisExport(plan()); if (value) { setOutput(value.output); setNotice(`已导出 ${value.charts} 张图表、描述统计 CSV、离线 HTML 和分析方案。`); } }
    catch (e) { setError(cancelled.current ? '已停止导出' : message(e)); } finally { setBusy(''); }
  }
  return <section className="analysis-workspace" aria-label="统计分析工作台">
    <div className="analysis-source-bar"><div><strong>{reading ? '正在读取数据源…' : source?.name || '选择要分析的 Excel'}</strong><span>{source ? `${source.workbooks.length} 个文件 · ${choices.length} 张可见 Sheet，图表可分别选区` : '可使用一个工作簿，或包含多个工作簿的文件夹'}</span></div><button className="button" disabled={!!busy || reading} onClick={() => onChoose('file')}><FolderOpen size={15} />{source ? '更换分析文件' : '选择分析文件'}</button><button className="button" disabled={!!busy || reading} onClick={() => onChoose('directory')}>选择分析文件夹</button></div>
    {error && <p className="analysis-error" role="alert">{error}</p>}{notice && <p className="analysis-notice" role="status">{notice}</p>}
    <div className="analysis-editor" ref={editor}>
      <div className="analysis-section-heading"><div><span className="analysis-step">01</span><strong>{editingId ? '编辑看板图表' : '选择图表，定义分析口径'}</strong></div><button className="text-button" disabled={!!busy} onClick={() => startNew()}><Plus size={14} />新建图表</button></div>
      <div className="analysis-templates">{templates.map((t) => <button className={draft.kind === t.kind ? 'active' : ''} key={t.kind} disabled={!!busy} aria-label={`选择${t.name}`} aria-pressed={draft.kind === t.kind} onClick={() => patch({ kind: t.kind, title: t.name, aggregation: t.kind === 'pie' && !['sum','count'].includes(draft.aggregation) ? 'sum' : draft.aggregation, y: ['metric','pie','histogram','scatter'].includes(t.kind) ? draft.y.slice(0, 1) : draft.y })}><t.icon size={20} /><strong>{t.name}</strong><small>{t.purpose}</small></button>)}</div>
      <div className="analysis-explanation"><strong>{template.name}怎样生成</strong><p>{template.method}</p><p><b>如何选数：</b>{template.help}</p></div>
      <div className="analysis-builder-grid"><div className="analysis-data-setup">
        <div className="analysis-section-heading"><div><span className="analysis-step">02</span><strong>选择数据区域</strong></div><button className="button small" disabled={!!busy || !source || draft.blocks.length >= 4} onClick={() => patch({ blocks: [...draft.blocks, newBlock(draft.blocks.length)], combine: draft.blocks.length ? 'append' : 'single' })}><Plus size={14} />添加区域 / Sheet</button></div>
        {!draft.blocks.length && <p className="analysis-empty">先选择数据源，再添加需要分析的区域。</p>}
        {draft.blocks.map((block, i) => <section className="analysis-block" key={block.id}>
          <div className="analysis-block-heading"><input aria-label={`数据块 ${i+1} 名称`} maxLength={80} value={block.name} disabled={!!busy} onChange={(e) => changeBlock(block.id, { name: e.target.value })} /><button className="icon-button" aria-label={`删除数据块 ${i+1}`} disabled={!!busy || draft.blocks.length === 1} onClick={() => { const blocks = draft.blocks.filter((b) => b.id !== block.id); patch({ blocks, combine: blocks.length === 1 ? 'single' : draft.combine }); }}><Trash2 size={15} /></button></div>
          <label>工作簿 / Sheet<select aria-label={`数据块 ${i+1} 工作表`} value={JSON.stringify([block.path, block.sheet])} disabled={!!busy || !source} onChange={(e) => { const [path, sheet] = JSON.parse(e.target.value) as string[]; changeBlock(block.id, { path, sheet, range: '' }); }}><option value="">选择工作表</option>{choices.map((s) => <option key={JSON.stringify([s.path,s.name])} value={JSON.stringify([s.path,s.name])}>{s.file} / {s.name}</option>)}</select></label>
          <div className="analysis-range-row"><label>数据范围<input aria-label={`数据块 ${i+1} 区域`} placeholder="例如 A6:K246 或 B:D" value={block.range} disabled={!!busy} onChange={(e) => changeBlock(block.id, { range: e.target.value })} /></label><button className="button small" disabled={!!busy || !block.path} aria-label={`选择数据块 ${i+1} 区域`} onClick={() => setPicker(block.id)}><MousePointer2 size={15} />选区</button><button className="text-button" disabled={!!busy || !block.path} onClick={() => void suggested(block)}>建议区域</button></div>
          <div className="analysis-two-fields"><label>数据排布<select aria-label={`数据块 ${i+1} 排布`} value={block.orientation} disabled={!!busy} onChange={(e) => changeBlock(block.id, { orientation: e.target.value as AnalysisBlock['orientation'] })}><option value="columns">按列 · 每行一条记录</option><option value="rows">按行 · 每列一条记录</option></select></label><label>{block.orientation === 'rows' ? '左侧表头列数' : '顶部表头行数'}<input aria-label={`数据块 ${i+1} 表头层数`} type="number" min={0} max={10} value={block.headerRows} disabled={!!busy} onChange={(e) => changeBlock(block.id, { headerRows: Number(e.target.value) })} /></label></div>
          <div className="analysis-checks"><label><input type="checkbox" checked={block.skipTotals} disabled={!!busy} onChange={(e) => changeBlock(block.id, { skipTotals: e.target.checked })} />排除合计 / 小计记录</label><label><input type="checkbox" checked={block.includeHidden} disabled={!!busy} onChange={(e) => changeBlock(block.id, { includeHidden: e.target.checked })} />包含隐藏行列</label></div>
          {fresh && schema?.blocks[i] && <details className="analysis-sample"><summary>{schema.blocks[i].rows} 条记录 · {schema.blocks[i].fields.length} 个字段 · {schema.blocks[i].resolvedRange}</summary><Table table={{ columns: schema.blocks[i].fields.map((f) => `${f.name} (${f.location})`), rows: schema.blocks[i].sample }} />{schema.blocks[i].warnings.map((w) => <p key={w}>{w}</p>)}</details>}
        </section>)}
        {draft.blocks.length > 1 && <div className="analysis-combine"><label>跨 Sheet 如何组合<select aria-label="跨表组合方式" value={draft.combine} disabled={!!busy} onChange={(e) => patch({ combine: e.target.value as AnalysisRequest['combine'], x: '', y: [], fieldLabels: {} })}><option value="append">追加同类记录 · 按相同字段名对齐</option><option value="join">按共同字段关联 · 一对一</option></select></label><p>{draft.combine === 'append' ? '适用于多月份明细等同类表，字段名须一致；按名字对齐，不按列位置猜测。' : '适用于预算表与实际表等不同指标。每张表的关联键必须唯一；不按行号配对。'}</p></div>}
        <button className="button primary" disabled={!!busy || !ready || !draft.blocks.length || draft.blocks.some((b) => !b.range.trim())} onClick={() => void inspect()}><RefreshCw size={15} />读取数据并检查字段</button>
        {!fresh && schema && <p className="analysis-warning">数据范围已变更，请重新读取后选字段。</p>}
      </div><div className="analysis-chart-setup">
        <div className="analysis-section-heading"><div><span className="analysis-step">03</span><strong>绑定字段与生成图表</strong></div></div>
        <label>图表标题<input aria-label="统计图标题" maxLength={120} value={draft.title} disabled={!!busy} onChange={(e) => patch({ title: e.target.value })} /></label>
        {draft.combine === 'join' && fresh && <div className="analysis-join-keys">{schema?.blocks.map((b) => <label key={b.id}>{b.name} 的共同字段<select aria-label={`${b.name} 关联字段`} value={draft.joinKeys[b.id] || ''} disabled={!!busy} onChange={(e) => patch({ joinKeys: { ...draft.joinKeys, [b.id]: e.target.value } })}><option value="">选择关联字段</option>{b.fields.map((f) => <option value={f.id} key={f.id}>{f.name} · {f.location}</option>)}</select></label>)}<label>保留哪些记录<select aria-label="关联保留方式" value={draft.joinType} disabled={!!busy} onChange={(e) => patch({ joinType: e.target.value as 'inner'|'left' })}><option value="inner">交集 · 只保留全部表都匹配的记录</option><option value="left">保留首表 · 未匹配指标记为缺失</option></select></label></div>}
        {needsX && <label>{draft.kind === 'scatter' ? 'X 轴 · 数值字段' : draft.kind === 'line' ? '横轴 · 时间 / 有序字段' : '分类字段'}<select aria-label="统计图分类或 X 轴" value={draft.x} disabled={!!busy || !fresh} onChange={(e) => patch({ x: e.target.value })}><option value="">选择字段</option>{fields.map((f) => <option key={f.id} value={f.id}>{f.name} · {f.location}</option>)}</select></label>}
        <fieldset className="analysis-value-fields" disabled={!!busy || !fresh}><legend>{singleY ? '数值字段 · 选 1 个' : draft.kind === 'correlation' ? '数值字段 · 选 2～6 个' : '数值字段 · 最多 6 个'}</legend>{!fields.length && <p>读取数据后，这里显示可选择的字段。</p>}{fields.map((f) => <label key={f.id}><input type="checkbox" aria-label={`数值字段 ${f.name}`} checked={draft.y.includes(f.id)} onChange={(e) => patch({ y: e.target.checked ? singleY ? [f.id] : [...draft.y, f.id].slice(0, 6) : draft.y.filter((c) => c !== f.id) })} /><span>{f.name}<small>{f.location} · { {number:'数值',date:'日期',mixed:'混合',text:'文本'}[f.kind]}</small></span></label>)}</fieldset>
        {['bar','line','pie','metric'].includes(draft.kind) && <label>统计方式<select aria-label="统计汇总方式" value={draft.aggregation} disabled={!!busy} onChange={(e) => patch({ aggregation: e.target.value as AnalysisRequest['aggregation'] })}>{Object.entries({ sum:'合计',mean:'平均值',median:'中位数',min:'最小值',max:'最大值',count:'有效数值计数' }).filter(([k]) => draft.kind !== 'pie' || ['sum','count'].includes(k)).map(([k,v]) => <option key={k} value={k}>{v}</option>)}</select></label>}
        {['bar','line','pie'].includes(draft.kind) && <div className="analysis-two-fields"><label>排序<select aria-label="图表排序" value={draft.sort} disabled={!!busy} onChange={(e) => patch({ sort: e.target.value as AnalysisRequest['sort'] })}><option value="source">首次出现顺序</option><option value="label">标签文字升序</option><option value="value_desc">第一指标降序</option></select></label><label>图上分类数<input aria-label="图上分类数" type="number" min={1} max={60} value={draft.limit} disabled={!!busy} onChange={(e) => patch({ limit: Number(e.target.value) })} /></label></div>}
        {draft.kind === 'line' && <label>日期汇总粒度<select aria-label="日期汇总粒度" value={draft.timeBucket} disabled={!!busy} onChange={(e) => patch({ timeBucket: e.target.value as AnalysisRequest['timeBucket'] })}><option value="none">不按日期重分组</option><option value="day">按日</option><option value="month">按月</option><option value="year">按年</option></select></label>}
        {draft.kind === 'histogram' && <label>等宽区间数量<input aria-label="直方图组数" type="number" min={3} max={50} value={draft.bins} disabled={!!busy} onChange={(e) => patch({ bins: Number(e.target.value) })} /></label>}
        <label>单位标注<input aria-label="统计图单位" maxLength={30} value={draft.unit} placeholder="例如 元、万元、件（仅标注，不换算）" disabled={!!busy} onChange={(e) => patch({ unit: e.target.value })} /></label>
        <label>图表注解<textarea aria-label="统计图注解" maxLength={1000} rows={2} value={draft.note} placeholder="记录统计口径、异常原因或业务解释" disabled={!!busy} onChange={(e) => patch({ note: e.target.value })} /></label>
        <p className="analysis-cache-note">使用公式已保存的结果，不执行重算。数字文本支持千分位、货币符号和百分号；无法识别的值会计入说明。</p>
        <div className="analysis-generate-actions"><button className="button primary" disabled={!!busy || !ready || !fresh || !draft.y.length || !draft.title.trim() || (needsX && !draft.x)} onClick={() => void generate()}><BarChart3 size={16} />生成图表预览</button><button className="button" disabled={!!busy || !preview || (!editingId && cards.length >= 12)} onClick={addToBoard}><Plus size={15} />{editingId ? '更新看板图表' : '加入 dashboard'}</button></div>
      </div></div>
      {preview && <section className="analysis-live-preview" aria-label="统计图生成预览"><h3>{preview.title}</h3><ResultView result={preview} /></section>}
    </div>
    {busy && <div className="analysis-busy" role="status"><LoaderCircle size={19} className="spin" />{busy}，数据在本机处理…{['生成图表','刷新看板','导出报告'].includes(busy) && <button className="button small" onClick={() => { cancelled.current = true; void api?.analysisCancel(); }}><Square size={12} />停止</button>}</div>}
    <section className="analysis-board" aria-label="统计分析 dashboard">
      <div className="analysis-board-heading"><div><span className="eyebrow">DASHBOARD · 本地看板</span><input aria-label="Dashboard 标题" value={title} maxLength={120} disabled={!!busy} onChange={(e) => setTitle(e.target.value)} /><p>{cards.length} / 12 个组件 · 配置自动保存在本机，重新打开后刷新数据</p></div><div className="analysis-board-actions"><button className="button small" disabled={!!busy} onClick={() => void importPlan()}><Upload size={14} />导入方案</button><button className="button small" disabled={!!busy || !cards.length} onClick={async () => { try { const path = await api?.analysisSavePlan(plan()); if (path) setNotice('分析方案已保存：' + path); } catch (e) { setError(message(e)); } }}><Save size={14} />保存方案</button><button className="button small" disabled={!!busy || !cards.length || !ready} onClick={() => void refreshCards(cards)}><RefreshCw size={14} />刷新全部图表</button><button className="button primary small" disabled={!!busy || !cards.length || cards.some((c) => !c.result || cardErrors[c.id]) || !title.trim()} onClick={() => void exportReport()}><Download size={14} />导出离线报告</button></div></div>
      <p className="analysis-report-note">报告包含真实统计结果、图表和来源，保存在本机；需要给 AI 的脱敏资料请使用“脱敏导出”。</p>
      {output && <button className="button small" onClick={() => void api?.openOutput(output).catch((e) => setError(message(e)))}><FolderOpen size={14} />打开报告文件夹</button>}
      {!cards.length ? <div className="analysis-empty-board"><Grid2X2 size={40} /><h3>把你的分析组合成一张看板</h3><p>上方生成预览后点击“加入 dashboard”。例如：总金额指标卡 + 部门对比 + 月度趋势。</p></div> : <DashboardCanvas key={boardEpoch} title={title} cards={cards} errors={cardErrors} busy={!!busy} ready={ready} onChange={setCards} onEdit={editCard} onRefresh={(card) => void refreshCards([card])} renderDetails={(result) => <ResultView result={result} />} />}
    </section>
    {picker && draft.blocks.find((b) => b.id === picker) && <AnalysisRangePicker inputPath={draft.inputPath} block={draft.blocks.find((b) => b.id === picker)!} onClose={() => setPicker(null)} onApply={(range) => { changeBlock(picker, { range }); setPicker(null); }} />}
  </section>;
}
