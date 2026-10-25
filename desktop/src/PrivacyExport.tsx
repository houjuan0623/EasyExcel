// Copyright (c) 2026 AbeLin · MIT License
import { useEffect, useRef, useState } from 'react';
import { ArrowDownToLine, Check, FolderOpen, LoaderCircle, ShieldCheck, X } from 'lucide-react';
import type { PrivacyMode, PrivacyPlan, PrivacyResult, SelectionRequest, VbaReviewItem, VbaSheetPlan } from '../shared/types';
import './privacy-export.css';

const api = window.desktop;
const errorText = (error: unknown) => error instanceof Error ? error.message.replace(/^Error invoking remote method '[^']+': (?:Error: )?/, '') : String(error);

function HeaderReview({ sheet, value, disabled, onChange }: { sheet: VbaSheetPlan; value: string; disabled: boolean; onChange: (value: string) => void }) {
  const [page, setPage] = useState(0);
  const count = Math.ceil(sheet.headers.length / 40);
  useEffect(() => { setPage(0); }, [sheet.headers]);
  return <details className="vba-sheet-review">
    <summary><strong>{sheet.name}</strong><span>{sheet.state === 'visible' ? '可见' : '隐藏'} · {sheet.formulaCount} 个公式 · {sheet.dataExtent}</span></summary>
    <label className="vba-range-label">保留文字的表头区域<input aria-label={`${sheet.id} 表头区域`} value={value} disabled={disabled} placeholder="例如 B7:N9；多个区域用分号分隔" onChange={(event) => onChange(event.target.value)} /></label>
    <p>这些区域内的常量会原样导出。请排除客户、日期、金额等真实记录；多层表头和并排表头可添加多个区域。</p>
    {sheet.codeName && <p>原 VBA 工作表代号：{sheet.codeName}。包内说明优先使用原 Sheet 名定位。</p>}
    {sheet.tables.map((table) => <p key={table.name}>Excel 表对象：<strong>{table.displayName}</strong> · {table.ref}<br />字段：{table.columns.map((column) => column.name).join(' / ')}</p>)}
    <table><thead><tr><th>原单元格</th><th>将保留的表头内容</th></tr></thead><tbody>{sheet.headers.slice(page * 40, (page + 1) * 40).map((header) => <tr key={header.cell}><td>{header.cell}</td><td className="vba-wrap">{header.value}</td></tr>)}</tbody></table>
    {!sheet.headers.length && <p>当前区域没有常量表头；原公式仍会保留。</p>}
    {count > 1 && <div className="vba-pagination"><button className="button small" disabled={page === 0} onClick={() => setPage(page - 1)}>上一页表头</button><span>{page + 1} / {count} · 共 {sheet.headers.length} 个</span><button className="button small" disabled={page + 1 >= count} onClick={() => setPage(page + 1)}>下一页表头</button></div>}
  </details>;
}

function FormulaReview({ items }: { items: VbaReviewItem[] }) {
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(0);
  const filtered = items.filter((item) => [item.kind, item.value, item.formula, ...item.locations].join(' ').toLowerCase().includes(search.toLowerCase()));
  const count = Math.ceil(filtered.length / 20);
  useEffect(() => { setPage(0); }, [items, search]);
  return <section className="vba-formula-review">
    <h3>公式检查清单 <span>{items.length} 类</span></h3>
    <p>原公式不会改写。相同固定值合并显示，每类列出示例公式和最多 8 个位置；这不是匿名导出。</p>
    <input aria-label="搜索公式检查项" placeholder="搜索固定值、公式或 Sheet" value={search} onChange={(event) => setSearch(event.target.value)} />
    {!items.length && <p>未检测到公式固定值或外部引用，仍需检查原名称与表头。</p>}
    {filtered.slice(page * 20, (page + 1) * 20).map((item) => <details key={item.kind + item.value} className="vba-review-item">
      <summary><strong>{item.kind}</strong><span>{item.count} 处</span><code>{item.value}</code></summary>
      {item.formula && <pre>{item.formula}</pre>}<ul>{item.locations.map((location, index) => <li key={index}>{location}</li>)}</ul>
    </details>)}
    {count > 1 && <div className="vba-pagination"><button className="button small" disabled={page === 0} onClick={() => setPage(page - 1)}>上一页检查项</button><span>{page + 1} / {count}</span><button className="button small" disabled={page + 1 >= count} onClick={() => setPage(page + 1)}>下一页检查项</button></div>}
  </section>;
}

export default function PrivacyExport({ request, disabled, onBusy }: { request: SelectionRequest; disabled: boolean; onBusy: (busy: boolean) => void }) {
  const [open, setOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [plan, setPlan] = useState<PrivacyPlan | null>(null);
  const [mode, setMode] = useState<PrivacyMode>('vba');
  const [rows, setRows] = useState(20);
  const [kinds, setKinds] = useState<Record<string, string[]>>({});
  const [headerRanges, setHeaderRanges] = useState<Record<string, string>>({});
  const [reviewed, setReviewed] = useState(false);
  const [error, setError] = useState('');
  const [result, setResult] = useState<PrivacyResult | null>(null);
  const dialog = useRef<HTMLDialogElement>(null);
  const version = useRef(0);
  const selectionKey = JSON.stringify(request);
  useEffect(() => {
    version.current += 1;
    setPlan(null); setResult(null); setOpen(false); setLoading(false); dialog.current?.close();
  }, [selectionKey]);
  useEffect(() => { if (open) dialog.current?.showModal(); }, [open]);
  const isVba = mode === 'vba';
  const rangeDirty = !!plan?.vba?.workbooks.some((book) => book.sheets.some((sheet) => headerRanges[sheet.id] !== sheet.headerRanges));
  async function prepare(nextMode = mode, ranges?: Record<string, string>) {
    if (!api || disabled) return;
    const generation = ++version.current;
    setMode(nextMode); setOpen(true); setLoading(true); setError(''); setResult(null); setReviewed(false);
    if (!ranges) setPlan(null);
    try {
      const value = await api.privacyInspect({ ...request, mode: nextMode, ...(ranges ? { headerRanges: ranges } : {}) });
      if (generation !== version.current) return;
      setPlan(value);
      setKinds(Object.fromEntries(value.tables.map((table) => [table.id, table.fields.map((field) => field.kind)])));
      setHeaderRanges(Object.fromEntries(value.vba?.workbooks.flatMap((book) => book.sheets.map((sheet) => [sheet.id, sheet.headerRanges])) || []));
    } catch (e) { if (generation === version.current) setError(errorText(e)); }
    finally { if (generation === version.current) setLoading(false); }
  }
  async function exportPackage() {
    if (!api || !plan || loading || exporting || (isVba && (!reviewed || rangeDirty || !plan.vba))) return;
    setExporting(true); onBusy(true); setError('');
    try {
      const value = await api.privacyExport({ ...request, token: plan.token, mode, sampleRows: rows, fieldKinds: kinds,
        ...(isVba ? { headerRanges, reviewConfirmed: reviewed } : {}) });
      if (value) setResult(value);
    } catch (e) { setError(errorText(e)); }
    finally { setExporting(false); onBusy(false); }
  }
  async function openOutput() {
    if (!api || !result) return;
    try { await api.openOutput(result.output); } catch (e) { setError(errorText(e)); }
  }
  return <>
    <section className="privacy-export-card" aria-label="给 AI 的脱敏导出">
      <ShieldCheck size={24} /><div><strong>给 AI 编程，回到原表运行</strong><p>VBA 模板保留原字段、位置和公式，清空业务数据；也可选择匿名 Python 演示包。</p></div>
      <button className="button" disabled={disabled || !api || !request.selection.length} onClick={() => void prepare()}><ArrowDownToLine size={15} />脱敏导出</button>
    </section>
    <dialog className="privacy-dialog" ref={dialog} aria-label="脱敏与 VBA 编程导出" onCancel={(event) => { if (exporting) event.preventDefault(); }} onClose={() => { setOpen(false); version.current += 1; }}>
      {open && <>
        <div className="card-title"><h2><ShieldCheck size={19} />{isVba ? '原表 VBA · 结构模板' : '严格脱敏 · Python 演示包'}</h2><button className="icon-button" aria-label="关闭脱敏导出" disabled={exporting} onClick={() => dialog.current?.close()}><X size={19} /></button></div>
        <div className="privacy-body">
          <p className="privacy-intro">{isVba ? '让 AI 按真实字段和原公式编写 VBA，再粘贴到原工作簿的标准模块运行。名称、表头、位置和公式保持对应，其余数据常量及公式缓存清空。' : '不抽取真实记录、不使用真实统计或分布。姓名、金额、分类等示例独立生成；Sheet 和字段统一匿名。'}</p>
          <div className="privacy-options">
            <label>编程用途<select aria-label="脱敏导出内容" value={mode} disabled={exporting} onChange={(e) => void prepare(e.target.value as PrivacyMode)}>
              <option value="vba">VBA 原表模板 · 保留字段和原公式</option><option value="synthetic">Python 匿名样例 + 安全公式</option><option value="blank">Python 匿名空白结构 + 安全公式</option><option value="schema">仅匿名字段说明（不含公式和行数）</option>
            </select></label>
            {!isVba && <label>演示行数<select aria-label="脱敏演示行数" value={rows} disabled={exporting || mode === 'schema'} onChange={(e) => { setRows(Number(e.target.value)); setResult(null); }}>
              <option value={5}>5 行</option><option value={20}>20 行</option><option value={50}>50 行</option>
            </select></label>}
          </div>
          <div className="privacy-policy"><Check size={16} /><span>{isVba ? '保留所选文件的全部 Sheet 结构（含隐藏 Sheet）、原始行列范围、表对象和命名公式。原公式中的固定文本、数字和外部路径也会保留，请检查下方清单。图表、透视、宏代码及缓存不复制。' : '清除原数据、统计值、图表缓存、图片、批注、链接和元数据。只保留表头布局与固定演示行，公式中的原常量留在本机；复杂公式可能移除。真实总行数不导出。'}</span></div>
          {loading && <p className="privacy-loading"><LoaderCircle size={19} className="spin" />{isVba ? '正在核对全部 Sheet、表头和原公式…' : '正在准备匿名字段…'}</p>}
          {plan?.vba && isVba && <div className="vba-plan">
            <h3>1. 检查原名称和表头 <span>{plan.vba.formulaCount} 个单元格公式原样保留</span></h3>
            <p>展开每张 Sheet 查看将保留的字段。建议区域可修改；修改后点击“更新检查清单”。表头区域外的非公式数据均清空。</p>
            {plan.vba.workbooks.map((book) => <section key={book.id} className="vba-workbook"><h4>{book.name} · {book.sheets.length} 张 Sheet</h4>
              {book.sheets.map((sheet) => <HeaderReview key={sheet.id} sheet={sheet} value={headerRanges[sheet.id] || ''} disabled={loading || exporting} onChange={(value) => { setHeaderRanges((old) => ({ ...old, [sheet.id]: value })); setReviewed(false); setResult(null); }} />)}
              {!!book.definedNames.length && <details className="vba-names"><summary>保留的命名区域 / 名称公式（{book.definedNames.length}）</summary>{book.definedNames.map((name, index) => <p key={index}><strong>{name.name}</strong><code>{name.formula}</code></p>)}</details>}
            </section>)}
            <button className="button small" disabled={loading || exporting} onClick={() => void prepare('vba', headerRanges)}>更新检查清单</button>
            {rangeDirty && <p className="notice">表头区域已修改，请更新清单后检查保留内容。</p>}
            <FormulaReview items={plan.vba.review} />
            <label className="vba-review-confirm"><input type="checkbox" aria-label="已检查 VBA 导出保留内容" checked={reviewed} disabled={loading || exporting || rangeDirty} onChange={(event) => setReviewed(event.target.checked)} /><span>我已检查原名称、表头、公式固定值和外部引用，允许它们进入这个结构包；其余真实数据应清空。</span></label>
          </div>}
          {plan && !isVba && <div className="privacy-fields">
            <p>下方原名称仅在本机显示。确认字段类型后生成虚构样例；类型可按用途修改。</p>
            {plan.tables.map((table) => <section key={table.id}><h3><span>{table.id}</span><small title={table.source}>{table.source}（仅本机）</small></h3>
              <table><thead><tr><th>本机字段</th><th>AI 可见名称</th><th>演示类型</th></tr></thead><tbody>{table.fields.map((field, index) => <tr key={field.id}>
                <td title={field.name}>{field.name}</td><td>{field.id}</td><td><select aria-label={`${table.id} ${field.id} 演示类型`} value={kinds[table.id]?.[index] || field.kind} disabled={exporting} onChange={(e) => { const type = e.target.value; setKinds((old) => ({ ...old, [table.id]: old[table.id].map((kind, i) => i === index ? type : kind) })); setResult(null); }}>{Object.entries(plan.kinds).map(([key, value]) => <option key={key} value={key}>{value}</option>)}</select></td>
              </tr>)}</tbody></table>
            </section>)}
          </div>}
          {error && <p className="notice error" role="alert">{error}</p>}
          {result && <div className="privacy-success" role="status"><strong>{result.mode === 'vba' ? 'VBA 结构包已生成' : '脱敏包已生成，只分享这个 ZIP'}</strong><p>{result.output}</p><p>{result.mode === 'vba' ? `${result.sheets} 张 Sheet、${result.formulaCount} 个单元格公式。将结构模板、结构说明和 VBA 编程说明交给 AI，生成的完整模块粘贴到原工作簿。无需匿名字段映射。` : '真实字段与参数映射保存在本机应用目录。AI 按包内说明调用 load_ai_context 后，才能在本机处理对应真实数据。'}</p>{result.warnings.map((warning) => <p key={warning}>{warning}</p>)}<button className="button small" onClick={() => void openOutput()}><FolderOpen size={14} />定位导出包</button></div>}
        </div>
        <div className="privacy-actions"><span>{isVba ? '生成代码前补充业务目标；先在原工作簿副本验证 VBA。' : '分析结果和运行日志仍可能含真实数据，请留在本机。'}</span><button className="button primary" disabled={!plan || loading || exporting || (isVba && (!reviewed || rangeDirty || !plan.vba))} onClick={() => void exportPackage()}>{exporting ? <LoaderCircle size={16} className="spin" /> : <ArrowDownToLine size={16} />}{exporting ? '正在生成导出包…' : isVba ? '生成 VBA 结构包 .zip' : '生成脱敏包 .zip'}</button></div>
      </>}
    </dialog>
  </>;
}
