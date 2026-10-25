// Copyright (c) 2026 AbeLin · MIT License
import { useEffect, useRef, useState } from 'react';
import { ArrowDownToLine, Check, FolderOpen, LoaderCircle, ShieldCheck, X } from 'lucide-react';
import type { PrivacyMode, PrivacyPlan, PrivacyResult, SelectionRequest } from '../shared/types';
import './privacy-export.css';

const api = window.desktop;
const errorText = (error: unknown) => error instanceof Error ? error.message.replace(/^Error invoking remote method '[^']+': (?:Error: )?/, '') : String(error);

export default function PrivacyExport({ request, disabled, onBusy }: { request: SelectionRequest; disabled: boolean; onBusy: (busy: boolean) => void }) {
  const [open, setOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [plan, setPlan] = useState<PrivacyPlan | null>(null);
  const [mode, setMode] = useState<PrivacyMode>('synthetic');
  const [rows, setRows] = useState(20);
  const [kinds, setKinds] = useState<Record<string, string[]>>({});
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
  async function prepare() {
    if (!api || disabled) return;
    const generation = ++version.current;
    setOpen(true); setLoading(true); setError(''); setResult(null); setPlan(null);
    try {
      const value = await api.privacyInspect(request);
      if (generation !== version.current) return;
      setPlan(value);
      setKinds(Object.fromEntries(value.tables.map((table) => [table.id, table.fields.map((field) => field.kind)])));
    } catch (e) { if (generation === version.current) setError(errorText(e)); }
    finally { if (generation === version.current) setLoading(false); }
  }
  async function exportPackage() {
    if (!api || !plan || exporting) return;
    setExporting(true); onBusy(true); setError('');
    try {
      const value = await api.privacyExport({ ...request, token: plan.token, mode, sampleRows: rows, fieldKinds: kinds });
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
      <ShieldCheck size={24} /><div><strong>给 AI 看结构，让数据留在本机</strong><p>匿名字段 + 完全虚构样例。AI 写代码，离线电脑运行真实数据。</p></div>
      <button className="button" disabled={disabled || !api || !request.selection.length} onClick={() => void prepare()}><ArrowDownToLine size={15} />脱敏导出</button>
    </section>
    <dialog className="privacy-dialog" ref={dialog} aria-label="严格脱敏导出" onCancel={(event) => { if (exporting) event.preventDefault(); }} onClose={() => { setOpen(false); version.current += 1; }}>
      {open && <>
        <div className="card-title"><h2><ShieldCheck size={19} />严格脱敏 · AI 演示包</h2><button className="icon-button" aria-label="关闭脱敏导出" disabled={exporting} onClick={() => dialog.current?.close()}><X size={19} /></button></div>
        <div className="privacy-body">
          <p className="privacy-intro">不抽取真实记录、不使用真实统计或分布。姓名、金额、分类等示例独立生成；Sheet 和字段统一匿名。</p>
          <div className="privacy-options">
            <label>导出内容<select aria-label="脱敏导出内容" value={mode} disabled={exporting} onChange={(e) => { setMode(e.target.value as PrivacyMode); setResult(null); }}>
              <option value="synthetic">虚构样例 + 安全公式（推荐）</option><option value="blank">空白结构 + 安全公式</option><option value="schema">仅字段说明（不含公式和行数）</option>
            </select></label>
            <label>演示行数<select aria-label="脱敏演示行数" value={rows} disabled={exporting || mode === 'schema'} onChange={(e) => { setRows(Number(e.target.value)); setResult(null); }}>
              <option value={5}>5 行</option><option value={20}>20 行</option><option value={50}>50 行</option>
            </select></label>
          </div>
          <div className="privacy-policy"><Check size={16} /><span>清除原数据、统计值、图表缓存、图片、批注、链接和元数据。只保留表头布局与固定演示行，公式中的原常量留在本机；复杂公式可能移除。真实总行数不导出。</span></div>
          {loading && <p className="privacy-loading"><LoaderCircle size={19} className="spin" />正在准备匿名字段…</p>}
          {plan && <div className="privacy-fields">
            <p>下方原名称仅在本机显示。确认字段类型后生成虚构样例；类型可按用途修改。</p>
            {plan.tables.map((table) => <section key={table.id}><h3><span>{table.id}</span><small title={table.source}>{table.source}（仅本机）</small></h3>
              <table><thead><tr><th>本机字段</th><th>AI 可见名称</th><th>演示类型</th></tr></thead><tbody>{table.fields.map((field, index) => <tr key={field.id}>
                <td title={field.name}>{field.name}</td><td>{field.id}</td><td><select aria-label={`${table.id} ${field.id} 演示类型`} value={kinds[table.id]?.[index] || field.kind} disabled={exporting} onChange={(e) => { const type = e.target.value; setKinds((old) => ({ ...old, [table.id]: old[table.id].map((kind, i) => i === index ? type : kind) })); setResult(null); }}>{Object.entries(plan.kinds).map(([key, value]) => <option key={key} value={key}>{value}</option>)}</select></td>
              </tr>)}</tbody></table>
            </section>)}
          </div>}
          {error && <p className="notice error" role="alert">{error}</p>}
          {result && <div className="privacy-success" role="status"><strong>脱敏包已生成，只分享这个 ZIP</strong><p>{result.output}</p><p>真实字段与参数映射保存在本机应用目录。AI 按包内说明调用 load_ai_context 后，才能在本机处理对应真实数据。</p>{result.warnings.map((warning) => <p key={warning}>{warning}</p>)}<button className="button small" onClick={() => void openOutput()}><FolderOpen size={14} />定位脱敏包</button></div>}
        </div>
        <div className="privacy-actions"><span>分析结果和运行日志仍可能含真实数据，请留在本机。</span><button className="button primary" disabled={!plan || loading || exporting} onClick={() => void exportPackage()}>{exporting ? <LoaderCircle size={16} className="spin" /> : <ArrowDownToLine size={16} />}{exporting ? '正在生成脱敏包…' : '生成脱敏包 .zip'}</button></div>
      </>}
    </dialog>
  </>;
}
