// Copyright (c) 2026 AbeLin · MIT License
import { useEffect, useRef, useState, type ReactNode, type PointerEvent as ReactPointerEvent } from 'react';
import { ArrowDown, ArrowUp, Expand, Eye, Grid2X2, GripHorizontal, Pencil, RefreshCw, RotateCcw, RotateCw, Trash2, X } from 'lucide-react';
import type { AnalysisCard, AnalysisResult } from '../shared/analysis';
import { arrangeLayouts, constrainLayout, transformLayout, type ChartLayout, type ResizeHandle } from '../shared/dashboard-layout';
import '../shared/dashboard-canvas.css';
import './dashboard.css';

type Card = AnalysisCard & { result?: AnalysisResult };
type Snapshot = { id: string; layout: ChartLayout }[];
type Gesture = { id: string; original: ChartLayout; current: ChartLayout; x: number; y: number; scale: number; mode: ResizeHandle | 'move' };
const handles: ResizeHandle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const snapshot = (cards: Card[]): Snapshot => cards.map(({ id, layout }) => ({ id, layout: { ...layout } }));
const number = (value: number | null) => value === null ? '—' : new Intl.NumberFormat('zh-CN', { maximumFractionDigits: 4 }).format(value);

function Modal({ title, children, onClose, presentation = false }: { title: string; children: ReactNode; onClose: () => void; presentation?: boolean }) {
  const ref = useRef<HTMLDialogElement>(null);
  useEffect(() => { ref.current?.showModal(); }, []);
  return <dialog ref={ref} className={`dashboard-dialog ${presentation ? 'dashboard-presentation' : ''}`} onCancel={onClose}>
    <header><strong>{title}</strong><button className="icon-button" aria-label="关闭看板窗口" onClick={onClose}><X size={20} /></button></header>{children}
  </dialog>;
}

// 同一画布只使用一个缩放系数；预览窗口多余的空间自然留白。
function FitCanvas({ children, grid = false }: { children: ReactNode; grid?: boolean }) {
  const host = useRef<HTMLDivElement>(null), [scale, setScale] = useState(0.5);
  useEffect(() => {
    const observer = new ResizeObserver(([entry]) => {
      const { width, height } = entry.contentRect;
      if (width > 0 && height > 0) setScale(Math.min(width / 1920, height / 1080));
    });
    if (host.current) observer.observe(host.current);
    return () => observer.disconnect();
  }, []);
  return <div className="dashboard-fit" ref={host} data-scale={scale}>
    <div className="dashboard-scaled" style={{ width: 1920 * scale, height: 1080 * scale }}>
      <div className={`dashboard-artboard ${grid ? 'dashboard-grid' : ''}`} style={{ transform: `scale(${scale})` }}>{children}</div>
    </div>
  </div>;
}

function TileBody({ card, error }: { card: Card; error?: string }) {
  const result = card.result;
  return <>
    <div className="dashboard-tile-body">{result
      ? result.image ? <img src={result.image} alt={result.title} draggable={false} /> : <div className="dashboard-metric"><strong>{number(result.metric)}</strong><span>{result.unit}</span></div>
      : <div className="analysis-card-pending">刷新图表后显示本机数据</div>}</div>
    {card.config.note && <p className="dashboard-tile-note" title={card.config.note}>{card.config.note}</p>}
    {error ? <p className="dashboard-tile-status error" title={error}>刷新失败 · 当前为上次结果，请查看详情</p>
      : !!result?.warnings.length && <p className="dashboard-tile-status">{result.warnings.length} 条数据处理说明 · 查看详情</p>}
  </>;
}

function LayoutFields({ value, disabled, onChange }: { value: ChartLayout; disabled: boolean; onChange: (v: ChartLayout) => void }) {
  const [draft, setDraft] = useState<Record<keyof ChartLayout, string>>({ x: '', y: '', width: '', height: '' });
  const reset = () => setDraft({ x: String(value.x), y: String(value.y), width: String(value.width), height: String(value.height) });
  useEffect(reset, [value.x, value.y, value.width, value.height]);
  function apply(key: keyof ChartLayout) {
    const v = Number(draft[key]);
    if (draft[key].trim() && Number.isFinite(v)) {
      const next = constrainLayout({ ...value, [key]: v });
      onChange(next); setDraft({ x: String(next.x), y: String(next.y), width: String(next.width), height: String(next.height) });
    } else reset();
  }
  return <div className="dashboard-layout-fields">{([['x', 'X'], ['y', 'Y'], ['width', '宽度'], ['height', '高度']] as const).map(([key, label]) =>
    <label key={key}>{label}<input type="number" aria-label={`图表${label}`} value={draft[key]} disabled={disabled}
      onChange={(e) => setDraft((old) => ({ ...old, [key]: e.target.value }))} onBlur={() => apply(key)}
      onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); }} /></label>)}</div>;
}

export default function DashboardCanvas({ title, cards, errors, busy, ready, onChange, onEdit, onRefresh, renderDetails }: {
  title: string; cards: Card[]; errors: Record<string, string>; busy: boolean; ready: boolean;
  onChange: (cards: Card[]) => void; onEdit: (card: Card) => void; onRefresh: (card: Card) => void;
  renderDetails: (result: AnalysisResult) => ReactNode;
}) {
  const [selected, setSelected] = useState(cards.at(-1)?.id || '');
  const [editing, setEditing] = useState(true), [snap, setSnap] = useState(true);
  const [presenting, setPresenting] = useState(false), [detailId, setDetailId] = useState<string | null>(null);
  const [history, setHistory] = useState<{ past: Snapshot[]; future: Snapshot[] }>({ past: [], future: [] });
  const [drag, setDrag] = useState<{ id: string; layout: ChartLayout } | null>(null);
  const gesture = useRef<Gesture | null>(null), lastIds = useRef(cards.map((c) => c.id));
  const membership = cards.map((c) => c.id).sort().join(',');
  useEffect(() => {
    const added = cards.find((c) => !lastIds.current.includes(c.id));
    if (added) setSelected(added.id);
    else if (!cards.some((c) => c.id === selected)) setSelected(cards.at(-1)?.id || '');
    lastIds.current = cards.map((c) => c.id); setHistory({ past: [], future: [] });
    gesture.current = null; setDrag(null);
  }, [membership]);
  useEffect(() => {
    const cancel = (e: KeyboardEvent) => { if (e.key === 'Escape' && gesture.current) { gesture.current = null; setDrag(null); } };
    window.addEventListener('keydown', cancel); return () => window.removeEventListener('keydown', cancel);
  }, []);
  useEffect(() => { if (busy) { gesture.current = null; setDrag(null); } }, [busy]);
  const active = cards.find((c) => c.id === selected), activeIndex = cards.findIndex((c) => c.id === selected);
  const detail = cards.find((c) => c.id === detailId);

  function commit(next: Card[]) {
    if (JSON.stringify(snapshot(next)) === JSON.stringify(snapshot(cards))) return;
    setHistory((old) => ({ past: [...old.past.slice(-29), snapshot(cards)], future: [] })); onChange(next);
  }
  function changeLayout(cardId: string, layout: ChartLayout) { commit(cards.map((c) => c.id === cardId ? { ...c, layout } : c)); }
  function travel(undo: boolean) {
    const stack = undo ? history.past : history.future, target = stack.at(-1);
    if (!target) return;
    const current = snapshot(cards), byId = new Map(cards.map((c) => [c.id, c]));
    onChange(target.map(({ id, layout }) => ({ ...byId.get(id)!, layout })));
    setHistory(undo ? { past: history.past.slice(0, -1), future: [...history.future, current] }
      : { past: [...history.past, current], future: history.future.slice(0, -1) });
  }
  function layer(delta: number) {
    const next = [...cards], target = Math.max(0, Math.min(cards.length - 1, activeIndex + delta));
    const [item] = next.splice(activeIndex, 1); next.splice(target, 0, item); commit(next);
  }
  function start(e: ReactPointerEvent<HTMLElement>, card: Card, mode: Gesture['mode']) {
    if (!editing || busy || e.button !== 0) return;
    e.preventDefault(); e.stopPropagation(); setSelected(card.id);
    e.currentTarget.closest<HTMLElement>('.dashboard-tile')?.focus({ preventScroll: true });
    const stage = e.currentTarget.closest('.dashboard-artboard')!.getBoundingClientRect();
    gesture.current = { id: card.id, original: card.layout, current: card.layout, x: e.clientX, y: e.clientY, scale: stage.width / 1920, mode };
    e.currentTarget.setPointerCapture(e.pointerId);
  }
  function move(e: ReactPointerEvent<HTMLElement>) {
    const g = gesture.current; if (!g) return;
    g.current = transformLayout(g.original, (e.clientX - g.x) / g.scale, (e.clientY - g.y) / g.scale, g.mode, snap && !e.altKey);
    setDrag({ id: g.id, layout: g.current });
  }
  function finish(cancel = false) {
    const g = gesture.current; gesture.current = null; setDrag(null);
    if (g && !cancel) changeLayout(g.id, g.current);
  }
  const pointer = { onPointerMove: move, onPointerUp: () => finish(), onPointerCancel: () => finish(true), onLostPointerCapture: () => finish(true) };
  function artboard(interactive: boolean) {
    return <><header className="dashboard-canvas-heading"><span className="dashboard-brand">EASYEXCEL</span><h2 className="dashboard-title">{title || '统计分析看板'}</h2></header>
      {cards.map((card, index) => {
        const layout = interactive && drag?.id === card.id ? drag.layout : card.layout;
        const selectedCard = interactive && editing && card.id === selected;
        return <article key={card.id} data-card-id={card.id} aria-label={`看板图表 ${index + 1}：${card.config.title}`}
          className={`dashboard-tile analysis-dashboard-card ${card.config.kind === 'metric' ? 'metric-card' : ''} ${selectedCard ? 'selected' : ''}`}
          style={{ left: layout.x, top: layout.y, width: layout.width, height: layout.height, zIndex: index + 1 }} tabIndex={interactive ? 0 : undefined}
          onClick={() => interactive && setSelected(card.id)} onDoubleClick={() => setDetailId(card.id)}
          onKeyDown={(e) => {
            if (!interactive || !editing || busy || !['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(e.key)) return;
            e.preventDefault(); setSelected(card.id); const step = e.shiftKey ? 12 : 1;
            changeLayout(card.id, transformLayout(card.layout, e.key === 'ArrowLeft' ? -step : e.key === 'ArrowRight' ? step : 0,
              e.key === 'ArrowUp' ? -step : e.key === 'ArrowDown' ? step : 0, 'move', false));
          }}>
          <header className={`dashboard-tile-header ${interactive && editing && !busy ? 'draggable' : ''}`}
            onPointerDown={(e) => interactive && start(e, card, 'move')} {...pointer}>
            {interactive && editing && <GripHorizontal size={22} className="dashboard-grip" />}<h3 title={card.config.title}>{card.config.title}</h3>
          </header><TileBody card={card} error={errors[card.id]} />
          {selectedCard && !busy && handles.map((handle) => <span key={handle} className={`dashboard-handle handle-${handle}`} data-handle={handle}
            onPointerDown={(e) => start(e, card, handle)} {...pointer} />)}
        </article>;
      })}</>;
  }
  return <div className="dashboard-layout-editor">
    <div className="dashboard-layout-toolbar"><div className="dashboard-size-label"><strong>16:9</strong><span>1920 × 1080 · 等比缩放</span></div>
      <label className="dashboard-snap"><input type="checkbox" checked={snap} disabled={!editing || busy} onChange={(e) => setSnap(e.target.checked)} />吸附网格</label>
      <button className="button small" disabled={busy || !editing || !cards.length} onClick={() => { const layouts = arrangeLayouts(cards.length); commit(cards.map((c, i) => ({ ...c, layout: layouts[i] }))); }}><Grid2X2 size={14} />自动排列</button>
      <button className="icon-button" aria-label="撤销布局" disabled={busy || !editing || !history.past.length} onClick={() => travel(true)}><RotateCcw size={16} /></button>
      <button className="icon-button" aria-label="重做布局" disabled={busy || !editing || !history.future.length} onClick={() => travel(false)}><RotateCw size={16} /></button>
      <button className="button small" disabled={busy} onClick={() => setEditing(!editing)}>{editing ? <Eye size={14} /> : <Pencil size={14} />}{editing ? '预览布局' : '编辑布局'}</button>
      <button className="button small" onClick={() => setPresenting(true)}><Expand size={14} />大屏预览</button>
    </div>
    <p className="dashboard-help">{editing ? '拖动图表标题调整位置，拖动边框控制点改变大小。方向键微调 1 px，Shift + 方向键移动 12 px；Alt 拖动临时关闭吸附。双击图表查看完整数据和注解。' : '布局已锁定，双击图表查看完整数据和注解。'}</p>
    <div className="dashboard-editor-viewport"><FitCanvas grid={editing && snap}>{artboard(true)}</FitCanvas></div>
    {active && <div className="dashboard-inspector">
      <label className="dashboard-card-select">选中图表<select aria-label="布局图表" value={selected} disabled={busy} onChange={(e) => setSelected(e.target.value)}>{cards.map((c, i) => <option key={c.id} value={c.id}>{i + 1}. {c.config.title}</option>)}</select></label>
      <LayoutFields key={active.id} value={drag?.id === active.id ? drag.layout : active.layout} disabled={busy || !editing} onChange={(layout) => changeLayout(active.id, layout)} />
      <div className="dashboard-inspector-actions">
        <button className="button small" disabled={busy || !editing || activeIndex === 0} onClick={() => layer(-1)}><ArrowDown size={14} />后移一层</button>
        <button className="button small" disabled={busy || !editing || activeIndex === cards.length - 1} onClick={() => layer(1)}><ArrowUp size={14} />前移一层</button>
        <button className="button small" onClick={() => setDetailId(active.id)}><Eye size={14} />查看详情</button>
        <button className="icon-button" aria-label={`编辑图表 ${activeIndex + 1}`} disabled={busy} onClick={() => onEdit(active)}><Pencil size={15} /></button>
        <button className="icon-button" aria-label={`刷新图表 ${activeIndex + 1}`} disabled={busy || !ready} onClick={() => onRefresh(active)}><RefreshCw size={15} /></button>
        <button className="icon-button" aria-label={`删除图表 ${activeIndex + 1}`} disabled={busy} onClick={() => onChange(cards.filter((c) => c.id !== active.id))}><Trash2 size={15} /></button>
      </div>
    </div>}
    {active && errors[active.id] && <p className="analysis-error">{errors[active.id]}。画布如有图表，是上次成功生成的结果。</p>}
    <p className="dashboard-footnote">布局自动保存。新增图表会放入空位；空间不足时叠放在上层，可使用“自动排列”整理。导出的 HTML 保留此布局和完整统计明细。</p>
    {presenting && <Modal title="16:9 看板预览 · Esc 退出" presentation onClose={() => setPresenting(false)}><FitCanvas>{artboard(false)}</FitCanvas></Modal>}
    {detail && <Modal title={detail.config.title} onClose={() => setDetailId(null)}><div className="dashboard-detail-content">{errors[detail.id] && <p className="analysis-error">{errors[detail.id]}</p>}{detail.result ? renderDetails(detail.result) : <p>刷新图表后可查看完整数据、来源和统计口径。</p>}</div></Modal>}
  </div>;
}
