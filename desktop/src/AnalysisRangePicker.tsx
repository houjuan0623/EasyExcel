// Copyright (c) 2026 AbeLin · MIT License
import { useEffect, useRef, useState } from 'react';
import { ChevronLeft, ChevronRight, LoaderCircle, Minus, Plus, RefreshCw, X } from 'lucide-react';
import type { AnalysisBlock, AnalysisGrid, NativePreview } from '../shared/analysis';
import NativeSelectionGrid, { columnLabel, type GridPosition } from './NativeSelectionGrid';
export { columnLabel } from './NativeSelectionGrid';
const columnIndex = (label: string) => [...label.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const errorText = (e: unknown) => e instanceof Error ? e.message.replace(/^Error invoking remote method '[^']+': (?:Error: )?/, '') : String(e);
type Bounds = [number, number, number, number];
function rangeBounds(value: string): Bounds | null {
  const s = value.trim().toUpperCase().replaceAll('$', '');
  let m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(s);
  if (m) return [columnIndex(m[1]), Number(m[2]), columnIndex(m[3] || m[1]), Number(m[4] || m[2])];
  m = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(s); if (m) return [columnIndex(m[1]), 1, columnIndex(m[2]), 1048576];
  m = /^(\d+):(\d+)$/.exec(s); if (m) return [1, Number(m[1]), 16384, Number(m[2])];
  return null;
}
export default function AnalysisRangePicker({ inputPath, block, onClose, onApply }: {
  inputPath: string; block: AnalysisBlock; onClose: () => void; onApply: (range: string) => void;
}) {
  const dialog = useRef<HTMLDialogElement>(null), viewport = useRef<HTMLDivElement>(null);
  const [mode, setMode] = useState<'native' | 'basic'>('native');
  const [grid, setGrid] = useState<AnalysisGrid | null>(null), [native, setNative] = useState<NativePreview | null>(null);
  const [top, setTop] = useState(1), [left, setLeft] = useState(1), [reload, setReload] = useState(0);
  const [jumpRow, setJumpRow] = useState('1'), [jumpCol, setJumpCol] = useState('A');
  const [range, setRange] = useState(block.range), [zoom, setZoom] = useState(100), [showFormulas, setShowFormulas] = useState(false);
  const [error, setError] = useState(''), [loading, setLoading] = useState(false), [imageReady, setImageReady] = useState(false);
  const [active, setActive] = useState<{ address: string; formula: string; text: string; missingCache?: boolean } | null>(null);
  const anchor = useRef<GridPosition | null>(null), dragging = useRef(false), previousRows = useRef<number[]>([]);
  const [bounds, setBounds] = useState<Bounds | null>(() => rangeBounds(block.range));
  const ready = !loading && (mode === 'basic' ? !!grid : !!native && imageReady);
  const extent = mode === 'native' ? native : grid;
  useEffect(() => { dialog.current?.showModal(); const up = () => { dragging.current = false; }; window.addEventListener('pointerup', up); return () => window.removeEventListener('pointerup', up); }, []);
  useEffect(() => {
    let current = true; setLoading(true); setError(''); setNative(null); setGrid(null); setImageReady(false); setActive(null);
    const request = { inputPath, path: block.path, sheet: block.sheet, startRow: top, startColumn: left, includeHidden: block.includeHidden, showFormulas };
    const run = mode === 'native' ? window.desktop?.analysisNativePreview(request).then(value => { if (current) setNative(value); }) : window.desktop?.analysisGrid(request).then(value => { if (current) setGrid(value); });
    void run?.catch(e => { if (current) setError(errorText(e)); }).finally(() => { if (current) { setLoading(false); viewport.current?.scrollTo(0, 0); } });
    return () => { current = false; };
  }, [inputPath, block.path, block.sheet, block.includeHidden, top, left, mode, showFormulas, reload]);
  function extend(position: GridPosition) {
    if (!anchor.current || anchor.current.type !== position.type) return;
    const a = anchor.current, r1 = Math.min(a.row, position.row), r2 = Math.max(a.endRow || a.row, position.endRow || position.row);
    const c1 = Math.min(a.col, position.col), c2 = Math.max(a.endCol || a.col, position.endCol || position.col);
    if (position.type === 'row') { setRange(`${r1}:${r2}`); setBounds([1, r1, 16384, r2]); }
    else if (position.type === 'column') { setRange(`${columnLabel(c1)}:${columnLabel(c2)}`); setBounds([c1, 1, c2, 1048576]); }
    else {
      setRange(`${columnLabel(c1)}${r1}:${columnLabel(c2)}${r2}`); setBounds([c1, r1, c2, r2]);
      const address = `${columnLabel(position.col)}${position.row}`;
      setActive({ address, formula: '', text: '', ...native?.cells[address] });
    }
  }
  function select(position: GridPosition, shift: boolean) {
    if (!ready) return;
    if (!shift || anchor.current?.type !== position.type) anchor.current = position;
    extend(position);
  }
  const pointer = (position: GridPosition) => ({
    onPointerDown: (event: React.PointerEvent) => { if (event.button !== 0 || !ready) return; event.preventDefault(); dragging.current = true; select(position, event.shiftKey); },
    onPointerEnter: () => { if (dragging.current && ready) extend(position); },
  });
  const nextRow = native ? native.endRow + 1 : top + 40, nextCol = native ? native.endColumn + 1 : left + 20;
  function jump() {
    const row = Number(jumpRow), col = columnIndex(jumpCol);
    if (Number.isInteger(row) && row >= 1 && row <= 100000 && /^[A-Za-z]{1,3}$/.test(jumpCol) && col <= 16384) { previousRows.current = []; setTop(row); setLeft(col); }
    else setError('起始行须为 1～100000，起始列须为 A～XFD');
  }
  return <dialog className="analysis-picker" ref={dialog} aria-label="选择分析数据区域" onClose={onClose}>
    <header><div><strong>{block.sheet} · 原表选区</strong><p>{mode === 'native' ? 'Excel 在后台渲染，直接在当前画面拖选；图表上的鼠标操作会选择其下方单元格。' : '基础模式显示数据值，不还原样式和图表。'} Shift + 点击可跨页扩展。</p></div><button className="icon-button" aria-label="关闭区域选择" onClick={() => dialog.current?.close()}><X size={20} /></button></header>
    <div className="analysis-picker-tools">
      <label>视图<select aria-label="选区渲染方式" value={mode} onChange={e => setMode(e.target.value as 'native' | 'basic')}><option value="native">Excel 原样视图</option><option value="basic">基础值视图</option></select></label>
      <label>起始行<input aria-label="选区预览起始行" type="number" min={1} max={100000} value={jumpRow} onChange={e => setJumpRow(e.target.value)} /></label>
      <label>起始列<input aria-label="选区预览起始列" value={jumpCol} onChange={e => setJumpCol(e.target.value)} /></label>
      <button className="button small" disabled={loading} onClick={jump}>跳转</button>
      <button className="button small" disabled={loading || top <= 1} aria-label="上一页数据行" onClick={() => { const row = previousRows.current.pop() || Math.max(1, top - 40); setTop(row); setJumpRow(String(row)); }}><ChevronLeft size={14} />上一页</button>
      <button className="button small" disabled={!ready || nextRow > Math.min(100000, extent?.totalRows || 100000)} aria-label="下一页数据行" onClick={() => { previousRows.current.push(top); setTop(nextRow); setJumpRow(String(nextRow)); }}>下一页<ChevronRight size={14} /></button>
      <button className="button small" disabled={loading || left <= 1} aria-label="向左查看数据列" onClick={() => { const col = Math.max(1, left - 20); setLeft(col); setJumpCol(columnLabel(col)); }}>左 20 列</button>
      <button className="button small" disabled={!ready || nextCol > Math.min(16384, extent?.totalColumns || 16384)} aria-label="向右查看数据列" onClick={() => { setLeft(nextCol); setJumpCol(columnLabel(nextCol)); }}>右 20 列</button>
      <button className="button small" disabled={loading} aria-label="重新渲染原表" onClick={() => setReload(n => n + 1)}><RefreshCw size={13} /></button>
      {mode === 'native' && <><label className="native-formula-toggle"><input type="checkbox" checked={showFormulas} disabled={loading} onChange={e => setShowFormulas(e.target.checked)} />显示公式原文</label><div className="preview-zoom"><button aria-label="缩小选区预览" disabled={zoom <= 50} onClick={() => setZoom(z => z - 10)}><Minus size={13} /></button><button aria-label="原始比例选区预览" onClick={() => setZoom(100)}>{zoom}%</button><button aria-label="放大选区预览" disabled={zoom >= 200} onClick={() => setZoom(z => z + 10)}><Plus size={13} /></button></div></>}
    </div>
    {mode === 'native' && <div className="native-formula-bar"><span>{active?.address || '单元格'}</span><b>fx</b><textarea rows={1} aria-label="单元格公式或内容" readOnly value={active?.formula || active?.text || ''} placeholder="点击单元格查看公式原文或已保存的值" />{active?.missingCache && <small>公式结果未缓存</small>}</div>}
    {error && <p className="analysis-error" role="alert">{error}</p>}
    <div className="analysis-picker-grid" ref={viewport} aria-busy={loading || (mode === 'native' && !!native && !imageReady)}>
      {loading && <div className="analysis-loading"><LoaderCircle className="spin" size={20} />{mode === 'native' ? 'Excel 正在后台还原当前区域，完成后即可选区…' : '读取该页数据…'}</div>}
      {native && !loading && <NativeSelectionGrid frame={native} zoom={zoom} bounds={bounds} ready={ready} pointer={pointer} keyboard={select} onReady={() => setImageReady(true)} onImageError={() => setError('原表图片未能加载，请重新渲染')} />}
      {grid && !loading && <table aria-label="分析区域选择网格"><thead><tr><th aria-label="行列坐标" />{grid.columns.map((label, c) => <th key={label} {...pointer({ row: 1, col: grid.startColumn + c, type: 'column' })}><button aria-label={`选择整列 ${label}`} onClick={e => { if (e.detail === 0) select({ row: 1, col: grid.startColumn + c, type: 'column' }, e.shiftKey); }}>{label}</button></th>)}</tr></thead><tbody>{grid.rows.map((row, i) => {
        const r = grid.startRow + i;
        return <tr key={r}><th {...pointer({ row: r, col: 1, type: 'row' })}><button aria-label={`选择整行 ${r}`} onClick={e => { if (e.detail === 0) select({ row: r, col: 1, type: 'row' }, e.shiftKey); }}>{r}</button></th>{row.map((value, j) => {
          const c = grid.startColumn + j, selected = bounds && c >= bounds[0] && c <= bounds[2] && r >= bounds[1] && r <= bounds[3];
          const hidden = grid.hiddenRows.includes(r) || grid.hiddenColumns.includes(c);
          return <td key={c} data-analysis-cell={`${columnLabel(c)}${r}`} className={`${selected ? 'selected' : ''} ${hidden ? 'hidden-cell' : ''}`} title={value} {...pointer({ row: r, col: c, type: 'cell' })}>{value || '\u00a0'}</td>;
        })}</tr>;
      })}</tbody></table>}
    </div>
    {native && <div className="native-render-notes"><span>Microsoft Excel {native.excelVersion} · 第 {native.startRow}–{native.endRow} 行 / {columnLabel(native.startColumn)}–{columnLabel(native.endColumn)} 列 · 原文件只读，公式不重算</span>{native.warnings.map(w => <span key={w}>{w}</span>)}</div>}
    <footer><label>所选区域<input aria-label="所选分析区域" value={range} onChange={e => { setRange(e.target.value); setBounds(rangeBounds(e.target.value)); }} /></label><p>行号、列标可选整行整列；合并格按完整范围选择。选区使用真实坐标，与缩放比例无关。</p><button className="button primary" disabled={!range.trim() || !ready} onClick={() => onApply(range.trim())}>使用此区域</button></footer>
  </dialog>;
}
