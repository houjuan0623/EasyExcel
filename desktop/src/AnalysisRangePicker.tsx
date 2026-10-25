// Copyright (c) 2026 AbeLin · MIT License
import { useEffect, useRef, useState } from 'react';
import { ChevronLeft, ChevronRight, LoaderCircle, X } from 'lucide-react';
import type { AnalysisBlock, AnalysisGrid } from '../shared/analysis';

export function columnLabel(index: number): string {
  let label = ''; for (let n = index; n > 0; n = Math.floor((n - 1) / 26)) label = String.fromCharCode(65 + (n - 1) % 26) + label;
  return label;
}
const columnIndex = (label: string) => [...label.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const errorText = (e: unknown) => e instanceof Error ? e.message.replace(/^Error invoking remote method '[^']+': (?:Error: )?/, '') : String(e);
type Position = { row: number; col: number; type: 'cell' | 'row' | 'column' };

export default function AnalysisRangePicker({ inputPath, block, onClose, onApply }: {
  inputPath: string; block: AnalysisBlock; onClose: () => void; onApply: (range: string) => void;
}) {
  const dialog = useRef<HTMLDialogElement>(null);
  const [grid, setGrid] = useState<AnalysisGrid | null>(null);
  const [top, setTop] = useState(1), [left, setLeft] = useState(1);
  const [jumpRow, setJumpRow] = useState('1'), [jumpCol, setJumpCol] = useState('A');
  const [range, setRange] = useState(block.range);
  const [error, setError] = useState(''), [loading, setLoading] = useState(false);
  const anchor = useRef<Position | null>(null), dragging = useRef(false);
  const [bounds, setBounds] = useState<[number, number, number, number] | null>(null);
  useEffect(() => { dialog.current?.showModal(); const up = () => { dragging.current = false; }; window.addEventListener('pointerup', up); return () => window.removeEventListener('pointerup', up); }, []);
  useEffect(() => {
    let current = true; setLoading(true); setError('');
    void window.desktop?.analysisGrid({ inputPath, path: block.path, sheet: block.sheet, startRow: top, startColumn: left, includeHidden: block.includeHidden })
      .then((value) => { if (current) setGrid(value); }).catch((e) => { if (current) setError(errorText(e)); }).finally(() => { if (current) setLoading(false); });
    return () => { current = false; };
  }, [inputPath, block.path, block.sheet, block.includeHidden, top, left]);
  function extend(position: Position) {
    if (!anchor.current || anchor.current.type !== position.type) return;
    const r1 = Math.min(anchor.current.row, position.row), r2 = Math.max(anchor.current.row, position.row);
    const c1 = Math.min(anchor.current.col, position.col), c2 = Math.max(anchor.current.col, position.col);
    if (position.type === 'row') { setRange(`${r1}:${r2}`); setBounds([1, r1, 16384, r2]); }
    else if (position.type === 'column') { setRange(`${columnLabel(c1)}:${columnLabel(c2)}`); setBounds([c1, 1, c2, 1048576]); }
    else { setRange(`${columnLabel(c1)}${r1}:${columnLabel(c2)}${r2}`); setBounds([c1, r1, c2, r2]); }
  }
  const pointer = (position: Position) => ({
    onPointerDown: (event: React.PointerEvent) => { if (event.button !== 0 || loading) return; event.preventDefault(); if (!event.shiftKey || anchor.current?.type !== position.type) anchor.current = position; dragging.current = true; extend(position); },
    onPointerEnter: () => { if (dragging.current) extend(position); },
  });
  return <dialog className="analysis-picker" ref={dialog} aria-label="选择分析数据区域" onClose={onClose}>
    <header><div><strong>{block.sheet} · 选择数据</strong><p>拖选矩形；点击或拖选行号、列标可选整行整列；Shift + 点击扩展到另一页。</p></div><button className="icon-button" aria-label="关闭区域选择" onClick={() => dialog.current?.close()}><X size={20} /></button></header>
    <div className="analysis-picker-tools">
      <label>起始行<input aria-label="选区预览起始行" type="number" min={1} max={100000} value={jumpRow} onChange={(e) => setJumpRow(e.target.value)} /></label>
      <label>起始列<input aria-label="选区预览起始列" value={jumpCol} onChange={(e) => setJumpCol(e.target.value)} /></label>
      <button className="button small" disabled={loading} onClick={() => { const row = Number(jumpRow); const col = columnIndex(jumpCol); if (Number.isInteger(row) && row >= 1 && row <= 100000 && /^[A-Za-z]{1,3}$/.test(jumpCol) && col <= 16384) { setTop(row); setLeft(col); } else setError('起始行须为 1～100000，起始列须为 A～XFD'); }}>跳转</button>
      <button className="button small" disabled={loading || top <= 1} aria-label="上一页数据行" onClick={() => { setTop(Math.max(1, top - 40)); setJumpRow(String(Math.max(1, top - 40))); }}><ChevronLeft size={14} />上一页</button>
      <button className="button small" disabled={loading || top + 40 > Math.min(100000, grid?.totalRows || 100000)} aria-label="下一页数据行" onClick={() => { setTop(top + 40); setJumpRow(String(top + 40)); }}>下一页<ChevronRight size={14} /></button>
      <span>{grid && `实际范围至 ${columnLabel(grid.totalColumns)}${grid.totalRows} · 每页 40 行 × 20 列`}</span>
    </div>
    {error && <p className="analysis-error" role="alert">{error}</p>}
    <div className="analysis-picker-grid" aria-busy={loading}>
      {loading && <div className="analysis-loading"><LoaderCircle className="spin" size={20} />读取该页数据…</div>}
      {grid && !loading && <table aria-label="分析区域选择网格"><thead><tr><th aria-label="行列坐标" />{grid.columns.map((label, c) => <th key={label} {...pointer({ row: 1, col: grid.startColumn + c, type: 'column' })}><button tabIndex={0} aria-label={`选择整列 ${label}`} onClick={(e) => { if (e.detail === 0) { anchor.current = { row: 1, col: grid.startColumn + c, type: 'column' }; extend(anchor.current); } }}>{label}</button></th>)}</tr></thead><tbody>{grid.rows.map((row, i) => {
        const r = grid.startRow + i;
        return <tr key={r}><th {...pointer({ row: r, col: 1, type: 'row' })}><button aria-label={`选择整行 ${r}`} onClick={(e) => { if (e.detail === 0) { anchor.current = { row: r, col: 1, type: 'row' }; extend(anchor.current); } }}>{r}</button></th>{row.map((value, j) => {
          const c = grid.startColumn + j, selected = bounds && c >= bounds[0] && c <= bounds[2] && r >= bounds[1] && r <= bounds[3];
          const hidden = grid.hiddenRows.includes(r) || grid.hiddenColumns.includes(c);
          return <td key={c} data-analysis-cell={`${columnLabel(c)}${r}`} className={`${selected ? 'selected' : ''} ${hidden ? 'hidden-cell' : ''}`} title={value} {...pointer({ row: r, col: c, type: 'cell' })}>{value || '\u00a0'}</td>;
        })}</tr>;
      })}</tbody></table>}
    </div>
    <footer><label>所选区域<input aria-label="所选分析区域" value={range} onChange={(e) => { setRange(e.target.value); setBounds(null); }} /></label><p>整行/列读取完整实际范围；标题、合计行可通过精确区域排除。浅灰行列为隐藏内容。</p><button className="button primary" disabled={!range.trim() || loading} onClick={() => onApply(range.trim())}>使用此区域</button></footer>
  </dialog>;
}
