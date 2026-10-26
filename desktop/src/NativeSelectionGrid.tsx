// Copyright (c) 2026 AbeLin · MIT License
import { useMemo, type PointerEvent } from 'react';
import type { NativePreview } from '../shared/analysis';

export type GridPosition = { row: number; col: number; endRow?: number; endCol?: number; type: 'cell' | 'row' | 'column' };
export const columnLabel = (index: number): string => {
  let label = ''; for (let n = index; n > 0; n = Math.floor((n - 1) / 26)) label = String.fromCharCode(65 + (n - 1) % 26) + label;
  return label;
};

export default function NativeSelectionGrid({ frame, zoom, bounds, ready, pointer, keyboard, onReady, onImageError }: {
  frame: NativePreview; zoom: number; bounds: [number, number, number, number] | null; ready: boolean;
  pointer: (point: GridPosition) => { onPointerDown: (event: PointerEvent) => void; onPointerEnter: () => void };
  keyboard: (point: GridPosition, shift: boolean) => void; onReady: () => void; onImageError: () => void;
}) {
  const cells = useMemo(() => {
    const result: { point: GridPosition; x: number; y: number; width: number; height: number; address: string }[] = [];
    const seen = new Set<string>();
    for (const row of frame.rows.filter(r => r.size > 0)) for (const col of frame.columns.filter(c => c.size > 0)) {
      const merge = frame.merges.find(m => col.index >= m[0] && col.index <= m[2] && row.index >= m[1] && row.index <= m[3]);
      const [c1, r1, c2, r2] = merge || [col.index, row.index, col.index, row.index];
      const address = `${columnLabel(c1)}${r1}`;
      if (seen.has(address)) continue;
      seen.add(address);
      const cols = frame.columns.filter(c => c.index >= c1 && c.index <= c2 && c.size > 0);
      const rows = frame.rows.filter(r => r.index >= r1 && r.index <= r2 && r.size > 0);
      result.push({ point: { row: r1, col: c1, endRow: r2, endCol: c2, type: 'cell' }, address,
        x: cols[0].start, y: rows[0].start, width: cols.at(-1)!.start + cols.at(-1)!.size - cols[0].start,
        height: rows.at(-1)!.start + rows.at(-1)!.size - rows[0].start });
    }
    return result;
  }, [frame]);
  const key = (point: GridPosition) => ({ onKeyDown: (event: React.KeyboardEvent) => {
    if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); keyboard(point, event.shiftKey); }
  } });
  return <div className="native-sheet-surface" style={{ width: frame.width + 44, height: frame.height + 25, zoom: zoom / 100 }} data-native-sheet={frame.sheet}>
    <img className="native-sheet-image" src={frame.image} alt={`${frame.sheet} · Microsoft Excel 原样画面`} width={frame.width} height={frame.height} draggable={false} onLoad={onReady} onError={onImageError} />
    <div className="native-column-headers">
      {frame.columns.filter(c => c.size > 0).map(col => <button key={col.index} disabled={!ready} aria-label={`选择整列 ${columnLabel(col.index)}`} style={{ left: 44 + col.start, width: col.size }} {...pointer({ row: 1, col: col.index, type: 'column' })} {...key({ row: 1, col: col.index, type: 'column' })}>{columnLabel(col.index)}</button>)}
      <span className="native-corner" aria-label="行列坐标" />
    </div>
    <div className="native-row-headers" style={{ height: frame.height }}>{frame.rows.filter(r => r.size > 0).map(row => <button key={row.index} disabled={!ready} aria-label={`选择整行 ${row.index}`} style={{ top: row.start, height: row.size }} {...pointer({ row: row.index, col: 1, type: 'row' })} {...key({ row: row.index, col: 1, type: 'row' })}>{row.index}</button>)}</div>
    <div className="native-cell-layer" style={{ width: frame.width, height: frame.height }} aria-label="Excel 原样区域选择网格">
      {cells.map(cell => {
        const p = cell.point, selected = bounds && p.col <= bounds[2] && (p.endCol || p.col) >= bounds[0] && p.row <= bounds[3] && (p.endRow || p.row) >= bounds[1];
        const detail = frame.cells[cell.address];
        return <button key={cell.address} data-analysis-cell={cell.address} disabled={!ready} className={selected ? 'native-hit-cell selected' : 'native-hit-cell'} aria-label={`选择单元格 ${cell.address}`} title={`${cell.address}${detail ? '\n' + (detail.formula || detail.text) : ''}`} style={{ left: cell.x, top: cell.y, width: cell.width, height: cell.height }} {...pointer(p)} {...key(p)} />;
      })}
    </div>
  </div>;
}
