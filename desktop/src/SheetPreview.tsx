// Copyright (c) 2026 AbeLin · MIT License
import { useEffect, useMemo, useRef, useState, type CSSProperties } from 'react';
import { Expand, Minus, Plus, Table2, X } from 'lucide-react';
import type { Preview } from '../shared/types';
import './preview.css';

function Grid({ preview, zoom, name }: { preview: Preview; zoom: number; name: string }) {
  const widths = useMemo(() => new Map(preview.columns.map((col) => [col.index, col.width])), [preview]);
  const measure = useMemo(() => document.createElement('canvas').getContext('2d'), []);
  const width = 44 + preview.columns.reduce((sum, col) => sum + col.width, 0);
  return <div className="sheet-viewport" tabIndex={0} role="region" aria-label={`${name} 原表预览，可滚动查看`}>
    {!preview.rows.length || !preview.columns.length ? <div className="preview-placeholder">前 100 行没有可见单元格</div> :
      <table className={`sheet-grid ${preview.showGridLines ? '' : 'without-gridlines'}`} aria-label={`${name} 前 100 行`} style={{ width, zoom: zoom / 100 }}>
        <colgroup><col style={{ width: 44 }} />{preview.columns.map((col) => <col key={col.index} style={{ width: col.width }} />)}</colgroup>
        <thead><tr><th className="sheet-corner" aria-label="行号与列标" />{preview.columns.map((col) => <th key={col.index} scope="col">{col.label}</th>)}</tr></thead>
        <tbody>{preview.rows.map((row) => <tr key={row.index} style={{ height: row.height }}>
          <th scope="row" className="sheet-row-number">{row.index}</th>
          {row.cells.map((cell) => {
            const { wrapText, shrinkToFit, indent = 0, rotation = 0, verticalAlign = 'bottom', ...style } = preview.styles[cell.style];
            const colIndex = preview.columns.findIndex((col) => col.index === cell.column);
            const cellWidth = cell.colSpan === 1 ? widths.get(cell.column)! : preview.columns.slice(colIndex, colIndex + cell.colSpan).reduce((sum, col) => sum + col.width, 0);
            let fontSize = style.fontSize || 11 * 4 / 3;
            if (shrinkToFit && !wrapText && cell.text && measure) {
              measure.font = `${style.fontStyle || 'normal'} ${style.fontWeight || 'normal'} ${fontSize}px ${style.fontFamily || 'Calibri'}`;
              fontSize *= Math.min(1, Math.max(1, cellWidth - indent - 6) / Math.max(1, measure.measureText(cell.text).width));
            }
            const textStyle: CSSProperties = { whiteSpace: wrapText ? 'pre-wrap' : 'pre', overflowWrap: wrapText ? 'anywhere' : 'normal', fontSize };
            if (rotation === 255) { textStyle.writingMode = 'vertical-rl'; textStyle.textOrientation = 'upright'; }
            else if (rotation) { textStyle.transform = `rotate(${rotation <= 90 ? -rotation : rotation - 90}deg)`; textStyle.transformOrigin = 'center'; }
            return <td key={cell.column} className="sheet-cell" data-cell={`${preview.columns[colIndex].label}${row.index}`} rowSpan={cell.rowSpan} colSpan={cell.colSpan} style={style} title={cell.text}>
              <div className="sheet-cell-box" style={{ alignItems: verticalAlign === 'top' ? 'flex-start' : verticalAlign === 'middle' ? 'center' : 'flex-end', paddingLeft: 3 + indent }}>
                <span style={textStyle}>{cell.text}</span>
              </div>
            </td>;
          })}
        </tr>)}</tbody>
      </table>}
  </div>;
}

export default function SheetPreview({ preview, name }: { preview: Preview; name: string }) {
  const [zoom, setZoom] = useState(100);
  const [expanded, setExpanded] = useState(false);
  const dialog = useRef<HTMLDialogElement>(null);
  useEffect(() => { if (expanded) dialog.current?.showModal(); }, [expanded]);
  const status = `原表第 1–${preview.endRow} 行${preview.totalRows > preview.endRow ? ' · 仅显示前 100 行' : ''}`;
  const zoomControls = <div className="preview-zoom">
    <button type="button" title="缩小预览" aria-label="缩小预览" disabled={zoom <= 50} onClick={() => setZoom((n) => Math.max(50, n - 10))}><Minus size={14} /></button>
    <button type="button" title="恢复原始比例" aria-label="恢复原始比例" onClick={() => setZoom(100)}>{zoom}%</button>
    <button type="button" title="放大预览" aria-label="放大预览" disabled={zoom >= 200} onClick={() => setZoom((n) => Math.min(200, n + 10))}><Plus size={14} /></button>
  </div>;
  const notes = <div className="preview-notes">
    <span>{preview.endRow ? status : '空白工作表'}{preview.hiddenRows + preview.hiddenColumns > 0 && ` · 已隐藏 ${preview.hiddenRows} 行 / ${preview.hiddenColumns} 列`}</span>
    {preview.warnings.map((warning) => <span className="preview-warning" key={warning}>{warning}</span>)}
  </div>;
  return <>
    <div className="preview-tabbar"><span className="preview-tab" title={name}><Table2 size={14} />{name}</span><div className="preview-tools">{zoomControls}<button className="preview-expand" onClick={() => setExpanded(true)}><Expand size={14} />展开预览</button></div></div>
    <Grid preview={preview} zoom={zoom} name={name} />
    {notes}
    <dialog className="preview-dialog" ref={dialog} onClose={() => setExpanded(false)} aria-label={`${name} 展开预览`}>
      {expanded && <><div className="preview-dialog-heading"><div><Table2 size={19} /><strong>{name}</strong><span>前 100 行 · 原表视图</span></div><div className="preview-tools">{zoomControls}<button className="icon-button" aria-label="关闭展开预览" onClick={() => dialog.current?.close()}><X size={20} /></button></div></div><Grid preview={preview} zoom={zoom} name={name} />{notes}</>}
    </dialog>
  </>;
}
