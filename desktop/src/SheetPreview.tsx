// Copyright (c) 2026 AbeLin · MIT License
import { useEffect, useMemo, useRef, useState, type CSSProperties } from 'react';
import { Expand, Minus, Plus, Table2, X } from 'lucide-react';
import type { Preview, PreviewDrawing } from '../shared/types';
import './preview.css';

function Grid({ preview, zoom, name, showDrawings, openDrawing }: {
  preview: Preview; zoom: number; name: string; showDrawings: boolean; openDrawing: (drawing: PreviewDrawing) => void;
}) {
  const widths = useMemo(() => new Map(preview.columns.map((col) => [col.index, col.width])), [preview]);
  const measure = useMemo(() => document.createElement('canvas').getContext('2d'), []);
  const width = 44 + preview.columns.reduce((sum, col) => sum + col.width, 0);
  const drawings = showDrawings ? preview.drawings.filter((drawing) => drawing.inline && drawing.src) : [];
  const surfaceWidth = Math.max(width, ...drawings.map((drawing) => 44 + drawing.x! + drawing.width!));
  const height = Math.max(25 + preview.rows.reduce((sum, row) => sum + row.height, 0), ...drawings.map((drawing) => 25 + drawing.y! + drawing.height!));
  return <div className="sheet-viewport" tabIndex={0} role="region" aria-label={`${name} 原表预览，可滚动查看`}>
    <div className="sheet-surface" style={{ width: surfaceWidth, minHeight: height, zoom: zoom / 100 }}>
    {!preview.rows.length || !preview.columns.length ? <div className="preview-placeholder">前 100 行没有可见单元格</div> :
      <table className={`sheet-grid ${preview.showGridLines ? '' : 'without-gridlines'}`} aria-label={`${name} 前 100 行`} style={{ width }}>
        <colgroup><col style={{ width: 44 }} />{preview.columns.map((col) => <col key={col.index} style={{ width: col.width }} />)}</colgroup>
        <thead><tr><th className="sheet-corner" aria-label="行号与列标" />{preview.columns.map((col) => <th key={col.index} scope="col">{col.label}</th>)}</tr></thead>
        <tbody>{preview.rows.map((row) => <tr key={row.index} style={{ height: row.height }}>
          <th scope="row" className="sheet-row-number">{row.index}</th>
          {row.cells.map((cell) => {
            const { wrapText, shrinkToFit, indent = 0, rotation = 0, verticalAlign = 'bottom', diagonalUp, diagonalDown, diagonalColor, script, ...style } = preview.styles[cell.style];
            const colIndex = preview.columns.findIndex((col) => col.index === cell.column);
            const cellWidth = cell.colSpan === 1 ? widths.get(cell.column)! : preview.columns.slice(colIndex, colIndex + cell.colSpan).reduce((sum, col) => sum + col.width, 0);
            let fontSize = style.fontSize || 11 * 4 / 3;
            if (shrinkToFit && !wrapText && cell.text && measure) {
              measure.font = `${style.fontStyle || 'normal'} ${style.fontWeight || 'normal'} ${fontSize}px ${style.fontFamily || 'Calibri'}`;
              fontSize *= Math.min(1, Math.max(1, cellWidth - indent - 6) / Math.max(1, measure.measureText(cell.text).width));
            }
            const textStyle: CSSProperties = { whiteSpace: wrapText ? 'pre-wrap' : 'pre', overflowWrap: wrapText ? 'anywhere' : 'normal', fontSize };
            if (script === 'superscript' || script === 'subscript') { textStyle.fontSize = fontSize * .75; textStyle.transform = `translateY(${script === 'superscript' ? '-25%' : '25%'})`; }
            if (rotation === 255) { textStyle.writingMode = 'vertical-rl'; textStyle.textOrientation = 'upright'; }
            else if (rotation) { textStyle.transform = `rotate(${rotation <= 90 ? -rotation : rotation - 90}deg)`; textStyle.transformOrigin = 'center'; }
            return <td key={cell.column} className="sheet-cell" data-cell={`${preview.columns[colIndex].label}${row.index}`} rowSpan={cell.rowSpan} colSpan={cell.colSpan} style={style} title={cell.text}>
              {cell.dataBar && <div className="sheet-data-bar" aria-hidden="true" style={{ left: `${cell.dataBar.start}%`, width: `${cell.dataBar.width}%`, backgroundColor: cell.dataBar.color }} />}
              {(diagonalUp || diagonalDown) && <svg className="sheet-diagonal" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">
                {diagonalUp && <path d="M0 100 L100 0" stroke={diagonalColor} vectorEffect="non-scaling-stroke" />}
                {diagonalDown && <path d="M0 0 L100 100" stroke={diagonalColor} vectorEffect="non-scaling-stroke" />}
              </svg>}
              <div className="sheet-cell-box" style={{ alignItems: verticalAlign === 'top' ? 'flex-start' : verticalAlign === 'middle' ? 'center' : 'flex-end', paddingLeft: 3 + indent }}>
                <span style={textStyle}>{cell.icon && <span className="sheet-condition-icon" style={{ color: cell.icon.color }} aria-hidden="true">{cell.icon.text}</span>}{!cell.hideText && (cell.runs ? cell.runs.map((run, index) => <span key={index} style={run.style}>{run.text}</span>) : cell.text)}</span>
              </div>
            </td>;
          })}
        </tr>)}</tbody>
      </table>}
      {drawings.map((drawing) => <button key={drawing.id} className="sheet-drawing" aria-label={`放大查看 ${drawing.title}`} title={`${drawing.title} · 点击放大`} style={{ left: 44 + drawing.x!, top: 25 + drawing.y!, width: drawing.width, height: drawing.height }} onClick={() => openDrawing(drawing)}>
        <img src={drawing.src} alt={drawing.title} draggable={false} />
      </button>)}
    </div>
  </div>;
}

export default function SheetPreview({ preview, name }: { preview: Preview; name: string }) {
  const [zoom, setZoom] = useState(100);
  const [expanded, setExpanded] = useState(false);
  const [showDrawings, setShowDrawings] = useState(true);
  const [selectedDrawing, setSelectedDrawing] = useState<PreviewDrawing | null>(null);
  const dialog = useRef<HTMLDialogElement>(null);
  const drawingDialog = useRef<HTMLDialogElement>(null);
  useEffect(() => { if (expanded) dialog.current?.showModal(); }, [expanded]);
  useEffect(() => { if (selectedDrawing) drawingDialog.current?.showModal(); }, [selectedDrawing]);
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
  const drawingToggle = preview.drawings.some((drawing) => drawing.inline && drawing.src) && <label className="preview-drawing-toggle"><input type="checkbox" checked={showDrawings} onChange={(event) => setShowDrawings(event.target.checked)} />显示图表和图片</label>;
  const grid = <Grid preview={preview} zoom={zoom} name={name} showDrawings={showDrawings} openDrawing={setSelectedDrawing} />;
  return <>
    <div className="preview-tabbar"><span className="preview-tab" title={name}><Table2 size={14} />{name}</span><div className="preview-tools">{drawingToggle}{zoomControls}<button className="preview-expand" onClick={() => setExpanded(true)}><Expand size={14} />展开预览</button></div></div>
    {grid}
    {preview.drawings.length > 0 && <details className="preview-drawings" open>
      <summary>统计图与图片 <span>{preview.drawings.length} 个 · 点击放大，也包含第 100 行之后的图表</span></summary>
      <div className="preview-drawing-list">{preview.drawings.map((drawing) => <div className="preview-drawing-card" key={drawing.id}>
        {drawing.src ? <button onClick={() => setSelectedDrawing(drawing)} aria-label={`查看图表图片 ${drawing.title}`}><img src={drawing.src} alt={drawing.title} loading="lazy" /><span>{drawing.title}</span></button> : <div className="preview-drawing-error"><strong>{drawing.title}</strong><span>{drawing.error || '此对象暂不支持预览'}</span></div>}
      </div>)}</div>
    </details>}
    {notes}
    <dialog className="preview-dialog" ref={dialog} onClose={() => setExpanded(false)} aria-label={`${name} 展开预览`}>
      {expanded && <><div className="preview-dialog-heading"><div><Table2 size={19} /><strong>{name}</strong><span>前 100 行 · 原表视图</span></div><div className="preview-tools">{drawingToggle}{zoomControls}<button className="icon-button" aria-label="关闭展开预览" onClick={() => dialog.current?.close()}><X size={20} /></button></div></div>{grid}{notes}</>}
    </dialog>
    <dialog className="preview-image-dialog" ref={drawingDialog} onClose={() => setSelectedDrawing(null)} aria-label="图表图片放大查看">
      {selectedDrawing && <><div className="preview-dialog-heading"><strong>{selectedDrawing.title}</strong><button className="icon-button" aria-label="关闭图表图片" onClick={() => drawingDialog.current?.close()}><X size={20} /></button></div><div className="preview-image-body"><img src={selectedDrawing.src} alt={selectedDrawing.title} /></div></>}
    </dialog>
  </>;
}
