// Copyright (c) 2026 AbeLin · MIT License
import { useEffect, useState } from 'react';
import { Expand, LoaderCircle } from 'lucide-react';
import type { Preview } from '../shared/types';
import WorkbookView from './WorkbookView';
import SheetPreview from './SheetPreview';

export default function ExcelSheetPreview({ inputPath, path, sheet, headerRow }: { inputPath: string; path: string; sheet: string; headerRow: number | null }) {
  const [mode, setMode] = useState<'native' | 'portable'>('native'), [expanded, setExpanded] = useState(false);
  const [preview, setPreview] = useState<Preview | null>(null), [loading, setLoading] = useState(false), [error, setError] = useState('');
  useEffect(() => {
    if (mode !== 'portable') return;
    let current = true; setLoading(true); setError(''); setPreview(null);
    void window.desktop?.preview(path, sheet, headerRow).then(v => { if (current) setPreview(v); }).catch(e => { if (current) setError(String(e.message || e)); }).finally(() => { if (current) setLoading(false); });
    return () => { current = false; };
  }, [mode, path, sheet, headerRow]);
  const block = { id: 'preview', name: sheet, path, sheet, range: '', headerRows: headerRow || 1, orientation: 'columns' as const, includeHidden: false, skipTotals: false };
  return <div className="excel-sheet-preview">
    <div className="excel-preview-mode"><label>预览方式<select aria-label="工作表预览方式" value={mode} onChange={e => setMode(e.target.value as 'native' | 'portable')}><option value="native">Excel 原样视图（推荐）</option><option value="portable">基础样式预览（无需 Excel）</option></select></label>
      {mode === 'native' && <button className="button small" onClick={() => setExpanded(true)}><Expand size={14} />展开预览</button>}</div>
    {mode === 'native' ? <><p className="excel-preview-tip">本机需安装 Microsoft Excel。无法使用时可在上方切换基础样式预览；基础预览对复杂图表和样式的还原有限。</p><WorkbookView inputPath={inputPath} block={block} selection={false} embedded={!expanded} modeChoice={false} onClose={() => setExpanded(false)} /></>
      : loading ? <div className="preview-placeholder"><LoaderCircle className="spin" size={20} />正在读取基础样式…</div> : error ? <div className="analysis-error" role="alert">{error}</div> : preview && <SheetPreview preview={preview} name={sheet} />}
  </div>;
}
