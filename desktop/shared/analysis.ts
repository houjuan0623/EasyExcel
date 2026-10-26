// Copyright (c) 2026 AbeLin · MIT License
export type ChartKind = 'metric' | 'bar' | 'line' | 'pie' | 'histogram' | 'scatter' | 'box' | 'correlation';
export interface AnalysisBlock {
  id: string; name: string; path: string; sheet: string; range: string; headerRows: number;
  orientation: 'columns' | 'rows'; includeHidden: boolean; skipTotals: boolean;
}
export interface AnalysisField { id: string; name: string; location: string; kind: 'number' | 'date' | 'mixed' | 'text' }
export interface AnalysisSchema {
  token: string; blocks: { id: string; name: string; fields: AnalysisField[]; rows: number; resolvedRange: string;
    sample: (string | number | boolean | null)[][]; audit: Record<string, number>; warnings: string[] }[];
}
export interface AnalysisRequest {
  inputPath: string; blocks: AnalysisBlock[]; token: string; kind: ChartKind; title: string; unit: string; note: string;
  combine: 'single' | 'append' | 'join'; joinKeys: Record<string, string>; joinType: 'inner' | 'left';
  x: string; y: string[]; aggregation: 'sum' | 'mean' | 'median' | 'min' | 'max' | 'count';
  sort: 'source' | 'label' | 'value_desc'; timeBucket: 'none' | 'day' | 'month' | 'year'; bins: number; limit: number;
  fieldLabels: Record<string, string>;
}
export interface AnalysisTable { columns: string[]; rows: (string | number | boolean | null)[][] }
export interface AnalysisResult {
  title: string; kind: ChartKind; image: string | null; metric: number | null; unit: string; note: string; method: string;
  statistics: AnalysisTable; table: AnalysisTable; counts: Record<string, number>; warnings: string[];
  sources: { name: string; file: string; sheet: string; range: string; rows: number; headerRows: number; orientation: string }[];
  generatedAt: string; token: string;
}
export interface AnalysisGrid {
  startRow: number; startColumn: number; totalRows: number; totalColumns: number;
  columns: string[]; rows: string[][]; hiddenRows: number[]; hiddenColumns: number[];
  suggestedRange: string; suggestedHeaderRows: number;
}
export interface NativePreview {
  engine: 'excel'; sheet: string; image: string; width: number; height: number; excelVersion: string;
  startRow: number; startColumn: number; endRow: number; endColumn: number; totalRows: number; totalColumns: number;
  rows: { index: number; start: number; size: number }[];
  columns: { index: number; start: number; size: number }[];
  cells: Record<string, { formula: string; text: string; missingCache: boolean }>;
  merges: [number, number, number, number][]; warnings: string[];
}
export interface AnalysisCard { id: string; config: AnalysisRequest }
export interface AnalysisPlan { version: 1; title: string; cards: AnalysisCard[] }

// 方案只含配置，不接受代码或结果图片；导入后必须由 Python 重新核对数据和字段。
export function parseAnalysisPlan(value: unknown): AnalysisPlan {
  const obj = (v: unknown): Record<string, any> => { if (!v || typeof v !== 'object' || Array.isArray(v)) throw new Error('分析方案结构无效'); return v as Record<string, any>; };
  const str = (v: unknown, max = 200): string => { if (typeof v !== 'string' || v.length > max) throw new Error('分析方案文字格式无效'); return v; };
  const choice = <T extends string>(v: unknown, options: readonly T[]): T => { if (!options.includes(v as T)) throw new Error('分析方案含不支持的设置'); return v as T; };
  const num = (v: unknown, low: number, high: number): number => { if (!Number.isInteger(v) || (v as number) < low || (v as number) > high) throw new Error('分析方案数值超出范围'); return v as number; };
  const identifier = (v: unknown): string => { const s = str(v, 80); if (!/^[a-zA-Z0-9_-]+$/.test(s) || ['__proto__', 'constructor', 'prototype'].includes(s)) throw new Error('分析方案标识无效'); return s; };
  const map = (v: unknown): Record<string, string> => Object.fromEntries(Object.entries(obj(v)).map(([k, x]) => [str(k), str(x, 1000)]));
  const root = obj(value);
  if (root.version !== 1 || !Array.isArray(root.cards) || root.cards.length > 12) throw new Error('分析方案版本或图表数量无效');
  const cards = root.cards.map((raw: unknown): AnalysisCard => {
    const card = obj(raw), c = obj(card.config);
    if (!Array.isArray(c.blocks) || c.blocks.length < 1 || c.blocks.length > 4 || !Array.isArray(c.y) || c.y.length > 6) throw new Error('分析方案数据区域或数值字段数量无效');
    const blocks = c.blocks.map((b: unknown): AnalysisBlock => {
      const x = obj(b);
      const id = identifier(x.id);
      return { id, name: str(x.name, 80), path: str(x.path, 3000), sheet: str(x.sheet, 80), range: str(x.range, 80),
        headerRows: num(x.headerRows, 0, 10), orientation: choice(x.orientation, ['columns', 'rows']), includeHidden: x.includeHidden === true, skipTotals: x.skipTotals !== false };
    });
    if (new Set(blocks.map((b: AnalysisBlock) => b.id)).size !== blocks.length) throw new Error('数据块标识重复');
    return { id: identifier(card.id), config: { inputPath: str(c.inputPath, 3000), blocks, token: str(c.token, 100),
      kind: choice(c.kind, ['metric', 'bar', 'line', 'pie', 'histogram', 'scatter', 'box', 'correlation']),
      title: str(c.title, 120), unit: str(c.unit, 30), note: str(c.note, 1000),
      combine: choice(c.combine, ['single', 'append', 'join']), joinKeys: map(c.joinKeys), joinType: choice(c.joinType, ['inner', 'left']),
      x: str(c.x), y: c.y.map((v: unknown) => str(v)), aggregation: choice(c.aggregation, ['sum', 'mean', 'median', 'min', 'max', 'count']),
      sort: choice(c.sort, ['source', 'label', 'value_desc']), timeBucket: choice(c.timeBucket, ['none', 'day', 'month', 'year']),
      bins: num(c.bins, 3, 50), limit: num(c.limit, 1, 60), fieldLabels: map(c.fieldLabels) } };
  });
  if (new Set(cards.map((c: AnalysisCard) => c.id)).size !== cards.length) throw new Error('图表标识重复');
  return { version: 1, title: str(root.title, 120), cards };
}
