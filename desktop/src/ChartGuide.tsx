// Copyright (c) 2026 AbeLin · MIT License
import type { ChartKind } from '../shared/analysis';
import guides from '../shared/chart-guides.json';
export { guides };
export default function ChartGuide({ kind, expanded = false }: { kind: ChartKind; expanded?: boolean }) {
  const guide = guides[kind];
  return <details className="chart-guide" open={expanded || undefined}>
    <summary>{guide.name}使用说明 · {guide.question}</summary>
    <div className="chart-guide-grid"><section><h4>什么情况下使用</h4><p>{guide.useWhen}</p><p className="chart-guide-caution">{guide.avoidWhen}</p><h4>例如</h4><p>{guide.example}</p></section>
      <section><h4>怎样选择数据并生成</h4><ol>{guide.steps.map(step => <li key={step}>{step}</li>)}</ol></section>
      <section><h4>怎么计算、怎么看图</h4><p>{guide.method}</p><p>{guide.reading}</p></section>
      <section><h4>遇到问题怎样调整</h4><ul>{guide.fixes.map(fix => <li key={fix}>{fix}</li>)}</ul></section></div>
  </details>;
}
