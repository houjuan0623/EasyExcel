// Copyright (c) 2026 AbeLin · MIT License
import type { AnalysisBlock } from '../shared/analysis';
import WorkbookView from './WorkbookView';
export { columnLabel } from './NativeSelectionGrid';
export default function AnalysisRangePicker(props: { inputPath: string; block: AnalysisBlock; onClose: () => void; onApply: (range: string) => void }) {
  return <WorkbookView {...props} />;
}
