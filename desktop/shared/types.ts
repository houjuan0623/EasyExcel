// Copyright (c) 2026 AbeLin · MIT License
export type Mode = 'split' | 'merge';
export interface SheetInfo {
  name: string;
  state: string;
  headerRow: number;
  columns: string[];
  estimatedRows: number | null;
  selectable: boolean;
  reason: string;
}
export interface WorkbookInfo {
  path: string;
  name: string;
  size: number;
  sheets: SheetInfo[];
  error?: string;
}
export interface SourceInfo {
  path: string;
  name: string;
  isDirectory: boolean;
  workbooks: WorkbookInfo[];
  suggestedOutput: string;
}
export interface SheetChoice { path: string; sheets: string[] }
export interface SelectionRequest {
  inputPath: string;
  selection: SheetChoice[];
  headerRow: number | null;
}
export interface PreviewStyle {
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: 'bold' | 'normal';
  fontStyle?: 'italic' | 'normal';
  color?: string;
  backgroundColor?: string;
  backgroundImage?: string;
  backgroundSize?: string;
  textAlign?: 'left' | 'center' | 'right' | 'justify';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  textDecoration?: string;
  borderLeft?: string;
  borderRight?: string;
  borderTop?: string;
  borderBottom?: string;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  rotation?: number;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
  diagonalColor?: string;
  script?: 'superscript' | 'subscript' | 'baseline';
}
export interface PreviewCell {
  column: number; text: string; style: number; rowSpan: number; colSpan: number;
  runs?: { text: string; style: Pick<PreviewStyle, 'fontFamily' | 'fontSize' | 'fontWeight' | 'fontStyle' | 'color' | 'textDecoration'> & { verticalAlign?: 'super' | 'sub' | 'baseline' } }[];
  dataBar?: { start: number; width: number; color: string };
  icon?: { text: string; color: string };
  hideText?: boolean;
}
export interface PreviewDrawing {
  id: string; kind: 'chart' | 'image'; title: string; src?: string; error?: string;
  inline: boolean; x?: number; y?: number; width?: number; height?: number;
}
export interface Preview {
  columns: { index: number; label: string; width: number }[];
  rows: { index: number; height: number; cells: PreviewCell[] }[];
  styles: PreviewStyle[];
  rowLimit: number;
  endRow: number;
  totalRows: number;
  hiddenRows: number;
  hiddenColumns: number;
  truncatedColumns: boolean;
  showGridLines: boolean;
  warnings: string[];
  drawings: PreviewDrawing[];
}
export interface GroupPreview { values: string[]; missing: string[] }
export interface RunRequest extends SelectionRequest {
  mode: Mode;
  outputPath: string;
  column: string;
  values: string[] | null;
  preserveFormat: boolean;
  keepFormulas: boolean;
  makeZip: boolean;
  mergeAcrossFiles: boolean;
  bundleSheets: boolean;
  mergeMode: 'rows' | 'sheets';
}
export interface JobResult {
  output: string;
  summary: Record<string, unknown>;
  stopped: boolean;
}
export interface JobEvent {
  jobId: string;
  type: 'log' | 'progress' | 'result' | 'error';
  message?: string;
  value?: number;
  result?: JobResult;
}
export interface HistoryItem extends JobResult {
  id: string;
  name: string;
  mode: Mode | 'code';
  date: string;
}
export interface LibraryCapability { name: string; version: string; importName: string; purpose: string; example: string }
export interface CodeEnvironment { python: string; libraries: Record<string, string>; capabilities: LibraryCapability[] }
export interface CodeRequest extends SelectionRequest { code: string; outputPath: string; timeoutSeconds: number; scriptName?: string }
export type CodeOutput = { kind: 'image'; title: string; src: string } | {
  kind: 'table'; title: string; columns: string[]; rows: string[][]; index: string[]; totalRows: number; totalColumns: number;
};
export interface CodeResult {
  output: string; status: 'success' | 'error' | 'stopped' | 'timeout'; error?: string | null;
  artifacts: string[]; artifactCount: number; rows: number; duration: number;
}
export interface CodeEvent {
  jobId: string; type: 'log' | 'output' | 'result'; message?: string; data?: CodeOutput; result?: CodeResult;
}
export type ScriptStatus = CodeResult['status'] | 'never' | 'running';
export interface ScriptInfo {
  id: string; name: string; category: string; notes: string; favorite: boolean; archived: boolean;
  createdAt: string; updatedAt: string; lastRunAt: string | null; runCount: number; status: ScriptStatus; lines: number;
}
export interface ScriptRecord extends ScriptInfo { code: string }
export interface ScriptSave { code: string; name: string; category?: string; notes?: string; favorite?: boolean; rememberOnly?: boolean }
export interface ScriptLibrarySnapshot { items: ScriptInfo[]; warnings: string[]; path: string }
export interface DesktopAPI {
  health(): Promise<{ python: string; version: string; frozen: boolean }>;
  chooseSource(kind: 'file' | 'directory'): Promise<string | null>;
  chooseOutput(): Promise<string | null>;
  inspect(path: string, headerRow: number | null): Promise<SourceInfo>;
  preview(path: string, sheet: string, headerRow: number | null): Promise<Preview>;
  groups(request: SelectionRequest & { column: string }): Promise<GroupPreview>;
  start(request: RunRequest): Promise<string>;
  cancel(jobId: string): Promise<void>;
  onJob(callback: (event: JobEvent) => void): () => void;
  openOutput(path: string): Promise<void>;
  history(): Promise<HistoryItem[]>;
  codeEnvironment(): Promise<CodeEnvironment>;
  startCode(request: CodeRequest): Promise<string>;
  cancelCode(jobId: string): Promise<void>;
  onCode(callback: (event: CodeEvent) => void): () => void;
  importScript(): Promise<{ name: string; code: string } | null>;
  saveText(kind: 'script' | 'instructions', text: string): Promise<string | null>;
  copyText(text: string): Promise<void>;
  scripts(): Promise<ScriptLibrarySnapshot>;
  script(id: string): Promise<ScriptRecord>;
  saveScript(request: ScriptSave): Promise<ScriptRecord>;
  updateScript(id: string, changes: Partial<Pick<ScriptInfo, 'name' | 'category' | 'notes' | 'favorite' | 'archived'>>): Promise<ScriptRecord>;
  exportScripts(): Promise<string | null>;
  importScripts(): Promise<{ added: number; merged: number } | null>;
  pathForFile(file: File): string;
  window(action: 'minimize' | 'maximize' | 'close'): void;
}
declare global { interface Window { desktop?: DesktopAPI } }
