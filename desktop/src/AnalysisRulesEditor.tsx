// Copyright (c) 2026 AbeLin · MIT License
import { useState } from 'react';
import { ArrowDown, ArrowUp, MousePointer2, Plus, Trash2 } from 'lucide-react';
import type { AnalysisBlock } from '../shared/analysis';
import { RULE_OPERATIONS, type AnalysisRule, type RuleAudit } from '../shared/analysis-rules';
const arithmetic = ['add', 'subtract', 'multiply', 'divide'];
const display = (v: unknown) => v === null || v === undefined ? '（空白）' : String(v);

export function RulesAudit({ audits }: { audits: RuleAudit[] }) {
  if (!audits.length) return null;
  return <details className="analysis-rule-audit" open><summary>规则处理记录 · 修改的是本次分析数据</summary>{audits.map(a => <section key={a.rule}>
    <strong>规则 {a.rule} · {a.target} · {a.description}</strong><p>匹配 {a.matched} 格，改变 {a.changed} 格，保持原样 {a.skipped} 格；下方最多展示 6 个变化样例。</p>
    {!!a.examples.length && <div className="analysis-table-scroll"><table><thead><tr><th>单元格</th><th>处理前</th><th>处理后</th></tr></thead><tbody>{a.examples.map(e => <tr key={e.cell}><td>{e.cell}</td><td>{display(e.before)}</td><td>{display(e.after)}</td></tr>)}</tbody></table></div>}
  </section>)}</details>;
}

export default function AnalysisRulesEditor({ block, index, disabled, suggestedColumn, onChange, onPick }: {
  block: AnalysisBlock; index: number; disabled: boolean; suggestedColumn?: string;
  onChange: (rules: AnalysisRule[]) => void; onPick: (ruleId: string) => void;
}) {
  const rules = block.rules || [], [open, setOpen] = useState(!!rules.length);
  const patch = (id: string, changes: Partial<AnalysisRule>) => onChange(rules.map(r => r.id === id ? { ...r, ...changes } : r));
  function reorder(i: number, delta: number) { const next = [...rules]; [next[i], next[i + delta]] = [next[i + delta], next[i]]; onChange(next); }
  return <details className="analysis-rules-editor" open={open} onToggle={e => setOpen(e.currentTarget.open)}>
    <summary>数据处理规则 · 可选 {rules.length ? `（${rules.filter(r => r.enabled).length} 条启用）` : '（默认使用原数据）'}</summary>
    <p>先复制选区数据 → 按下列顺序处理 → 排除空行、隐藏内容和合计 → 统计绘图。原表预览仍显示原始内容，处理后的样例在重新读取字段后显示。</p>
    <p>地址使用 Excel 行号、列标，如 D:D、7:7、D7、D7:F20。整行整列仅处理当前数据块范围；引用也须在该范围内。每条规则从上一步结果读取，文字不会当作公式执行。</p>
    {rules.map((rule, i) => {
      const prefix = `数据块 ${index + 1} 规则 ${i + 1}`;
      const math = arithmetic.includes(rule.operation);
      const reference = math || rule.operation === 'copy';
      return <fieldset key={rule.id} disabled={disabled} className={`analysis-rule ${rule.enabled ? '' : 'rule-off'}`}>
        <legend>规则 {i + 1}</legend><div className="analysis-rule-heading"><label><input type="checkbox" aria-label={`${prefix} 启用`} checked={rule.enabled} onChange={e => patch(rule.id, { enabled: e.target.checked })} />启用</label><span>{RULE_OPERATIONS[rule.operation]} · {rule.target}</span>
          <button className="icon-button" aria-label={`${prefix} 上移`} disabled={disabled || i === 0} onClick={() => reorder(i, -1)}><ArrowUp size={14} /></button><button className="icon-button" aria-label={`${prefix} 下移`} disabled={disabled || i === rules.length - 1} onClick={() => reorder(i, 1)}><ArrowDown size={14} /></button><button className="icon-button" aria-label={`${prefix} 删除`} onClick={() => onChange(rules.filter(r => r.id !== rule.id))}><Trash2 size={14} /></button></div>
        <div className="analysis-range-row"><label>作用于行、列或单元格<input aria-label={`${prefix} 作用区域`} value={rule.target} maxLength={80} placeholder="D:D / 7:7 / D7 / D7:F20" onChange={e => patch(rule.id, { target: e.target.value })} /></label><button className="button small" aria-label={`${prefix} 从原表选择`} onClick={() => onPick(rule.id)}><MousePointer2 size={13} />从原表选择</button></div>
        <div className="analysis-two-fields"><label>操作<select aria-label={`${prefix} 操作`} value={rule.operation} onChange={e => patch(rule.id, { operation: e.target.value as AnalysisRule['operation'], operand: e.target.value === 'copy' ? 'column' : 'constant' })}>{Object.entries(RULE_OPERATIONS).map(([v, label]) => <option key={v} value={v}>{label}</option>)}</select></label>
          {reference && <label>取值来源<select aria-label={`${prefix} 运算来源`} value={rule.operand} onChange={e => patch(rule.id, { operand: e.target.value as AnalysisRule['operand'], value: '' })}>{rule.operation !== 'copy' && <option value="constant">固定数值</option>}<option value="cell">固定单元格</option><option value="column">同一行的另一列</option><option value="row">同一列的另一行</option></select></label>}</div>
        {rule.operation === 'replace' && <label>查找文字（按原文匹配）<input aria-label={`${prefix} 查找文字`} maxLength={1000} value={rule.find} onChange={e => patch(rule.id, { find: e.target.value })} /></label>}
        {rule.operation !== 'clear' && <label>{reference ? { constant: '数值，例如 10000 或 10%', cell: '单元格地址，例如 D7', column: '列标，例如 E（每行各取 E 列的值）', row: '行号，例如 8（每列各取第 8 行的值）' }[rule.operand] : rule.operation === 'set_number' ? '要设置的数字' : rule.operation === 'replace' ? '替换为' : '文字内容'}<input aria-label={`${prefix} 参数`} maxLength={1000} value={rule.value} onChange={e => patch(rule.id, { value: e.target.value })} /></label>}
        <div className="analysis-rule-options"><label><input type="checkbox" aria-label={`${prefix} 同时处理表头`} checked={rule.includeHeaders} onChange={e => patch(rule.id, { includeHeaders: e.target.checked })} />同时处理表头（默认不处理）</label>
          {(math || ['replace', 'prefix', 'suffix'].includes(rule.operation)) && <label>空白或类型不符时<select aria-label={`${prefix} 类型不符处理`} value={rule.invalid} onChange={e => patch(rule.id, { invalid: e.target.value as AnalysisRule['invalid'] })}><option value="skip">保持原样并提示</option><option value="stop">停止并指出单元格</option></select></label>}</div>
        {math && <small>例如 D:D 减去同一行的 E 列，得到每行差额；要新建差额列，先将原列复制到选区内的空白列，再计算。除数为 0 会停止并提示具体位置。</small>}
      </fieldset>;
    })}
    <button className="button small" disabled={disabled || rules.length >= 20} aria-label={`数据块 ${index + 1} 添加处理规则`} onClick={() => {
      onChange([...rules, { id: crypto.randomUUID(), enabled: true, target: suggestedColumn ? `${suggestedColumn}:${suggestedColumn}` : block.range || 'D:D', operation: 'multiply', operand: 'constant', value: '1', find: '', includeHeaders: false, invalid: 'skip' }]); setOpen(true);
    }}><Plus size={14} />添加处理规则</button><span className="analysis-rules-count">{rules.length} / 20 条</span>
  </details>;
}
