// Copyright (c) 2026 AbeLin · MIT License
export const RULE_OPERATIONS = { add: '加上', subtract: '减去', multiply: '乘以', divide: '除以', set_number: '设置数字', set_text: '设置文本', replace: '替换文字', prefix: '添加前缀', suffix: '添加后缀', clear: '清空', copy: '复制另一处的值' } as const;
export interface AnalysisRule {
  id: string; enabled: boolean; target: string; operation: keyof typeof RULE_OPERATIONS;
  operand: 'constant' | 'cell' | 'column' | 'row'; value: string; find: string;
  includeHeaders: boolean; invalid: 'skip' | 'stop';
}
export interface RuleAudit {
  rule: number; target: string; description: string; matched: number; changed: number; skipped: number;
  examples: { cell: string; before: string | number | boolean | null; after: string | number | boolean | null }[];
}
export function parseRules(value: unknown): AnalysisRule[] {
  if (value === undefined) return [];
  if (!Array.isArray(value) || value.length > 20) throw new Error('每个数据块最多设置 20 条处理规则');
  const rules = value.map((v): AnalysisRule => {
    if (!v || typeof v !== 'object' || Array.isArray(v)) throw new Error('数据处理规则格式无效');
    const r = v as AnalysisRule;
    if (typeof r.id !== 'string' || !/^[a-zA-Z0-9_-]{1,80}$/.test(r.id) || ['__proto__', 'constructor', 'prototype'].includes(r.id)) throw new Error('规则标识无效');
    if (typeof r.target !== 'string' || r.target.length > 80 || !r.target.trim()) throw new Error('请填写规则作用区域');
    if (!Object.hasOwn(RULE_OPERATIONS, r.operation) || !['constant', 'cell', 'column', 'row'].includes(r.operand) || !['skip', 'stop'].includes(r.invalid)) throw new Error('规则包含不支持的操作');
    if (typeof r.value !== 'string' || r.value.length > 1000 || typeof r.find !== 'string' || r.find.length > 1000) throw new Error('规则文字最多 1000 个字符');
    if (typeof r.enabled !== 'boolean' || typeof r.includeHeaders !== 'boolean') throw new Error('规则开关格式无效');
    return { id: r.id, enabled: r.enabled, target: r.target, operation: r.operation, operand: r.operand, value: r.value, find: r.find, includeHeaders: r.includeHeaders, invalid: r.invalid };
  });
  if (new Set(rules.map(r => r.id)).size !== rules.length) throw new Error('规则标识重复');
  return rules;
}
