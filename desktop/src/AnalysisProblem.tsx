// Copyright (c) 2026 AbeLin · MIT License
export function recoverySteps(reason: string): string[] {
  if (/规则.*除|除数/.test(reason)) return ['到对应数据块展开“数据处理规则”，找到提示中的规则编号和单元格。', '将固定除数或引用单元格改为正确的非零数字，再点击“读取数据并检查字段”。'];
  if (/规则/.test(reason)) return ['检查提示中的规则编号、作用区域和引用地址；所有地址都应位于该数据块选区内。', '涉及表头时开启“同时处理表头”；不适合运算的内容可选择保持原样并提示，或停用该规则。', '调整后重新读取字段，查看处理前后样例，再生成图表。'];
  if (/没有有效数字|数值字段|配对记录|非数值/.test(reason)) return ['检查数值字段是否选中了金额、数量等数字列，且表头层数正确。', '公式单元格需先在本机 Excel 计算并保存；散点图的 X、Y 要在同一行都有数字。', '展开读取后的样例和数据处理说明，修正后重新生成。'];
  if (/饼图|分类数超过|非负|总值/.test(reason)) return ['饼图只选一个非负数值字段，统计方式选择合计或有效值计数，总额须大于零。', '分类数超限时提高“图上分类数”；有退款或负数时可改用柱形图。'];
  if (/关联|追加|字段名必须/.test(reason)) return ['同结构明细选择追加；不同指标选择按共同字段关联。', '关联列应为双方都有的唯一编号，数字与文字类型应一致；重复编号先整理成每个编号一条。', '检查所有数据块的表头层数和选区，再重新读取。'];
  if (/日期|时间字段/.test(reason)) return ['时间轴选择真实日期或可识别的日期文字，不要选择金额或普通序号。', '文字月份可先整理为真实日期；只想按标签顺序连线时选择“不按日期重分组”。'];
  if (/改变|重新读取|原表字段/.test(reason)) return ['点击“读取数据并检查字段”，核对更新后的字段名和处理样例。', '重新选择分类和数值字段；看板图表可先点击编辑，再更新图表。'];
  if (/区域|表头|没有数据|清理后/.test(reason)) return ['区域可填写 A6:K246、D:D 或 7:20；每块最多 256 列、最远第 100000 行。', '核对表头层数，确认范围中有明细；检查隐藏行列和合计行是否被排除。'];
  if (/Excel|COM|原样|工作簿/.test(reason)) return ['确认本机 Microsoft Excel 可以正常启动，当前文件可以打开。', '可切换基础预览继续查看；原样预览的限制和具体原因见上方提示。'];
  return ['根据上方原因检查数据源、选区及字段设置，然后重新读取数据。', '若仍然失败，可保留这段错误信息以便排查；当前操作不会写回原 Excel 文件。'];
}
export default function AnalysisProblem({ reason }: { reason: string }) {
  return <div className="analysis-error analysis-problem" role="alert"><strong>本次操作未完成</strong><p><b>原因：</b>{reason}</p><b>可以这样处理：</b><ol>{recoverySteps(reason).map(step => <li key={step}>{step}</li>)}</ol></div>;
}
