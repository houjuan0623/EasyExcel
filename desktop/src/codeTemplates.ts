// Copyright (c) 2026 AbeLin · MIT License
import type { CodeEnvironment } from '../shared/types';

export const templates = [
  { id: 'analysis', title: '数据分析', code: `# df 是第一个勾选 Sheet 的完整数据，tables 包含所有勾选的表。
# 可直接修改代码，或导入另一台电脑上由 AI 生成的 .py 文件。
print("数据形状：", df.shape)
display(df.head(20), "前 20 条数据")

overview = pd.DataFrame({
    "字段": df.columns,
    "数据类型": df.dtypes.astype(str).values,
    "空值数": df.isna().sum().values,
})
display(overview, "字段概况")
save_excel(overview, "数据概况.xlsx")

numbers = df.select_dtypes(include="number")
if not numbers.empty:
    statistics = numbers.describe().round(2)
    display(statistics, "数值统计")
    save_excel(statistics, "数值统计.xlsx", index=True)
` },
  { id: 'plot', title: '绘制图表', code: `# 参数：将 None 改成你想画图的数值字段名，例如 "销售额"。
VALUE_COLUMN = None

numeric_columns = list(df.select_dtypes(include="number").columns)
if not numeric_columns:
    raise ValueError("没有数值列，请先用 pd.to_numeric 转换需要绘图的字段")
column = VALUE_COLUMN or numeric_columns[0]
values = pd.to_numeric(df[column], errors="coerce")

fig, ax = plt.subplots(figsize=(9, 4.8))
ax.plot(np.arange(1, len(values) + 1), values, color="#258365", linewidth=1.6)
ax.set(title=f"{column} · 数据趋势", xlabel="数据行", ylabel=str(column))
ax.grid(axis="y", alpha=0.2)
fig.tight_layout()
save_figure(fig, "数据趋势.png")
` },
  { id: 'split', title: '拆分表格', code: `# 参数：改成需要分组的字段名，例如 "部门"。
GROUP_COLUMN = df.columns[0]

paths = split_table(df, GROUP_COLUMN, prefix="拆分")
print(f"共生成 {len(paths)} 个文件")
# 每个不同取值单独保存一份 Excel；空值也会保留为一组。
# 此方法导出数据值。保留原表样式请使用左侧的 Excel 拆分。
` },
  { id: 'merge', title: '合并表格', code: `# 勾选多个 Sheet 后，按列名纵向合并，并记录数据来源。
merged = merge_tables(tables, source_column="来源")
print(f"已合并 {len(tables)} 个 Sheet，共 {len(merged)} 行")
display(merged, "合并结果")
save_excel(merged, "合并结果.xlsx")

# 按共同字段横向关联时，可自行使用 pd.merge(left, right, on="编号")。
` },
  { id: 'loan', title: '贷款试算', code: `# 示例参数：固定名义年利率、按月等额本息、每月期末还款，不含额外费用。
PRINCIPAL = 100000
ANNUAL_RATE = 0.036
MONTHS = 36

if PRINCIPAL <= 0 or MONTHS <= 0 or ANNUAL_RATE < 0:
    raise ValueError("请检查本金、期数与利率")
rate = ANNUAL_RATE / 12
periods = np.arange(1, MONTHS + 1)
payment = float(-npf.pmt(rate, MONTHS, PRINCIPAL))
schedule = pd.DataFrame({
    "期数": periods,
    "月供": payment,
    "本金": -npf.ppmt(rate, periods, MONTHS, PRINCIPAL),
    "利息": -npf.ipmt(rate, periods, MONTHS, PRINCIPAL),
})
schedule["剩余本金"] = (PRINCIPAL - schedule["本金"].cumsum()).clip(lower=0)
display(schedule.round(2), "还款计划（显示保留两位）")
save_excel(schedule.round(2), "还款计划.xlsx")
print("总利息（未逐期舍入）：", round(float(schedule["利息"].sum()), 2))
# 这是数学试算；逐期分币、尾期调整等应按你的业务规则另外明确。
` },
  { id: 'money', title: '金额核对', code: `from decimal import ROUND_HALF_UP, InvalidOperation

# 参数：指定需要核对的列名。逐笔保留两位、再汇总；不自动剔除合计行。
AMOUNT_COLUMN = "金额"
if AMOUNT_COLUMN not in df.columns:
    raise ValueError(f"请将 AMOUNT_COLUMN 改为实际字段，可选：{list(df.columns)}")

amounts = []
for row_number, value in enumerate(df[AMOUNT_COLUMN], start=1):
    if pd.isna(value):
        raise ValueError(f"第 {row_number} 条数据的金额为空，请先确认空值口径")
    try:
        amount = Decimal(str(value).replace(",", "").strip())
        if not amount.is_finite():
            raise InvalidOperation()
        amounts.append(amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
    except InvalidOperation:
        raise ValueError(f"第 {row_number} 条数据不是有效金额") from None

total = sum(amounts, Decimal("0.00"))
report = pd.DataFrame({"项目": ["笔数", "合计"], "结果": [str(len(amounts)), str(total)]})
display(report, "金额核对")
save_excel(report, "金额核对.xlsx")
` },
];

export function aiInstructions(environment: CodeEnvironment | null, sheets: { name: string; bookName: string; columns: string[] }[]) {
  return `请为 EasyExcel 的离线 Python 工作台编写一个可直接运行的 UTF-8 .py 脚本。

我的需求：［在这里填写目标、字段含义、计算口径、需要的图表或输出文件］

运行环境：
Python ${environment?.python || '3.12'}；${environment ? Object.entries(environment.libraries).map(([key, value]) => `${key} ${value}`).join('；') : 'pandas、numpy、matplotlib、openpyxl、xlrd'}。
目标电脑没有网络。只用上述已内置的第三方库，以及 json、csv、math、statistics、datetime、pathlib、re、collections、itertools、functools、random、decimal、fractions、sqlite3 等常用标准库。程序使用 PyInstaller 打包，不保证未列出的模块可用。不要 pip install、下载数据、调用 AI API、使用 Jupyter、GUI 或 input()。不要启动其他 Python 进程。

工作台已准备这些变量，无需硬编码文件路径：
- df：第一个勾选 Sheet 的 pandas.DataFrame。
- tables：{ "文件相对路径 / Sheet名": DataFrame }，包含所有勾选 Sheet，顺序和界面一致。
- inputs：[{key, path, sheet, headerRow, rows, columns}]，每张表的实际输入信息。
- pd、np、plt：pandas、numpy、matplotlib.pyplot；npf：numpy_financial；Decimal：decimal.Decimal。
- OUTPUT_DIR：本次运行独立结果文件夹的 pathlib.Path；当前工作目录也是这里。

读入规则：读取表头以下的完整数据，不限 100 行。隐藏 Sheet 不参与，隐藏列不载入，隐藏行保留；合计行不自动剔除。Excel 公式读已保存的缓存值，不重算。默认 pandas 推断类型和空值，编号前导零、合并表头等请根据需求明确处理。界面识别的字段可能与 df.columns 有差异，以 inputs / df.columns 为准。

可直接调用的辅助函数：
- display(data, title="分析结果")：显示 DataFrame / Series / dict（界面最多 100 行、40 列），其他对象打印为文字。
- save_excel(data, filename="结果.xlsx", index=False)：保存 DataFrame 或 {Sheet名: DataFrame} 到结果目录，自动避免重名覆盖，返回 Path。
- save_figure(figure=None, filename="图表.png")：保存并展示 PNG，关闭该图；也支持 plt.show()。未关闭的图在脚本结束时自动保存。
- split_table(data, column, prefix="拆分")：按字段分组导出多个 xlsx，保留空值组，返回 Path 列表。
- merge_tables(source_tables=None, source_column="来源")：按列名纵向合并字典中的表（默认 tables），返回 DataFrame；来源列重名时报错，可另起列名或传 None。

代码要求：
1. 在开头集中定义需要调整的参数，并添加中文注释。
2. 校验需要的字段，缺失时用中文 ValueError 说明；不要猜测真实业务计算口径。
3. 所有输出使用上述保存函数或 OUTPUT_DIR；不写回、删除或移动源文件。
4. 图表使用已配置的 Agg 后端和本机中文字体；不要切换为 GUI 后端。
5. print() 用于日志，display() 用于预览；完整结果请 save_excel()。
6. 每次运行都是新进程，变量不会沿用。导出为数据值，不自动保留原 Excel 样式。
7. 不要导入不在内置清单里的第三方库；如果做不到，先说明需要增加什么库。
8. 财务计算请明确金额单位、期间、利率口径、现金流正负方向和舍入规则。精确金额使用 Decimal(str(value))，不要 Decimal(float)。numpy_financial.npv 将第一个现金流视为 t=0；XIRR / XNPV 使用实际日期。不要静默跳过无效金额或重复计算表中合计行。

可用库及导入方式：
${environment?.capabilities?.map((item) => `- ${item.name} ${item.version}：${item.purpose}；${item.example}`).join('\n') || '运行环境检查完成后会显示内置库的详细清单。'}

当前勾选的表头说明（没有单元格数据；带到联网电脑前可自行替换字段名）：
${sheets.length ? sheets.map((s) => `- ${s.bookName} / ${s.name}：${s.columns.filter(Boolean).join('、')}`).join('\n') : '尚未选择数据源，请自行填写字段名和含义。'}

请只给完整可执行代码，不要 Markdown 代码围栏。
`;
}
