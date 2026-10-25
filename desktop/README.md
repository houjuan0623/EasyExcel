# EasyExcel · Electron 桌面界面

Electron + React + TypeScript 只负责界面、系统文件对话框和进程通信。**读取工作簿、获取 Sheet、预览单元格、统计分组、拆分和汇总全部由 Python 完成**，沿用根目录 `core/` 的处理引擎。没有 JavaScript Excel 库，也没有远程接口。

0.3.1 起使用 EasyExcel 名称；0.6.1 起图标更新为绿色表格与字母 E 结合的专属 Logo，统一用于界面、窗口、可执行文件与安装快捷方式。为兼容旧版本，用户数据仍保存在 `%APPDATA%/excelrouter-desktop`，原脚本库、历史记录、编辑草稿以及旧版 JSON 备份继续可用。

0.6.2 修复复杂合并表头在字段预览与实际拆分时识别不一致的问题：拆分复用同一份有限表头扫描，保留纵向合并的字段名；没有生成分表时明确报错，不会只返回一个空的输出目录。

0.7.0 新增原表 VBA 结构包与导出前检查清单；保留真实字段和原公式，支持读取 .xlsm，详见下文。

0.8.0 新增统计分析与 dashboard。按图表用途和注解选区，支持跨 Sheet 组合、描述统计、方案复用和离线报告，全部计算仍由本机 Python 完成。

图标资源和重新转换说明见 [assets/README.md](assets/README.md)。

## 启动

需要 Node.js 22.12+（推荐 24）和 Python 3.9+。在项目根目录运行：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r desktop/requirements.txt
cd desktop
npm ci
npm run dev
```

`npm run dev` 启动 Vite 和 Electron，界面修改自动刷新；修改 `electron/` 主进程或 preload 后重启该命令。第一次启动 Electron 会下载官方运行时，需要联网；依赖准备好后，表格处理完全离线。

已经有 Python 环境时，可在启动前指定解释器：

```powershell
$env:EXCELROUTER_PYTHON = 'C:\你的环境\python.exe'
npm run dev
```

未设置时优先使用项目根目录 `.venv`，否则使用 PATH 中的 `python`（Windows）或 `python3`（Linux/macOS）。Linux 虚拟环境的解释器位于 `.venv/bin/python`。

构建并运行本地生产界面：

```powershell
npm run build
npm start
```

`build` 只生成 `dist/` 和 `dist-electron/`，供源码运行。生成独立安装包请使用下面的 `dist:win`。

## 打包 Windows 安装包

在 **Windows x64** 上构建，需要 64 位 Python 3.9+、Node.js 22.12+（推荐 24）。在项目根目录依次执行；已有 `.venv` 时跳过创建虚拟环境：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r desktop/requirements-build.txt
cd desktop
npm ci
npm run dist:win
```

`dist:win` 依次编译 TypeScript 界面、用 PyInstaller 打包 Python 引擎并自检、用 electron-builder 生成 NSIS 安装程序。第一次构建需要联网下载 Electron 和打包工具。构建产物位于：

```text
desktop/release/
├── EasyExcel-Setup-0.8.0-x64.exe    可分发的完整安装程序
└── win-unpacked/
    └── EasyExcel.exe              打包后程序，可直接运行验证
```

安装包名称中的版本来自 `desktop/package.json`。后续修改代码后，在 `desktop` 中重新执行 `npm run dist:win` 即可。`npm run pack:win` 只生成 `win-unpacked` 目录，省略安装程序。

**安装与分发：**将 `EasyExcel-Setup-0.8.0-x64.exe` 发给使用者，双击、选择安装目录并完成向导，再通过桌面或开始菜单的 EasyExcel 快捷方式启动。默认安装到当前用户目录；使用者不需要安装 Python、Node.js 或项目源码。若使用 `win-unpacked`，必须复制整个目录，不能只拷贝其中的 `.exe`。

Excel 操作仍由安装包中的 Python 引擎执行：`resources/backend/EasyExcelBackend.exe` 及同目录依赖。安装版使用该内置引擎，不依赖 `EXCELROUTER_PYTHON` 或系统 PATH。安装包目前未做代码签名，Windows 可能显示“未知发布者”；正式公开分发时可配置自己的签名证书。

卸载可通过 Windows「设置 → 应用 → 已安装的应用 → EasyExcel」完成。卸载保留本机最近任务记录，用户保存的 Excel 结果文件也不属于程序卸载范围。

打包配置见 `electron-builder.cjs`、`scripts/build-backend.cjs`、`scripts/package-win.cjs`。这里的构建命令只产出本地文件，不自动发布。根目录的 `build.bat` 和现有发行工作流用于原 Python GUI；Electron 新版请使用本节命令。

相关配置参考 [electron-builder NSIS 文档](https://www.electron.build/v26/docs/nsis/) 与 [PyInstaller 打包文档](https://pyinstaller.org/en/stable/usage.html)。

## 打包 Windows 免安装版

准备好上一节的打包依赖后，在 `desktop` 目录运行：

```powershell
npm run dist:portable
```

输出 `desktop/release/EasyExcel-Portable-0.8.0-x64.zip`，包含完整的 Electron 程序、Python 引擎、离线分析依赖和中文使用说明（版本号随 `package.json` 更新）。将 ZIP 放到 Windows 64 位电脑，右键「全部解压缩」到独立文件夹，再双击其中的 `EasyExcel.exe` 即可。无需安装程序、Python、Node.js 或 Microsoft Excel，可离线处理表格。

请保留解压后的全部文件，不要只复制 `.exe`，也不要直接在压缩包内运行。最近任务记录仍保存在当前 Windows 用户的应用数据目录；移动程序文件夹不会迁移这些记录。

验证解压到其他目录的免安装程序时，可指定实际入口：

```powershell
$env:EXCELROUTER_PACKAGED_EXE = 'D:\工具\EasyExcel\EasyExcel.exe'
npm run test:packaged
Remove-Item Env:\EXCELROUTER_PACKAGED_EXE
```

## 已实现

- 文件选择、文件夹递归扫描、拖放，以及 `Ctrl+O` 快捷键。
- 按文件分组的 Sheet 多选、全选、清空、搜索；默认选中每个文件的首个可处理 Sheet。
- 隐藏 Sheet 列出但不可选。空白页和未识别表头的 Sheet 不处理，可在高级设置中指定表头行后重新识别。
- 点击「预览」查看原表第 1～100 行，包含标题、表头和空白行；保留合并单元格、行高列宽、字体、颜色、边框、对齐、换行及常见数字格式。支持 50%～200% 缩放、展开预览和双向滚动；隐藏行列不显示。Sheet 列表行数来自文件声明的尺寸，因此标为「约」。
- 拆分字段取所有勾选 Sheet 的共同列；分组统计由 Python 全量扫描所选 Sheet 后返回。分组支持勾选和搜索。
- 标准模式保留格式；极速模式仅保留数据值；可选保留同行公式、目录批量 ZIP、跨文件合并。
- 普通汇总按列名对齐，只合并勾选的 Sheet，保留来源列。表头差异过大的表由原引擎跳过并记入汇总说明。
- 拆分设置中可开启「同时生成总工作簿」，额外生成 `总工作簿/拆分结果总表.xlsx`；每份分表单独占一个 Sheet，原分表与 ZIP 照常保留。
- Excel 汇总新增「每个表独立一个 Sheet」方式，可将已有分表或不同结构的表整理成一个工作簿。
- 始终可见的底部操作栏、真实进度、协作式停止、最多 500 条界面日志、结果目录入口、本机最近 20 次任务。
- UTF-8 JSON Lines 通信；更换文件、Sheet 或表头后忽略过期查询。空 Sheet / 空分组选择禁止执行，Python 再次校验。

新版当前覆盖 **Excel 拆分、普通汇总和离线 Python 工作台**。查收、方案、二级到人、加密分发等原有功能仍可运行根目录 `python main.py` 使用。旧界面和原有 CLI 调用方式不变。

## 将分表整理成一个多 Sheet 工作簿（0.4.0）

有两种入口，Excel 读写仍由 Python 完成：

1. **拆分时同时整理**：在 Excel 拆分的设置中开启「同时生成总工作簿」，然后照常拆分。结果目录下的 `总工作簿/拆分结果总表.xlsx` 包含本次成功导出的分表；无归属核对表和 ZIP 中的副本不会重复收入。
2. **整理已经拆好的文件**：打开 Excel 汇总，选择保存分表的文件夹（可直接选某次拆分的日期结果目录），勾选需要的 Sheet，在「汇总方式」里选择「每个表独立一个 Sheet」，点击开始汇总。结果为 `工作表合集_日期_时间.xlsx`，各表无需具有相同的列结构。

每个文件只选择一张 Sheet 时，以文件名命名；同一文件选择多张时，以「文件名_原 Sheet 名」命名，每张都会保留。非法字符自动替换，名称最长 31 字，重名自动加序号。所有数据行都会处理，前 100 行仅是界面预览限制。

保留常见单元格样式、数据类型、合并区域、行高列宽、冻结窗格、筛选范围和打印布局；隐藏 Sheet 不处理，隐藏列的数据沿用现有规则不带出。普通整理读取公式已保存的计算值，请先在 Excel/WPS 计算并保存；拆分时启用「保留同行公式」的分表，收入总工作簿后仍保留这些公式。图片、图表和 Excel 表对象不复制；依赖其他表/定义名称的验证和条件格式会跳过并提示。`.xls` 主要保留数据值。

## 统计分析与 dashboard（0.8.0）

左侧「统计分析」支持 8 类组件：指标卡、柱形、折线、饼图、直方图、散点、箱线与 Pearson 相关性热力图。每类提供用途、计算方法和字段选择说明。数据块可选择不同工作簿/Sheet 的矩形、整行或整列，支持鼠标拖选、Shift 扩展、分页跳转、多层合并表头和横向数据转置。

点击「读取数据并检查字段」确认样例和口径，再绑定字段生成预览、加入看板。同结构记录按字段名追加；不同指标通过单一共同字段一对一关联，拒绝重复键造成的记录放大。空键、未匹配、缺失公式缓存、隐藏行列和合计记录的处理会显示。数值不自动补零，单位只作标注。

看板最多 12 张图，可编辑、排序、刷新、移除；配置自动本机保存并支持导出/导入 JSON。离线报告包含自包含 HTML、图表 PNG、图表数据与描述统计 CSV。真实结果只写用户选择的本机目录，不属于脱敏包。选区每张图最多 4 块、500,000 个单元格，每块最多 256 列（转置后最多 256 字段），最远第 100,000 行，单任务最多 5 分钟。

完整步骤、图表对照表、关联方式、统计口径和测试入口见 [统计分析指南](../docs/统计分析指南.md)。源表内嵌图表的「预览」和此处根据用户选区新生成的统计图相互独立。

## 原表 VBA 结构包（0.7.0）

「脱敏导出」现在默认选择「VBA 原表模板」。用于让 AI 生成粘贴到原工作簿标准模块中的 VBA。

1. 选择 `.xlsx` 或 `.xlsm`，勾选至少一个可处理 Sheet，再点击「脱敏导出」。原有宏不执行。
2. 所选文件的全部 Sheet 都会列出，包括隐藏 Sheet；跨表公式保持引用目标。展开各 Sheet 核对表头常量、表对象名称、字段和工作表代号。
3. 根据实际表头调整范围，例如 `A7:K9;M6:R6`。范围内常量原样保留，范围外常量清空；Excel 表对象的表头必须包含。修改后点「更新检查清单」。
4. 查看命名区域和公式检查清单；固定文本、金额等数字、外部路径会按原公式保留。相同项归类显示，可搜索、翻页，每类显示最多 8 个位置和示例公式。
5. 勾选「已检查」后生成 ZIP。文件内容或表头范围变动时，必须重新检查。源文件不修改。

ZIP 包含编程说明、完整结构 JSON，以及每本源文件独立的 `VBA结构模板.xlsx` 和 `运行前结构检查.bas`。多个文件的同名 Sheet 不会合并或改名。保留原表头、Sheet 名/顺序/隐藏状态、单元格位置、合并区域、全部单元格公式（含共享/数组公式）、表对象名称/范围/列名/公式、命名区域和名称公式。

非表头数据常量及公式计算缓存均清空；图表、透视缓存、图片、批注、条件格式、验证、筛选条件、原宏、文档属性和外部连接移除。外部公式引用路径保留供检查，但移除外部值缓存，模板设置手动计算且不自动更新链接。基础格式会清理，不承诺样式无损。原保护状态写入结构说明，模板本身不加锁。

交给 AI 时说明具体任务，让它按包内约定使用 `ThisWorkbook`、原表头和命名区域，运行时确定数据范围，避免从空模板推断末行，避免覆盖公式列。`.bas` 只提供少量表头定位检查，无业务逻辑，可在 VBE 中导入；AI 应提供完整可粘贴代码（去掉 `Attribute VB_Name` 行）。`.xlsx` 另存为 `.xlsm` 后，将代码放入原工作簿标准模块，先在本机副本验证。此功能不执行 VBA，VBA 运行需要桌面 Excel。

本模式按用户确认保留名称、表头、公式固定值、外部路径和原行列范围，不是完全匿名模式。每次最多 16 本工作簿，每本最多 64 张工作表、100 万个实际单元格、解压后 200 MB；每表最多 20,000 个表头单元格。旧 `.xls`、独立图表 Sheet、Excel 4.0 宏表、DDE/OLE 外部链接暂不支持，会报错而非静默丢失。

## 严格脱敏与 AI 演示包（0.6.0）

1. 在离线电脑选择原表和需要的 Sheet，确认运行设置中的表头行。
2. 点击工作表列表下方「脱敏导出」，选择「Python 匿名样例 + 安全公式」。此模式的原文件名、Sheet 名和字段名只在本机选择器里显示。
3. 默认选择「虚构样例 + 安全公式」和 20 行。按列选择编号、姓名、类别、整数、金额、小数、比例、日期、布尔、电话、邮箱或文本；也可选择 5/50 行、空白结构或仅字段说明。
4. 保存 ZIP。里面只有匿名的 `虚构演示.xlsx`（或 `空白模板.xlsx`）、`结构说明.json`、`给AI的说明.txt`；仅字段说明模式没有 Excel 文件。
5. 只把 ZIP 交给 AI，补充抽象的字段含义和分析目标。AI 使用匿名名编写脚本，真实业务结论要在本机运行后才能得到。
6. 将 `.py` 带回原电脑，选择相同文件、Sheet 和表头后在 Python 工作台运行。包内说明会让 AI 使用 `ai_tables, ai_params = load_ai_context("导出标识")`；`ai_tables["Sheet_001"]` 的列名保持 `字段_001` 等匿名形式，但数据来自本机原表，参数也从本机映射读取。

虚构生成器不使用原记录的值、分类、数量、极值或分布；固定示例中的金额、类别和日期不能反映真实情况。模板保留上方表头布局、列宽及可转换的代表性公式，数据区使用统一演示样式和固定行数。原表含大量汇总区、多块布局或逐行不同公式时，演示模板不会完整复制它们；以匿名结构和包内限制说明为准，不要用演示表覆盖原表。

公式只接受允许的常见函数和所选表内的 A1 引用。数字、文字和逻辑常量变成演示参数，原值保存在本机；跨行范围规范为固定演示范围，不泄露源总行数。不支持外部/命名/结构化/动态引用和部分高级公式，不能安全转换时用虚构值或空白替代并提示，不导出原公式文本。图表、透视缓存、图片、批注、超链接、验证、条件规则、宏、连接、文档属性和原公式缓存均不复制；自定义数字格式中的原文字也不复制。导出后会检查 XLSX 内部部件。

本机映射保存在 `%APPDATA%/excelrouter-desktop/privacy-mappings/<导出标识>.json`，包含真实路径、名称及原公式常量，不在对外 ZIP 中。`load_ai_context` 会核对文件、Sheet、表头、列名及代表性公式；这些结构变化时需重新导出。映射不会随脚本库备份迁移，换电脑应在新离线电脑重新导出并更新脚本里的标识。只分享指定 ZIP，原文件、用户目录、日志和实际分析结果都不属于脱敏包。

当前仅支持 `.xlsx`，一次最多 32 张 Sheet，每表最多 256 列，表头须位于前 100 行，工作簿解压后不超过 200 MB，本机映射不超过 8 MB。隐私模式使用新建文件重建模板，源文件不修改。普通「给 AI 的说明」保留原名称和表头，与本节的严格脱敏导出不同。

新功能默认关闭，原拆分与按列名汇总的行为保持兼容。总工作簿停止生成或保存失败不会删除分表；已有同名结果不会覆盖。

## 离线 Python 工作台

用于“在有网电脑让 AI 写脚本，拷到无网电脑运行”。在界面选择 Excel 和 Sheet 后，打开 **Python 工作台 → 给 AI 的说明 → 导出说明 .txt**，填写目标、字段含义和计算口径，让外部 AI 按这份说明生成 UTF-8 `.py` 文件，再拷回本机点击 **导入 .py → 运行代码**。说明包含选中表的名称和表头，不含数据行；可自行将名称改成通用名称并提供虚构样例。

编辑器支持 Python 语法高亮、行号、缩进、撤销、查找和 `Ctrl+Enter` 运行。草稿自动保存在当前用户应用数据中；保存为 `.py` 可跨电脑复用。内置数据分析、绘图、拆分、合并、贷款试算和金额核对示例。运行日志可复制，报错附完整 Python 调用栈和脚本行号。新增运行不继承上次变量。

免安装包携带 pandas、numpy、matplotlib（Agg）、openpyxl、xlrd 及下文的财务依赖，具体版本显示在运行设置中，并写入每次结果目录的 `运行环境.json`。**目标电脑无需 pip、联网或外部 Python**。额外第三方库要在有网开发机加入打包依赖后重建；本版本不支持在界面在线安装库或使用 Jupyter / GUI / 交互式 `input()`。PyInstaller 只收集所需模块，打包时显式加入了 statistics、sqlite3、fractions、decimal、csv、random 等常用标准库。

| 变量 / 方法 | 用法 |
| --- | --- |
| `df` / `tables` | 第一张勾选表的 DataFrame / 所有勾选表的字典，键为 `文件相对路径 / Sheet名` |
| `inputs` | 各表的来源、Sheet、表头行、实际行数和实际列名 |
| `pd`, `np`, `plt` | pandas、numpy、matplotlib.pyplot |
| `npf`, `Decimal` | numpy_financial、decimal.Decimal，可直接用于财务计算 |
| `OUTPUT_DIR` | 本次独立结果目录的 Path；当前工作目录也是这里 |
| `display(data, title="分析结果")` | 展示 DataFrame / Series / dict；其他对象输出文字 |
| `save_excel(data, filename="结果.xlsx", index=False)` | 导出单表或 `{Sheet名: DataFrame}`；自动避免同名覆盖，返回 Path |
| `save_figure(figure=None, filename="图表.png")` | 保存 PNG 并在界面预览，然后关闭该图；也支持 `plt.show()` |
| `split_table(data, column, prefix="拆分")` | 按指定列分组导出，包含空值组，返回 Path 列表 |
| `merge_tables(source_tables=None, source_column="来源")` | 按列名纵向合并（默认 tables）；来源列已存在时报错，可另起名或传 None |

运行读取所选 Sheet **表头以下的全部数据**，前 100 行仅限制原表预览；隐藏 Sheet 不参与、隐藏列不载入，隐藏行和合计行保留。表头可在运行设置中手动指定。pandas 默认推断类型与空值，编号前导零等需在脚本中按业务要求处理；公式读取文件已保存的缓存值，不进行 Excel 公式重算。脚本导出保留数据值；需要保留原表样式时用原来的拆分/汇总页面。

每次运行建立 `Python_日期_时间_随机号` 文件夹，保存 `script.py`、运行环境、输入说明、日志及生成文件。辅助保存函数只写本次目录；用户 Python 代码拥有当前用户权限，**不是文件访问沙箱**，因此直接使用路径读写时应明确输出到 `OUTPUT_DIR`。目录扫描会跳过这些运行目录。可停止运行，默认 2 分钟到时终止，可调整为 30 秒、10 分钟、1 小时；终止的文件可能未写完整。最近任务中保留脚本结果目录入口。

结果表最多显示 100 行、40 列，单格最多 1,000 字符；完整结果应保存 Excel。界面最多展示 40 项结果，PNG 单张最多 4 MB、合计 20 MB（超过只保存文件）。控制台最多记录 200,000 字符，界面保留末尾 500 条。图表使用本机中文字体，通常为 Windows 微软雅黑；没有中文字体时应安装合适字体或使用英文标签。

## 本地脚本库

在 **Python 工作台 → 我的脚本库** 中管理脚本。每次点击运行都会先保存代码，再执行；运行成功、报错、停止和超时都保留记录。按代码内容去重（统一换行符），相同代码累计运行次数，修改后的代码保留为另一份记录。也可点击 **存入脚本库** 保存尚未运行的代码。

- 可修改名称、分类、使用说明，按名称 / 分类 / 说明搜索，按分类筛选和收藏。
- **载入编辑器**只加载；**载入并运行**加载后立即使用当前工作台勾选的 Sheet。脚本不绑定旧数据源路径。未选择数据源时仍可加载编辑，运行按钮保持禁用。
- 加载前，当前非内置示例、尚未收录的编辑内容会自动保存为「草稿」，避免切换脚本丢失修改。
- 移入回收站后可恢复；回收站不会删除历史结果文件。每份记录显示运行次数、最近时间和状态。
- 备份导出为 JSON，包含全部脚本、说明、收藏和回收站状态，不包含 Excel 输入数据。脚本内自己写入的业务常量仍属于代码，会随备份带出。
- 导入备份会按代码内容合并，已有名称 / 说明 / 分类优先保留，收藏合并；运行次数取最大值，重复导入不会重复累计。支持最大 64 MB 的备份。

存储位于 Electron 用户数据目录下的 `script-library/`，与应用安装目录和运行结果目录独立；更新 / 移动程序后，当前 Windows 用户的脚本库仍然保留。换电脑请使用「备份 → 导入备份」。每份记录为独立 JSON，通过临时文件替换写入；损坏记录会提示并保留原文件，可用同一脚本的备份恢复。导入前保存一份 `before-import.backup`；修复的原损坏文件另存为 `.damaged`。

升级时，程序会尝试从最近任务中仍存在的 `script.py` 自动收录旧版本代码。旧结果已移动、删除或超出最近 20 次记录时，可手动导入 `.py` 再存入脚本库。

## 内置离线财务库

下面是 0.3.0 开发与打包环境实际验证的版本。界面 **运行设置与内置环境 → 查看内置库的用途与导入方式** 和导出的 AI 编程说明均使用实际运行版本。

| 库 | 版本 | 适用工作 | 导入方式 |
| --- | --- | --- | --- |
| [numpy-financial](https://numpy.org/numpy-financial/) | 1.0.0 | NPV / IRR、贷款月供、本金利息、现值终值 | `import numpy_financial as npf` |
| [PyXIRR](https://anexen.github.io/pyxirr/) | 0.10.8 | 实际日期现金流的 XIRR / XNPV | `import pyxirr` |
| [SciPy](https://docs.scipy.org/doc/scipy/reference/optimize.html) | 1.18.1 | 统计检验、拟合、求解和优化 | `from scipy import stats, optimize` |
| [statsmodels](https://www.statsmodels.org/stable/user-guide.html) | 0.15.0 | 回归、时间序列和统计预测 | `import statsmodels.api as sm` |
| [QuantLib](https://www.quantlib.org/) | 1.44 | 计息日历、现金流、债券与利率估值 | `import QuantLib as ql` |
| [XlsxWriter](https://xlsxwriter.readthedocs.io/) | 3.2.9 | 带格式、公式和图表的 Excel 财务报表 | `import xlsxwriter` |
| [Seaborn](https://seaborn.pydata.org/) | 0.13.2 | 统计分布、比较、相关性图表 | `import seaborn as sns` |
| [Babel](https://babel.pocoo.org/en/latest/api/numbers.html) | 2.18.0 | 多币种金额和数字格式化 | `from babel.numbers import format_currency` |
| decimal | Python 内置 | 精确十进制金额与显式舍入规则 | `from decimal import Decimal, ROUND_HALF_UP` |

原有 pandas / NumPy / Matplotlib / openpyxl / xlrd 继续可用。金额、汇率、利率、交易日和估值数据由本地文件或脚本参数提供；这些计算库不提供自动联网获取行情的流程。金额计算应明确逐笔还是合计后舍入；用 `Decimal(str(value))` 保留十进制意义。`npf.npv` 的第一个现金流对应 t=0，日期不规则时可用 PyXIRR；模板中已注明还款频率、利率和舍入假设。

打包保留所用依赖的元数据与许可证，在 `resources/backend/离线依赖清单.json` 提供版本清单。`core/` 继续只使用原来的 Excel 依赖，财务库限于 Electron 的脚本工作台。

## 代码边界

```text
src/                 React + TypeScript 界面和样式
shared/types.ts      界面 / preload / 主进程共用类型
electron/preload.ts  最小化的白名单 API
electron/main.ts     窗口、IPC、文件对话框、本机历史记录
electron/python.ts   Python 子进程、JSON Lines、退出回收
electron/script-library.ts  脚本存储、去重、收藏、回收站、备份迁移
bridge.py            Python 参数校验、元数据、预览、核心调用
code_runner.py       离线脚本运行、DataFrame、结果展示与保存
runtime_libraries.py 离线分析与财务库清单（pandas 等只在 desktop 内）
../core/             Excel 拆分 / 汇总引擎
```

渲染器启用 sandbox、contextIsolation，关闭 nodeIntegration；阻止新窗口、页面跳转和浏览器权限申请。生产构建使用仅本地脚本的 CSP，开发环境只为本机 Vite 注入 nonce。Excel 单元格按文字渲染。路径和字段通过 stdin JSON 传入 Python，不拼接 shell 命令。只有本次或最近任务的结果路径能通过「查看结果」打开。

新增核心配置是可选的，省略时保留原来处理所有可见 Sheet 的行为：

- `selected_sheets: ["Sheet1"]`：所有输入文件按名称筛选。空数组表示不选任何 Sheet。
- `sheet_selection: {规范化绝对文件路径: ["Sheet1"]}`：精确到每个文件的 Sheet 选择，优先于 `selected_sheets`。文件不在映射中或列表为空时不处理。路径键使用 Python 的 `os.path.normcase(os.path.abspath(path))`。

标准拆分 / 汇总和内置脚本示例不会写入原文件。自定义 Python 脚本按其代码执行。最近任务只在 Electron 用户数据目录保存结果路径与处理摘要，不上传。

预览读取和格式转换位于 `preview.py`、`preview_format.py`、`preview_styles.py`、`preview_drawings.py`，仍全部使用 Python；TypeScript 展示返回的网格和 PNG。`.xlsx` 使用流式读取，仅保留前 100 行单元格；`.xls` 通过 xlrd 读取基础格式。预览最多 256 列，超出会明确提示。公式优先显示文件保存的缓存结果，没有缓存时显示公式文本并提示，不重新计算公式；字体取决于本机可用字体。

### 统计图与复杂样式预览（0.5.0）

选择文件后，点击某个 Sheet 的「预览」。前 100 行中的图片、图表显示在原锚点附近，可双向滚动、缩放或「展开预览」。取消「显示图表和图片」可查看被遮住的单元格。下方「统计图与图片」包含该工作表的所有已读取对象（也包括第 100 行之后的图表），点击缩略图即可单独放大，Esc 关闭。

- 统计图：二维柱形/条形、折线、面积、饼图、单系列圆环、散点、雷达，支持常见堆积、百分比堆积和柱形/折线/面积组合。双坐标轴保留独立比例，并读取标题、图例、系列颜色、轴标题、常见数字轴格式和边界。Python 使用 Matplotlib 离线重绘，外观不保证逐像素一致。
- 样式：原有字体、边框、对齐、合并区域之外，新增渐变/图案填充、斜线边框、单格富文本、上下标、负数格式颜色。条件格式支持数值比较、文字、空白、简单比较公式、色阶、数据条和图标集；按优先级及 `stopIfTrue` 应用。图标外观、渐变和数据条可能近似显示。
- 图表和条件格式按引用范围读取数据，不局限于前 100 行；隐藏行、列和 Sheet 不进入图表数据。不读取外部文件或网络链接，不执行工作簿宏或条件公式代码。无单元格公式缓存时，部分图表可使用文件内已保存的图表缓存并提示。
- 本次最多读取 16 个图片/图表，返回 PNG 合计最多 8 MB；图表每系列最多 5,000 个点、每组最多 20 个系列。图表与条件格式引用累计最多读取 100,000 个单元格、最远第 100,000 行，条件格式最多 200 组。超过上限会提示，单元格仍可预览。
- 三维图、透视/扩展图、复杂组合、SmartArt、命名区域引用、复杂条件公式及旧 `.xls` 图表暂不支持。特殊数字格式和部分高级 Excel 样式可能回退显示。此功能仅改进预览；拆分和多 Sheet 汇总仍按原规则处理，**不自动复制图片或图表**。

样式读取依据 [openpyxl 样式文档](https://openpyxl.readthedocs.io/en/stable/styles.html) 和 [xlrd API 文档](https://xlrd.readthedocs.io/en/latest/api.html)。

## 验证

```powershell
# 在 desktop 目录
npm run build
npm run test:smoke
npm run test:library

# 生成安装包后验证打包版（打包版应用不使用开发机的 Python 环境）
npm run test:packaged

# 可选：验证 Vite 开发模式的完整流程
$env:EXCELROUTER_SMOKE_DEV = '1'
npm run test:smoke
Remove-Item Env:\EXCELROUTER_SMOKE_DEV

# 在项目根目录，安装 pytest 后
.\.venv\Scripts\python.exe -m pytest tests/test_desktop_bridge.py -q
.\.venv\Scripts\python.exe -m pytest tests -q
```

端到端测试用 Playwright 驱动真实 Electron 窗口，仅替代系统文件选择对话框。工作簿由 Python 生成到系统临时目录，验证 Sheet 切换、空选禁用、拆分格式与内容、汇总行数、停止任务、最近任务、小窗口布局及进程隔离，截图也保存在临时目录。可以用 `EXCELROUTER_TEST_DIR` 指定仓库外的测试根目录，每次建立独立子目录。测试无需读取用户表格。

图表样式回归位于 `tests/test_desktop_preview_visuals.py`，验证全范围统计、隐藏数据排除、双轴比例、7 类图表、内嵌图片、条件格式优先级、富文本及外部关系隔离。Electron 测试另检查图表缩略图、放大弹窗、显示开关、渐变、数据条和图标集的真实 DOM 与截图。

脚本工作台的端到端检查还包含 `.py` 导入与保存、AI 编程说明导出、真实分析表和 PNG、拆分/合并辅助函数、`Ctrl+Enter`、中文错误日志复制、无限循环停止和运行超时。`tests/test_desktop_code.py` 验证超过 100 行的完整数据、隐藏列、源文件哈希不变、空值分组、输出重名、日志及预览容量限制。

`test:packaged` 直接启动 `release/win-unpacked/EasyExcel.exe`，移除应用 PATH 中的开发工具并指定无效的外部 Python 路径，确认实际使用内置引擎。测试脚本自身仍使用开发环境中的 Python 创建和校验测试表格；安装后的最终用户无需此环境。

通信和隔离方式参考 [Electron IPC 文档](https://www.electronjs.org/docs/latest/tutorial/ipc) 与 [Context Isolation 文档](https://www.electronjs.org/docs/latest/tutorial/context-isolation)。
