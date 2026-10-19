# EasyExcel

<img src="desktop/assets/app.png" width="72" height="72" alt="EasyExcel 图标" />

**可离线使用的 Excel 与 Python 工作台。** Electron + TypeScript 负责界面，Python 负责 Excel 操作、绘图和数据分析。当前桌面版为 **0.4.0**，面向 Windows x64。

## 功能

- 选择 Excel 文件或文件夹，按文件勾选多个 Sheet。
- 预览原表前 100 行，保留常见样式与合并单元格，支持放大查看。
- 按字段拆分和汇总 Excel，沿用 Python 表格处理引擎。
- 可将拆分结果整理成一个总工作簿，每份分表独立占一个 Sheet；也可直接整理已经拆好的文件。
- Python 编辑器支持导入、保存和运行 `.py`，展示日志、结果表和图表。
- 脚本库自动收录运行过的代码，支持去重、分类、收藏、搜索、再次加载、回收站和备份迁移。
- 内置财务与分析依赖；可在联网电脑让 AI 编写脚本，再导入离线电脑运行。

脚本读取所选 Sheet 的完整数据，前 100 行仅为原表预览限制。文件在本机处理，程序不需要连接远程服务。

## 开发运行

使用 Windows x64、Python 3.12 和 Node.js 22.12+（推荐 24）。在项目根目录执行：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r desktop/requirements-build.txt
cd desktop
npm ci
npm run dev
```

首次准备依赖需要联网。最终免安装包携带 Electron、Python 引擎和所需库，目标电脑无需另装 Python、Node.js 或 Microsoft Excel。

## 生成免安装包

在 `desktop/` 目录运行：

```powershell
npm run dist:portable
```

生成 `desktop/release/EasyExcel-Portable-0.4.0-x64.zip`（版本号随 `desktop/package.json` 更新）。完整解压后双击 `EasyExcel.exe`，不要只复制 EXE。

`npm run dist:win` 可生成安装程序；`npm run build` 只编译界面。详细说明见 [桌面版文档](desktop/README.md)。构建产物和依赖目录不提交到 Git。

## 离线 Python 与 AI 编程

已内置 pandas、NumPy、Matplotlib、openpyxl、xlrd、numpy-financial、PyXIRR、SciPy、statsmodels、Seaborn、XlsxWriter、Babel、QuantLib，以及 Python 内置 decimal。

将 [AI 编程说明](desktop/EasyExcel-AI编程说明.txt) 交给 AI，填写字段和需求，即可按工作台的 `df`、`tables`、`OUTPUT_DIR`、`display()`、`save_excel()`、`save_figure()` 等接口生成代码。软件中选择好 Sheet 后，还可通过“给 AI 的说明”导出包含当前表头的说明。

脚本以当前用户权限运行；结果应写到本次 `OUTPUT_DIR`。公式读取已保存的缓存值，不执行 Excel 公式重算。自定义脚本导出不会自动保留原表样式，具体行为见桌面版文档。

更名后继续使用 `%APPDATA%/excelrouter-desktop` 保存脚本库、历史和草稿，兼容旧版备份。换电脑时使用脚本库的“备份 / 导入备份”。

## 验证

```powershell
# 项目根目录
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m ruff check .
.\.venv\Scripts\python.exe -m pytest tests -q

# desktop 目录
npm run build
npm run test:library
npm run test:smoke
# 先生成免安装包，再检查内置引擎
npm run test:packaged
```

测试使用临时数据，不要将真实业务表格或运行结果提交到仓库。上游附带的演示工作簿不随 EasyExcel 源码提交；可使用自己的表格，或通过桌面端测试脚本生成临时样例。

## 项目结构与来源

- `desktop/`：EasyExcel 的 Electron 界面、Python 桥接、离线脚本运行器与打包工具。
- `core/`：Excel 核心处理逻辑。
- `tests/`：核心与桌面后端验证。
- `gui/`、根目录 `main.py` 和原有构建脚本：保留的上游 Python 界面。

本项目基于 [AbeLin / ExcelRouter](https://github.com/MarsandSea/ExcelRouter) 扩展，保留原项目版权声明与 [MIT 许可证](LICENSE)。[原项目说明](README.upstream.md) 保留供参考；其中下载链接和旧版说明属于上游项目。

原有 GitHub Actions 包含上游的 Gitee 镜像、发行版和技能市场发布流程，已限制只在原仓库运行。EasyExcel 当前使用上述本地构建命令，不会通过这些旧流程向上游服务发布。
