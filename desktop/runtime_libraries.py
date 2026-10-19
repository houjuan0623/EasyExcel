"""The documented, tested offline script environment (never imported by core).

Copyright (c) 2026 AbeLin · MIT License
"""
import sys


def environment():
    # Explicit imports also make PyInstaller collect the user-facing packages.
    import babel
    import matplotlib
    import numpy
    import numpy_financial
    import openpyxl
    import pandas
    import pyxirr
    import QuantLib
    import scipy
    import seaborn
    import statsmodels
    import xlrd
    import xlsxwriter

    packages = [
        ("pandas", pandas, "pandas", "对账、清洗、分组汇总与数据分析", "import pandas as pd"),
        ("numpy", numpy, "numpy", "数组运算与数值计算", "import numpy as np"),
        ("matplotlib", matplotlib, "matplotlib.pyplot", "财务趋势、构成与比较图表", "import matplotlib.pyplot as plt"),
        ("openpyxl", openpyxl, "openpyxl", "读取和修改 Excel 内容与格式", "import openpyxl"),
        ("xlrd", xlrd, "xlrd", "读取旧版 .xls 工作簿", "import xlrd"),
        ("numpy-financial", numpy_financial, "numpy_financial", "NPV、IRR、现值终值、贷款还款与利息", "import numpy_financial as npf"),
        ("pyxirr", pyxirr, "pyxirr", "按实际日期计算 XIRR、XNPV 等现金流指标", "import pyxirr"),
        ("scipy", scipy, "scipy", "统计检验、优化、拟合与方程求解", "from scipy import stats, optimize"),
        ("statsmodels", statsmodels, "statsmodels.api", "回归、时间序列与统计预测", "import statsmodels.api as sm"),
        ("seaborn", seaborn, "seaborn", "分布、相关性和分组统计图表", "import seaborn as sns"),
        ("XlsxWriter", xlsxwriter, "xlsxwriter", "输出含格式、公式和图表的 Excel 报表", "import xlsxwriter"),
        ("Babel", babel, "babel.numbers", "多币种金额及数字的本地化格式", "from babel.numbers import format_currency"),
        ("QuantLib", QuantLib, "QuantLib", "计息日历、债券、利率与金融估值", "import QuantLib as ql"),
    ]
    capabilities = [{"name": name, "version": module.__version__, "importName": import_name,
                     "purpose": purpose, "example": example}
                    for name, module, import_name, purpose, example in packages]
    capabilities.append({"name": "decimal", "version": "Python 内置", "importName": "decimal",
                         "purpose": "精确十进制金额、明确舍入规则", "example": "from decimal import Decimal, ROUND_HALF_UP"})
    return {"python": sys.version.split()[0],
            "libraries": {item["name"]: item["version"] for item in capabilities}, "capabilities": capabilities}
