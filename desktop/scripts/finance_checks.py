"""Executed inside the real code worker for source AND portable integration tests."""
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
import math
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import numpy_financial as npf
import pyxirr
import QuantLib as ql
from scipy import optimize, stats
import seaborn as sns
import statsmodels.api as sm
import statsmodels.formula.api as smf
from statsmodels.tsa.holtwinters import ExponentialSmoothing
from babel.numbers import format_currency

# Supplied by the workbench when this script is executed.
OUTPUT_DIR = Path(globals()["OUTPUT_DIR"])
display = globals()["display"]
save_figure = globals()["save_figure"]
df = globals()["df"]

assert Decimal("0.1") + Decimal("0.2") == Decimal("0.3")
assert Decimal("1.005").quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) == Decimal("1.01")
assert math.isclose(npf.npv(0.1, [-100, 110]), 0, abs_tol=1e-9)
assert math.isclose(npf.irr([-100, 110]), 0.1, abs_tol=1e-9)
assert math.isclose(npf.ipmt(0.01, 1, 12, 1200), -12, abs_tol=1e-9)
dates = [date(2025, 1, 1), date(2026, 1, 1)]
assert math.isclose(pyxirr.xirr(dates, [-100, 110]), 0.1, abs_tol=1e-8)
assert math.isclose(pyxirr.xnpv(0.1, dates, [-100, 110]), 0, abs_tol=1e-8)
assert math.isclose(stats.norm.cdf(0), 0.5)
assert math.isclose(optimize.brentq(lambda rate: -100 + 110 / (1 + rate), 0, 1), 0.1, abs_tol=1e-8)
linear = optimize.linprog([1, 2], bounds=[(1, None), (2, None)])
assert linear.success and math.isclose(linear.fun, 5)

sample = pd.DataFrame({"period": np.arange(12), "amount": 100 + 20 * np.arange(12)})
fit = sm.OLS(sample["amount"], sm.add_constant(sample["period"])).fit()
assert np.allclose(fit.params, [100, 20])
formula = smf.ols("amount ~ period", data=sample).fit()
assert np.allclose(formula.params, [100, 20])
forecast = ExponentialSmoothing(sample["amount"].astype(float), trend="add").fit().forecast(2)
assert len(forecast) == 2 and np.isfinite(forecast).all()

formatted = format_currency(Decimal("1234.50"), "CNY", locale="zh_CN")
assert "1,234.50" in formatted
start = ql.Date(1, 1, 2025)
end = ql.Date(1, 1, 2026)
ql.Settings.instance().evaluationDate = start
dc = ql.Actual365Fixed()
assert math.isclose(dc.yearFraction(start, end), 1)
curve = ql.FlatForward(start, 0.05, dc)
assert math.isclose(curve.discount(end), math.exp(-0.05))
bond = ql.ZeroCouponBond(0, ql.NullCalendar(), 100, end, ql.Unadjusted, 100, start)
bond.setPricingEngine(ql.DiscountingBondEngine(ql.YieldTermStructureHandle(curve)))
assert math.isclose(bond.NPV(), 100 * math.exp(-0.05))

with pd.ExcelWriter(OUTPUT_DIR / "财务库报表.xlsx", engine="xlsxwriter") as writer:
    sample.to_excel(writer, sheet_name="报表", index=False)
    workbook = writer.book
    worksheet = writer.sheets["报表"]
    worksheet.set_column("B:B", 18, workbook.add_format({"num_format": "#,##0.00"}))
    chart = workbook.add_chart({"type": "line"})
    chart.add_series({"values": ["报表", 1, 1, 12, 1]})
    worksheet.insert_chart("D2", chart)

figure, ax = plt.subplots(figsize=(7, 3.5))
sns.barplot(data=sample, x="period", y="amount", color="#428a68", errorbar=None, ax=ax)
ax.set_title("财务库验证")
save_figure(figure, "财务分析图.png")
display(pd.DataFrame({"检查项": ["现金流", "统计优化", "回归与预测", "货币格式", "金融估值", "报表图表"], "状态": ["通过"] * 6}), "财务库验证结果")
print("FINANCE_CHECKS_OK", formatted, len(df))
