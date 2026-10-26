"""原样选区测试用虚构工作簿，产物只写入调用方指定的仓库外目录。

Copyright (c) 2026 AbeLin · MIT License
"""

import json
from pathlib import Path
import sys

from PIL import Image
import xlsxwriter


def create(folder):
    """包括图表跨页、合并格、缓存公式、隐藏行列和超过第 100 行的数据。"""
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / "native-preview-fixture.xlsx"
    picture = folder / "test-icon.png"
    Image.new("RGB", (36, 36), "#ff8800").save(picture)
    with xlsxwriter.Workbook(path) as book:
        ws = book.add_worksheet("原样测试")
        title = book.add_format(
            {
                "bold": True,
                "font_size": 23,
                "font_color": "white",
                "bg_color": "#16785A",
                "align": "center",
                "valign": "vcenter",
            }
        )
        money = book.add_format({"num_format": "#,##0.00;[Red](#,##0.00)", "border": 1})
        ws.merge_range("A1:F2", "EasyExcel · 原样选区测试", title)
        ws.set_row(0, 27)
        ws.set_column("A:A", 20)
        ws.set_column("B:C", 16)
        ws.set_column("D:D", 12, None, {"hidden": True})
        ws.set_row(6, None, None, {"hidden": True})
        ws.write_row("A3", ["项目", "预算", "实际", "隐藏备注"])
        for r in range(3, 130):
            ws.write(r, 0, f"项目 {r - 2}")
            ws.write(r, 1, (r + 1) * 100, money)
            ws.write_formula(r, 2, f"=B{r + 1}*2", money, (r + 1) * 200)
            ws.write(r, 3, "HIDDEN_TEST_VALUE")
        ws.merge_range("E39:F43", "跨页合并", title)
        ws.write("AB121", "第 28 列")
        ws.insert_image("H17", str(picture))
        for anchor in ("F5", "F35"):
            chart = book.add_chart({"type": "column"})
            chart.add_series(
                {
                    "name": "预算",
                    "categories": "='原样测试'!$A$4:$A$10",
                    "values": "='原样测试'!$B$4:$B$10",
                    "fill": {"color": "#FF00EE"},
                }
            )
            chart.set_title({"name": "原始统计图"})
            chart.set_size({"width": 420, "height": 260})
            ws.insert_chart(anchor, chart)
        other = book.add_worksheet("另一个Sheet")
        other.write_row("A1", ["编号", "指标"])
        other.write_row("A2", ["A01", 10])
    return path


if __name__ == "__main__":
    print(json.dumps({"path": str(create(sys.argv[1]))}, ensure_ascii=False))
