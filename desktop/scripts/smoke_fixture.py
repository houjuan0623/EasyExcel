"""桌面端到端测试样本，只写测试指定的临时目录。

Copyright (c) 2026 AbeLin · MIT License
"""
import datetime
import json
from pathlib import Path
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


def create(directory):
    """构造多 Sheet 的合成数据，不使用用户业务数据。"""
    path = directory / "2026年销售明细.xlsx"
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    departments = ["华东区域", "华南区域", "华北区域", "西南区域"]
    for index, name in enumerate(["10月销售明细", "9月销售明细", "8月销售明细", "归档数据"]):
        ws = wb.create_sheet(name)
        ws.append(["订单编号", "区域", "负责人", "产品名称", "销售金额", "日期"])
        for n in range(24):
            ws.append([f"SO-{index + 8:02d}{n + 1:04d}", departments[n % 4], f"负责人{n % 5 + 1}",
                       ["办公设备", "企业服务", "数据产品"][n % 3], 1280 + n * 350,
                       datetime.datetime(2026, 10 - index, n + 1)])
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="21845E")
            cell.font = Font(color="FFFFFF", bold=True)
        if name == "归档数据":
            ws.sheet_state = "hidden"
    ws = wb.create_sheet("原样预览")
    ws.merge_cells("A1:F2")
    ws["A1"] = "十月销售总览"
    ws["A1"].font = Font(name="Microsoft YaHei", size=20, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="155D45")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 22
    for label, width in zip("ABCDEF", [22, 17, 20, 16, 14, 34]):
        ws.column_dimensions[label].width = width
    for index, label in enumerate(["订单编号", "日期", "销售金额", "区域", "完成率", "备注"], 1):
        cell = ws.cell(4, index, label)
        cell.font = Font(name="Microsoft YaHei", size=11, bold=True, color="155D45")
        cell.fill = PatternFill("solid", fgColor="DFEEE7")
        cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[4].height = 27
    edge = Side(style="thin", color="D4DED8")
    for row in range(5, 151):
        values = [f"ORDER-{row:04d}", datetime.datetime(2026, 10, (row % 28) + 1), 1200 + row * 125.5,
                  departments[row % 4], 0.72 + row / 1000, "预览保留原表样式"]
        for col, value in enumerate(values, 1):
            cell = ws.cell(row, col, value)
            cell.font = Font(name="Microsoft YaHei", size=10, color="243C31")
            cell.fill = PatternFill("solid", fgColor="F2F7F4" if row % 2 else "FFFFFF")
            cell.border = Border(left=edge, right=edge, top=edge, bottom=edge)
            cell.alignment = Alignment(vertical="center")
        ws.row_dimensions[row].height = 24
        ws.cell(row, 2).number_format = "yyyy/mm/dd"
        ws.cell(row, 3).number_format = '"¥"#,##0.00'
        ws.cell(row, 5).number_format = "0.0%"
    ws["F5"] = "文字自动换行\n第二行保持原样"
    ws["F5"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[5].height = 38
    ws.row_dimensions[8].hidden = True
    ws.column_dimensions["G"].hidden = True
    ws["G5"] = "隐藏字段"
    ws.merge_cells("A99:B105")
    ws["A99"] = "跨越第 100 行的合并区域"
    ws["F100"] = "第 100 行，完整保留"
    ws["F101"] = "第 101 行，不应显示"
    wb.save(path)
    wb.close()
    print(json.dumps({"path": str(path)}, ensure_ascii=False))


def verify(directory):
    """检查仅所选月份导出，内容、表头样式均由 Python 保存。"""
    files = list((directory / "拆分结果").rglob("*.xlsx"))
    bundles = [path for path in files if path.parent.name == '总工作簿']
    files = [path for path in files if path.parent.name != '总工作簿']
    assert len(files) == 4, len(files)
    assert len(bundles) == 1, len(bundles)
    total = 0
    for path in files:
        wb = openpyxl.load_workbook(path)
        assert wb.sheetnames == ["9月销售明细"], wb.sheetnames
        assert wb.active["A1"].fill.fgColor.rgb == "0021845E"
        total += wb.active.max_row - 1
        wb.close()
    assert total == 24, total
    bundled = openpyxl.load_workbook(bundles[0])
    assert set(bundled.sheetnames) == {'华东区域', '华南区域', '华北区域', '西南区域'}
    assert sum(ws.max_row - 1 for ws in bundled) == total
    for ws in bundled:
        assert ws['A1'].fill.fgColor.rgb == '0021845E'
    bundled.close()
    print(json.dumps({"files": len(files), "rows": total}))


if __name__ == "__main__":
    directory = Path(sys.argv[2])
    directory.mkdir(parents=True, exist_ok=True)
    (create if sys.argv[1] == "create" else verify)(directory)
