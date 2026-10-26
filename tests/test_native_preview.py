"""Excel 原样渲染的边界、只读副本和真实坐标回归。

Copyright (c) 2026 AbeLin · MIT License
"""

import base64
from hashlib import sha256
from io import BytesIO
import os
from pathlib import Path
import sys
from zipfile import ZipFile

import openpyxl
from PIL import Image
import pytest

from desktop.native_preview import cell_details, expanded_bounds, preview, preview_copy
from desktop.scripts.native_preview_fixture import create


def test_copy_keeps_styles_charts_formulas_and_caches(tmp_path):
    source = create(tmp_path)
    before = sha256(source.read_bytes()).hexdigest()
    target = tmp_path / "copy.xlsx"
    preview_copy(source, target)
    with ZipFile(source) as a, ZipFile(target) as b:
        assert a.namelist() == b.namelist()
        for name in a.namelist():
            if name != "xl/workbook.xml":
                assert a.read(name) == b.read(name)
    wb = openpyxl.load_workbook(target, read_only=True, data_only=False)
    try:
        assert wb.calculation.calcMode == "manual"
        assert wb.calculation.fullCalcOnLoad is False
        assert wb["原样测试"]["C121"].value == "=B121*2"
    finally:
        wb.close()
    assert sha256(source.read_bytes()).hexdigest() == before


@pytest.mark.parametrize(
    "part", ["xl/connections.xml", "xl/macrosheets/sheet1.xml", "xl/activeX/activeX1.xml", "xl/embeddings/object.bin"]
)
def test_active_parts_rejected_before_excel_starts(tmp_path, part):
    source = create(tmp_path)
    with ZipFile(source, "a") as z:
        z.writestr(part, "<data/>")
    with pytest.raises(ValueError, match="活动对象"):
        preview_copy(source, tmp_path / "copy.xlsx")


@pytest.mark.parametrize(
    "formula", ['=WEBSERVICE("https://example.test")', '=RTD("server",,"x")', '=_xlfn.IMAGE("https://example.test/a")']
)
def test_network_formula_rejected_before_excel_starts(tmp_path, formula):
    source = tmp_path / "network.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = formula
    wb.save(source)
    with pytest.raises(ValueError, match="联网"):
        preview_copy(source, tmp_path / "copy.xlsx")


def test_page_expands_merged_cells_and_keeps_empty_chart_area():
    assert expanded_bounds(1, 1, 11, 6, []) == (1, 1, 20, 40)
    assert expanded_bounds(41, 1, 130, 20, [(5, 39, 6, 43)]) == (1, 39, 20, 80)
    assert expanded_bounds(1, 21, 130, 28, [(20, 3, 22, 4)]) == (20, 1, 40, 40)
    with pytest.raises(ValueError, match="合并区域过大"):
        expanded_bounds(1, 1, 300, 100, [(1, 1, 100, 300)])


def test_exact_formula_and_formatted_cached_value_after_row_100(tmp_path):
    path = create(tmp_path)
    meta = {"hiddenRows": {7}, "hiddenCols": {4}}
    cells = cell_details(path, "原样测试", (1, 101, 28, 130), False, meta)
    assert cells["C121"] == {"formula": "=B121*2", "text": "24,200.00", "missingCache": False}
    assert "D121" not in cells
    assert cells["AB121"]["text"] == "第 28 列"
    assert "D121" in cell_details(path, "原样测试", (1, 121, 4, 121), True, meta)


@pytest.mark.skipif(
    sys.platform != "win32" or os.environ.get("EASYEXCEL_TEST_NATIVE") != "1",
    reason="需要本机 Microsoft Excel，显式启用真实渲染测试",
)
def test_excel_image_coordinates_formulas_and_cleanup(tmp_path):
    import gc
    import win32api
    import win32com.client
    import win32con
    import win32event
    import win32process

    path = create(tmp_path)
    before = sha256(path.read_bytes()).hexdigest()
    # 保留一个独立的“用户 Excel”实例，验证预览不接管、不关闭其他实例。
    user_excel = win32com.client.DispatchEx("Excel.Application")
    user_excel.Visible = False
    user_excel.DisplayAlerts = False
    user_book = user_excel.Workbooks.Add()
    user_excel.Workbooks.Add()  # 保留空白工作簿，释放 COM 前不触发 Excel 的“最后一本关闭”自动退出。
    user_process = win32api.OpenProcess(
        win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE,
        False,
        win32process.GetWindowThreadProcessId(user_excel.Hwnd)[1],
    )
    user_book.Worksheets(1).Range("A1").Value = "user-session-kept"
    try:
        for top, left, formulas in ((1, 1, False), (41, 1, False), (101, 21, False), (1, 1, True)):
            result = preview(
                {
                    "inputPath": str(path),
                    "path": str(path),
                    "sheet": "原样测试",
                    "startRow": top,
                    "startColumn": left,
                    "showFormulas": formulas,
                }
            )
            assert result["engine"] == "excel"
            image = Image.open(BytesIO(base64.b64decode(result["image"].split(",")[1])))
            image.save(tmp_path / f"native-{top}-{left}-{formulas}.png")
            assert abs(sum(c["size"] for c in result["columns"]) - (image.width - 1)) < 0.1
            assert abs(sum(r["size"] for r in result["rows"]) - (image.height - 1)) < 0.1
            if top == 1:
                pixels = {color: count for count, color in image.convert("RGB").getcolors(image.width * image.height)}
                assert pixels.get((22, 120, 90), 0) > 100
                assert pixels.get((255, 0, 238), 0) > 100  # Excel 原生图表的系列颜色。
                assert result["cells"]["C4"]["formula"] == "=B4*2"
                assert result["columns"][3]["size"] == 0
            if top == 41:
                assert result["startRow"] == 39
                pixels = {color: count for count, color in image.convert("RGB").getcolors(image.width * image.height)}
                assert pixels.get((255, 0, 238), 0) > 100
            if left == 21:
                assert result["cells"]["AB121"]["text"] == "第 28 列"
            image.close()
            assert user_book.Worksheets(1).Range("A1").Value == "user-session-kept"
            assert win32process.GetWindowThreadProcessId(user_excel.Hwnd)[1] > 0
        assert sha256(path.read_bytes()).hexdigest() == before
    finally:
        user_book.Close(False)
        user_book = user_excel = None
        gc.collect()
        if win32event.WaitForSingleObject(user_process, 1000) == win32event.WAIT_TIMEOUT:
            win32api.TerminateProcess(user_process, 0)
        user_process.Close()
