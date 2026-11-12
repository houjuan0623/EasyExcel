# Copyright (c) 2026 AbeLin · MIT License
"""规则按 Excel 地址处理分析副本；失败位置、计算顺序、保存和导出均可追溯。"""
import copy
import hashlib
import json
import re
from pathlib import Path

import openpyxl
import pytest

from desktop import analysis
from desktop.analysis_rules import apply_rules


def rule(target="B:B", operation="multiply", value="2", **changes):
    return {"id": "rule", "enabled": True, "target": target, "operation": operation, "operand": "constant",
            "value": value, "find": "", "includeHeaders": False, "invalid": "skip", **changes}


def process(rows, rules, **settings):
    data = copy.deepcopy(rows)
    spec = {"name": "明细", "orientation": "columns", "includeHidden": False, "rules": rules, **settings}
    audits, warnings = apply_rules(data, spec, (1, 1, len(rows[0]), len(rows)),
                                  {"hiddenRows": settings.get("hiddenRows", set()), "hiddenCols": settings.get("hiddenCols", set())},
                                  settings.get("header", 1), analysis.numeric, analysis.plain)
    return data, audits, warnings


def test_default_and_disabled_rules_keep_values():
    rows = [["项目", "金额"], ["甲", 10], ["乙", None]]
    assert process(rows, [])[0] == rows
    assert process(rows, [rule(enabled=False, value="bad")])[0] == rows


def test_cell_column_row_and_rectangular_operations():
    rows = [["项目", "预算", "实际"], ["甲", 100, 40], ["乙", 200, 90]]
    steps = [rule("B:B", "subtract", "C", operand="column"),
             rule("2:2", "multiply", "2", id="row"), rule("B2", "add", "5", id="cell"),
             rule("B2:C3", "divide", "10", id="area")]
    data, audit, warnings = process(rows, steps)
    assert data == [rows[0], ["甲", 12.5, 8], ["乙", 11, 9]]
    assert audit[1]["skipped"] == 1 and warnings
    assert audit[2]["examples"][0] == {"cell": "B2", "before": 120.0, "after": 125.0}
    assert rows[1][1] == 100


def test_reference_reads_snapshot_and_later_rule_sees_previous_result():
    rows = [["项目", "数值"], ["甲", 10], ["乙", 20]]
    data, _, _ = process(rows, [rule("B:B", "divide", "B2", operand="cell"), rule("B3", "add", "B2", operand="cell", id="later")])
    assert data[1][1] == 1 and data[2][1] == 3
    data, _, _ = process([["名称", 10, 20], ["甲", 30, 80]], [rule("2:2", "subtract", "1", operand="row")])
    assert data[1] == ["甲", 20, 60]


def test_copy_calculated_column_text_edits_and_header_protection():
    rows = [["部门", "预算", "实际", None], ["销售", 100, 70, None], ["销售", 200, 120, None]]
    steps = [rule("D:D", "copy", "B", operand="column"), rule("D:D", "subtract", "C", operand="column", id="difference"),
             rule("D1", "set_text", "差额", includeHeaders=True, id="header"),
             rule("A:A", "replace", "销售中心", find="销售", id="rename")]
    data, _, _ = process(rows, steps)
    assert data == [["部门", "预算", "实际", "差额"], ["销售中心", 100, 70, 30], ["销售中心", 200, 120, 80]]
    with pytest.raises(ValueError, match="表头"):
        process(rows, [rule("A1", "set_text", "部门名称")])
    assert process(rows, [rule("A2", "set_text", "=SUM(B2:C2)")])[0][1][0] == "=SUM(B2:C2)"


def test_text_operations_numeric_set_and_clear_are_explicit():
    rows = [["名称", "金额"], ["甲", 10], ["乙", None]]
    data, _, _ = process(rows, [rule("A:A", "prefix", "组-"), rule("A:A", "suffix", "号", id="suffix"),
                              rule("B3", "set_number", "12.5%", id="number"), rule("B2", "clear", "", id="clear")])
    assert data == [rows[0], ["组-甲号", None], ["组-乙号", 0.125]]


def test_hidden_targets_excluded_and_hidden_references_rejected():
    rows = [["部门", "金额"], ["甲", 10], ["乙", 20]]
    assert process(rows, [rule()], hiddenRows={3})[0][2][1] == 20
    with pytest.raises(ValueError, match="B2.*隐藏"):
        process(rows, [rule("B2", "divide", "B3", operand="cell")], hiddenRows={3})


def test_transposed_headers_use_original_coordinates():
    rows = [["项目", "甲", "乙"], ["预算", 100, 200], ["实际", 40, 90]]
    data, _, _ = process(rows, [rule("2:2", "subtract", "3", operand="row")], orientation="rows")
    assert data[1] == ["预算", 60, 110]


@pytest.mark.parametrize("r,match", [
    (rule("B2", "divide", "0"), "除数不能为 0"),
    (rule("B2", "divide", "C2", operand="cell"), "B2.*除数为 0"),
    (rule("B3", "add", "1", invalid="stop"), "B3.*没有有效数字"),
    (rule("Z:Z"), "超出"), (rule("B2;B3"), "作用区域"),
    (rule(value="__import__('os').system('whoami')"), "有限数字"),
    (rule(value="1e309"), "有限数字"), (rule(operation="execute"), "不支持"),
    (rule(operation="replace", find=""), "查找"),
    (rule("B2", "add", "C9", operand="cell"), "超出"),
])
def test_rule_failures_identify_rule_and_cell(r, match):
    rows = [["项目", "数字", "引用"], ["甲", 10, 0], ["乙", None, 2]]
    with pytest.raises(ValueError, match="规则 1") as exc:
        process(rows, [r])
    assert re.search(match, str(exc.value))


def test_order_can_be_changed_and_unchanged_rules_are_reported():
    rows = [["项", "数"], ["甲", 10]]
    add, mul = rule("B2", "add", "2"), rule("B2", "multiply", "3", id="mul")
    assert process(rows, [add, mul])[0][1][1] == 36
    assert process(rows, [mul, add])[0][1][1] == 32
    assert "没有改变" in process(rows, [rule(value="1")])[2][0]


def test_rules_participate_in_schema_freshness_render_and_report(tmp_path):
    path = tmp_path / "rules.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "明细"
    for row in [["项目", "预算", "实际", "差额"], ["甲", 100, 60, None], ["乙", 200, 90, None]]:
        ws.append(row)
    wb.save(path); wb.close()
    original = hashlib.sha256(path.read_bytes()).hexdigest()
    block = {"id": "b", "name": "明细", "path": str(path), "sheet": "明细", "range": "A1:D3",
             "headerRows": 1, "orientation": "columns", "includeHidden": False, "skipTotals": True,
             "rules": [rule("D:D", "copy", "B", operand="column"), rule("D:D", "subtract", "C", operand="column", id="sub")]}
    req = {"inputPath": str(path), "blocks": [block]}
    schema = analysis.inspect(req)
    assert schema["blocks"][0]["sample"][0][-1] == 40
    config = {**req, "token": schema["token"], "title": "差额合计", "kind": "metric", "x": "", "y": ["b:f3"], "unit": "元", "note": "规则示例", "aggregation": "sum", "combine": "single"}
    result = analysis.render(config)
    assert result["metric"] == 150 and len(result["ruleAudit"][0]["rules"]) == 2
    output = Path(analysis.export_dashboard({"title": "规则看板", "cards": [{"id": "card", "config": config}], "outputPath": str(tmp_path)})["output"])
    saved = json.loads((output / "dashboard-plan.json").read_text(encoding="utf-8"))
    assert saved["cards"][0]["config"]["blocks"][0]["rules"] == block["rules"]
    html = (output / "dashboard.html").read_text(encoding="utf-8")
    assert "处理前" in html and "选择与生成步骤" in html
    assert "ruleAudit" not in (output / "dashboard-plan.json").read_text(encoding="utf-8")
    block["rules"][1]["enabled"] = False
    with pytest.raises(ValueError, match="已改变"):
        analysis.render(config)
    assert hashlib.sha256(path.read_bytes()).hexdigest() == original
