# Copyright (c) 2026 AbeLin · MIT License
"""分析副本的顺序处理规则。按 Excel 原始坐标定位，不执行代码、不写回工作簿。"""
import math
import re
from openpyxl.utils.cell import column_index_from_string, get_column_letter, range_boundaries

OPERATIONS = {"add": "加上", "subtract": "减去", "multiply": "乘以", "divide": "除以",
              "set_number": "设置数字", "set_text": "设置文本", "replace": "替换文字", "prefix": "添加前缀",
              "suffix": "添加后缀", "clear": "清空", "copy": "复制另一处的值"}
ARITHMETIC = {"add", "subtract", "multiply", "divide"}


def target_bounds(value, bounds):
    """整行整列限定在分析区域内，矩形或单元格必须全部位于当前区域。"""
    if not isinstance(value, str) or len(value) > 80:
        raise ValueError("作用区域格式无效")
    area = value.strip().upper().replace("$", "")
    left, top, right, bottom = bounds
    if re.fullmatch(r"[A-Z]{1,3}:[A-Z]{1,3}", area):
        a, c = map(column_index_from_string, area.split(":"))
        b, d = top, bottom
    elif re.fullmatch(r"[1-9][0-9]{0,5}:[1-9][0-9]{0,5}", area):
        b, d = map(int, area.split(":"))
        a, c = left, right
    elif re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]{0,5}(?::[A-Z]{1,3}[1-9][0-9]{0,5})?", area):
        a, b, c, d = range_boundaries(area)
    else:
        raise ValueError("作用区域应为 D:D、7:7、D7 或 D7:F20")
    if not (left <= a <= c <= right and top <= b <= d <= bottom):
        raise ValueError("作用区域超出本数据块，请调整规则区域或扩大数据块选区")
    return a, b, c, d


def apply_rules(rows, spec, bounds, meta, header, numeric, plain):
    """每条规则读取开始时的快照，后续规则读取上一步结果；默认跳过非数值并留痕。"""
    rules = spec.get("rules", [])
    if not isinstance(rules, list) or len(rules) > 20:
        raise ValueError("每个数据块最多设置 20 条处理规则")
    left, top, right, bottom = bounds
    audit, warnings, identifiers = [], [], set()
    for index, rule in enumerate(rules, 1):
        context = f"「{spec['name']}」规则 {index}"
        try:
            if not isinstance(rule, dict):
                raise ValueError("规则格式无效")
            identifier = rule.get("id")
            if (not isinstance(identifier, str) or not re.fullmatch(r"[a-zA-Z0-9_-]{1,80}", identifier)
                    or identifier in identifiers or identifier in {"__proto__", "prototype", "constructor"}):
                raise ValueError("规则标识无效或重复")
            identifiers.add(identifier)
            if type(rule.get("enabled")) is not bool or type(rule.get("includeHeaders")) is not bool:
                raise ValueError("规则开关格式无效")
            operation, operand, value = rule.get("operation"), rule.get("operand"), rule.get("value")
            find, invalid = rule.get("find"), rule.get("invalid")
            if (operation not in OPERATIONS or operand not in ("constant", "cell", "column", "row")
                    or invalid not in ("skip", "stop")):
                raise ValueError("规则包含不支持的操作")
            if any(not isinstance(v, str) or len(v) > 1000 for v in (value, find)):
                raise ValueError("规则文字最多 1000 个字符")
            if not rule["enabled"]:
                continue
            a, b, c, d = target_bounds(rule.get("target"), bounds)
            if operation == "replace" and not find:
                raise ValueError("替换文字需要填写要查找的内容")
            constant, reference = None, None
            if operation in ARITHMETIC or operation in ("set_number", "copy"):
                if operation == "set_number" and operand != "constant":
                    raise ValueError("设置数字仅接受固定数值")
                if operation == "copy" and operand == "constant":
                    raise ValueError("复制值需要选择固定单元格、同一行的另一列或同一列的另一行")
                if operand == "constant":
                    constant = numeric(value)
                    if not math.isfinite(constant):
                        raise ValueError("请填写有限数字，例如 10000、0.1 或 10%")
                    if operation == "divide" and constant == 0:
                        raise ValueError("除数不能为 0，请修改固定数值")
                elif operand == "cell":
                    if not re.fullmatch(r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,5}", value.strip()):
                        raise ValueError("固定单元格应填写 D7 这样的地址")
                    x, y, _, _ = target_bounds(value, bounds)
                    reference = x, y
                elif operand == "column":
                    if not re.fullmatch(r"[A-Za-z]{1,3}", value.strip()):
                        raise ValueError("同一行的另一列应填写 E 这样的列标")
                    x = column_index_from_string(value.strip().upper())
                    if not left <= x <= right:
                        raise ValueError("引用列超出本数据块选区")
                    reference = x, None
                else:
                    if not re.fullmatch(r"[1-9][0-9]{0,5}", value.strip()) or not top <= int(value) <= bottom:
                        raise ValueError("引用行须位于本数据块选区内，例如 8")
                    reference = None, int(value)
            previous = [r[:] for r in rows]
            record = {"rule": index, "target": rule["target"], "description": OPERATIONS[operation],
                      "matched": 0, "changed": 0, "skipped": 0, "examples": []}
            for r in range(b, d + 1):
                for col in range(a, c + 1):
                    is_header = r < top + header if spec.get("orientation", "columns") == "columns" else col < left + header
                    if is_header and not rule["includeHeaders"]:
                        continue
                    if not spec.get("includeHidden") and (r in meta["hiddenRows"] or col in meta["hiddenCols"]):
                        continue
                    address = f"{get_column_letter(col)}{r}"
                    old = previous[r - top][col - left]
                    record["matched"] += 1
                    source_value = None
                    if reference:
                        x, y = reference[0] or col, reference[1] or r
                        if not spec.get("includeHidden") and (y in meta["hiddenRows"] or x in meta["hiddenCols"]):
                            raise ValueError(f"{address} 引用了隐藏单元格，请开启包含隐藏行列或更换引用")
                        source_value = previous[y - top][x - left]
                    if operation == "copy":
                        new = source_value
                    elif operation in ARITHMETIC:
                        n, other = numeric(old), constant
                        if reference:
                            other = numeric(source_value)
                        if not math.isfinite(n) or not math.isfinite(other):
                            if invalid == "stop":
                                raise ValueError(f"{address} 或其引用没有有效数字（可能为空、文本或公式无缓存）；修改数据或选择保持原样并提示")
                            record["skipped"] += 1
                            continue
                        if operation == "divide" and other == 0:
                            raise ValueError(f"{address} 的除数为 0，请修改引用或先设置正确的除数")
                        new = n + other if operation == "add" else n - other if operation == "subtract" else n * other if operation == "multiply" else n / other
                        if not math.isfinite(new):
                            raise ValueError(f"{address} 的计算结果超出可表示范围，请降低运算数值")
                    elif operation == "set_number":
                        new = constant
                    elif operation == "set_text":
                        new = value
                    elif operation == "clear":
                        new = None
                    else:
                        if not isinstance(old, str):
                            if invalid == "stop":
                                raise ValueError(f"{address} 不是文字；请改用设置文本或选择保持原样并提示")
                            record["skipped"] += 1
                            continue
                        new = old.replace(find, value) if operation == "replace" else value + old if operation == "prefix" else old + value
                        if len(new) > 32767:
                            raise ValueError(f"{address} 的文字超过 32767 字符，请缩短内容")
                    rows[r - top][col - left] = new
                    if new != old or isinstance(new, bool) != isinstance(old, bool):
                        record["changed"] += 1
                        if len(record["examples"]) < 6:
                            record["examples"].append({"cell": address, "before": plain(old), "after": plain(new)})
            if not record["matched"]:
                raise ValueError("没有可处理的单元格；检查作用区域、表头开关和隐藏行列设置")
            if record["skipped"]:
                warnings.append(f"{context}：{record['skipped']} 个不符合类型的单元格保持原样，未当作零参与运算")
            if not record["changed"]:
                warnings.append(f"{context}：没有改变任何值，请查看规则样例和查找条件")
            audit.append(record)
        except (ValueError, TypeError, OverflowError) as exc:
            raise ValueError(f"{context}：{exc}") from exc
    return audit, warnings
