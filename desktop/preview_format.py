"""将常见 Excel 数字格式转换为预览文字，不计算工作簿公式。

Copyright (c) 2026 AbeLin · MIT License
"""
import calendar
import datetime as dt
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import math
import re

from openpyxl.styles.numbers import is_date_format


def plain(value):
    """缺少受支持的格式时仍保留完整数值，不伪造计算结果。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, dt.datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, (dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, float) and math.isfinite(value) and value.is_integer():
        return str(int(value))
    return str(value)


def date_text(value, pattern):
    """解释日期 / 时间占位符，区分月份 m 与分钟 m。"""
    if isinstance(value, dt.timedelta):
        seconds = int(value.total_seconds())
        date = dt.datetime(1900, 1, 1) + value
    else:
        seconds = getattr(value, "hour", 0) * 3600 + getattr(value, "minute", 0) * 60 + getattr(value, "second", 0)
        date = value
    tokens = re.findall(r'"[^"]*"|\\.|\[[^]]+\]|AM/PM|A/P|[ymdhs]+|.', pattern, flags=re.I)
    timed = any(token.upper() in ("AM/PM", "A/P") for token in tokens)
    fields = [(index, token.lower()[0]) for index, token in enumerate(tokens) if re.fullmatch(r"[ymdhs]+", token, re.I)]
    result = []
    for index, token in enumerate(tokens):
        low = token.lower()
        n = len(token)
        if token.startswith('"') or token.startswith("\\"):
            result.append(token[1:-1] if token.startswith('"') else token[1:])
        elif low in ("[h]", "[hh]", "[m]", "[mm]", "[s]", "[ss]"):
            divisor = 3600 if "h" in low else 60 if "m" in low else 1
            result.append(str(seconds // divisor).zfill(len(low) - 2))
        elif token.startswith("["):
            continue
        elif low in ("am/pm", "a/p"):
            result.append(("AM" if getattr(date, "hour", 0) < 12 else "PM")[:1 if low == "a/p" else 2])
        elif low.startswith("y"):
            year = getattr(date, "year", 1900)
            result.append(str(year) if n > 2 else f"{year % 100:02d}")
        elif low.startswith("m"):
            previous = next((field for pos, field in reversed(fields) if pos < index), "")
            following = next((field for pos, field in fields if pos > index), "")
            minute = previous == "h" or following == "s"
            number = getattr(date, "minute" if minute else "month", 0 if minute else 1)
            if not minute and n >= 3:
                result.append(calendar.month_abbr[number] if n == 3 else calendar.month_name[number][:1] if n >= 5 else calendar.month_name[number])
            else:
                result.append(str(number).zfill(2 if n >= 2 else 1))
        elif low.startswith("d"):
            if n >= 3 and hasattr(date, "weekday"):
                result.append((calendar.day_abbr if n == 3 else calendar.day_name)[date.weekday()])
            else:
                result.append(str(getattr(date, "day", 1)).zfill(2 if n >= 2 else 1))
        elif low.startswith("h"):
            hour = getattr(date, "hour", 0)
            result.append(str((hour % 12 or 12) if timed else hour).zfill(2 if n >= 2 else 1))
        elif low.startswith("s"):
            result.append(str(getattr(date, "second", 0)).zfill(2 if n >= 2 else 1))
        else:
            result.append(token)
    return "".join(result)


def display(value, number_format="General"):
    """返回显示文字；第二项标记不支持的特殊格式，供界面提示。"""
    fmt = number_format or "General"
    if value is None or isinstance(value, (str, bool)):
        return plain(value), False
    if isinstance(value, (dt.date, dt.time, dt.timedelta)):
        return (date_text(value, fmt.split(";")[0]), False) if is_date_format(fmt) else (plain(value), False)
    if not isinstance(value, (int, float)) or not math.isfinite(value) or fmt.lower() in ("general", "@"):
        return plain(value), False
    sections = re.split(r';(?=(?:[^"]*"[^"]*")*[^"]*$)', fmt)
    negative_section = value < 0 and len(sections) > 1
    pattern = sections[1] if negative_section else sections[2] if value == 0 and len(sections) > 2 else sections[0]
    # 条件、分数、分段编号等格式保留原值，避免近似解释为错误的数值。
    if re.search(r"\[(?:[<>]=?|=)", pattern) or re.search(r"[?0#]/[?0#]", pattern):
        return plain(value), True
    literals = []

    def protect(match):
        token = match.group()
        literals.append(token[1:-1] if token.startswith('"') else token[1:])
        return chr(0xE000 + len(literals) - 1)

    pattern = re.sub(r'"[^"]*"|\\.', protect, pattern)
    pattern = re.sub(r"\[\$([^\]-]*)[^\]]*\]", r"\1", pattern)
    pattern = re.sub(r"\[[^]]*\]", "", pattern)
    pattern = re.sub(r"_.", " ", pattern)
    pattern = re.sub(r"\*.", "", pattern)

    def restore(text):
        for index, literal in enumerate(literals):
            text = text.replace(chr(0xE000 + index), literal)
        return text

    matches = list(re.finditer(r"[0#?][0#?,]*(?:\.[0#?]+)?(?:[Ee][+-]0+)?", pattern))
    if not matches:
        return restore(pattern), False
    if len(matches) != 1:
        return plain(value), True
    match = matches[0]
    mask = match.group()
    prefix, suffix = pattern[:match.start()], pattern[match.end():]
    trailing_commas = len(mask) - len(mask.rstrip(","))
    scale_suffix = re.match(r",+", suffix)
    if scale_suffix:
        trailing_commas += len(scale_suffix.group())
        suffix = suffix[len(scale_suffix.group()):]
    mask = mask.rstrip(",")
    try:
        number = Decimal(str(abs(value) if negative_section else value)) * (100 ** pattern.count("%")) / (1000 ** trailing_commas)
        decimals = len(mask.split(".", 1)[1].split("E")[0].split("e")[0]) if "." in mask else 0
        decimals = min(decimals, 15)
        if "e" in mask.lower():
            text = f"{number:.{decimals}E}"
            head, exponent = text.split("E")
            digits = len(re.split(r"[Ee][+-]", mask)[-1])
            text = head + "E" + ("+" if int(exponent) >= 0 else "-") + str(abs(int(exponent))).zfill(digits)
        else:
            number = number.quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)
            text = f"{number:,.{decimals}f}" if "," in mask else f"{number:.{decimals}f}"
            integer_mask = mask.split(".")[0]
            if "," not in mask and integer_mask.count("0") > 1:
                sign = "-" if text.startswith("-") else ""
                text = sign + text.lstrip("-").zfill(integer_mask.count("0") + (decimals + 1 if decimals else 0))
            if decimals and mask.endswith("#"):
                optional = len(mask) - len(mask.rstrip("#"))
                for _ in range(optional):
                    if text.endswith("0"):
                        text = text[:-1]
                text = text.rstrip(".")
        return restore(prefix + text + suffix).strip(), False
    except (InvalidOperation, ValueError, OverflowError):
        return plain(value), True
