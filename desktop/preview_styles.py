"""预览的复杂填充与条件格式；只解释有限规则，不执行工作簿代码。

Copyright (c) 2026 AbeLin · MIT License
"""
from copy import copy
import math
import operator
import re
from types import SimpleNamespace

from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import range_boundaries

from desktop.preview_format import display

OPS = {'>': operator.gt, '>=': operator.ge, '<': operator.lt, '<=': operator.le,
       '=': operator.eq, '<>': operator.ne, 'greaterThan': operator.gt, 'greaterThanOrEqual': operator.ge,
       'lessThan': operator.lt, 'lessThanOrEqual': operator.le, 'equal': operator.eq, 'notEqual': operator.ne}


def number(value):
    """布尔、空值和错误值不参加数值规则。"""
    return isinstance(value, (float, int)) and not isinstance(value, bool) and math.isfinite(value)


def fill_css(fill, palette, color):
    """将渐变和常见图案填充转成仅包含受控颜色与尺寸的 CSS。"""
    if getattr(fill, 'type', None) in ('linear', 'path') and getattr(fill, 'stop', None):
        stops = ', '.join(f'{color(s.color, palette, "#FFFFFF")} {s.position * 100:.2f}%' for s in fill.stop)
        if fill.type == 'path':
            return {'backgroundImage': f'radial-gradient(ellipse, {stops})'}
        return {'backgroundImage': f'linear-gradient({(90 + float(fill.degree or 0)) % 360}deg, {stops})'}
    pattern = getattr(fill, 'patternType', None)
    if pattern and pattern != 'solid':
        fg, bg = color(fill.fgColor, palette, '#000000'), color(fill.bgColor, palette, '#FFFFFF')
        angle = 0 if 'Horizontal' in pattern else 90 if 'Vertical' in pattern else 135 if 'Up' in pattern else 45
        spacing = 3 if pattern.startswith('dark') else 6
        gradient = f'repeating-linear-gradient({angle}deg, {fg} 0px, {fg} 1px, transparent 1px, transparent {spacing}px)'
        if 'Grid' in pattern or 'Trellis' in pattern:
            gradient += f', repeating-linear-gradient({angle + 90}deg, {fg} 0px, {fg} 1px, transparent 1px, transparent {spacing}px)'
        if pattern.startswith('gray'):
            gradient = f'radial-gradient({fg} 1px, transparent 1px)'
            return {'backgroundColor': bg, 'backgroundImage': gradient, 'backgroundSize': '4px 4px'}
        return {'backgroundColor': bg, 'backgroundImage': gradient}
    return {}


def format_color(value, fmt):
    """读取正/负/零对应格式段中的命名颜色。"""
    sections = re.split(r';(?=(?:[^"]*"[^"]*")*[^"]*$)', fmt or '')
    index = 1 if number(value) and value < 0 and len(sections) > 1 else 2 if value == 0 and len(sections) > 2 else 0
    match = re.search(r'\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow)\]', sections[index], re.I)
    colors = {'black': '#000000', 'blue': '#0000FF', 'cyan': '#00FFFF', 'green': '#008000',
              'magenta': '#FF00FF', 'red': '#FF0000', 'white': '#FFFFFF', 'yellow': '#FFFF00'}
    return colors.get(match[1].lower()) if match else None


def _threshold(item, values):
    """将色阶/数据条/图标阈值转换成真实数值。"""
    if not values:
        raise ValueError('条件格式范围没有可计算的数值')
    low, high = min(values), max(values)
    if item.type in ('min', 'autoMin'):
        return low
    if item.type in ('max', 'autoMax'):
        return high
    val = float(item.val or 0)
    if item.type == 'num':
        return val
    if item.type == 'percent':
        return low + (high - low) * val / 100
    if item.type == 'percentile':
        ordered = sorted(values)
        index = max(0, min(len(values) - 1, (len(values) - 1) * val / 100))
        lo, hi = math.floor(index), math.ceil(index)
        return ordered[lo] + (ordered[hi] - ordered[lo]) * (index - lo)
    raise ValueError('引用公式的条件格式阈值暂不支持')


def _blend(a, b, ratio):
    ratio = max(0, min(1, ratio))
    return '#' + ''.join(f'{round(int(a[i:i+2], 16) * (1-ratio) + int(b[i:i+2], 16) * ratio):02X}' for i in (1, 3, 5))


def _operand(text, current_sheet, reader):
    text = str(text).strip()
    if re.fullmatch(r'"(?:[^"]|"")*"', text):
        return text[1:-1].replace('""', '"')
    if text.upper() in ('TRUE', 'FALSE'):
        return text.upper() == 'TRUE'
    try:
        return float(text)
    except ValueError:
        values, hidden = reader.read(text, current_sheet)
        if len(values) != 1 or hidden[0]:
            raise ValueError('条件公式引用不可用') from None
        return values[0]


def _matches(rule, value, coordinate, anchor, sheet, reader):
    """支持比较和常见文字规则；复杂公式返回明确的不支持提示。"""
    if rule.type in ('containsText', 'notContainsText', 'beginsWith', 'endsWith'):
        text, needle = str(value or '').casefold(), str(rule.text or '').casefold()
        return {'containsText': needle in text, 'notContainsText': needle not in text,
                'beginsWith': text.startswith(needle), 'endsWith': text.endswith(needle)}[rule.type]
    if rule.type in ('containsBlanks', 'notContainsBlanks'):
        blank = value is None or str(value).strip() == ''
        return blank if rule.type == 'containsBlanks' else not blank
    if rule.type == 'expression':
        formula = str(rule.formula[0]).lstrip('=')
        formula = Translator('=' + formula, origin=anchor).translate_formula(coordinate)[1:]
        match = re.fullmatch(r'(.+?)(>=|<=|<>|=|>|<)(.+)', formula)
        if not match:
            raise ValueError('复杂条件公式暂不支持')
        return OPS[match[2]](_operand(match[1], sheet, reader), _operand(match[3], sheet, reader))
    if rule.type == 'cellIs':
        if value is None:
            return False
        values = [_operand(Translator('=' + f, origin=anchor).translate_formula(coordinate)[1:], sheet, reader)
                  for f in rule.formula]
        if rule.operator in ('between', 'notBetween'):
            matched = values[0] <= value <= values[1]
            return matched if rule.operator == 'between' else not matched
        if rule.operator in OPS and values:
            return OPS[rule.operator](value, values[0])
    raise ValueError('该类型条件格式暂不支持')


def differential_style(source, dxf, palette, style_fn):
    """仅覆盖差异样式中指定的属性，保留原字体的其他设置。"""
    merged = SimpleNamespace(**{key: copy(getattr(source, key)) for key in ('font', 'fill', 'border', 'alignment')})
    keys = set()
    mappings = {
        'font': {'name': 'fontFamily', 'sz': 'fontSize', 'b': 'fontWeight', 'i': 'fontStyle',
                 'color': 'color', 'u': 'textDecoration', 'strike': 'textDecoration', 'vertAlign': 'script'},
        'alignment': {'horizontal': 'textAlign', 'vertical': 'verticalAlign', 'wrapText': 'wrapText',
                      'shrinkToFit': 'shrinkToFit', 'indent': 'indent', 'textRotation': 'rotation'},
    }
    for part, attrs in mappings.items():
        change = getattr(dxf, part)
        if change is None:
            continue
        for attr, css in attrs.items():
            value = getattr(change, attr)
            # Font 的 b/i 默认值是 False，差异格式不应把未指定的粗体清除。
            if value is not None and not (part == 'font' and attr in ('b', 'i') and value is False):
                setattr(getattr(merged, part), attr, value)
                keys.add(css)
    if dxf.fill is not None:
        merged.fill = dxf.fill
        keys.update(('backgroundColor', 'backgroundImage', 'backgroundSize'))
    if dxf.border is not None:
        for edge in ('left', 'right', 'top', 'bottom', 'diagonal'):
            side = getattr(dxf.border, edge)
            if side is not None:
                setattr(merged.border, edge, copy(side))
                keys.add('border' + edge.title())
        if dxf.border.diagonal is not None:
            merged.border.diagonalUp, merged.border.diagonalDown = dxf.border.diagonalUp, dxf.border.diagonalDown
            keys.update(('diagonalUp', 'diagonalDown', 'diagonalColor'))
    result = style_fn(merged, palette)
    if dxf.fill is not None:
        result.setdefault('backgroundImage', 'none')
        result.setdefault('backgroundSize', 'auto')
    return {key: result[key] for key in keys if key in result}


def apply_conditions(groups, cells, raw_cells, palette, sheet, reader, style_fn, color, warnings):
    """按 Excel 优先级应用支持的规则，完整引用范围只用于计算阈值。"""
    stopped, claimed = set(), {}
    ordered = sorted(((sqref, rule) for sqref, rules in groups for rule in rules), key=lambda pair: pair[1].priority or 999999)
    for sqref, rule in ordered:
        try:
            ranges = [range_boundaries(part) for part in sqref.split()]
            positions = [p for p in cells if p not in stopped and any((a or 1) <= p[1] <= (c or 16384) and (b or 1) <= p[0] <= (d or 1048576) for a, b, c, d in ranges)]
            if not positions:
                continue
            numeric = []
            if rule.type in ('colorScale', 'dataBar', 'iconSet'):
                for part in sqref.split():
                    values, _hidden = reader.read(part, sheet)
                    numeric.extend(v for v in values if number(v))
            for r, c in positions:
                source, item = raw_cells[r, c], cells[r, c]
                value = item['value']
                if source.data_type == 'f' and reader.prime.get((sheet, r, c)) is None:
                    warnings.add('部分条件格式依赖未缓存的公式结果，暂不渲染')
                    continue
                style = {}
                if rule.type == 'colorScale' and number(value):
                    scale = rule.colorScale
                    limits = [_threshold(t, numeric) for t in scale.cfvo]
                    colors = [color(t, palette, '#FFFFFF') for t in scale.color]
                    if len(colors) != len(limits) or len(colors) < 2:
                        raise ValueError('色阶设置不完整')
                    step = min(next((i for i in range(len(limits)-1) if value <= limits[i+1]), len(limits)-2), len(colors)-2)
                    ratio = (value - limits[step]) / (limits[step+1] - limits[step]) if limits[step+1] != limits[step] else 0
                    style = {'backgroundColor': _blend(colors[step], colors[step+1], ratio), 'backgroundImage': 'none'}
                elif rule.type == 'dataBar' and number(value):
                    bar = rule.dataBar
                    low, high = _threshold(bar.cfvo[0], numeric), _threshold(bar.cfvo[1], numeric)
                    if low >= 0 or high <= 0:
                        minimum, maximum = bar.minLength if bar.minLength is not None else 10, bar.maxLength if bar.maxLength is not None else 90
                        ratio = max(0, min(1, (value-low) / (high-low))) if high != low else 1
                        length = minimum + (maximum-minimum) * (1-ratio if high <= 0 else ratio)
                        start = 100-length if high <= 0 else 0
                    else:
                        zero = -low / (high-low) * 100
                        end = max(0, min(100, (value-low) / (high-low) * 100))
                        start, length = min(zero, end), abs(end-zero)
                    item.setdefault('dataBar', {'start': start, 'width': length, 'color': '#D65B5B' if value < 0 else color(bar.color, palette, '#638EC6')})
                    if bar.showValue is False:
                        item['hideText'] = True
                elif rule.type == 'iconSet' and number(value):
                    icons = rule.iconSet
                    thresholds = [_threshold(t, numeric) for t in icons.cfvo]
                    level = sum(value >= t if threshold.gte is not False else value > t for t, threshold in zip(thresholds[1:], icons.cfvo[1:]))
                    count = len(thresholds)
                    if icons.reverse:
                        level = count - level - 1
                    name = icons.iconSet or '3TrafficLights1'
                    shapes = ['↓', '→', '↑'] if count == 3 else ['↓', '↘', '↗', '↑'] if count == 4 else ['↓', '↘', '→', '↗', '↑']
                    if 'TrafficLights' in name or 'Signs' in name or 'Symbols' in name:
                        shapes = ['●'] * count
                    elif 'Rating' in name or 'Quarters' in name:
                        shapes = ['○', '◔', '◑', '◕', '●'][-count:]
                    elif 'Flags' in name:
                        shapes = ['⚑'] * count
                    item.setdefault('icon', {'text': shapes[max(0, min(level, len(shapes)-1))], 'color': '#C94D46' if level == 0 else '#27905E' if level == count-1 else '#BA8A26'})
                    if icons.showValue is False:
                        item['hideText'] = True
                elif rule.type in ('colorScale', 'dataBar', 'iconSet'):
                    continue
                elif _matches(rule, value, source.coordinate, sqref.split()[0].split(':')[0].replace('$', ''), sheet, reader):
                    if rule.dxf:
                        style = differential_style(source, rule.dxf, palette, style_fn)
                        if rule.dxf.numFmt and 'numberFormat' not in claimed.setdefault((r, c), set()):
                            item['text'] = display(value, rule.dxf.numFmt.formatCode)[0]
                            claimed[r, c].add('numberFormat')
                            tint = format_color(value, rule.dxf.numFmt.formatCode)
                            if tint:
                                style.setdefault('color', tint)
                    if rule.stopIfTrue:
                        stopped.add((r, c))
                used = claimed.setdefault((r, c), set())
                for key, val in style.items():
                    if key not in used:
                        item['style'][key] = val
                        used.add(key)
        except (ValueError, TypeError, IndexError, AttributeError, KeyError, OverflowError):
            warnings.add('部分复杂条件格式未渲染（支持比较、文字、简单比较公式、色阶、数据条和图标集）')
