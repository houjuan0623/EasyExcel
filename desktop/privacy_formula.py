"""脱敏公式：仅保留允许的函数和本包内引用，原常量只进入本机映射。

Copyright (c) 2026 AbeLin · MIT License
"""
import re
import math

from openpyxl.formula import Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import range_boundaries

FUNCTIONS = set('SUM SUMIF SUMIFS SUMPRODUCT SUMSQ AVERAGE AVERAGEIF AVERAGEIFS COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS '
                'MIN MAX MEDIAN MODE STDEV STDEV.S STDEV.P VAR VAR.S VAR.P IF IFS IFERROR IFNA AND OR NOT TRUE FALSE '
                'ROUND ROUNDUP ROUNDDOWN INT TRUNC ABS MOD POWER SQRT LOG LN EXP SIGN CEILING FLOOR CEILING.MATH FLOOR.MATH '
                'INDEX MATCH XMATCH XLOOKUP VLOOKUP HLOOKUP CHOOSE ROW ROWS COLUMN COLUMNS COLUMNS CONCAT CONCATENATE '
                'LEFT RIGHT MID LEN TRIM CLEAN UPPER LOWER SUBSTITUTE REPLACE FIND SEARCH EXACT VALUE TEXT TEXTJOIN '
                'DATE YEAR MONTH DAY HOUR MINUTE SECOND TIME EDATE EOMONTH DATEDIF DAYS DAYS360 NETWORKDAYS WORKDAY '
                'ISBLANK ISNUMBER ISTEXT ISERROR ISNA ISLOGICAL N T NA SUBTOTAL AGGREGATE ROUND NPV IRR PMT PV FV '
                'RATE NPER IPMT PPMT RANK RANK.EQ LARGE SMALL PERCENTILE.INC QUARTILE.INC TRANSPOSE FILTER SORT UNIQUE'.split())
ERRORS = {'#N/A', '#REF!', '#VALUE!', '#DIV/0!', '#NUM!', '#NAME?', '#NULL!'}
CELL_RANGE = re.compile(r'\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*', re.I)


class FormulaSanitizer:
    """保留代表性计算关系；演示范围固定，不暴露原表记录总数。"""
    def __init__(self, sheets, row_count, mode):
        self.sheets, self.row_count, self.mode = sheets, row_count, mode
        self.parameters = {}
        self._keys = {}

    def reference(self, reference, source):
        """拒绝命名范围、结构化引用、外部工作簿和未导出的 Sheet。"""
        target, span = source, reference
        if '!' in reference:
            name, span = reference.rsplit('!', 1)
            name = name[1:-1].replace("''", "'") if name.startswith("'") and name.endswith("'") else name
            target = next((item for item in self.sheets if item['path'] == source['path'] and item['sheet'] == name), None)
            if target is None:
                raise ValueError('不支持的引用')
        if not CELL_RANGE.fullmatch(span):
            raise ValueError('不支持的引用')
        left, top, right, bottom = range_boundaries(span)
        if left is not None and (right > target['width'] or any(c not in target['visible'] for c in range(left, right+1))):
            raise ValueError('引用了未导出的列')
        header, end = target['header'], target['header'] + self.row_count
        # 跨行范围统一为演示范围；逐行引用保留相对关系，超出范围的固定位置落到首个演示行。
        multi_row = top is not None and bottom is not None and top != bottom
        parts = span.upper().split(':')
        rendered = []
        for i, part in enumerate(parts):
            match = re.fullmatch(r'(\$?[A-Z]{1,3})?(\$?)(\d+)?', part)
            if not match:
                raise ValueError('不支持的引用')
            col, absolute, row = match.groups()
            if row:
                value = int(row)
                if value > header:
                    value = (header+1 if i == 0 else end) if multi_row else value if value <= end else header+1
                rendered.append((col or '') + absolute + str(value))
            else:
                rendered.append(col)
        prefix = f"'{target['id']}'!" if target is not source else ''
        return prefix + ':'.join(rendered)

    def sanitize(self, formula, origin, destination, source, field_id):
        """常量改成匿名参数引用，非法表达式由调用方统一回退。"""
        if not isinstance(formula, str) or len(formula) > 16000:
            raise ValueError('不支持的公式')
        translated = Translator(formula, origin=origin).translate_formula(destination)
        tokens, output, pending = Tokenizer(translated).items, [], []
        for index, token in enumerate(tokens):
            value = token.value
            if token.type == 'FUNC':
                if token.subtype == 'OPEN':
                    function = value[:-1].upper()
                    if function not in FUNCTIONS:
                        raise ValueError('不支持的函数')
                    output.append(function + '(')
                else:
                    output.append(')')
            elif token.type == 'OPERAND' and token.subtype == 'RANGE':
                output.append(self.reference(value, source))
            elif token.type == 'OPERAND' and token.subtype in ('TEXT', 'NUMBER', 'LOGICAL'):
                key = source['id'], field_id, index
                name = self._keys.get(key)
                if name is None:
                    name = f'参数_{len(self.parameters) + len(pending) + 1:04d}'
                    real = value[1:-1].replace('""', '"') if token.subtype == 'TEXT' else value == 'TRUE' if token.subtype == 'LOGICAL' else int(value) if re.fullmatch(r'[+-]?\d+', value) else float(value)
                    if isinstance(real, float) and not math.isfinite(real):
                        raise ValueError('无效公式常量')
                    pending.append((key, name, {'kind': token.subtype, 'value': real}))
                row = int(name.split('_')[-1]) + 1
                output.append(f"'演示参数'!$B${row}")
            elif token.type == 'OPERAND' and token.subtype == 'ERROR' and value in ERRORS:
                output.append(value)
            elif token.type in ('OPERATOR-PREFIX', 'OPERATOR-INFIX', 'OPERATOR-POSTFIX', 'PAREN', 'SEP', 'WHITE-SPACE') and re.fullmatch(r'[+\-*/^&=<>%(),;\s]+', value):
                output.append(value)
            else:
                raise ValueError('不支持的公式语法')
        if not output or len(self.parameters) + len(pending) > 5000:
            raise ValueError('公式参数过多')
        for key, name, item in pending:
            self._keys[key] = name
            self.parameters[name] = item
        return '=' + ''.join(output)


def demo_parameter(kind, mode):
    """演示参数完全固定，与原常量的值无关。"""
    if mode != 'synthetic':
        return None
    return '示例条件' if kind == 'TEXT' else True if kind == 'LOGICAL' else 1
