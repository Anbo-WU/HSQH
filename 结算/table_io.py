"""读取 HTML 导出、二进制 XLS 和 XLSX，统一为既有业务布局。"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal
import fnmatch
import html
from io import BytesIO
from pathlib import Path
import re

import openpyxl
import xlrd

EXTENSIONS = {'.xls', '.xlsx'}
# 原计算中的 O/U/V 分别是成交数量、了结远期收益、了结数量。
EXPORT_HEADERS = (
    '远期交易编号', '交易对手方', '交易员', '成交日期', '到期日期', '买卖方向',
    '多空方向', '结构类型', '标的代码', '标的名称', '标的品种', '期初标的价格',
    '交割价格', '远期价值(成交)', '成交数量', '成交手续费', '了结序号', '了结方式',
    '了结日期', '了结标的价格', '了结远期收益', '了结数量', '了结名义本金',
    '每手了结费用', '了结资金成本', '结算汇率', '了结总费用', '实现盈亏', '备注', '簿记账户',
)
ROW_RE = re.compile(r'<tr\b[^>]*>.*?</tr\s*>', re.I | re.S)
CELL_RE = re.compile(r'<(td|th)\b([^>]*)>(.*?)</\1\s*>', re.I | re.S)
TABLE_RE = re.compile(r'<table\b[^>]*>.*?</table\s*>', re.I | re.S)


def cell_text(cell):
    return html.unescape(re.sub(r'<[^>]+>', '', cell)).strip()


def text_value(value):
    if value is None:
        return ''
    if isinstance(value, datetime):
        return value.isoformat(sep=' ', timespec='seconds') if value.time() != time() else value.date().isoformat()
    if isinstance(value, (date, time)):
        return value.isoformat()
    if isinstance(value, float):
        # Excel 数值精度为 15 位有效数字，去掉二进制浮点尾差后再交给 Decimal。
        return format(Decimal(format(value, '.15g')), 'f')
    return str(value)


@dataclass
class Sheet:
    name: str
    rows: list[list[str]]

    @property
    def ncols(self):
        return max(map(len, self.rows), default=0)


def _html_text(raw):
    encodings = ('utf-16',) if raw.startswith((b'\xff\xfe', b'\xfe\xff')) else ('utf-8-sig', 'gb18030')
    for encoding in encodings:
        try:
            text = raw.decode(encoding)
        except UnicodeError:
            continue
        if TABLE_RE.search(text):
            return text
    raise ValueError('无法识别表格内容；支持 HTML 表格、二进制 .xls 及 .xlsx。文件可能损坏或不是表格，请重新导出。')


def _read_excel(raw, path):
    sheets = []
    if raw.startswith(b'PK'):
        # 使用文件内容而非扩展名，兼容被错误命名为 .xls 的 XLSX。
        book = openpyxl.load_workbook(BytesIO(raw), read_only=True, data_only=True)
        try:
            for sheet in book:
                rows = []
                for row in sheet.iter_rows():
                    if any(c.data_type == 'e' for c in row):
                        raise ValueError(f'{sheet.title} 中存在 Excel 错误单元格，请先修正')
                    rows.append([text_value(c.value) for c in row])
                sheets.append(Sheet(sheet.title, rows))
        finally:
            book.close()
    elif raw.startswith(bytes.fromhex('d0cf11e0a1b11ae1')):
        book = xlrd.open_workbook(file_contents=raw, on_demand=True)
        try:
            for sheet in book.sheets():
                rows = []
                for row_index in range(sheet.nrows):
                    values = []
                    for cell in sheet.row(row_index):
                        if cell.ctype == xlrd.XL_CELL_ERROR:
                            raise ValueError(f'{sheet.name} 第 {row_index+1} 行含 Excel 错误单元格')
                        value = (xlrd.xldate_as_datetime(cell.value, book.datemode)
                                 if cell.ctype == xlrd.XL_CELL_DATE else cell.value)
                        values.append(text_value(value))
                    rows.append(values)
                sheets.append(Sheet(sheet.name, rows))
        finally:
            book.release_resources()
    else:
        return None
    return sheets


def read_sheets(path: Path):
    try:
        raw = path.read_bytes()
        sheets = _read_excel(raw, path)
        if sheets is not None:
            return sheets
        text = _html_text(raw)
        return [Sheet(f'表格{i}', [[cell_text(c.group(3)) for c in CELL_RE.finditer(row)]
                                  for row in ROW_RE.findall(table.group())])
                for i, table in enumerate(TABLE_RE.finditer(text), 1)]
    except Exception as exc:
        raise ValueError(f'无法读取表格 {path}：{exc}') from exc


def _header_key(value):
    return re.sub(r'\s+', '', value).replace('（', '(').replace('）', ')')


def _column_order(headers):
    keys = [_header_key(h) for h in headers]
    missing = [h for h in EXPORT_HEADERS if _header_key(h) not in keys]
    if missing:
        raise ValueError('导出表缺少必需列：' + '、'.join(missing))
    duplicates = [h for h in EXPORT_HEADERS if keys.count(_header_key(h)) != 1]
    if duplicates:
        raise ValueError('导出表存在重复表头：' + '、'.join(duplicates))
    return [keys.index(_header_key(h)) for h in EXPORT_HEADERS]


def export_html(path: Path):
    """按表头对齐成原 30 列布局；HTML 输入保留原单元格标记/外层样式。"""
    raw = path.read_bytes()
    try:
        sheets = _read_excel(raw, path)
        if sheets is None:
            text = _html_text(raw)
        else:
            candidates = []
            errors = []
            for sheet in sheets:
                rows = [r for r in sheet.rows if any(v.strip() for v in r)]
                if not rows:
                    continue
                try:
                    _column_order(rows[0])
                    candidates.append((sheet, rows))
                except ValueError as exc:
                    errors.append(f'{sheet.name}：{exc}')
            if len(candidates) != 1:
                raise ValueError('需有且仅有一个包含完整导出表头的工作表。' + '；'.join(errors))
            _, rows = candidates[0]
            text = '<html><head><meta charset="utf-8"></head><body><table border="1">'
            text += ''.join('<tr>'+''.join('<td>'+html.escape(v)+'</td>' for v in r)+'</tr>' for r in rows)
            text += '</table></body></html>'
        table = TABLE_RE.search(text)
        rows = list(ROW_RE.finditer(table.group()))
        if not rows:
            raise ValueError('导出表没有表头')
        headers = list(CELL_RE.finditer(rows[0].group()))
        order = _column_order([cell_text(c.group(3)) for c in headers])
        output = []
        for number, row in enumerate(rows, 1):
            cells = list(CELL_RE.finditer(row.group()))
            if not cells or not any(cell_text(c.group(3)) for c in cells):
                continue
            if len(cells) <= max(order):
                raise ValueError(f'第 {number} 行缺少必需列，已停止，未生成不完整结果')
            output.append('<tr>'+''.join(cells[i].group() for i in order)+'</tr>')
        start, end = rows[0].start(), rows[-1].end()
        replacement = table.group()[:start] + '\n'.join(output) + table.group()[end:]
        return text[:table.start()] + replacement + text[table.end():]
    except Exception as exc:
        raise ValueError(f'无法读取导出表 {path}：{exc}') from exc


def spreadsheet_files(directory):
    return sorted(p for p in directory.iterdir() if p.is_file() and p.suffix.lower() in EXTENSIONS
                  and not p.name.startswith(('~$', '.')))


def find_one_file(directory: Path, preferred_names: list[str], fallback: str):
    files = spreadsheet_files(directory)
    # 沿用原 .xls 优先级；同名有两种格式时明确打印选择，并提供 CLI 指定入口。
    for name in preferred_names:
        stem = Path(name).stem
        matches = [p for p in files if p.stem == stem]
        if matches:
            matches.sort(key=lambda p: (p.suffix.lower() != '.xls', p.name))
            if len(matches) > 1:
                print(f'提示：同名存在 .xls/.xlsx，使用 {matches[0].name}；可用 --export-file 或 --due-file 指定其他文件。')
            return matches[0]
    stem_pattern = Path(fallback).stem
    matches = [p for p in files if fnmatch.fnmatchcase(p.stem, stem_pattern)]
    if len(matches) != 1:
        raise ValueError(f'{directory} 中需要唯一的 {stem_pattern}.xls/.xlsx，实际找到：'
                         + ('、'.join(p.name for p in matches) or '无'))
    return matches[0]
