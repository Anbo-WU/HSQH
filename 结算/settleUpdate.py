#!/usr/bin/env python3
"""按《需求文档.txt》批量改版当前目录中的 HTML 格式 .xls 文件。"""

from __future__ import annotations

import argparse
import html
import re
from decimal import Decimal, InvalidOperation
from pathlib import Path


DELETE_COLUMNS = {2, 5, 23, 24, 25, 26, 27, 29}  # C/F/X/Y/Z/AA/AB/AD
SUM_COLUMNS = {14, 20, 21}  # O/U/V（均为删除列之前的列号）
INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
ROW_RE = re.compile(r"<tr\b[^>]*>.*?</tr\s*>", re.I | re.S)
CELL_RE = re.compile(r"<(td|th)\b([^>]*)>(.*?)</\1\s*>", re.I | re.S)


def cell_text(cell_html: str) -> str:
    """取得单元格的可见文字（这些导出表通常没有嵌套标签）。"""
    match = CELL_RE.fullmatch(cell_html.strip())
    if not match:
        return ""
    without_tags = re.sub(r"<[^>]+>", "", match.group(3))
    return html.unescape(without_tags).strip()


def replace_cell_text(cell_html: str, value: str) -> str:
    match = CELL_RE.fullmatch(cell_html.strip())
    if not match:
        raise ValueError("无法识别单元格")
    tag, attrs = match.group(1), match.group(2)
    return f"<{tag}{attrs}>{html.escape(value, quote=False)}</{tag}>"


def decimal_value(value: str, *, location: str) -> Decimal:
    normalized = value.strip().replace(",", "")
    if not normalized:
        return Decimal("0")
    try:
        return Decimal(normalized)
    except InvalidOperation as exc:
        raise ValueError(f"{location} 不是有效数字：{value!r}") from exc


def decimal_places(value: str) -> int:
    value = value.strip().replace(",", "")
    return len(value.rsplit(".", 1)[1]) if "." in value else 0


def format_decimal(value: Decimal, places: int) -> str:
    # 避免 Decimal('-0') 输出为负零。
    if value == 0:
        value = abs(value)
    return f"{value:.{places}f}"


def split_row(row_html: str) -> list[str]:
    return [match.group(0) for match in CELL_RE.finditer(row_html)]


def make_row(cells: list[str]) -> str:
    return "<tr>" + "".join(cells) + "</tr>"


def safe_name_part(value: str) -> str:
    return INVALID_FILENAME_CHARS.sub("_", value).strip().rstrip(". ")


def transform(source: Path, output_dir: Path) -> Path:
    raw = source.read_bytes()
    if not raw.lstrip().lower().startswith((b"<html", b"<!doctype html")):
        raise ValueError("不是本工具支持的 HTML 格式 .xls 文件")
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError("文件不是 UTF-8 编码的 HTML 格式 .xls") from exc

    table_match = re.search(r"<table\b[^>]*>.*?</table\s*>", text, re.I | re.S)
    if not table_match:
        raise ValueError("未找到表格")
    rows = ROW_RE.findall(table_match.group(0))
    if len(rows) < 2:
        raise ValueError("表格没有数据行")

    # 某些 Excel 导出文件会在数据末尾附带仅用于保存格式的空白行，
    # 其单元格数量甚至可能少于表头；这类行不属于业务数据，应忽略。
    parsed = []
    for row in rows:
        cells = split_row(row)
        if cells and any(cell_text(cell) for cell in cells):
            parsed.append(cells)
    if len(parsed) < 2:
        raise ValueError("表格没有有效数据行")
    if any(len(row) < 30 for row in parsed):
        raise ValueError("表格列数不足 30 列，无法按需求处理")

    headers, data_rows = parsed[0], parsed[1:]
    headers[20] = replace_cell_text(headers[20], "客户了解远期收益")

    # U 列切换为客户视角：数值乘以 -1，并保留源数据的小数位数。
    for row_number, row in enumerate(data_rows, start=2):
        original = cell_text(row[20])
        changed = -decimal_value(original, location=f"{source.name} U{row_number}")
        row[20] = replace_cell_text(row[20], format_decimal(changed, decimal_places(original)))

    totals: dict[int, Decimal] = {}
    total_places: dict[int, int] = {}
    for column in SUM_COLUMNS:
        values = [cell_text(row[column]) for row in data_rows]
        totals[column] = sum(
            (decimal_value(value, location=f"{source.name} 第{column + 1}列") for value in values),
            Decimal("0"),
        )
        total_places[column] = max((decimal_places(value) for value in values), default=2)

    # O、V 合计不一致时，用 V 列逐行覆盖 O 列（包含随后生成的合计）。
    if totals[14] != totals[21]:
        for row in data_rows:
            row[14] = replace_cell_text(row[14], cell_text(row[21]))
        totals[14] = totals[21]
        total_places[14] = total_places[21]

    total_cells: list[str] = []
    for column in range(30):
        if column == 0:
            value = "合计"
        elif column in SUM_COLUMNS:
            value = format_decimal(totals[column], total_places[column])
        else:
            value = "/"
        total_cells.append(f"<td>{html.escape(value, quote=False)}</td>")

    all_rows = [headers, *data_rows, total_cells]
    kept_rows = [[cell for index, cell in enumerate(row) if index not in DELETE_COLUMNS] for row in all_rows]
    new_table = "<table><thead>" + make_row(kept_rows[0]) + "</thead><tbody>"
    new_table += "".join(make_row(row) for row in kept_rows[1:]) + "</tbody></table>"
    result = text[: table_match.start()] + new_table + text[table_match.end() :]

    customer = safe_name_part(cell_text(parsed[1][1])[:4])
    reference = safe_name_part(cell_text(parsed[1][28])[-10:])
    if not customer or not reference:
        raise ValueError("B2 或 AC2 内容不足，无法生成新文件名")
    output_dir.mkdir(parents=True, exist_ok=True)
    destination = output_dir / f"{customer}{reference}{source.stem}{source.suffix}"
    destination.write_text(result, encoding="utf-8")
    return destination


def main() -> int:
    project_dir = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description="批量改版远期了结记录 .xls 文件")
    parser.add_argument("--input-dir", type=Path, default=project_dir / "老版本")
    parser.add_argument("--output-dir", type=Path, help="输出目录，默认是项目目录下的“新版本”")
    args = parser.parse_args()

    input_dir = args.input_dir.resolve()
    output_dir = (args.output_dir or project_dir / "新版本").resolve()
    sources = sorted(path for path in input_dir.glob("*.xls") if path.is_file())
    if not sources:
        print(f"未在 {input_dir} 找到 .xls 文件")
        return 1

    failures = 0
    for source in sources:
        try:
            destination = transform(source, output_dir)
            print(f"成功：{source.name} -> {destination}")
        except Exception as exc:
            failures += 1
            print(f"失败：{source.name}：{exc}")
    print(f"处理完成：成功 {len(sources) - failures} 个，失败 {failures} 个")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
