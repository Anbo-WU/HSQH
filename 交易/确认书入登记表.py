#!/usr/bin/env python3
"""递归读取“确认书文件”中的 PDF 文件名，填入合同章登记表的 B5 起始单元格。"""

from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path
from typing import Iterator


PDF_FOLDER_NAME = "确认书文件"
WORKBOOK_NAME = "登记表模版.xls"
START_ROW = 5
COLUMN_INDEX = 1  # B 列（从 0 开始计数）
TEMPLATE_ROW_COUNT = 29
DATA_END_ROW = 27  # 每页第 28、29 行是模板底部说明

# 兼容实际文件名中的下划线、空格和连字符。
REMOVE_PATTERN = re.compile(
    r"[\s_-]*商品交易确认书[\s_-]*【HFSY】[\s_-]*",
    flags=re.IGNORECASE,
)


def natural_key(path: Path) -> list[tuple[int, object]]:
    """让 7.9 排在 7.10 前，同时保持中英文文件名排序稳定。"""
    parts = re.split(r"(\d+)", path.name.casefold())
    return [(0, int(part)) if part.isdigit() else (1, part) for part in parts]


def iter_pdfs(folder: Path) -> Iterator[Path]:
    """先读当前层 PDF，再按名称自然排序递归所有子文件夹。"""
    paths = sorted(folder.iterdir(), key=natural_key)
    for path in paths:
        if path.is_file() and path.suffix.casefold() == ".pdf":
            yield path
    for path in paths:
        if path.is_dir():
            yield from iter_pdfs(path)


def clean_pdf_name(pdf_path: Path) -> str:
    """去掉 .pdf 扩展名和指定文字及其相邻分隔符。"""
    return REMOVE_PATTERN.sub("", pdf_path.stem).strip()


def collect_names(folder: Path) -> list[str]:
    return [clean_pdf_name(path) for path in iter_pdfs(folder)]


def load_xls_modules():
    try:
        import xlrd  # type: ignore[import-not-found]
        from xlutils.filter import (  # type: ignore[import-not-found]
            XLRDReader,
            XLWTWriter,
            process,
        )
    except ImportError:
        print(
            "缺少写入 .xls 所需的 Python 库。请先执行：\n"
            "  python3 -m pip install xlrd==1.2.0 xlwt==1.3.0 xlutils==2.0.0",
            file=sys.stderr,
        )
        raise SystemExit(1)
    return xlrd, XLRDReader, XLWTWriter, process


def write_names(workbook_path: Path, names: list[str]) -> None:
    xlrd, XLRDReader, XLWTWriter, process = load_xls_modules()
    read_book = xlrd.open_workbook(str(workbook_path), formatting_info=True)

    if read_book.nsheets == 0:
        raise RuntimeError("表格中没有工作表。")

    sheet_index = 0
    read_sheet = read_book.sheet_by_index(sheet_index)

    # 部分 WPS 生成的 xls 会带有空的数字格式名，xlwt 无法直接保存。
    # 这些空格式在 Excel/WPS 中等同于常规格式。
    for cell_format in read_book.format_map.values():
        if cell_format.format_str is None:
            cell_format.format_str = "General"

    writer = XLWTWriter()
    process(XLRDReader(read_book, workbook_path.name), writer)
    write_book = writer.output[0][1]
    write_sheet = write_book.get_sheet(sheet_index)
    start_index = START_ROW - 1

    if read_sheet.nrows < TEMPLATE_ROW_COUNT:
        raise RuntimeError(
            f"模板行数不足：需要至少 {TEMPLATE_ROW_COUNT} 行，"
            f"实际只有 {read_sheet.nrows} 行。"
        )

    rows_per_page = DATA_END_ROW - START_ROW + 1
    page_count = (len(names) + rows_per_page - 1) // rows_per_page

    def original_style(row_index: int):
        source_row = row_index % TEMPLATE_ROW_COUNT
        xf_index = read_sheet.cell_xf_index(source_row, COLUMN_INDEX)
        return writer.style_list[xf_index]

    # 如果之前已运行过脚本，先删掉第一页之后的旧页。
    rows = write_sheet._Worksheet__rows
    for row_index in [index for index in rows if index >= TEMPLATE_ROW_COUNT]:
        del rows[row_index]
    write_sheet._Worksheet__merged_ranges = [
        merged_range
        for merged_range in write_sheet._Worksheet__merged_ranges
        if merged_range[0] < TEMPLATE_ROW_COUNT
    ]
    write_sheet.last_used_row = TEMPLATE_ROW_COUNT - 1

    merged_cells = set()
    for row_low, row_high, col_low, col_high in read_sheet.merged_cells:
        if row_low >= TEMPLATE_ROW_COUNT:
            continue
        for row_index in range(row_low, min(row_high, TEMPLATE_ROW_COUNT)):
            for col_index in range(col_low, col_high):
                merged_cells.add((row_index, col_index))

    def copy_template_page(page_number: int) -> None:
        """复制整页模板，包括值、样式、行高和合并单元格。"""
        row_offset = page_number * TEMPLATE_ROW_COUNT

        for source_row in range(TEMPLATE_ROW_COUNT):
            source_info = read_sheet.rowinfo_map.get(source_row)
            if source_info is not None:
                target_row = write_sheet.row(row_offset + source_row)
                target_row.height = source_info.height
                target_row.has_default_height = source_info.has_default_height
                target_row.height_mismatch = source_info.height_mismatch
                target_row.level = source_info.outline_level
                target_row.collapse = source_info.outline_group_starts_ends
                target_row.hidden = source_info.hidden
                target_row.space_above = source_info.additional_space_above
                target_row.space_below = source_info.additional_space_below
                if source_info.has_default_xf_index:
                    target_row.set_style(writer.style_list[source_info.xf_index])

            for col_index in range(read_sheet.ncols):
                if (source_row, col_index) in merged_cells:
                    continue
                source_cell = read_sheet.cell(source_row, col_index)
                style = writer.style_list[source_cell.xf_index]
                target_row = write_sheet.row(row_offset + source_row)
                if source_cell.ctype == xlrd.XL_CELL_EMPTY:
                    continue
                if source_cell.ctype == xlrd.XL_CELL_BLANK:
                    target_row.set_cell_blank(col_index, style)
                elif source_cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    target_row.set_cell_boolean(col_index, source_cell.value, style)
                elif source_cell.ctype == xlrd.XL_CELL_ERROR:
                    target_row.set_cell_error(col_index, source_cell.value, style)
                elif source_cell.ctype in (xlrd.XL_CELL_NUMBER, xlrd.XL_CELL_DATE):
                    target_row.set_cell_number(col_index, source_cell.value, style)
                else:
                    target_row.set_cell_text(col_index, str(source_cell.value), style)

        for row_low, row_high, col_low, col_high in read_sheet.merged_cells:
            if row_low >= TEMPLATE_ROW_COUNT:
                continue
            top_left = read_sheet.cell(row_low, col_low)
            style = writer.style_list[top_left.xf_index]
            write_sheet.write_merge(
                row_offset + row_low,
                row_offset + min(row_high, TEMPLATE_ROW_COUNT) - 1,
                col_low,
                col_high - 1,
                top_left.value,
                style,
            )

    for page_number in range(1, page_count):
        copy_template_page(page_number)

    # 每页只清空 B5:B27，不覆盖第 28、29 行的页尾说明。
    for page_number in range(page_count):
        row_offset = page_number * TEMPLATE_ROW_COUNT
        for source_row in range(start_index, DATA_END_ROW):
            row_index = row_offset + source_row
            write_sheet.write(row_index, COLUMN_INDEX, "", original_style(row_index))

    for offset, name in enumerate(names):
        page_number, row_on_page = divmod(offset, rows_per_page)
        row_index = page_number * TEMPLATE_ROW_COUNT + start_index + row_on_page
        write_sheet.write(row_index, COLUMN_INDEX, name, original_style(row_index))

    # 在每个模板块之间加手动分页，确保双面连续打印。
    write_sheet.set_horz_page_breaks(
        [
            (page_number * TEMPLATE_ROW_COUNT, 0, read_sheet.ncols - 1)
            for page_number in range(1, page_count)
        ]
    )

    temp_path = workbook_path.with_name(f".{workbook_path.stem}.tmp.xls")
    try:
        write_book.save(str(temp_path))
        os.replace(temp_path, workbook_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--preview",
        action="store_true",
        help="只显示排序和提取结果，不修改表格",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    base_folder = Path(__file__).resolve().parent
    pdf_folder = base_folder / PDF_FOLDER_NAME
    workbook_path = base_folder / WORKBOOK_NAME

    if not pdf_folder.is_dir():
        print(f"找不到文件夹：{pdf_folder}", file=sys.stderr)
        return 1
    if not workbook_path.is_file():
        print(f"找不到表格：{workbook_path}", file=sys.stderr)
        return 1

    names = collect_names(pdf_folder)
    if not names:
        print(f"未在 {pdf_folder} 及其子文件夹中找到 PDF。", file=sys.stderr)
        return 1

    for number, name in enumerate(names, start=1):
        print(f"{number:>3}. {name}")

    if args.preview:
        print(f"\n预览完成：共 {len(names)} 个 PDF，未修改表格。")
        return 0

    write_names(workbook_path, names)
    rows_per_page = DATA_END_ROW - START_ROW + 1
    page_count = (len(names) + rows_per_page - 1) // rows_per_page
    print(
        f"\n已完成：共 {len(names)} 条，生成 {page_count} 页模板，"
        f"每页最多填写 {rows_per_page} 条。"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
