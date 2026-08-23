#!/usr/bin/env python3
"""递归读取“确认书文件”中的 PDF，从登记表模板生成当日登记表。

每个模板页可填 23 条（第 5–27 行）；数量超出时，会在同一张
工作表中向下复制足够的整页模板，并在页之间设置手动分页。
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import tempfile
from copy import copy
from datetime import date
from pathlib import Path
from typing import Iterator


START_ROW = 5
DATA_END_ROW = 27
TEMPLATE_ROW_COUNT = 29
FIRST_COLUMN_INDEX = 0
LAST_COLUMN_INDEX = 8  # I 列

DATE_COLUMN_INDEX = 0
NAME_COLUMN_INDEX = 1
COPY_COUNT_COLUMN_INDEX = 2
STAMP_COUNT_COLUMN_INDEX = 3
APPROVAL_COLUMN_INDEX = 4
REVIEW_COLUMN_INDEX = 5

CHECKED_TEXT = "是☑   否☐"
DATE_NUMBER_FORMAT = "yyyy-mm-dd"

# 兼容实际文件名中的下划线、空格和连字符。
REMOVE_PATTERN = re.compile(
    r"[\s_-]*商品交易确认书[\s_-]*【HFSY】[\s_-]*",
    flags=re.IGNORECASE,
)


def natural_key(path: Path) -> list[tuple[int, object]]:
    """让 7.9 排在 7.10 之前，同时保持中英文文件名排序稳定。"""
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


def required_page_count(item_count: int) -> int:
    rows_per_page = DATA_END_ROW - START_ROW + 1
    return (item_count + rows_per_page - 1) // rows_per_page


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


def restore_template_print_settings(sheet) -> None:
    """恢复 xlutils 复制时会丢失的模板打印设置。"""
    sheet.print_headers = 0
    sheet.print_grid = 0
    sheet.header_str = b""
    sheet.footer_str = b""
    sheet.print_centered_vert = 0
    sheet.print_centered_horz = 0

    sheet.left_margin = 0.75
    sheet.right_margin = 0.75
    sheet.top_margin = 0.275
    sheet.bottom_margin = 0.19652777777777777

    sheet.paper_size_code = 9  # A4
    sheet.print_scaling = 92
    sheet.start_page_number = 1
    sheet.fit_width_to_pages = 1
    sheet.fit_height_to_pages = 1
    sheet.fit_num_pages = 0
    sheet.print_in_rows = 0
    sheet.portrait = 0  # 横向

    sheet.print_colour = True
    sheet.print_draft = 0
    sheet.print_notes = 0
    sheet.print_notes_at_end = 0
    sheet.print_omit_errors = 0
    sheet.print_hres = 600
    sheet.print_vres = 600
    sheet.header_margin = 0.3145833333333333
    sheet.footer_margin = 0.3145833333333333
    sheet.copies_num = 1


def write_registration(
    template_path: Path,
    output_path: Path,
    names: list[str],
    stamp_date: date,
) -> int:
    """只读模板，将完整结果写入按日期命名的新 xls 文件。"""
    xlrd, XLRDReader, XLWTWriter, process = load_xls_modules()
    read_book = xlrd.open_workbook(str(template_path), formatting_info=True)

    if read_book.nsheets == 0:
        raise RuntimeError("登记表模板中没有工作表。")

    sheet_index = 0
    read_sheet = read_book.sheet_by_index(sheet_index)
    if read_sheet.nrows < TEMPLATE_ROW_COUNT:
        raise RuntimeError(
            f"模板行数不足：需要至少 {TEMPLATE_ROW_COUNT} 行，"
            f"实际只有 {read_sheet.nrows} 行。"
        )
    if read_sheet.ncols <= LAST_COLUMN_INDEX:
        raise RuntimeError("模板列数不足：需要至少 A:I 九列。")

    # 部分 WPS 生成的 xls 会带有空的数字格式名，xlwt 无法直接保存。
    for cell_format in read_book.format_map.values():
        if cell_format.format_str is None:
            cell_format.format_str = "General"

    writer = XLWTWriter()
    process(XLRDReader(read_book, template_path.name), writer)
    write_book = writer.output[0][1]
    write_sheet = write_book.get_sheet(sheet_index)
    restore_template_print_settings(write_sheet)

    rows_per_page = DATA_END_ROW - START_ROW + 1
    page_count = required_page_count(len(names))
    start_index = START_ROW - 1

    merged_cells = set()
    for row_low, row_high, col_low, col_high in read_sheet.merged_cells:
        if row_low >= TEMPLATE_ROW_COUNT:
            continue
        for row_index in range(row_low, min(row_high, TEMPLATE_ROW_COUNT)):
            for column_index in range(col_low, col_high):
                merged_cells.add((row_index, column_index))

    def cell_style(source_row: int, column_index: int):
        xf_index = read_sheet.cell_xf_index(source_row, column_index)
        return writer.style_list[xf_index]

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

            for column_index in range(read_sheet.ncols):
                if (source_row, column_index) in merged_cells:
                    continue
                source_cell = read_sheet.cell(source_row, column_index)
                style = writer.style_list[source_cell.xf_index]
                target_row = write_sheet.row(row_offset + source_row)
                if source_cell.ctype == xlrd.XL_CELL_EMPTY:
                    continue
                if source_cell.ctype == xlrd.XL_CELL_BLANK:
                    target_row.set_cell_blank(column_index, style)
                elif source_cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    target_row.set_cell_boolean(column_index, source_cell.value, style)
                elif source_cell.ctype == xlrd.XL_CELL_ERROR:
                    target_row.set_cell_error(column_index, source_cell.value, style)
                elif source_cell.ctype in (xlrd.XL_CELL_NUMBER, xlrd.XL_CELL_DATE):
                    target_row.set_cell_number(column_index, source_cell.value, style)
                else:
                    target_row.set_cell_text(column_index, str(source_cell.value), style)

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

    date_style = copy(cell_style(start_index, DATE_COLUMN_INDEX))
    date_style.num_format_str = DATE_NUMBER_FORMAT

    # 先清空每页 A5:I27 的模板旧数据，仅保留单元格样式。
    for page_number in range(page_count):
        row_offset = page_number * TEMPLATE_ROW_COUNT
        for source_row in range(start_index, DATA_END_ROW):
            target_row = row_offset + source_row
            for column_index in range(FIRST_COLUMN_INDEX, LAST_COLUMN_INDEX + 1):
                write_sheet.write(
                    target_row,
                    column_index,
                    "",
                    cell_style(source_row, column_index),
                )

    # 只填充存在真实 PDF 数据的行；G、H、I 列继续留空。
    for offset, name in enumerate(names):
        page_number, row_on_page = divmod(offset, rows_per_page)
        source_row = start_index + row_on_page
        row_index = page_number * TEMPLATE_ROW_COUNT + source_row

        write_sheet.write(row_index, DATE_COLUMN_INDEX, stamp_date, date_style)
        write_sheet.write(
            row_index,
            NAME_COLUMN_INDEX,
            name,
            cell_style(source_row, NAME_COLUMN_INDEX),
        )
        write_sheet.write(
            row_index,
            COPY_COUNT_COLUMN_INDEX,
            1,
            cell_style(source_row, COPY_COUNT_COLUMN_INDEX),
        )
        write_sheet.write(
            row_index,
            STAMP_COUNT_COLUMN_INDEX,
            1,
            cell_style(source_row, STAMP_COUNT_COLUMN_INDEX),
        )
        write_sheet.write(
            row_index,
            APPROVAL_COLUMN_INDEX,
            CHECKED_TEXT,
            cell_style(source_row, APPROVAL_COLUMN_INDEX),
        )
        write_sheet.write(
            row_index,
            REVIEW_COLUMN_INDEX,
            CHECKED_TEXT,
            cell_style(source_row, REVIEW_COLUMN_INDEX),
        )

    # 在每个模板块之间加手动分页。
    write_sheet.set_horz_page_breaks(
        [
            (page_number * TEMPLATE_ROW_COUNT, 0, read_sheet.ncols - 1)
            for page_number in range(1, page_count)
        ]
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_output: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{output_path.stem}-",
            suffix=".xls",
            dir=output_path.parent,
            delete=False,
        ) as temp_file:
            temp_output = Path(temp_file.name)

        write_book.save(str(temp_output))
        os.replace(temp_output, output_path)
        temp_output = None
    finally:
        if temp_output is not None and temp_output.exists():
            temp_output.unlink()

    return page_count


def run_registration(
    pdf_folder: Path,
    template_path: Path,
    output_path: Path,
    stamp_date: date,
    preview: bool = False,
) -> tuple[int, int]:
    """生成登记表，返回（PDF 数量，登记表页数）。"""
    pdf_folder = pdf_folder.expanduser().resolve()
    template_path = template_path.expanduser().resolve()
    output_path = output_path.expanduser().resolve()

    required_paths = [
        (pdf_folder, "PDF 文件夹"),
        (template_path, "登记表模板"),
    ]
    for path, description in required_paths:
        if not path.exists():
            raise RuntimeError(f"找不到{description}：{path}")

    names = collect_names(pdf_folder)
    if not names:
        raise RuntimeError(f"未在 {pdf_folder} 及其子文件夹中找到 PDF。")

    page_count = required_page_count(len(names))
    for number, name in enumerate(names, start=1):
        print(f"{number:>3}. {name}")

    print(
        f"\nPDF 数量：{len(names)}\n"
        f"需要页数：{page_count}（每页最多 23 条）\n"
        f"盖章日期：{stamp_date:%Y-%m-%d}\n"
        f"输出文件：{output_path}"
    )

    if preview:
        print("\n预览完成：未修改模板，也未生成登记表。")
        return len(names), page_count

    write_registration(
        template_path,
        output_path,
        names,
        stamp_date,
    )
    print(f"\n已完成：{output_path}")
    return len(names), page_count


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdf_folder", type=Path, help="要读取的确认书目录")
    parser.add_argument("--template", type=Path, required=True, help="登记表模板路径")
    parser.add_argument("--output", type=Path, required=True, help="登记表输出路径")
    parser.add_argument(
        "--preview",
        action="store_true",
        help="只显示排序、页数和输出文件名，不生成登记表",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        run_registration(
            args.pdf_folder,
            args.template,
            args.output,
            date.today(),
            preview=args.preview,
        )
        return 0
    except (OSError, RuntimeError, ValueError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
