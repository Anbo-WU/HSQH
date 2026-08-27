#!/usr/bin/env python3
"""递归读取“确认书文件”中的 PDF，从登记表模板生成当日登记表。

每个模板页可填 23 条（第 5–27 行）；数量超出时，会在同一张
工作表中向下复制足够的整页模板，并在页之间设置手动分页。
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
from copy import copy
from datetime import date
from pathlib import Path
from typing import Iterator


PDF_FOLDER_NAME = "确认书文件"
TEMPLATE_NAME = "登记表模版.xls"
OUTPUT_PREFIX = "登记表"
LU_SIGNATURE_NAME = "Lu.png"
WU_SIGNATURE_NAME = "Wu.png"

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
APPLICANT_SIGNATURE_COLUMN_INDEX = 6
REVIEWER_SIGNATURE_COLUMN_INDEX = 7

CHECKED_TEXT = "是☑   否☐"
DATE_NUMBER_FORMAT = "yyyy-mm-dd"
SIGNATURE_HEIGHT_SCALE = 0.82
SIGNATURE_RESOLUTION_SCALE = 4

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


def png_dimensions(path: Path) -> tuple[int, int]:
    """不依赖第三方库读取 PNG 宽高。"""
    with path.open("rb") as image_file:
        header = image_file.read(24)
    if len(header) < 24 or header[:8] != b"\x89PNG\r\n\x1a\n":
        raise RuntimeError(f"签名图不是有效的 PNG 文件：{path}")
    width, height = struct.unpack(">II", header[16:24])
    return width, height


def column_width_pixels(sheet, column_index: int) -> int:
    """将 xls 的 1/256 字符列宽近似换算为像素。"""
    column_info = sheet.colinfo_map.get(column_index)
    width_units = column_info.width if column_info is not None else 8 * 256
    return max(1, int(round(width_units / 256 * 7 + 5)))


def row_height_pixels(sheet, row_index: int) -> int:
    """将 xls 的 twip 行高换算为 96 DPI 像素。"""
    row_info = sheet.rowinfo_map.get(row_index)
    height_twips = (
        row_info.height
        if row_info is not None
        else getattr(sheet, "default_row_height", 255)
    )
    return max(1, int(round(height_twips / 15)))


def signature_layout(
    sheet, row_index: int, column_index: int, image_path: Path
) -> tuple[int, int, int, int]:
    """计算签名的显示尺寸和居中偏移，不修改单元格尺寸。"""
    image_width, image_height = png_dimensions(image_path)
    cell_width = column_width_pixels(sheet, column_index)
    cell_height = row_height_pixels(sheet, row_index)
    horizontal_margin = 3
    vertical_margin = 2
    max_width = max(1, cell_width - horizontal_margin * 2)
    max_height = max(1, cell_height - vertical_margin * 2)
    scale = min(max_width / image_width, max_height / image_height)
    display_width = max(1, int(round(image_width * scale)))
    display_height = max(1, int(round(image_height * scale)))
    # 保持原有宽度，仅把签名略微压矮，给上下边框留出更明显的空隙。
    display_height = max(1, int(round(display_height * SIGNATURE_HEIGHT_SCALE)))
    x_offset = max(0, (cell_width - display_width) // 2)
    y_offset = max(0, (cell_height - display_height) // 2)
    return display_width, display_height, x_offset, y_offset


def normalize_bitmap_for_xlwt(bitmap_path: Path) -> None:
    """将 macOS 的负高度 BMP 转为 xlwt 支持的正高度 24 位 BMP。"""
    bitmap = bytearray(bitmap_path.read_bytes())
    if len(bitmap) < 54 or bitmap[:2] != b"BM":
        raise RuntimeError(f"无法生成有效的 BMP 签名图：{bitmap_path}")

    pixel_offset = struct.unpack_from("<I", bitmap, 10)[0]
    width = struct.unpack_from("<i", bitmap, 18)[0]
    height = struct.unpack_from("<i", bitmap, 22)[0]
    bits_per_pixel = struct.unpack_from("<H", bitmap, 28)[0]
    compression = struct.unpack_from("<I", bitmap, 30)[0]
    if bits_per_pixel != 24 or compression != 0:
        raise RuntimeError("签名 BMP 必须是未压缩的 24 位图像。")

    if height < 0:
        absolute_height = -height
        row_size = ((abs(width) * bits_per_pixel + 31) // 32) * 4
        pixel_length = row_size * absolute_height
        pixel_data = bitmap[pixel_offset : pixel_offset + pixel_length]
        rows = [
            pixel_data[row_start : row_start + row_size]
            for row_start in range(0, pixel_length, row_size)
        ]
        bitmap[pixel_offset : pixel_offset + pixel_length] = b"".join(reversed(rows))
        struct.pack_into("<i", bitmap, 22, absolute_height)
        bitmap_path.write_bytes(bitmap)


def prepare_signature_bitmap(
    image_path: Path,
    bitmap_path: Path,
    display_width: int,
    display_height: int,
) -> None:
    """
    将透明 PNG 转换为 xlwt 支持的 24 位 BMP。

    使用 4 倍像素后在 xls 中缩放显示，保持打印清晰度；优先使用
    Pillow，未安装时在 macOS 上自动使用系统 sips。
    """
    resolution_scale = SIGNATURE_RESOLUTION_SCALE
    bitmap_width = max(1, display_width * resolution_scale)
    bitmap_height = max(1, display_height * resolution_scale)

    try:
        from PIL import Image  # type: ignore[import-not-found]
    except ImportError:
        sips_path = shutil.which("sips")
        if sips_path is None:
            raise RuntimeError(
                "生成签名图需要 Pillow。请执行：\n"
                "python3 -m pip install Pillow"
            )

        source_width, source_height = png_dimensions(image_path)
        jpeg_path = bitmap_path.with_suffix(".jpg")
        resized_jpeg_path = bitmap_path.with_name(f"{bitmap_path.stem}-resized.jpg")
        commands = [
            [
                sips_path,
                "-s",
                "format",
                "jpeg",
                "--padColor",
                "FFFFFF",
                "--padToHeightWidth",
                str(source_height),
                str(source_width),
                str(image_path),
                "--out",
                str(jpeg_path),
            ],
            [
                sips_path,
                "-z",
                str(bitmap_height),
                str(bitmap_width),
                str(jpeg_path),
                "--out",
                str(resized_jpeg_path),
            ],
            [
                sips_path,
                "-s",
                "format",
                "bmp",
                str(resized_jpeg_path),
                "--out",
                str(bitmap_path),
            ],
        ]
        for command in commands:
            subprocess.run(command, check=True, capture_output=True, text=True)
    else:
        with Image.open(image_path) as source_image:
            rgba_image = source_image.convert("RGBA")
            white_background = Image.new("RGBA", rgba_image.size, "white")
            flattened = Image.alpha_composite(white_background, rgba_image).convert("RGB")
            resized = flattened.resize(
                (bitmap_width, bitmap_height), Image.Resampling.LANCZOS
            )
            resized.save(bitmap_path, format="BMP")

    normalize_bitmap_for_xlwt(bitmap_path)


def build_picture_object_record(
    sheet,
    row: int,
    column: int,
    width: int,
    height: int,
    x_offset: int,
    y_offset: int,
) -> bytes:
    """生成在登记表白底数据区没有可见边框的 BIFF5/7 位图对象记录。"""
    from xlwt.Bitmap import _position_image  # type: ignore[import-not-found]

    coordinates = _position_image(
        sheet, row, column, x_offset, y_offset, width, height
    )
    col_start, x1, row_start, y1, col_end, x2, row_end, y2 = coordinates

    values = [
        struct.pack("<L", 1),
        struct.pack("<H", 0x0008),
        struct.pack("<H", 0x0001),
        struct.pack("<H", 0x0614),
        struct.pack(
            "<HHHHHHHH",
            col_start,
            x1,
            row_start,
            y1,
            col_end,
            x2,
            row_end,
            y2,
        ),
        struct.pack("<H", 0),
        struct.pack("<L", 0),
        struct.pack("<H", 0),
        struct.pack(
            "<BBBBBBBB",
            0x09,
            0x09,
            0x00,
            0x00,
            0x09,
            0x00,
            0x00,
            0x00,
        ),
        # WPS 会把 BIFF 的 Null line 显示为对象虚框；改用白色 hairline，
        # 在当前白底数据区不显示边线，也不会出现 xlwt 原始的黑色实框。
        struct.pack("<H", 0),
        struct.pack("<L", 0x0009),  # Bitmap
        struct.pack("<HHHH", 0, 0, 0, 1),
        struct.pack("<L", 0),
    ]
    data = b"".join(values)
    return struct.pack("<HH", 0x005D, len(data)) + data


def insert_signature_without_border(
    sheet,
    bitmap_path: Path,
    row: int,
    column: int,
    display_width: int,
    display_height: int,
    x_offset: int,
    y_offset: int,
) -> None:
    """沿用 WPS 兼容的 BMP 图片记录，同时去掉可见的黑色对象框。"""
    from xlwt.Bitmap import ImRawDataBmpRecord  # type: ignore[import-not-found]

    image_record = ImRawDataBmpRecord(bitmap_path.read_bytes()).get()
    object_record = build_picture_object_record(
        sheet,
        row,
        column,
        display_width,
        display_height,
        x_offset,
        y_offset,
    )
    sheet._Worksheet__bmp_rec += object_record + image_record


def write_registration(
    template_path: Path,
    output_path: Path,
    names: list[str],
    stamp_date: date,
    lu_signature_path: Path,
    wu_signature_path: Path,
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

    lu_layout = signature_layout(
        read_sheet,
        start_index,
        APPLICANT_SIGNATURE_COLUMN_INDEX,
        lu_signature_path,
    )
    wu_layout = signature_layout(
        read_sheet,
        start_index,
        REVIEWER_SIGNATURE_COLUMN_INDEX,
        wu_signature_path,
    )

    with tempfile.TemporaryDirectory(prefix="registration-signatures-") as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        lu_bitmap = temp_dir / "Lu.bmp"
        wu_bitmap = temp_dir / "Wu.bmp"
        prepare_signature_bitmap(
            lu_signature_path, lu_bitmap, lu_layout[0], lu_layout[1]
        )
        prepare_signature_bitmap(
            wu_signature_path, wu_bitmap, wu_layout[0], wu_layout[1]
        )

        # 只填充存在真实 PDF 数据的行；I 列继续留空。
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

            insert_signature_without_border(
                write_sheet,
                lu_bitmap,
                row_index,
                APPLICANT_SIGNATURE_COLUMN_INDEX,
                lu_layout[0],
                lu_layout[1],
                lu_layout[2],
                lu_layout[3],
            )
            insert_signature_without_border(
                write_sheet,
                wu_bitmap,
                row_index,
                REVIEWER_SIGNATURE_COLUMN_INDEX,
                wu_layout[0],
                wu_layout[1],
                wu_layout[2],
                wu_layout[3],
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


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--preview",
        action="store_true",
        help="只显示排序、页数和输出文件名，不生成登记表",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    base_folder = Path(__file__).resolve().parent
    pdf_folder = base_folder / PDF_FOLDER_NAME
    template_path = base_folder / TEMPLATE_NAME
    lu_signature_path = base_folder / LU_SIGNATURE_NAME
    wu_signature_path = base_folder / WU_SIGNATURE_NAME
    today = date.today()
    output_path = base_folder / f"{OUTPUT_PREFIX}{today:%Y%m%d}.xls"

    required_paths = [
        (pdf_folder, "PDF 文件夹"),
        (template_path, "登记表模板"),
        (lu_signature_path, "Lu 签名图"),
        (wu_signature_path, "Wu 签名图"),
    ]
    for path, description in required_paths:
        if not path.exists():
            print(f"找不到{description}：{path}", file=sys.stderr)
            return 1

    names = collect_names(pdf_folder)
    if not names:
        print(f"未在 {pdf_folder} 及其子文件夹中找到 PDF。", file=sys.stderr)
        return 1

    page_count = required_page_count(len(names))
    for number, name in enumerate(names, start=1):
        print(f"{number:>3}. {name}")

    print(
        f"\nPDF 数量：{len(names)}\n"
        f"需要页数：{page_count}（每页最多 23 条）\n"
        f"盖章日期：{today:%Y-%m-%d}\n"
        f"输出文件：{output_path.name}"
    )

    if args.preview:
        print("\n预览完成：未修改模板，也未生成登记表。")
        return 0

    write_registration(
        template_path,
        output_path,
        names,
        today,
        lu_signature_path,
        wu_signature_path,
    )
    print(f"\n已完成：{output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
