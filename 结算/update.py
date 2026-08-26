#!/usr/bin/env python3
"""批量改版老版本远期了结记录，校验客户收益，并打包交付。"""

from __future__ import annotations

import argparse
import html
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from zoneinfo import ZoneInfo

from workflow_utils import archive_and_clear, create_archive

try:
    import xlrd
except ImportError:
    xlrd = None  # type: ignore[assignment]


DELETE_COLUMNS = {2, 5, 23, 24, 25, 26, 27, 29}  # C/F/X/Y/Z/AA/AB/AD
SUM_COLUMNS = {14, 20, 21}  # O/U/V（均为删除列之前的列号）
INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
IDENTIFIER_RE = re.compile(r"【[^\r\n】]+】[A-Za-z0-9]+-JY-\d+")
ROW_RE = re.compile(r"<tr\b[^>]*>.*?</tr\s*>", re.I | re.S)
CELL_RE = re.compile(r"<(td|th)\b([^>]*)>(.*?)</\1\s*>", re.I | re.S)


@dataclass(frozen=True)
class TransformResult:
    source: Path
    destination: Path
    identifier: str


def parse_date(value: str | None) -> datetime:
    if value is None:
        return datetime.now(ZoneInfo("Asia/Shanghai"))
    try:
        return datetime.strptime(value, "%Y%m%d").replace(
            tzinfo=ZoneInfo("Asia/Shanghai")
        )
    except ValueError as exc:
        raise ValueError("--date 必须是 YYYYMMDD 格式，例如 20260826") from exc


def find_business_root(explicit: Path | None) -> Path:
    if explicit is not None:
        candidates = [explicit.expanduser().resolve()]
    else:
        script_dir = Path(__file__).resolve().parent
        current_dir = Path.cwd().resolve()
        candidates = [
            script_dir,
            current_dir,
            script_dir / "forward",
            current_dir / "forward",
        ]

    checked: list[Path] = []
    for candidate in candidates:
        candidate = candidate.resolve()
        if candidate in checked:
            continue
        checked.append(candidate)
        if (
            (candidate / "累计远期到期").is_dir()
            and (candidate / "老版本").is_dir()
            and (candidate / "新版本").is_dir()
        ):
            return candidate

    locations = "\n  ".join(str(path) for path in checked)
    raise FileNotFoundError(
        "未找到同时包含“累计远期到期”、“老版本”和“新版本”"
        f"的业务母文件夹。已检查：\n  {locations}"
    )


def find_one_file(directory: Path, preferred_names: list[str], fallback: str) -> Path:
    for name in preferred_names:
        path = directory / name
        if path.is_file():
            return path
    matches = sorted(path for path in directory.glob(fallback) if path.is_file())
    if not matches:
        expected = " 或 ".join(preferred_names)
        raise FileNotFoundError(f"未在 {directory} 找到 {expected}")
    if len(matches) > 1:
        names = "、".join(path.name for path in matches)
        raise RuntimeError(f"{directory} 中匹配到多个文件：{names}")
    return matches[0]


def extract_identifier(value: object) -> str | None:
    match = IDENTIFIER_RE.search(str(value).strip())
    return match.group(0).strip() if match else None


def cell_text(cell_html: str) -> str:
    """取得单元格的可见文字。"""
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
    if value == 0:
        value = abs(value)  # 避免输出负零。
    return f"{value:.{places}f}"


def split_row(row_html: str) -> list[str]:
    return [match.group(0) for match in CELL_RE.finditer(row_html)]


def make_row(cells: list[str]) -> str:
    return "<tr>" + "".join(cells) + "</tr>"


def safe_name_part(value: str) -> str:
    return INVALID_FILENAME_CHARS.sub("_", value).strip().rstrip(". ")


def transform(source: Path, output_dir: Path) -> TransformResult:
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

    parsed: list[list[str]] = []
    for row in rows:
        cells = split_row(row)
        if cells and any(cell_text(cell) for cell in cells):
            parsed.append(cells)
    if len(parsed) < 2:
        raise ValueError("表格没有有效数据行")
    if any(len(row) < 30 for row in parsed):
        raise ValueError("表格列数不足 30 列，无法按需求处理")

    headers, data_rows = parsed[0], parsed[1:]
    identifiers = {
        identifier
        for row in data_rows
        if (identifier := extract_identifier(cell_text(row[28]))) is not None
    }
    if not identifiers:
        raise ValueError("AC 列中未找到累计远期标志符")
    if len(identifiers) != 1:
        raise ValueError(f"AC 列存在多个不同标志符：{sorted(identifiers)}")
    identifier = next(iter(identifiers))

    headers[20] = replace_cell_text(headers[20], "客户了结远期收益")

    # U 列切换为客户视角：数值乘以 -1，并保留源数据小数位。
    for row_number, row in enumerate(data_rows, start=2):
        original = cell_text(row[20])
        changed = -decimal_value(original, location=f"{source.name} U{row_number}")
        row[20] = replace_cell_text(
            row[20], format_decimal(changed, decimal_places(original))
        )

    totals: dict[int, Decimal] = {}
    total_places: dict[int, int] = {}
    for column in SUM_COLUMNS:
        values = [cell_text(row[column]) for row in data_rows]
        totals[column] = sum(
            (
                decimal_value(value, location=f"{source.name} 第{column + 1}列")
                for value in values
            ),
            Decimal("0"),
        )
        total_places[column] = max(
            (decimal_places(value) for value in values), default=2
        )

    # O、V 合计不一致时，用 V 列逐行覆盖 O 列。
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
    kept_rows = [
        [cell for index, cell in enumerate(row) if index not in DELETE_COLUMNS]
        for row in all_rows
    ]
    new_table = "<table><thead>" + make_row(kept_rows[0]) + "</thead><tbody>"
    new_table += "".join(make_row(row) for row in kept_rows[1:]) + "</tbody></table>"
    result_text = text[: table_match.start()] + new_table + text[table_match.end() :]

    customer = safe_name_part(cell_text(parsed[1][1])[:4])
    reference = safe_name_part(identifier[-10:])
    if not customer or not reference:
        raise ValueError("B2 或 AC2 内容不足，无法生成新文件名")
    output_dir.mkdir(parents=True, exist_ok=True)
    destination = output_dir / f"{customer}{reference}{source.stem}{source.suffix}"
    destination.write_text(result_text, encoding="utf-8")
    return TransformResult(source, destination, identifier)


def read_final_s_total(path: Path) -> Decimal:
    """从改版完成的新版本文件最后一行读取 S 列合计。"""
    text = path.read_text(encoding="utf-8-sig")
    table_match = re.search(r"<table\b[^>]*>.*?</table\s*>", text, re.I | re.S)
    if not table_match:
        raise ValueError("未找到表格")
    rows = []
    for row_html in ROW_RE.findall(table_match.group(0)):
        cells = split_row(row_html)
        if cells and any(cell_text(cell) for cell in cells):
            rows.append(cells)
    if not rows or len(rows[-1]) < 19:
        raise ValueError("最后一行列数不足，无法读取 S 列")
    if cell_text(rows[-1][0]) != "合计":
        raise ValueError("最后一行不是合计行")
    return decimal_value(cell_text(rows[-1][18]), location=f"{path.name} 合计行 S 列")


def read_company_totals(path: Path) -> tuple[dict[str, Decimal], str]:
    """按 C 列标志符建立累计到期表 I 列数值索引。"""
    if xlrd is None:
        raise RuntimeError("缺少 xlrd，请先执行：python3 -m pip install xlrd==1.2.0")
    try:
        workbook = xlrd.open_workbook(str(path), on_demand=True)
    except Exception as exc:
        raise RuntimeError(f"无法读取累计到期表 {path}：{exc}") from exc

    try:
        usable_sheets = [sheet for sheet in workbook.sheets() if sheet.ncols >= 9]
        if not usable_sheets:
            raise RuntimeError("工作簿中没有至少 9 列的工作表")
        sheet = max(
            usable_sheets,
            key=lambda sh: sum(
                bool(extract_identifier(sh.cell_value(row, 2)))
                for row in range(sh.nrows)
            ),
        )
        totals: dict[str, Decimal] = {}
        for row_index in range(sheet.nrows):
            identifier = extract_identifier(sheet.cell_value(row_index, 2))
            if identifier is None:
                continue
            raw_value = str(sheet.cell_value(row_index, 8)).strip()
            if not raw_value:
                raise ValueError(f"{path.name} I{row_index + 1} 为空")
            value = decimal_value(raw_value, location=f"{path.name} I{row_index + 1}")
            if identifier in totals:
                raise ValueError(f"累计到期表 C 列存在重复标志符：{identifier}")
            totals[identifier] = value
        sheet_name = sheet.name
    finally:
        workbook.release_resources()

    if not totals:
        raise ValueError(f"{path.name} 中未找到可用的 C/I 列数据")
    return totals, sheet_name


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="批量改版远期了结记录、校验收益并打包新版本"
    )
    parser.add_argument("--date", help="处理日期 YYYYMMDD，默认为上海时区当天")
    parser.add_argument("--base-dir", type=Path, help="forward 业务母文件夹")
    parser.add_argument("--input-dir", type=Path, help="老版本输入目录")
    parser.add_argument("--output-dir", type=Path, help="新版本输出目录")
    parser.add_argument(
        "--archive-dir",
        type=Path,
        help="历史压缩包目录，默认为脚本同级 archive",
    )
    parser.add_argument(
        "--desktop-dir",
        type=Path,
        help="最终压缩包交付目录，默认为当前用户桌面",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        processing_date = parse_date(args.date)
        date_yyyymmdd = processing_date.strftime("%Y%m%d")
        date_mmdd = processing_date.strftime("%m%d")
        project_dir = Path(__file__).resolve().parent
        business_root = find_business_root(args.base_dir)
        input_dir = (
            args.input_dir.expanduser().resolve()
            if args.input_dir
            else business_root / "老版本"
        )
        output_dir = (
            args.output_dir.expanduser().resolve()
            if args.output_dir
            else business_root / "新版本"
        )
        archive_dir = (
            args.archive_dir.expanduser().resolve()
            if args.archive_dir
            else project_dir / "archive"
        )
        desktop_dir = (
            args.desktop_dir.expanduser().resolve()
            if args.desktop_dir
            else Path.home() / "Desktop"
        )

        if input_dir.resolve() == output_dir.resolve():
            raise ValueError("老版本输入目录和新版本输出目录不能相同")
        protected_dirs = {
            Path("/").resolve(),
            Path.home().resolve(),
            project_dir.resolve(),
            business_root.resolve(),
            (business_root / "累计远期到期").resolve(),
            input_dir.resolve(),
            archive_dir.resolve(),
        }
        if output_dir.resolve() in protected_dirs:
            raise ValueError(f"新版本输出目录不安全：{output_dir}")

        sources = sorted(path for path in input_dir.glob("*.xls") if path.is_file())
        if not sources:
            raise FileNotFoundError(f"未在 {input_dir} 找到 .xls 文件")

        due_dir = business_root / "累计远期到期"
        due_path = find_one_file(
            due_dir,
            [f"累计远期到期{date_mmdd}.xls", f"累计文件到期{date_mmdd}.xls"],
            f"*到期{date_mmdd}.xls",
        )
        company_totals, sheet_name = read_company_totals(due_path)
        print(f"老版本目录：{input_dir}；待处理 {len(sources)} 个文件")
        print(f"累计到期表：{due_path}（{sheet_name}）")
        print(f"新版本目录：{output_dir}")

        archived = archive_and_clear(
            output_dir, archive_dir, "新版本", date_yyyymmdd
        )
        if archived is None:
            print(f"清理：{output_dir} 没有旧文件，无需归档。")
        else:
            print(
                f"归档：已将 {archived.file_count} 个旧文件压缩至 "
                f"{archived.path}，并清空新版本目录。"
            )

        failures = 0
        verification_issues = 0
        successes: list[TransformResult] = []
        for source in sources:
            try:
                result = transform(source, output_dir)
                successes.append(result)
                print(f"成功：{source.name} -> {result.destination}")
            except Exception as exc:
                failures += 1
                print(f"失败：{source.name}：{exc}")

        # 改版完成后，从新版本实际文件的最后一行 S 列取数校验。
        for result in successes:
            try:
                client_total = read_final_s_total(result.destination)
                if result.identifier not in company_totals:
                    raise ValueError(
                        f"累计到期表 C 列未找到标志符 {result.identifier}"
                    )
                company_total = company_totals[result.identifier]
                if client_total != -company_total:
                    difference = client_total + company_total
                    raise ValueError(
                        f"收益方向/金额不一致：新版本 S 列客户合计="
                        f"{format(client_total, 'f')}；累计表 I 列公司金额="
                        f"{format(company_total, 'f')}；两者之和={format(difference, 'f')}"
                    )
                print(
                    f"校验通过：{result.destination.name}；{result.identifier}；"
                    f"客户={format(client_total, 'f')}，公司={format(company_total, 'f')}"
                )
            except Exception as exc:
                verification_issues += 1
                print(
                    f"收益校验异常：{result.destination.name}；"
                    f"{result.identifier}；{exc}"
                )

        package = create_archive(
            output_dir, desktop_dir, "新版本", date_yyyymmdd
        )
        if package is None:
            print("交付压缩包未生成：新版本目录中没有文件。")
        else:
            print(
                f"交付压缩包：{package.path}（{package.file_count} 个文件）"
            )

        print(
            f"处理完成：成功 {len(successes)} 个，失败 {failures} 个；"
            f"收益校验异常 {verification_issues} 个。"
        )
        return 1 if failures or verification_issues else 0
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
