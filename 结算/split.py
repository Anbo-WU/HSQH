#!/usr/bin/env python3
"""按累计远期到期表中的编号，拆分当日的远期了结记录导出表。

默认目录结构：
    forward/
    ├── 累计远期到期/累计远期到期MMDD.xls
    ├── 远期镒链导出/远期了结记录导出YYYYMMDD.xls
    └── 老版本/远期了结记录导出01.xls ...

累计到期表是二进制 .xls，需要 xlrd；远期了结记录导出表是 HTML
格式的 .xls。输出仍保留为 HTML 格式 .xls，可供后续 update.py 处理。
正式运行前会先将老版本的旧文件压缩归档，校验 zip 成功后再清空。
"""

from __future__ import annotations

import argparse
import html
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from workflow_utils import archive_and_clear

try:
    import xlrd
except ImportError:  # 让缺少依赖时的提示比 Python traceback 更直接。
    xlrd = None  # type: ignore[assignment]


IDENTIFIER_RE = re.compile(r"【[^\r\n】]+】[A-Za-z0-9]+-JY-\d+")
ROW_RE = re.compile(r"<tr\b[^>]*>.*?</tr\s*>", re.IGNORECASE | re.DOTALL)
CELL_RE = re.compile(
    r"<(td|th)\b[^>]*>.*?</\1\s*>", re.IGNORECASE | re.DOTALL
)
CELL_CONTENT_RE = re.compile(
    r"<(?:td|th)\b[^>]*>(.*?)</(?:td|th)\s*>",
    re.IGNORECASE | re.DOTALL,
)
TABLE_RE = re.compile(r"<table\b[^>]*>.*?</table\s*>", re.IGNORECASE | re.DOTALL)


@dataclass(frozen=True)
class DueItem:
    excel_row: int
    identifier: str | None
    c_text: str
    e_text: str


@dataclass(frozen=True)
class HtmlTable:
    source_text: str
    table_start: int
    table_end: int
    table_text: str
    rows: list[str]
    row_start: int
    row_end: int


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
            and (candidate / "远期镒链导出").is_dir()
        ):
            return candidate

    locations = "\n  ".join(str(path) for path in checked)
    raise FileNotFoundError(
        "未找到同时包含“累计远期到期”和“远期镒链导出”"
        f"的母文件夹。已检查：\n  {locations}"
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
        raise RuntimeError(f"{directory} 中匹配到多个文件，无法确定使用哪个：{names}")
    return matches[0]


def extract_identifier(value: object) -> str | None:
    """从两种单元格文本中取统一的编号。

    支持：
      【HFSY】0112-JY-2026051201转远期明细表
      累计期权:【HFSY】0112-JY-2026051201
    """
    text = str(value).strip()
    match = IDENTIFIER_RE.search(text)
    return match.group(0).strip() if match else None


def read_due_items(path: Path) -> tuple[list[DueItem], str]:
    if xlrd is None:
        raise RuntimeError("缺少 xlrd，请先执行：python3 -m pip install xlrd==1.2.0")

    try:
        workbook = xlrd.open_workbook(str(path), on_demand=True)
    except Exception as exc:
        raise RuntimeError(f"无法读取累计到期表 {path}：{exc}") from exc

    try:
        usable_sheets = [sheet for sheet in workbook.sheets() if sheet.ncols >= 5]
        if not usable_sheets:
            raise RuntimeError("工作簿中没有至少 5 列的工作表")

        # 如果有多个 sheet，选 C/E 列中有效编号数量最多的一张。
        sheet = max(
            usable_sheets,
            key=lambda sh: sum(
                bool(extract_identifier(sh.cell_value(row, 2)))
                or bool(extract_identifier(sh.cell_value(row, 4)))
                for row in range(sh.nrows)
            ),
        )

        items: list[DueItem] = []
        for row_index in range(sheet.nrows):
            c_text = str(sheet.cell_value(row_index, 2)).strip()
            e_text = str(sheet.cell_value(row_index, 4)).strip()
            c_identifier = extract_identifier(c_text)
            e_identifier = extract_identifier(e_text)

            # 表头和纯空行不是待拆分数据。若 C/E 中任一格像编号，
            # 或任一格包含“转远期明细表”，则保留该行以便报警。
            looks_like_data = bool(c_identifier or e_identifier) or (
                "转远期明细表" in c_text or "转远期明细表" in e_text
            )
            if not looks_like_data:
                continue

            items.append(
                DueItem(
                    excel_row=row_index + 1,
                    identifier=c_identifier,  # 无论 C/E 是否一致，始终以 C 列为准。
                    c_text=c_text,
                    e_text=e_text,
                )
            )
    finally:
        workbook.release_resources()

    if not items:
        raise RuntimeError(f"{path.name} 的 C/E 列中没有找到待处理的编号")
    return items, sheet.name


def decode_html_xls(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            text = raw.decode(encoding)
        except UnicodeDecodeError:
            continue
        if "<html" in text[:2000].lower() or "<table" in text[:5000].lower():
            return text
    raise RuntimeError(f"{path} 不是本程序支持的 HTML 格式 .xls 文件")


def parse_html_table(path: Path) -> HtmlTable:
    source_text = decode_html_xls(path)
    table_match = TABLE_RE.search(source_text)
    if not table_match:
        raise RuntimeError(f"{path.name} 中未找到表格")

    table_text = table_match.group(0)
    row_matches = list(ROW_RE.finditer(table_text))
    if not row_matches:
        raise RuntimeError(f"{path.name} 的表格中未找到任何行")

    rows = [match.group(0) for match in row_matches]
    header_cells = CELL_RE.findall(rows[0])
    if len(header_cells) < 30:
        raise RuntimeError(
            f"{path.name} 表头只有 {len(header_cells)} 列，不足 A–AD 的 30 列"
        )

    return HtmlTable(
        source_text=source_text,
        table_start=table_match.start(),
        table_end=table_match.end(),
        table_text=table_text,
        rows=rows,
        row_start=row_matches[0].start(),
        row_end=row_matches[-1].end(),
    )


def cell_text(cell_html: str) -> str:
    match = CELL_CONTENT_RE.fullmatch(cell_html.strip())
    if not match:
        return ""
    without_tags = re.sub(r"<[^>]+>", "", match.group(1))
    return html.unescape(without_tags).strip()


def row_cells(row_html: str) -> list[str]:
    return [match.group(0) for match in CELL_RE.finditer(row_html)]


def limit_row_to_ad(row_html: str) -> str:
    cells = row_cells(row_html)
    if len(cells) < 30:
        raise ValueError(f"该行只有 {len(cells)} 列")
    if len(cells) == 30:
        return row_html
    open_tag = re.match(r"<tr\b[^>]*>", row_html, re.IGNORECASE | re.DOTALL)
    if not open_tag:
        raise ValueError("无法识别 <tr> 行标签")
    return open_tag.group(0) + "".join(cells[:30]) + "</tr>"


def build_match_index(table: HtmlTable, source_name: str) -> dict[str, list[str]]:
    matches: dict[str, list[str]] = {}
    malformed_rows = 0

    for excel_row, row_html in enumerate(table.rows[1:], start=2):
        cells = row_cells(row_html)
        # 某些导出表末尾会有仅保存格式的空行，可直接忽略。
        if not cells or not any(cell_text(cell) for cell in cells):
            continue
        if len(cells) < 30:
            malformed_rows += 1
            print(
                f"警告：{source_name} 第 {excel_row} 行只有 {len(cells)} 列，"
                "已跳过该行。"
            )
            continue

        identifier = extract_identifier(cell_text(cells[28]))  # AC 列：备注
        if identifier is None:
            continue  # 导出大表中与累计期权无关的废数据。
        matches.setdefault(identifier, []).append(limit_row_to_ad(row_html))

    if malformed_rows:
        print(f"警告：共跳过 {malformed_rows} 条列数不足 A–AD 的数据。")
    return matches


def make_output_text(table: HtmlTable, selected_rows: list[str]) -> str:
    header = limit_row_to_ad(table.rows[0])
    replacement_rows = header + "\n" + "\n".join(selected_rows)
    new_table = (
        table.table_text[: table.row_start]
        + replacement_rows
        + table.table_text[table.row_end :]
    )
    return (
        table.source_text[: table.table_start]
        + new_table
        + table.source_text[table.table_end :]
    )


def write_outputs(
    items: list[DueItem],
    match_index: dict[str, list[str]],
    table: HtmlTable,
    output_dir: Path,
    dry_run: bool,
) -> tuple[int, int]:
    width = max(2, len(str(len(items))))
    mismatch_count = 0
    no_match_count = 0

    if not dry_run:
        output_dir.mkdir(parents=True, exist_ok=True)

    for sequence, item in enumerate(items, start=1):
        c_identifier = extract_identifier(item.c_text)
        e_identifier = extract_identifier(item.e_text)
        identifiers_differ = c_identifier != e_identifier
        malformed_texts_differ = (
            c_identifier is None
            and e_identifier is None
            and item.c_text != item.e_text
        )
        if identifiers_differ or malformed_texts_differ:
            mismatch_count += 1
            print(
                f"警告：累计到期表第 {item.excel_row} 行 C/E 编号不一致："
                f"C={c_identifier or item.c_text or '<空>'}；"
                f"E={e_identifier or item.e_text or '<空>'}。按 C 列处理。"
            )

        selected_rows = match_index.get(item.identifier, []) if item.identifier else []
        filename = f"远期了结记录导出{sequence:0{width}d}.xls"
        destination = output_dir / filename

        if not selected_rows:
            no_match_count += 1
            shown = item.identifier or item.c_text or "<C 列为空>"
            print(
                f"警告：编号 {shown} 在远期了结记录导出表中"
                f"未找到匹配；{filename} 将只包含表头。"
            )

        if dry_run:
            print(f"预览：{filename} <- {len(selected_rows)} 条匹配数据")
        else:
            destination.write_text(make_output_text(table, selected_rows), encoding="utf-8")
            print(f"生成：{destination} <- {len(selected_rows)} 条匹配数据")

    return mismatch_count, no_match_count


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="按累计远期到期表 C 列编号拆分当日远期了结记录"
    )
    parser.add_argument(
        "--date",
        help="指定处理日期（YYYYMMDD），默认为上海时区当天",
    )
    parser.add_argument(
        "--base-dir",
        type=Path,
        help="母文件夹，其下应有“累计远期到期”和“远期镒链导出”",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="输出目录，默认为母文件夹下的“老版本”",
    )
    parser.add_argument(
        "--archive-dir",
        type=Path,
        help="历史压缩包目录，默认为脚本同级的 archive",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只检查匹配结果，不归档、不清理、不写入文件",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        processing_date = parse_date(args.date)
        date_yyyymmdd = processing_date.strftime("%Y%m%d")
        date_mmdd = processing_date.strftime("%m%d")
        business_root = find_business_root(args.base_dir)

        due_dir = business_root / "累计远期到期"
        export_dir = business_root / "远期镒链导出"
        output_dir = (
            args.output_dir.expanduser().resolve()
            if args.output_dir
            else business_root / "老版本"
        )
        archive_dir = (
            args.archive_dir.expanduser().resolve()
            if args.archive_dir
            else Path(__file__).resolve().parent / "archive"
        )

        project_dir = Path(__file__).resolve().parent
        protected_dirs = {
            Path("/").resolve(),
            Path.home().resolve(),
            project_dir.resolve(),
            business_root.resolve(),
            due_dir.resolve(),
            export_dir.resolve(),
            archive_dir.resolve(),
        }
        if output_dir.resolve() in protected_dirs:
            raise ValueError(f"输出目录不能是业务根目录或源数据目录：{output_dir}")

        due_path = find_one_file(
            due_dir,
            [f"累计远期到期{date_mmdd}.xls", f"累计文件到期{date_mmdd}.xls"],
            f"*到期{date_mmdd}.xls",
        )
        export_path = find_one_file(
            export_dir,
            [f"远期了结记录导出{date_yyyymmdd}.xls"],
            f"远期了结记录导出*{date_yyyymmdd}*.xls",
        )

        print(f"累计到期表：{due_path}")
        print(f"了结记录表：{export_path}")
        print(f"输出目录：{output_dir}")

        items, sheet_name = read_due_items(due_path)
        print(f"使用工作表：{sheet_name}；待拆分数据：{len(items)} 条")

        table = parse_html_table(export_path)
        match_index = build_match_index(table, export_path.name)
        indexed_row_count = sum(len(rows) for rows in match_index.values())
        print(
            f"导出表可匹配编号：{len(match_index)} 个；"
            f"已索引数据：{indexed_row_count} 条"
        )

        if not args.dry_run:
            archived = archive_and_clear(
                output_dir, archive_dir, "老版本", date_yyyymmdd
            )
            if archived is None:
                print(f"清理：{output_dir} 没有旧文件，无需归档。")
            else:
                print(
                    f"归档：已将 {archived.file_count} 个旧文件压缩至 "
                    f"{archived.path}，并清空老版本目录。"
                )

        mismatch_count, no_match_count = write_outputs(
            items, match_index, table, output_dir, args.dry_run
        )
        action = "预览" if args.dry_run else "生成"
        print(
            f"处理完成：{action} {len(items)} 个文件；"
            f"C/E 不一致 {mismatch_count} 条；零匹配 {no_match_count} 条。"
        )
        return 0
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
