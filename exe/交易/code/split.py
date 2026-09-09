#!/usr/bin/env python3
"""把整本盖章扫描件拆分为独立 PDF。

支持两种固定格式：
- 场外（商品）期权交易确认书：4 页；
- 场外商品远期交易确认书：2 页。

Pan 结算单使用 A 登记表中的原文件名和原文件页数直接拆分，不运行 OCR。

Windows 版使用 RapidOCR、ONNX Runtime、PyMuPDF 和 pypdf，不依赖
macOS PDFKit、Vision、AppKit 或 Swift。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import sys
import tempfile
import unicodedata
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path

from pypdf import PdfReader, PdfWriter

from windows_ocr import OCRLine, recognize_split_pages


OCR_SEPARATOR = r"[\s._·•,，:：/\\\-‐‑‒–—−]*"
TRANSACTION_PATTERN = re.compile(
    r"(?<![0-9A-Z])([0-9OQILSZBGD|]{4})"
    + OCR_SEPARATOR
    + r"([A-Z]"
    + OCR_SEPARATOR
    + r"[A-Z](?:"
    + OCR_SEPARATOR
    + r"[A-Z]"
    + OCR_SEPARATOR
    + r"[A-Z])?)"
    + OCR_SEPARATOR
    + r"([0-9OQILSZBGD|]{10})(?![0-9A-Z])",
    flags=re.IGNORECASE,
)
OCR_DIGIT_TRANSLATION = str.maketrans(
    {
        "O": "0",
        "Q": "0",
        "D": "0",
        "I": "1",
        "L": "1",
        "|": "1",
        "Z": "2",
        "S": "5",
        "G": "6",
        "B": "8",
    }
)


@dataclass(frozen=True)
class PageOCR:
    page: int
    lines: tuple[OCRLine, ...]


@dataclass(frozen=True)
class StartMarker:
    kind: str
    page_count: int
    transaction_number: str
    title_text: str


@dataclass(frozen=True)
class TitleMarker:
    kind: str
    page_count: int
    title_text: str
    line_index: int


@dataclass(frozen=True)
class SplitPlan:
    number: int
    start_page: int
    end_page: int
    kind: str
    transaction_number: str
    filename: str
    footer_confirmed: bool

    @property
    def page_count(self) -> int:
        return self.end_page - self.start_page + 1


@dataclass(frozen=True)
class NamedDocument:
    """无需 OCR 命名时，由 A 功能原文件提供的文件名和固定页数。"""

    filename: str
    page_count: int


@dataclass(frozen=True)
class SplitIssue:
    number: int | None
    start_page: int
    end_page: int
    reason: str
    title_text: str | None = None


class PartialSplitError(RuntimeError):
    """部分 PDF 已安全生成，但仍有页段需要人工处理。"""


def normalize_text(text: str) -> str:
    normalized = unicodedata.normalize("NFKC", text).casefold()
    return "".join(character for character in normalized if character.isalnum())


def source_pdf_for_date(folder: Path, scan_date: date) -> Path:
    """正式运行只选择当天 MMDD.pdf，目录中的其他文件一律不参与。"""
    source = folder / f"{scan_date:%m%d}.pdf"
    if not source.is_file():
        raise RuntimeError(
            f"找不到当天扫描件：{source}\n"
            "确认书扫描目录中的其他 PDF 和文件不会参与本次工作。"
        )
    return source


def resolve_source_pdf(
    scan_folder: Path,
    scan_date: date,
    source_pdf: Path | None,
) -> Path:
    if source_pdf is None:
        return source_pdf_for_date(scan_folder, scan_date)
    candidate = source_pdf.expanduser()
    if not candidate.is_absolute():
        candidate = scan_folder / candidate
    candidate = candidate.resolve()
    if not candidate.is_file() or candidate.suffix.casefold() != ".pdf":
        raise RuntimeError(f"找不到指定的测试扫描 PDF：{candidate}")
    return candidate


def recognize_pages(source: Path) -> list[PageOCR]:
    recognized = recognize_split_pages(source)
    return [
        PageOCR(page=number, lines=lines)
        for number, lines in enumerate(recognized, start=1)
    ]


def ordered_lines(page: PageOCR) -> list[OCRLine]:
    return sorted(page.lines, key=lambda line: (-(line.y + line.height), line.x))


def normalize_trade_type(raw_type: str, expected_type: str | None = None) -> str | None:
    trade_type = re.sub(r"[^A-Z]", "", raw_type.upper())
    if trade_type in {"JY", "FWJY"}:
        return trade_type
    if (
        expected_type in {"JY", "FWJY"}
        and len(trade_type) == len(expected_type)
        and sum(left != right for left, right in zip(trade_type, expected_type)) <= 1
    ):
        return expected_type
    return None


def transaction_number(
    text: str,
    expected_trade_type: str | None = None,
) -> str | None:
    matches = {
        (match.group(1), match.group(2).upper(), match.group(3))
        for match in TRANSACTION_PATTERN.finditer(text)
    }
    if len(matches) != 1:
        return None
    company_code, raw_trade_type, serial = matches.pop()
    company_code = company_code.upper().translate(OCR_DIGIT_TRANSLATION)
    trade_type = normalize_trade_type(raw_trade_type, expected_trade_type)
    serial = serial.upper().translate(OCR_DIGIT_TRANSLATION)
    if not company_code.isdigit() or trade_type is None or not serial.isdigit():
        return None
    return f"【HFSY】{company_code}-{trade_type}-{serial}"


def detect_title(page: PageOCR) -> TitleMarker | None:
    lines = ordered_lines(page)
    for index, line in enumerate(lines):
        if line.y < 0.55:
            continue
        normalized = normalize_text(line.text)
        if "场外" not in normalized or "交易确认书" not in normalized:
            continue
        if "远期" in normalized:
            kind = "场外商品远期交易确认书"
            expected_pages = 2
        elif "期权" in normalized:
            kind = "场外期权交易确认书"
            expected_pages = 4
        else:
            continue
        if line.height < 0.014:
            continue
        return TitleMarker(kind, expected_pages, line.text, index)
    return None


def detect_start(page: PageOCR) -> StartMarker | None:
    lines = ordered_lines(page)
    title = detect_title(page)
    if title is None:
        return None
    title_line = lines[title.line_index]
    nearby = [
        candidate
        for candidate in lines[title.line_index + 1 : title.line_index + 7]
        if candidate.y < title_line.y and title_line.y - candidate.y <= 0.20
    ]
    expected_type = "FWJY" if title.kind == "场外商品远期交易确认书" else "JY"
    number = transaction_number(
        "\n".join(item.text for item in nearby),
        expected_type,
    )
    if number is None:
        return None
    return StartMarker(title.kind, title.page_count, number, title.title_text)


def detect_footer(page: PageOCR) -> bool:
    # 坐标统一为左下角原点，只检查页面右下区域。
    region = [
        line
        for line in page.lines
        if line.y <= 0.35 and line.x + line.width / 2 >= 0.48
    ]
    normalized = normalize_text("\n".join(line.text for line in region))
    return "盖章" in normalized


def analyze_split_plans(
    pages: list[PageOCR],
    expected_count: int,
) -> tuple[list[SplitPlan], list[SplitIssue]]:
    if expected_count < 1:
        raise RuntimeError("确认书文件目录中没有 PDF，无法取得预期确认书数量。")
    if not pages:
        raise RuntimeError("扫描 PDF 没有页面。")

    titles = {
        page.page: marker
        for page in pages
        if (marker := detect_title(page)) is not None
    }
    starts = {
        page.page: marker
        for page in pages
        if (marker := detect_start(page)) is not None
    }
    footers = {page.page for page in pages if detect_footer(page)}

    plans: list[SplitPlan] = []
    issues: list[SplitIssue] = []
    current_page = 1
    document_number = 1
    while current_page <= len(pages):
        title = titles.get(current_page)
        if title is None:
            next_title = min(
                (number for number in titles if number > current_page),
                default=len(pages) + 1,
            )
            issues.append(
                SplitIssue(
                    None,
                    current_page,
                    next_title - 1,
                    "未识别到确认书大标题，无法可靠确定该页段包含几份文件",
                )
            )
            current_page = next_title
            continue

        end_page = current_page + title.page_count - 1
        if end_page > len(pages):
            issues.append(
                SplitIssue(
                    document_number,
                    current_page,
                    len(pages),
                    f"识别为{title.kind}，应有 {title.page_count} 页，但扫描 PDF 已结束",
                    title.title_text,
                )
            )
            document_number += 1
            break

        unexpected_titles = [
            number for number in titles if current_page < number <= end_page
        ]
        if unexpected_titles:
            next_title = min(unexpected_titles)
            issues.append(
                SplitIssue(
                    document_number,
                    current_page,
                    next_title - 1,
                    f"识别为{title.kind}，固定应有 {title.page_count} 页，"
                    f"但第 {next_title} 页又出现新标题",
                    title.title_text,
                )
            )
            document_number += 1
            current_page = next_title
            continue

        marker = starts.get(current_page)
        if marker is None:
            issues.append(
                SplitIssue(
                    document_number,
                    current_page,
                    end_page,
                    "大标题已识别，但其下方交易编号无法通过格式校验",
                    title.title_text,
                )
            )
        else:
            plans.append(
                SplitPlan(
                    document_number,
                    current_page,
                    end_page,
                    marker.kind,
                    marker.transaction_number,
                    f"{document_number}.pdf",
                    end_page in footers,
                )
            )
        document_number += 1
        current_page = end_page + 1

    identified_documents = document_number - 1
    if identified_documents != expected_count:
        issues.append(
            SplitIssue(
                None,
                0,
                0,
                f"根据页面边界识别出 {identified_documents} 份，"
                f"确认书文件目录中有 {expected_count} 份；数量不一致",
            )
        )
    return plans, issues


def format_issue(issue: SplitIssue) -> str:
    if issue.start_page > 0:
        page_range = (
            f"第 {issue.start_page} 页"
            if issue.start_page == issue.end_page
            else f"第 {issue.start_page}-{issue.end_page} 页"
        )
    else:
        page_range = "全局校验"
    number = f"第 {issue.number} 份，" if issue.number is not None else ""
    return f"{number}{page_range}：{issue.reason}"


def build_plans(pages: list[PageOCR], expected_count: int) -> list[SplitPlan]:
    plans, issues = analyze_split_plans(pages, expected_count)
    if issues:
        raise RuntimeError("拆分计划存在异常：\n  " + "\n  ".join(map(format_issue, issues)))
    return plans


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source_file:
        while chunk := source_file.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def print_plans(
    source: Path,
    plans: list[SplitPlan],
    issues: list[SplitIssue],
    output_folder: Path,
    total_pages: int,
) -> None:
    print(f"扫描源文件：{source}")
    print(f"扫描总页数：{total_pages}")
    print(f"可自动拆分数量：{len(plans)}")
    print(f"异常数量：{len(issues)}")
    print(f"拆分输出目录：{output_folder}")
    print("\n拆分计划：")
    for plan in plans:
        verification = (
            "原文件页数已匹配"
            if plan.kind == "按 A 原文件名拆分"
            else "账户首页/资金末页已核对"
            if plan.kind == "持仓报告账户首页/资金末页"
            else f"末页盖章OCR={'已确认' if plan.footer_confirmed else '未识别'}"
        )
        print(
            f"  {plan.number:>3}. 第 {plan.start_page}-{plan.end_page} 页 "
            f"({plan.page_count} 页，{plan.kind}，"
            f"{verification}) "
            f"-> {plan.filename}"
        )
        print(f"       {plan.transaction_number}")
    if not plans:
        print("  无可安全自动拆分的文件。")
    if issues:
        print("\n需要人工处理的异常：")
        for issue in issues:
            print(f"  [跳过] {format_issue(issue)}")
            if issue.title_text:
                print(f"         OCR 标题：{issue.title_text}")


def reusable_output(
    output_folder: Path,
    source: Path,
    source_hash: str,
    expected_count: int,
) -> dict[str, object] | None:
    if not output_folder.exists():
        return None
    if not output_folder.is_dir():
        raise RuntimeError(f"拆分输出路径已存在且不是文件夹：{output_folder}")
    manifest_path = output_folder / "拆分记录.json"
    if not manifest_path.is_file():
        raise RuntimeError(
            f"拆分输出目录已经存在，但缺少拆分记录，无法确认能否安全复用：{output_folder}"
        )
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"无法读取已有拆分记录：{manifest_path}") from exc
    documents = manifest.get("documents", [])
    issues = manifest.get("issues", [])
    pdf_count = sum(
        path.is_file() and path.suffix.casefold() == ".pdf"
        for path in output_folder.iterdir()
    )
    matches = (
        manifest.get("source_name") == source.name
        and manifest.get("source_sha256") == source_hash
        and manifest.get("expected_count") == expected_count
        and isinstance(documents, list)
        and isinstance(issues, list)
        and pdf_count == len(documents)
    )
    if not matches:
        raise RuntimeError(
            f"已有拆分目录与本次源文件或预期数量不一致，未覆盖：{output_folder}"
        )
    return manifest


def write_split_pdfs(
    source: Path,
    output_folder: Path,
    plans: list[SplitPlan],
    issues: list[SplitIssue],
    source_hash: str,
    expected_count: int,
    total_pages: int,
) -> None:
    output_folder.parent.mkdir(parents=True, exist_ok=True)
    temp_path = Path(
        tempfile.mkdtemp(
            prefix=f".{output_folder.name}.split-",
            dir=output_folder.parent,
        )
    )
    completed = False
    try:
        reader = PdfReader(str(source), strict=False)
        for plan in plans:
            writer = PdfWriter()
            for page_index in range(plan.start_page - 1, plan.end_page):
                writer.add_page(reader.pages[page_index])
            target = temp_path / plan.filename
            with target.open("wb") as output_file:
                writer.write(output_file)

        generated = sorted(temp_path.glob("*.pdf"), key=lambda path: path.name)
        if len(generated) != len(plans):
            raise RuntimeError(
                f"实际生成 {len(generated)} 个 PDF，拆分计划为 {len(plans)} 个。"
            )
        manifest = {
            "source_name": source.name,
            "source_size": source.stat().st_size,
            "source_sha256": source_hash,
            "expected_count": expected_count,
            "generated_count": len(plans),
            "status": "complete"
            if not issues and len(plans) == expected_count
            else "partial",
            "total_pages": total_pages,
            "documents": [asdict(plan) for plan in plans],
            "issues": [asdict(issue) for issue in issues],
        }
        (temp_path / "拆分记录.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        if output_folder.exists():
            raise RuntimeError(f"拆分输出目录已存在，未覆盖：{output_folder}")
        temp_path.rename(output_folder)
        completed = True
    finally:
        if not completed and temp_path.exists():
            shutil.rmtree(temp_path)


def run_named_split(
    scan_folder: Path,
    output_folder: Path,
    documents: list[NamedDocument],
    preview: bool = False,
    scan_date: date | None = None,
    source_pdf: Path | None = None,
) -> tuple[int, int]:
    """按 A 功能原文件的顺序、页数和名称拆分，不运行 OCR。"""
    scan_folder = scan_folder.expanduser().resolve()
    output_folder = output_folder.expanduser().resolve()
    if not scan_folder.is_dir():
        raise RuntimeError(f"找不到确认书扫描目录：{scan_folder}")
    if not documents:
        raise RuntimeError("A 批次登记表中没有可用于命名的文件。")

    names: list[str] = []
    seen_names: set[str] = set()
    for document in documents:
        filename = document.filename
        if (
            Path(filename).name != filename
            or Path(filename).suffix.casefold() != ".pdf"
            or filename in {".", ".."}
        ):
            raise RuntimeError(f"A 批次中存在不安全的 PDF 文件名：{filename!r}")
        if document.page_count < 1:
            raise RuntimeError(f"A 批次文件页数无效：{filename}={document.page_count}")
        name_key = filename.casefold()
        if name_key in seen_names:
            raise RuntimeError(f"A 批次中存在重复文件名：{filename}")
        seen_names.add(name_key)
        names.append(filename)

    effective_scan_date = scan_date or date.today()
    source = resolve_source_pdf(scan_folder, effective_scan_date, source_pdf)
    source_hash = file_sha256(source)
    expected_count = len(documents)

    existing_manifest = reusable_output(
        output_folder,
        source,
        source_hash,
        expected_count,
    )
    if existing_manifest is not None:
        manifest_documents = existing_manifest.get("documents", [])
        existing_names = [
            item.get("filename")
            for item in manifest_documents
            if isinstance(item, dict)
        ]
        generated_names = sorted(
            path.name
            for path in output_folder.iterdir()
            if path.is_file() and path.suffix.casefold() == ".pdf"
        )
        if existing_names != names or sorted(names) != generated_names:
            raise RuntimeError(
                f"已有拆分目录的文件名与 A 批次不一致，未覆盖：{output_folder}"
            )
        print(f"已有按原文件名拆分的结果可直接复用：{output_folder}")
        return expected_count, 0

    try:
        reader = PdfReader(str(source), strict=False)
        total_pages = len(reader.pages)
    except Exception as exc:
        raise RuntimeError(f"无法读取扫描 PDF：{source}") from exc

    expected_pages = sum(document.page_count for document in documents)
    if total_pages != expected_pages:
        raise RuntimeError(
            f"扫描 PDF 共 {total_pages} 页，但 A 批次 {expected_count} 份原文件"
            f"合计 {expected_pages} 页；为防止文件名错位，未执行自动拆分。"
        )

    plans: list[SplitPlan] = []
    start_page = 1
    for number, document in enumerate(documents, start=1):
        end_page = start_page + document.page_count - 1
        plans.append(
            SplitPlan(
                number=number,
                start_page=start_page,
                end_page=end_page,
                kind="按 A 原文件名拆分",
                transaction_number=Path(document.filename).stem,
                filename=document.filename,
                footer_confirmed=False,
            )
        )
        start_page = end_page + 1

    print_plans(source, plans, [], output_folder, total_pages)
    if preview:
        print("\n预览完成：文件名和页数匹配，未生成拆分 PDF。")
        return expected_count, total_pages

    write_split_pdfs(
        source,
        output_folder,
        plans,
        [],
        source_hash,
        expected_count,
        total_pages,
    )
    print(f"\n拆分完成：已按 A 原文件名生成 {expected_count} 个 PDF，无需 OCR 重命名。")
    return expected_count, total_pages


def run_split(
    scan_folder: Path,
    output_folder: Path,
    expected_count: int,
    preview: bool = False,
    scan_date: date | None = None,
    source_pdf: Path | None = None,
) -> tuple[int, int]:
    """拆分扫描件，返回（确认书数量，扫描总页数）。"""
    scan_folder = scan_folder.expanduser().resolve()
    output_folder = output_folder.expanduser().resolve()
    if not scan_folder.is_dir():
        raise RuntimeError(f"找不到确认书扫描目录：{scan_folder}")
    effective_scan_date = scan_date or date.today()
    source = resolve_source_pdf(scan_folder, effective_scan_date, source_pdf)
    source_hash = file_sha256(source)

    existing_manifest = reusable_output(
        output_folder,
        source,
        source_hash,
        expected_count,
    )
    if existing_manifest is not None:
        documents = existing_manifest.get("documents", [])
        issues = existing_manifest.get("issues", [])
        status = existing_manifest.get(
            "status",
            "complete" if len(documents) == expected_count else "partial",
        )
        if status == "complete" and len(documents) == expected_count:
            print(f"已有拆分结果与本次扫描件一致，直接复用：{output_folder}")
            return expected_count, 0
        raise PartialSplitError(
            f"已有部分拆分结果可供 continue 使用：{output_folder}\n"
            f"已生成 {len(documents)} 个 PDF，记录 {len(issues)} 项异常。"
        )

    print(f"正在 OCR 识别 {source.name} 的全部页面……")
    pages = recognize_pages(source)
    if preview:
        start_pages = [page.page for page in pages if detect_start(page) is not None]
        footer_pages = [page.page for page in pages if detect_footer(page)]
        print(f"识别到确认书首页：{start_pages}")
        print(f"识别到右下角盖章结束语：{footer_pages}")
    plans, issues = analyze_split_plans(pages, expected_count)
    print_plans(source, plans, issues, output_folder, len(pages))
    if preview:
        if issues:
            raise PartialSplitError(
                f"预览发现 {len(issues)} 项异常；预览模式未生成拆分 PDF。"
            )
        print("\n预览完成：未生成拆分 PDF。")
        return len(plans), len(pages)

    if not plans:
        raise RuntimeError("没有可安全自动拆分的确认书，未生成拆分目录。")
    write_split_pdfs(
        source,
        output_folder,
        plans,
        issues,
        source_hash,
        expected_count,
        len(pages),
    )
    if issues:
        raise PartialSplitError(
            f"部分拆分完成：已生成 {len(plans)} 个正常 PDF，"
            f"跳过 {len(issues)} 项异常。\n"
            f"部分结果保存在：{output_folder}\n"
            "请选择 continue 继续执行 scan.py，或选择 resplit 清空后重试。"
        )
    print(f"\n拆分完成：已生成 {len(plans)} 个 PDF。")
    return len(plans), len(pages)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("scan_folder", type=Path, help="放置整本扫描 PDF 的目录")
    parser.add_argument("output_folder", type=Path, help="拆分 PDF 输出目录")
    parser.add_argument(
        "--expected-count",
        type=int,
        required=True,
        help="确认书文件目录中的原始确认书数量",
    )
    parser.add_argument(
        "--preview",
        action="store_true",
        help="只 OCR 并显示拆分计划，不生成 PDF",
    )
    parser.add_argument(
        "--scan-date",
        type=lambda value: datetime.strptime(value, "%Y%m%d").date(),
        metavar="YYYYMMDD",
        default=date.today(),
        help="扫描日期，用于选择 MMDD.pdf（默认：今天）",
    )
    parser.add_argument(
        "--source-pdf",
        type=Path,
        help="仅用于测试/排障；指定后覆盖 MMDD.pdf 的源文件选择",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        run_split(
            args.scan_folder,
            args.output_folder,
            args.expected_count,
            preview=args.preview,
            scan_date=args.scan_date,
            source_pdf=args.source_pdf,
        )
        return 0
    except PartialSplitError as exc:
        print(f"\n{exc}", file=sys.stderr)
        return 3
    except (OSError, RuntimeError, ValueError) as exc:
        print(f"\n拆分已停止，未生成新的拆分目录：\n{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
