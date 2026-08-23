#!/usr/bin/env python3
"""OCR 识别整本盖章扫描件，并按确认书边界拆分为独立 PDF。

支持两种固定格式：
- 场外（商品）期权交易确认书：4 页；
- 场外商品远期交易确认书：2 页。

脚本使用 macOS PDFKit 和 Vision，不依赖 Python 第三方库。标题、交易编号、
固定页数、下一份首页位置和总份数用于硬性校验；末页右下角盖章文字作为
辅助 OCR 证据记录，避免红章遮挡文字时把正确扫描件误判为失败。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path


TRANSACTION_PATTERN = re.compile(
    r"(?<![0-9A-Z])([0-9OQILSZBGD|]{4})\s*[-‐‑‒–—−]?\s*"
    r"(FWJY|JY)\s*[-‐‑‒–—−]?\s*([0-9OQILSZBGD|]{10})(?![0-9A-Z])",
    flags=re.IGNORECASE,
)

OCR_DIGIT_TRANSLATION = str.maketrans(
    {"O": "0", "Q": "0", "D": "0", "I": "1", "L": "1", "|": "1", "Z": "2", "S": "5", "G": "6", "B": "8"}
)


SWIFT_OCR_SOURCE = r'''
import Foundation
import PDFKit
import Vision
import AppKit

func fail(_ message: String, code: Int32 = 1) -> Never {
    FileHandle.standardError.write((message + "\n").data(using: .utf8)!)
    exit(code)
}

let arguments = Array(CommandLine.arguments.dropFirst())
guard arguments.count == 2 else {
    fail("用法：pdf_all_pages_ocr 输入.pdf 输出.json", code: 2)
}

let inputURL = URL(fileURLWithPath: arguments[0])
let outputURL = URL(fileURLWithPath: arguments[1])
guard let document = PDFDocument(url: inputURL), document.pageCount > 0 else {
    fail("无法打开 PDF 或 PDF 没有页面：\(inputURL.path)", code: 3)
}

var output: [[String: Any]] = []
for pageIndex in 0..<document.pageCount {
    autoreleasepool {
        guard let page = document.page(at: pageIndex) else {
            output.append(["page": pageIndex + 1, "error": "无法读取页面"])
            return
        }

        let bounds = page.bounds(for: .mediaBox)
        let width: CGFloat = 3000
        let height = width * bounds.height / bounds.width
        let image = page.thumbnail(
            of: NSSize(width: width, height: height),
            for: .mediaBox
        )
        var rect = NSRect(origin: .zero, size: image.size)
        guard let cgImage = image.cgImage(
            forProposedRect: &rect,
            context: nil,
            hints: nil
        ) else {
            output.append(["page": pageIndex + 1, "error": "无法渲染页面"])
            return
        }

        do {
            let request = VNRecognizeTextRequest()
            request.recognitionLevel = .accurate
            request.recognitionLanguages = ["zh-Hans", "en-US"]
            request.usesLanguageCorrection = true
            request.customWords = ["徽丰实业（上海）有限公司", "盖章"]
            try VNImageRequestHandler(cgImage: cgImage, options: [:]).perform([request])

            var lines: [[String: Any]] = []
            for observation in request.results ?? [] {
                guard let candidate = observation.topCandidates(1).first else {
                    continue
                }
                let box = observation.boundingBox
                lines.append([
                    "text": candidate.string,
                    "confidence": Double(candidate.confidence),
                    "x": Double(box.origin.x),
                    "y": Double(box.origin.y),
                    "width": Double(box.size.width),
                    "height": Double(box.size.height),
                ])
            }
            output.append(["page": pageIndex + 1, "lines": lines])
        } catch {
            output.append(["page": pageIndex + 1, "error": error.localizedDescription])
        }
    }

    if pageIndex == 0 || (pageIndex + 1) % 5 == 0 || pageIndex + 1 == document.pageCount {
        fputs("OCR 进度：\(pageIndex + 1)/\(document.pageCount)\n", stderr)
        fflush(stderr)
    }
}

do {
    let data = try JSONSerialization.data(withJSONObject: output)
    try data.write(to: outputURL, options: .atomic)
} catch {
    fail("无法写入 OCR 结果：\(error.localizedDescription)", code: 4)
}
'''


SWIFT_SPLIT_SOURCE = r'''
import Foundation
import PDFKit

func fail(_ message: String, code: Int32 = 1) -> Never {
    FileHandle.standardError.write((message + "\n").data(using: .utf8)!)
    exit(code)
}

let arguments = Array(CommandLine.arguments.dropFirst())
guard arguments.count >= 5 && (arguments.count - 2) % 3 == 0 else {
    fail("用法：pdf_splitter 输入.pdf 输出目录 起始页 页数 文件名 [...]", code: 2)
}

let inputURL = URL(fileURLWithPath: arguments[0])
let outputFolder = URL(fileURLWithPath: arguments[1], isDirectory: true)
guard let source = PDFDocument(url: inputURL), source.pageCount > 0 else {
    fail("无法打开源 PDF：\(inputURL.path)", code: 3)
}

var position = 2
while position < arguments.count {
    guard let start = Int(arguments[position]),
          let count = Int(arguments[position + 1]),
          start >= 0,
          count > 0,
          start + count <= source.pageCount else {
        fail("无效的拆分页码参数", code: 4)
    }
    let filename = arguments[position + 2]
    let output = PDFDocument()
    for pageIndex in start..<(start + count) {
        guard let page = source.page(at: pageIndex),
              let copiedPage = page.copy() as? PDFPage else {
            fail("无法复制第 \(pageIndex + 1) 页", code: 5)
        }
        output.insert(copiedPage, at: output.pageCount)
    }
    let target = outputFolder.appendingPathComponent(filename)
    guard output.write(to: target) else {
        fail("无法写入：\(target.path)", code: 6)
    }
    position += 3
}
'''


@dataclass(frozen=True)
class OCRLine:
    text: str
    confidence: float
    x: float
    y: float
    width: float
    height: float


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


def normalize_text(text: str) -> str:
    normalized = unicodedata.normalize("NFKC", text).casefold()
    return "".join(character for character in normalized if character.isalnum())


def source_pdf_for_date(folder: Path, scan_date: date) -> Path:
    """只选择当天 MMDD.pdf；目录中的其他文件一律不参与。"""
    source = folder / f"{scan_date:%m%d}.pdf"
    if not source.is_file():
        raise RuntimeError(
            f"找不到当天扫描件：{source}\n"
            "确认书扫描目录中的其他 PDF 和文件不会参与本次工作。"
        )
    return source


def compile_swift_helper(
    temp_folder: Path,
    source_name: str,
    executable_name: str,
    source_text: str,
    frameworks: tuple[str, ...],
) -> Path:
    swiftc = shutil.which("swiftc")
    if swiftc is None:
        raise RuntimeError(
            "找不到 swiftc。请先安装 macOS Command Line Tools：xcode-select --install"
        )

    source_path = temp_folder / source_name
    executable_path = temp_folder / executable_name
    source_path.write_text(source_text, encoding="utf-8")
    command = [swiftc]
    for framework in frameworks:
        command.extend(("-framework", framework))
    command.extend((str(source_path), "-o", str(executable_path)))

    environment = os.environ.copy()
    environment["SWIFT_MODULECACHE_PATH"] = str(temp_folder / "swift-cache")
    environment["CLANG_MODULE_CACHE_PATH"] = str(temp_folder / "clang-cache")
    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        env=environment,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(f"Swift 组件编译失败：\n{result.stderr.strip()}")
    return executable_path


def recognize_pages(source: Path) -> list[PageOCR]:
    if sys.platform != "darwin":
        raise RuntimeError("split.py 使用 macOS Vision OCR，需要在 Mac 上运行。")

    with tempfile.TemporaryDirectory(prefix="confirm_split_ocr_") as temp_name:
        temp_folder = Path(temp_name)
        helper = compile_swift_helper(
            temp_folder,
            "pdf_all_pages_ocr.swift",
            "pdf_all_pages_ocr",
            SWIFT_OCR_SOURCE,
            ("PDFKit", "Vision", "AppKit"),
        )
        json_path = temp_folder / "ocr.json"
        result = subprocess.run(
            (str(helper), str(source.resolve()), str(json_path)),
            check=False,
        )
        if result.returncode != 0:
            raise RuntimeError(f"OCR 运行失败（退出码 {result.returncode}）。")
        try:
            records = json.loads(json_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise RuntimeError("OCR 未返回有效结果。") from exc

    pages: list[PageOCR] = []
    errors: list[str] = []
    for record in records:
        page_number = int(record["page"])
        if "error" in record:
            errors.append(f"第 {page_number} 页：{record['error']}")
            continue
        lines = tuple(
            OCRLine(
                text=str(line.get("text", "")),
                confidence=float(line.get("confidence", 0)),
                x=float(line.get("x", 0)),
                y=float(line.get("y", 0)),
                width=float(line.get("width", 0)),
                height=float(line.get("height", 0)),
            )
            for line in record.get("lines", [])
        )
        pages.append(PageOCR(page_number, lines))
    if errors:
        raise RuntimeError("部分页面 OCR 失败：\n  " + "\n  ".join(errors))
    pages.sort(key=lambda page: page.page)
    if [page.page for page in pages] != list(range(1, len(pages) + 1)):
        raise RuntimeError("OCR 返回的页码不连续。")
    return pages


def ordered_lines(page: PageOCR) -> list[OCRLine]:
    return sorted(page.lines, key=lambda line: (-(line.y + line.height), line.x))


def transaction_number(text: str) -> str | None:
    matches = {
        (match.group(1), match.group(2).upper(), match.group(3))
        for match in TRANSACTION_PATTERN.finditer(text)
    }
    if len(matches) != 1:
        return None
    company_code, trade_type, serial = matches.pop()
    company_code = company_code.upper().translate(OCR_DIGIT_TRANSLATION)
    serial = serial.upper().translate(OCR_DIGIT_TRANSLATION)
    if not (company_code.isdigit() and serial.isdigit()):
        return None
    return f"【HFSY】{company_code}-{trade_type}-{serial}"


def detect_start(page: PageOCR) -> StartMarker | None:
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

        # 标题应显著大于同页上半部分普通文字。
        if line.height < 0.014:
            continue

        # 交易编号必须紧跟在标题之后，并且位于标题下方。
        nearby = [
            candidate
            for candidate in lines[index + 1 : index + 7]
            if candidate.y < line.y and line.y - candidate.y <= 0.20
        ]
        number = transaction_number("\n".join(item.text for item in nearby))
        if number is None:
            continue
        return StartMarker(kind, expected_pages, number, line.text)
    return None


def detect_footer(page: PageOCR) -> bool:
    # Vision 坐标从页面左下角开始；只检查页面右下区域。
    region = [
        line
        for line in page.lines
        if line.y <= 0.35 and line.x + line.width / 2 >= 0.48
    ]
    normalized = normalize_text("\n".join(line.text for line in region))
    # 真实盖章件中红章可能完全遮住公司名，Vision 仍能稳定识别下一行
    # “盖章：”。位置、固定页数和下一份首页会共同约束边界。
    return "盖章" in normalized


def build_plans(pages: list[PageOCR], expected_count: int) -> list[SplitPlan]:
    if expected_count < 1:
        raise RuntimeError("确认书文件目录中没有 PDF，无法取得预期确认书数量。")
    if not pages:
        raise RuntimeError("扫描 PDF 没有页面。")

    starts = {
        page.page: marker
        for page in pages
        if (marker := detect_start(page)) is not None
    }
    footers = {page.page for page in pages if detect_footer(page)}
    if 1 not in starts:
        raise RuntimeError("第 1 页未同时识别到确认书大标题和其下方交易编号。")

    plans: list[SplitPlan] = []
    current_page = 1
    while current_page <= len(pages):
        marker = starts.get(current_page)
        if marker is None:
            raise RuntimeError(f"第 {current_page} 页未识别为确认书首页。")

        end_page = current_page + marker.page_count - 1
        if end_page > len(pages):
            raise RuntimeError(
                f"第 {current_page} 页识别为{marker.kind}，应有 {marker.page_count} 页，"
                "但扫描 PDF 已结束。"
            )

        unexpected_starts = [
            page for page in starts if current_page < page <= end_page
        ]
        if unexpected_starts:
            raise RuntimeError(
                f"第 {current_page} 页开始的确认书固定应有 {marker.page_count} 页，"
                f"但第 {min(unexpected_starts)} 页又识别到新标题。"
            )

        number = len(plans) + 1
        plans.append(
            SplitPlan(
                number=number,
                start_page=current_page,
                end_page=end_page,
                kind=marker.kind,
                transaction_number=marker.transaction_number,
                filename=f"{number}.pdf",
                footer_confirmed=end_page in footers,
            )
        )
        current_page = end_page + 1

    if len(plans) != expected_count:
        raise RuntimeError(
            f"拆分识别出 {len(plans)} 份确认书，但确认书文件目录中有 "
            f"{expected_count} 份；数量不一致。"
        )
    return plans


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source_file:
        while chunk := source_file.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def print_plans(source: Path, plans: list[SplitPlan], output_folder: Path) -> None:
    print(f"扫描源文件：{source}")
    print(f"扫描总页数：{plans[-1].end_page if plans else 0}")
    print(f"确认书数量：{len(plans)}")
    print(f"拆分输出目录：{output_folder}")
    print("\n拆分计划：")
    for plan in plans:
        print(
            f"  {plan.number:>3}. 第 {plan.start_page}-{plan.end_page} 页 "
            f"({plan.page_count} 页，{plan.kind}，"
            f"末页盖章OCR={'已确认' if plan.footer_confirmed else '未识别'}) "
            f"-> {plan.filename}"
        )
        print(f"       {plan.transaction_number}")


def reusable_output(
    output_folder: Path,
    source: Path,
    source_hash: str,
    expected_count: int,
) -> bool:
    if not output_folder.exists():
        return False
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
    pdf_count = sum(
        path.is_file() and path.suffix.casefold() == ".pdf"
        for path in output_folder.iterdir()
    )
    matches = (
        manifest.get("source_name") == source.name
        and manifest.get("source_sha256") == source_hash
        and manifest.get("expected_count") == expected_count
        and len(manifest.get("documents", [])) == expected_count
        and pdf_count == expected_count
    )
    if not matches:
        raise RuntimeError(
            f"已有拆分目录与本次源文件或预期数量不一致，未覆盖：{output_folder}"
        )
    return True


def write_split_pdfs(
    source: Path,
    output_folder: Path,
    plans: list[SplitPlan],
    source_hash: str,
    expected_count: int,
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
        with tempfile.TemporaryDirectory(prefix="confirm_split_helper_") as helper_name:
            helper_folder = Path(helper_name)
            helper = compile_swift_helper(
                helper_folder,
                "pdf_splitter.swift",
                "pdf_splitter",
                SWIFT_SPLIT_SOURCE,
                ("PDFKit",),
            )
            command = [str(helper), str(source.resolve()), str(temp_path.resolve())]
            for plan in plans:
                command.extend(
                    (
                        str(plan.start_page - 1),
                        str(plan.page_count),
                        plan.filename,
                    )
                )
            result = subprocess.run(command, capture_output=True, text=True, check=False)
            if result.returncode != 0:
                raise RuntimeError(result.stderr.strip() or "PDF 拆分失败。")

        generated = sorted(temp_path.glob("*.pdf"), key=lambda path: path.name)
        if len(generated) != expected_count:
            raise RuntimeError(
                f"实际生成 {len(generated)} 个 PDF，预期 {expected_count} 个。"
            )
        manifest = {
            "source_name": source.name,
            "source_size": source.stat().st_size,
            "source_sha256": source_hash,
            "expected_count": expected_count,
            "total_pages": plans[-1].end_page,
            "documents": [asdict(plan) for plan in plans],
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


def run_split(
    scan_folder: Path,
    output_folder: Path,
    expected_count: int,
    preview: bool = False,
    scan_date: date | None = None,
) -> tuple[int, int]:
    """拆分扫描件，返回（确认书数量，扫描总页数）。"""
    scan_folder = scan_folder.expanduser().resolve()
    output_folder = output_folder.expanduser().resolve()
    if not scan_folder.is_dir():
        raise RuntimeError(f"找不到确认书扫描目录：{scan_folder}")
    effective_scan_date = scan_date or date.today()
    source = source_pdf_for_date(scan_folder, effective_scan_date)
    source_hash = file_sha256(source)

    if reusable_output(output_folder, source, source_hash, expected_count):
        print(f"已有拆分结果与本次扫描件一致，直接复用：{output_folder}")
        return expected_count, 0

    print(f"正在 OCR 识别 {source.name} 的全部页面……")
    pages = recognize_pages(source)
    if preview:
        start_pages = [page.page for page in pages if detect_start(page) is not None]
        footer_pages = [page.page for page in pages if detect_footer(page)]
        print(f"识别到确认书首页：{start_pages}")
        print(f"识别到右下角盖章结束语：{footer_pages}")
    plans = build_plans(pages, expected_count)
    print_plans(source, plans, output_folder)
    if preview:
        print("\n预览完成：未生成拆分 PDF。")
        return len(plans), len(pages)

    write_split_pdfs(source, output_folder, plans, source_hash, expected_count)
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
        )
        return 0
    except (OSError, RuntimeError, ValueError) as exc:
        print(f"\n拆分已停止，未生成新的拆分目录：\n{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
