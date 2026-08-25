#!/usr/bin/env python3
"""识别“确认书扫描”中每份 PDF 首页的交易编号，并按规则重命名。

例：
【HFSY】0009-FWJY-2026072401
-> 【HFSY】0009-FWJY-202607240120260724.pdf

脚本使用 macOS 自带的 PDFKit 和 Vision OCR，不需要安装 Python 第三方库。
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import uuid
from dataclasses import dataclass
from pathlib import Path


# 不依赖 OCR 必须准确识别书名号和 HFSY。Vision 偶尔会把 FWJY
# 识别成 FW.JY，因此允许编号各固定字段之间混入少量分隔标点。
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
    {"O": "0", "Q": "0", "D": "0", "I": "1", "L": "1", "|": "1", "Z": "2", "S": "5", "G": "6", "B": "8"}
)


SWIFT_OCR_SOURCE = r'''
import Foundation
import PDFKit
import Vision
import AppKit

func recognize(_ path: String) -> [String: Any] {
    let url = URL(fileURLWithPath: path)
    guard let document = PDFDocument(url: url), let page = document.page(at: 0) else {
        return ["path": path, "error": "无法打开 PDF 或 PDF 没有首页"]
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
        return ["path": path, "error": "无法渲染 PDF 首页"]
    }

    do {
        let request = VNRecognizeTextRequest()
        request.recognitionLevel = .accurate
        request.recognitionLanguages = ["zh-Hans", "en-US"]
        request.usesLanguageCorrection = true
        try VNImageRequestHandler(cgImage: cgImage, options: [:]).perform([request])

        // 交易编号固定在首页大标题下，只取页面最上方的识别结果。
        var lines: [String] = []
        for observation in (request.results ?? []).prefix(10) {
            if let candidate = observation.topCandidates(1).first {
                lines.append(candidate.string)
            }
        }
        return ["path": path, "text": lines.joined(separator: "\n")]
    } catch {
        return ["path": path, "error": error.localizedDescription]
    }
}

var output: [[String: Any]] = []
for path in CommandLine.arguments.dropFirst() {
    autoreleasepool {
        output.append(recognize(path))
    }
}

do {
    let data = try JSONSerialization.data(withJSONObject: output)
    FileHandle.standardOutput.write(data)
} catch {
    fputs("JSON output failed: \(error)\n", stderr)
    exit(4)
}
'''


@dataclass(frozen=True)
class RenamePlan:
    source: Path
    target: Path
    transaction_number: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("folder", type=Path, help="要识别并重命名的确认书扫描目录")
    parser.add_argument(
        "--preview",
        action="store_true",
        help="只识别并显示新旧文件名，不执行重命名",
    )
    return parser.parse_args()


def collect_pdfs(folder: Path) -> list[Path]:
    return sorted(
        (
            path
            for path in folder.rglob("*")
            if path.is_file() and path.suffix.casefold() == ".pdf"
        ),
        key=lambda path: str(path.relative_to(folder)).casefold(),
    )


def compile_ocr_helper(temp_folder: Path) -> Path:
    swiftc = shutil.which("swiftc")
    if swiftc is None:
        raise RuntimeError(
            "找不到 swiftc。请先安装 macOS Command Line Tools："
            "xcode-select --install"
        )

    source_path = temp_folder / "pdf_first_page_ocr.swift"
    executable_path = temp_folder / "pdf_first_page_ocr"
    source_path.write_text(SWIFT_OCR_SOURCE, encoding="utf-8")

    environment = os.environ.copy()
    environment["SWIFT_MODULECACHE_PATH"] = str(temp_folder / "swift-cache")
    environment["CLANG_MODULE_CACHE_PATH"] = str(temp_folder / "clang-cache")
    command = [
        swiftc,
        "-framework",
        "PDFKit",
        "-framework",
        "Vision",
        "-framework",
        "AppKit",
        str(source_path),
        "-o",
        str(executable_path),
    ]
    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        env=environment,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(f"OCR 组件编译失败：\n{result.stderr.strip()}")
    return executable_path


def recognize_all(pdfs: list[Path]) -> dict[Path, str]:
    with tempfile.TemporaryDirectory(prefix="confirm_ocr_") as temp_name:
        helper = compile_ocr_helper(Path(temp_name))
        command = [str(helper), *(str(path.resolve()) for path in pdfs)]
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"OCR 运行失败（退出码 {result.returncode}）：\n"
                f"{result.stderr.strip()}"
            )
        try:
            records = json.loads(result.stdout)
        except json.JSONDecodeError as exc:
            raise RuntimeError("OCR 未返回有效结果。") from exc

    recognized: dict[Path, str] = {}
    errors: list[str] = []
    for record in records:
        path = Path(record["path"])
        if "error" in record:
            errors.append(f"{path.name}: {record['error']}")
        else:
            recognized[path] = record.get("text", "")
    if errors:
        raise RuntimeError("OCR 失败：\n  " + "\n  ".join(errors))
    return recognized


def normalize_text(text: str) -> str:
    normalized = unicodedata.normalize("NFKC", text).casefold()
    return "".join(character for character in normalized if character.isalnum())


def expected_trade_type(text: str) -> str | None:
    normalized = normalize_text(text)
    if "远期交易确认书" in normalized:
        return "FWJY"
    if "期权交易确认书" in normalized:
        return "JY"
    return None


def normalize_trade_type(raw_type: str, expected_type: str | None) -> str | None:
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


def extract_transaction_number(text: str, pdf_name: str) -> str:
    matches = {
        (match.group(1), match.group(2).upper(), match.group(3))
        for match in TRANSACTION_PATTERN.finditer(text)
    }
    if len(matches) != 1:
        detail = "未识别到" if not matches else "识别到多个候选编号"
        raise ValueError(f"{pdf_name}: {detail}")

    company_code, raw_trade_type, serial = matches.pop()
    company_code = company_code.upper().translate(OCR_DIGIT_TRANSLATION)
    trade_type = normalize_trade_type(raw_trade_type, expected_trade_type(text))
    serial = serial.upper().translate(OCR_DIGIT_TRANSLATION)
    if (
        not company_code.isdigit()
        or trade_type is None
        or not serial.isdigit()
    ):
        raise ValueError(f"{pdf_name}: 交易编号包含无法安全纠正的字符")
    return f"【HFSY】{company_code}-{trade_type}-{serial}"


def build_plans(pdfs: list[Path], recognized: dict[Path, str]) -> list[RenamePlan]:
    plans: list[RenamePlan] = []
    errors: list[str] = []

    for source in pdfs:
        try:
            transaction_number = extract_transaction_number(
                recognized[source.resolve()], source.name
            )
            # 编号末尾 10 位是 YYYYMMDDNN，去掉最后 2 位得到 YYYYMMDD。
            serial = transaction_number.rsplit("-", maxsplit=1)[1]
            date_suffix = serial[-10:-2]
            target = source.with_name(f"{transaction_number}{date_suffix}.pdf")
            plans.append(RenamePlan(source, target, transaction_number))
        except (KeyError, ValueError) as exc:
            errors.append(str(exc))

    if errors:
        raise RuntimeError("编号提取失败：\n  " + "\n  ".join(errors))

    target_keys: dict[str, Path] = {}
    for plan in plans:
        key = str(plan.target.resolve()).casefold()
        if key in target_keys and target_keys[key] != plan.source:
            raise RuntimeError(
                f"新文件名重复：{plan.target.name}\n"
                f"  {target_keys[key].name}\n  {plan.source.name}"
            )
        target_keys[key] = plan.source

    source_keys = {str(plan.source.resolve()).casefold() for plan in plans}
    for plan in plans:
        target_key = str(plan.target.resolve()).casefold()
        if plan.target.exists() and target_key not in source_keys:
            raise RuntimeError(f"目标文件已存在：{plan.target}")

    return plans


def apply_plans(plans: list[RenamePlan]) -> None:
    """分两阶段改名，避免文件名互相占用；失败时尽量回滚。"""
    active = [plan for plan in plans if plan.source != plan.target]
    temporary: dict[RenamePlan, Path] = {}
    completed: list[RenamePlan] = []

    try:
        for plan in active:
            temp_path = plan.source.with_name(
                f".{plan.source.name}.rename-{uuid.uuid4().hex}.tmp"
            )
            plan.source.rename(temp_path)
            temporary[plan] = temp_path

        for plan in active:
            temporary[plan].rename(plan.target)
            completed.append(plan)
    except Exception:
        for plan in reversed(completed):
            if plan.target.exists() and not plan.source.exists():
                plan.target.rename(plan.source)
        for plan, temp_path in temporary.items():
            if temp_path.exists() and not plan.source.exists():
                temp_path.rename(plan.source)
        raise


def run_scan(pdf_folder: Path, preview: bool = False) -> tuple[int, int]:
    """识别并重命名扫描件，返回（PDF 总数，实际改名数量）。"""
    pdf_folder = pdf_folder.expanduser().resolve()
    if sys.platform != "darwin":
        raise RuntimeError("本脚本使用 macOS Vision OCR，需要在 Mac 上运行。")
    if not pdf_folder.is_dir():
        raise RuntimeError(f"找不到文件夹：{pdf_folder}")

    pdfs = collect_pdfs(pdf_folder)
    if not pdfs:
        raise RuntimeError(f"未在 {pdf_folder} 中找到 PDF。")

    print(f"正在识别 {len(pdfs)} 份 PDF 的首页，请稍候……")
    recognized = recognize_all(pdfs)
    plans = build_plans(pdfs, recognized)

    print()
    for number, plan in enumerate(plans, start=1):
        status = "（已是目标名称）" if plan.source == plan.target else ""
        print(f"{number:>3}. {plan.source.name}")
        print(f"     -> {plan.target.name}{status}")

    if preview:
        print(f"\n预览完成：共 {len(plans)} 份，未重命名任何文件。")
        return len(plans), 0

    apply_plans(plans)

    changed = sum(plan.source != plan.target for plan in plans)
    print(f"\n已完成：成功重命名 {changed} 份 PDF，PDF 内容未修改。")
    return len(plans), changed


def main() -> int:
    args = parse_args()
    try:
        run_scan(args.folder, preview=args.preview)
        return 0
    except (OSError, RuntimeError, ValueError) as exc:
        print(f"\n已停止，未重命名任何文件：\n{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
