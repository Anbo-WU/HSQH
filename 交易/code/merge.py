#!/usr/bin/env python3
"""递归扫描指定确认书文件夹，转换 Word 文档并按自然顺序合并全部 PDF。

Word 转 PDF 使用 LibreOffice 的无界面模式；macOS 上的 PDF 合并优先使用
系统 PDFKit，因此通常不需要额外安装 Python 包。
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Iterable


WORD_SUFFIXES = {".doc", ".docx"}


SWIFT_PDF_MERGER = r'''
import Foundation
import PDFKit

func fail(_ message: String, code: Int32 = 1) -> Never {
    FileHandle.standardError.write((message + "\n").data(using: .utf8)!)
    exit(code)
}

let arguments = Array(CommandLine.arguments.dropFirst())
guard arguments.count >= 2 else {
    fail("用法：pdf_merger 输出.pdf 输入1.pdf [输入2.pdf ...]", code: 2)
}

let outputURL = URL(fileURLWithPath: arguments[0])
let merged = PDFDocument()
var outputPageIndex = 0

for inputPath in arguments.dropFirst() {
    let inputURL = URL(fileURLWithPath: inputPath)
    guard let document = PDFDocument(url: inputURL) else {
        fail("无法打开 PDF：\(inputPath)", code: 3)
    }
    guard document.pageCount > 0 else {
        fail("PDF 没有页面：\(inputPath)", code: 4)
    }
    for pageIndex in 0..<document.pageCount {
        guard let page = document.page(at: pageIndex) else {
            fail("无法读取 PDF 第 \(pageIndex + 1) 页：\(inputPath)", code: 5)
        }
        merged.insert(page, at: outputPageIndex)
        outputPageIndex += 1
    }
}

guard merged.write(to: outputURL) else {
    fail("无法写入合并后的 PDF：\(outputURL.path)", code: 6)
}

print(outputPageIndex)
'''


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "folder",
        type=Path,
        help="要扫描的确认书目录",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        required=True,
        help="合并结果路径",
    )
    parser.add_argument(
        "--overwrite-word-pdf",
        action="store_true",
        help="Word 的同名 PDF 已存在时，用新转换的文件覆盖它",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只扫描并显示转换计划、合并顺序，不转换或合并",
    )
    return parser.parse_args()


def natural_tokens(text: str) -> tuple[tuple[int, object], ...]:
    """将文本拆成字符和数字，以便 8.7 排在 8.10 前面。"""
    return tuple(
        (0, int(part)) if part.isdigit() else (1, part.casefold())
        for part in re.split(r"(\d+)", text)
        if part
    )


def natural_path_key(path: Path, root: Path) -> tuple[tuple[tuple[int, object], ...], ...]:
    relative = path.relative_to(root)
    return tuple(natural_tokens(part) for part in relative.parts)


def scan_files(root: Path) -> list[Path]:
    return sorted(
        (path for path in root.rglob("*") if path.is_file()),
        key=lambda path: natural_path_key(path, root),
    )


def same_path(left: Path, right: Path) -> bool:
    return os.path.normcase(str(left.resolve())) == os.path.normcase(str(right.resolve()))


def print_scan_report(files: list[Path], root: Path, output: Path) -> tuple[list[Path], list[Path]]:
    pdfs: list[Path] = []
    word_files: list[Path] = []
    non_pdfs: list[Path] = []

    for path in files:
        suffix = path.suffix.casefold()
        if suffix == ".pdf":
            if not same_path(path, output):
                pdfs.append(path)
        else:
            non_pdfs.append(path)
            if suffix in WORD_SUFFIXES:
                word_files.append(path)

    print(f"扫描目录：{root}")
    print(f"共发现 {len(files)} 个文件：{len(pdfs)} 个待合并 PDF，{len(non_pdfs)} 个非 PDF。")
    if non_pdfs:
        print("\n发现以下非 PDF 文件：")
        for path in non_pdfs:
            kind = "Word，待转换" if path.suffix.casefold() in WORD_SUFFIXES else "其他文件"
            print(f"  [{kind}] {path.relative_to(root)}")
    else:
        print("\n没有发现非 PDF 文件。")

    return pdfs, word_files


def find_libreoffice() -> Path | None:
    for command in ("libreoffice", "soffice"):
        found = shutil.which(command)
        if found:
            return Path(found)

    mac_app_binary = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    return mac_app_binary if mac_app_binary.is_file() else None


def check_word_targets(word_files: list[Path]) -> None:
    targets: dict[str, Path] = {}
    for source in word_files:
        target = source.with_suffix(".pdf")
        key = os.path.normcase(str(target.resolve()))
        previous = targets.get(key)
        if previous is not None:
            raise RuntimeError(
                "多个 Word 文件会生成同一个 PDF，无法确定应保留哪个：\n"
                f"  {previous}\n  {source}\n  -> {target}"
            )
        targets[key] = source


def convert_one_word(
    source: Path,
    target: Path,
    libreoffice: Path,
    overwrite: bool,
) -> str:
    if target.exists() and not overwrite:
        return "同名 PDF 已存在，直接使用"

    with tempfile.TemporaryDirectory(
        prefix=".word_to_pdf_", dir=target.parent
    ) as temp_name:
        temp_folder = Path(temp_name)
        profile_folder = temp_folder / "libreoffice_profile"
        output_folder = temp_folder / "output"
        output_folder.mkdir()

        command = [
            str(libreoffice),
            "--headless",
            f"-env:UserInstallation={profile_folder.resolve().as_uri()}",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            str(output_folder),
            str(source.resolve()),
        ]
        result = subprocess.run(command, capture_output=True, text=True, check=False)
        converted = output_folder / f"{source.stem}.pdf"
        if result.returncode != 0 or not converted.is_file():
            detail = (result.stderr or result.stdout).strip() or "LibreOffice 未生成 PDF"
            raise RuntimeError(detail)
        with converted.open("rb") as converted_file:
            signature = converted_file.read(5)
        if converted.stat().st_size < 5 or signature != b"%PDF-":
            raise RuntimeError("转换结果不是有效的 PDF 文件")

        # 转换先在临时目录完成，成功后才替换目标，避免留下半成品。
        os.replace(converted, target)
    return "转换完成"


def convert_word_files(
    word_files: list[Path],
    root: Path,
    overwrite: bool,
) -> None:
    if not word_files:
        print("\n没有 Word 文件需要转换。")
        return

    check_word_targets(word_files)
    needs_conversion = [
        source for source in word_files if overwrite or not source.with_suffix(".pdf").exists()
    ]
    libreoffice = find_libreoffice() if needs_conversion else None
    if needs_conversion and libreoffice is None:
        raise RuntimeError(
            "发现需要转换的 Word 文件，但找不到 LibreOffice。\n"
            "请先安装 LibreOffice（macOS 可运行：brew install --cask libreoffice），然后重试。\n"
            "已安装的 WPS Mac 版没有官方稳定的批量转换命令行接口，"
            "因此脚本不使用容易误操作的界面自动点击。"
        )

    print("\n开始处理 Word 文件：")
    failures: list[str] = []
    for source in word_files:
        target = source.with_suffix(".pdf")
        try:
            assert libreoffice is not None or target.exists()
            status = convert_one_word(source, target, libreoffice, overwrite)  # type: ignore[arg-type]
            print(f"  [成功] {source.relative_to(root)} -> {target.name}（{status}）")
        except Exception as exc:
            failures.append(f"{source.relative_to(root)}：{exc}")
            print(f"  [失败] {source.relative_to(root)}：{exc}")

    if failures:
        raise RuntimeError(
            "部分 Word 文件转换失败；为保证文件完整，本次没有开始合并：\n  "
            + "\n  ".join(failures)
        )


def planned_pdfs(current_pdfs: list[Path], word_files: list[Path], root: Path) -> list[Path]:
    by_path = {os.path.normcase(str(path.resolve())): path for path in current_pdfs}
    for source in word_files:
        target = source.with_suffix(".pdf")
        by_path[os.path.normcase(str(target.resolve()))] = target
    return sorted(by_path.values(), key=lambda path: natural_path_key(path, root))


def merge_with_pdfkit(pdfs: list[Path], temporary_output: Path) -> int:
    swiftc = shutil.which("swiftc")
    if swiftc is None or sys.platform != "darwin":
        raise FileNotFoundError("当前系统不能使用 macOS PDFKit")

    helper_folder = temporary_output.parent / "pdfkit_helper"
    helper_folder.mkdir()
    source = helper_folder / "pdf_merger.swift"
    executable = helper_folder / "pdf_merger"
    source.write_text(SWIFT_PDF_MERGER, encoding="utf-8")

    environment = os.environ.copy()
    environment["SWIFT_MODULECACHE_PATH"] = str(helper_folder / "swift-cache")
    environment["CLANG_MODULE_CACHE_PATH"] = str(helper_folder / "clang-cache")
    compile_result = subprocess.run(
        [swiftc, "-framework", "PDFKit", str(source), "-o", str(executable)],
        capture_output=True,
        text=True,
        env=environment,
        check=False,
    )
    if compile_result.returncode != 0:
        raise RuntimeError(f"PDFKit 合并组件编译失败：{compile_result.stderr.strip()}")

    merge_result = subprocess.run(
        [str(executable), str(temporary_output), *(str(path.resolve()) for path in pdfs)],
        capture_output=True,
        text=True,
        check=False,
    )
    if merge_result.returncode != 0:
        raise RuntimeError(merge_result.stderr.strip() or "PDFKit 合并失败")
    return int(merge_result.stdout.strip())


def merge_with_python(pdfs: list[Path], temporary_output: Path) -> int:
    try:
        from pypdf import PdfReader, PdfWriter  # type: ignore[import-not-found]
    except ImportError:
        try:
            from PyPDF2 import PdfReader, PdfWriter  # type: ignore[import-not-found,no-redef]
        except ImportError as exc:
            raise RuntimeError(
                "无法使用 macOS PDFKit，且未安装 pypdf。请运行：python3 -m pip install pypdf"
            ) from exc

    writer = PdfWriter()
    page_count = 0
    try:
        for path in pdfs:
            reader = PdfReader(str(path))
            if reader.is_encrypted:
                try:
                    reader.decrypt("")
                except Exception as exc:
                    raise RuntimeError(f"PDF 已加密，无法读取：{path}") from exc
            for page in reader.pages:
                writer.add_page(page)
                page_count += 1
        with temporary_output.open("wb") as output_file:
            writer.write(output_file)
    finally:
        close = getattr(writer, "close", None)
        if callable(close):
            close()
    return page_count


def merge_pdfs(pdfs: list[Path], output: Path) -> int:
    if not pdfs:
        raise RuntimeError("没有找到可合并的 PDF 文件。")

    output.parent.mkdir(parents=True, exist_ok=True)
    # 临时结果与正式结果放在同一文件系统，最后可原子替换。
    with tempfile.TemporaryDirectory(prefix=".merge_pdf_", dir=output.parent) as temp_name:
        temporary_output = Path(temp_name) / "merged.pdf"
        try:
            page_count = merge_with_pdfkit(pdfs, temporary_output)
        except FileNotFoundError:
            page_count = merge_with_python(pdfs, temporary_output)

        if not temporary_output.is_file() or temporary_output.stat().st_size == 0:
            raise RuntimeError("合并程序没有生成有效的输出文件。")
        os.replace(temporary_output, output)
    return page_count


def print_merge_order(pdfs: Iterable[Path], root: Path) -> None:
    print("\nPDF 合并顺序：")
    for number, path in enumerate(pdfs, start=1):
        print(f"  {number:>3}. {path.relative_to(root)}")


def run_merge(
    root: Path,
    output: Path,
    overwrite_word_pdf: bool = False,
    dry_run: bool = False,
) -> tuple[int, int]:
    """转换并合并确认书，返回（PDF 数量，合并页数）。"""
    root = root.expanduser().resolve()
    output = output.expanduser().resolve()
    if not root.is_dir():
        raise RuntimeError(f"扫描目录不存在或不是文件夹：{root}")

    initial_files = scan_files(root)
    current_pdfs, word_files = print_scan_report(initial_files, root, output)

    if dry_run:
        check_word_targets(word_files)
        print("\nWord 转换计划：")
        if word_files:
            for source in word_files:
                target = source.with_suffix(".pdf")
                status = "将覆盖" if target.exists() and overwrite_word_pdf else (
                    "已存在，将复用" if target.exists() else "将转换"
                )
                print(f"  [{status}] {source.relative_to(root)} -> {target.name}")
        else:
            print("  无")
        pdfs = planned_pdfs(current_pdfs, word_files, root)
        print_merge_order(pdfs, root)
        print(f"\n预演完成：不会修改文件。计划输出到：{output}")
        return len(pdfs), 0

    # 严格分阶段：扫描全部结束 -> 转换全部 Word -> 重新扫描 -> 合并。
    convert_word_files(word_files, root, overwrite_word_pdf)
    pdfs = [
        path
        for path in scan_files(root)
        if path.suffix.casefold() == ".pdf" and not same_path(path, output)
    ]
    print_merge_order(pdfs, root)
    page_count = merge_pdfs(pdfs, output)
    print(
        f"\n完成：已将 {len(pdfs)} 个 PDF（共 {page_count} 页）合并到：\n{output}"
    )
    return len(pdfs), page_count


def main() -> int:
    args = parse_args()
    root = args.folder.expanduser().resolve()
    output = args.output.expanduser().resolve()

    try:
        run_merge(
            root,
            output,
            overwrite_word_pdf=args.overwrite_word_pdf,
            dry_run=args.dry_run,
        )
        return 0
    except (OSError, RuntimeError, ValueError) as exc:
        print(f"\n错误：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
