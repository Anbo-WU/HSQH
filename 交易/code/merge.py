#!/usr/bin/env python3
"""递归扫描指定确认书文件夹，转换 Word 文档并按自然顺序合并全部 PDF。

Word 转 PDF 使用本机 WPS 的原生排版引擎；macOS 上的 PDF 合并优先使用
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
import time
import uuid
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Iterable


WORD_SUFFIXES = {".doc", ".docx"}
WPS_APP = Path("/Applications/wpsoffice.app")


WPS_EXPORT_SCRIPT = r'''
on run argv
    set sourcePath to item 1 of argv
    set sourceName to item 2 of argv

    tell application id "com.kingsoft.wpsoffice.mac"
        activate
        open POSIX file sourcePath
    end tell

    tell application "System Events"
        tell process "wpsoffice"
            repeat 240 times
                if exists window sourceName then exit repeat
                delay 0.25
            end repeat
            if not (exists window sourceName) then
                error "WPS 未在 60 秒内打开指定 Word：" & sourceName
            end if

            set frontmost to true
            tell menu 1 of menu bar item 3 of menu bar 1
                if exists menu item "Export to PDF..." then
                    click menu item "Export to PDF..."
                else if exists menu item "Export to PDF…" then
                    click menu item "Export to PDF…"
                else if exists menu item "输出为 PDF..." then
                    click menu item "输出为 PDF..."
                else if exists menu item "输出为PDF..." then
                    click menu item "输出为PDF..."
                else if exists menu item "导出为 PDF..." then
                    click menu item "导出为 PDF..."
                else if exists menu item "导出为PDF..." then
                    click menu item "导出为PDF..."
                else
                    error "找不到 WPS 的导出 PDF 菜单"
                end if
            end tell

            set exportWindowName to ""
            repeat 240 times
                if exists window "Export to PDF" then
                    set exportWindowName to "Export to PDF"
                    exit repeat
                else if exists window "输出为 PDF" then
                    set exportWindowName to "输出为 PDF"
                    exit repeat
                else if exists window "输出为PDF" then
                    set exportWindowName to "输出为PDF"
                    exit repeat
                else if exists window "导出为 PDF" then
                    set exportWindowName to "导出为 PDF"
                    exit repeat
                else if exists window "导出为PDF" then
                    set exportWindowName to "导出为PDF"
                    exit repeat
                end if
                delay 0.25
            end repeat
            if exportWindowName is "" then
                error "WPS 未在 60 秒内打开导出 PDF 窗口"
            end if

            tell group 1 of window exportWindowName
                repeat 240 times
                    if (count of pop up buttons) is greater than or equal to 2 then
                        if (exists button "Export") or (exists button "导出") then
                            exit repeat
                        end if
                    end if
                    delay 0.25
                end repeat
                if (count of pop up buttons) is less than 2 then
                    error "WPS 导出窗口尚未加载完成"
                end if

                set outputPopup to pop up button 2
                click outputPopup
                delay 0.5
                -- 使用 WPS 已获写入权限的自定义目录。临时 Word 使用随机文件名，
                -- 因此不会覆盖该目录中的任何既有 PDF。
                key code 115
                key code 125
                key code 36
                delay 0.5

                set outputPopup to pop up button 2
                set outputChoice to name of outputPopup
                if outputChoice is not "Custom Folder" and outputChoice is not "自定义文件夹" then
                    error "无法把 WPS 输出位置切换到自定义文件夹，当前为：" & outputChoice
                end if

                set outputFolder to ""
                repeat with labelItem in every static text
                    try
                        set labelText to value of labelItem as text
                        if labelText starts with "/" then set outputFolder to labelText
                    end try
                end repeat
                if outputFolder is "" then
                    error "无法读取 WPS 自定义输出目录"
                end if

                if exists button "Export" then
                    click button "Export"
                else if exists button "导出" then
                    click button "导出"
                else
                    error "找不到 WPS 导出按钮"
                end if
            end tell
            return "WPS_OUTPUT_FOLDER=" & outputFolder
        end tell
    end tell
end run
'''


WPS_CLEANUP_SCRIPT = r'''
on run argv
    set sourceName to item 1 of argv
    tell application "System Events"
        if not (exists process "wpsoffice") then return
        tell process "wpsoffice"
            set frontmost to true
            if exists window "Task Completed" then
                tell window "Task Completed"
                    set taskCloseButtons to every button whose description is "close button"
                    if (count of taskCloseButtons) > 0 then click item 1 of taskCloseButtons
                end tell
                delay 0.25
            end if
            repeat with dialogName in {"Export to PDF", "输出为 PDF", "输出为PDF", "导出为 PDF", "导出为PDF"}
                if exists window dialogName then
                    tell window dialogName
                        set closeButtons to every button whose description is "close button"
                        if (count of closeButtons) > 0 then click item 1 of closeButtons
                    end tell
                    exit repeat
                end if
            end repeat
            delay 0.5
            if exists window sourceName then
                set frontmost to true
                keystroke "w" using command down
            end if
        end tell
    end tell
end run
'''


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


def find_wps_office() -> Path | None:
    return WPS_APP if WPS_APP.is_dir() else None


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


def declared_docx_page_count(source: Path) -> int | None:
    """读取 Word 保存时记录的页数；旧版 .doc 没有对应的 OOXML 元数据。"""
    if source.suffix.casefold() != ".docx":
        return None
    try:
        with zipfile.ZipFile(source) as document:
            properties = ET.fromstring(document.read("docProps/app.xml"))
    except (KeyError, OSError, ET.ParseError, zipfile.BadZipFile) as exc:
        raise RuntimeError(f"无法读取 DOCX 页数信息：{source.name}") from exc

    pages = properties.find(
        "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Pages"
    )
    if pages is None or not (pages.text or "").isdigit():
        return None
    page_count = int(pages.text or "0")
    return page_count if page_count > 0 else None


def pdf_page_count(pdf: Path) -> int | None:
    """读取常规 PDF 的页面对象数；对象流 PDF 无法可靠读取时返回 None。"""
    try:
        data = pdf.read_bytes()
    except OSError as exc:
        raise RuntimeError(f"无法读取转换后的 PDF：{pdf}") from exc
    count = len(re.findall(rb"/Type\s*/Page\b", data))
    return count if count > 0 else None


def validate_word_pdf(source: Path, target: Path) -> int | None:
    if not target.is_file() or target.stat().st_size < 5:
        raise RuntimeError("WPS 未生成有效的 PDF 文件")
    with target.open("rb") as converted_file:
        if converted_file.read(5) != b"%PDF-":
            raise RuntimeError("WPS 转换结果不是有效的 PDF 文件")

    expected_pages = declared_docx_page_count(source)
    actual_pages = pdf_page_count(target)
    if (
        expected_pages is not None
        and actual_pages is not None
        and expected_pages != actual_pages
    ):
        raise RuntimeError(
            f"转换页数异常：Word 记录为 {expected_pages} 页，"
            f"PDF 实际为 {actual_pages} 页"
        )
    return actual_pages


def wait_for_pdf(target: Path, timeout: float = 180.0) -> None:
    deadline = time.monotonic() + timeout
    previous_size = -1
    stable_checks = 0
    while time.monotonic() < deadline:
        if target.is_file():
            size = target.stat().st_size
            if size >= 5 and size == previous_size:
                stable_checks += 1
                if stable_checks >= 3:
                    return
            else:
                stable_checks = 0
            previous_size = size
        time.sleep(0.5)
    raise RuntimeError(f"等待 WPS 输出超时：{target}")


def run_wps_script(script_text: str, *arguments: str) -> subprocess.CompletedProcess[str]:
    osascript = shutil.which("osascript")
    if osascript is None:
        raise RuntimeError("找不到 macOS osascript，无法调用 WPS 导出 PDF。")
    with tempfile.TemporaryDirectory(prefix=".wps_automation_") as temp_name:
        script_path = Path(temp_name) / "wps_export.applescript"
        script_path.write_text(script_text, encoding="utf-8")
        return subprocess.run(
            [osascript, str(script_path), *arguments],
            capture_output=True,
            text=True,
            check=False,
            timeout=190,
        )


def cleanup_wps_window(source: Path) -> None:
    try:
        run_wps_script(WPS_CLEANUP_SCRIPT, source.name)
    except (OSError, subprocess.SubprocessError):
        # 转换结果已经单独校验；清理 WPS 窗口失败不应破坏正确 PDF。
        pass


def wps_output_folder(stdout: str) -> Path:
    prefix = "WPS_OUTPUT_FOLDER="
    for line in reversed(stdout.splitlines()):
        if line.startswith(prefix):
            folder = Path(line.removeprefix(prefix)).expanduser()
            if folder.is_dir():
                return folder
            raise RuntimeError(f"WPS 输出目录不存在：{folder}")
    raise RuntimeError("WPS 未返回实际输出目录。")


def convert_one_word(
    source: Path,
    target: Path,
    wps_office: Path,
    overwrite: bool,
) -> str:
    if target.exists() and not overwrite:
        pages = validate_word_pdf(source, target)
        detail = f"，{pages} 页" if pages is not None else ""
        return f"同名 PDF 已存在且校验通过，直接使用{detail}"

    if not wps_office.is_dir():
        raise RuntimeError(f"找不到 WPS Office：{wps_office}")
    if not same_path(target, source.with_suffix(".pdf")):
        raise RuntimeError("WPS 自动导出仅支持生成 Word 所在目录中的同名 PDF。")

    converted: Path | None = None
    staged_target = target.with_name(f".{target.name}.wps-{uuid.uuid4().hex}.tmp")
    with tempfile.TemporaryDirectory(
        prefix=".wps_source_", dir=target.parent
    ) as temp_name:
        temp_source = Path(temp_name) / f"wps-{uuid.uuid4().hex}{source.suffix.casefold()}"
        shutil.copy2(source, temp_source)
        try:
            result = run_wps_script(
                WPS_EXPORT_SCRIPT,
                str(temp_source.resolve()),
                temp_source.name,
            )
            if result.returncode != 0:
                detail = (result.stderr or result.stdout).strip()
                if "辅助功能" in detail or "not authorized" in detail.casefold():
                    detail += (
                        "\n请在“系统设置 → 隐私与安全性 → 辅助功能”中允许当前终端控制 WPS。"
                    )
                raise RuntimeError(detail or "WPS 自动导出失败")

            output_folder = wps_output_folder(result.stdout)
            converted = output_folder / temp_source.with_suffix(".pdf").name
            wait_for_pdf(converted)
            pages = validate_word_pdf(source, converted)

            shutil.copy2(converted, staged_target)
            validate_word_pdf(source, staged_target)
            os.replace(staged_target, target)
        finally:
            cleanup_wps_window(temp_source)
            if staged_target.exists():
                staged_target.unlink()
            if converted is not None and converted.exists():
                converted.unlink()

    detail = f"，{pages} 页" if pages is not None else ""
    return f"WPS 转换完成{detail}"


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
    wps_office = find_wps_office() if needs_conversion else None
    if needs_conversion and wps_office is None:
        raise RuntimeError(
            "发现需要转换的 Word 文件，但找不到 WPS Office：\n"
            f"  {WPS_APP}\n"
            "为保证字体、行距和分页与原 Word 一致，本程序不会回退到 LibreOffice。"
        )

    print("\n开始处理 Word 文件：")
    failures: list[str] = []
    for source in word_files:
        target = source.with_suffix(".pdf")
        try:
            assert wps_office is not None or target.exists()
            status = convert_one_word(source, target, wps_office or WPS_APP, overwrite)
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
