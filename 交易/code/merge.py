#!/usr/bin/env python3
"""在 Windows 上转换 Word，并按自然顺序合并确认书 PDF。

目录中没有 Word 时会完全跳过转换阶段。只有发现尚无同名 PDF 的 .doc/.docx
时，才通过 Windows PowerShell 调用 Microsoft Word 的 COM 接口导出 PDF。
合并由 pypdf 完成；奇数页确认书会先用 PyMuPDF 渲染末页判断是否空白：
空白尾页删除，非空白尾页后补一张同尺寸空白页。
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
import uuid
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Iterable


WORD_SUFFIXES = {".doc", ".docx"}


WORD_TO_PDF_POWERSHELL = r'''
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath
)

$ErrorActionPreference = "Stop"
$word = $null
$document = $null

try {
    $jobs = @(Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json)
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    # msoAutomationSecurityForceDisable：自动转换外部文档时禁止执行宏。
    $word.AutomationSecurity = 3

    foreach ($job in $jobs) {
        try {
            $source = [string]$job.source
            $target = [string]$job.target
            # ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False
            $document = $word.Documents.Open($source, $false, $true, $false)
            # wdExportFormatPDF = 17
            $document.ExportAsFixedFormat($target, 17)
        }
        finally {
            if ($null -ne $document) {
                $document.Close(0)
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
                $document = $null
            }
        }
    }
}
finally {
    if ($null -ne $document) {
        $document.Close(0)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    if ($null -ne $word) {
        $word.Quit(0)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
'''


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("folder", type=Path, help="要扫描的确认书目录")
    parser.add_argument(
        "-o", "--output", type=Path, required=True, help="合并结果路径"
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


def print_scan_report(
    files: list[Path], root: Path, output: Path
) -> tuple[list[Path], list[Path]]:
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
    print(
        f"共发现 {len(files)} 个文件：{len(pdfs)} 个待合并 PDF，"
        f"{len(word_files)} 个 Word，{len(non_pdfs) - len(word_files)} 个其他文件。"
    )
    if non_pdfs:
        print("\n发现以下非 PDF 文件：")
        for path in non_pdfs:
            kind = "Word，待检查转换" if path.suffix.casefold() in WORD_SUFFIXES else "其他文件"
            print(f"  [{kind}] {path.relative_to(root)}")
    else:
        print("\n全部输入都是 PDF，将完全跳过 Word 转 PDF 阶段。")

    return pdfs, word_files


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


def load_pypdf():
    try:
        from pypdf import PdfReader, PdfWriter  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError(
            "缺少 PDF 处理库 pypdf。请执行：python -m pip install pypdf"
        ) from exc
    return PdfReader, PdfWriter


def pdf_page_count(pdf: Path) -> int:
    PdfReader, _ = load_pypdf()
    try:
        reader = PdfReader(str(pdf))
        if reader.is_encrypted:
            reader.decrypt("")
        return len(reader.pages)
    except Exception as exc:
        raise RuntimeError(f"无法读取转换后的 PDF：{pdf}") from exc


def validate_word_pdf(source: Path, target: Path) -> int:
    if not target.is_file() or target.stat().st_size < 5:
        raise RuntimeError("Word 未生成有效的 PDF 文件")
    with target.open("rb") as converted_file:
        if converted_file.read(5) != b"%PDF-":
            raise RuntimeError("Word 转换结果不是有效的 PDF 文件")

    expected_pages = declared_docx_page_count(source)
    actual_pages = pdf_page_count(target)
    if actual_pages < 1:
        raise RuntimeError("Word 转换结果没有页面")
    if expected_pages is not None and expected_pages != actual_pages:
        raise RuntimeError(
            f"转换页数异常：Word 记录为 {expected_pages} 页，"
            f"PDF 实际为 {actual_pages} 页"
        )
    return actual_pages


def find_windows_powershell() -> str | None:
    for command in ("powershell.exe", "powershell", "pwsh.exe", "pwsh"):
        found = shutil.which(command)
        if found is not None:
            return found
    return None


def run_word_conversion(jobs: list[tuple[Path, Path]]) -> None:
    if not jobs:
        return
    if sys.platform != "win32":
        raise RuntimeError("当前 Word 转 PDF 实现仅支持 Windows。")

    powershell = find_windows_powershell()
    if powershell is None:
        raise RuntimeError("找不到 Windows PowerShell，无法调用 Microsoft Word。")

    with tempfile.TemporaryDirectory(prefix="word_to_pdf_") as temp_name:
        temp_folder = Path(temp_name)
        script_path = temp_folder / "word_to_pdf.ps1"
        manifest_path = temp_folder / "jobs.json"
        script_path.write_text(WORD_TO_PDF_POWERSHELL, encoding="utf-8-sig")
        manifest_path.write_text(
            json.dumps(
                [
                    {"source": str(source.resolve()), "target": str(target.resolve())}
                    for source, target in jobs
                ],
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8-sig",
        )

        result = subprocess.run(
            [
                powershell,
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
                "-ManifestPath",
                str(manifest_path),
            ],
            capture_output=True,
            text=True,
            errors="replace",
            check=False,
            timeout=max(180, len(jobs) * 120),
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    if result.returncode != 0:
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError(
            "无法通过 Microsoft Word 转换文档。请确认已安装桌面版 Microsoft Word，"
            "且文档未被其他程序占用。"
            + (f"\n系统信息：{detail}" if detail else "")
        )


def convert_word_files(
    word_files: list[Path], root: Path, overwrite: bool
) -> None:
    if not word_files:
        print("\nWord 转 PDF：已跳过（目录中没有 Word）。")
        return

    check_word_targets(word_files)
    reusable: list[tuple[Path, Path, int]] = []
    conversion_jobs: list[tuple[Path, Path, Path]] = []
    for source in word_files:
        target = source.with_suffix(".pdf")
        if target.exists() and not overwrite:
            reusable.append((source, target, validate_word_pdf(source, target)))
            continue
        staged = target.with_name(f".{target.stem}.word-{uuid.uuid4().hex}.pdf")
        conversion_jobs.append((source, target, staged))

    print("\n开始处理 Word 文件：")
    for source, target, pages in reusable:
        print(
            f"  [复用] {source.relative_to(root)} -> {target.name}"
            f"（同名 PDF 已存在，{pages} 页）"
        )

    if not conversion_jobs:
        print("  没有需要启动 Microsoft Word 的文件。")
        return

    staged_paths = [staged for _, _, staged in conversion_jobs]
    try:
        # 一批文件只启动一次 Word，避免逐个启动造成额外等待。
        run_word_conversion(
            [(source, staged) for source, _, staged in conversion_jobs]
        )
        validated: list[tuple[Path, Path, Path, int]] = []
        for source, target, staged in conversion_jobs:
            pages = validate_word_pdf(source, staged)
            validated.append((source, target, staged, pages))

        # 全部转换并校验通过后，才替换正式 PDF。
        for source, target, staged, pages in validated:
            os.replace(staged, target)
            print(
                f"  [转换] {source.relative_to(root)} -> {target.name}"
                f"（Microsoft Word，{pages} 页）"
            )
    finally:
        for staged in staged_paths:
            if staged.exists():
                staged.unlink()


def planned_pdfs(
    current_pdfs: list[Path], word_files: list[Path], root: Path
) -> list[Path]:
    by_path = {os.path.normcase(str(path.resolve())): path for path in current_pdfs}
    for source in word_files:
        target = source.with_suffix(".pdf")
        by_path[os.path.normcase(str(target.resolve()))] = target
    return sorted(by_path.values(), key=lambda path: natural_path_key(path, root))


def rendered_page_is_blank(pdf_path: Path, page_index: int) -> bool:
    """渲染页面并保守判断是否视觉空白，覆盖扫描图和矢量内容。"""
    try:
        import pymupdf  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError(
            "缺少视觉空白页检测库 PyMuPDF。请执行：python -m pip install PyMuPDF"
        ) from exc

    try:
        with pymupdf.open(str(pdf_path)) as document:
            page = document[page_index]
            bounds = page.rect
            longest_side = max(float(bounds.width), float(bounds.height))
            if longest_side <= 0:
                return False
            scale = min(1000.0 / longest_side, 2.0)
            pixmap = page.get_pixmap(
                matrix=pymupdf.Matrix(scale, scale),
                colorspace=pymupdf.csGRAY,
                alpha=False,
            )
            pixels = pixmap.samples
    except Exception:
        # 判断不确定时绝不删除页面，后续会在其后补空白页。
        return False

    ink_pixels = sum(value < 250 for value in pixels)
    allowance = max(8, len(pixels) // 200_000)
    return ink_pixels <= allowance


def merge_with_pypdf(
    pdfs: list[Path], temporary_output: Path
) -> tuple[int, list[dict[str, object]]]:
    PdfReader, PdfWriter = load_pypdf()
    writer = PdfWriter()
    page_count = 0
    reports: list[dict[str, object]] = []
    try:
        for path in pdfs:
            try:
                reader = PdfReader(str(path))
                if reader.is_encrypted:
                    decrypt_result = reader.decrypt("")
                    if not decrypt_result:
                        raise RuntimeError(f"PDF 已加密，无法读取：{path}")
                original_page_count = len(reader.pages)
            except RuntimeError:
                raise
            except Exception as exc:
                raise RuntimeError(f"无法读取 PDF：{path}") from exc

            if original_page_count < 1:
                raise RuntimeError(f"PDF 没有页面：{path}")

            pages_to_copy = original_page_count
            should_append_blank = False
            action = "unchanged"
            last_page = reader.pages[-1]
            if original_page_count % 2 == 1:
                if rendered_page_is_blank(path, original_page_count - 1):
                    if original_page_count == 1:
                        raise RuntimeError(
                            f"唯一一页是空白页，无法作为确认书合并：{path}"
                        )
                    pages_to_copy -= 1
                    action = "removed_blank"
                else:
                    should_append_blank = True
                    action = "added_blank"

            for page in reader.pages[:pages_to_copy]:
                writer.add_page(page)
                page_count += 1
            if should_append_blank:
                media_box = last_page.mediabox
                blank_page = writer.add_blank_page(
                    width=float(media_box.width), height=float(media_box.height)
                )
                rotation = int(last_page.get("/Rotate", 0) or 0) % 360
                if rotation:
                    blank_page.rotate(rotation)
                page_count += 1

            output_page_count = pages_to_copy + int(should_append_blank)
            if output_page_count % 2:
                raise RuntimeError(f"奇偶页校验失败：{path}")
            reports.append(
                {
                    "path": str(path.resolve()),
                    "original_pages": original_page_count,
                    "output_pages": output_page_count,
                    "action": action,
                }
            )

        with temporary_output.open("wb") as output_file:
            writer.write(output_file)
    finally:
        close = getattr(writer, "close", None)
        if callable(close):
            close()
    return page_count, reports


def print_page_adjustment_report(reports: list[dict[str, object]]) -> None:
    removed = 0
    added = 0
    print("\n奇数页确认书处理：")
    for report in reports:
        action = report.get("action")
        if action == "unchanged":
            continue
        path = Path(str(report.get("path", "")))
        original_pages = report.get("original_pages")
        output_pages = report.get("output_pages")
        if action == "removed_blank":
            removed += 1
            label = "删除空白尾页"
        elif action == "added_blank":
            added += 1
            label = "补入空白页"
        else:
            raise RuntimeError(f"未知的 PDF 页面处理结果：{action}")
        print(f"  [{label}] {path.name}（{original_pages} 页 -> {output_pages} 页）")
    if not removed and not added:
        print("  所有确认书原本均为偶数页，无需调整。")
    print(f"  汇总：删除 {removed} 个空白尾页，补入 {added} 个空白页。")


def merge_pdfs(pdfs: list[Path], output: Path) -> int:
    if not pdfs:
        raise RuntimeError("没有找到可合并的 PDF 文件。")

    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".merge_pdf_", dir=output.parent) as temp_name:
        temporary_output = Path(temp_name) / "merged.pdf"
        page_count, reports = merge_with_pypdf(pdfs, temporary_output)

        if not temporary_output.is_file() or temporary_output.stat().st_size == 0:
            raise RuntimeError("合并程序没有生成有效的输出文件。")
        actual_page_count = pdf_page_count(temporary_output)
        if actual_page_count != page_count:
            raise RuntimeError(
                f"合并页数校验失败：计划 {page_count} 页，实际 {actual_page_count} 页。"
            )
        if page_count % 2:
            raise RuntimeError("合并结果仍为奇数页，已停止输出。")
        os.replace(temporary_output, output)
    print_page_adjustment_report(reports)
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

    # 只递归扫描一次。没有 Word 时直接使用本次结果进入合并。
    initial_files = scan_files(root)
    current_pdfs, word_files = print_scan_report(initial_files, root, output)

    if dry_run:
        check_word_targets(word_files)
        print("\nWord 转换计划：")
        if word_files:
            for source in word_files:
                target = source.with_suffix(".pdf")
                status = "将覆盖" if target.exists() and overwrite_word_pdf else (
                    "已存在，将复用" if target.exists() else "将用 Microsoft Word 转换"
                )
                print(f"  [{status}] {source.relative_to(root)} -> {target.name}")
        else:
            print("  无（完全跳过转换器）")
        pdfs = planned_pdfs(current_pdfs, word_files, root)
        print_merge_order(pdfs, root)
        print(f"\n预演完成：不会修改文件。计划输出到：{output}")
        return len(pdfs), 0

    convert_word_files(word_files, root, overwrite_word_pdf)
    pdfs = planned_pdfs(current_pdfs, word_files, root)
    missing = [path for path in pdfs if not path.is_file()]
    if missing:
        raise RuntimeError(
            "Word 转换后仍缺少计划中的 PDF：\n  "
            + "\n  ".join(str(path) for path in missing)
        )

    print_merge_order(pdfs, root)
    page_count = merge_pdfs(pdfs, output)
    print(f"\n完成：已将 {len(pdfs)} 个 PDF（共 {page_count} 页）合并到：\n{output}")
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
    except (OSError, RuntimeError, ValueError, subprocess.SubprocessError) as exc:
        print(f"\n错误：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
