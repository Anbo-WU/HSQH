"""Pan 持仓报告：只读表格、按 sheet 打印，以及按账户首页拆分扫描件。"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import asdict, dataclass, replace
from pathlib import Path

import pymupdf
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import RowBreak, ColBreak

import merge
import registration
import split
from windows_ocr import recognize_split_pages


SHEETS = ("账户状况", "持仓明细", "历史交易", "资金明细")
MANIFEST_NAME = "持仓批次记录.json"
ROOT = Path(__file__).resolve().parents[1]
# 匹配打印内容，而不是只匹配 Excel 标签；原始账户页未印“账户状况”。
BOUNDARY_RULES = {
    "account": (("持仓报告", "客户编号", "客户名称"), ("账户状况", "客户名称")),
    "funds": (("资金明细",), ("发生时间", "收支方向", "资金类型")),
}


@dataclass(frozen=True)
class SheetInfo:
    name: str
    area: str
    scale: int
    landscape: bool
    empty: bool


def normalize(text: str) -> str:
    return re.sub(r"[\s\W_]+", "", text).casefold()


def collect_workbooks(folder: Path) -> list[Path]:
    if not folder.is_dir():
        raise RuntimeError(f"持仓报告目录不存在：{folder}")
    paths = sorted(
        (p for p in folder.iterdir() if p.is_file() and not p.name.startswith("~$")
         and p.suffix.casefold() in {".xlsx", ".xlsm", ".xls"}),
        key=registration.natural_key,
    )
    if not paths:
        raise RuntimeError(f"没有找到持仓报告表格：{folder}")
    seen: set[str] = set()
    for path in paths:
        if path.stem.casefold() in seen:
            raise RuntimeError(f"存在同名表格，会生成重复 PDF：{path.stem}")
        seen.add(path.stem.casefold())
    return paths


def has_value(value: object) -> bool:
    return value is not None and (not isinstance(value, str) or bool(value.strip()))


def sheet_has_data(sheet) -> bool:
    rows = [[c.value for c in row] for row in sheet.iter_rows()]
    if not any(has_value(value) for row in rows for value in row):
        return False
    if sheet.title.strip() in {"持仓明细", "历史交易"}:
        # 导出器也可能留下栏目和零值汇总，只有交易记录才算非空。
        header_rows = [i for i, row in enumerate(rows) if "交易编号" in row]
        if header_rows:
            for i in header_rows:
                col = rows[i].index("交易编号")
                for row in rows[i + 1:]:
                    value = str(row[col] or "").strip()
                    if value and value != "交易编号" and sum(has_value(v) for v in row) >= 3:
                        return True
            return False
    if sheet.title.strip() == "资金明细":
        headers = [i for i, row in enumerate(rows) if "发生时间" in row]
        if headers:
            col = rows[headers[0]].index("发生时间")
            return any(has_value(row[col]) for row in rows[headers[0] + 1:])
    return True


def column_width_points(sheet, column: int) -> float:
    width = sheet.sheet_format.defaultColWidth or 8.43
    for dim in sheet.column_dimensions.values():
        if (dim.min or 0) <= column <= (dim.max or dim.min or 0):
            if dim.hidden:
                return 0.0
            width = dim.width
            break
    # Excel 列宽以数字字符宽度计量；留出 2% 的渲染器差异余量。
    return (int(width * 7 + 5) * 72 / 96)


def inspect_sheet(sheet) -> SheetInfo:
    cells = [c for row in sheet.iter_rows() for c in row if has_value(c.value)]
    if not sheet_has_data(sheet):
        return SheetInfo(sheet.title, "", 100, sheet.title.strip() == "历史交易", True)
    last_row = max(c.row for c in cells)
    last_col = max(c.column for c in cells)
    for area in sheet.merged_cells.ranges:
        if has_value(sheet.cell(area.min_row, area.min_col).value):
            last_row = max(last_row, area.max_row)
            last_col = max(last_col, area.max_col)
    landscape = sheet.title.strip() == "历史交易"
    page_width = 841.89 if landscape else 595.28
    available = page_width - 72 * (sheet.page_margins.left + sheet.page_margins.right)
    width = sum(column_width_points(sheet, c) for c in range(1, last_col + 1))
    original = sheet.page_setup.scale or 100
    estimated_scale = min(original, 100, int(available * 0.98 / max(width, 1) * 100))
    if estimated_scale < 10:
        raise RuntimeError(f"{sheet.title} 太宽，缩至 10% 仍无法完整放入 A4。")
    return SheetInfo(sheet.title, f"A1:{get_column_letter(last_col)}{last_row}", original, landscape, False)


def find_soffice() -> Path:
    candidates = [os.environ.get("TRADER_SOFFICE"), shutil.which("soffice")]
    candidates.extend(str(ROOT.parent / ".tools" / "LibreOffice" / "program" / exe)
                      for exe in ("soffice.com", "soffice.exe"))
    for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
        if base:
            candidates.append(str(Path(base) / "LibreOffice/program/soffice.com"))
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return Path(candidate)
    raise RuntimeError(
        "持仓表格转 PDF 需要 LibreOffice。请安装桌面版，或将独立解压版放到 "
        "exe/.tools/LibreOffice；也可用 TRADER_SOFFICE 指定 soffice.com 路径。"
    )


def office_convert(inputs: list[Path], outdir: Path, profile: Path, fmt: str) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    command = [str(find_soffice()), f"-env:UserInstallation={profile.resolve().as_uri()}",
               "--headless", "--nologo", "--nodefault", "--nofirststartwizard",
               "--convert-to", fmt, "--outdir", str(outdir.resolve()),
               *(str(p.resolve()) for p in inputs)]
    try:
        result = subprocess.run(command, capture_output=True, encoding="utf-8", errors="replace",
                                timeout=max(180, len(inputs) * 90),
                                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (OSError, subprocess.SubprocessError) as exc:
        raise RuntimeError(f"表格转 PDF 失败：{exc}") from exc
    extension = fmt.split(":", 1)[0]
    missing = [p.name for p in inputs if not (outdir / f"{p.stem}.{extension}").is_file()]
    if result.returncode or missing:
        raise RuntimeError(f"LibreOffice 转换失败：{', '.join(missing)}\n"
                           f"{result.stderr or result.stdout}")


def open_source(path: Path, temporary: Path):
    if path.suffix.casefold() == ".xls":
        office_convert([path], temporary / "xlsx", temporary / "profile", "xlsx")
        path = temporary / "xlsx" / f"{path.stem}.xlsx"
    return load_workbook(path)


def export_sheet_pdfs(jobs: list[dict], temporary: Path) -> dict[str, int]:
    soffice = find_soffice()
    runtime = soffice.with_name("python.exe")
    if not runtime.is_file():
        raise RuntimeError(f"找不到 LibreOffice 自带的 Python：{runtime}")
    manifest = temporary / "export-jobs.json"
    manifest.write_text(json.dumps(jobs, ensure_ascii=False), encoding="utf-8")
    (temporary / "pdf").mkdir(exist_ok=True)
    try:
        result = subprocess.run(
            [str(runtime), str(Path(__file__).with_name("holdings_export.py")),
             str(manifest), str(soffice), str(temporary / "export-profile")],
            capture_output=True, encoding="utf-8", errors="replace",
            timeout=max(180, len(jobs) * 90), creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except (OSError, subprocess.SubprocessError) as exc:
        raise RuntimeError(f"持仓表格 PDF 排版失败：{exc}") from exc
    if result.returncode or any(not Path(job["target"]).is_file() for job in jobs):
        raise RuntimeError(f"持仓表格 PDF 排版失败：{result.stderr or result.stdout}")
    return {item["source"]: item["scale"] for item in
            json.loads(manifest.with_suffix(".result.json").read_text(encoding="utf-8"))}


def inspect_workbook(workbook) -> list[SheetInfo]:
    by_title = {s.title.strip(): s for s in workbook}
    missing = [name for name in SHEETS if name not in by_title]
    if missing:
        raise RuntimeError("缺少指定工作表：" + "、".join(missing))
    infos = [inspect_sheet(by_title[name]) for name in SHEETS]
    if infos[0].empty or infos[-1].empty:
        raise RuntimeError("账户状况或资金明细为空，无法保留要求的首页和末页，请先核对表格。")
    return infos


def prepare_sheet_copy(workbook, info: SheetInfo, target: Path) -> None:
    sheet = workbook[info.name]
    for other in workbook:
        other.sheet_state = "visible" if other is sheet else "hidden"
        if other is not sheet:
            # Calc 会导出隐藏但仍有显式打印区域的 sheet，必须一并清除。
            other.print_area = ""
    workbook.active = workbook.index(sheet)
    sheet.print_area = info.area
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.orientation = "landscape" if info.landscape else "portrait"
    # 列宽估算只用于预览；最终由排版器按实际字体度量适配一页宽，避免截列。
    sheet.page_setup.scale = None
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    # 只在临时打印副本中清除旧打印机的分页，原行高、列宽、字体、边框不变。
    sheet.row_breaks = RowBreak()
    sheet.col_breaks = ColBreak()
    if sheet.title.strip() == "资金明细":
        sheet.print_title_rows = "1:2"
    workbook.save(target)


def assemble_document(parts: list[tuple[SheetInfo, Path]], target: Path) -> list[dict]:
    document = pymupdf.open()
    details: list[dict] = []
    try:
        for info, part in parts:
            first = len(document) + 1
            with pymupdf.open(part) as source:
                for page in source:
                    if info.landscape:
                        page.set_rotation((page.rotation + 270) % 360)
                        page.remove_rotation()
                    # 输出必须是物理 A4 竖版；历史交易的文字方向保持左转 90 度。
                    if abs(page.rect.width - 595.28) > 3 or abs(page.rect.height - 841.89) > 3:
                        raise RuntimeError(f"{target.name}/{info.name} 未生成 A4 竖版页面：{page.rect}")
                    if not page.get_text().strip() and merge.rendered_page_is_blank(part, page.number):
                        continue
                    document.insert_pdf(source, from_page=page.number, to_page=page.number)
            count = len(document) - first + 1
            if count < 1:
                raise RuntimeError(f"{info.name} 有数据但转换后没有有效页面。")
            details.append({**asdict(info), "start_page": first, "end_page": len(document), "pages": count})
        document.save(target, garbage=4, deflate=True)
    finally:
        document.close()
    return details


def run_before(paths, dry_run: bool = False) -> tuple[int, int]:
    sources = collect_workbooks(paths.confirmation_folder)
    names = [p.stem for p in sources]
    source_hashes = {p: split.file_sha256(p) for p in sources}
    individual = paths.holdings_pdf_folder
    # Windows 下办公引擎的深层缓存可能清理失败，不让缓存清理使已完成批次报错。
    with tempfile.TemporaryDirectory(prefix=".holdings-", dir=paths.teacher_folder,
                                     ignore_cleanup_errors=True) as temp:
        temporary = Path(temp)
        jobs: list[tuple[Path, list[tuple[SheetInfo, Path]], list[str], str, str]] = []
        copies: list[dict] = []
        for number, source in enumerate(sources, 1):
            original_hash = source_hashes[source]
            book = open_source(source, temporary)
            try:
                infos = inspect_workbook(book)
                account = book[infos[0].name]
                company = str(account["J5"].value or "").strip()
                customer_id = str(account["D5"].value or "").strip()
                print(f"{number:>3}. {source.stem}")
                parts: list[tuple[SheetInfo, Path]] = []
                for index, info in enumerate(infos):
                    print(f"     {info.name}：" + ("空表，跳过" if info.empty else
                          f"{info.area}，适配一页宽，" + ("左转 90°" if info.landscape else "A4 竖版")))
                    if info.empty or dry_run:
                        continue
                    copy = temporary / f"{number:03d}_{index}.xlsx"
                    prepare_sheet_copy(book, info, copy)
                    copies.append({"source": str(copy), "target": str(temporary / "pdf" / f"{copy.stem}.pdf"),
                                   "sheet": info.name, "area": info.area, "max_scale": info.scale})
                    parts.append((info, temporary / "pdf" / f"{copy.stem}.pdf"))
                jobs.append((source, parts, [i.name for i in infos if i.empty], company, customer_id))
            finally:
                book.close()
            if split.file_sha256(source) != original_hash:
                raise RuntimeError(f"原表格在处理过程中发生变化，请关闭编辑后重试：{source.name}")
        if dry_run:
            print(f"持仓 PDF 单独目录：{individual}\n合并 PDF：{paths.merged_pdf_output}")
            return len(sources), 0
        if individual.exists() or paths.merged_pdf_output.exists() or paths.registration_output.exists():
            # 成功批次不可原地重建，避免 A 记录和已盖章扫描件失去对应。
            if batch_complete(individual, paths.registration_output, paths.merged_pdf_output):
                manifest = load_manifest(individual)
                if [(p.name, split.file_sha256(p)) for p in sources] == [
                    (d["source_name"], d["source_sha256"]) for d in manifest["documents"]
                ]:
                    print("该批次已完成且源文件未变化，复用已有输出。")
                    return len(sources), manifest["total_pages"]
            raise RuntimeError("本批次已有输出或源文件发生变化，未覆盖；请使用新的 --batch。")
        print(f"开始转换 {len(copies)} 个非空 sheet……", flush=True)
        actual_scales = export_sheet_pdfs(copies, temporary)
        for source in sources:
            if split.file_sha256(source) != source_hashes[source]:
                raise RuntimeError(f"转换期间原表格被编辑，未发布本批次：{source.name}")
        staged = temporary / "individual"
        staged.mkdir()
        records: list[dict] = []
        merged = pymupdf.open()
        try:
            for source, parts, skipped, company, customer_id in jobs:
                target = staged / f"{source.stem}.pdf"
                parts = [(replace(info, scale=actual_scales[str(temporary / f"{part.stem}.xlsx")]), part)
                         for info, part in parts]
                details = assemble_document(parts, target)
                with pymupdf.open(target) as document:
                    count = len(document)
                    merged.insert_pdf(document)
                records.append({"source_name": source.name, "source_sha256": source_hashes[source],
                                "filename": target.name, "company": company, "customer_id": customer_id,
                                "pages": count, "sheets": details, "skipped_sheets": skipped,
                                "pdf_sha256": split.file_sha256(target)})
                print(f"  {target.name}：{count} 页；" + "，".join(f"{d['name']} {d['pages']} 页" for d in details))
            total = len(merged)
            merged.save(temporary / "merged.pdf", garbage=4, deflate=True)
        finally:
            merged.close()
        registration.write_registration(paths.template_path, temporary / "registration.xls", names, paths.task_date)
        manifest = {"version": 1, "mode": "ccbg", "task_date": paths.task_date.isoformat(),
                    "batch": paths.batch, "total_pages": total, "documents": records,
                    "merged_sha256": split.file_sha256(temporary / "merged.pdf"),
                    "registration_sha256": split.file_sha256(temporary / "registration.xls")}
        (staged / MANIFEST_NAME).write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        individual.parent.mkdir(parents=True, exist_ok=True)
        paths.registration_output.parent.mkdir(parents=True, exist_ok=True)
        paths.merged_pdf_output.parent.mkdir(parents=True, exist_ok=True)
        staged.rename(individual)
        os.replace(temporary / "merged.pdf", paths.merged_pdf_output)
        os.replace(temporary / "registration.xls", paths.registration_output)
    return len(sources), total


def load_manifest(folder: Path) -> dict:
    try:
        manifest = json.loads((folder / MANIFEST_NAME).read_text(encoding="utf-8"))
        docs = manifest["documents"]
        if manifest["mode"] != "ccbg" or not docs or len({d["filename"].casefold() for d in docs}) != len(docs):
            raise ValueError("记录类型或文件名称异常")
        for doc in docs:
            if Path(doc["filename"]).name != doc["filename"] or doc["pages"] < 1:
                raise ValueError("记录文件名或页数异常")
        return manifest
    except (OSError, ValueError, KeyError, TypeError) as exc:
        raise RuntimeError(f"无法读取持仓 A 批次记录：{folder / MANIFEST_NAME}：{exc}") from exc


def batch_complete(folder: Path, registration_path: Path, merged: Path) -> bool:
    if not registration_path.is_file() or not merged.is_file() or not (folder / MANIFEST_NAME).is_file():
        return False
    manifest = load_manifest(folder)
    return (all((folder / d["filename"]).is_file()
                and split.file_sha256(folder / d["filename"]) == d["pdf_sha256"]
                for d in manifest["documents"])
            and split.file_sha256(registration_path) == manifest["registration_sha256"]
            and split.file_sha256(merged) == manifest["merged_sha256"])


def matches_rule(text: str, kind: str) -> bool:
    text = normalize(text)
    return any(all(normalize(word) in text for word in group) for group in BOUNDARY_RULES[kind])


def build_scan_plans(texts: list[str], documents: list[dict]) -> list[split.SplitPlan]:
    starts = [i for i, text in enumerate(texts) if matches_rule(text, "account")]
    if len(starts) != len(documents) or not starts or starts[0] != 0:
        raise RuntimeError(f"识别到账户首页 {len(starts)} 个，A 批次为 {len(documents)} 份；"
                           "或扫描开头存在多余页面。未执行拆分，请检查扫描件完整性。")
    plans: list[split.SplitPlan] = []
    used: set[str] = set()
    for number, start in enumerate(starts, 1):
        end = starts[number] if number < len(starts) else len(texts)
        first_text = normalize(texts[start])
        matches = [d for d in documents if d.get("company") and normalize(d["company"]) in first_text]
        if len(matches) > 1:
            matches = [d for d in matches if d.get("customer_id") and normalize(d["customer_id"]) in first_text]
        if len(matches) != 1:
            matches = [d for d in documents if d.get("customer_id") and normalize(d["customer_id"]) in first_text]
        if len(matches) != 1:
            raise RuntimeError(f"第 {start + 1} 页无法唯一匹配 A 批次客户名称/编号，未按位置猜测文件名。")
        doc = matches[0]
        if doc["filename"] in used:
            raise RuntimeError(f"扫描件出现重复账户首页：{doc['filename']}")
        used.add(doc["filename"])
        if end - start != doc["pages"]:
            raise RuntimeError(f"{doc['filename']} 扫描 {end - start} 页，A 为 {doc['pages']} 页，"
                               "可能漏扫、重扫或多出空白背面，未拆分。")
        if not matches_rule(texts[end - 1], "funds"):
            raise RuntimeError(f"第 {end} 页未识别到资金明细末页，未拆分。")
        # 资金明细每页重复表头；要求它只出现在 A 记录的最后一组 sheet。
        fund = doc["sheets"][-1]
        fund_start = start + fund["start_page"] - 1
        if any(not matches_rule(texts[i], "funds") for i in range(fund_start, end)):
            raise RuntimeError(f"{doc['filename']} 的资金明细页不完整或顺序异常，未拆分。")
        if any(matches_rule(texts[i], "funds") for i in range(start, fund_start)):
            raise RuntimeError(f"{doc['filename']} 在预期位置之前出现资金明细，未拆分。")
        plans.append(split.SplitPlan(number, start + 1, end, "持仓报告账户首页/资金末页",
                                     Path(doc["filename"]).stem, doc["filename"], True))
    return plans


def run_after(paths, names: list[str], dry_run: bool = False) -> tuple[int, int]:
    manifest = load_manifest(paths.holdings_pdf_folder)
    docs = manifest["documents"]
    if [Path(d["filename"]).stem for d in docs] != names:
        raise RuntimeError("持仓 A 批次记录与登记表名称/顺序不一致。")
    source = split.resolve_source_pdf(paths.scanned_folder, paths.scan_date,
                                     paths.scan_source_pdf if paths.scan_source_is_override else None)
    source_hash = split.file_sha256(source)
    existing = split.reusable_output(paths.split_output_folder, source, source_hash, len(docs))
    if existing is not None:
        expected = {d["filename"]: d["pages"] for d in docs}
        actual = {p.name: merge.pdf_page_count(p) for p in paths.split_output_folder.glob("*.pdf")}
        if actual != expected:
            raise RuntimeError("已有持仓拆分文件与 A 批次不一致，未覆盖。")
        return len(docs), existing["total_pages"]
    with pymupdf.open(source) as document:
        texts = [page.get_text() for page in document]
    # 扫描图没有文本层；沿用原项目本地 OCR 引擎，不上传文档。
    if not all(text.strip() for text in texts):
        print("正在 OCR 识别持仓报告账户首页、资金明细及客户信息……", flush=True)
        # 表格的边界标记均在页顶，避免对密集的历史交易正文逐格 OCR。
        recognized = recognize_split_pages(source, regions=((0.0, 0.22),))
        texts = ["\n".join(line.text for line in lines) for lines in recognized]
    plans = build_scan_plans(texts, docs)
    split.print_plans(source, plans, [], paths.split_output_folder, len(texts))
    if not dry_run:
        split.write_split_pdfs(source, paths.split_output_folder, plans, [], source_hash, len(docs), len(texts))
    return len(plans), len(texts)
