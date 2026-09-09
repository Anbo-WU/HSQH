"""用七份真实表格验证 A 和 B；测试产物留在此目录，不触发打印或桌面复制。"""
from __future__ import annotations

import json
import sys
from dataclasses import replace
from datetime import date
from pathlib import Path

import pymupdf

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import main
import holdings
import split

destination = Path(__file__).resolve().parent / 'actual'
destination.mkdir(exist_ok=True)
paths = main.build_task_paths('Pan', date(2026, 9, 9), 1, mode='ccbg')
# TaskPaths 的正式个别 PDF 目录使用 teacher_folder 属性；用隔离目录验证。
paths = replace(paths, teacher_folder=destination,
                registration_output=destination / 'registration.xls',
                merged_pdf_output=destination / 'merged.pdf',
                split_output_folder=destination / 'split',
                scan_source_pdf=destination / 'merged.pdf', scan_source_is_override=True)
before = {p: split.file_sha256(p) for p in holdings.collect_workbooks(paths.confirmation_folder)}
count, pages = holdings.run_before(paths)
assert count == 7
assert before == {p: split.file_sha256(p) for p in before}
names = main.registration_output_names(paths.registration_output)
assert names == [p.stem for p in before]
manifest = holdings.load_manifest(paths.holdings_pdf_folder)
assert sum(d['pages'] for d in manifest['documents']) == pages
assert sum(len(d['skipped_sheets']) for d in manifest['documents']) == 4
for doc in manifest['documents']:
    pdf = paths.holdings_pdf_folder / doc['filename']
    with pymupdf.open(pdf) as rendered:
        assert len(rendered) == doc['pages']
        for info in doc['sheets']:
            for i in range(info['start_page'] - 1, info['end_page']):
                page = rendered[i]
                assert 590 < page.rect.width < 598 and 838 < page.rect.height < 845
                lines = [line for b in page.get_text('dict')['blocks'] if 'lines' in b for line in b['lines']]
                assert lines, (doc['filename'], i, 'blank page')
                direction = lines[0]['dir']
                if info['name'] == '历史交易':
                    assert direction[1] < -0.9, direction
                else:
                    assert direction[0] > 0.9, direction
        # 账户首页、历史交易首尾、资金末页供人工检查。
        if doc is manifest['documents'][0] or '江苏中恩' in doc['filename']:
            for info in doc['sheets']:
                page = rendered[info['start_page'] - 1]
                page.get_pixmap(matrix=pymupdf.Matrix(1.4, 1.4)).save(
                    destination / f"{doc['company']}_{info['name']}.png")
holdings.run_after(paths, names)
# OCR 样本使用三个短报告的全部页面，覆盖四张空表的分支和资金表头识别。
sample = manifest['documents'][-3:]
scanned = pymupdf.open()
for doc in sample:
    with pymupdf.open(paths.holdings_pdf_folder / doc['filename']) as source:
        for original in source:
            pix = original.get_pixmap(matrix=pymupdf.Matrix(2, 2))
            page = scanned.new_page(width=original.rect.width, height=original.rect.height)
            page.insert_image(page.rect, stream=pix.tobytes('png'))
scanned.save(destination / 'ocr-sample.pdf', deflate=True)
scanned.close()
recognized = holdings.recognize_split_pages(destination / 'ocr-sample.pdf')
texts = ['\n'.join(line.text for line in lines) for lines in recognized]
(destination / 'ocr-texts.json').write_text(json.dumps(texts, ensure_ascii=False, indent=2), encoding='utf-8')
plans = holdings.build_scan_plans(texts, sample)
assert len(plans) == 3
print(f'PASS: {count} workbooks, {pages} portrait A4 pages, 7 splits, 3 image-only OCR reports; sources unchanged.')
