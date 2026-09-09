"""七份报告的完整 75 页模拟扫描与印章保留验证。"""
from pathlib import Path
from datetime import date
from dataclasses import replace
import sys
import json
import pymupdf

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import main
import holdings

root = Path(__file__).resolve().parent / 'actual'
paths = main.build_task_paths('Pan', date(2026, 9, 9), 1, mode='ccbg')
paths = replace(paths, teacher_folder=root, registration_output=root / 'registration.xls',
                merged_pdf_output=root / 'merged.pdf', split_output_folder=root / 'scan-split',
                scan_source_pdf=root / 'full-scan.pdf', scan_source_is_override=True)
scanned = pymupdf.open()
with pymupdf.open(root / 'merged.pdf') as original:
    for source in original:
        pix = source.get_pixmap(matrix=pymupdf.Matrix(2, 2))
        page = scanned.new_page(width=source.rect.width, height=source.rect.height)
        page.insert_image(page.rect, stream=pix.tobytes('png'))
        # 红色圆圈用于验证拆分过程保留附加图形，模拟盖章后的 PDF 内容。
        page.draw_circle(pymupdf.Point(525, 110), 20, color=(0.8, 0, 0), width=2)
scanned.save(paths.scan_source_pdf, deflate=True)
scanned.close()
names = main.registration_output_names(paths.registration_output)
count, pages = holdings.run_after(paths, names)
assert (count, pages) == (7, 75)
with pymupdf.open(paths.scan_source_pdf) as source:
    manifest = json.loads((paths.split_output_folder / '拆分记录.json').read_text(encoding='utf-8'))
    for doc in manifest['documents']:
        with pymupdf.open(paths.split_output_folder / doc['filename']) as result:
            for i, page in enumerate(result):
                a = source[doc['start_page'] - 1 + i].get_pixmap(matrix=pymupdf.Matrix(.5, .5)).samples
                b = page.get_pixmap(matrix=pymupdf.Matrix(.5, .5)).samples
                assert a == b, (doc['filename'], i)
main.create_desktop_zip(paths.split_output_folder, root / 'scan-split.zip')
print('PASS: 75 scanned image pages -> 7 correctly named files; every page including red mark matches source; ZIP valid.')
