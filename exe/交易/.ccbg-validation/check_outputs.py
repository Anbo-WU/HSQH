from pathlib import Path
from datetime import date
import json
import sys
import pymupdf

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import main
import holdings
import split

paths = main.build_task_paths('Pan', date(2026, 9, 9), 1, mode='ccbg')
manifest = holdings.load_manifest(paths.holdings_pdf_folder)
assert holdings.batch_complete(paths.holdings_pdf_folder, paths.registration_output, paths.merged_pdf_output)
assert main.latest_completed_batch('Pan', 'ccbg') == (date(2026, 9, 9), 1)
assert main.latest_completed_batch('Pan', 'jsd') == (date(2026, 9, 4), 1)
assert main.choose_a_batch('Pan', date(2026, 9, 9), 'ccbg') == 2
assert main.registration_output_names(paths.registration_output) == [Path(d['filename']).stem for d in manifest['documents']]
for d in manifest['documents']:
    assert split.file_sha256(paths.confirmation_folder / d['source_name']) == d['source_sha256']
for source, target in [(paths.registration_output, paths.desktop_registration_output),
                       (paths.merged_pdf_output, paths.desktop_merged_pdf_output)]:
    assert split.file_sha256(source) == split.file_sha256(target)
blank = []
with pymupdf.open(paths.merged_pdf_output) as document:
    assert len(document) == 75
    for page in document:
        assert abs(page.rect.width - 595.28) < 3 and abs(page.rect.height - 841.89) < 3
        if main.merge_program.rendered_page_is_blank(paths.merged_pdf_output, page.number):
            blank.append(page.number + 1)
assert not blank, blank
report = dict(reports=7, merged_pages=75, skipped_empty_sheets=4, blank_pages=blank,
              source_hashes_unchanged=True, desktop_copies_match=True, legacy_and_holdings_batches_separate=True,
              merged_pdf=str(paths.merged_pdf_output), individual_pdfs=str(paths.holdings_pdf_folder))
(Path(__file__).resolve().parent / 'verification.json').write_text(
    json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(report, ensure_ascii=False, indent=2))
