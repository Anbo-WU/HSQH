#!/usr/bin/env python3
"""Windows-compatible local OCR helpers for scanned confirmation PDFs."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import numpy as np
import pymupdf
from rapidocr import RapidOCR


RENDER_SCALE = 2.5


@dataclass(frozen=True)
class OCRLine:
    text: str
    confidence: float
    x: float
    y: float
    width: float
    height: float


def create_engine() -> RapidOCR:
    """Create one reusable ONNX OCR engine for the current process."""
    try:
        return RapidOCR()
    except Exception as exc:
        raise RuntimeError(f"无法初始化 Windows OCR：{exc}") from exc


def _render_page(page: pymupdf.Page) -> np.ndarray:
    pixmap = page.get_pixmap(
        matrix=pymupdf.Matrix(RENDER_SCALE, RENDER_SCALE),
        colorspace=pymupdf.csRGB,
        alpha=False,
    )
    return np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(
        pixmap.height,
        pixmap.width,
        pixmap.n,
    )


def _recognize_crop(
    engine: RapidOCR,
    image: np.ndarray,
    full_width: int,
    full_height: int,
    top_offset: int,
) -> list[OCRLine]:
    try:
        result = engine(image)
    except Exception as exc:
        raise RuntimeError(f"OCR 识别失败：{exc}") from exc

    if result.txts is None or result.boxes is None or result.scores is None:
        return []

    lines: list[OCRLine] = []
    for text, score, box in zip(result.txts, result.scores, result.boxes):
        points = np.asarray(box, dtype=float)
        min_x = float(points[:, 0].min())
        max_x = float(points[:, 0].max())
        min_y = float(points[:, 1].min()) + top_offset
        max_y = float(points[:, 1].max()) + top_offset
        lines.append(
            OCRLine(
                text=str(text),
                confidence=float(score),
                x=min_x / full_width,
                y=1.0 - max_y / full_height,
                width=(max_x - min_x) / full_width,
                height=(max_y - min_y) / full_height,
            )
        )
    return lines


def recognize_split_pages(
    source: Path,
    *,
    regions: tuple[tuple[float, float], ...] = ((0.0, 0.5), (0.6, 1.0)),
) -> list[tuple[OCRLine, ...]]:
    """OCR the top and bottom regions needed to determine document boundaries."""
    engine = create_engine()
    try:
        document = pymupdf.open(source)
    except Exception as exc:
        raise RuntimeError(f"无法打开扫描 PDF：{source}") from exc
    try:
        if document.page_count < 1:
            raise RuntimeError(f"扫描 PDF 没有页面：{source}")
        recognized: list[tuple[OCRLine, ...]] = []
        for page_index in range(document.page_count):
            image = _render_page(document[page_index])
            height, width = image.shape[:2]
            lines: list[OCRLine] = []
            for start, end in regions:
                first = max(0, min(height - 1, round(height * start)))
                last = max(first + 1, min(height, round(height * end)))
                lines.extend(
                    _recognize_crop(engine, image[first:last], width, height, first)
                )
            recognized.append(tuple(lines))
            page_number = page_index + 1
            if page_number == 1 or page_number % 5 == 0 or page_number == document.page_count:
                print(f"OCR 进度：{page_number}/{document.page_count}")
        return recognized
    finally:
        document.close()


def recognize_first_pages(pdfs: Iterable[Path]) -> dict[Path, str]:
    """OCR the upper part of each PDF's first page with one shared engine."""
    paths = list(pdfs)
    engine = create_engine()
    recognized: dict[Path, str] = {}
    errors: list[str] = []
    for number, path in enumerate(paths, start=1):
        resolved = path.resolve()
        try:
            with pymupdf.open(resolved) as document:
                if document.page_count < 1:
                    raise RuntimeError("PDF 没有首页")
                image = _render_page(document[0])
            height, width = image.shape[:2]
            top_end = max(1, round(height * 0.45))
            lines = _recognize_crop(
                engine,
                image[:top_end],
                width,
                height,
                0,
            )
            lines.sort(key=lambda line: (-(line.y + line.height), line.x))
            recognized[resolved] = "\n".join(line.text for line in lines[:12])
            print(f"首页 OCR：{number}/{len(paths)}  {path.name}")
        except Exception as exc:
            errors.append(f"{path.name}: {exc}")
    if errors:
        raise RuntimeError("OCR 失败：\n  " + "\n  ".join(errors))
    return recognized
