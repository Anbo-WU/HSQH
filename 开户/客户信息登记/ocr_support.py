"""本地 PDF OCR，保留位置、置信度和可重复使用的识别缓存。"""
from __future__ import annotations

from dataclasses import asdict, dataclass
import hashlib
import json
from pathlib import Path

import numpy as np
import pymupdf


@dataclass(frozen=True)
class Line:
    text: str
    score: float
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2


class PDFReader:
    def __init__(self, cache: Path):
        self.cache = cache
        self.engine = None
        self.hashes: dict[Path, str] = {}

    def page(self, path: Path, index: int, crop=None, scale=2.5, rotation=0) -> list[Line]:
        if path not in self.hashes:
            self.hashes[path] = hashlib.sha256(path.read_bytes()).hexdigest()
        key = hashlib.sha256(
            repr((self.hashes[path], index, crop, scale, 'v1' if rotation == 0 else rotation)).encode()
        ).hexdigest()
        target = self.cache / f'{key}.json'
        if target.exists():
            return [Line(**item) for item in json.loads(target.read_text('utf-8'))]
        with pymupdf.open(path) as doc:
            page = doc[index]
            if rotation:
                page.set_rotation((page.rotation + rotation) % 360)
            width, height = page.rect.width, page.rect.height
            rect = pymupdf.Rect(0, 0, width, height)
            if crop is not None:
                rect = pymupdf.Rect(crop[0]*width, crop[1]*height,
                                    crop[2]*width, crop[3]*height)
            words = page.get_text('words', clip=rect)
            if words:
                lines = [Line(w[4], 1.0, w[0]/width, w[1]/height,
                              w[2]/width, w[3]/height) for w in words]
            else:
                if self.engine is None:
                    from rapidocr import RapidOCR
                    self.engine = RapidOCR()
                print(f'OCR：{path.name} 第 {index+1} 页' + ('（局部）' if crop else ''), flush=True)
                pix = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale),
                                      clip=rect, colorspace=pymupdf.csRGB, alpha=False)
                arr = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, 3)
                result = self.engine(arr)
                lines = []
                from rapidocr.utils.output import RapidOCROutput
                if not isinstance(result, RapidOCROutput):
                    raise RuntimeError('OCR 未返回完整文字识别结果，请检查检测和识别配置')
                if result.txts is not None:
                    if result.scores is None or result.boxes is None:
                        raise RuntimeError('OCR 返回了文字，但缺少置信度或文字位置')
                    for txt, score, box in zip(result.txts, result.scores, result.boxes, strict=True):
                        pts = np.asarray(box)
                        lines.append(Line(str(txt), float(score),
                            (float(pts[:, 0].min())/scale+rect.x0)/width,
                            (float(pts[:, 1].min())/scale+rect.y0)/height,
                            (float(pts[:, 0].max())/scale+rect.x0)/width,
                            (float(pts[:, 1].max())/scale+rect.y0)/height))
        lines.sort(key=lambda line: (line.cy, line.x0))
        self.cache.mkdir(parents=True, exist_ok=True)
        target.write_text(json.dumps([asdict(line) for line in lines], ensure_ascii=False, indent=2), 'utf-8')
        return lines


def rows(lines: list[Line]) -> list[list[Line]]:
    """按文字高度聚合同一行的表格单元格，避免混用上一行或下一行数值。"""
    groups: list[list[Line]] = []
    for line in sorted(lines, key=lambda item: (item.cy, item.x0)):
        if groups and abs(line.cy - sum(x.cy for x in groups[-1])/len(groups[-1])) <= max(
                0.004, min(line.y1-line.y0, groups[-1][0].y1-groups[-1][0].y0)*0.6):
            groups[-1].append(line)
        else:
            groups.append([line])
    return [sorted(group, key=lambda item: item.x0) for group in groups]
