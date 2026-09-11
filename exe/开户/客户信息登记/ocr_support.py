"""本地 PDF OCR，保留位置、置信度和可重复使用的识别缓存。"""
from __future__ import annotations

from dataclasses import asdict, dataclass
import hashlib
import json
import os
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
        self.cache = Path(os.environ['ACCOUNT_OCR_CACHE']) if os.environ.get('ACCOUNT_OCR_CACHE') else cache
        self.engine = None
        self.hashes: dict[Path, str] = {}

    def page(self, path: Path, index: int, crop=None, scale=2.5, rotation=0,
             use_cls=True, suppress_red=False, force_ocr=False) -> list[Line]:
        if path not in self.hashes:
            self.hashes[path] = hashlib.sha256(path.read_bytes()).hexdigest()
        key = hashlib.sha256(
            repr(('v2', self.hashes[path], index, crop, scale, rotation,
                  use_cls, suppress_red) + (('force_ocr',) if force_ocr else ())).encode()
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
            words = [] if force_ocr else page.get_text('words')
            if words:
                # PDF 文本坐标不随页面 rotation 改变，须转换到显示坐标后再裁剪。
                lines = []
                for word in words:
                    box = pymupdf.Rect(word[:4]) * page.rotation_matrix
                    if rect.contains(box):
                        lines.append(Line(word[4], 1.0, box.x0/width, box.y0/height,
                                          box.x1/width, box.y1/height))
            else:
                if self.engine is None:
                    self.engine = create_ocr_engine()
                print(f'OCR：{path.name} 第 {index+1} 页 / 旋转 {rotation}°'
                      + ('（局部）' if crop else ''), flush=True)
                pix = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale),
                                      clip=rect, colorspace=pymupdf.csRGB, alpha=False)
                arr = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, 3)
                if suppress_red:
                    # 红色印章在红通道中接近白色；仅用于补读被印章遮挡的表头。
                    arr = np.repeat(arr[:, :, :1], 3, axis=2)
                result = self.engine(arr, use_cls=use_cls)
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

    def oriented_page(self, path: Path, index: int, keywords=(), force_ocr=False) -> tuple[list[Line], int]:
        """用文字方向及可读关键词选择页面方向，不依据纸张横竖或固定页码。"""
        lines = self.page(path, index, use_cls=False, force_ocr=force_ocr)

        def quality(items):
            horizontal = [x for x in items if x.x1-x.x0 > x.y1-x.y0]
            return (sum(len(x.text)*x.score**3 for x in horizontal)
                    + 15*sum(word in x.text for x in horizontal for word in keywords))

        total = sum(len(x.text) for x in lines)
        tall = sum(len(x.text) for x in lines if x.y1-x.y0 > (x.x1-x.x0)*1.5)
        if (not total or tall > total*.20 or quality(lines) < 12
                or (keywords and not any(w in x.text for x in lines for w in keywords))):
            choices = [(lines, 0)]
            for rotation in (90, 270, 180):
                choices.append((self.page(path, index, rotation=rotation, use_cls=False,
                                          force_ocr=force_ocr), rotation))
            return max(choices, key=lambda item: quality(item[0]))
        return lines, 0


def create_ocr_engine():
    """桌面版明确指定随软件分发的模型，避免运行时下载或写入安装目录。"""
    from rapidocr import RapidOCR
    model_dir = os.environ.get('ACCOUNT_OCR_MODEL_DIR')
    if not model_dir:
        return RapidOCR()
    params = {'Global.log_level': 'warning'}
    for kind, name in [('Det', 'PP-OCRv6_det_small.onnx'),
                       ('Cls', 'ch_ppocr_mobile_v2.0_cls_mobile.onnx'),
                       ('Rec', 'PP-OCRv6_rec_small.onnx')]:
        path = Path(model_dir)/name
        if not path.is_file():
            raise RuntimeError(f'缺少识别模型 {name}，请重新解压完整软件包')
        params[f'{kind}.model_path'] = str(path)
    return RapidOCR(params=params)


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
