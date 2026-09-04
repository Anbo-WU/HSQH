#!/usr/bin/env python3
"""识别每份拆分 PDF 首页的交易编号，并按规则安全重命名。

例：
【HFSY】0009-FWJY-2026072401
-> 【HFSY】0009-FWJY-202607240120260724.pdf

特殊分段编号也会保留：
【HFSY】0009-FWJY-2026072401-1
-> 【HFSY】0009-FWJY-2026072401-120260724.pdf

Windows 版使用本地 RapidOCR 和 ONNX Runtime，不依赖 macOS Vision。
"""

from __future__ import annotations

import argparse
import re
import sys
import unicodedata
import uuid
from dataclasses import dataclass
from pathlib import Path

from windows_ocr import recognize_first_pages


OCR_SEPARATOR = r"[\s._·•,，:：/\\\-‐‑‒–—−]*"
OCR_SUFFIX_SEPARATOR = r"[\s._·•,，:：/\\]*[-‐‑‒–—−]+[\s._·•,，:：/\\]*"
TRANSACTION_PATTERN = re.compile(
    r"(?<![0-9A-Z])([0-9OQILSZBGD|]{4})"
    + OCR_SEPARATOR
    + r"([A-Z]"
    + OCR_SEPARATOR
    + r"[A-Z](?:"
    + OCR_SEPARATOR
    + r"[A-Z]"
    + OCR_SEPARATOR
    + r"[A-Z])?)"
    + OCR_SEPARATOR
    + r"([0-9OQILSZBGD|]{10})"
    + r"(?:"
    + OCR_SUFFIX_SEPARATOR
    + r"([0-9OQILSZBGD|]+)"
    + r")?"
    + r"(?![0-9A-Z])",
    flags=re.IGNORECASE,
)
OCR_DIGIT_TRANSLATION = str.maketrans(
    {
        "O": "0",
        "Q": "0",
        "D": "0",
        "I": "1",
        "L": "1",
        "|": "1",
        "Z": "2",
        "S": "5",
        "G": "6",
        "B": "8",
    }
)


@dataclass(frozen=True)
class RenamePlan:
    source: Path
    target: Path
    transaction_number: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("folder", type=Path, help="要识别并重命名的拆分 PDF 目录")
    parser.add_argument(
        "--preview",
        action="store_true",
        help="只识别并显示新旧文件名，不执行重命名",
    )
    return parser.parse_args()


def collect_pdfs(folder: Path) -> list[Path]:
    return sorted(
        (
            path
            for path in folder.rglob("*")
            if path.is_file() and path.suffix.casefold() == ".pdf"
        ),
        key=lambda path: str(path.relative_to(folder)).casefold(),
    )


def recognize_all(pdfs: list[Path]) -> dict[Path, str]:
    return recognize_first_pages(pdfs)


def normalize_text(text: str) -> str:
    normalized = unicodedata.normalize("NFKC", text).casefold()
    return "".join(character for character in normalized if character.isalnum())


def expected_trade_type(text: str) -> str | None:
    normalized = normalize_text(text)
    if "远期交易确认书" in normalized:
        return "FWJY"
    if "期权交易确认书" in normalized:
        return "JY"
    return None


def normalize_trade_type(raw_type: str, expected_type: str | None) -> str | None:
    trade_type = re.sub(r"[^A-Z]", "", raw_type.upper())
    if trade_type in {"JY", "FWJY"}:
        return trade_type
    if (
        expected_type in {"JY", "FWJY"}
        and len(trade_type) == len(expected_type)
        and sum(left != right for left, right in zip(trade_type, expected_type)) <= 1
    ):
        return expected_type
    return None


def extract_transaction_number(text: str, pdf_name: str) -> str:
    matches = {
        (
            match.group(1),
            match.group(2).upper(),
            match.group(3),
            match.group(4) or "",
        )
        for match in TRANSACTION_PATTERN.finditer(text)
    }
    if len(matches) != 1:
        detail = "未识别到" if not matches else "识别到多个候选编号"
        raise ValueError(f"{pdf_name}: {detail}")

    company_code, raw_trade_type, serial, raw_suffix = matches.pop()
    company_code = company_code.upper().translate(OCR_DIGIT_TRANSLATION)
    trade_type = normalize_trade_type(raw_trade_type, expected_trade_type(text))
    serial = serial.upper().translate(OCR_DIGIT_TRANSLATION)
    suffix = raw_suffix.upper().translate(OCR_DIGIT_TRANSLATION)
    if (
        not company_code.isdigit()
        or trade_type is None
        or not serial.isdigit()
        or (suffix and not suffix.isdigit())
    ):
        raise ValueError(f"{pdf_name}: 交易编号包含无法安全纠正的字符")
    suffix_text = f"-{suffix}" if suffix else ""
    return f"【HFSY】{company_code}-{trade_type}-{serial}{suffix_text}"


def build_plans(pdfs: list[Path], recognized: dict[Path, str]) -> list[RenamePlan]:
    plans: list[RenamePlan] = []
    errors: list[str] = []
    for source in pdfs:
        try:
            number = extract_transaction_number(
                recognized[source.resolve()],
                source.name,
            )
            # 主编号中的 10 位流水号是 YYYYMMDDNN；末尾可能还有 -1、-2
            # 等项目分段，不能用最后一个短横线直接切流水号。
            serial_match = re.search(r"-(\d{10})(?:-\d+)?$", number)
            if serial_match is None:
                raise ValueError(f"{source.name}: 无法从交易编号提取日期")
            date_suffix = serial_match.group(1)[:-2]
            target = source.with_name(f"{number}{date_suffix}.pdf")
            plans.append(RenamePlan(source, target, number))
        except (KeyError, ValueError) as exc:
            errors.append(str(exc))
    if errors:
        raise RuntimeError("编号提取失败：\n  " + "\n  ".join(errors))

    target_keys: dict[str, Path] = {}
    for plan in plans:
        key = str(plan.target.resolve()).casefold()
        if key in target_keys and target_keys[key] != plan.source:
            raise RuntimeError(
                f"新文件名重复：{plan.target.name}\n"
                f"  {target_keys[key].name}\n  {plan.source.name}"
            )
        target_keys[key] = plan.source

    source_keys = {str(plan.source.resolve()).casefold() for plan in plans}
    for plan in plans:
        target_key = str(plan.target.resolve()).casefold()
        if plan.target.exists() and target_key not in source_keys:
            raise RuntimeError(f"目标文件已存在：{plan.target}")
    return plans


def apply_plans(plans: list[RenamePlan]) -> None:
    """分两阶段改名，避免文件名互相占用；失败时尽量回滚。"""
    active = [plan for plan in plans if plan.source != plan.target]
    temporary: dict[RenamePlan, Path] = {}
    completed: list[RenamePlan] = []
    try:
        for plan in active:
            temp_path = plan.source.with_name(
                f".{plan.source.name}.rename-{uuid.uuid4().hex}.tmp"
            )
            plan.source.rename(temp_path)
            temporary[plan] = temp_path
        for plan in active:
            temporary[plan].rename(plan.target)
            completed.append(plan)
    except Exception:
        for plan in reversed(completed):
            if plan.target.exists() and not plan.source.exists():
                plan.target.rename(plan.source)
        for plan, temp_path in temporary.items():
            if temp_path.exists() and not plan.source.exists():
                temp_path.rename(plan.source)
        raise


def run_scan(pdf_folder: Path, preview: bool = False) -> tuple[int, int]:
    """识别并重命名拆分件，返回（PDF 总数，实际改名数量）。"""
    pdf_folder = pdf_folder.expanduser().resolve()
    if not pdf_folder.is_dir():
        raise RuntimeError(f"找不到文件夹：{pdf_folder}")
    pdfs = collect_pdfs(pdf_folder)
    if not pdfs:
        raise RuntimeError(f"未在 {pdf_folder} 中找到 PDF。")

    print(f"正在识别 {len(pdfs)} 份 PDF 的首页，请稍候……")
    recognized = recognize_all(pdfs)
    plans = build_plans(pdfs, recognized)
    print()
    for number, plan in enumerate(plans, start=1):
        status = "（已是目标名称）" if plan.source == plan.target else ""
        print(f"{number:>3}. {plan.source.name}")
        print(f"     -> {plan.target.name}{status}")
    if preview:
        print(f"\n预览完成：共 {len(plans)} 份，未重命名任何文件。")
        return len(plans), 0

    apply_plans(plans)
    changed = sum(plan.source != plan.target for plan in plans)
    print(f"\n已完成：成功重命名 {changed} 份 PDF，PDF 内容未修改。")
    return len(plans), changed


def main() -> int:
    args = parse_args()
    try:
        run_scan(args.folder, preview=args.preview)
        return 0
    except (OSError, RuntimeError, ValueError) as exc:
        print(f"\n已停止，未重命名任何文件：\n{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
