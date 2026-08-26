"""远期了结处理流程共用的安全归档工具。"""

from __future__ import annotations

import re
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class ArchiveResult:
    path: Path
    file_count: int


def _next_batch_path(destination_dir: Path, label: str, date_text: str) -> Path:
    destination_dir.mkdir(parents=True, exist_ok=True)
    pattern = re.compile(
        rf"^{re.escape(label)}{re.escape(date_text)}批次(\d+)\.zip$"
    )
    batches = []
    for path in destination_dir.iterdir():
        match = pattern.fullmatch(path.name)
        if path.is_file() and match:
            batches.append(int(match.group(1)))
    batch = max(batches, default=0) + 1
    width = max(2, len(str(batch)))
    return destination_dir / f"{label}{date_text}批次{batch:0{width}d}.zip"


def _source_files(source_dir: Path) -> list[Path]:
    if not source_dir.exists():
        source_dir.mkdir(parents=True, exist_ok=True)
        return []
    if not source_dir.is_dir():
        raise RuntimeError(f"待归档路径不是文件夹：{source_dir}")

    paths = list(source_dir.rglob("*"))
    symlinks = [path for path in paths if path.is_symlink()]
    if symlinks:
        names = "、".join(str(path) for path in symlinks[:5])
        raise RuntimeError(f"待归档目录中存在符号链接，为避免误操作已停止：{names}")
    return [path for path in paths if path.is_file()]


def create_archive(
    source_dir: Path,
    destination_dir: Path,
    label: str,
    date_text: str,
) -> ArchiveResult | None:
    """将 source_dir 的全部内容压缩到新批次 zip，但不删除源文件。"""
    source_dir = source_dir.expanduser().resolve()
    destination_dir = destination_dir.expanduser().resolve()
    if source_dir == destination_dir or source_dir in destination_dir.parents:
        raise RuntimeError("压缩包目录不能位于待归档目录内")

    files = _source_files(source_dir)
    if not files:
        return None

    destination = _next_batch_path(destination_dir, label, date_text)
    temporary = destination.with_name(destination.name + ".tmp")
    if temporary.exists():
        temporary.unlink()

    try:
        with zipfile.ZipFile(
            temporary, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6
        ) as archive:
            for path in files:
                archive.write(path, path.relative_to(source_dir))

        with zipfile.ZipFile(temporary, "r") as archive:
            bad_file = archive.testzip()
            if bad_file is not None:
                raise RuntimeError(f"压缩包校验失败：{bad_file}")
            if len(archive.infolist()) != len(files):
                raise RuntimeError("压缩包内文件数量与源目录不一致")
        temporary.replace(destination)
    except Exception:
        if temporary.exists():
            temporary.unlink()
        raise

    return ArchiveResult(destination, len(files))


def archive_and_clear(
    source_dir: Path,
    destination_dir: Path,
    label: str,
    date_text: str,
) -> ArchiveResult | None:
    """压缩并校验成功后，清空 source_dir 中的旧内容。"""
    result = create_archive(source_dir, destination_dir, label, date_text)
    if result is None:
        return None

    # 只在 zip 已经完成且通过校验后才删除源内容。
    for entry in source_dir.iterdir():
        if entry.is_symlink() or entry.is_file():
            entry.unlink()
        elif entry.is_dir():
            shutil.rmtree(entry)
        else:
            raise RuntimeError(f"无法识别待删除项目的类型：{entry}")
    return result
