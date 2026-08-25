#!/usr/bin/env python3
"""按人员统一运行确认书处理流程。"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import time
import uuid
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Callable, TypeVar


PROJECT_ROOT = Path(__file__).resolve().parent
CODE_FOLDER = PROJECT_ROOT / "code"
TEACHERS = ("LuX", "Tan", "Zhang", "LuT")

# 功能脚本仍可独立运行；总入口通过它们公开的函数传入本次任务路径。
sys.path.insert(0, str(CODE_FOLDER))
import merge as merge_program  # noqa: E402
import precheck as precheck_program  # noqa: E402
import registration as registration_program  # noqa: E402
import scan as scan_program  # noqa: E402
import split as split_program  # noqa: E402


T = TypeVar("T")


@dataclass(frozen=True)
class TaskPaths:
    teacher: str
    task_date: date
    scan_date: date
    batch: int
    teacher_folder: Path
    confirmation_folder: Path
    registration_folder: Path
    merged_pdf_folder: Path
    scanned_folder: Path
    scan_source_pdf: Path
    split_output_folder: Path
    template_path: Path
    registration_output: Path
    merged_pdf_output: Path
    desktop_registration_output: Path
    desktop_merged_pdf_output: Path
    desktop_scan_zip_output: Path
    log_path: Path


def output_paths(
    teacher: str,
    task_date: date,
    batch: int,
) -> tuple[Path, Path, Path, Path]:
    teacher_folder = PROJECT_ROOT / teacher
    date_text = task_date.strftime("%Y%m%d")
    batch_text = f"{batch:02d}"
    registration_name = f"登记表_{teacher}_{date_text}_{batch_text}.xls"
    merged_pdf_name = f"确认书合并_{teacher}_{date_text}_{batch_text}.pdf"
    desktop = Path.home() / "Desktop"
    return (
        teacher_folder / "登记表文件" / registration_name,
        teacher_folder / "合并PDF" / merged_pdf_name,
        desktop / registration_name,
        desktop / merged_pdf_name,
    )


def used_batches(teacher: str, task_date: date) -> set[int]:
    """从两类正式输出中提取已经出现过的批次。"""
    teacher_folder = PROJECT_ROOT / teacher
    date_text = task_date.strftime("%Y%m%d")
    patterns = (
        (
            teacher_folder / "登记表文件",
            re.compile(
                rf"^登记表_{re.escape(teacher)}_{date_text}_(\d+)\.xls$",
                flags=re.IGNORECASE,
            ),
        ),
        (
            teacher_folder / "合并PDF",
            re.compile(
                rf"^确认书合并_{re.escape(teacher)}_{date_text}_(\d+)\.pdf$",
                flags=re.IGNORECASE,
            ),
        ),
    )
    batches: set[int] = set()
    for folder, pattern in patterns:
        if not folder.is_dir():
            continue
        for path in folder.iterdir():
            match = pattern.fullmatch(path.name)
            if path.is_file() and match is not None:
                batches.add(int(match.group(1)))
    return batches


def choose_a_batch(teacher: str, task_date: date) -> int:
    """优先续跑未完成批次，否则返回下一个新批次。"""
    batches = used_batches(teacher, task_date)
    for batch in sorted(batches):
        if not batch_outputs_complete(teacher, task_date, batch):
            return batch
    return max(batches, default=0) + 1


def batch_outputs_complete(teacher: str, task_date: date, batch: int) -> bool:
    """以人员目录中的两份正式输出判断 A 功能是否完成。"""
    registration_output, merged_pdf_output, _, _ = output_paths(
        teacher, task_date, batch
    )
    return registration_output.is_file() and merged_pdf_output.is_file()


def completed_batches(teacher: str, task_date: date) -> set[int]:
    completed: set[int] = set()
    for batch in used_batches(teacher, task_date):
        if batch_outputs_complete(teacher, task_date, batch):
            completed.add(batch)
    return completed


def latest_completed_batch(teacher: str) -> tuple[date, int] | None:
    """为 B 功能查找该老师最近完成的 A 批次。"""
    candidates: list[tuple[date, int]] = []
    registration_folder = PROJECT_ROOT / teacher / "登记表文件"
    pattern = re.compile(
        rf"^登记表_{re.escape(teacher)}_(\d{{8}})_(\d+)\.xls$",
        flags=re.IGNORECASE,
    )
    if registration_folder.is_dir():
        for path in registration_folder.iterdir():
            match = pattern.fullmatch(path.name)
            if not path.is_file() or match is None:
                continue
            task_date = datetime.strptime(match.group(1), "%Y%m%d").date()
            batch = int(match.group(2))
            if batch_outputs_complete(teacher, task_date, batch):
                candidates.append((task_date, batch))
    if not candidates:
        return None
    return max(candidates)


def build_task_paths(
    teacher: str,
    task_date: date,
    batch: int,
    scan_date: date | None = None,
) -> TaskPaths:
    if batch < 1:
        raise RuntimeError("批次必须是大于 0 的整数。")
    teacher_folder = PROJECT_ROOT / teacher
    confirmation_folder = teacher_folder / "确认书文件"
    registration_folder = teacher_folder / "登记表文件"
    merged_pdf_folder = teacher_folder / "合并PDF"
    scanned_folder = teacher_folder / "确认书扫描"
    template_path = PROJECT_ROOT / "登记表模版.xls"
    desktop_folder = Path.home() / "Desktop"
    effective_scan_date = scan_date or date.today()
    split_folder_name = (
        f"扫描件_{effective_scan_date:%m%d}_{teacher}_{batch:02d}"
    )

    required = [
        (teacher_folder, "人员目录"),
        (confirmation_folder, "确认书文件目录"),
        (registration_folder, "登记表文件目录"),
        (merged_pdf_folder, "合并PDF目录"),
        (scanned_folder, "确认书扫描目录"),
        (template_path, "登记表模板"),
        (desktop_folder, "桌面目录"),
    ]
    missing = [f"{description}：{path}" for path, description in required if not path.exists()]
    if missing:
        raise RuntimeError("项目目录不完整：\n  " + "\n  ".join(missing))

    (
        registration_output,
        merged_pdf_output,
        desktop_registration_output,
        desktop_merged_pdf_output,
    ) = output_paths(teacher, task_date, batch)
    return TaskPaths(
        teacher=teacher,
        task_date=task_date,
        scan_date=effective_scan_date,
        batch=batch,
        teacher_folder=teacher_folder,
        confirmation_folder=confirmation_folder,
        registration_folder=registration_folder,
        merged_pdf_folder=merged_pdf_folder,
        scanned_folder=scanned_folder,
        scan_source_pdf=scanned_folder / f"{effective_scan_date:%m%d}.pdf",
        split_output_folder=scanned_folder / split_folder_name,
        template_path=template_path,
        registration_output=registration_output,
        merged_pdf_output=merged_pdf_output,
        desktop_registration_output=desktop_registration_output,
        desktop_merged_pdf_output=desktop_merged_pdf_output,
        desktop_scan_zip_output=desktop_folder / f"{split_folder_name}.zip",
        log_path=teacher_folder / "运行记录.json",
    )


class RunLogger:
    def __init__(self, paths: TaskPaths, flow: str) -> None:
        self.paths = paths
        self.flow = flow
        self.run_id = datetime.now().astimezone().strftime("%Y%m%d_%H%M%S_%f")
        self.started_at = datetime.now().astimezone()
        self.started_timer = time.perf_counter()
        self.events: list[dict[str, object]] = []
        self.migrate_legacy_if_needed()

    def run_step(
        self,
        program: str,
        action: Callable[[], T],
        detail: str = "",
    ) -> T:
        print(f"\n{'=' * 18} {program} {'=' * 18}")
        started_at = datetime.now().astimezone()
        started_timer = time.perf_counter()
        status = "success"
        error = ""
        try:
            return action()
        except Exception as exc:
            status = "failed"
            error = str(exc)
            raise
        finally:
            finished_at = datetime.now().astimezone()
            elapsed = time.perf_counter() - started_timer
            event: dict[str, object] = {
                "program": program,
                "status": status,
                "started_at": started_at.isoformat(timespec="seconds"),
                "finished_at": finished_at.isoformat(timespec="seconds"),
                "elapsed_seconds": round(elapsed, 3),
            }
            if detail:
                event["detail"] = detail
            if error:
                event["error"] = error
            self.events.append(event)
            print(f"[运行时间] {program}：{elapsed:.3f} 秒（{status}）")

    def save(self, status: str) -> None:
        finished_at = datetime.now().astimezone()
        elapsed = time.perf_counter() - self.started_timer
        record = {
            "run_id": self.run_id,
            "teacher": self.paths.teacher,
            "flow": self.flow,
            "task_date": self.paths.task_date.isoformat(),
            "scan_date": self.paths.scan_date.isoformat(),
            "batch": f"{self.paths.batch:02d}",
            "status": status,
            "started_at": self.started_at.isoformat(timespec="seconds"),
            "finished_at": finished_at.isoformat(timespec="seconds"),
            "elapsed_seconds": round(elapsed, 3),
            "paths": {
                "confirmation_folder": str(self.paths.confirmation_folder),
                "registration_output": str(self.paths.registration_output),
                "merged_pdf_output": str(self.paths.merged_pdf_output),
                "desktop_registration_output": str(
                    self.paths.desktop_registration_output
                ),
                "desktop_merged_pdf_output": str(
                    self.paths.desktop_merged_pdf_output
                ),
                "scanned_folder": str(self.paths.scanned_folder),
                "scan_source_pdf": str(self.paths.scan_source_pdf),
                "split_output_folder": str(self.paths.split_output_folder),
                "desktop_scan_zip_output": str(
                    self.paths.desktop_scan_zip_output
                ),
            },
            "programs": self.events,
        }
        records = self.load_records()
        records.append(record)
        self.write_records(records)
        print(f"\n[总运行时间] {elapsed:.3f} 秒")
        print(f"[运行记录] {self.paths.log_path}")

    def write_records(self, records: list[dict[str, object]]) -> None:
        """格式化并原子写入完整 JSON 历史记录。"""
        temporary: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(
                prefix=".运行记录-",
                suffix=".json",
                dir=self.paths.teacher_folder,
                delete=False,
            ) as temp_file:
                temporary = Path(temp_file.name)
            temporary.write_text(
                json.dumps(records, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            os.replace(temporary, self.paths.log_path)
            temporary = None

            # 首次使用新版日志时，成功写入 JSON 后再移除旧 JSONL。
            legacy_path = self.paths.teacher_folder / "运行记录.jsonl"
            if legacy_path.exists():
                legacy_path.unlink()
        finally:
            if temporary is not None and temporary.exists():
                temporary.unlink()

    def migrate_legacy_if_needed(self) -> None:
        legacy_path = self.paths.teacher_folder / "运行记录.jsonl"
        if legacy_path.exists():
            self.write_records(self.load_records())

    def load_records(self) -> list[dict[str, object]]:
        """读取格式化 JSON，并兼容迁移旧版 JSONL 历史记录。"""
        records: list[dict[str, object]] = []
        if self.paths.log_path.exists():
            try:
                loaded = json.loads(self.paths.log_path.read_text(encoding="utf-8"))
            except json.JSONDecodeError as exc:
                raise RuntimeError(
                    f"运行记录不是有效 JSON：{self.paths.log_path}"
                ) from exc
            if not isinstance(loaded, list) or not all(
                isinstance(item, dict) for item in loaded
            ):
                raise RuntimeError(
                    f"运行记录必须是 JSON 数组：{self.paths.log_path}"
                )
            records.extend(loaded)

        legacy_path = self.paths.teacher_folder / "运行记录.jsonl"
        if legacy_path.exists():
            try:
                for line_number, line in enumerate(
                    legacy_path.read_text(encoding="utf-8").splitlines(),
                    start=1,
                ):
                    if not line.strip():
                        continue
                    item = json.loads(line)
                    if not isinstance(item, dict):
                        raise RuntimeError(
                            f"旧运行记录第 {line_number} 行不是 JSON 对象。"
                        )
                    records.append(item)
            except json.JSONDecodeError as exc:
                raise RuntimeError(
                    f"旧运行记录不是有效 JSONL：{legacy_path}"
                ) from exc

        # 防止迁移中断后下次重复导入相同运行记录。
        deduplicated: list[dict[str, object]] = []
        seen_run_ids: set[object] = set()
        for item in records:
            run_id = item.get("run_id")
            if run_id is not None and run_id in seen_run_ids:
                continue
            if run_id is not None:
                seen_run_ids.add(run_id)
            deduplicated.append(item)
        return deduplicated


def choose_teacher(value: str | None) -> str:
    if value is not None:
        matches = {teacher.casefold(): teacher for teacher in TEACHERS}
        selected = matches.get(value.strip().casefold())
        if selected is None:
            raise RuntimeError(
                f"未知人员：{value}。可选人员：{', '.join(TEACHERS)}"
            )
        return selected

    print("请选择本次处理人员：")
    for number, teacher in enumerate(TEACHERS, start=1):
        print(f"  {number}. {teacher}")
    answer = input("输入序号或人员文件夹名称：").strip()
    if answer.isdigit() and 1 <= int(answer) <= len(TEACHERS):
        return TEACHERS[int(answer) - 1]
    return choose_teacher(answer)


def choose_flow(value: str | None) -> str:
    if value is not None:
        return value
    print("\n请选择流程：")
    print("  A. 盖章前：precheck → registration → merge")
    print("  B. 盖章后：split → scan")
    answer = input("输入 A 或 B：").strip().upper()
    if answer in {"A", "B"}:
        return answer
    raise RuntimeError(f"无效流程选项：{answer}")


def issue_action(policy: str) -> str:
    if policy == "stop":
        return "stop"
    if policy == "continue":
        return "continue"
    if not sys.stdin.isatty():
        print("当前不是交互终端，发现异常后默认 stop。")
        return "stop"

    while True:
        answer = input(
            "\n预检查发现异常，请选择 stop / recheck / continue："
        ).strip().casefold()
        aliases = {
            "s": "stop",
            "stop": "stop",
            "r": "recheck",
            "recheck": "recheck",
            "c": "continue",
            "continue": "continue",
        }
        selected = aliases.get(answer)
        if selected is not None:
            return selected
        print("请输入 stop、recheck 或 continue。")


def copy_output(source: Path, target: Path) -> None:
    """先复制到桌面临时文件，完成后再原子替换正式副本。"""
    if not source.is_file():
        raise RuntimeError(f"找不到需要复制的输出文件：{source}")
    temporary = target.with_name(f".{target.name}.copy-{uuid.uuid4().hex}.tmp")
    try:
        shutil.copy2(source, temporary)
        os.replace(temporary, target)
    finally:
        if temporary.exists():
            temporary.unlink()
    print(f"桌面副本：{target}")


def create_desktop_zip(source_folder: Path, target_zip: Path) -> None:
    """把扫描件子文件夹压缩到桌面，保留但不打包拆分记录。"""
    if not source_folder.is_dir():
        raise RuntimeError(f"找不到需要压缩的扫描件目录：{source_folder}")
    pdfs = sorted(
        (
            path
            for path in source_folder.rglob("*")
            if path.is_file() and path.suffix.casefold() == ".pdf"
        ),
        key=lambda path: str(path.relative_to(source_folder)).casefold(),
    )
    if not pdfs:
        raise RuntimeError(f"扫描件目录中没有可压缩的 PDF：{source_folder}")

    temporary = target_zip.with_name(
        f".{target_zip.name}.zip-{uuid.uuid4().hex}.tmp"
    )
    try:
        with zipfile.ZipFile(
            temporary,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=6,
        ) as archive:
            archive.writestr(f"{source_folder.name}/", b"")
            for path in sorted(
                (
                    item
                    for item in source_folder.rglob("*")
                    if item.is_file() and item.name != "拆分记录.json"
                ),
                key=lambda item: str(item.relative_to(source_folder)).casefold(),
            ):
                archive.write(
                    path,
                    arcname=str(
                        Path(source_folder.name) / path.relative_to(source_folder)
                    ),
                )

        with zipfile.ZipFile(temporary, "r") as archive:
            if archive.testzip() is not None:
                raise RuntimeError("桌面 ZIP 完整性检查失败。")
            archived_pdf_count = sum(
                Path(name).suffix.casefold() == ".pdf"
                for name in archive.namelist()
            )
        if archived_pdf_count != len(pdfs):
            raise RuntimeError(
                f"桌面 ZIP 中有 {archived_pdf_count} 个 PDF，"
                f"原扫描件目录中有 {len(pdfs)} 个。"
            )
        os.replace(temporary, target_zip)
    finally:
        if temporary.exists():
            temporary.unlink()
    print(f"桌面压缩包：{target_zip}")


def safe_clear_split_output(paths: TaskPaths) -> None:
    """只允许删除本次 B 功能对应的直接子目录。"""
    target = paths.split_output_folder.resolve()
    parent = paths.scanned_folder.resolve()
    expected_name = f"扫描件_{paths.scan_date:%m%d}_{paths.teacher}_{paths.batch:02d}"
    if target.parent != parent or target.name != expected_name:
        raise RuntimeError(f"拒绝清理未通过路径校验的目录：{target}")
    if target.is_symlink():
        raise RuntimeError(f"拒绝清理符号链接目录：{target}")
    if target.exists():
        if not target.is_dir():
            raise RuntimeError(f"拆分输出路径不是文件夹：{target}")
        shutil.rmtree(target)
        print(f"已清空上一次拆分目录：{target}")
    else:
        print("本批次尚无拆分目录，将直接重新拆分。")


def split_issue_action(policy: str) -> str:
    if policy == "stop":
        return "stop"
    if policy == "continue":
        return "continue"
    if not sys.stdin.isatty():
        print("当前不是交互终端，split 异常后默认 stop。")
        return "stop"

    while True:
        answer = input(
            "\nsplit 输出异常，请选择 stop / resplit / continue："
        ).strip().casefold()
        aliases = {
            "s": "stop",
            "stop": "stop",
            "r": "resplit",
            "resplit": "resplit",
            "c": "continue",
            "continue": "continue",
        }
        selected = aliases.get(answer)
        if selected is not None:
            return selected
        print("请输入 stop、resplit 或 continue。")


def run_before_flow(
    paths: TaskPaths,
    logger: RunLogger,
    on_issue: str,
    dry_run: bool,
) -> str:
    attempt = 1
    while True:
        def precheck_action() -> int:
            report, issue_count = precheck_program.检查PDF(paths.confirmation_folder)
            print(report)
            return issue_count

        issue_count = logger.run_step(
            "precheck.py",
            precheck_action,
            detail=f"第 {attempt} 次检查",
        )
        if issue_count == 0:
            print("\n预检查通过，继续生成登记表。")
            break

        print(f"\n预检查共检测到 {issue_count} 项/组异常。")
        action = issue_action(on_issue)
        if action == "stop":
            print("流程已暂停。人工处理完成后，请重新运行本流程。")
            return "stopped"
        if action == "recheck":
            attempt += 1
            print("即将重新检查当前目录……")
            continue
        print("已选择 continue，将带着当前异常继续运行。")
        break

    def registration_action() -> tuple[int, int]:
        result = registration_program.run_registration(
            paths.confirmation_folder,
            paths.template_path,
            paths.registration_output,
            paths.task_date,
            preview=dry_run,
        )
        if dry_run:
            print(f"计划桌面副本：{paths.desktop_registration_output}")
        else:
            copy_output(
                paths.registration_output,
                paths.desktop_registration_output,
            )
        return result

    logger.run_step(
        "registration.py",
        registration_action,
    )

    def merge_action() -> tuple[int, int]:
        result = merge_program.run_merge(
            paths.confirmation_folder,
            paths.merged_pdf_output,
            dry_run=dry_run,
        )
        if dry_run:
            print(f"计划桌面副本：{paths.desktop_merged_pdf_output}")
        else:
            copy_output(paths.merged_pdf_output, paths.desktop_merged_pdf_output)
        return result

    logger.run_step(
        "merge.py",
        merge_action,
    )
    return "dry_run" if dry_run else "success"


def run_after_flow(
    paths: TaskPaths,
    logger: RunLogger,
    dry_run: bool,
    on_split_issue: str,
) -> str:
    expected_count = precheck_program.统计PDF数量(paths.confirmation_folder)
    if expected_count < 1:
        raise RuntimeError(
            f"{paths.confirmation_folder} 中没有原始确认书，"
            "无法校验拆分后的 PDF 数量。"
        )
    print(f"原始确认书数量：{expected_count}")

    while True:
        try:
            def split_action() -> tuple[int, int]:
                result = split_program.run_split(
                    paths.scanned_folder,
                    paths.split_output_folder,
                    expected_count,
                    preview=dry_run,
                    scan_date=paths.scan_date,
                )
                if result[0] != expected_count:
                    raise RuntimeError(
                        f"split.py 返回 {result[0]} 份，预期 {expected_count} 份。"
                    )
                return result

            split_count, _ = logger.run_step("split.py", split_action)
            break
        except (OSError, RuntimeError, ValueError) as exc:
            print(f"\nsplit.py 异常：{exc}", file=sys.stderr)
            should_resplit = False
            while True:
                action = split_issue_action(on_split_issue)
                if action == "stop":
                    print("B 功能已停止，scan.py 未运行。")
                    return "stopped"
                if action == "resplit":
                    safe_clear_split_output(paths)
                    should_resplit = True
                    break

                available_count = sum(
                    path.is_file() and path.suffix.casefold() == ".pdf"
                    for path in paths.split_output_folder.rglob("*")
                ) if paths.split_output_folder.is_dir() else 0
                if available_count < 1:
                    print(
                        "本次异常前没有生成任何可安全使用的拆分 PDF，"
                        "因此无法 continue；请选择 resplit 或 stop。"
                    )
                    if on_split_issue != "ask":
                        return "stopped"
                    continue
                split_count = available_count
                issue_count = 0
                manifest_path = paths.split_output_folder / "拆分记录.json"
                if manifest_path.is_file():
                    try:
                        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
                        issues = manifest.get("issues", [])
                        if isinstance(issues, list):
                            issue_count = len(issues)
                    except (OSError, json.JSONDecodeError):
                        pass
                print(
                    f"已选择 continue：将跳过 {issue_count} 项拆分异常，"
                    f"使用其余 {available_count} 个 PDF 继续执行 scan.py。"
                )
                break
            if should_resplit:
                continue
            break

    if dry_run:
        print(f"计划 scan 目录：{paths.split_output_folder}")
        print(f"计划桌面压缩包：{paths.desktop_scan_zip_output}")
        print("B 功能预览完成：dry-run 模式下不运行 scan.py 或生成 ZIP。")
        return "dry_run"

    def scan_action() -> tuple[int, int]:
        result = scan_program.run_scan(paths.split_output_folder)
        create_desktop_zip(
            paths.split_output_folder,
            paths.desktop_scan_zip_output,
        )
        return result

    logger.run_step(
        "scan.py",
        scan_action,
    )
    return "success"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__, allow_abbrev=False)
    teacher_group = parser.add_mutually_exclusive_group()
    teacher_group.add_argument(
        "--teacher",
        help=f"人员文件夹名称：{', '.join(TEACHERS)}",
    )
    for teacher in TEACHERS:
        teacher_group.add_argument(
            f"--{teacher}",
            dest="teacher",
            action="store_const",
            const=teacher,
            help=f"选择 {teacher} 老师",
        )

    flow_group = parser.add_mutually_exclusive_group()
    flow_group.add_argument(
        "--flow",
        choices=("A", "B"),
        help="A=盖章前流程，B=盖章后流程",
    )
    flow_group.add_argument(
        "--A",
        dest="flow",
        action="store_const",
        const="A",
        help="运行 A 功能：precheck → registration → merge",
    )
    flow_group.add_argument(
        "--B",
        dest="flow",
        action="store_const",
        const="B",
        help="运行 B 功能：split → scan",
    )
    parser.add_argument(
        "--on-issue",
        choices=("ask", "stop", "continue"),
        default="ask",
        help="precheck 发现异常时的处理方式（默认：交互询问）",
    )
    parser.add_argument(
        "--on-split-issue",
        choices=("ask", "stop", "continue"),
        default="ask",
        help="split 异常时的处理方式（默认：交互询问，可选择 resplit）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只预览生成和合并计划，不写入登记表或合并 PDF",
    )
    parser.add_argument(
        "--task-date",
        type=lambda value: datetime.strptime(value, "%Y%m%d").date(),
        metavar="YYYYMMDD",
        help="指定任务日期；A 默认今天，B 默认最近完成的 A 批次日期",
    )
    parser.add_argument(
        "--batch",
        type=int,
        help="手动指定批次；默认自动选择",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    logger: RunLogger | None = None
    status = "failed"
    try:
        teacher = choose_teacher(args.teacher)
        flow = choose_flow(args.flow)
        if args.batch is not None and args.batch < 1:
            raise RuntimeError("--batch 必须是大于 0 的整数。")

        if flow == "A":
            task_date = args.task_date or date.today()
            batch = args.batch or choose_a_batch(teacher, task_date)
        else:
            if args.task_date is not None or args.batch is not None:
                task_date = args.task_date or date.today()
                available = completed_batches(teacher, task_date)
                if args.batch is not None:
                    batch = args.batch
                    if batch not in available:
                        raise RuntimeError(
                            f"{teacher} 在 {task_date:%Y-%m-%d} 的 {batch:02d} 批次"
                            "尚未完成 A 功能。"
                        )
                else:
                    if not available:
                        raise RuntimeError(
                            f"{teacher} 在 {task_date:%Y-%m-%d} 没有已完成的 A 功能批次。"
                        )
                    batch = max(available)
            else:
                latest = latest_completed_batch(teacher)
                if latest is None:
                    raise RuntimeError(
                        f"没有找到 {teacher} 已完成的 A 功能批次，无法确定 B 功能批次。"
                    )
                task_date, batch = latest

        scan_date = date.today()
        paths = build_task_paths(
            teacher,
            task_date,
            batch,
            scan_date=scan_date,
        )
        logger = RunLogger(paths, flow)

        print(f"\n本次处理人员：{teacher}")
        print(f"任务日期：{task_date:%Y-%m-%d}")
        print(f"任务批次：{batch:02d}")
        print(f"确认书目录：{paths.confirmation_folder}")
        if flow == "A":
            status = run_before_flow(paths, logger, args.on_issue, args.dry_run)
        else:
            print(f"当天扫描件：{paths.scan_source_pdf}")
            print(f"拆分目录：{paths.split_output_folder}")
            status = run_after_flow(
                paths,
                logger,
                args.dry_run,
                args.on_split_issue,
            )
        return 0
    except KeyboardInterrupt:
        status = "interrupted"
        print("\n操作已由用户中断。", file=sys.stderr)
        return 130
    except (EOFError, OSError, RuntimeError, ValueError) as exc:
        status = "failed"
        print(f"\n错误：{exc}", file=sys.stderr)
        return 1
    finally:
        if logger is not None:
            try:
                logger.save(status)
            except OSError as exc:
                print(f"警告：无法写入运行记录：{exc}", file=sys.stderr)


if __name__ == "__main__":
    raise SystemExit(main())
