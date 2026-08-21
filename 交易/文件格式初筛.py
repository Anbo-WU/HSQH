#!/usr/bin/env python3
"""递归检查“确认书文件”中的重名或疑似重复 PDF。"""

from __future__ import annotations

import argparse
import hashlib
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path


脚本目录 = Path(__file__).resolve().parent
默认检查目录 = 脚本目录 / "确认书文件"

# 浏览器或系统复制文件时常见的结尾，例如“文件 (1).pdf”“文件_副本.pdf”。
复制件后缀 = re.compile(
    r"(?:\s*[（(]\d+[）)]|[\s_-]*(?:副本|复制件|copy)(?:\s*\d+)?)$",
    flags=re.IGNORECASE,
)

# 标准示例：公司名称_商品交易确认书_【HFSY】0183-JY-2026080701.pdf
标准文件名 = re.compile(
    r"^.+_商品交易确认书_【HFSY】\d{4}-(?:JY|FWJY)-\d{10}\.pdf$"
)


def 统一名称(文件名: str) -> str:
    """忽略英文大小写及全角、半角字符差异。"""
    return unicodedata.normalize("NFKC", 文件名).casefold()


def 去掉复制件后缀(文件名: str) -> str:
    路径 = Path(文件名)
    主体 = 复制件后缀.sub("", 路径.stem).rstrip()
    return 统一名称(主体 + 路径.suffix)


def 计算哈希(文件路径: Path) -> str:
    摘要 = hashlib.sha256()
    with 文件路径.open("rb") as 文件:
        while 数据 := 文件.read(1024 * 1024):
            摘要.update(数据)
    return 摘要.hexdigest()


def 添加分组(分组: dict[str, list[Path]], 标题: str, 输出: list[str]) -> int:
    重复组 = [路径列表 for 路径列表 in 分组.values() if len(路径列表) > 1]
    重复组.sort(key=lambda 路径列表: 统一名称(路径列表[0].name))

    输出.append(f"\n{标题}：{len(重复组)} 组")
    输出.append("=" * 72)
    if not 重复组:
        输出.append("未发现。")
        return 0

    for 序号, 路径列表 in enumerate(重复组, start=1):
        输出.append(f"[{序号}] {路径列表[0].name}")
        for 文件路径 in sorted(路径列表, key=lambda 路径: str(路径).casefold()):
            输出.append(f"    {文件路径.resolve()}")
        输出.append("")
    return len(重复组)


def 添加异常文件(文件列表: list[Path], 输出: list[str]) -> int:
    输出.append(f"\n四、命名格式异常：{len(文件列表)} 个")
    输出.append("=" * 72)
    输出.append("标准格式：公司名称_商品交易确认书_【HFSY】四位编号-JY或FWJY-八位日期两位流水号.pdf")
    if not 文件列表:
        输出.append("未发现。")
        return 0

    for 序号, 文件路径 in enumerate(文件列表, start=1):
        原因 = "带有常见复制件后缀" if 复制件后缀.search(文件路径.stem) else "不符合标准命名结构"
        输出.append(f"[{序号}] {原因}")
        输出.append(f"    {文件路径.resolve()}")
    return len(文件列表)


def 检查PDF(检查目录: Path) -> tuple[str, int]:
    PDF列表 = sorted(
        (路径 for 路径 in 检查目录.rglob("*") if 路径.is_file() and 路径.suffix.casefold() == ".pdf"),
        key=lambda 路径: str(路径).casefold(),
    )

    同名分组: dict[str, list[Path]] = defaultdict(list)
    疑似复制件分组: dict[str, list[Path]] = defaultdict(list)
    同内容分组: dict[str, list[Path]] = defaultdict(list)
    命名异常列表: list[Path] = []

    for 文件路径 in PDF列表:
        同名分组[统一名称(文件路径.name)].append(文件路径)
        疑似复制件分组[去掉复制件后缀(文件路径.name)].append(文件路径)
        if 标准文件名.fullmatch(文件路径.name) is None:
            命名异常列表.append(文件路径)

        try:
            同内容分组[计算哈希(文件路径)].append(文件路径)
        except OSError as 错误:
            print(f"警告：无法读取 {文件路径.resolve()}：{错误}", file=sys.stderr)

    # “疑似复制件”部分只保留至少有一个复制后缀，且并非已经完全同名的组。
    疑似复制件分组 = {
        名称: 路径列表
        for 名称, 路径列表 in 疑似复制件分组.items()
        if len(路径列表) > 1
        and len({统一名称(路径.name) for 路径 in 路径列表}) > 1
        and any(复制件后缀.search(路径.stem) for 路径 in 路径列表)
    }

    输出 = [
        "PDF 重复文件初筛报告",
        f"检查目录：{检查目录.resolve()}",
        f"PDF 总数：{len(PDF列表)}",
    ]
    发现组数 = 0
    发现组数 += 添加分组(同名分组, "一、文件名完全相同", 输出)
    发现组数 += 添加分组(疑似复制件分组, "二、带 (1)、副本、copy 等后缀的疑似复制件", 输出)
    发现组数 += 添加分组(同内容分组, "三、文件内容完全相同（SHA-256 一致）", 输出)
    发现组数 += 添加异常文件(命名异常列表, 输出)

    输出.append("说明：本脚本只检查并报告，不会移动、删除或修改任何文件。")
    return "\n".join(输出), 发现组数


def main() -> int:
    解析器 = argparse.ArgumentParser(
        description="递归检查确认书目录中的重名、疑似复制及内容相同的 PDF。"
    )
    解析器.add_argument(
        "目录",
        nargs="?",
        type=Path,
        default=默认检查目录,
        help=f"要检查的目录（默认：{默认检查目录}）",
    )
    参数 = 解析器.parse_args()
    检查目录 = 参数.目录.expanduser()

    if not 检查目录.is_dir():
        print(f"错误：找不到目录：{检查目录.resolve()}", file=sys.stderr)
        return 2

    报告, _ = 检查PDF(检查目录)
    print(报告)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
