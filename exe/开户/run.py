#!/usr/bin/env python3
"""account 的终端菜单；使用同一解释器启动原有业务脚本。"""
from __future__ import annotations

import importlib
from importlib.metadata import version
from pathlib import Path
import subprocess
import sys

BASE = Path(__file__).resolve().parent
SCRIPTS = {
    'register': BASE/'客户信息登记'/'登记.py',
    'extract': BASE/'问卷分数等级'/'分数提取.py',
    'grade': BASE/'问卷分数等级'/'等级计算.py',
}
HELP = """account：输入编号或使用下面的命令。
  1  account register   生成客户信息登记表
  2  account extract    提取问卷答案，更新分数统计.xlsx 的 A:U
  3  account grade      根据现有答案更新总分和等级（V:W）
  4  account score      先提取问卷，再计算总分和等级
  5  account check      检查依赖和必要文件，不处理客户材料
  6  account models     准备并检查登记功能的 OCR 模型（首次需联网）
  0  退出

Windows 问卷请使用 .docx；当前问卷 PDF 读取依赖 macOS。
登记功能支持 PDF。运行前请关闭正在使用的结果 Excel。
查看单项参数：account register --help / account extract --help / account grade --help
"""


def check() -> int:
    print(f'程序目录：{BASE}\nPython：{sys.executable}\n版本：{sys.version.split()[0]}', flush=True)
    errors = []
    if sys.version_info[:2] != (3, 12):
        errors.append('本安装说明使用 Python 3.12，请检查虚拟环境版本')
    for package, module in [('openpyxl', 'openpyxl'), ('PyMuPDF', 'pymupdf'),
                            ('rapidocr', 'rapidocr'), ('onnxruntime', 'onnxruntime'), ('numpy', 'numpy')]:
        try:
            importlib.import_module(module)
            print(f'依赖正常：{package} {version(package)}', flush=True)
        except Exception as exc:
            errors.append(f'{package} 无法加载：{exc}')
    for path in [*SCRIPTS.values(), BASE/'客户信息登记'/'ocr_support.py',
                 BASE/'客户信息登记'/'场外衍生品中心客户信息登记表-模板.xlsx',
                 BASE/'问卷分数等级'/'分数统计.xlsx']:
        if not path.is_file():
            errors.append(f'缺少文件：{path}')
    if errors:
        print('\n'.join(errors), file=sys.stderr)
        return 1
    print('依赖和必要文件检查通过。OCR 模型另用 account models 检查。')
    return 0


def run_script(command: str, args: list[str]) -> int:
    path = SCRIPTS[command]
    if not path.is_file():
        print(f'缺少程序：{path}', file=sys.stderr)
        return 1
    # 不切换工作目录；原脚本用自身位置查找默认材料，外部参数按调用者目录解析。
    return subprocess.run([sys.executable, str(path), *args], check=False).returncode


def main() -> int:
    args = sys.argv[1:]
    if not args:
        print(HELP, flush=True)
        try:
            choice = input('请输入编号并按回车：').strip()
        except EOFError:
            return 0
        if choice == '0':
            return 0
        args = [{'1': 'register', '2': 'extract', '3': 'grade', '4': 'score',
                 '5': 'check', '6': 'models'}.get(choice, choice)]
    command, *extra = args
    if command in ('--help', '-h', 'help'):
        print(HELP)
        return 0
    if command in SCRIPTS:
        return run_script(command, extra)
    if extra:
        print(f'{command} 不接受附加参数，请使用单项命令传参。', file=sys.stderr)
        return 2
    if command == 'check':
        return check()
    if command == 'models':
        from rapidocr import RapidOCR
        RapidOCR()
        print('OCR 模型已准备好，登记功能可以使用本地识别。')
        return 0
    if command == 'score':
        status = run_script('extract', [])
        if status:
            print('提取失败，已停止，未启动等级计算。', file=sys.stderr)
            return status
        return run_script('grade', [])
    print(HELP)
    print(f'未识别的命令或编号：{command}', file=sys.stderr)
    return 2


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print('\n已中止。', file=sys.stderr)
        raise SystemExit(130)
    except Exception as exc:
        print(f'启动失败：{exc}\n请参照安装说明，或把完整报错发给维护同事。', file=sys.stderr)
        raise SystemExit(1)
