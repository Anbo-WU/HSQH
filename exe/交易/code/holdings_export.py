"""由 LibreOffice 自带 Python 运行，按实际列宽导出临时打印副本。"""
from __future__ import annotations

import json
import math
from pathlib import Path
import subprocess
import sys
import time
import uuid

import uno  # LibreOffice 的 Python UNO 接口，不在项目 venv 内导入。


def prop(name, value):
    result = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    result.Name, result.Value = name, value
    return result


def export(jobs_path: Path, soffice: Path, profile: Path) -> None:
    jobs = json.loads(jobs_path.read_text(encoding="utf-8"))
    pipe = "holdings_" + uuid.uuid4().hex
    process = subprocess.Popen(
        [str(soffice), f"-env:UserInstallation={profile.resolve().as_uri()}",
         "--headless", "--nologo", "--nodefault", "--nofirststartwizard",
         f"--accept=pipe,name={pipe};urp;StarOffice.ServiceManager"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    desktop = None
    results = []
    try:
        context = uno.getComponentContext()
        resolver = context.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", context)
        deadline = time.monotonic() + 45
        while True:
            try:
                remote = resolver.resolve(f"uno:pipe,name={pipe};urp;StarOffice.ComponentContext")
                break
            except Exception:
                if time.monotonic() > deadline or process.poll() is not None:
                    raise RuntimeError("无法连接本地 LibreOffice 排版进程。")
                time.sleep(0.2)
        desktop = remote.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", remote)
        for job in jobs:
            document = desktop.loadComponentFromURL(
                Path(job["source"]).resolve().as_uri(), "_blank", 0,
                (prop("Hidden", True), prop("ReadOnly", True), prop("UpdateDocMode", 0),
                 prop("MacroExecutionMode", uno.getConstantByName("com.sun.star.document.MacroExecMode.NEVER_EXECUTE"))),
            )
            if document is None:
                raise RuntimeError(f"无法打开打印副本：{job['source']}")
            try:
                sheets = document.getSheets()
                sheet = sheets.getByName(job["sheet"])
                for name in sheets.getElementNames():
                    if name != job["sheet"]:
                        sheets.getByName(name).setPrintAreas(())
                area = sheet.getCellRangeByName(job["area"])
                sheet.setPrintAreas((area.getRangeAddress(),))
                style = document.StyleFamilies.getByName("PageStyles").getByName(sheet.PageStyle)
                available = style.Width - style.LeftMargin - style.RightMargin
                columns = area.getColumns()
                width = sum(columns.getByIndex(i).Width for i in range(columns.getCount())
                            if columns.getByIndex(i).IsVisible)
                scale = min(job["max_scale"], 100, math.floor((available - 100) / max(width, 1) * 100))
                if scale < 10:
                    raise RuntimeError(f"{job['sheet']} 太宽，缩至 10% 仍无法打印于一页宽。")
                style.ScaleToPages = 0
                style.ScaleToPagesX = 0
                style.ScaleToPagesY = 0
                style.PageScale = scale
                document.storeToURL(Path(job["target"]).resolve().as_uri(), (
                    prop("FilterName", "calc_pdf_Export"), prop("Overwrite", True),
                    prop("FilterData", (prop("ExportHiddenSheets", False), prop("SinglePageSheets", False))),
                ))
                results.append({"source": job["source"], "scale": scale, "width_mm": width / 100,
                                "available_mm": available / 100})
            finally:
                document.close(True)
        jobs_path.with_suffix(".result.json").write_text(json.dumps(results, indent=2), encoding="utf-8")
    finally:
        if desktop is not None:
            desktop.terminate()
        try:
            process.wait(timeout=15)
        except subprocess.TimeoutExpired:
            process.terminate()
            process.wait(timeout=10)


if __name__ == "__main__":
    export(Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3]))
