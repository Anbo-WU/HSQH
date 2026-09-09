# account 项目安装与运行说明（Windows 新手版）

编写日期：2026-09-09。适用于 Windows 10 / 11、Intel 或 AMD 的 64 位电脑。本文只安装 `account` 所需环境，不需要复制整个 `exe` 项目。

完成一次安装后，平时只需：**按 Windows 键，搜索并打开 PowerShell，输入 `account`，按提示选择功能。** 不用打开项目文件夹，不用启动编程软件。

本文的主启动器是 Windows 自带终端可执行的 `account.bat`。另附 `account.sh`，供已经使用 Git Bash 的人选择；初次安装按 PowerShell 路线即可。

## 1. 先了解各项功能

| 功能 | 平时输入的命令 | 默认材料位置 | 结果位置 |
|---|---|---|---|
| 生成客户信息登记表 | `account register` | `account\客户信息登记\各客户文件夹\` | `客户信息登记\结果\` 中带时间戳的新文件 |
| 提取问卷答案 | `account extract` | `account\问卷分数等级\问卷\` | 更新 `问卷分数等级\分数统计.xlsx` 的 A:U 列 |
| 根据答案计算总分和等级 | `account grade` | `问卷分数等级\分数统计.xlsx` | 更新该表的 V:W 列 |
| 连续完成问卷提取和评分 | `account score` | 同上 | 同上；提取失败时不会继续评分 |

**现有代码的格式要求：**

- 客户信息登记：Windows 支持 PDF，包括扫描 PDF，识别使用本地 RapidOCR。
- 问卷分数等级：Windows 请使用 `.docx`。现有“分数提取.py”的 PDF 分支使用 macOS 组件，**扫描 PDF 和带文字层的 PDF 在 Windows 上都不能直接通过这个问卷程序处理**。这不是少装了 Python 库。请取得原始 `.docx`；只有 PDF 时，先用办公软件转换并人工核对题目和答案。
- 旧版 `.doc` 请用 Word/WPS “另存为” `.docx`，不能只修改扩展名。
- 问卷答案必须符合现有程序规则：20 道题，答案字母写在题目段落末尾；不能仅靠给选项打勾。第 13、17 题支持多选。

## 2. 下载安装 Python

### 2.1 需要哪些软件

| 软件 | 是否需要 | 用途 |
|---|---|---|
| Python 3.12（64 位） | 必须 | 运行程序 |
| PowerShell | Windows 已自带 | 输入命令 |
| Excel 或 WPS 表格 | 查看和填写结果时需要 | 沿用电脑已有的办公软件即可 |
| Word 或 WPS 文字 | 转换旧问卷格式时使用 | `.doc` 另存为 `.docx`，检查转换后的内容 |

无需安装 VS Code、PyCharm、Anaconda、Git Bash、Tesseract 或 CUDA 来完成本文的 Windows 流程。

### 2.2 安装 Python install manager

1. 用浏览器打开 [Python 官方下载页](https://www.python.org/downloads/)。
2. 选择 **Python install manager** 的正式版本，下载 Windows 安装包。
3. 双击下载的安装包，点击“安装 / Install”，等待完成。
4. 按 Windows 键，搜索 `PowerShell`，打开 **Windows PowerShell**。后面的代码块均在这个窗口粘贴、按回车执行，不要粘贴到 Python 的 `>>>` 窗口。

安装管理器及 `pymanager` 命令用法可参见 [Python 官方 Windows 安装说明](https://docs.python.org/3/using/windows.html)。

逐条执行下面两行，每条执行结束后再执行下一条：

```powershell
pymanager install 3.12
pymanager exec -V:3.12 --version
```

第二条应显示 `Python 3.12.x`，末尾的小版本数字可以不同。本文固定使用 3.12 系列，不要自行改成 3.10 或实验版本。

检查是 64 位：

```powershell
pymanager exec -V:3.12 -c "import struct; print(struct.calcsize('P') * 8)"
```

显示 `64` 才继续。若电脑是 ARM 架构，本教程未验证，请先交给维护同事确认环境。

如果提示找不到 `pymanager`，先关闭全部终端窗口后重新打开 PowerShell，再试一次；仍找不到时，检查 Python install manager 是否已安装、Windows“管理应用执行别名”中是否启用对应命令。单位限制软件安装时，把安装提示交给 IT 处理。

## 3. 复制完整 account 文件夹

本教程假设同事将文件夹放在：

```text
C:\Work\account
```

也可以换成其他位置。下面所有安装步骤中的 `$AccountDir` 都应指向**她自己电脑里的 account 文件夹**，不要使用原作者电脑的用户名或桌面路径。

复制后至少应包含：

```text
account\
├─ account.bat                         Windows 终端启动器（本次已提供）
├─ account.sh                          可选的 Git Bash 启动器（本次已提供）
├─ run.py                              中文菜单和命令入口（本次已提供）
├─ requirements.txt                    account 专用依赖清单（本次已提供）
├─ account安装与运行说明.md
├─ 客户信息登记\
│  ├─ 登记.py
│  ├─ ocr_support.py                    必须和登记.py 一起复制
│  ├─ 场外衍生品中心客户信息登记表-模板.xlsx
│  ├─ 某客户完整文件夹\
│  │  ├─ 0.OA\                        KYC、资信评估等原始材料
│  │  └─ 1.开户材料存档\               带原有编号的 PDF 材料
│  └─ 结果\                           运行后自动创建
└─ 问卷分数等级\
   ├─ 分数提取.py
   ├─ 等级计算.py
   ├─ 分数统计.xlsx                    必须保留现有模板表头
   └─ 问卷\                           本次要处理的 .docx 问卷
```

客户材料的文件名、编号和层级按原来的架构复制；不要只复制 `.py` 文件。问卷程序只读取“问卷”目录直接包含的文件，不会递归读取子文件夹。

**不要复制别人的 `.venv` 虚拟环境。** 在同事电脑上按下一步新建。`__pycache__` 和 `.ocr-cache` 是缓存，不是必需的程序文件；复制源码和必要材料即可。Python 官方说明也要求在新位置重新创建虚拟环境。[参考：venv 文档](https://docs.python.org/3.12/library/venv.html)

## 4. 创建 account 自己的 Python 环境

这一节到第 7 节请在**同一个 PowerShell 窗口**完成。修改下面第一行的路径，然后整块粘贴：

```powershell
$AccountDir = 'C:\Work\account'
$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath (Join-Path $AccountDir 'requirements.txt'))) {
    throw '没有找到 account 的 requirements.txt，请检查路径和复制的文件。'
}
$AccountDir = (Resolve-Path -LiteralPath $AccountDir).Path
$AccountEnv = Join-Path $AccountDir '.venv'
pymanager exec -V:3.12 -m venv "$AccountEnv"
if ($LASTEXITCODE -ne 0) { throw '创建环境失败，请先处理上面的报错。' }
$AccountPython = Join-Path $AccountEnv 'Scripts\python.exe'
& $AccountPython --version
```

执行成功后会生成 `account\.venv`，最后显示 `Python 3.12.x`。

我们始终直接使用这个环境里的 Python，不需要运行 `activate`，也不需要修改 PowerShell 执行策略。[参考：Python venv 用法](https://docs.python.org/3.12/library/venv.html)

## 5. 安装依赖和库

在刚才的 PowerShell 窗口粘贴：

```powershell
& $AccountPython -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw 'pip 更新失败，请先处理上面的报错。' }
& $AccountPython -m pip install -r (Join-Path $AccountDir 'requirements.txt')
if ($LASTEXITCODE -ne 0) { throw '依赖安装失败，请先处理上面的报错。' }
& $AccountPython -m pip check
if ($LASTEXITCODE -ne 0) { throw '依赖有冲突，请把完整报错发给维护同事。' }
```

首次需要联网，等待出现 `Successfully installed` 或 `Requirement already satisfied`；最后应显示 `No broken requirements found.`。安装过程中不要关闭窗口。

`account\requirements.txt` 已单独整理为：

```text
openpyxl==3.1.5
PyMuPDF==1.28.2
rapidocr==3.9.2
onnxruntime==1.29.0
numpy==2.5.2
```

这些版本取自原电脑当前可用的 Python 3.12 环境。openpyxl 读写 Excel；PyMuPDF 读取、渲染 PDF；RapidOCR、ONNX Runtime、NumPy 用于本地文字识别。间接依赖由 pip 自动安装，无需逐个查找。

这是 **account 自己的清单**，不需要安装根项目中给 trader 等其他项目使用的 xlrd、xlwt、xlutils，也不需要额外安装 python-docx；当前问卷代码使用 Python 自带组件读取 `.docx`。

## 6. 检查依赖，准备 OCR 模型

仍在同一窗口执行：

```powershell
& $AccountPython (Join-Path $AccountDir 'run.py') check
if ($LASTEXITCODE -ne 0) { throw '检查未通过，请根据上面的提示补齐环境或文件。' }
& $AccountPython (Join-Path $AccountDir 'run.py') models
if ($LASTEXITCODE -ne 0) { throw 'OCR 模型尚未准备好，请检查上面的下载或加载错误。' }
```

依次看到“依赖和必要文件检查通过”和“OCR 模型已准备好”即可。检查不会生成客户登记表，也不会改写问卷统计表。

`pip install` 安装库，`models` 初始化识别模型，两步都要完成。模型尚未缓存时需要联网下载；准备好之后，登记程序使用本地模型处理材料。缓存与模型是两类文件：业务 `.ocr-cache` 不是模型，单独复制它不能代替准备模型。[参考：RapidOCR 官方安装说明](https://rapidai.github.io/RapidOCRDocs/main/install_usage/rapidocr/install/)

## 7. 配置一次，以后在任意目录输入 account

下面这段把 **account 文件夹路径**加入当前用户的 PATH。PATH 的作用是让终端找到 `account.bat`；代码会保留原有 PATH 项，也不会重复添加同一路径。

```powershell
$AccountDir = (Resolve-Path -LiteralPath $AccountDir).Path
if (-not (Test-Path -LiteralPath (Join-Path $AccountDir 'account.bat'))) {
    throw 'account.bat 不存在，请补齐启动器。'
}
$AccountUserPath = [Environment]::GetEnvironmentVariable('Path', 'User')
$AccountPathParts = @($AccountUserPath -split ';' | Where-Object { $_.Trim() })
if ($AccountPathParts -notcontains $AccountDir) {
    $AccountNewPath = (@($AccountPathParts) + $AccountDir) -join ';'
    [Environment]::SetEnvironmentVariable('Path', $AccountNewPath, 'User')
}
if (@($env:Path -split ';') -notcontains $AccountDir) {
    $env:Path = $env:Path.TrimEnd(';') + ';' + $AccountDir
}
Get-Command account.bat
```

最后输出的 `Source` 应指向她电脑上的 `account\account.bat`。现在就可以执行：

```powershell
account check
account
```

`account` 会显示中文菜单，输入编号再回车即可。

安装完成后可以关闭窗口。下次从 Windows 开始菜单重新打开 PowerShell，直接输入 `account`。如果在旧的 Windows Terminal 里新建标签页仍找不到命令，先关闭整个 Terminal 再重开；仍不生效时注销 Windows 后重新登录。

如果同名命令冲突，可使用 `account.bat`；也可以用下面的完整路径启动，**仍不需要进入文件夹**：

```powershell
& 'C:\Work\account\account.bat'
```

## 8. 日常操作示例

### 客户信息登记

放好本次各客户的 `0.OA` 和 `1.开户材料存档` 后，输入：

```powershell
account register
```

它会处理 `客户信息登记` 目录下所有符合结构的客户文件夹。只处理某一家时，填写该文件夹的完整名称：

```powershell
account register --customer '398、成都向鸿商贸有限公司20260904-C4低买高'
```

运行后到 `客户信息登记\结果` 查看新生成的 Excel、`.核对.json` 和 `.待核对.txt`。终端会显示完整路径。登记程序使用时间戳生成新结果，不覆盖已有文件；无法确认的字段留空并标黄，请依据原材料人工填写。

### 问卷提取和评分

先将本次要处理的 `.docx` 放入 `问卷分数等级\问卷`，保留 `分数统计.xlsx` 的模板表头，关闭 Excel/WPS 中打开的统计表。

检查问卷，不写入：

```powershell
account extract --preview
```

确认材料符合格式后，完成提取和评分：

```powershell
account score
```

结果是原位置的 `分数统计.xlsx`。**提取会替换该表 A:U 的旧记录，评分会更新 V:W**，不会另外生成一个带日期的新统计表。如果要保留上一批结果，处理新一批前先另存一份；“问卷”里只放本次要处理的文件。

如果答案已经人工填好，只要计算等级：

```powershell
account grade
```

也可以只预览分数：

```powershell
account grade --preview
```

## 9. 常见问题

| 提示或现象 | 怎么处理 |
|---|---|
| `account` 不是可识别的命令 | 关闭整个终端重开；检查第 7 节 PATH；先用完整路径执行 `account.bat`。 |
| `Python environment not found` | 启动器找不到 `account\.venv\Scripts\python.exe`；按第 4～6 节创建本机环境。 |
| `$AccountDir` 或 `$AccountPython` 为空 | 可能换了 PowerShell 窗口，重新执行下面的“恢复路径变量”代码。 |
| `No module named ...` | 重新执行第 5 节，确认用的是 `$AccountPython -m pip`，不要使用其他 Python 的裸 `pip` 命令。 |
| `No matching distribution found` | 检查 Python 为 3.12、64 位，并确认版本号未输错。若单位软件源缺少指定版本，请把错误交给维护同事；不要随意删版本号。 |
| 下载超时、连接失败 | 检查网络；若单位要求代理或指定软件源，请向 IT 获取配置。pip 的库下载失败和 OCR 模型下载失败需分别处理。 |
| `PermissionError` / 文件正在使用 / 保存失败 | 关闭 Excel/WPS 中打开的结果表，再运行；材料放在当前用户可写入的本地文件夹。 |
| 问卷提示需要 macOS | 当前问卷 PDF 功能不支持 Windows，改用 `.docx`，不要继续安装 OCR 库尝试解决。 |
| 缺少模板、缺少客户文件夹、表头不一致 | 按第 3 节补齐完整目录与原模板；不要自行改模板表头或材料编号。 |
| 终端出现 `>>>` | 这是 Python 交互窗口；输入 `exit()` 回车退出，再回到 PowerShell 执行命令。 |
| OCR 首次运行较慢 | 等待日志继续输出；后续同一材料会使用识别缓存。 |

恢复路径变量（把路径换成实际位置）：

```powershell
$AccountDir = 'C:\Work\account'
$AccountPython = Join-Path $AccountDir '.venv\Scripts\python.exe'
```

如果提示 `DLL load failed` 且涉及 `onnxruntime`，先确认 Python 是 64 位。ONNX Runtime 的 Windows 构建需要 Visual C++ 运行库，可按 [ONNX Runtime 官方安装要求](https://onnxruntime.ai/docs/install/) 安装 [Microsoft 官方 Visual C++ Redistributable x64](https://aka.ms/vs/17/release/vc_redist.x64.exe)，然后重开 PowerShell 再执行 `account check`。

移动整个 account 文件夹后，原虚拟环境可能失效，PATH 也仍指向旧位置。应在新位置重新创建 `.venv`、安装依赖并更新 PATH；不要只移动文件夹后继续使用旧环境。

## 10. 已提供的启动脚本

以下文件已放进 `account`，同事复制完整文件夹后无需自己编写。启动器通过自身位置寻找 `.venv` 和 `run.py`，没有写死原作者电脑的路径。**这几个文件需一起复制，不能只拿下面一个 bat 文件。**

### Windows 主启动器：account.bat

```bat
@echo off
setlocal
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
chcp 65001 >nul
if not exist "%~dp0.venv\Scripts\python.exe" (
    echo Python environment not found: "%~dp0.venv\Scripts\python.exe"
    echo Please create the account .venv and install requirements.txt first.
    exit /b 1
)
"%~dp0.venv\Scripts\python.exe" "%~dp0run.py" %*
exit /b %errorlevel%
```

### 可选 Bash 脚本：account.sh

这个脚本用于 **Windows 的 Git Bash**，同样调用已经配置好的 Windows `.venv`，不是 macOS/Linux 安装教程。没有安装 Git Bash 的同事直接使用前面的 PowerShell 流程即可。

```bash
#!/usr/bin/env bash
set -eu
account_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
account_python="$account_dir/.venv/Scripts/python.exe"
if [[ ! -f "$account_python" ]]; then
    printf 'Python environment not found: %s\n' "$account_python" >&2
    exit 1
fi
export PYTHONUTF8=1
export PYTHONIOENCODING=utf-8
exec "$account_python" "$account_dir/run.py" "$@"
```

已使用 Git Bash 时，在任意目录执行以下命令即可，不需要先进入 account 文件夹：

```bash
bash /c/Work/account/account.sh
bash /c/Work/account/account.sh register
```

## 11. 安装完成的判断标准

1. `account check` 显示依赖和必要文件检查通过。
2. `account models` 显示 OCR 模型已准备好。
3. 从任意目录输入 `account` 能显示中文菜单。
4. 放入真实材料后，按第 8 节执行对应功能，在终端提示的位置找到结果。

原电脑验证环境为 Python 3.12.14 / Windows x64；新电脑仍需完成上述检查，尤其是模型下载和办公文件是否齐全。
