# 必赢 · 标讯雷达 —— 首次安装（给 Agent 执行）

目标：让本机能运行 `scripts/radar.py`，只需要 **Python 3.8 或更高版本**，不需要装任何第三方包。

全程原则：
- **能自己执行的就自己执行**，只在系统弹窗、需要用户点「安装 / 允许」时请用户操作，并用一句话告诉用户点哪里
- 每一步执行完都要看输出确认成功，**不得在没确认的情况下告诉用户「装好了」**
- 不修改系统设置，不关闭杀毒软件 / 安全软件

---

## 第 1 步：判断系统

- Windows → 看「A. Windows」
- macOS → 看「B. macOS」
- Linux → 执行 `python3 --version`，≥ 3.8 即可跳到「C. 验证」；没有就用系统包管理器安装 python3（如 `sudo apt install python3`，密码由用户输入）

---

## A. Windows

### A1. 检查是否已有 Python

依次执行，**哪条输出了 `Python 3.x`（x ≥ 8）就用哪条**，作为 `<PY>`：

```
python --version
py -3 --version
```

> ⚠️ 如果 `python` 打开了微软应用商店、或没有任何输出，说明是系统占位程序，不算装了 Python。

有 → 跳到「C. 验证」。没有 → A2。

### A2. 安装 Python

**方式一（优先）：winget**

```
winget install -e --id Python.Python.3.12 --source winget --accept-source-agreements --accept-package-agreements
```

**方式二（winget 不可用或下载失败时）：国内镜像安装包**

在 PowerShell 执行：

```
$f = "$env:TEMP\python-3.12.7-amd64.exe"
Invoke-WebRequest -Uri "https://registry.npmmirror.com/-/binary/python/3.12.7/python-3.12.7-amd64.exe" -OutFile $f
Start-Process -FilePath $f -ArgumentList "/quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1" -Wait
```

如果弹出「是否允许此应用对你的设备进行更改」，请用户点「是」。

### A3. 确认安装

刚装完，当前终端可能还找不到新 Python（环境变量没刷新）。依次尝试，哪条成功就用哪条作为 `<PY>`：

```
py -3 --version
python --version
& "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe" --version
```

都失败 → 请用户**重启当前使用的 AI 助手**后再试一次；仍失败 → 告诉用户「Python 安装没成功」，请用户到 https://www.python.org/downloads/ 手动下载安装，并提醒**勾选 Add python.exe to PATH**。

---

## B. macOS

### B1. 检查是否已有 Python

```
python3 --version
```

输出 `Python 3.x`（x ≥ 8）→ 跳到「C. 验证」。

> ⚠️ 如果执行后系统弹窗提示「需要安装命令行开发者工具」，这就是在装 Python，进入 B2。

### B2. 安装 Python

**方式一（优先）：系统自带的命令行开发者工具**

```
xcode-select --install
```

会弹出系统窗口 → 请用户点「安装」并同意协议，等它装完（通常几分钟）。装完再执行 `python3 --version` 确认。

**方式二（方式一失败时）：国内镜像安装包**

```
curl -L -o ~/Downloads/python-3.12.7-macos11.pkg "https://registry.npmmirror.com/-/binary/python/3.12.7/python-3.12.7-macos11.pkg"
open ~/Downloads/python-3.12.7-macos11.pkg
```

会打开安装向导 → 请用户一路点「继续」→「安装」（需要输入电脑开机密码，**由用户自己输入，你不要代填**）。装完执行 `python3 --version` 确认。

---

## C. 验证

```
<PY> <RADAR> check
```

看到 `"python_ok": true` 即 Python 就绪。

- `key_configured` 为 false → 回到 SKILL.md「写入 key」，请用户把 key 发给你
- `key_configured` 为 true → 已完成，告诉用户可以直接开始查标

安装完成后回复用户（示例）：

> 标讯雷达已装好 ✅ 请把你的标讯雷达 key 发给我，我帮你配置好就能开始用。
