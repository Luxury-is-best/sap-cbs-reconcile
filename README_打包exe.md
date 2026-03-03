# Windows 打包 exe 说明

## 目标
- 将桌面工具 `gui_app.py` 打包为 Windows 可直接运行的 `exe`。
- 同事无需安装 Python。

## 1. 打包环境
- Windows 10/11
- Python 3.10+
- 项目目录：`SAP和CBS对账整月`

## 2. 安装依赖
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 3. 执行打包（推荐 onedir）
```powershell
pyinstaller --noconfirm --clean --windowed --name 对账助手 --onedir gui_app.py
```

产物路径：
- `dist\对账助手\对账助手.exe`

## 4. 一键脚本打包
可直接运行：
```powershell
.\build_exe.bat
```

## 5. 分发建议
- 将 `dist\对账助手` 整个文件夹打包给同事。
- 同事双击 `对账助手.exe` 即可使用。

## 6. 可选：单文件 onefile
```powershell
pyinstaller --noconfirm --clean --windowed --name 对账助手 --onefile gui_app.py
```

说明：
- `onefile` 启动速度通常慢于 `onedir`。
- 首选 `onedir`，稳定性更好。
