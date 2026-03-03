# SAP和CBS对账助手（桌面版）快速运行

## 适用对象
- 在 Cursor 中运行项目的同事（不熟悉 Python 也可按步骤执行）。

## 1. 环境准备
- 安装 Python 3.10+（Windows 勾选 `Add Python to PATH`）。
- 打开项目目录：`SAP和CBS对账整月`。

## 2. 一句话运行（推荐）
在 Cursor 终端执行：

```bash
python -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt && python gui_app.py
```

Windows PowerShell 可用：

```powershell
python -m venv .venv; .\.venv\Scripts\Activate.ps1; pip install -r requirements.txt; python gui_app.py
```

## 3. 界面使用
- 选择模式：
  - `明细对账`：上传 `CBS明细` + `SAP明细`
  - `总和对账`：上传 `CBS历史余额` + `SAP科余` + `银行科目`
- 选择输出目录。
- 点击 `开始对账`。
- 程序会输出：
  - Excel 对账结果文件（带时间戳）
  - 文本对账报告（带时间戳）

## 4. 常见问题
- 报错“列名不匹配”：
  - 检查上传文件是否是正确模板，确认表头名称未被改动。
- 报错“不是Excel文件”：
  - 只支持 `.xlsx` / `.xls`。
- 运行时闪退：
  - 请在终端执行 `python gui_app.py`，查看详细错误信息。
