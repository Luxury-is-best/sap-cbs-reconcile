@echo off
chcp 65001 >nul
echo.
echo === SAP/CBS 对账助手 - Windows 打包 ===
echo 说明：必须在 Windows 电脑上运行本脚本（Mac 无法直接生成 .exe）。
echo.
where py >nul 2>&1
if errorlevel 1 (
  echo 未找到 py 启动器，请安装 Python 3.10+ 并勾选 "Add to PATH"。
  echo 下载：https://www.python.org/downloads/
  pause
  exit /b 1
)

cd /d "%~dp0"
py -3 -m pip install -U pip
py -3 -m pip install -r requirements_gui_build.txt
py -3 -m PyInstaller --noconfirm sap_cbs_gui.spec

if errorlevel 1 (
  echo 打包失败，请把上方报错截图发给你同事或开发者。
  pause
  exit /b 1
)

echo.
echo 完成。可执行文件位置：
echo   %~dp0dist\SAP_CBS_对账助手.exe
echo 把整个 dist 里的 exe 发给同事即可（单文件，无需 Python）。
echo 若杀软误报，可让同事添加信任或改用「文件夹模式」打包（需再说明）。
echo.
pause
