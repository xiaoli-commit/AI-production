@echo off
setlocal
cd /d "%~dp0"
set "PYTHON=%LocalAppData%\Python\pythoncore-3.14-64\python.exe"

if not exist "%PYTHON%" (
  echo 未找到本机 Python 3.14：%PYTHON%
  echo 请安装 Python 3.14 后重新运行此脚本。
  pause
  exit /b 1
)

"%PYTHON%" -c "import fastapi, uvicorn, multipart, docx, openpyxl, numpy, pymupdf, rapidocr_onnxruntime" >nul 2>&1
if errorlevel 1 (
  echo 正在安装文档提取所需组件...
  "%PYTHON%" -m pip install -r "%~dp0requirements.txt"
  if errorlevel 1 (
    echo 依赖安装失败，请检查网络后重试。
    pause
    exit /b 1
  )
)

set "READY="
"%PYTHON%" -c "from urllib.request import urlopen; urlopen('http://127.0.0.1:8765/api/health', timeout=1)" >nul 2>&1
if not errorlevel 1 set "READY=1"
if not defined READY start "采购审查平台本地服务" /min "%PYTHON%" -m uvicorn app:app --host 127.0.0.1 --port 8765
for /l %%i in (1,1,30) do (
  if not defined READY (
    "%PYTHON%" -c "from urllib.request import urlopen; urlopen('http://127.0.0.1:8765/api/health', timeout=1)" >nul 2>&1
    if not errorlevel 1 set "READY=1"
    if not defined READY timeout /t 1 /nobreak >nul
  )
)

if not defined READY (
  echo 本地服务启动失败。请确认 8765 端口未被其他程序占用。
  pause
  exit /b 1
)

start "" "http://127.0.0.1:8765/"
echo 采购审查平台已在本机打开。关闭“采购审查平台本地服务”窗口即可停止服务。