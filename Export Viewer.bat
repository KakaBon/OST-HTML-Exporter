@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
  echo Python launcher "py" was not found.
  echo.
  pause
  exit /b 1
)

py -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Installing required package: openpyxl
  py -m pip install openpyxl
  if errorlevel 1 (
    echo.
    echo Failed to install openpyxl.
    pause
    exit /b 1
  )
)

if "%~1"=="" (
  py "%~dp0export_viewer.py"
) else (
  py "%~dp0export_viewer.py" "%~1"
)

if errorlevel 1 (
  echo.
  echo Exporter stopped with an error.
  pause
)

endlocal
