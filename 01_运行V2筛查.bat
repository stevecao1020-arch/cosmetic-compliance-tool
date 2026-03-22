@echo off
setlocal EnableExtensions
cd /d "%~dp0"

call :find_python
if errorlevel 1 goto :end

echo.
echo ========================================
echo Run V2 Excel Screening
echo ========================================
echo Python: %PYTHON_EXE% %PYTHON_ARG%
echo.

"%PYTHON_EXE%" %PYTHON_ARG% "cosmetic_screening_windows_qwen_inline_v2.py"

:end
echo.
pause
exit /b

:find_python
set "PYTHON_EXE="
set "PYTHON_ARG="

if exist "%~dp0.venv\Scripts\python.exe" (
    set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"
    exit /b 0
)

if exist "%LocalAppData%\Programs\Python\Python313\python.exe" (
    set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python313\python.exe"
    exit /b 0
)

if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
    set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python312\python.exe"
    exit /b 0
)

if exist "%LocalAppData%\Programs\Python\Python311\python.exe" (
    set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python311\python.exe"
    exit /b 0
)

for /f "delims=" %%I in ('where python.exe 2^>nul') do (
    if not defined PYTHON_EXE set "PYTHON_EXE=%%~I"
)
if defined PYTHON_EXE exit /b 0

for /f "delims=" %%I in ('where py.exe 2^>nul') do (
    if not defined PYTHON_EXE (
        set "PYTHON_EXE=%%~I"
        set "PYTHON_ARG=-3"
    )
)
if defined PYTHON_EXE exit /b 0

echo [ERROR] Python was not found.
echo Please contact the project owner.
exit /b 1
