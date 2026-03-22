@echo off
setlocal EnableExtensions
cd /d "%~dp0"

call :find_python
if errorlevel 1 goto :end

echo.
echo ========================================
echo Generate V2 Word Reports
echo ========================================
echo Leave formula blank to generate all formula reports.
echo Leave workbook path blank to use the latest workbook in output.
echo.

set "FORMULA="
set "WORKBOOK="
set "LATEST_WORKBOOK="

set /p FORMULA=Formula code (blank = all): 
set /p WORKBOOK=Workbook path (blank = latest): 

if not defined WORKBOOK call :find_latest_workbook

echo.
echo Python: %PYTHON_EXE% %PYTHON_ARG%
echo.

if defined FORMULA (
    if defined WORKBOOK (
        "%PYTHON_EXE%" %PYTHON_ARG% "cosmetic_screening_windows_qwen_inline_report_v2.py" --workbook "%WORKBOOK%" --formula "%FORMULA%"
    ) else if defined LATEST_WORKBOOK (
        "%PYTHON_EXE%" %PYTHON_ARG% "cosmetic_screening_windows_qwen_inline_report_v2.py" --workbook "%LATEST_WORKBOOK%" --formula "%FORMULA%"
    ) else (
        "%PYTHON_EXE%" %PYTHON_ARG% "cosmetic_screening_windows_qwen_inline_report_v2.py" --formula "%FORMULA%"
    )
) else (
    if defined WORKBOOK (
        "%PYTHON_EXE%" %PYTHON_ARG% "cosmetic_screening_windows_qwen_inline_report_v2.py" --workbook "%WORKBOOK%" --all
    ) else if defined LATEST_WORKBOOK (
        "%PYTHON_EXE%" %PYTHON_ARG% "cosmetic_screening_windows_qwen_inline_report_v2.py" --workbook "%LATEST_WORKBOOK%" --all
    ) else (
        "%PYTHON_EXE%" %PYTHON_ARG% "cosmetic_screening_windows_qwen_inline_report_v2.py" --all
    )
)

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

:find_latest_workbook
for /f "delims=" %%I in ('dir /b /a-d /o-d "output\\Cosmetic_Compliance_Output_*.xlsx" 2^>nul') do (
    if not defined LATEST_WORKBOOK set "LATEST_WORKBOOK=%~dp0output\%%~I"
)
exit /b 0
