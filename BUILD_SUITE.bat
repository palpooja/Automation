@echo off
rem ------------------------------------------------------------------
rem  Builds dist\URA_Suite.exe   (no bundled browser: uses the Google
rem  Chrome or Microsoft Edge already installed on the PC).
rem  --collect-all playwright is REQUIRED: it packs Playwright's driver
rem  (node + package); without it the program cannot start any browser.
rem ------------------------------------------------------------------
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" py -3 -m venv .venv || goto :fail
set "PY=.venv\Scripts\python.exe"
echo [1/3] packages...
"%PY%" -m pip install --quiet --upgrade pip
"%PY%" -m pip install --quiet -r requirements-build.txt || goto :fail
echo [2/3] checking the code loads...
"%PY%" -c "import app" || goto :fail
echo [3/3] building (3-6 minutes)...
"%PY%" -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name URA_Suite --icon ura_icon.ico --add-data "ura_icon.ico;." ^
    --paths . --hidden-import engines.ledger_engine --hidden-import engines.returns_common ^
    --hidden-import engines.wht_outputs --hidden-import engines.wht_engine --hidden-import engines.returns_pw --hidden-import engines.paye_engine ^
    --hidden-import engines.dialogs --hidden-import engines.sheet_safety --hidden-import engines.excel_files --hidden-import engines.ledger_workbook --hidden-import lxml --hidden-import lxml.etree ^
    --hidden-import lxml.html --hidden-import lxml._elementpath --add-data "fonts;fonts" ^
    --collect-all playwright ^
    app.py > build_log.txt 2>&1 || goto :fail
echo FINISHED: %cd%\dist\URA_Suite.exe
start "" "%cd%\dist"
pause
exit /b 0
:fail
echo Something went wrong - send the text above and build_log.txt.
pause
exit /b 1
