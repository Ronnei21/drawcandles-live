@echo off
chcp 65001 >nul
setlocal

echo ============================================================
echo  DrawCandles Live - Gerador de EXE
echo ============================================================
echo.

echo [1/5] Verificando PyInstaller...
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo       Instalando PyInstaller...
    python -m pip install --upgrade pyinstaller
    if errorlevel 1 (
        echo ERRO: Nao consegui instalar o PyInstaller.
        pause
        exit /b 1
    )
) else (
    echo       OK.
)
echo.

REM ---------- Arquivos ----------
set SCRIPT=drawcandleslive_v7.0.1.py
set ICONE_PNG=icone.png
set ICONE_ICO=icon-2.ico
set ASSETS_DIR=assets

if not exist "%SCRIPT%" (
    echo ERRO: Nao encontrei "%SCRIPT%".
    pause
    exit /b 1
)
if not exist "%ICONE_PNG%" (
    echo ERRO: Nao encontrei "%ICONE_PNG%".
    pause
    exit /b 1
)
if not exist "%ICONE_ICO%" (
    echo ERRO: Nao encontrei "%ICONE_ICO%".
    pause
    exit /b 1
)
if not exist "%ASSETS_DIR%" (
    echo ERRO: Nao encontrei a pasta "%ASSETS_DIR%".
    pause
    exit /b 1
)

echo [2/5] Limpando builds anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "drawcandleslive.spec" del /q "drawcandleslive.spec"
echo       OK.
echo.

echo [3/5] Gerando o EXE com icone e assets...
echo.

python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "drawcandleslive" ^
    --icon "%ICONE_ICO%" ^
    --add-data "%ICONE_PNG%;." ^
    --add-data "%ASSETS_DIR%;assets" ^
    --hidden-import "win32timezone" ^
    --hidden-import "win32api" ^
    --hidden-import "win32con" ^
    --hidden-import "win32gui" ^
    --collect-all "win32com" ^
    --collect-all "pythoncom" ^
    --clean ^
    --noconfirm ^
    "%SCRIPT%"

if errorlevel 1 (
    echo.
    echo ERRO: Falha ao gerar o EXE.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  PRONTO!
echo ============================================================
echo.
echo O EXE esta em:  dist\drawcandleslive.exe
echo.
echo Ele ja vem com:
echo   - Icone personalizado
echo   - Pasta assets embutida (toolbar + relogio)
echo   - Todos os modulos win32
echo.

if exist dist start "" "dist"
pause
endlocal