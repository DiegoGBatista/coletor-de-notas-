@echo off
REM ============================================================
REM  Gera a pasta "dist\Coletor de Notas Fiscais" (o programa pronto)
REM  Requisitos: Python 3.11+ instalado ("py" no PATH)
REM  Coloque nesta pasta: icon.ico e logo_automacao2000.png
REM ============================================================
setlocal
cd /d "%~dp0"

if not exist icon.ico (
  echo ERRO: copie o icon.ico para esta pasta.
  pause & exit /b 1
)
if not exist logo_automacao2000.png (
  echo ERRO: copie o logo_automacao2000.png para esta pasta.
  pause & exit /b 1
)

py -m pip install --upgrade pip
py -m pip install -r requirements.txt || (echo Falha ao instalar as bibliotecas. & pause & exit /b 1)

py -m PyInstaller --noconfirm --clean --onedir --windowed ^
  --name "Coletor de Notas Fiscais" ^
  --icon icon.ico ^
  --add-data "icon.ico;." ^
  --add-data "logo_automacao2000.png;." ^
  --collect-all customtkinter ^
  --collect-all reportlab ^
  --hidden-import certifi ^
  nfce_manager_fixed.py || (echo Falha no PyInstaller. & pause & exit /b 1)

echo.
echo Pronto! Agora abra "instalador.iss" no Inno Setup e clique em Build ^> Compile.
echo O instalador sera criado na pasta "Output".
pause
