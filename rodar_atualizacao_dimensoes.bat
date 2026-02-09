@echo off
chcp 65001 >nul
setlocal EnableExtensions

set "PROJECT_DIR=C:\Users\felipe.maffra\Desktop\Python\Dimensões Indicador"
set "SCRIPT=main.py"

cd /d "%PROJECT_DIR%" || (
  echo Falha ao acessar "%PROJECT_DIR%"
  pause
  exit /b 1
)

if not exist ".venv\Scripts\activate.bat" (
  echo ERRO: nao encontrei ".venv\Scripts\activate.bat"
  pause
  exit /b 1
)

call ".venv\Scripts\activate.bat"

if not exist "%SCRIPT%" (
  echo ERRO: nao encontrei "%SCRIPT%"
  pause
  exit /b 1
)

python "%SCRIPT%"
set "EC=%ERRORLEVEL%"

if not "%EC%"=="0" (
  echo ERRO: exit code %EC%
  pause
  exit /b %EC%
)

echo OK: concluido.
pause
exit /b 0

