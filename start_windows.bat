@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Sanal ortam bulunamadi. Once README_TR.md icindeki kurulum adimlarini uygulayin.
  pause
  exit /b 1
)
call ".venv\Scripts\activate.bat"
python scripts\preflight.py
if errorlevel 1 (
  echo On kontrol basarisiz. Yukaridaki hatalari duzeltin.
  pause
  exit /b 1
)
python -m streamlit run app.py
