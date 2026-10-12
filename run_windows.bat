@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo Dependency installation failed. Check your Python installation and internet connection.
  pause
  exit /b 1
)
python -m streamlit run app.py
pause
