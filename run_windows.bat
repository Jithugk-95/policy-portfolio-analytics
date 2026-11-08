@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\streamlit.exe" (
  echo Virtual environment not found.
  echo Run setup commands in README.md first.
  pause
  exit /b 1
)
call ".venv\Scripts\activate.bat"
streamlit run app.py
pause
