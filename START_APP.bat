@echo off
cd /d "%~dp0"

python -c "import streamlit, docx" 2>nul
if errorlevel 1 (
    echo Installing required components...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo Installation failed.
        pause
        exit /b 1
    )
)

python -m streamlit run app.py
if errorlevel 1 pause
