@echo off
chcp 65001 >nul
cd /d "%~dp0"

python -c "import streamlit, docx" 2>nul
if errorlevel 1 (
    echo Устанавливаем необходимые компоненты...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo Не удалось установить компоненты.
        pause
        exit /b 1
    )
)

python -m streamlit run app.py
if errorlevel 1 pause
