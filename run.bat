cd /d "%~dp0"

:: prefer virtual environment python if it exists
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" main.py
) else (
    python main.py
)