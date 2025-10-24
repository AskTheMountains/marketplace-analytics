import os
import sys
from pathlib import Path

if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).absolute().parent.parent

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))


# Путь до venv\Scripts\python.exe (Windows)
VENV_PYTHON = BASE_DIR / ".venv" / "Scripts" / "python.exe"

if VENV_PYTHON.exists():
    PYTHON_EXE = str(VENV_PYTHON)
else:
    PYTHON_EXE = "python"  # fallback на глобальный python, если venv не найден

# Иконка приложения
icon_path = BASE_DIR / 'application' / 'icon.ico'

# Файл настроек
SETTINGS_JSON = (BASE_DIR / "settings.json").resolve()
