@echo off
chcp 65001

REM === Переходим в папку, где лежит этот .bat ===
cd /d %~dp0

REM === Поднимаемся на уровень выше (корень проекта) ===
cd ..

REM === Сохраняем корень проекта в переменную ===
set ROOT=%CD%
echo Корень проекта: %ROOT%

REM === Параметры сборки ===
set MAIN_SCRIPT=application\main_window.py
set ICON=application\icon.ico
set NAME=my_app
set PYTHONPATH=%ROOT%

REM === Сборка exe в корень проекта ===
pyinstaller --onefile --windowed --icon "%ICON%" --distpath . --name "%NAME%" --paths=%ROOT% "%MAIN_SCRIPT%"

REM === Удаляем временную папку build ===
rmdir /S /Q build

echo.
echo Сборка завершена! Готовый exe: %ROOT%\%NAME%.exe
