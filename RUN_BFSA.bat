@echo off
chcp 65001 >nul
title BFSA Excel Екстрактор

echo ========================================
echo    BFSA Excel Екстрактор
echo       by staticmnt
echo ========================================
echo.

:: Проверка дали Python е инсталиран
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python не е инсталиран!
    echo Моля, инсталирайте Python от https://python.org
    echo Уверете се, че сте отметнали "Add Python to PATH" по време на инсталацията
    pause
    exit /b 1
)

:: Проверка дали скрипт файловете съществуват
if not exist "bfsa_excel_app.py" (
    echo ❌ Скрипт файл 'bfsa_excel_app.py' не е намерен!
    pause
    exit /b 1
)

if not exist "bfsa_main.py" (
    echo ❌ Скрипт файл 'bfsa_main.py' не е намерен!
    pause
    exit /b 1
)

echo ✅ Python е инсталиран
echo 🚀 Стартиране на BFSA Excel Екстрактор...
echo.

:: Стартиране на основния скрипт
python bfsa_excel_app.py

pause