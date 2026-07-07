@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title Обзвон CRM - МТС

cd /d "%~dp0"

echo.
echo ========================================
echo   Выгрузка CRM и автообзвон через МТС
echo ========================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python не найден!
    echo Скачай и установи Python отсюда:
    echo https://www.python.org/downloads/
    echo.
    echo ВАЖНО: при установке поставь галочку "Add Python to PATH".
    echo После установки запусти этот файл снова.
    echo.
    start https://www.python.org/downloads/
    pause
    exit
)

echo Устанавливаю/проверяю нужные библиотеки...
pip install selenium beautifulsoup4 webdriver-manager openpyxl requests pandas -q
echo.

:menu
echo ========================================
echo   Выбери действие:
echo ========================================
echo   1. Выгрузить заказы из CRM (и подготовить файл номеров для МТС)
echo   2. Запустить автообзвон в МТС (новое задание, загрузка номеров)
echo   3. Скачать отчёт МТС и объединить с заказами CRM
echo   4. Объединить уже скачанный отчёт вручную
echo   5. ПОЛНЫЙ ЦИКЛ: пункты 1+2+3 подряд
echo   6. Выход
echo ========================================
set /p choice="Введи цифру от 1 до 6: "

if "%choice%"=="1" goto crm
if "%choice%"=="2" goto call
if "%choice%"=="3" goto fetchreport
if "%choice%"=="4" goto merge
if "%choice%"=="5" goto fullcycle
if "%choice%"=="6" goto end
echo Неверный выбор, попробуй ещё раз.
goto menu

:crm
echo.
echo Выгружаю заказы из CRM...
echo.
python scrape_crm.py crm
goto finish

:call
echo.
echo Запускаю автообзвон в МТС (создаю новое задание)...
echo.
python scrape_crm.py call
goto finish

:fetchreport
echo.
if exist last_task_url.txt (
    set /p taskurl=<last_task_url.txt
    echo Использую ссылку последнего задания:
    echo %taskurl%
    echo Если нужна другая ссылка — введи её ниже и нажми Enter.
    echo Если нужна эта — просто нажми Enter.
    set /p "newtaskurl=Ссылка (или Enter для текущей): "
    if not "!newtaskurl!"=="" set taskurl=!newtaskurl!
) else (
    set /p "taskurl=Введи ссылку на задание МТС: "
)
echo.
python scrape_crm.py fetch_report "%taskurl%"
goto finish

:merge
echo.
set "reportfile=report.csv"
set /p "reportfile=Введи имя файла отчёта (Enter - по умолчанию report.csv): "
echo.
echo Использую файл: %reportfile%
python scrape_crm.py merge "%reportfile%"
goto finish

:fullcycle
echo.
echo ========================================
echo   ПОЛНЫЙ ЦИКЛ: Шаг 1 из 3 - выгрузка из CRM
echo ========================================
echo.
python scrape_crm.py crm
if %errorlevel% neq 0 (
    echo.
    echo Шаг 1 завершился с ошибкой, останавливаюсь.
    goto finish
)

echo.
echo ========================================
echo   ПОЛНЫЙ ЦИКЛ: Шаги 2+3 - обзвон и отчёт в одном браузере
echo ========================================
echo.
echo   Обзвон запустится автоматически. После создания задания
echo   скрипт подождёт 15 минут и скачает отчёт в том же окне.
echo ========================================
echo.
python scrape_crm.py full_cycle
goto finish

:finish
echo.
echo ========================================
echo Готово.
echo ========================================
pause
goto menu

:end
exit
