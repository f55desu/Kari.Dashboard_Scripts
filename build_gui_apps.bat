@echo off
REM Подключаем сетевой диск временно для запуска скрипта из сетевой папки
pushd "%~dp0"

echo ================================================
echo Сборка GUI приложений для дашбордов
echo ================================================
echo.

REM Настройка папки вывода (измените на нужную)
set OUTPUT_DIR=.\apps

REM Создаем папку если не существует
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

REM Проверка установлен ли PyInstaller
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo PyInstaller не найден. Устанавливаем...
    python -m pip install pyinstaller
    echo.
)

REM Проверка установлена ли библиотека pyarrow
python -m pip show pyarrow >nul 2>&1
if errorlevel 1 (
    echo pyarrow не найден. Устанавливаем...
    python -m pip install pyarrow
    echo.
)

@REM echo [1/4] Сборка Dashboard_GUI.py (WB Wrappers)...
@REM pyinstaller --onefile --windowed --icon=app_WB.ico --distpath="%OUTPUT_DIR%" --name="Dashboard_WB_Wrappers" Dashboard_GUI.py
@REM echo.

@REM echo [2/4] Сборка Dashboard_OZON_GUI.py (OZON Wrappers)...
@REM pyinstaller --onefile --windowed --icon=app_OZ.ico --distpath="%OUTPUT_DIR%" --name="Dashboard_OZON_Wrappers" Dashboard_OZON_GUI.py
@REM echo.

echo [3/4] Сборка DashboardAssemblyWB_GUI.py (WB Assembly)...
pyinstaller --onefile --windowed --icon=app_WB.ico --hidden-import=pyarrow --collect-all=pyarrow --distpath="%OUTPUT_DIR%" --name="Dashboard_WB_Assembly" DashboardAssemblyWB_GUI.py
echo.

echo [4/4] Сборка DashboardAssemblyOZ_GUI.py (OZON Assembly)...
pyinstaller --onefile --windowed --icon=app_OZ.ico --hidden-import=pyarrow --collect-all=pyarrow --distpath="%OUTPUT_DIR%" --name="Dashboard_OZ_Assembly" DashboardAssemblyOZ_GUI.py
echo.

echo ================================================
echo Готово! Приложения находятся в папке %OUTPUT_DIR%\
echo ================================================
echo.
echo Файлы:
echo - %OUTPUT_DIR%\Dashboard_WB_Wrappers.exe
echo - %OUTPUT_DIR%\Dashboard_OZON_Wrappers.exe
echo - %OUTPUT_DIR%\Dashboard_WB_Assembly.exe
echo - %OUTPUT_DIR%\Dashboard_OZ_Assembly.exe
echo.
popd
pause
