# Инструкция по сборке GUI приложений

## Подготовка

### 1. Установка PyInstaller

Откройте командную строку и выполните:

```bash
pip install pyinstaller
```

### 2. Подготовка иконок

Убедитесь, что в директории со скриптами есть файлы иконок:
- `app_WB.ico` - для приложений Wildberries
- `app_OZ.ico` - для приложений Ozon

Если иконок нет, создайте их или используйте стандартную иконку Python.

## Быстрая сборка (рекомендуется)

**Просто запустите:**
```bash
build_gui_apps.bat
```

Этот скрипт автоматически:
1. Проверит установлен ли PyInstaller (и установит если нет)
2. Соберет все 4 GUI приложения
3. Создаст .exe файлы в папке `apps\` (по умолчанию)

### Изменение папки вывода

Откройте `build_gui_apps.bat` и измените переменную `OUTPUT_DIR`:

```batch
set OUTPUT_DIR=.\apps          REM По умолчанию - папка apps в текущей директории
set OUTPUT_DIR=C:\MyApps       REM Абсолютный путь
set OUTPUT_DIR=..\..\Programs  REM Относительный путь
```

## Ручная сборка отдельных приложений

### Dashboard_GUI.py (WB Wrappers)
```bash
pyinstaller --onefile --windowed --icon=app_WB.ico --name="Dashboard_WB_Wrappers" Dashboard_GUI.py
```

### Dashboard_OZON_GUI.py (OZON Wrappers)
```bash
pyinstaller --onefile --windowed --icon=app_OZ.ico --name="Dashboard_OZON_Wrappers" Dashboard_OZON_GUI.py
```

### DashboardAssemblyWB_GUI.py (WB Assembly)
```bash
pyinstaller --onefile --windowed --icon=app_WB.ico --name="Dashboard_WB_Assembly" DashboardAssemblyWB_GUI.py
```

### DashboardAssemblyOZ_GUI.py (OZON Assembly)
```bash
pyinstaller --onefile --windowed --icon=app_OZ.ico --name="Dashboard_OZ_Assembly" DashboardAssemblyOZ_GUI.py
```

## Параметры PyInstaller

- `--onefile` - создает один .exe файл (все зависимости внутри)
- `--windowed` - запуск без консоли (для GUI приложений)
- `--icon=` - путь к файлу иконки (.ico)
- `--name=` - имя выходного файла

## Структура после сборки

```
Kari.Dashboard_Scripts/
├── build/          # Временные файлы сборки (можно удалить)
├── dist/           # Готовые .exe файлы
│   ├── Dashboard_WB_Wrappers.exe
│   ├── Dashboard_OZON_Wrappers.exe
│   ├── Dashboard_WB_Assembly.exe
│   └── Dashboard_OZ_Assembly.exe
└── *.spec          # Спецификации PyInstaller (можно сохранить для повторных сборок)
```

## Распространение приложений

Готовые .exe файлы из папки `dist\` можно:
1. Копировать в любую папку и запускать
2. Создать ярлыки на рабочем столе
3. Распространять между коллегами

**ВАЖНО:** .exe файлы должны находиться в той же директории, что и Python скрипты (Dashboard_*.py), так как они создают временные файлы и запускают оригинальные скрипты.

## Очистка временных файлов

После успешной сборки можно удалить:
- Папку `build\`
- Файлы `*.spec` (если не планируете изменять параметры сборки)

Папку `dist\` **НЕ УДАЛЯЙТЕ** - в ней находятся готовые приложения!

## Обновление приложений

При изменении исходных Python файлов просто запустите `build_gui_apps.bat` заново - старые .exe будут перезаписаны новыми версиями.

## Устранение проблем

### "PyInstaller не найден"
```bash
pip install --upgrade pyinstaller
```

### "Не удается найти модуль XXX"
Убедитесь, что все необходимые библиотеки установлены:
```bash
pip install tkinter pandas xlwings win32com pyodbc sqlalchemy openpyxl
```

### Иконка не применяется
- Проверьте, что файлы `app_WB.ico` и `app_OZ.ico` существуют в директории
- Используйте абсолютный путь к иконке: `--icon=C:\path\to\app_WB.ico`

### Приложение не запускается
- Попробуйте собрать без флага `--windowed` чтобы видеть ошибки в консоли
- Проверьте, что все зависимости доступны на целевой машине
