import os
import re
import xlwings as xw
from datetime import datetime, timedelta

from pathlib import Path
import shutil
import win32com.client as win32
import time

import pandas as pd

import subprocess

logs = open('logs.log', 'a')
logs.write(f'{datetime.now()} - OZON Dashboard Wrapper ran\n')

# 🔧 Параметры (КОНСТАНТЫ ПУТЕЙ)
query_name = "Воронка"
# folder_воронка = "ВЫГРУЗКА воронка Озон"
# folder_показатели_по_дням = "Показатели по дням"
# folder_затраты_Озон = "Затраты\\Озон. Затраты из Аналитики"

folder_воронка = r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!\ВЫГРУЗКА воронка Озон"
folder_показатели_по_дням = r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!\Показатели по дням"
folder_затраты_Озон = r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!\Затраты\\Озон. Затраты из Аналитики"
folder_затраты_Озон_NewFormat = r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!\Затраты\\Озон. Затраты из Аналитики New Format"

def copy_to_directory(source_path, target_dir, filename):
    src = os.path.join(source_path, filename)
    dst = os.path.join(target_dir, filename)

    shutil.copy2(src, dst)
    print(f"Скопирован: {src} в {dst}")
    return dst

def get_latest_filename(folder, pattern=r'analytics_report_(\d{4}-\d{2}-\d{2})_\d{2}_\d{2}\.xlsx', format="%Y-%m-%d"):
    files = [f for f in os.listdir(folder) if re.match(pattern, f)]
    if not files:
        raise FileNotFoundError("Нет файлов, соответствующих шаблону")

    def extract_date(file):
        match = re.search(pattern, file)
        if match:
            return datetime.strptime(match.group(1), format)
        return datetime.min

    latest_file = max(files, key=extract_date)
    return latest_file

def get_latest_file_by_pattern(folder, pattern):
    files = [f for f in os.listdir(folder) if re.match(pattern, f)]
    if not files:
        raise FileNotFoundError(f"Нет файлов по шаблону: {pattern}")
    files.sort(key=lambda f: os.path.getmtime(os.path.join(folder, f)), reverse=True)
    return files[0], files[1], files[2]

def copy_and_convert_sku_statistics(downloads_path, filename):
    import os, re, shutil
    from datetime import datetime
    import pandas as pd

    src = os.path.join(downloads_path, filename)

    # 📥 Загружаем Excel. Заголовок — со второй строки (header=1), как у тебя.
    df = pd.read_excel(src, header=1)

    # Удаляем столбец "Название товара в продвижении", если есть
    column_name = "Название товара"
    if column_name in df.columns:
        df.drop(column_name, axis=1, inplace=True)
        print(f"✅ Столбец '{column_name}' был удалён.")
    else:
        print(f"⚠️ Столбец '{column_name}' не найден.")

    column_name = "Место размещения"
    if column_name in df.columns:
        df.drop(column_name, axis=1, inplace=True)
        print(f"✅ Столбец '{column_name}' был удалён.")
    else:
        print(f"⚠️ Столбец '{column_name}' не найден.")

    column_name = "Инструмент"
    target_column_name = "Тип продвижения"
    if column_name in df.columns:
        df.rename(columns={column_name: target_column_name}, inplace=True)
        print(f"✅ Столбец '{column_name}' был изменен на '{target_column_name}'.")
    else:
        print(f"⚠️ Столбец '{column_name}' не найден.")
    df.replace('-', 0.0, inplace=True)
    df['CTR, %'] = pd.to_numeric(df['CTR, %'])
    df['Конверсия в корзину, %'] = pd.to_numeric(df['Конверсия в корзину, %'])
    df['Затраты на заказ, ₽'] = pd.to_numeric(df['Затраты на заказ, ₽'])
    df['Стоимость клика, ₽'] = pd.to_numeric(df['Стоимость клика, ₽'])
    
    # Извлекаем дату из имени файла
    match = re.search(r"Аналитика продвижения_(\d{2}\.\d{2}\.\d{4})\.xlsx", filename)
    if not match:
        raise ValueError(f"Не удалось извлечь дату из имени файла {filename}")
    file_date = match.group(1)

    # Формат "DD.MM.YYYY"
    date_obj = datetime.strptime(file_date, "%d.%m.%Y") - timedelta(days=1)
    formatted_date = date_obj.strftime("%d.%m.%Y")

    # Путь к CSV
    csv_filename = f"{formatted_date}.csv"
    csv_path = os.path.join(downloads_path, csv_filename)

    # Сколько точек с запятой нужно в служебных строках?
    # Должно быть (число_колонок - 1)
    num_cols = len(df.columns)
    semis = ";" * (num_cols - 1)

    # Пишем файл
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        # 1-я строка: "Период ... ;;;;;;;;"
        f.write(f"Период {formatted_date} - {formatted_date}{semis}\n")
        # 2-я строка: только ;;;;;;;;
        f.write(f"{semis}\n")

        # Дальше — заголовок и данные из DataFrame
        df.to_csv(
            f,
            index=False,
            header=True,
            sep=';',
            decimal=',',        # запятая вместо точки
            float_format="%.2f",# 2 знака после запятой
            lineterminator="\r\n"
        )

    print(f"✅ CSV сохранён: {csv_filename}")

    # Копируем куда нужно
    dst_csv = os.path.join(folder_затраты_Озон, csv_filename)
    dst_xlsx = os.path.join(folder_затраты_Озон_NewFormat, filename)
    shutil.copy2(csv_path, dst_csv)
    shutil.copy2(src, dst_xlsx)
    print(f"✅ Копирован CSV файл статистики: {csv_filename}")

    return os.path.join(folder_затраты_Озон, csv_filename)

def process_last_3_sku_files(downloads_path):
    # Шаблон для поиска файлов
    pattern = r"Аналитика продвижения_(\d{2}\.\d{2}\.\d{4})\.xlsx"

    # Получаем последние 3 файла по шаблону
    files = [f for f in os.listdir(downloads_path) if re.match(pattern, f)]
    if len(files) < 3:
        raise FileNotFoundError("Не найдено достаточного количества файлов для обработки.")

    # Сортируем файлы по дате (по имени файла, где дата в начале)
    files.sort(key=lambda x: datetime.strptime(re.search(r"(\d{2}.\d{2}.\d{4})", x).group(1), "%d.%m.%Y"), reverse=True)
    
    # Берем последние 3 файла
    file_sku_1, file_sku_2, file_sku_3 = files[:3]

    # Обрабатываем файлы
    file_sku_1_path = copy_and_convert_sku_statistics(downloads_path, file_sku_1)
    file_sku_2_path = copy_and_convert_sku_statistics(downloads_path, file_sku_2)
    file_sku_3_path = copy_and_convert_sku_statistics(downloads_path, file_sku_3)

    return file_sku_1_path, file_sku_2_path, file_sku_3_path

def duplicate_file_with_date_format(folder, filename):
    yesterday = datetime.now() - timedelta(days=1)
    new_filename = yesterday.strftime("%d.%m.%Y") + ".xlsx"
    src_path = os.path.join(folder, filename)
    dst_path = os.path.join(folder, new_filename)

    shutil.copy2(src_path, dst_path)
    print(f"✅ Создана копия файла: {new_filename}")
    return new_filename

def update_power_query_filter():
    latest_filename_воронка = get_latest_filename(folder_воронка)
    latest_filename_по_дням = get_latest_filename(folder_показатели_по_дням, pattern=r'(\d{2}.\d{2}.\d{4})', format="%d.%m.%Y")

    # 🗂 Создаём копию файла с названием "dd.mm.YYYY.xlsx"
    dated_filename_по_дням = duplicate_file_with_date_format(folder_показатели_по_дням, latest_filename_по_дням)
    dated_path = os.path.join(folder_показатели_по_дням, dated_filename_по_дням)

    print(f"📂 Актуальный файл воронки: {latest_filename_воронка}")
    print(f"📄 Работаем с файлом-копией: {dated_filename_по_дням}")

    # 🔧 Этап 1: Открываем через xlwings и меняем M-код
    app = None
    try:
        app = xw.App(visible=False)
        wb = app.books.open(dated_path)
        xl = wb.app.api
        xl_wb = xl.Workbooks(os.path.basename(dated_filename_по_дням))

        # Найдём запрос
        for q in xl_wb.Queries:
            if q.Name == query_name:
                m_code = q.Formula
                break
        else:
            raise ValueError("❌ Запрос не найден")

        # Заменяем имя файла в фильтре
        updated_code = re.sub(
            r'Table\.SelectRows\([^)]*Name[^"]*"[^"]+\.xlsx"',
            lambda m: re.sub(r'"[^"]+\.xlsx"', f'"{latest_filename_воронка}"', m.group(0)),
            m_code
        )

        if updated_code == m_code:
            raise ValueError("⚠️ Не удалось заменить имя файла в фильтре")

        # Обновляем формулу
        q.Formula = updated_code

        # Сохраняем и закрываем книгу
        wb.save()
        wb.close()
        print("✏️ M-код обновлён и сохранён")

    finally:
        if app:
            app.quit()
        time.sleep(10)  # 💤 Пауза, чтобы Excel освободил ресурсы

    # 🔄 Этап 2: Перезапуск Excel и обновление данных
    if not os.path.exists(dated_path):
        raise FileNotFoundError(f"❌ Файл не найден: {dated_path}")

    excel = win32.Dispatch("Excel.Application")
    excel.DisplayAlerts = False

    try:
        dated_path = os.path.abspath(dated_path)
        wb = excel.Workbooks.Open(dated_path)
        if wb is None:
            raise RuntimeError(f"❌ Excel не смог открыть файл: {dated_path}")

        time.sleep(5)  # ⏳ Дать Power Query и подключению инициализироваться
        wb.RefreshAll()
        print("✏️ Обновляем данные Power Querry")
                # # Обновить сводные таблицы
        sheet = wb.Sheets("Лист3")
        for pivot in sheet.PivotTables():
            pivot.RefreshTable()
        excel.CalculateUntilAsyncQueriesDone()
        wb.Save()
        wb.Close()
        print(f"✅ Данные в '{dated_path}' обновлены")
    finally:
        excel.Quit()

# Подготавливаем 3 дня сразу, пятница, суббота и воскресенье
copyFrom = r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Taldykin"
pattern = r"Аналитика продвижения_(\d{2}\.\d{2}\.\d{4})\.xlsx"
# file_sku_1, file_sku_2, file_sku_3 = get_latest_file_by_pattern(copyFrom, pattern)
# file_sku_1 = copy_and_convert_sku_statistics(filename=file_sku_1, downloads_path=copyFrom)
# file_sku_2 = copy_and_convert_sku_statistics(filename=file_sku_2, downloads_path=copyFrom)
# file_sku_3 = copy_and_convert_sku_statistics(filename=file_sku_3, downloads_path=copyFrom)

file_sku_1, file_sku_2, file_sku_3 = process_last_3_sku_files(copyFrom)

pattern = r"analytics_report_\d{4}-\d{2}-\d{2}_\d{2}_\d{2}\.xlsx"
file_report_1, file_report_2, file_report_3 = get_latest_file_by_pattern(copyFrom, pattern)

# Начинаем с 3 файла (с пятницы). Т.е. в порядке пятница-суббота-понедельник
file_report_3 = copy_to_directory(copyFrom, folder_воронка, file_report_3)
# update_power_query_filter() # Обновляем данные в новом файле
file_report_2 = copy_to_directory(copyFrom, folder_воронка, file_report_2)
# update_power_query_filter() # Обновляем данные в новом файле
file_report_1 = copy_to_directory(copyFrom, folder_воронка, file_report_1)
# update_power_query_filter() # Обновляем данные в новом файле

# Запускаем скрипт
# subprocess.run(['DB_OZ_1.1.exe'], check=True)a

logs = open('logs.log', 'a')
logs.write(f'{datetime.now()} - OZON Dashboard Wrapper completed\n')