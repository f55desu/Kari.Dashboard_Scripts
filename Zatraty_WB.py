import os
import pandas as pd
from sqlalchemy import create_engine
import pyodbc
import win32com.client as win32
import time  # Для измерения времени выполнения
import shutil
import re
from functools import reduce

# Функция для форматирования даты в строковый формат 'DD.MM.YYYY'
def format_date_column(df, date_column):
    if date_column in df.columns:
        df[date_column] = pd.to_datetime(df[date_column], errors='coerce', dayfirst=True).dt.strftime('%d.%m.%Y')
    return df

def format_elapsed_time(seconds):
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{int(hours)} часа(ов) {int(minutes)} минут(ы) {seconds:.2f} секунд"
# Функция для подключения к SQL Server с аутентификацией Windows
def connect_to_sql(server, database):
    connection_string = (
        f"mssql+pyodbc://{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server"
        "&trusted_connection=yes"
    )
    Engine = create_engine(connection_string)
    return Engine
# Функция для проверки, является ли файл скрытым (для Windows)
def is_hidden(file_path):
    try:
        # Получаем атрибуты файла
        file_attributes = os.stat(file_path).st_file_attributes
        # Проверяем, установлен ли флаг "скрытый"
        return file_attributes & 2 != 0  # 2 соответствует атрибуту "скрытый"
    except Exception:
        # Если возникла ошибка, считаем файл не скрытым
        return False
# Функция для обработки ошибок и замены их на null
def handle_errors(df):
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].apply(lambda x: None if isinstance(x, str) and x.strip() == '' else x)
    return df

# 1. Объявление констант
# FOLDER_PATH = os.path.normpath(r"!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!")
# FOLDER_PATH_FEATURES = os.path.normpath(r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Taldykin\Дашбоард по рекламным кампаниям")
# FOLDER_PATH_FOR_DB = os.path.normpath(r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Taldykin\Дашбоард по рекламным кампаниям")
# FOLDER_PATH_DUDL = os.path.normpath(r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Taldykin")

FOLDER_PATH = os.path.normpath(r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!")
FOLDER_PATH_FEATURES = r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Дашбоард по рекламным кампаниям"
FOLDER_PATH_FOR_DB = os.path.normpath(r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям")
FOLDER_PATH_DUDL = os.path.normpath(r"\\kari.local\public\all\Агрегаторы\Дашборд реклама WB_OZ")

SQL_SERVER = "cl01sql"
SQL_DATABASE_DBREPORT = "DBReport"
SQL_DATABASE_DBPARTNERS = "DBPartners"

start_all_time = time.time()
# 2. Получить данные из файла "Затраты ВБ_2.xlsx"
try:
    print("Начинаем получать данные для Затрат ВБ...")
    start_time = time.time()  # Запускаем таймер
    file_path_expenses = os.path.join(FOLDER_PATH, "Затраты ВБ_2.xlsx")

    if os.path.exists(file_path_expenses):
        # Список столбцов, которые нужно взять из файла
        columns_to_read = [
            "Дата", "Артикул WB", "ID кампании", "Показы", "Клики",
            "Кол-во добавлений в корзину", "Заказы, шт",
            "Кол-во заказаных товаров, шт", "Заказов на сумму",
            "Расход, с НДС"
        ]

        # Переименование столбцов
        column_mapping = {
            "Показы": "Рекламные показы",
            "Клики": "Рекламные клики",
            "Кол-во добавлений в корзину": "Рекламные в корзину",
            "Заказы, шт": "Рекламные Заказы, шт",
            "Кол-во заказаных товаров, шт": "Рекламные заказаных товаров, шт",
            "Заказов на сумму": "Рекламные заказов на сумму",
            "Расход, с НДС": "Расход, руб"
        }

        # Типы данных для столбцов
        column_dtypes = {
            "Артикул WB": str,
            "Рекламные показы": int,
            "Рекламные клики": int,
            "Рекламные в корзину": int,
            "Рекламные Заказы, шт": int,
            "Рекламные заказаных товаров, шт": int,
            "Рекламные заказов на сумму": int,
            "Расход, руб": int
        }

        # Чтение файла с указанием нужных столбцов
        df_expenses = pd.read_excel(file_path_expenses, sheet_name="Затраты ВБ", engine="openpyxl",
                                    usecols=columns_to_read)

        # Переименование столбцо
        df_expenses.rename(columns=column_mapping, inplace=True)

        # Форматирование даты
        df_expenses = format_date_column(df_expenses, 'Дата')
        # Вывод первых 5 строк
        print("Первые 5 строк таблицы Затраты:")
        print(df_expenses.head())

        # Сохраняем результат
        elapsed_time = time.time() - start_time  # Вычисляем затраченное время
        print(f"Данные для Затрат ВБ успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
    else:
        print("Файл 'Затраты ВБ_2.xlsx' не найден.")
except Exception as e:
    print(f"Ошибка при получении данных для Затрат ВБ: {e}")

# 3. Файл Campaign_Info
file_path_campaing = os.path.join(FOLDER_PATH, "Campaign_Info.xlsx")
df_campaing = pd.read_excel(file_path_campaing, engine='openpyxl')

df_campaing = df_campaing[['name', 'advertId', 'type']]
df_campaing = df_campaing.rename(columns={'type':'Раздел'})
df_campaing['Раздел'] = df_campaing['Раздел'].astype(str)
type_dictionary = {"8":"Автоматическое", 
                   "9":"Аукцион",
                   "4":"Каталог",
                   "5":"Карточка",
                   "6":"Поиск"}

df_campaing["Раздел"] = df_campaing["Раздел"].map(type_dictionary)
df_campaing

# 4. Подготовка файла ЗатратыВБ_2
# 1) Подготовим кампании: one-hot по "Раздел" (динамический набор столбцов)
df_c = df_campaing[['advertId', 'Раздел']].dropna().copy()
dummies = pd.get_dummies(df_c['Раздел'])          # столбцы = уникальные значения "Раздел"
df_c = pd.concat([df_c[['advertId']], dummies], axis=1)

# Если на один advertId есть несколько строк с разными "Раздел" — агрегируем флагами
df_c = df_c.groupby('advertId', as_index=False).max()

# 2) Присоединим к таблице с артикулами
cols_campaign = dummies.columns.tolist()           # имена динамических столбцов
df_result = (
    df_expenses.merge(df_c, left_on='ID кампании', right_on='advertId', how='left')
         .drop(columns=['advertId'])
)

# 3) Заполним пропуски нулями и приведём к int (0/1)
if cols_campaign:  # на случай, если внезапно "Раздел" пустой
    df_result[cols_campaign] = df_result[cols_campaign].fillna(0).astype(int)

# Готово: df_result имеет вид
# АртикулWB | ID кампании | Автоматическое | Аукцион | Каталог | Карточка | Поиск
df_expenses = df_result

# Список датафреймов по категориям
categories = df_campaing['Раздел'].unique().tolist()
dfs_by_cat = []
for cat in categories:
    if cat in df_result.columns:   # проверим, что колонка существует
        df_group = df_result[df_result[cat] == 1].copy()
        dfs_by_cat.append(df_group)

# 5. Получить данные таблицы с SQL (Цены)
try:
    print("Начинаем получать данные для Цен...")
    start_time = time.time()  # Запускаем таймер
    Engine = connect_to_sql(SQL_SERVER, SQL_DATABASE_DBPARTNERS)
    query_price = f"""
        SELECT DT AS [Дата], ITEMID AS [Артикул], AVG(SITE_PRICE) AS [Цена]
        FROM [DBPartners].[dbo].[WblmRepPriceDiscountWbReport]
        WHERE dt >= '{(pd.Timestamp.today() - pd.DateOffset(months=3)).strftime("%Y-%m-%d")}'
        GROUP BY DT, ITEMID
        ORDER BY DT
        DESC
    """
    df_price = pd.read_sql(query_price, Engine)
    df_price = format_date_column(df_price, 'Дата')

    # Вывод первых 5 строк

    print("Первые 5 строк таблицы Цена:")
    print(df_price.head())

    elapsed_time = time.time() - start_time  # Вычисляем затраченное время
    print(f"Данные для Цен успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
except Exception as e:
    print(f"Ошибка при получении данных для Цен: {e}")

# 6. Получить данные из файла "Справочник.xlsx"
try:
    print("Начинаем получать данные для Справочника...")
    start_time = time.time()  # Запускаем таймер
    file_path_reference = os.path.join(FOLDER_PATH, "Справочник.xlsx")

    if os.path.exists(file_path_reference):
        # Список столбцов, которые нужно взять из файла
        columns_to_read = [
            "Артикул", "Наименование", "Коллекция",
            "Бренд", "Размер", "Сезон", "Направление", "Розничный отдел",
            "Модель", "Группа", "Бизнес-группа", "Техсегмент",
            "Байер", "Две последние коллекции", "Основной артикул",
            "Артикул WB", "Себестоимость с НДС",
            "Процент выкупа ВБ", "НДС", "Ответственный за группу", "Группа для отчетов"
        ]

        # Типы данных для столбцов
        column_dtypes = {
            "Артикул": str,
            "Наименование": str,
            "Коллекция": str,
            "Размер": str,
            "Бренд": str,
            "Сезон": str,
            "Направление": str,
            "Розничный отдел": str,
            "Модель": str,
            "Группа": str,
            "Бизнес-группа": str,
            "Техсегмент": str,
            "Байер": str,
            "Две последние коллекции": str,
            "Основной артикул": str,
            "Артикул WB": str,
            "Себестоимость с НДС": float,
            "Процент выкупа ВБ": float,
            "НДС": int,
            "Ответственный за группу": str,
            "Группа для отчетов":str
        }

        # Чтение файла с указанием нужных столбцов и типов данных
        df_reference = pd.read_excel(
            file_path_reference,
            sheet_name="Выгрузка для справочника",
            engine="openpyxl",
            usecols=columns_to_read,
            dtype=column_dtypes
        )

        # Удаление дубликатов
        df_reference = df_reference.drop_duplicates()

        # Вывод первых 5 строк
        print("Первые 5 строк таблицы Справочник:")
        print(df_reference.head())

        # Сохраняем результат
        elapsed_time = time.time() - start_time  # Вычисляем затраченное время
        print(f"Данные для Справочника успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
    else:
        print("Файл 'Справочник.xlsx' не найден.")
except Exception as e:
    print(f"Ошибка при получении данных для Справочника: {e}")

# 5. Получить данные из файлов вложенной папки "Показатели по неделям ВБ"
try:
    print("Начинаем получать данные для Воронки...")
    start_time = time.time()  # Запускаем таймер
    folder_path_weeks = os.path.join(FOLDER_PATH, "Показатели по неделям ВБ")
    df_funnel = pd.DataFrame()

    if os.path.exists(folder_path_weeks):
        for file in os.listdir(folder_path_weeks):
            file_path = os.path.join(folder_path_weeks, file)

            # Пропускаем скрытые файлы
            if is_hidden(file_path):
                print(f"Пропущен скрытый файл: {file}")
                continue

            # Проверяем расширение файла
            if file.endswith((".xlsx", ".xls")):
                try:
                    # Список всех возможных столбцов
                    all_columns = [
                        "Номенклатура", "Рейтинг карточки", "Переходы в карточку",
                        "Положили в корзину", "Заказали, шт", "Выкупили, шт",
                        "Отменили, шт", "Заказали на сумму, руб", "Выкупили на сумму, руб",
                        "Отменили на сумму, руб", "Средняя цена, руб", "Дата", "Рейтинг по отзывам"
                    ]

                    if file.endswith(".xlsx"):
                        temp_df = pd.read_excel(file_path, sheet_name="воронка ОЗОН", engine="openpyxl")
                    elif file.endswith(".xls"):
                        temp_df = pd.read_excel(file_path, sheet_name="воронка ОЗОН", engine="xlrd")

                    # Добавляем отсутствующие столбцы со значением None
                    for col in all_columns:
                        if col not in temp_df.columns:
                            temp_df[col] = None

                    # Выбираем только нужные столбцы
                    temp_df = temp_df[all_columns]

                    # Переименование столбцов
                    temp_df.rename(columns={
                        "Номенклатура": "Артикул WB",
                        "Рейтинг карточки": "Рейтинг карточки",
                        "Переходы в карточку": "Показы на карточке товара",
                        "Положили в корзину": "Положили в корзину",
                        "Заказали, шт": "Заказали, шт",
                        "Выкупили, шт": "Выкупили, шт",
                        "Отменили, шт": "Отменили, шт",
                        "Заказали на сумму, руб": "Заказали на сумму, руб",
                        "Выкупили на сумму, руб": "Выкупили на сумму, руб",
                        "Отменили на сумму, руб": "Отменили на сумму, руб",
                        "Средняя цена, руб": "Средняя цена, руб",
                        "Рейтинг по отзывам": "Рейтинг по отзывам"  # Новый столбец
                    }, inplace=True)

                    # Типы данных для столбцов
                    column_dtypes = {
                        "Артикул WB": str,
                        "Рейтинг карточки": float,  # Рейтинг может быть дробным числом
                        "Показы на карточке товара": int,
                        "Положили в корзину": int,
                        "Заказали, шт": int,
                        "Выкупили, шт": int,
                        "Отменили, шт": int,
                        "Заказали на сумму, руб": float,
                        "Выкупили на сумму, руб": float,
                        "Отменили на сумму, руб": float,
                        "Средняя цена, руб": float,
                        "Рейтинг по отзывам": float  # Рейтинг по отзывам также может быть дробным
                    }

                    # Применяем типы данных
                    for col, dtype in column_dtypes.items():
                        if col in temp_df.columns:
                            temp_df[col] = temp_df[col].astype(dtype, errors='ignore')

                    # Форматирование даты
                    temp_df = format_date_column(temp_df, 'Дата')

                    # Объединяем временный DataFrame с основным
                    df_funnel = pd.concat([df_funnel, temp_df], ignore_index=True)

                except Exception as e:
                    print(f"Ошибка при чтении файла {file}: {e}")

        # Выбираем финальные столбцы
        final_columns = [
            "Дата", "Артикул WB", "Рейтинг карточки", "Показы на карточке товара",
            "Положили в корзину", "Заказали, шт", "Выкупили, шт",
            "Отменили, шт", "Заказали на сумму, руб", "Выкупили на сумму, руб",
            "Отменили на сумму, руб", "Средняя цена, руб", "Рейтинг по отзывам"
        ]
        df_funnel = df_funnel[final_columns]

        # Заполняем пустые значения в столбце "Рейтинг по отзывам" если они есть
        df_funnel["Рейтинг по отзывам"] = df_funnel["Рейтинг по отзывам"].fillna(0)

        # Вывод первых 5 строк
        print("Первые 5 строк таблицы Воронка:")
        print(df_funnel.head())

        elapsed_time = time.time() - start_time  # Вычисляем затраченное время
        print(f"Данные для Воронки успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
    else:
        print("Папка 'Показатели по неделям ВБ' не найдена.")
except Exception as e:
    print(f"Ошибка при получении данных для Воронки: {e}")


# 4. Правильный джойн кампаний 
from functools import reduce
# ключи для объединения
maybe_keys = ['Артикул WB', 'ID кампании', 'Дата'] + cols_campaign #Базовые столбцы + кампании
keys = [c for c in maybe_keys if c in dfs_by_cat[0].columns]

# категории (в том же порядке, что и список dfs_by_cat)
categories = df_campaing['Раздел'].dropna().unique().tolist()

# вспомогательная функция: префиксовать все НЕключевые колонки
def prefix_non_keys(df, prefix, keep_cols):
    keep = set(keep_cols)
    rename_map = {c: f"{prefix}_{c}" for c in df.columns if c not in keep}
    return df.rename(columns=rename_map)

# 1. Префиксуем все датафреймы в списке
dfs_prefixed = []
for cat, df_cat in zip(categories, dfs_by_cat):
    df_cat_pref = prefix_non_keys(df_cat, cat, keep_cols=keys)
    dfs_prefixed.append(df_cat_pref)

# 2. Смерджим все датафреймы из списка по ключам
df_final = reduce(lambda left, right: pd.merge(left, right, on=keys, how="outer"), dfs_prefixed)

# 3. Опционально — заполнить NaN нулями только в числовых колонках
for col in df_final.columns:
    if col not in keys and pd.api.types.is_numeric_dtype(df_final[col]):
        df_final[col] = df_final[col].fillna(0)

# 5. Добавляем итоговые столбцы метрик по всем кампаниям
from pandas.api.types import is_numeric_dtype
def add_metric_totals_safe(df, decimal_comma=False):
    df = df.copy()
    pattern = re.compile(r'(.+?)_(.+)')   # "<категория>_<метрика>"
    col_map = {}                          # {метрика: [список колонок]}

    # Соберём группы колонок по метрикам
    for col in df.columns:
        m = pattern.fullmatch(col)
        if m:
            metric = m.group(2)
            col_map.setdefault(metric, []).append(col)

    def is_numeric_like(s: pd.Series) -> bool:
        s = s.dropna()
        if s.empty:
            return True
        if is_numeric_dtype(s):
            return True
        if s.dtype == 'object':
            x = s.astype(str)
            if decimal_comma:
                x = x.str.replace(',', '.', regex=False)
            conv = pd.to_numeric(x, errors='coerce')
            return conv.notna().all()   # если есть нечисловые строки → False
        return False

    for metric, cols in col_map.items():
        # Если ХОТЬ В ОДНОМ столбце метрики есть текст — пропускаем всю метрику
        if not all(is_numeric_like(df[c]) for c in cols):
            continue

        # Конвертируем в числа и суммируем
        num_df = pd.DataFrame(index=df.index)
        for c in cols:
            s = df[c]
            if is_numeric_dtype(s):
                num_df[c] = s
            else:
                x = s.astype(str)
                if decimal_comma:
                    x = x.str.replace(',', '.', regex=False)
                num_df[c] = pd.to_numeric(x, errors='coerce')

        total_col = f"{metric}"
        df[total_col] = num_df.sum(axis=1, skipna=True)

        # Опционально: если без дробей — сделать int
        if pd.api.types.is_float_dtype(df[total_col]):
            s = df[total_col]
            if s.dropna().apply(float.is_integer).all():
                df[total_col] = s.astype('int64')

    return df

df_sum = add_metric_totals_safe(df_final)

# 13. Связать "Воронка" с "Справочник"
try:
    print("Начинаем создавать таблицу ВоронкаСправочник...")
    start_time = time.time()  # Запускаем таймер

    # Проверка наличия необходимых столбцов
    required_columns = ["Дата", "Артикул WB"]
    for col in required_columns:
        if col not in df_funnel.columns:
            print(f"Ошибка: Отсутствует столбец '{col}' в df_funnel.")
            exit()

    # Список столбцов, которые нужно взять из справочника
    reference_columns = [
        "Артикул", "Наименование", "Коллекция", "Бренд", "Сезон", "Направление",
        "Розничный отдел", "Модель", "Группа", "Бизнес-группа", "Техсегмент",
        "Байер", "Две последние коллекции", "Основной артикул", "Артикул WB",
        "Себестоимость с НДС", "Процент выкупа ВБ", "НДС", "Ответственный за группу", "Группа для отчетов"
    ]

    # Фильтруем справочник, оставляя только нужные столбцы
    df_reference_filtered = df_reference[reference_columns]

    # Приводим типы данных к строковому формату
    df_funnel["Артикул WB"] = df_funnel["Артикул WB"].fillna('').astype(str)  # Заменяем NaN на пустые строки
    df_reference_filtered["Артикул WB"] = df_reference_filtered["Артикул WB"].fillna('').astype(str)

    # Объединение таблиц
    df_funnel_reference = pd.merge(
        df_funnel,
        df_reference_filtered,
        left_on="Артикул WB",
        right_on="Артикул WB",
        how="left"
    )

    # Удаление дубликатов
    df_funnel_reference = df_funnel_reference.drop_duplicates()

    # Удаление лишних столбцов (если они остались)
    # df_funnel_reference = df_funnel_reference.drop(columns=["Артикул WB"], errors="ignore")

    # Форматирование даты
    df_funnel_reference = format_date_column(df_funnel_reference, 'Дата')

    # Вывод первых 5 строк
    print("Первые 5 строк таблицы ВоронкаСправочник:")
    print(df_funnel_reference.head())

    elapsed_time = time.time() - start_time  # Вычисляем затраченное время
    print(f"Таблица ВоронкаСправочник успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
except Exception as e:
    print(f"Ошибка при создании таблицы ВоронкаСправочник: {e}")

# 14. Связать "ВоронкаСправочник" с "Затраты ВБ"
try:
    print("Начинаем создавать таблицу ВоронкаСправочникЗатраты...")

    # Проверка наличия необходимых столбцов
    required_columns = ["Дата", "Артикул WB"]
    for col in required_columns:
        if col not in df_funnel_reference.columns:
            raise ValueError(f"Ошибка: Отсутствует столбец '{col}' в df_funnel_reference.")
        if col not in df_sum.columns:
            raise ValueError(f"Ошибка: Отсутствует столбец '{col}' в df_expenses.")

    # Очистка и нормализация данных
    df_funnel_reference["Артикул WB"] = (
        df_funnel_reference["Артикул WB"]
        .fillna('')
        .astype(str)
        .str.strip()
        .str.upper()
    )
    df_sum["Артикул WB"] = (
        df_sum["Артикул WB"]
        .fillna('')
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # Объединение таблиц
    df_funnel_expenses = pd.merge(
        df_funnel_reference,
        df_sum,
        left_on=["Дата", "Артикул WB"],
        right_on=["Дата", "Артикул WB"],
        how="left",
        indicator=True  # Для диагностики
    )

    # Проверка результата объединения
    print("Результат объединения:")
    print(df_funnel_expenses['_merge'].value_counts())

    # Удаление лишних столбцов
    df_funnel_expenses = df_funnel_expenses.drop(columns=["_merge"], errors="ignore")

    # Форматирование даты
    df_funnel_expenses = format_date_column(df_funnel_expenses, 'Дата')

    # Вывод первых 5 строк
    print("Первые 5 строк таблицы ВоронкаСправочникЗатраты:")
    print(df_funnel_expenses.head())

    print("Таблица ВоронкаСправочникЗатраты успешно создана.")

except Exception as e:
    print(f"Ошибка при создании таблицы ВоронкаСправочникЗатраты: {e}")
del df_funnel_reference

# 15. Связать "ВоронкаСправочникЗатраты" с "Цены"
try:
    print("Начинаем создавать таблицу ВоронкаСправочникЗатратыЦены...")
    start_time = time.time()  # Запускаем таймер
    df_funnel_prices = pd.merge(df_funnel_expenses, df_price, on=["Дата", "Артикул"], how="left")
    df_funnel_prices = format_date_column(df_funnel_prices, 'Дата')

    # Вывод первых 5 строк
    print("Первые 5 строк таблицы ВоронкаСправочникЗатратыЦены:")
    print(df_funnel_prices.head())

    elapsed_time = time.time() - start_time  # Вычисляем затраченное время
    print(f"Таблица ВоронкаСправочникЗатратыЦены успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
except Exception as e:
    print(f"Ошибка при создании таблицы ВоронкаСправочникЗатратыЦены: {e}")
del df_funnel_expenses

# 9. Получить данные таблицы с SQL (РазмерыНаАгрегаторе)
try:
    print("Начинаем получать данные для РазмеровНаАгрегаторе...")
    start_time = time.time()  # Запускаем таймер
    query_sizes = f"""
        SELECT a.[dt] AS [Дата], a.[itemid] AS [Артикул], COUNT(DISTINCT(a.[INVENTSIZEID])) AS [Колво размеров]
        FROM [DBPartners].[dbo].[WblmRepGetStockWildberries] a
        WHERE [dt] >= '{(pd.Timestamp.today() - pd.DateOffset(months=3)).strftime("%Y-%m-%d")}'
        GROUP BY [dt], [itemid]
        ORDER BY [dt]
        DESC
    """
    df_sizes = pd.read_sql(query_sizes, Engine)
    df_sizes = format_date_column(df_sizes, 'Дата')
    # Вывод первых 5 строк
    print("Первые 5 строк таблицы РазмерыНаАгрегаторе:")
    print(df_sizes.head())

    elapsed_time = time.time() - start_time  # Вычисляем затраченное время
    print(f"Данные для РазмеровНаАгрегаторе успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
except Exception as e:
    print(f"Ошибка при получении данных для РазмеровНаАгрегаторе: {e}")

# 2. Получить данные таблицы с SQL (Остатки)
try:
    print("Начинаем получать данные для Остатков...")
    start_time = time.time()  # Запускаем таймер
    Engine = connect_to_sql(SQL_SERVER, SQL_DATABASE_DBPARTNERS)
    query_stock = f"""
        SELECT a.[dt] AS [Дата], a.[itemid] AS [Артикул], SUM(a.qte) AS [Остаток Агрегатора]
        FROM [DBPartners].[dbo].[WblmRepGetStockWildberries] a
        WHERE [dt] >= '{(pd.Timestamp.today() - pd.DateOffset(months=3)).strftime("%Y-%m-%d")}'
        GROUP BY [dt], [itemid]
        ORDER BY [dt]
        DESC
    """
    df_stock = pd.read_sql(query_stock, Engine)
    df_stock = format_date_column(df_stock, 'Дата')

    # Вывод первых 5 строк
    print("Первые 5 строк таблицы Остатков:")
    print(df_stock.head())

    elapsed_time = time.time() - start_time  # Вычисляем затраченное время
    print(f"Данные для Остатков успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
except Exception as e:
    print(f"Ошибка при получении данных для Остатков: {e}")

# 10. Создание таблицы "ВсегоРазмеров"
try:
    print("Начинаем создавать таблицу ВсегоРазмеров...")
    start_time = time.time()  # Запускаем таймер

    # Проверка наличия необходимых столбцов
    required_columns = ["Артикул", "Размер"]
    for col in required_columns:
        if col not in df_reference.columns:
            print(f"Ошибка: Отсутствует столбец '{col}' в df_reference.")
            exit()

    # Очищаем столбец "Размер":
    # - Преобразуем в строковый формат
    # - Удаляем лишние пробелы
    # - Заменяем пустые строки на None
    df_reference["Размер"] = df_reference["Размер"].astype(str).str.strip().replace('', None)

    # Создаем DataFrame с количеством размеров для каждого артикула
    df_reference_unique = (
        df_reference
        .drop_duplicates(subset=["Артикул", "Размер"])  # Удаляем дубликаты Артикул-Размер
        .groupby("Артикул")["Размер"]  # Группируем по артикулу
        .apply(lambda sizes: len(sizes.dropna().unique()) if len(sizes.dropna()) > 0 else 1)  # Подсчитываем размеры
        .reset_index(name="Всего размеров")
    )

    # Вывод первых 5 строк
    print("Первые 5 строк таблицы ВсегоРазмеров:")
    print(df_reference_unique.head())

    elapsed_time = time.time() - start_time  # Вычисляем затраченное время
    print(f"Таблица ВсегоРазмеров успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
except Exception as e:
    print(f"Ошибка при создании таблицы ВсегоРазмеров: {e}")

# 11. Связать "РазмерыНаАгрегаторе" с "ВсегоРазмеров"
try:
    print("Начинаем создавать таблицу Дистрибуция...")
    start_time = time.time()  # Запускаем таймер

    # Объединяем таблицы по полю "Артикул"
    df_distribution = pd.merge(df_sizes, df_reference_unique, on="Артикул", how="left")

    # Вычисляем дистрибуцию с проверкой на деление на ноль
    df_distribution["Дистрибуция"] = df_distribution.apply(
        lambda row: row["Колво размеров"] / row["Всего размеров"] if row["Всего размеров"] != 0 else 0,
        axis=1
    )

    # Оставляем только нужные столбцы
    df_distribution = df_distribution[["Дата", "Артикул", "Дистрибуция"]]

    # Форматирование даты
    df_distribution = format_date_column(df_distribution, 'Дата')

    # Вывод первых 5 строк
    print("Первые 5 строк таблицы Дистрибуция:")
    print(df_distribution.head())

    elapsed_time = time.time() - start_time  # Вычисляем затраченное время
    print(f"Таблица Дистрибуция успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
except Exception as e:
    print(f"Ошибка при создании таблицы Дистрибуция: {e}")
del df_reference_unique

# 12. Связать "Остатки" с "Дистрибуция"
try:
    print("Начинаем создавать таблицу Остатки с дистрибуцией...")
    start_time = time.time()  # Запускаем таймер
    df_stock_with_distribution = pd.merge(df_stock, df_distribution, on=["Дата", "Артикул"], how="left")
    df_stock_with_distribution = format_date_column(df_stock_with_distribution, 'Дата')
    # Вывод первых 5 строк
    print("Первые 5 строк таблицы Остатки с дистрибуцией:")
    print(df_stock_with_distribution.head())

    elapsed_time = time.time() - start_time  # Вычисляем затраченное время
    print(f"Таблица Остатки с дистрибуцией успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
except Exception as e:
    print(f"Ошибка при создании таблицы Остатки с дистрибуцией: {e}")
del df_distribution

# 7. Получить данные из файлов вложенной папки "Воронка с показами ВБ"
try:
    print("Начинаем получать данные для Воронки с показами ВБ...")
    start_time = time.time()  # Запускаем таймер
    folder_path_funnel_shows = os.path.join(FOLDER_PATH, "Воронка с показами ВБ")
    df_funnel_shows = pd.DataFrame()

    if os.path.exists(folder_path_funnel_shows):
        for file in os.listdir(folder_path_funnel_shows):
            file_path = os.path.join(folder_path_funnel_shows, file)

            # Пропускаем скрытые файлы
            if is_hidden(file_path):
                print(f"Пропущен скрытый файл: {file}")
                continue

            # Проверяем расширение файла
            if file.endswith((".xlsx", ".xls")):
                try:
                    # Чтение файла с явным преобразованием столбца "nm" в строку
                    if file.endswith(".xlsx"):
                        temp_df = pd.read_excel(
                            file_path,
                            sheet_name="Лист1",
                            engine="openpyxl",
                            usecols=range(6),
                            dtype={"nm": str}  # Явно задаем тип данных для столбца "nm"
                        )
                    elif file.endswith(".xls"):
                        temp_df = pd.read_excel(
                            file_path,
                            sheet_name="Лист1",
                            engine="xlrd",
                            usecols=range(6),
                            dtype={"nm": str}  # Явно задаем тип данных для столбца "nm"
                        )

                    # Переименование столбцов
                    temp_df.rename(columns={
                        "data_day": "Дата",
                        "nm": "Артикул WB",
                        "pokaz": "Показы из выгрузки",
                        "click": "Клики из выгрузки",
                        "basket": "Корзины из выгрузки",
                        "order": "Заказы из выгрузки"
                    }, inplace=True, errors="ignore")  # Игнорируем ошибки при переименовании

                    # Очистка данных в столбце "Артикул WB":
                    # - Удаляем лишние пробелы
                    # - Удаляем символы после точки (если есть)
                    temp_df["Артикул WB"] = temp_df["Артикул WB"].str.strip().str.split('.').str[0]

                    # Форматирование даты
                    temp_df = format_date_column(temp_df, 'Дата')

                    # Добавляем временный DataFrame к основному
                    df_funnel_shows = pd.concat([df_funnel_shows, temp_df])
                except Exception as e:
                    print(f"Ошибка при чтении файла {file}: {e}")

        # Вывод первых 5 строк
        print("Первые 5 строк таблицы Воронка с показами:")
        print(df_funnel_shows.head())

        elapsed_time = time.time() - start_time  # Вычисляем затраченное время
        print(f"Данные для Воронки с показами ВБ успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
    else:
        print("Папка 'Воронка с показами ВБ' не найдена.")
except Exception as e:
    print(f"Ошибка при получении данных для Воронки с показами ВБ: {e}")

# 8. Получить данные из файла !!!_Признаки для артикула и даты для ВБ
try:
    print("Начинаем получать данные для Признаков...")
    start_time = time.time()  # Запускаем таймер
    file_path_features = os.path.join(FOLDER_PATH_FEATURES, "!!!_Признаки для артикула и даты для ВБ.xlsx")
    if os.path.exists(file_path_features):
        df_item_features = pd.read_excel(file_path_features, sheet_name="Признаки для артикула", dtype=str, engine="openpyxl")
        df_date_features = pd.read_excel(file_path_features, sheet_name="Признаки для дат", dtype={0: "datetime64[ns]", **{i: str for i in range(1, 6)}}, engine="openpyxl")

        # Обработка ошибок
        df_item_features = handle_errors(df_item_features)
        df_date_features = handle_errors(df_date_features)

        # Форматирование даты
        df_date_features = format_date_column(df_date_features, 'Дата')

        # Вывод первых 5 строк
        print("Первые 5 строк таблицы Признаки артикула:")
        print(df_item_features.head())

        # Вывод первых 5 строк
        print("Первые 5 строк таблицы Признаки дат:")
        print(df_date_features.head())

        elapsed_time = time.time() - start_time  # Вычисляем затраченное время
        print(f"Данные для Признаков успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
    else:
        print("Файл '!!!_Признаки для артикула и даты для ВБ.xlsx' не найден.")
except Exception as e:
    print(f"Ошибка при получении данных для Признаков: {e}")

# 16. Связать "ВоронкаСправочникЗатратыЦены" с "Остатки с дистрибуцией"
try:
    print("Начинаем создавать таблицу ДБбезПризнаков...")
    start_time = time.time()  # Запускаем таймер
    df_final_db = pd.merge(df_funnel_prices, df_stock_with_distribution, on=["Дата", "Артикул"], how="left")
    df_final_db = format_date_column(df_final_db, 'Дата')
    # Вывод первых 5 строк
    print("Первые 5 строк таблицы ДБбезПризнаков:")
    print(df_final_db.head())

    elapsed_time = time.time() - start_time  # Вычисляем затраченное время
    print(f"Таблица ДБбезПризнаков успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
except Exception as e:
    print(f"Ошибка при создании таблицы ДБбезПризнаков: {e}")
del df_funnel_prices
# 16.5 Связать "ДБбезПризнаков" с "Воронка с показами ВБ"
try:
    print("Начинаем создавать таблицу ДБсПоказами...")
    start_time = time.time()  # Запускаем таймер

    # Проверяем наличие необходимых столбцов в df_funnel_shows
    required_columns = ["Дата", "Артикул WB"]
    for col in required_columns:
        if col not in df_funnel_shows.columns:
            print(f"Ошибка: Отсутствует столбец '{col}' в df_funnel_shows.")
            exit()

    # Приводим типы данных к строковому формату
    df_final_db["Артикул WB"] = (df_final_db["Артикул WB"]
        .fillna('')
        .astype(str)
        .str.strip()
        .str.upper()
    )
    df_funnel_shows["Артикул WB"] = (df_funnel_shows["Артикул WB"]
        .fillna('')
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # Объединяем таблицы по полям "Дата" и "Артикул WB"
    df_final_db_with_shows = pd.merge(
        df_final_db,
        df_funnel_shows,
        left_on=["Дата", "Артикул WB"],
        right_on=["Дата", "Артикул WB"],
        how="left"
    )

    # Переименовываем столбец "Дата" (если нужно)
    #df_final_db_with_shows = df_final_db_with_shows.drop(columns=["Дата"], errors="ignore")

    # Форматирование даты
    df_final_db_with_shows = format_date_column(df_final_db_with_shows, 'Дата')

    # Вывод первых 5 строк
    print("Первые 5 строк таблицы ДБсПоказами:")
    print(df_final_db_with_shows.head())

    elapsed_time = time.time() - start_time  # Вычисляем затраченное время
    print(f"Таблица ДБсПоказами успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
except Exception as e:
    print(f"Ошибка при создании таблицы ДБсПоказами: {e}")

# 17. Связать "ДБсПоказами" с "Признаки для артикула"
try:
    print("Начинаем создавать таблицу ДБсПризнакамиАртикула...")
    start_time = time.time()  # Запускаем таймер
    df_final_db_item_features = pd.merge(df_final_db_with_shows, df_item_features, on="Артикул", how="left")
    df_final_db_item_features = format_date_column(df_final_db_item_features, 'Дата')

    # Вывод первых 5 строк
    print("Первые 5 строк таблицы ДБсПризнакамиАртикула:")
    print(df_final_db_item_features.head())

    elapsed_time = time.time() - start_time  # Вычисляем затраченное время
    print(f"Таблица ДБсПризнакамиАртикула успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
except Exception as e:
    print(f"Ошибка при создании таблицы ДБсПризнакамиАртикула: {e}")
del df_final_db_with_shows
# 18. Связать "ДБсПризнакамиАртикула" с "Признаки для дат"
try:
    print("Начинаем создавать таблицу ДБсПризнаками...")
    start_time = time.time()  # Запускаем таймер
    df_final_db_all_features = pd.merge(df_final_db_item_features, df_date_features, on="Дата", how="left")
    df_final_db_all_features = format_date_column(df_final_db_all_features, 'Дата')


    # Вывод первых 5 строк
    print("Первые 5 строк таблицы ДБсПризнаками:")
    print(df_final_db_all_features.head())

    df_final_db_all_features.to_csv(os.path.join(FOLDER_PATH, "ДБсПризнаками.csv"), index=False)
    elapsed_time = time.time() - start_time  # Вычисляем затраченное время
    print(f"Таблица ДБсПризнаками успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
except Exception as e:
    print(f"Ошибка при создании таблицы ДБсПризнаками: {e}")
del df_final_db_item_features

# Функция для обновления Excel-файла с циклом попыток
def update_and_save_excel(file_path, new_file_path):
    max_attempts = 10  # Максимальное количество попыток
    attempt = 0

    while attempt < max_attempts:
        attempt += 1
        print(f"Попытка {attempt} обновить файл '{os.path.basename(file_path)}'...")

        try:
            # Открываем Excel приложение
            excel = win32.Dispatch("Excel.Application")
            excel.DisplayAlerts = False  # Отключает предупреждения Excel

            try:
                # Открываем книгу
                workbook = excel.Workbooks.Open(file_path)

                # Выполняем обновление всех данных (эквивалентно "Обновить всё" в Excel)
                print("Выполняем обновление данных...")
                workbook.RefreshAll()
                excel.CalculateUntilAsyncQueriesDone()  # Дожидаемся завершения обновления

                # Сохраняем оригинальный файл в FOLDER_PATH_FOR_DB
                workbook.SaveAs(file_path)
                print(f"Файл успешно сохранен с оригинальным именем в '{os.path.dirname(file_path)}'.")

                # Сохраняем файл с новым именем в FOLDER_PATH_FEATURES
                workbook.SaveAs(new_file_path)
                print(f"Файл успешно сохранен как '{os.path.basename(new_file_path)}'.")

                return True  # Успешное завершение

            except Exception as e:
                print(f"Ошибка при обновлении или сохранении файла: {e}")
            finally:
                # Закрываем книгу и выходим из Excel
                if 'workbook' in locals():
                    workbook.Close(SaveChanges=False)
                excel.Quit()

        except Exception as e:
            print(f"Ошибка при работе с Excel: {e}")

        # Если произошла ошибка, ждем перед следующей попыткой
        if attempt < max_attempts:
            print(f"Пауза перед следующей попыткой ({attempt + 1}/{max_attempts})...")
            time.sleep(60)  # Пауза 5 секунд

    return False  # Все попытки завершились неудачно

# 19. Обновить файл "Показы и затраты ОЗ_2.0.xlsx"
try:
    print("Подготовка данных для ДБ завершена.")
    input("Начать обновление файлов ДБ? Для подтверждения нажмите Enter...")
    print("Начинаем обновлять файл 'Показы и затраты ВБ_2.0.xlsx'...")
    start_time = time.time()  # Запускаем таймер

    # Путь к исходному файлу
    file_path_shows_expenses = os.path.join(FOLDER_PATH_FOR_DB, "Показы и затраты ВБ_2.0.xlsx")

    if os.path.exists(file_path_shows_expenses):
        # Создаем новое имя файла с текущей датой без года
        current_month_day = time.strftime("%d.%m")  # Текущая дата в формате ДД.ММ
        new_file_name = f"Показы и затраты ВБ_2.0 {current_month_day}.xlsx"
        new_file_path = os.path.join(FOLDER_PATH_FEATURES, new_file_name)

        # Путь для сохранения в дополнительную папку FOLDER_PATH_DUDL
        dudl_file_path = os.path.join(FOLDER_PATH_DUDL, new_file_name)

        # Удаляем старые файлы из FOLDER_PATH_DUDL
        try:
            if os.path.exists(FOLDER_PATH_DUDL):
                for filename in os.listdir(FOLDER_PATH_DUDL):
                    # Ищем файлы с шаблоном "Показы и затраты ОЗ_2.0 DD.MM.xlsx"
                    match = re.match(r"Показы и затраты ВБ_2\.0 (\d{2}\.\d{2})\.xlsx", filename)
                    if match:
                        file_date = match.group(1)  # Извлекаем дату из имени файла
                        if file_date != current_month_day:  # Сравниваем с текущей датой
                            file_to_delete = os.path.join(FOLDER_PATH_DUDL, filename)
                            os.remove(file_to_delete)
                            print(f"Файл '{filename}' удален из папки '{FOLDER_PATH_DUDL}'.")
        except Exception as delete_error:
            print(f"Ошибка при удалении старых файлов из папки '{FOLDER_PATH_DUDL}': {delete_error}")

        # Удаляем старые файлы из FOLDER_PATH_FEATURES
        try:
            if os.path.exists(FOLDER_PATH_FEATURES):
                for filename in os.listdir(FOLDER_PATH_FEATURES):
                    # Ищем файлы с шаблоном "Показы и затраты ОЗ_2.0 DD.MM.xlsx"
                    match = re.match(r"Показы и затраты ВБ_2\.0 (\d{2}\.\d{2})\.xlsx", filename)
                    if match:
                        file_date = match.group(1)  # Извлекаем дату из имени файла
                        if file_date != current_month_day:  # Сравниваем с текущей датой
                            file_to_delete = os.path.join(FOLDER_PATH_FEATURES, filename)
                            os.remove(file_to_delete)
                            print(f"Файл '{filename}' удален из папки '{FOLDER_PATH_FEATURES}'.")
        except Exception as delete_error:
            print(f"Ошибка при удалении старых файлов из папки '{FOLDER_PATH_FEATURES}': {delete_error}")

        # Пытаемся обновить и сохранить файл
        success = update_and_save_excel(file_path_shows_expenses, new_file_path)

        if not success:
            # Если все попытки неудачны, выводим сообщение пользователю
            while not success:
                input("Обновить Excel файл не получилось. Закройте все открытые файлы и нажмите любую кнопку для повторной попытки.")
                success = update_and_save_excel(file_path_shows_expenses, new_file_path)

            print("Файл успешно обновлен после повторной попытки.")

        # После успешного обновления копируем файл в папку FOLDER_PATH_DUDL
        if success:
            try:
                shutil.copy(new_file_path, dudl_file_path)
                print(f"Файл успешно скопирован в папку '{FOLDER_PATH_DUDL}'.")
            except Exception as copy_error:
                print(f"Ошибка при копировании файла в папку '{FOLDER_PATH_DUDL}': {copy_error}")

        elapsed_time = time.time() - start_time  # Вычисляем затраченное время
        print(f"Файл успешно обновлен и сохранен. Время выполнения: {format_elapsed_time(elapsed_time)}")
    else:
        print("Файл 'Показы и затраты ВБ_2.0.xlsx' не найден.")
except Exception as e:
    print(f"Ошибка при обработке файла 'Показы и затраты ВБ_2.0.xlsx': {e}")
elapsed_all_time = time.time() - start_all_time  # Вычисляем затраченное время
print(f"Файл успешно обновлен и сохранен. Время выполнения: {format_elapsed_time(elapsed_time)}")