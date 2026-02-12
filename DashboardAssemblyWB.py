import os
import pandas as pd
from sqlalchemy import create_engine
import pyodbc
import win32com.client as win32
import time  # Для измерения времени выполнения
import shutil
import re
from functools import reduce
import numpy as np

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

FOLDER_PATH = os.path.normpath(r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!")
FOLDER_PATH_FEATURES = r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Дашбоард по рекламным кампаниям"
FOLDER_PATH_FOR_DB = os.path.normpath(r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям")
FOLDER_PATH_DUDL = os.path.normpath(r"\\kari.local\public\all\Агрегаторы\Дашборд реклама WB_OZ")

SQL_SERVER = "cl01sql"
SQL_DATABASE_DBREPORT = "DBReport"
SQL_DATABASE_DBPARTNERS = "DBPartners"

# 1. Получить данные из файла "Затраты ВБ_2.xlsx"
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
        df_expenses = pd.read_excel(file_path_expenses, sheet_name="Затраты ВБ", engine="calamine",
                                    usecols=columns_to_read)

        # Переименование столбцо
        df_expenses.rename(columns=column_mapping, inplace=True)

        # Форматирование даты
        df_expenses = format_date_column(df_expenses, 'Дата')
        
        # df_expenses = df_expenses.drop_duplicates(subset=['Артикул WB', 'Дата'], keep='first')
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


# 2. Получить данные из файла "История-затрат-Все.xlsx"
pattern = r"История-затрат-Все*\.xlsx"

folder = os.path.join(FOLDER_PATH, "Затраты", "Затраты ВБ")
pattern = re.compile(r"^История-затрат-Все.*\.xlsx$", re.IGNORECASE)

files = [e.name for e in os.scandir(folder) if e.is_file() and pattern.match(e.name)]
# Если нужны полные пути:
# files = [e.path for e in os.scandir(folder) if e.is_file() and pattern.match(e.name)]

# Если файлы не найдены, выбрасываем ошибку
if not files:
    raise FileNotFoundError("Файл История-затрат не найден")
# Находим последний файл по времени изменения
latest = max(files, key=lambda f: os.path.getmtime(os.path.join(os.path.join(FOLDER_PATH, "Затраты/Затраты ВБ"), f)))

# Путь к последнему найденному файлу
latest_file_path = os.path.join(os.path.join(FOLDER_PATH, "Затраты/Затраты ВБ"), latest)

df_costs_hist = pd.read_excel(latest_file_path, engine='calamine')
df_costs_hist = df_costs_hist[['Кампания', 'ID кампании', 'Раздел']]
df_costs_hist['Раздел'] = df_costs_hist['Раздел'].astype(str)
type_dictionary = {"Единая Ставка":"Автоматическое", 
                   "Ручная Ставка":"Аукцион"}

df_costs_hist["Раздел"] = df_costs_hist["Раздел"].map(type_dictionary)
df_costs_hist

# 3. Создание широкой таблицы по типам активности
# Маппинг «сырых» значений разделов -> нормализованные
TYPE_DICT_RAW = {
    "Единая Ставка": "Автоматическое",
    "Ручная Ставка": "Аукцион"
}

# Приоритет при конфликте (к одному ID привязаны разные разделы)
SECTION_PRIORITY = {'Аукцион': 2, 'Автоматическое': 1, np.nan: 0}

# Как называть «неопределённый» раздел
MISC_SECTION_NAME = 'Прочее'

# Кандидаты для названия столбца расхода в df_expenses
SPEND_COL_CANDIDATES = ['Расход, ₽', 'Расход, руб', 'Расход, Р', 'Расход']

# Ключи в wide-таблице
KEY_COLS = ['Дата', 'Артикул WB']


# ===================== ВСПОМОГАТЕЛЬНЫЕ =====================

def _pick_spend_col(df: pd.DataFrame) -> str:
    for c in SPEND_COL_CANDIDATES:
        if c in df.columns:
            return c
    raise KeyError("Не найдена колонка расхода среди: " + ", ".join(SPEND_COL_CANDIDATES))


def build_camp_ref(df_costs_hist: pd.DataFrame) -> pd.DataFrame:
    """
    На вход: df_costs_hist с колонками как минимум ['ID кампании', 'Раздел'].
    Возвращает camp_ref: по одному 'Раздел' на каждый 'ID кампании'.
    - Нормализует 'Раздел' через TYPE_DICT_RAW (прочее -> NaN)
    - При нескольких значениях для одного ID выбирает по SECTION_PRIORITY
    """
    if df_costs_hist.empty:
        return pd.DataFrame({'ID кампании': pd.Series(dtype='Int64'),
                             'Раздел': pd.Series(dtype='object')})

    dfp = df_costs_hist[['ID кампании', 'Раздел']].copy()
    dfp['ID кампании'] = pd.to_numeric(dfp['ID кампании'], errors='coerce').astype('Int64')

    # Нормализация разделов (если уже нормализованы, невидимые значения станут NaN — это ок)
    dfp['Раздел'] = dfp['Раздел'].map(TYPE_DICT_RAW).where(dfp['Раздел'].map(TYPE_DICT_RAW).notna(), dfp['Раздел'])

    # Приоритет и выбор одного значения на ID
    dfp['_prio'] = dfp['Раздел'].map(SECTION_PRIORITY).fillna(0).astype(int)
    dfp = dfp.sort_values(['ID кампании', '_prio'], ascending=[True, False])
    camp_ref = dfp.drop_duplicates(subset=['ID кампании'], keep='first')[['ID кампании', 'Раздел']].copy()

    # sanity: уникальность справа
    assert not camp_ref['ID кампании'].duplicated().any(), "camp_ref содержит дубликаты ID кампании"
    return camp_ref


def merge_expenses_with_sections(df_expenses: pd.DataFrame, camp_ref: pd.DataFrame) -> pd.DataFrame:
    """
    Левый merge строго m:1. Исключает размножение строк.
    """
    out = df_expenses.copy()
    out['ID кампании'] = pd.to_numeric(out['ID кампании'], errors='coerce').astype('Int64')

    # Строгая валидация 'm:1' – упадёт, если camp_ref не уникален по ID
    out = out.merge(camp_ref, on='ID кампании', how='left', validate='m:1')
    out['Раздел'] = out['Раздел'].fillna(MISC_SECTION_NAME)
    return out


import numpy as np
import pandas as pd

# ===================== ПАРАМЕТРЫ =====================

# Если в справочнике встречаются "Единая Ставка"/"Ручная Ставка" — нормализуем
TYPE_DICT_RAW = {
    "Единая Ставка": "Автоматическое",
    "Ручная Ставка": "Аукцион"
}

# При конфликте нескольких "Раздел" на один ID — выбираем по приоритету
SECTION_PRIORITY = {'Аукцион': 2, 'Автоматическое': 1, np.nan: 0}

# Как называть неопределённый раздел (после merge)
MISC_SECTION_NAME = 'Прочее'

# Кандидаты названия столбца расхода в df_expenses
SPEND_COL_CANDIDATES = ['Расход, ₽', 'Расход, руб', 'Расход, Р', 'Расход']

# Ключи формирования wide
KEY_COLS = ['Дата', 'Артикул WB']

# Разделы, которые считаем платными (для логической "Органики")
PAID_SECTIONS = ['Автоматическое', 'Аукцион']


# ===================== ВСПОМОГАТЕЛЬНЫЕ =====================

def _pick_spend_col(df: pd.DataFrame) -> str:
    for c in SPEND_COL_CANDIDATES:
        if c in df.columns:
            return c
    raise KeyError("Не найдена колонка расхода среди: " + ", ".join(SPEND_COL_CANDIDATES))


def build_camp_ref(df_costs_hist: pd.DataFrame) -> pd.DataFrame:
    """
    На вход: df_costs_hist с колонками как минимум ['ID кампании','Раздел'].
    Возвращает camp_ref: по одному 'Раздел' на каждый 'ID кампании' (по приоритету).
    """
    if df_costs_hist.empty:
        return pd.DataFrame({'ID кампании': pd.Series(dtype='Int64'),
                             'Раздел': pd.Series(dtype='object')})

    dfp = df_costs_hist[['ID кампании', 'Раздел']].copy()
    dfp['ID кампании'] = pd.to_numeric(dfp['ID кампании'], errors='coerce').astype('Int64')

    # Нормализация: заменяем только известные сырьевые значения, остальное оставляем как есть
    dfp['Раздел'] = dfp['Раздел'].replace(TYPE_DICT_RAW)

    # Приоритет и выбор одного значения на ID
    dfp['_prio'] = dfp['Раздел'].map(SECTION_PRIORITY).fillna(0).astype(int)
    dfp = dfp.sort_values(['ID кампании', '_prio'], ascending=[True, False])
    camp_ref = dfp.drop_duplicates(subset=['ID кампании'], keep='first')[['ID кампании', 'Раздел']].copy()

    # sanity: справа строго 1 строка на ID
    assert not camp_ref['ID кампании'].duplicated().any(), "camp_ref содержит дубликаты ID кампании"
    return camp_ref


def merge_expenses_with_sections(df_expenses: pd.DataFrame, camp_ref: pd.DataFrame) -> pd.DataFrame:
    """
    Левый merge строго m:1. Исключает размножение строк (и, как следствие, раздутые суммы).
    """
    out = df_expenses.copy()
    out['ID кампании'] = pd.to_numeric(out['ID кампании'], errors='coerce').astype('Int64')
    out = out.merge(camp_ref, on='ID кампании', how='left', validate='m:1')
    out['Раздел'] = out['Раздел'].fillna(MISC_SECTION_NAME)
    return out


def build_wide_by_section(df_merged: pd.DataFrame, key_cols=KEY_COLS) -> pd.DataFrame:
    """
    Формирует wide-таблицу:
      - <Раздел>_<Метрика> по каждой метрике
      - Итоги (без префиксов) как сумма по всем разделам
      - ID по разделам, 'Тип активности'
      - Бинарные флаги по разделам + логическая 'Органика'
    """
    df = df_merged.copy()
    key_cols = list(key_cols)
    aux_cols = ['ID кампании', 'Раздел']
    metric_cols = [c for c in df.columns if c not in set(key_cols + aux_cols)]

    # Приведение метрик к числам
    for c in metric_cols:
        df[c] = pd.to_numeric(df[c], errors='coerce')

    # 1) Суммы по (Дата, Артикул WB, Раздел)
    agg = (df.groupby(key_cols + ['Раздел'], as_index=False)[metric_cols]
             .sum(min_count=1))

    # 2) Pivot → (метрика, раздел) → Раздел_Метрика
    wide = agg.pivot_table(index=key_cols, columns='Раздел', values=metric_cols,
                           aggfunc='sum', fill_value=0)
    wide.columns = [f"{sec}_{met}" for met, sec in wide.columns.to_flat_index()]
    wide = wide.reset_index()

    # Актуальные разделы
    sections = list(pd.unique(df['Раздел'].dropna()))

    # 3) Итоги без префиксов (ОДИН раз по каждой метрике)
    for met in metric_cols:
        cols_to_sum = [f"{sec}_{met}" for sec in sections if f"{sec}_{met}" in wide.columns]
        if cols_to_sum:
            wide[met] = wide[cols_to_sum].sum(axis=1)

    # 4) ID по разделам (первый попавшийся в группе)
    id_wide = (df.groupby(key_cols + ['Раздел'], as_index=False)['ID кампании']
                 .first()
                 .pivot(index=key_cols, columns='Раздел', values='ID кампании')
                 .rename(columns=lambda sec: f"{sec}_ID кампании")
                 .reset_index())
    wide = wide.merge(id_wide, on=key_cols, how='left')

    # 5) Тип активности (по наличию ID в платных разделах)
    presence = {}
    for sec in PAID_SECTIONS:
        col = f'{sec}_ID кампании'
        presence[sec] = wide[col].notna() if col in wide.columns else pd.Series(False, index=wide.index)

    if len(PAID_SECTIONS) >= 2:
        both = presence[PAID_SECTIONS[0]] & presence[PAID_SECTIONS[1]]
        only_first = presence[PAID_SECTIONS[0]] & ~presence[PAID_SECTIONS[1]]
        only_second = ~presence[PAID_SECTIONS[0]] & presence[PAID_SECTIONS[1]]
        wide['Тип активности'] = np.select(
            [both, only_first, only_second],
            [f'{PAID_SECTIONS[0]} и {PAID_SECTIONS[1]}', PAID_SECTIONS[0], PAID_SECTIONS[1]],
            default='Органика'
        )
    else:
        # если платный только один тип
        only = presence[PAID_SECTIONS[0]]
        wide['Тип активности'] = np.where(only, PAID_SECTIONS[0], 'Органика')

    # 6) БИНАРНЫЕ ФЛАГИ ПО РАЗДЕЛАМ (one-hot) + ЛОГИЧЕСКАЯ "ОРГАНИКА"
    flags_block = (
        df.assign(_hit=1)
          .pivot_table(index=key_cols, columns='Раздел', values='_hit',
                       aggfunc='max', fill_value=0)
          .astype('int8')
          .reset_index()
    )
    wide = wide.merge(flags_block, on=key_cols, how='left')

    # Гарантируем наличие флагов для платных разделов
    for sec in PAID_SECTIONS:
        if sec not in wide.columns:
            wide[sec] = 0

    # Логическая "Органика" = 1, если ни один платный раздел не активен
    wide['Органика'] = (1 - wide[PAID_SECTIONS].max(axis=1)).astype('int8')

    # 7) Порядок колонок
    section_order = PAID_SECTIONS + [s for s in sections if s not in PAID_SECTIONS]
    ordered = key_cols + ['Тип активности']

    # Бинарные флаги сначала: платные, затем прочие, затем "Органика"
    ordered += [s for s in PAID_SECTIONS if s in wide.columns]
    ordered += [s for s in sections if (s not in PAID_SECTIONS and s in wide.columns and s != 'Органика')]
    ordered += ['Органика'] if 'Органика' in wide.columns else []

    # ID и метрики по разделам → итоги
    for sec in section_order:
        id_col = f"{sec}_ID кампании"
        if id_col in wide.columns:
            ordered.append(id_col)
        ordered += [c for c in wide.columns if c.startswith(f"{sec}_") and c != id_col]
    ordered += [c for c in metric_cols if c in wide.columns]

    wide = wide[[c for c in ordered if c in wide.columns]].copy()
    return wide

def build_sum_table(df_expenses: pd.DataFrame, df_costs_hist: pd.DataFrame) -> pd.DataFrame:
    """
    End-to-end: camp_ref (one-row-per-ID) -> merge (m:1) -> wide (+флаги).
    Печатает контроль суммы расходов до/после.
    """
    spend_col = _pick_spend_col(df_expenses)
    base_sum = pd.to_numeric(df_expenses[spend_col], errors='coerce').sum()

    camp_ref = build_camp_ref(df_costs_hist)
    df_merged = merge_expenses_with_sections(df_expenses, camp_ref)
    df_sum = build_wide_by_section(df_merged, key_cols=KEY_COLS)

    if spend_col in df_sum.columns:
        after_sum = pd.to_numeric(df_sum[spend_col], errors='coerce').sum()
        print(f"[CHECK] Сумма расходов до/после: {base_sum:,.0f} → {after_sum:,.0f}")
        if not np.isclose(after_sum, base_sum, rtol=1e-3, atol=1e-2):
            print("[WARN] Расхождение суммы расходов. Проверьте конфликтные ID кампаний или строки без ID.")
    else:
        print(f"[INFO] В df_sum нет '{spend_col}' для проверки суммы.")

    return df_sum



def build_sum_table(df_expenses: pd.DataFrame, df_costs_hist: pd.DataFrame) -> pd.DataFrame:
    """
    End-to-end: camp_ref (one-row-per-ID) -> merge (m:1) -> wide.
    Печатает контроль суммы расходов до/после.
    """
    # Базовая сумма расходов до пайплайна
    spend_col = _pick_spend_col(df_expenses)
    base_sum = pd.to_numeric(df_expenses[spend_col], errors='coerce').sum()

    # camp_ref и merge
    camp_ref = build_camp_ref(df_costs_hist)
    df_merged = merge_expenses_with_sections(df_expenses, camp_ref)

    # wide
    df_sum = build_wide_by_section(df_merged, key_cols=KEY_COLS)

    # Контроль суммы расходов
    if spend_col in df_sum.columns:
        after_sum = pd.to_numeric(df_sum[spend_col], errors='coerce').sum()
        print(f"[CHECK] Сумма расходов до/после: {base_sum:,.0f} → {after_sum:,.0f}")
        if not np.isclose(after_sum, base_sum, rtol=1e-3, atol=1e-2):
            print("[WARN] Расхождение суммы расходов. Диагностика конфликтных ID:")
            bad = (df_costs_hist.groupby('ID кампании')['Раздел']
                   .nunique().sort_values(ascending=False))
            bad = bad[bad > 1]
            if not bad.empty:
                print("ID с несколькими значениями 'Раздел':")
                print(bad.head(20))
            else:
                null_id_sum = pd.to_numeric(df_expenses.loc[df_expenses['ID кампании'].isna(), spend_col], errors='coerce').sum()
                print(f"Сумма по строкам без ID кампании (ушли в '{MISC_SECTION_NAME}'): {null_id_sum:,.0f}")
    else:
        print(f"[INFO] В df_sum нет колонки '{spend_col}' для проверки суммы.")

    return df_sum

wide = build_sum_table(df_expenses=df_expenses, df_costs_hist=df_costs_hist)
wide

df_expenses = wide

# 4. Получить данные таблицы с SQL (Цены)
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



# 5. Получить данные из файла "Справочник.xlsx"
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



# 6. Получить данные из файлов вложенной папки "Показатели по неделям ВБ"
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
                        "Номенклатура", "Рейтинг карточки", "Показы", "Переходы в карточку",
                        "Положили в корзину", "Заказали, шт", "Выкупили, шт",
                        "Отменили, шт", "Заказали на сумму, руб", "Выкупили на сумму, руб",
                        "Отменили на сумму, руб", "Средняя цена, руб", "Дата", "Рейтинг по отзывам"
                    ]

                    if file.endswith(".xlsx"):
                        temp_df = pd.read_excel(file_path, sheet_name="воронка ОЗОН", engine="calamine")
                    elif file.endswith(".xls"):
                        temp_df = pd.read_excel(file_path, sheet_name="воронка ОЗОН", engine="xlrd")

                    # Добавляем отсутствующие столбцы со значением None
                    for col in all_columns:
                        if col not in temp_df.columns:
                            # если числовой столбец — проставляем 0, иначе None
                            if col in ["Показы", "Показы на карточке товара", "Положили в корзину", 
                                    "Заказали, шт", "Выкупили, шт", "Отменили, шт"]:
                                temp_df[col] = 0
                            else:
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
                        "Показы": int, 
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
            "Дата", "Артикул WB", "Рейтинг карточки", "Показы", "Показы на карточке товара",
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



# 7. Связать "Воронка" с "Затраты ВБ"

# ----------- настройки -----------
KEY_COLS = ['Дата', 'Артикул WB']
SPEND_CANDS = ['Расход, руб', 'Автоматическое_Расход, руб', 'Аукцион_Расход, руб']  # названия возможного столбца расхода
ALLOC_MODE = 'even'   # 'first' — записать расход в первую строку группы; 'even' — равномерно разделить
NO_SUM_CONTAINS = ['ID кампании']  # числовые техполя, которые НЕЛЬЗЯ суммировать при агрегации справа

# ----------- утилиты -----------
def _pick_spend_col(df: pd.DataFrame) -> str:
    for c in SPEND_CANDS:
        if c in df.columns:
            return c
    raise KeyError(f"Не найден столбец расхода среди: {SPEND_CANDS}")

def _clean_money_series(s: pd.Series) -> pd.Series:
    # убираем пробелы/неразрывные и заменяем запятую на точку
    if s.dtype.kind in 'biufc':  # уже число
        return s
    s = (s.astype(str)
           .str.replace('\u00A0', '', regex=False)  # NBSP
           .str.replace(' ', '', regex=False)
           .str.replace(',', '.', regex=False))
    # убираем всё, что не цифра/точка/минус
    s = s.str.replace(r'[^0-9\.\-]+', '', regex=True)
    return pd.to_numeric(s, errors='coerce')

def _norm_keys(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    # дата: сначала пытаемся dayfirst, затем обычный парс; оставляем дату без времени
    # d1 = pd.to_datetime(out['Дата'], errors='coerce', dayfirst=True)
    # d2 = pd.to_datetime(out['Дата'], errors='coerce', dayfirst=False)
    # out['Дата'] = d1.fillna(d2).dt.normalize()
    # артикул
    out['Артикул WB'] = (out['Артикул WB']
                         .fillna('').astype(str).str.strip().str.upper())
    return out

def _make_right_unique(df_expenses: pd.DataFrame) -> pd.DataFrame:
    """
    Агрегируем df_expenses до 1 строки на (Дата, Артикул WB):
    - числовые метрики суммируем (кроме тех, чьё имя содержит токены из NO_SUM_CONTAINS);
    - бинарные флаги ('Автоматическое','Аукцион','Органика') берём max;
    - прочие текстовые — first.
    """
    df = _norm_keys(df_expenses).copy()

    # приведение расхода к числу и выбор его столбца
    spend_col = _pick_spend_col(df)
    df[spend_col] = _clean_money_series(df[spend_col])

    key = KEY_COLS
    other = [c for c in df.columns if c not in key]

    flag_cols = [c for c in ['Автоматическое', 'Аукцион', 'Органика'] if c in df.columns]
    no_sum_cols = [c for c in other if any(tok in c for tok in NO_SUM_CONTAINS)]

    # что считать «числом» для суммирования
    num_cols = []
    for c in other:
        if c in flag_cols or c in no_sum_cols:
            continue
        # аккуратно приводим потенциально числовые
        df[c] = pd.to_numeric(df[c], errors='ignore')
        if pd.api.types.is_numeric_dtype(df[c]):
            num_cols.append(c)

    text_cols = [c for c in other if c not in num_cols and c not in flag_cols and c not in no_sum_cols]

    agg = {c: 'sum' for c in num_cols}
    agg.update({c: 'max' for c in flag_cols})
    agg.update({c: 'first' for c in text_cols})
    agg.update({c: 'first' for c in no_sum_cols})  # ID и подобные — берем первое

    right_u = (df.groupby(key, as_index=False).agg(agg))

    # sanity
    assert right_u.duplicated(key).sum() == 0, "Правая часть после агрегации не уникальна по ключу"
    return right_u

def _allocate_on_left_dupes(df_merged: pd.DataFrame, spend_col: str, mode: str) -> pd.DataFrame:
    """
    Корректируем 'Расход' только в тех группах (Дата, Артикул WB), где после OUTER-merge появилось
    >1 строки из левой таблицы, чтобы суммы по датам/итогу не искажались.
    """
    out = df_merged.copy()
    out[spend_col] = pd.to_numeric(out[spend_col], errors='coerce').fillna(0.0)

    # считаем размер группы по ключу
    counts = out.groupby(KEY_COLS, dropna=False)[spend_col].transform('size')

    if mode == 'first':
        # оставляем расход только в первой строке группы; прочие — 0
        mask_first = ~out.duplicated(subset=KEY_COLS, keep='first')
        out.loc[~mask_first, spend_col] = 0.0
    elif mode == 'even':
        out[spend_col] = out[spend_col] / counts.clip(lower=1)
    else:
        raise ValueError("ALLOC_MODE должен быть 'first' или 'even'.")

    return out

# ----------- основной блок -----------
def build_funnel_expenses_outer(df_funnel: pd.DataFrame,
                                df_expenses: pd.DataFrame,
                                alloc_mode: str = ALLOC_MODE) -> pd.DataFrame:
    """
    OUTER-merge: сохраняет ВСЮ сумму 'Расход, руб' по датам и в целом (как в df_expenses),
    исключает раздувание на дубликатах слева.
    """
    # 0) нормализация
    left = _norm_keys(df_funnel)
    right_u = _make_right_unique(df_expenses)     # 1 строка на ключ справа

    spend_col = _pick_spend_col(right_u)
    base_by_date = (right_u.groupby('Дата', as_index=False)[spend_col]
                          .sum(min_count=1).rename(columns={spend_col: 'Расход_до'}))
    base_total = right_u[spend_col].sum()

    # 1) OUTER-merge: чтобы не потерять расходы по ключам, которых нет в воронке
    df = left.merge(right_u, on=KEY_COLS, how='outer', indicator=True)

    # 2) корректируем расход только где есть дубликаты по ключу (т.е. несколько левых строк на один ключ)
    #    Для правых «одиноких» ключей (которых не было в воронке) размер группы = 1 — значение не меняется.
    df = _allocate_on_left_dupes(df, spend_col, alloc_mode)

    # 3) проверки инвариантов по датам и общая сумма
    after_by_date = (df.groupby('Дата', as_index=False)[spend_col]
                       .sum(min_count=1).rename(columns={spend_col: 'Расход_после'}))
    chk = base_by_date.merge(after_by_date, on='Дата', how='outer').fillna(0)
    drift = chk.loc[~np.isclose(chk['Расход_до'], chk['Расход_после'], rtol=1e-9, atol=1e-6)]
    after_total = df[spend_col].sum()

    if not drift.empty or not np.isclose(after_total, base_total, rtol=1e-9, atol=1e-6):
        # покажем топ расхождений для быстрой отладки
        print("[WARN] Найдены расхождения после объединения (первые 10 дат):")
        print(drift.sort_values('Расход_до', ascending=False).head(10))
        print(f"[WARN] Общая сумма: до={base_total:,.2f} | после={after_total:,.2f}")
        # не падаем, а выводим диагностику — чтобы вы видели конкретные даты

    # 4) (опционально) можно убрать служебную колонку индикатора
    df.drop(columns=['_merge'], inplace=True, errors='ignore')

    return df
                        
df_funnel_expenses = build_funnel_expenses_outer(df_funnel, df_expenses)



# 8. Связать "Воронка" с "Справочник"
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
        df_funnel_expenses,
        df_reference_filtered.drop_duplicates(subset=['Артикул WB']),
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



# 9. Связать "ВоронкаСправочникЗатраты" с "Цены"
try:
    print("Начинаем создавать таблицу ВоронкаСправочникЗатратыЦены...")
    start_time = time.time()  # Запускаем таймер
    df_funnel_prices = pd.merge(df_funnel_reference, df_price, on=["Дата", "Артикул"], how="left")
    df_funnel_prices = format_date_column(df_funnel_prices, 'Дата')

    # Вывод первых 5 строк
    print("Первые 5 строк таблицы ВоронкаСправочникЗатратыЦены:")
    print(df_funnel_prices.head())

    elapsed_time = time.time() - start_time  # Вычисляем затраченное время
    print(f"Таблица ВоронкаСправочникЗатратыЦены успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
except Exception as e:
    print(f"Ошибка при создании таблицы ВоронкаСправочникЗатратыЦены: {e}")



# STOCK WITH DISTRIBUTION

# 10. Получить данные таблицы с SQL (РазмерыНаАгрегаторе)
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

# 11. Получить данные таблицы с SQL (Остатки)
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

# 12. Создание таблицы "ВсегоРазмеров"
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

# 13. Связать "РазмерыНаАгрегаторе" с "ВсегоРазмеров"
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

# 14. Связать "Остатки" с "Дистрибуция"
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



# ДБбезПризнаков

# 15. Получить данные из файлов вложенной папки "Воронка с показами ВБ"
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
                            engine="calamine",
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



# 16. Получить данные из файла !!!_Признаки для артикула и даты для ВБ
try:
    print("Начинаем получать данные для Признаков...")
    start_time = time.time()  # Запускаем таймер
    file_path_features = "!!!_Признаки для артикула и даты для ВБ.xlsx"
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



# 17. Связать "ВоронкаСправочникЗатратыЦены" с "Остатки с дистрибуцией"
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
# del df_funnel_prices
# 17.5 Связать "ДБбезПризнаков" с "Воронка с показами ВБ"
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

# 18. Связать "ДБсПоказами" с "Признаки для артикула"
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

# 19. Связать "ДБсПризнакамиАртикула" с "Признаки для дат"
try:
    print("Начинаем создавать таблицу ДБсПризнаками...")
    start_time = time.time()  # Запускаем таймер
    df_final_db_all_features = pd.merge(df_final_db_item_features, df_date_features, on="Дата", how="left")
    df_final_db_all_features = format_date_column(df_final_db_all_features, 'Дата')


    # Вывод первых 5 строк
    print("Первые 5 строк таблицы ДБсПризнаками:")
    print(df_final_db_all_features.head())
    
    elapsed_time = time.time() - start_time  # Вычисляем затраченное время
    print(f"Таблица ДБсПризнаками успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
except Exception as e:
    print(f"Ошибка при создании таблицы ДБсПризнаками: {e}")


# 20. Получить данные по склейкам из SQL    
# === SQL СЦЕПКИ ВБ ===
sql = """
SELECT [updated_at] as [Дата Обновления], 
		[nmID] as [Артикул WB], 
		[imtID] as [Текущая склейка]
FROM [DBReport].[mp].[wb_scepka]
"""
df_links = pd.read_sql(sql, Engine)
df_links['Артикул WB'] = df_links['Артикул WB'].astype(str)
df_links.to_excel(os.path.join(FOLDER_PATH, f"Склейки товаров\\WB\\{df_links['Дата Обновления'].iloc[0].strftime('%d.%m.%Y')}_Склейка Товаров_WB.xlsx"))

df_final_db_all_features = pd.merge(df_final_db_all_features, df_links[["Артикул WB", "Текущая склейка"]], how='left', on='Артикул WB')

# 21. Получить данные по ассоциациям из SQL  
# === SQL АССОЦИАЦИИ ВБ ===
sql = """
SELECT [Дата]
      ,[nmId]
      ,[ID кампании]
      ,[Кол-во заказаных товаров, шт]
      ,[Заказов на сумму]
  FROM [DBReport].[mp].[wb_marketing_associated]
"""
df_associations = pd.read_sql(sql, Engine)


# 22. Преобразование данных по ассоциациям  
import numpy as np
import pandas as pd

DDMMYYYY = re.compile(r'^\d{2}\.\d{2}\.\d{4}$')   # 01.09.2025
ISO      = re.compile(r'^\d{4}-\d{2}-\d{2}$')     # 2025-09-01

def normalize_date_strict(s: pd.Series) -> pd.Series:
    """Строго нормализует дату к datetime64[ns] без времени.
       Поддерживает РОВНО 'DD.MM.YYYY' и 'YYYY-MM-DD'. Никаких эвристик."""
    ss = s.astype(str).str.strip()
    out = pd.Series(pd.NaT, index=ss.index, dtype='datetime64[ns]')

    m_dd = ss.str.match(DDMMYYYY)
    m_iso = ss.str.match(ISO)

    out[m_dd]  = pd.to_datetime(ss[m_dd],  format='%d.%m.%Y', errors='coerce')
    out[m_iso] = pd.to_datetime(ss[m_iso], format='%Y-%m-%d', errors='coerce')

    # диагностируем «левые» значения
    bad = ~(m_dd | m_iso)
    if bad.any():
        print("ВНИМАНИЕ: найдены строки с недопустимым форматом даты, примеры:")
        print(ss[bad].head(10).to_list())

    return out.dt.normalize()

def allocate_campaign_metrics_rowwise_intsafe(
    df_associations: pd.DataFrame,      # Дата, ID кампании, Кол-во заказаных товаров, Заказов на сумму
    df_item_features: pd.DataFrame,     # Дата, Артикул, Показы (или Показы, всего), Автоматическое_ID кампании / Аукцион_ID кампании
    shows_cols=('Показы, всего','Показы'),
    auto_col='Автоматическое_ID кампании',
    auction_col='Аукцион_ID кампании',
    article_col='Артикул',
    money_decimals=2
):
    A = df_associations.copy()
    I = df_item_features.copy()

    # 1) Даты
    # for c in ['Дата']:
    #     if c in A.columns: A[c] = pd.to_datetime(A[c], errors='coerce')
    #     if c in I.columns: I[c] = pd.to_datetime(I[c], errors='coerce')

    A['Дата'] = normalize_date_strict(A['Дата'])
    I['Дата'] = normalize_date_strict(I['Дата'])

    # 2) coalesce ID кампании
    I['ID кампании'] = pd.Series(index=I.index, dtype=object)
    if auto_col in I.columns:
        I['ID кампании'] = I[auto_col]
    if auction_col in I.columns:
        I['ID кампании'] = I['ID кампании'].combine_first(I[auction_col])

    # 3) колонка показов
    shows_col = next((c for c in shows_cols if c in I.columns), None)
    if shows_col is None:
        raise ValueError(f'Не найдена колонка показов ни из {shows_cols}')

    # 4) агрегируем кампанию по (Дата, ID кампании)
    A2 = (A.groupby(['Дата','ID кампании'], dropna=False)[
            ['Кол-во заказаных товаров, шт','Заказов на сумму']
         ].sum()
         .reset_index())

    # 5) LEFT merge — сохраним исходный индекс, чтобы потом ровно join'ить обратно
    keepI = ['Дата', article_col, 'ID кампании', shows_col]
    I_view = I.loc[:, [c for c in keepI if c in I.columns]].copy()
    I_view['_orig_index'] = I_view.index

    J = I_view.merge(A2, on=['Дата','ID кампании'], how='left')
    # Выравниваем индекс J обратно к исходному для безопасного .join
    J.index = I_view['_orig_index'].values

    # 6) типы
    J[shows_col] = pd.to_numeric(J[shows_col], errors='coerce').fillna(0.0)
    for m in ['Кол-во заказаных товаров, шт', 'Заказов на сумму']:
        if m not in J.columns:
            J[m] = 0.0
        J[m] = pd.to_numeric(J[m], errors='coerce').fillna(0.0)

    # 7) групповые итоги и доля
    keys = ['Дата','ID кампании']
    grp_total_shows = J.groupby(keys, dropna=False)[shows_col].transform('sum')  # это нормально: суммируем показы
    grp_total_qty   = J.groupby(keys, dropna=False)['Кол-во заказаных товаров, шт'].transform('first')
    grp_total_amt   = J.groupby(keys, dropna=False)['Заказов на сумму'].transform('first')


    m_has_total = (grp_total_qty > 0) | (grp_total_amt > 0)
    grp_size    = J.groupby(keys, dropna=False)[shows_col].transform('size').astype('float64')

    J['_share'] = 0.0
    m_shows = (grp_total_shows > 0) & m_has_total
    J.loc[m_shows, '_share'] = (J.loc[m_shows, shows_col] / grp_total_shows[m_shows])

    m_equal = (grp_total_shows == 0) & m_has_total
    denom = grp_size.where(grp_size > 0, 1.0)
    J.loc[m_equal, '_share'] = 1.0 / denom[m_equal]

    # ---------- ШТУКИ: Hamilton ----------
    J['__grp_total_qty'] = grp_total_qty
    J['__raw_qty'] = J['_share'] * J['__grp_total_qty']

    def _distribute_int(group, raw_col, total_col, out_name):
        raw = group[raw_col].to_numpy(dtype=float)
        base = np.floor(raw).astype(np.int64)
        need = int(round(group[total_col].iloc[0] - base.sum()))
        if need > 0:
            frac = raw - base
            order = np.argsort(-frac, kind='mergesort')
            base[order[:need]] += 1
        return pd.Series(base, index=group.index, name=out_name, dtype='int64')

    J['Ассоциированные заказы, шт'] = 0
    J.loc[m_has_total, 'Ассоциированные заказы, шт'] = (
        J[m_has_total]
          .groupby(keys, group_keys=False, dropna=False)
          .apply(_distribute_int, raw_col='__raw_qty', total_col='__grp_total_qty',
                 out_name='Ассоциированные заказы, шт')
    )

    # ---------- ДЕНЬГИ: в копейках, Hamilton ----------
    cents = 10 ** money_decimals
    J['__grp_total_cents'] = np.round(grp_total_amt * cents).astype('int64')
    J['__raw_cents'] = J['_share'] * J['__grp_total_cents']

    def _distribute_cents(group):
        raw = group['__raw_cents'].to_numpy(dtype=float)
        base = np.floor(raw).astype(np.int64)
        need = int(group['__grp_total_cents'].iloc[0] - base.sum())
        if need > 0:
            frac = raw - base
            order = np.argsort(-frac, kind='mergesort')
            base[order[:need]] += 1
        return pd.Series(base, index=group.index, name='__alloc_cents', dtype='int64')

    J['__alloc_cents'] = 0
    J.loc[m_has_total, '__alloc_cents'] = (
        J[m_has_total]
          .groupby(keys, group_keys=False, dropna=False)
          .apply(_distribute_cents)
    )
    J['Ассоциированные заказы, руб'] = (J['__alloc_cents'] / cents).astype('float64')

    # Нули там, где нечего распределять
    J.loc[~m_has_total, ['Ассоциированные заказы, шт','Ассоциированные заказы, руб']] = 0

    # очистка временных колонок
    J.drop(columns=[
        '_share','__grp_total_qty','__raw_qty','__grp_total_cents','__raw_cents','__alloc_cents'
    ], inplace=True, errors='ignore')

    # 8) Вернуть в исходный df_item_features 1:1 по индексу
    out = I.copy()
    out = out.join(J[['Ассоциированные заказы, шт','Ассоциированные заказы, руб']], how='left')
    out[['Ассоциированные заказы, шт','Ассоциированные заказы, руб']] = \
        out[['Ассоциированные заказы, шт','Ассоциированные заказы, руб']].fillna(0)

    assert len(out) == len(df_item_features), f"Row count changed: {len(out)} vs {len(df_item_features)}"
    out['Дата'] = pd.to_datetime(out['Дата'], errors='coerce')\
                 .dt.strftime('%d.%m.%Y')
    return out

df_alloc = allocate_campaign_metrics_rowwise_intsafe(df_associations, df_final_db_all_features)

rub = df_alloc['Ассоциированные заказы, руб']
# проверяем, что все кратны 1 копейке
mask_not_cents = (np.round(rub * 100) != rub * 100)

df_not_cents = df_alloc[mask_not_cents]
print(len(df_not_cents))
print(df_not_cents[['Дата', 'ID кампании', 'Артикул', 'Ассоциированные заказы, руб']].head(10))

df_alloc['Ассоциированные заказы, руб'] = pd.to_numeric(df_alloc['Ассоциированные заказы, руб'])

print('Shape df_alloc: ' + str(df_alloc.shape))


# 23. Сохранить df_alloc в csv
import pyarrow as pa
import pyarrow.csv as csv
table = pa.Table.from_pandas(df_alloc)
csv.write_csv(table, os.path.join(FOLDER_PATH, "ДБсПризнаками.csv"))

# 24. Обновляем и сохраняем Excel-файл
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
            time.sleep(60)  # Пауза 5 секундS

    return False  # Все попытки завершились неудачно

# 25. Обновить файл "Показы и затраты ВБ_2.0.xlsx"
try:
    print("Подготовка данных для ДБ завершена.")
    # input("Начать обновление файлов ДБ? Для подтверждения нажмите Enter...")
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
        # success = True

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