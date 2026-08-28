import os
from sqlalchemy import create_engine
import pyodbc
import win32com.client as win32
import time  # Для измерения времени выполнения
import shutil
import re
from functools import reduce
import numpy as np
import pandas as pd
import pyarrow as pa
import pyarrow.csv as csv

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
 # Constants
FOLDER_PATH = os.path.normpath(r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!")
FOLDER_PATH_FEATURES = r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Дашбоард по рекламным кампаниям"
FOLDER_PATH_FOR_DB = os.path.normpath(r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям")
FOLDER_PATH_DUDL = os.path.normpath(r"\\kari.local\public\all\Агрегаторы\Дашборд реклама WB_OZ")

SQL_SERVER = "cl01sql"
SQL_DATABASE_DBREPORT = "DBReport"
SQL_DATABASE_DBPARTNERS = "DBPartners"

# Период выгрузки из PostgreSQL: 1 месяц и 2 недели (≈45 дней)
PG_LOOKBACK_DAYS = 45

def _get_pg_engine():
    """Подключение к PostgreSQL через SSH-туннель (как в DashboardDBUploader*)."""
    from Preprocessing.ozon_union_db import load_db_config, get_engine
    cfg = load_db_config()
    return get_engine(cfg)

def assemble():
    """Главная функция сборки дашборда WB (SQL-based: воронка и затраты из PostgreSQL)"""

    pg_date_from = (pd.Timestamp.today() - pd.DateOffset(days=PG_LOOKBACK_DAYS)).strftime("%Y-%m-%d")

    # ================================================================
    # 1. Получить данные затрат из PostgreSQL work.wb_marketing
    #    (вместо "Затраты ВБ_2.xlsx")
    # ================================================================
    try:
        print(f"Начинаем получать данные для Затрат ВБ из PostgreSQL work.wb_marketing (с {pg_date_from})...")
        start_time = time.time()

        pg_engine = _get_pg_engine()

        query_expenses = f"""
            SELECT "Дата",
                   CAST("SKU" AS TEXT) AS "SKU",
                   "ID кампании",
                   "Показы",
                   "Клики",
                   "Кол-во добавлений в корзину",
                   "Заказы, шт",
                   "Кол-во заказаных товаров, шт",
                   "Заказов на сумму",
                   "Расход, с НДС"
            FROM work.wb_marketing
            WHERE "Дата" >= '{pg_date_from}'
        """
        df_expenses = pd.read_sql(query_expenses, pg_engine)
        pg_engine.dispose()

        # SKU (barcode) → nmID (product) mapping from SQL Server
        Engine = connect_to_sql(SQL_SERVER, SQL_DATABASE_DBREPORT)
        df_sku_map = pd.read_sql(
            "SELECT DISTINCT CAST(sku AS BIGINT) AS sku, nmID FROM [mp].[wb_sku] WHERE actual=1",
            Engine
        )
        df_sku_map['sku'] = df_sku_map['sku'].astype(str)
        df_sku_map['nmID'] = df_sku_map['nmID'].astype(str)
        df_sku_map = df_sku_map.drop_duplicates(subset='sku', keep='first')

        df_expenses['SKU'] = df_expenses['SKU'].astype(str)
        df_expenses = df_expenses.merge(df_sku_map, left_on='SKU', right_on='sku', how='left')
        df_expenses['Артикул WB'] = df_expenses['nmID'].fillna(df_expenses['SKU'])
        df_expenses.drop(columns=['SKU', 'sku', 'nmID'], inplace=True)

        # Переименование столбцов (как в оригинале)
        column_mapping = {
            "Показы": "Рекламные показы",
            "Клики": "Рекламные клики",
            "Кол-во добавлений в корзину": "Рекламные в корзину",
            "Заказы, шт": "Рекламные Заказы, шт",
            "Кол-во заказаных товаров, шт": "Рекламные заказаных товаров, шт",
            "Заказов на сумму": "Рекламные заказов на сумму",
            "Расход, с НДС": "Расход, руб"
        }
        df_expenses.rename(columns=column_mapping, inplace=True)

        # Форматирование даты
        df_expenses = format_date_column(df_expenses, 'Дата')

        print("Первые 5 строк таблицы Затраты:")
        print(df_expenses.head())

        elapsed_time = time.time() - start_time
        print(f"Данные для Затрат ВБ из PostgreSQL успешно загружены. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при получении данных для Затрат ВБ: {e}")


    # 2. Получить данные из файла "История-затрат-Все.xlsx"
    pattern = r"История-затрат-Все*\.xlsx"

    folder = os.path.join(FOLDER_PATH, "Затраты", "Затраты ВБ")
    pattern = re.compile(r"^История-затрат-Все.*\.xlsx$", re.IGNORECASE)

    files = [e.name for e in os.scandir(folder) if e.is_file() and pattern.match(e.name)]

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

    # ================================================================
    # 2.5 Получить данные Продаж из [DBReport].[mp].[wb_sales_report]
    #     Два общих итога по (Дата, Артикул WB), БЕЗ разбивки по разделам:
    #       Продажи, руб  = GMV net  (Продажа + / Возврат -)
    #       Продажи, шт = Кол-во net (Продажа + / Возврат -)
    #     Ключ wb_sales_report — [Артикул поставщика] (itemid Кари),
    #     маппим в [Артикул WB] (nmID) через WblmRepGetNomenclatureWildberries
    #     (по одному nmID на itemid, чтобы не размножать продажи).
    # ================================================================
    try:
        print(f"Начинаем получать данные Продаж из [mp].[wb_sales_report] (с {pg_date_from})...")
        start_time = time.time()

        Engine = connect_to_sql(SQL_SERVER, SQL_DATABASE_DBREPORT)
        query_sales = f"""
            SELECT
                CAST(s.[Дата продажи] AS date) AS [Дата],
                m.[NMID]                       AS [Артикул WB],
                SUM(CASE WHEN s.[Тип документа] = N'Продажа' THEN s.[Цена розничная]
                         WHEN s.[Тип документа] = N'Возврат' THEN -s.[Цена розничная]
                         ELSE 0 END)           AS [Продажи, руб],
                SUM(CASE WHEN s.[Тип документа] = N'Продажа' THEN s.[Кол-во]
                         WHEN s.[Тип документа] = N'Возврат' THEN -s.[Кол-во]
                         ELSE 0 END)           AS [Продажи, шт]
            FROM [DBReport].[mp].[wb_sales_report] s
            INNER JOIN (
                SELECT [ITEMID], MAX([NMID]) AS [NMID]
                FROM [DBPartners].[dbo].[WblmRepGetNomenclatureWildberries]
                WHERE [NMID] IS NOT NULL
                GROUP BY [ITEMID]
            ) m
                ON UPPER(s.[Артикул поставщика]) = UPPER(m.[ITEMID])
            WHERE s.[Дата продажи] >= '{pg_date_from}'
            GROUP BY CAST(s.[Дата продажи] AS date), m.[NMID]
        """
        df_sales = pd.read_sql(query_sales, Engine)
        df_sales['Артикул WB'] = df_sales['Артикул WB'].fillna('').astype(str).str.strip().str.upper()
        df_sales = format_date_column(df_sales, 'Дата')

        print("Первые 5 строк таблицы Продажи:")
        print(df_sales.head())

        elapsed_time = time.time() - start_time
        print(f"Данные Продаж успешно загружены. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при получении данных Продаж: {e}")
        df_sales = pd.DataFrame(columns=['Дата', 'Артикул WB', 'Продажи, руб', 'Продажи, шт'])

    # 3. Создание широкой таблицы по типам активности
    TYPE_DICT_RAW = {}

    SECTION_PRIORITY = {'Ручная Ставка': 2, 'Единая Ставка': 1, np.nan: 0}

    # Как называть «неопределённый» раздел
    MISC_SECTION_NAME = 'Прочее'

    # Кандидаты для названия столбца расхода в df_expenses
    SPEND_COL_CANDIDATES = ['Расход, ₽', 'Расход, руб', 'Расход, Р', 'Расход']

    # Ключи в wide-таблице
    KEY_COLS = ['Дата', 'Артикул WB']

    # [v2] Разделы, которые считаем платными
    PAID_SECTIONS = ['Единая Ставка', 'Ручная Ставка']


    # ===================== ВСПОМОГАТЕЛЬНЫЕ =====================

    def _pick_spend_col(df: pd.DataFrame) -> str:
        for c in SPEND_COL_CANDIDATES:
            if c in df.columns:
                return c
        raise KeyError("Не найдена колонка расхода среди: " + ", ".join(SPEND_COL_CANDIDATES))


    def build_camp_ref(df_costs_hist: pd.DataFrame) -> pd.DataFrame:
        if df_costs_hist.empty:
            return pd.DataFrame({'ID кампании': pd.Series(dtype='Int64'),
                                'Раздел': pd.Series(dtype='object')})

        dfp = df_costs_hist[['ID кампании', 'Раздел']].copy()
        dfp['ID кампании'] = pd.to_numeric(dfp['ID кампании'], errors='coerce').astype('Int64')

        dfp['Раздел'] = dfp['Раздел'].replace(TYPE_DICT_RAW)

        dfp['_prio'] = dfp['Раздел'].map(SECTION_PRIORITY).fillna(0).astype(int)
        dfp = dfp.sort_values(['ID кампании', '_prio'], ascending=[True, False])
        camp_ref = dfp.drop_duplicates(subset=['ID кампании'], keep='first')[['ID кампании', 'Раздел']].copy()

        assert not camp_ref['ID кампании'].duplicated().any(), "camp_ref содержит дубликаты ID кампании"
        return camp_ref


    def merge_expenses_with_sections(df_expenses: pd.DataFrame, camp_ref: pd.DataFrame) -> pd.DataFrame:
        out = df_expenses.copy()
        out['ID кампании'] = pd.to_numeric(out['ID кампании'], errors='coerce').astype('Int64')
        out = out.merge(camp_ref, on='ID кампании', how='left', validate='m:1')
        out['Раздел'] = out['Раздел'].fillna(MISC_SECTION_NAME)
        return out


    def build_wide_by_section(df_merged: pd.DataFrame, key_cols=KEY_COLS) -> pd.DataFrame:
        df = df_merged.copy()
        key_cols = list(key_cols)
        aux_cols = ['ID кампании', 'Раздел']
        metric_cols = [c for c in df.columns if c not in set(key_cols + aux_cols)]

        for c in metric_cols:
            df[c] = pd.to_numeric(df[c], errors='coerce')

        agg = (df.groupby(key_cols + ['Раздел'], as_index=False)[metric_cols]
                .sum(min_count=1))

        wide = agg.pivot_table(index=key_cols, columns='Раздел', values=metric_cols,
                            aggfunc='sum', fill_value=0)
        wide.columns = [f"{sec}_{met}" for met, sec in wide.columns.to_flat_index()]
        wide = wide.reset_index()

        sections = list(pd.unique(df['Раздел'].dropna()))

        for met in metric_cols:
            cols_to_sum = [f"{sec}_{met}" for sec in sections if f"{sec}_{met}" in wide.columns]
            if cols_to_sum:
                wide[met] = wide[cols_to_sum].sum(axis=1)

        id_wide = (df.groupby(key_cols + ['Раздел'], as_index=False)['ID кампании']
                    .first()
                    .pivot(index=key_cols, columns='Раздел', values='ID кампании')
                    .rename(columns=lambda sec: f"{sec}_ID кампании")
                    .reset_index())
        wide = wide.merge(id_wide, on=key_cols, how='left')

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
            only = presence[PAID_SECTIONS[0]]
            wide['Тип активности'] = np.where(only, PAID_SECTIONS[0], 'Органика')

        flags_block = (
            df.assign(_hit=1)
            .pivot_table(index=key_cols, columns='Раздел', values='_hit',
                        aggfunc='max', fill_value=0)
            .astype('int8')
            .reset_index()
        )
        wide = wide.merge(flags_block, on=key_cols, how='left')

        for sec in PAID_SECTIONS:
            if sec not in wide.columns:
                wide[sec] = 0

        wide['Органика'] = (1 - wide[PAID_SECTIONS].max(axis=1)).astype('int8')

        section_order = PAID_SECTIONS + [s for s in sections if s not in PAID_SECTIONS]
        ordered = key_cols + ['Тип активности']

        ordered += [s for s in PAID_SECTIONS if s in wide.columns]
        ordered += [s for s in sections if (s not in PAID_SECTIONS and s in wide.columns and s != 'Органика')]
        ordered += ['Органика'] if 'Органика' in wide.columns else []

        for sec in section_order:
            id_col = f"{sec}_ID кампании"
            if id_col in wide.columns:
                ordered.append(id_col)
            ordered += [c for c in wide.columns if c.startswith(f"{sec}_") and c != id_col]
        ordered += [c for c in metric_cols if c in wide.columns]

        wide = wide[[c for c in ordered if c in wide.columns]].copy()
        return wide


    def build_sum_table(df_expenses: pd.DataFrame, df_costs_hist: pd.DataFrame) -> pd.DataFrame:
        spend_col = _pick_spend_col(df_expenses)
        base_sum = pd.to_numeric(df_expenses[spend_col], errors='coerce').sum()

        camp_ref = build_camp_ref(df_costs_hist)
        df_merged = merge_expenses_with_sections(df_expenses, camp_ref)

        df_sum = build_wide_by_section(df_merged, key_cols=KEY_COLS)

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
        start_time = time.time()
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

        print("Первые 5 строк таблицы Цена:")
        print(df_price.head())

        elapsed_time = time.time() - start_time
        print(f"Данные для Цен успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при получении данных для Цен: {e}")



    # 5. Получить данные из файла "Справочник.xlsx"
    try:
        print("Начинаем получать данные для Справочника...")
        start_time = time.time()
        file_path_reference = os.path.join(FOLDER_PATH, "Справочник.xlsx")

        if os.path.exists(file_path_reference):
            columns_to_read = [
                "Артикул", "Наименование", "Коллекция",
                "Бренд", "Размер", "Сезон", "Направление", "Розничный отдел",
                "Модель", "Группа", "Бизнес-группа", "Техсегмент",
                "Байер", "Две последние коллекции", "Основной артикул",
                "Артикул WB", "Себестоимость с НДС",
                "Процент выкупа ВБ", "НДС", "Ответственный за группу", "Группа для отчетов"
            ]

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

            df_reference = pd.read_excel(
                file_path_reference,
                sheet_name="Выгрузка для справочника",
                engine="openpyxl",
                usecols=columns_to_read,
                dtype=column_dtypes
            )

            df_reference = df_reference.drop_duplicates()

            print("Первые 5 строк таблицы Справочник:")
            print(df_reference.head())

            elapsed_time = time.time() - start_time
            print(f"Данные для Справочника успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
        else:
            print("Файл 'Справочник.xlsx' не найден.")
    except Exception as e:
        print(f"Ошибка при получении данных для Справочника: {e}")



    # ================================================================
    # 6. Получить данные воронки из PostgreSQL work.wb_sales_funnel_lk
    #    (вместо файлов из папки "Показатели по неделям ВБ")
    # ================================================================
    try:
        print(f"Начинаем получать данные для Воронки из PostgreSQL work.wb_sales_funnel_lk (с {pg_date_from})...")
        start_time = time.time()

        pg_engine = _get_pg_engine()

        query_funnel = f"""
            SELECT "Дата",
                   CAST("Артикул WB" AS TEXT)                       AS "Артикул WB",
                   "Показы",
                   "Кол-во переходов в карточку товара"              AS "Показы на карточке товара",
                   "Положили в корзину, штук"                        AS "Положили в корзину",
                   "Заказали товаров, шт"                            AS "Заказали, шт",
                   "Выкупили товаров, шт."                           AS "Выкупили, шт",
                   "Отменили товаров, шт."                           AS "Отменили, шт",
                   "Заказали на сумму, руб."                         AS "Заказали на сумму, руб",
                   "Выкупили на сумму, руб."                         AS "Выкупили на сумму, руб",
                   "Отменили на сумму, руб."                         AS "Отменили на сумму, руб",
                   "Средняя цена, руб."                              AS "Средняя цена, руб"
            FROM work.wb_sales_funnel_lk
            WHERE "Дата" >= '{pg_date_from}'
        """
        df_funnel = pd.read_sql(query_funnel, pg_engine)
        pg_engine.dispose()

        # Столбцов «Рейтинг карточки» и «Рейтинг по отзывам» нет в таблице — заполняем нулями
        df_funnel["Рейтинг карточки"] = 0.0
        df_funnel["Рейтинг по отзывам"] = 0.0

        df_funnel = format_date_column(df_funnel, 'Дата')

        # Приводим к тому же набору колонок, что ожидает остальной пайплайн
        final_columns = [
            "Дата", "Артикул WB", "Рейтинг карточки", "Показы", "Показы на карточке товара",
            "Положили в корзину", "Заказали, шт", "Выкупили, шт",
            "Отменили, шт", "Заказали на сумму, руб", "Выкупили на сумму, руб",
            "Отменили на сумму, руб", "Средняя цена, руб", "Рейтинг по отзывам"
        ]
        df_funnel = df_funnel[final_columns]
        df_funnel["Рейтинг по отзывам"] = df_funnel["Рейтинг по отзывам"].fillna(0)

        print("Первые 5 строк таблицы Воронка:")
        print(df_funnel.head())

        elapsed_time = time.time() - start_time
        print(f"Данные для Воронки из PostgreSQL успешно загружены. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при получении данных для Воронки: {e}")



    # 7. Связать "Воронка" с "Затраты ВБ"

    # ----------- настройки -----------
    KEY_COLS = ['Дата', 'Артикул WB']
    SPEND_CANDS = ['Расход, руб', 'Единая Ставка_Расход, руб', 'Ручная Ставка_Расход, руб']
    ALLOC_MODE = 'even'
    NO_SUM_CONTAINS = ['ID кампании']

    # ----------- утилиты -----------
    def _pick_spend_col(df: pd.DataFrame) -> str:
        for c in SPEND_CANDS:
            if c in df.columns:
                return c
        raise KeyError(f"Не найден столбец расхода среди: {SPEND_CANDS}")

    def _clean_money_series(s: pd.Series) -> pd.Series:
        if s.dtype.kind in 'biufc':
            return s
        s = (s.astype(str)
            .str.replace(' ', '', regex=False)
            .str.replace(' ', '', regex=False)
            .str.replace(',', '.', regex=False))
        s = s.str.replace(r'[^0-9\.\-]+', '', regex=True)
        return pd.to_numeric(s, errors='coerce')

    def _norm_keys(df: pd.DataFrame) -> pd.DataFrame:
        out = df.copy()
        out['Артикул WB'] = (out['Артикул WB']
                            .fillna('').astype(str).str.strip().str.upper())
        return out

    def _make_right_unique(df_expenses: pd.DataFrame) -> pd.DataFrame:
        df = _norm_keys(df_expenses).copy()

        spend_col = _pick_spend_col(df)
        df[spend_col] = _clean_money_series(df[spend_col])

        key = KEY_COLS
        other = [c for c in df.columns if c not in key]

        flag_cols = [c for c in ['Единая Ставка', 'Ручная Ставка', 'Органика'] if c in df.columns]
        no_sum_cols = [c for c in other if any(tok in c for tok in NO_SUM_CONTAINS)]

        num_cols = []
        for c in other:
            if c in flag_cols or c in no_sum_cols:
                continue
            df[c] = pd.to_numeric(df[c], errors='ignore')
            if pd.api.types.is_numeric_dtype(df[c]):
                num_cols.append(c)

        text_cols = [c for c in other if c not in num_cols and c not in flag_cols and c not in no_sum_cols]

        agg = {c: 'sum' for c in num_cols}
        agg.update({c: 'max' for c in flag_cols})
        agg.update({c: 'first' for c in text_cols})
        agg.update({c: 'first' for c in no_sum_cols})

        right_u = (df.groupby(key, as_index=False).agg(agg))

        assert right_u.duplicated(key).sum() == 0, "Правая часть после агрегации не уникальна по ключу"
        return right_u

    def _allocate_on_left_dupes(df_merged: pd.DataFrame, spend_col: str, mode: str) -> pd.DataFrame:
        out = df_merged.copy()
        out[spend_col] = pd.to_numeric(out[spend_col], errors='coerce').fillna(0.0)

        counts = out.groupby(KEY_COLS, dropna=False)[spend_col].transform('size')

        if mode == 'first':
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
        left = _norm_keys(df_funnel)
        right_u = _make_right_unique(df_expenses)

        spend_col = _pick_spend_col(right_u)
        base_by_date = (right_u.groupby('Дата', as_index=False)[spend_col]
                            .sum(min_count=1).rename(columns={spend_col: 'Расход_до'}))
        base_total = right_u[spend_col].sum()

        df = left.merge(right_u, on=KEY_COLS, how='outer', indicator=True)

        df = _allocate_on_left_dupes(df, spend_col, alloc_mode)

        after_by_date = (df.groupby('Дата', as_index=False)[spend_col]
                        .sum(min_count=1).rename(columns={spend_col: 'Расход_после'}))
        chk = base_by_date.merge(after_by_date, on='Дата', how='outer').fillna(0)
        drift = chk.loc[~np.isclose(chk['Расход_до'], chk['Расход_после'], rtol=1e-9, atol=1e-6)]
        after_total = df[spend_col].sum()

        if not drift.empty or not np.isclose(after_total, base_total, rtol=1e-9, atol=1e-6):
            print("[WARN] Найдены расхождения после объединения (первые 10 дат):")
            print(drift.sort_values('Расход_до', ascending=False).head(10))
            print(f"[WARN] Общая сумма: до={base_total:,.2f} | после={after_total:,.2f}")

        df.drop(columns=['_merge'], inplace=True, errors='ignore')

        return df

    df_funnel_expenses = build_funnel_expenses_outer(df_funnel, df_expenses)

    # 7.5 Добавить Продажи (Продажи, руб, Продажи, шт) по (Дата, Артикул WB).
    #     Мёржим ПОСЛЕ широкой таблицы — как общие итоги, без разбивки по Типам активности.
    try:
        print("Начинаем добавлять Продажи в df_funnel_expenses...")
        start_time = time.time()

        df_funnel_expenses['Артикул WB'] = (df_funnel_expenses['Артикул WB']
            .fillna('').astype(str).str.strip().str.upper())

        df_funnel_expenses = pd.merge(
            df_funnel_expenses,
            df_sales,
            on=['Дата', 'Артикул WB'],
            how='left'
        )
        df_funnel_expenses[['Продажи, руб', 'Продажи, шт']] = (
            df_funnel_expenses[['Продажи, руб', 'Продажи, шт']].fillna(0)
        )

        print("Первые 5 строк df_funnel_expenses с Продажами:")
        print(df_funnel_expenses[['Дата', 'Артикул WB', 'Продажи, руб', 'Продажи, шт']].head())

        elapsed_time = time.time() - start_time
        print(f"Продажи успешно добавлены. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при добавлении Продаж в df_funnel_expenses: {e}")



    # 8. Связать "Воронка" с "Справочник"
    try:
        print("Начинаем создавать таблицу ВоронкаСправочник...")
        start_time = time.time()

        required_columns = ["Дата", "Артикул WB"]
        for col in required_columns:
            if col not in df_funnel.columns:
                print(f"Ошибка: Отсутствует столбец '{col}' в df_funnel.")
                exit()

        reference_columns = [
            "Артикул", "Наименование", "Коллекция", "Бренд", "Сезон", "Направление",
            "Розничный отдел", "Модель", "Группа", "Бизнес-группа", "Техсегмент",
            "Байер", "Две последние коллекции", "Основной артикул", "Артикул WB",
            "Себестоимость с НДС", "Процент выкупа ВБ", "НДС", "Ответственный за группу", "Группа для отчетов"
        ]

        df_reference_filtered = df_reference[reference_columns]

        df_funnel["Артикул WB"] = df_funnel["Артикул WB"].fillna('').astype(str)
        df_reference_filtered["Артикул WB"] = df_reference_filtered["Артикул WB"].fillna('').astype(str)

        df_funnel_reference = pd.merge(
            df_funnel_expenses,
            df_reference_filtered.drop_duplicates(subset=['Артикул WB']),
            left_on="Артикул WB",
            right_on="Артикул WB",
            how="left"
        )

        df_funnel_reference = df_funnel_reference.drop_duplicates()
        df_funnel_reference = format_date_column(df_funnel_reference, 'Дата')

        print("Первые 5 строк таблицы ВоронкаСправочник:")
        print(df_funnel_reference.head())

        elapsed_time = time.time() - start_time
        print(f"Таблица ВоронкаСправочник успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при создании таблицы ВоронкаСправочник: {e}")



    # 9. Связать "ВоронкаСправочникЗатраты" с "Цены"
    try:
        print("Начинаем создавать таблицу ВоронкаСправочникЗатратыЦены...")
        start_time = time.time()
        df_funnel_prices = pd.merge(df_funnel_reference, df_price, on=["Дата", "Артикул"], how="left")
        df_funnel_prices = format_date_column(df_funnel_prices, 'Дата')

        print("Первые 5 строк таблицы ВоронкаСправочникЗатратыЦены:")
        print(df_funnel_prices.head())

        elapsed_time = time.time() - start_time
        print(f"Таблица ВоронкаСправочникЗатратыЦены успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при создании таблицы ВоронкаСправочникЗатратыЦены: {e}")



    # STOCK WITH DISTRIBUTION

    # 10. Получить данные таблицы с SQL (РазмерыНаАгрегаторе)
    try:
        print("Начинаем получать данные для РазмеровНаАгрегаторе...")
        start_time = time.time()
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
        print("Первые 5 строк таблицы РазмерыНаАгрегаторе:")
        print(df_sizes.head())

        elapsed_time = time.time() - start_time
        print(f"Данные для РазмеровНаАгрегаторе успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при получении данных для РазмеровНаАгрегаторе: {e}")

    # 11. Получить данные таблицы с SQL (Остатки)
    try:
        print("Начинаем получать данные для Остатков...")
        start_time = time.time()
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

        print("Первые 5 строк таблицы Остатков:")
        print(df_stock.head())

        elapsed_time = time.time() - start_time
        print(f"Данные для Остатков успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при получении данных для Остатков: {e}")

    # 12. Создание таблицы "ВсегоРазмеров"
    try:
        print("Начинаем создавать таблицу ВсегоРазмеров...")
        start_time = time.time()

        required_columns = ["Артикул", "Размер"]
        for col in required_columns:
            if col not in df_reference.columns:
                print(f"Ошибка: Отсутствует столбец '{col}' в df_reference.")
                exit()

        df_reference["Размер"] = df_reference["Размер"].astype(str).str.strip().replace('', None)

        df_reference_unique = (
            df_reference
            .drop_duplicates(subset=["Артикул", "Размер"])
            .groupby("Артикул")["Размер"]
            .apply(lambda sizes: len(sizes.dropna().unique()) if len(sizes.dropna()) > 0 else 1)
            .reset_index(name="Всего размеров")
        )

        print("Первые 5 строк таблицы ВсегоРазмеров:")
        print(df_reference_unique.head())

        elapsed_time = time.time() - start_time
        print(f"Таблица ВсегоРазмеров успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при создании таблицы ВсегоРазмеров: {e}")

    # 13. Связать "РазмерыНаАгрегаторе" с "ВсегоРазмеров"
    try:
        print("Начинаем создавать таблицу Дистрибуция...")
        start_time = time.time()

        df_distribution = pd.merge(df_sizes, df_reference_unique, on="Артикул", how="left")

        df_distribution["Дистрибуция"] = df_distribution.apply(
            lambda row: row["Колво размеров"] / row["Всего размеров"] if row["Всего размеров"] != 0 else 0,
            axis=1
        )

        df_distribution = df_distribution[["Дата", "Артикул", "Дистрибуция"]]
        df_distribution = format_date_column(df_distribution, 'Дата')

        print("Первые 5 строк таблицы Дистрибуция:")
        print(df_distribution.head())

        elapsed_time = time.time() - start_time
        print(f"Таблица Дистрибуция успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при создании таблицы Дистрибуция: {e}")

    # 14. Связать "Остатки" с "Дистрибуция"
    try:
        print("Начинаем создавать таблицу Остатки с дистрибуцией...")
        start_time = time.time()
        df_stock_with_distribution = pd.merge(df_stock, df_distribution, on=["Дата", "Артикул"], how="left")
        df_stock_with_distribution = format_date_column(df_stock_with_distribution, 'Дата')
        print("Первые 5 строк таблицы Остатки с дистрибуцией:")
        print(df_stock_with_distribution.head())

        elapsed_time = time.time() - start_time
        print(f"Таблица Остатки с дистрибуцией успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при создании таблицы Остатки с дистрибуцией: {e}")



    # ДБбезПризнаков

    # 15. Получить данные из файлов вложенной папки "Воронка с показами ВБ"
    try:
        print("Начинаем получать данные для Воронки с показами ВБ...")
        start_time = time.time()
        folder_path_funnel_shows = os.path.join(FOLDER_PATH, "Воронка с показами ВБ")
        df_funnel_shows = pd.DataFrame()

        if os.path.exists(folder_path_funnel_shows):
            for file in os.listdir(folder_path_funnel_shows):
                file_path = os.path.join(folder_path_funnel_shows, file)

                if is_hidden(file_path):
                    print(f"Пропущен скрытый файл: {file}")
                    continue

                if file.endswith((".xlsx", ".xls")):
                    try:
                        if file.endswith(".xlsx"):
                            temp_df = pd.read_excel(
                                file_path,
                                sheet_name="Лист1",
                                engine="calamine",
                                usecols=range(6),
                                dtype={"nm": str}
                            )
                        elif file.endswith(".xls"):
                            temp_df = pd.read_excel(
                                file_path,
                                sheet_name="Лист1",
                                engine="xlrd",
                                usecols=range(6),
                                dtype={"nm": str}
                            )

                        temp_df.rename(columns={
                            "data_day": "Дата",
                            "nm": "Артикул WB",
                        }, inplace=True, errors="ignore")

                        temp_df = temp_df[[c for c in ("Дата", "Артикул WB") if c in temp_df.columns]]

                        temp_df["Артикул WB"] = temp_df["Артикул WB"].str.strip().str.split('.').str[0]

                        temp_df = format_date_column(temp_df, 'Дата')

                        df_funnel_shows = pd.concat([df_funnel_shows, temp_df])
                    except Exception as e:
                        print(f"Ошибка при чтении файла {file}: {e}")

            print("Первые 5 строк таблицы Воронка с показами:")
            print(df_funnel_shows.head())

            elapsed_time = time.time() - start_time
            print(f"Данные для Воронки с показами ВБ успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
        else:
            print("Папка 'Воронка с показами ВБ' не найдена.")
    except Exception as e:
        print(f"Ошибка при получении данных для Воронки с показами ВБ: {e}")



    # 16. Получить данные из файла !!!_Признаки для артикула и даты для ВБ
    try:
        print("Начинаем получать данные для Признаков...")
        start_time = time.time()
        file_path_features = "!!!_Признаки для артикула и даты для ВБ.xlsx"
        if os.path.exists(file_path_features):
            df_item_features = pd.read_excel(file_path_features, sheet_name="Признаки для артикула", dtype=str, engine="openpyxl")
            df_date_features = pd.read_excel(file_path_features, sheet_name="Признаки для дат", dtype={0: "datetime64[ns]", **{i: str for i in range(1, 6)}}, engine="openpyxl")

            df_item_features = handle_errors(df_item_features)
            df_date_features = handle_errors(df_date_features)

            df_date_features = format_date_column(df_date_features, 'Дата')

            print("Первые 5 строк таблицы Признаки артикула:")
            print(df_item_features.head())

            print("Первые 5 строк таблицы Признаки дат:")
            print(df_date_features.head())

            elapsed_time = time.time() - start_time
            print(f"Данные для Признаков успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
        else:
            print("Файл '!!!_Признаки для артикула и даты для ВБ.xlsx' не найден.")
    except Exception as e:
        print(f"Ошибка при получении данных для Признаков: {e}")



    # 17. Связать "ВоронкаСправочникЗатратыЦены" с "Остатки с дистрибуцией"
    try:
        print("Начинаем создавать таблицу ДБбезПризнаков...")
        start_time = time.time()
        df_final_db = pd.merge(df_funnel_prices, df_stock_with_distribution, on=["Дата", "Артикул"], how="left")
        df_final_db = format_date_column(df_final_db, 'Дата')
        print("Первые 5 строк таблицы ДБбезПризнаков:")
        print(df_final_db.head())

        elapsed_time = time.time() - start_time
        print(f"Таблица ДБбезПризнаков успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при создании таблицы ДБбезПризнаков: {e}")

    # 17.5 Связать "ДБбезПризнаков" с "Воронка с показами ВБ"
    try:
        print("Начинаем создавать таблицу ДБсПоказами...")
        start_time = time.time()

        required_columns = ["Дата", "Артикул WB"]
        for col in required_columns:
            if col not in df_funnel_shows.columns:
                print(f"Ошибка: Отсутствует столбец '{col}' в df_funnel_shows.")
                exit()

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

        df_final_db_with_shows = pd.merge(
            df_final_db,
            df_funnel_shows,
            left_on=["Дата", "Артикул WB"],
            right_on=["Дата", "Артикул WB"],
            how="left"
        )

        df_final_db_with_shows = format_date_column(df_final_db_with_shows, 'Дата')

        print("Первые 5 строк таблицы ДБсПоказами:")
        print(df_final_db_with_shows.head())

        elapsed_time = time.time() - start_time
        print(f"Таблица ДБсПоказами успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при создании таблицы ДБсПоказами: {e}")

    # 18. Связать "ДБсПоказами" с "Признаки для артикула"
    try:
        print("Начинаем создавать таблицу ДБсПризнакамиАртикула...")
        start_time = time.time()
        df_final_db_item_features = pd.merge(df_final_db_with_shows, df_item_features, on="Артикул", how="left")
        df_final_db_item_features = format_date_column(df_final_db_item_features, 'Дата')

        print("Первые 5 строк таблицы ДБсПризнакамиАртикула:")
        print(df_final_db_item_features.head())

        elapsed_time = time.time() - start_time
        print(f"Таблица ДБсПризнакамиАртикула успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при создании таблицы ДБсПризнакамиАртикула: {e}")

    # 19. Связать "ДБсПризнакамиАртикула" с "Признаки для дат"
    try:
        print("Начинаем создавать таблицу ДБсПризнаками...")
        start_time = time.time()
        df_final_db_all_features = pd.merge(df_final_db_item_features, df_date_features, on="Дата", how="left")
        df_final_db_all_features = format_date_column(df_final_db_all_features, 'Дата')

        print("Первые 5 строк таблицы ДБсПризнаками:")
        print(df_final_db_all_features.head())

        elapsed_time = time.time() - start_time
        print(f"Таблица ДБсПризнаками успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при создании таблицы ДБсПризнаками: {e}")


    # 20. Получить данные по склейкам из SQL
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

    DDMMYYYY = re.compile(r'^\d{2}\.\d{2}\.\d{4}$')
    ISO      = re.compile(r'^\d{4}-\d{2}-\d{2}$')

    def normalize_date_strict(s: pd.Series) -> pd.Series:
        ss = s.astype(str).str.strip()
        out = pd.Series(pd.NaT, index=ss.index, dtype='datetime64[ns]')

        m_dd = ss.str.match(DDMMYYYY)
        m_iso = ss.str.match(ISO)

        out[m_dd]  = pd.to_datetime(ss[m_dd],  format='%d.%m.%Y', errors='coerce')
        out[m_iso] = pd.to_datetime(ss[m_iso], format='%Y-%m-%d', errors='coerce')

        bad = ~(m_dd | m_iso)
        if bad.any():
            print("ВНИМАНИЕ: найдены строки с недопустимым форматом даты, примеры:")
            print(ss[bad].head(10).to_list())

        return out.dt.normalize()

    def allocate_campaign_metrics_rowwise_intsafe(
        df_associations: pd.DataFrame,
        df_item_features: pd.DataFrame,
        shows_cols=('Показы, всего','Показы'),
        auto_col='Единая Ставка_ID кампании',
        auction_col='Ручная Ставка_ID кампании',
        article_col='Артикул',
        money_decimals=2
    ):
        A = df_associations.copy()
        I = df_item_features.copy()

        A['Дата'] = normalize_date_strict(A['Дата'])
        I['Дата'] = normalize_date_strict(I['Дата'])

        I['ID кампании'] = pd.Series(index=I.index, dtype=object)
        if auto_col in I.columns:
            I['ID кампании'] = I[auto_col]
        if auction_col in I.columns:
            I['ID кампании'] = I['ID кампании'].combine_first(I[auction_col])

        shows_col = next((c for c in shows_cols if c in I.columns), None)
        if shows_col is None:
            raise ValueError(f'Не найдена колонка показов ни из {shows_cols}')

        A2 = (A.groupby(['Дата','ID кампании'], dropna=False)[
                ['Кол-во заказаных товаров, шт','Заказов на сумму']
            ].sum()
            .reset_index())

        keepI = ['Дата', article_col, 'ID кампании', shows_col]
        I_view = I.loc[:, [c for c in keepI if c in I.columns]].copy()
        I_view['_orig_index'] = I_view.index

        J = I_view.merge(A2, on=['Дата','ID кампании'], how='left')
        J.index = I_view['_orig_index'].values

        J[shows_col] = pd.to_numeric(J[shows_col], errors='coerce').fillna(0.0)
        for m in ['Кол-во заказаных товаров, шт', 'Заказов на сумму']:
            if m not in J.columns:
                J[m] = 0.0
            J[m] = pd.to_numeric(J[m], errors='coerce').fillna(0.0)

        keys = ['Дата','ID кампании']
        grp_total_shows = J.groupby(keys, dropna=False)[shows_col].transform('sum')
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

        J.loc[~m_has_total, ['Ассоциированные заказы, шт','Ассоциированные заказы, руб']] = 0

        J.drop(columns=[
            '_share','__grp_total_qty','__raw_qty','__grp_total_cents','__raw_cents','__alloc_cents'
        ], inplace=True, errors='ignore')

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
    mask_not_cents = (np.round(rub * 100) != rub * 100)

    df_not_cents = df_alloc[mask_not_cents]
    print(len(df_not_cents))
    print(df_not_cents[['Дата', 'ID кампании', 'Артикул', 'Ассоциированные заказы, руб']].head(10))

    df_alloc['Ассоциированные заказы, руб'] = pd.to_numeric(df_alloc['Ассоциированные заказы, руб'])

    print('Shape df_alloc: ' + str(df_alloc.shape))

    # =====================================================================
    # 23 [v2]. Переименование колонок согласно "Описание столбцов дашборда ВБ.xlsx"
    # =====================================================================
    DASH = ' — '

    WB_PREFIX_BODY_MAP = {
        'Расход, руб':                     'Расходы на рекламу',
        'Рекламные Заказы, шт':            'Количество рекламных заказов',
        'Рекламные в корзину':             'Добавлений в корзину от рекламы',
        'Рекламные заказаных товаров, шт': 'Количество товаров в рекламных заказах',
        'Рекламные заказов на сумму':      'Сумма рекламных заказов',
        'Рекламные клики':                 'Количество кликов по рекламе',
        'Рекламные показы':                'Количество рекламных показов',
        'ID кампании':                     'Идентификатор рекламной кампании',
    }
    WB_PREFIXES = ('Единая Ставка', 'Ручная Ставка')

    WB_DIRECT = {
        'Рейтинг карточки':           'Рейтинг карточки товара',
        'Показы':                     'Количество показов товара',
        'Показы на карточке товара':  'Просмотры карточки товара',
        'Положили в корзину':         'Добавлений в корзину',
        'Заказали, шт':               'Количество заказов',
        'Выкупили, шт':               'Количество выкупов',
        'Отменили, шт':               'Количество отмен',
        'Заказали на сумму, руб':     'Сумма заказов',
        'Выкупили на сумму, руб':     'Сумма выкупов',
        'Отменили на сумму, руб':     'Сумма отмен',
        'Средняя цена, руб':          'Средняя цена продажи',

        'Рекламные показы':                  'Количество рекламных показов',
        'Рекламные заказов на сумму':        'Сумма рекламных заказов',
        'Рекламные Заказы, шт':              'Количество рекламных заказов',
        'Рекламные клики':                   'Количество кликов по рекламе',
        'Рекламные в корзину':               'Добавлений в корзину от рекламы',
        'Рекламные заказаных товаров, шт':   'Количество товаров в рекламных заказах',
        'Расход, руб':                       'Расходы на рекламу',

        'Единая Ставка':                'Признак автоматической рекламной кампании',
        'Ручная Ставка':                'Признак аукционной рекламной кампании',
        'Органика':                     'Признак органического трафика',
        'Тип активности':               'Тип рекламной активности',

        'Наименование':                 'Название товара',
        'Группа':                       'Товарная группа',
        'Техсегмент':                   'Технологический сегмент',
        'Себестоимость с НДС':          'Себестоимость товара с НДС',
        'Процент выкупа ВБ':            'Процент выкупа на Wildberries',
        'НДС':                          'Ставка НДС',
        'Ответственный за группу':      'Ответственный за товарную группу',
        'Группа для отчетов':           'Группа для отчётов',

        'Цена':                         'Розничная цена товара',
        'Остаток Агрегатора':           'Остаток товара на складе агрегатора',
        'Дистрибуция':                  'Дистрибуция размеров',

        'Признак Артикула 1':           'Признак товара 1',
        'Признак Артикула 2':           'Признак товара 2',
        'Признак Артикула 3':           'Признак товара 3',
        'Признак Артикула 4':           'Признак товара 4',
        'Признак Артикула 5':           'Признак товара 5',
        'Признак Даты 1':               'Признак даты 1',
        'Признак Даты 2':               'Признак даты 2',
        'Признак Даты 3':               'Признак даты 3',
        'Признак Даты 4':               'Признак даты 4',
        'Признак Даты 5':               'Признак даты 5',

        'Текущая склейка':              'Идентификатор текущей склейки товара',
        'ID кампании':                  'Идентификатор рекламной кампании',
        'Ассоциированные заказы, шт':   'Количество ассоциированных заказов',
        'Ассоциированные заказы, руб': 'Сумма ассоциированных заказов',
    }

    def _build_wb_rename_map(columns):
        rename = {}
        for col in columns:
            if col in WB_DIRECT:
                rename[col] = WB_DIRECT[col]
                continue
            for p in WB_PREFIXES:
                pref = p + '_'
                if col.startswith(pref):
                    body = col[len(pref):]
                    new_body = WB_PREFIX_BODY_MAP.get(body, body)
                    rename[col] = f'{p}{DASH}{new_body}'
                    break
        return rename

    rename_map = _build_wb_rename_map(df_alloc.columns.tolist())
    df_alloc = df_alloc.rename(columns=rename_map)

    cols = list(df_alloc.columns)
    if len(cols) != len(set(cols)):
        from collections import Counter
        dup = [c for c, n in Counter(cols).items() if n > 1]
        raise RuntimeError(f"[v2] После переименования возникли дубликаты колонок: {dup}")
    print(f'[v2] Колонок после переименования: {len(cols)}; дубликатов: 0')

    # 23 [v2]. Сохранить df_alloc в CSV
    table = pa.Table.from_pandas(df_alloc)
    csv.write_csv(table, os.path.join(FOLDER_PATH, "ДБсПризнаками.csv"))
    print(f"[v2] Сохранён CSV: {os.path.join(FOLDER_PATH, 'ДБсПризнаками.csv')}")

    # 24. Обновляем и сохраняем Excel-файл
    def update_and_save_excel(file_path, new_file_path):
        max_attempts = 10
        attempt = 0

        while attempt < max_attempts:
            attempt += 1
            print(f"Попытка {attempt} обновить файл '{os.path.basename(file_path)}'...")

            try:
                excel = win32.Dispatch("Excel.Application")
                excel.DisplayAlerts = False

                try:
                    workbook = excel.Workbooks.Open(file_path)

                    print("Выполняем обновление данных...")
                    workbook.RefreshAll()
                    excel.CalculateUntilAsyncQueriesDone()

                    workbook.SaveAs(file_path)
                    print(f"Файл успешно сохранен с оригинальным именем в '{os.path.dirname(file_path)}'.")

                    workbook.SaveAs(new_file_path)
                    print(f"Файл успешно сохранен как '{os.path.basename(new_file_path)}'.")

                    return True

                except Exception as e:
                    print(f"Ошибка при обновлении или сохранении файла: {e}")
                finally:
                    if 'workbook' in locals():
                        workbook.Close(SaveChanges=False)
                    excel.Quit()

            except Exception as e:
                print(f"Ошибка при работе с Excel: {e}")

            if attempt < max_attempts:
                print(f"Пауза перед следующей попыткой ({attempt + 1}/{max_attempts})...")
                time.sleep(60)

        return False

    print("[v2] Тестовый прогон: обновление боевого Excel-файла пропущено.")

    # 25. Обновить файл "Показы и затраты ВБ_2.0.xlsx"
    try:
        print("Подготовка данных для ДБ завершена.")
        print("Начинаем обновлять файл 'Показы и затраты ВБ_2.0.xlsx'...")
        start_time = time.time()

        file_path_shows_expenses = os.path.join(FOLDER_PATH_FOR_DB, "Показы и затраты ВБ_2.0.xlsx")

        if os.path.exists(file_path_shows_expenses):
            current_month_day = time.strftime("%d.%m")
            new_file_name = f"Показы и затраты ВБ_2.0 {current_month_day}.xlsx"
            new_file_path = os.path.join(FOLDER_PATH_FEATURES, new_file_name)

            dudl_file_path = os.path.join(FOLDER_PATH_DUDL, new_file_name)

            try:
                if os.path.exists(FOLDER_PATH_DUDL):
                    for filename in os.listdir(FOLDER_PATH_DUDL):
                        match = re.match(r"Показы и затраты ВБ_2\.0 (\d{2}\.\d{2})\.xlsx", filename)
                        if match:
                            file_date = match.group(1)
                            if file_date != current_month_day:
                                file_to_delete = os.path.join(FOLDER_PATH_DUDL, filename)
                                os.remove(file_to_delete)
                                print(f"Файл '{filename}' удален из папки '{FOLDER_PATH_DUDL}'.")
            except Exception as delete_error:
                print(f"Ошибка при удалении старых файлов из папки '{FOLDER_PATH_DUDL}': {delete_error}")

            try:
                if os.path.exists(FOLDER_PATH_FEATURES):
                    for filename in os.listdir(FOLDER_PATH_FEATURES):
                        match = re.match(r"Показы и затраты ВБ_2\.0 (\d{2}\.\d{2})\.xlsx", filename)
                        if match:
                            file_date = match.group(1)
                            if file_date != current_month_day:
                                file_to_delete = os.path.join(FOLDER_PATH_FEATURES, filename)
                                os.remove(file_to_delete)
                                print(f"Файл '{filename}' удален из папки '{FOLDER_PATH_FEATURES}'.")
            except Exception as delete_error:
                print(f"Ошибка при удалении старых файлов из папки '{FOLDER_PATH_FEATURES}': {delete_error}")

            success = update_and_save_excel(file_path_shows_expenses, new_file_path)

            if not success:
                while not success:
                    input("Обновить Excel файл не получилось. Закройте все открытые файлы и нажмите любую кнопку для повторной попытки.")
                    success = update_and_save_excel(file_path_shows_expenses, new_file_path)

                print("Файл успешно обновлен после повторной попытки.")

            if success:
                try:
                    shutil.copy(new_file_path, dudl_file_path)
                    print(f"Файл успешно скопирован в папку '{FOLDER_PATH_DUDL}'.")
                except Exception as copy_error:
                    print(f"Ошибка при копировании файла в папку '{FOLDER_PATH_DUDL}': {copy_error}")

            elapsed_time = time.time() - start_time
            print(f"Файл успешно обновлен и сохранен. Время выполнения: {format_elapsed_time(elapsed_time)}")
        else:
            print("Файл 'Показы и затраты ВБ_2.0.xlsx' не найден.")
    except Exception as e:
        print(f"Ошибка при обработке файла 'Показы и затраты ВБ_2.0.xlsx': {e}")

if __name__ == "__main__":
    assemble()
