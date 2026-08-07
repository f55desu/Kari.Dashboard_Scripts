import os
import pandas as pd
import numpy as np
from sqlalchemy import create_engine
import win32com.client as win32
import time  # Для измерения времени выполнения
import shutil
import re
import pyodbc
from glob import glob
# from glob import glob
import gc
from datetime import timedelta
import datetime
# from datetime import datetime
import pyarrow as pa
import pyarrow.csv as csv

# Функция для форматирования времени в часы, минуты и секунды
def format_elapsed_time(seconds):
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{int(hours)} часа(ов) {int(minutes)} минут(ы) {seconds:.2f} секунд"

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

# Функция для форматирования даты в строковый формат 'YYYY-MM-DD'
def format_date_column(df, date_column):
    if date_column in df.columns:
        df[date_column] = pd.to_datetime(df[date_column], errors='coerce').dt.strftime('%Y-%m-%d')
    return df

# Функция для подключения к SQL Server с аутентификацией Windows
def connect_to_sql(server, database):
    connection_string = (
        f"mssql+pyodbc://{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server"
        "&trusted_connection=yes"
    )
    engine = create_engine(connection_string)
    return engine

# Функция для обработки ошибок и замены их на null
def handle_errors(df):
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].apply(lambda x: None if isinstance(x, str) and x.strip() == '' else x)
    return df


FOLDER_PATH = os.path.normpath(r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!")
FOLDER_PATH_FEATURES = r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Дашбоард по рекламным кампаниям"
FOLDER_PATH_FOR_DB= os.path.normpath(r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям")
FOLDER_PATH_DUDL = os.path.normpath(r"\\kari.local\public\all\Агрегаторы\Дашборд реклама WB_OZ")

SQL_SERVER = "cl01sql"
SQL_DATABASE_DBREPORT = "DBReport"
SQL_DATABASE_DBPARTNERS = "DBPartners"

def assemble():
    print("Начинаем собирать Базу Данных...")
    start_all_time = time.time()
    engine = connect_to_sql(SQL_SERVER, SQL_DATABASE_DBPARTNERS)

    # 1. Воронка
    # import pandas as pd
    # import glob
    # import os
    # import datetime
    # import pyodbc

    # Папки
    path_voronka = r"Z:\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!\ВЫГРУЗКА воронка Озон"
    path_zatraty = r"Z:\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!\Затраты\Озон. Затраты из Аналитики"

    # === 1. ВОРОНКА ===
    df_voronka_list = []
    files_voronka = glob(os.path.join(path_voronka, "analytics_report_*.xlsx"))

    for file in files_voronka:
        # достаём дату из имени файла
        fname = os.path.basename(file)
        try:
            report_date = str(datetime.datetime.strptime(fname.split("_")[2], "%Y-%m-%d").date() - timedelta(days=1))
        except Exception:
            raise Exception("Невозможно определить дату из имени файла:", fname)

        df = pd.read_excel(file, engine='calamine')

        # Чистим "Позиция в поиске и каталоге" от запятых
        df["Позиция в поиске и каталоге"] = (
            df["Позиция в поиске и каталоге"]
            .astype(str)
            .str.replace(",", ".", regex=False)
        )

        # df["Позиция в поиске и каталоге"] = pd.to_numeric(df["Позиция в поиске и каталоге"], errors="coerce")
        # print(df.head(20))
        # типы
        df = df.astype({
            "Артикул": "string",
            "Показы, всего": "Int64",
            "Показы на карточке товара": "Int64",
            "Показы в поиске и каталоге": "Int64",
            "Позиция в поиске и каталоге": "float64",
            "В корзину, всего": "Int64",
            "Заказано товаров": "Int64",
            "Отменено товаров": "Int64",
            "Доставлено товаров": "Int64",
            "Возвращено товаров": "Int64",
            "Заказано на сумму": "float64",
            "В корзину из карточки товара": "Int64"
        })

        df["Дата"] = report_date

        df["Выкупили ШТ"] = df["Заказано товаров"] - df["Отменено товаров"] - df["Возвращено товаров"]
        df["Артикул"] = df["Артикул"].astype(str).str.split("-").str[0]

        df_voronka_list.append(df)
    df_voronka = pd.concat(df_voronka_list, ignore_index=True)

    sum_cols_all = [
        "Показы, всего",
        "Показы на карточке товара",
        "Показы в поиске и каталоге",
        "Позиция в поиске и каталоге",          # суммируем, как в вашем примере
        "В корзину, всего",
        "Заказано товаров",
        "Отменено товаров",
        "Доставлено товаров",
        "Возвращено товаров",
        "Заказано на сумму",
        "В корзину из карточки товара",
        "В корзину из поиска или каталога",     # если есть в выгрузке
        "Выкупили ШТ"
    ]
    sum_cols = [c for c in sum_cols_all if c in df_voronka.columns]
    df_voronka.rename(columns={
                            "Артикул": "Артикул",
                            "Продажи, ₽": "Рекламные заказано на сумму",
                            "Показы": "Рекламные показы",
                            "Клики": "Рекламные показы на карточке товара",
                            "Заказы, шт": "Рекламные заказано товаров"
                        }, inplace=True, errors="ignore")
    # 2) безопасно приводим эти метрики к числам (NaN -> 0 перед суммированием)
    for c in sum_cols:
        df_voronka[c] = pd.to_numeric(df_voronka[c], errors="coerce").fillna(0)

    # 3) нечисловые поля, которые логично брать первыми в группе
    first_cols_all = ["Тип товара", "Товары", "Модель", "Ozon ID"]
    first_cols = [c for c in first_cols_all if c in df_voronka.columns]

    # 4) готовим словарь агрегаций
    agg_dict = {c: "sum" for c in sum_cols}
    agg_dict.update({c: "first" for c in first_cols})

    # 5) группируем и получаем одну строку на (Дата, Артикул)
    df_voronka = (
        df_voronka
        .groupby(["Дата", "Артикул"], as_index=False)
        .agg(agg_dict)
    )


    # 2. Получить данные из файла "Справочник.xlsx"
    try:
        print("Начинаем получать данные для Справочника...")
        start_time = time.time()  # Запускаем таймер
        file_path_reference = os.path.join(FOLDER_PATH, "Справочник.xlsx")

        if os.path.exists(file_path_reference):
            # Список столбцов, которые нужно взять из файла
            columns_to_read = [
                "Артикул", "Артикул OZ", "Наименование", "Коллекция",
                "Бренд", "Размер", "Сезон", "Направление", "Розничный отдел",
                "Модель", "Группа", "Бизнес-группа", "Техсегмент",
                "Байер", "Две последние коллекции", "Основной артикул", "Ответственный за группу", "Себестоимость с НДС",
                "Процент выкупа", "НДС", "Группа для отчетов"
            ]

            # Типы данных для столбцов
            column_dtypes = {
                "Артикул": str,
                "Артикул OZ": str,
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
                "Ответственный за группу": str,
                "Себестоимость с НДС": float,
                "Процент выкупа": float,
                "НДС": int,
                "Группа для отчетов": str
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



    # 3. Создание таблицы "ВсегоРазмеров"
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



    # 4. Получить данные таблицы с SQL (РазмерыНаАгрегаторе)
    try:
        print("Начинаем получать данные для РазмеровНаАгрегаторе...")
        start_time = time.time()  # Запускаем таймер
        query_sizes = f"""
            SELECT a.[dt] AS [Дата], a.[itemid] AS [Артикул], COUNT(DISTINCT(a.[INVENTSIZEID])) AS [Колво размеров]
            FROM [DBPartners].[dbo].[WblmRepGetStockOzon] a
            WHERE [dt] >= '{str(df_voronka['Дата'].min())}'
            GROUP BY [dt], [itemid]
        """
        df_sizes = pd.read_sql(query_sizes, engine)
        df_sizes = format_date_column(df_sizes, 'Дата')

        # Вывод первых 5 строк
        print("Первые 5 строк таблицы РазмерыНаАгрегаторе:")
        print(df_sizes.head())

        elapsed_time = time.time() - start_time  # Вычисляем затраченное время
        print(f"Данные для РазмеровНаАгрегаторе успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при получении данных для РазмеровНаАгрегаторе: {e}")



    # 5. Связать "РазмерыНаАгрегаторе" с "ВсегоРазмеров"
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


    # 6. Получить данные таблицы с SQL (Остатки)
    try:
        print("Начинаем получать данные для Остатков...")
        start_time = time.time()  # Запускаем таймер
        query_stock = f"""
            SELECT a.[dt] AS [Дата], a.[itemid] AS [Артикул], SUM(a.[free_to_sell_amount]) AS [Остаток Агрегатора]
            FROM [DBPartners].[dbo].[WblmRepGetStockOzon] a
            WHERE [dt] >= '{str(df_voronka['Дата'].min())}'
            GROUP BY [dt], [itemid]
        """
        df_stock = pd.read_sql(query_stock, engine)
        df_stock = format_date_column(df_stock, 'Дата')

        # Вывод первых 5 строк
        print("Первые 5 строк таблицы Остатков:")
        print(df_stock.head())

        elapsed_time = time.time() - start_time  # Вычисляем затраченное время
        print(f"Данные для Остатков успешно сохранены. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при получении данных для Остатков: {e}")
    # 7. Связать "Остатки" с "Дистрибуция"
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

    del df_stock



    # 8. Загрузка данных о затратах
    # === 2. ЗАТРАТЫ ===
    df_zatraty_list = []
    files_zatraty = glob(os.path.join(path_zatraty, "*.csv"))

    for f in files_zatraty:
        # дата из имени файла
        fname = os.path.basename(f).replace(".csv", "")
        file_date = pd.to_datetime(fname, dayfirst=True, errors="coerce")

        # читаем csv
        df = pd.read_csv(f, sep=";", skiprows=2)

        # чистим названия колонок
        df.columns = [c.replace(".csv","") if ".csv" in c else c for c in df.columns]

        # вставляем дату из имени файла
        df["Дата"] = file_date

        # переименования
        df = df.rename(columns={
            "Тип продвижения": "ТипАктивности",
            "Расход, ₽, с НДС": "Расход, ₽"
        })

        # Числовые колонки в CSV могут прийти как строки с запятой-разделителем ("0,00",
        # "1 234,56 ₽", "8103,00") или как int64 (когда все значения целые). После
        # pd.concat смешанные типы схлопываются в object, и groupby(...).sum() склеивает
        # строки вместо числового суммирования.
        # Начиная с выгрузок 24.04.2026 Ozon стал возвращать с запятой ВСЕ числовые
        # колонки (не только деньги), включая Показы / Клики / Заказы, шт. Поэтому
        # чистим весь набор метрик, которые потом уйдут в merge_voronka_costs_preserve_impressions
        # и build_funnel_wide.
        _numeric_cols = (
            "Расход, ₽", "Продажи, ₽", "Показы", "Клики", "Заказы, шт",
            "В корзину", "ДРР, %", "CTR, %", "Конверсия в корзину, %",
            "Затраты на заказ, ₽", "Стоимость клика, ₽",
        )
        for _num_col in _numeric_cols:
            if _num_col in df.columns:
                df[_num_col] = pd.to_numeric(
                    df[_num_col].astype(str)
                        .str.replace('\u00A0', '', regex=False)
                        .str.replace(' ', '', regex=False)
                        .str.replace('₽', '', regex=False)
                        .str.replace('%', '', regex=False)
                        .str.replace(',', '.', regex=False),
                    errors='coerce'
                )

        df_zatraty_list.append(df)

    df_zatraty = pd.concat(df_zatraty_list, ignore_index=True)
    # [v2] Раньше "Оплата за заказ: выбранные товары" склеивалось с "Оплата за клик".
    # Теперь сохраняем тип как есть — он соответствует одной из 4-х новых категорий.
    # df_zatraty.loc[df_zatraty['ТипАктивности'] == 'Оплата за заказ: выбранные товары', 'ТипАктивности'] = 'Оплата за клик'

    # 10. Получить данные таблицы с SQL (Цены)
    # === 3. SQL ЦЕНЫ ===
    sql = f"""
    select 'OZ' as AGREGATOR, DT, ITEMID, PRICE
    from [DBPartners].[dbo].[WblmRepPriceDiscountOzReport]
    where dt >= '{str(df_voronka['Дата'].min())}'
    """
    df_prices = pd.read_sql(sql, engine)

    df_prices = df_prices.groupby(["DT", "ITEMID"], as_index=False).agg({"PRICE": "max"})
    df_prices = df_prices.rename(columns={"DT": "Дата", "ITEMID": "Артикул", "PRICE": "Цена"})



    # 11. Собираем полный датафрейм воронки
    df_reference = df_reference.drop_duplicates(subset=["Артикул"])
    df_reference = df_reference[["Артикул", "Бизнес-группа", "Направление", "Розничный отдел", "Группа", "Модель", "Бренд", "Коллекция", "Сезон", "Себестоимость с НДС", "Процент выкупа", "Две последние коллекции", 'Артикул OZ', 'Наименование', 'Техсегмент', 'Байер', 'Основной артикул', 'НДС', 'Ответственный за группу', 'Группа для отчетов']]
    df_reference

    # import pandas as pd
    # import numpy as np

    NBSP = '\u00A0'

    # ---- утилиты ----
    def norm_date(s: pd.Series) -> pd.Series:
        d = pd.to_datetime(s, errors='coerce')
        try:
            d = d.dt.tz_localize(None)
        except Exception:
            pass
        return d.dt.normalize()

    def norm_code(s: pd.Series) -> pd.Series:
        return (s.astype(str)
                .str.replace(NBSP, '', regex=False)
                .str.strip()
                .str.upper())

    def norm_code_alnum(s: pd.Series) -> pd.Series:
        # только A–Z и 0–9: удобно для OZ/SKU
        return norm_code(s).str.replace(r'[^0-9A-Z]+', '', regex=True)

    def to_money(s: pd.Series) -> pd.Series:
        # «1 234,56 ₽», «1.234,56», «1234,56» → 1234.56
        ss = (s.astype(str)
                .str.replace(NBSP, '', regex=False)
                .str.replace(' ',  '', regex=False)
                .str.replace('₽',  '', regex=False))
        ss = ss.str.replace(r'(?<=\d)\.(?=\d{3}(?:\D|$))', '', regex=True)  # 1.234,56 → 1234,56
        ss = ss.str.replace(',', '.', regex=False)
        ss = ss.str.replace(r'[^0-9\.\-\+eE]', '', regex=True)
        return pd.to_numeric(ss, errors='coerce')

    # ---- основная функция: тип берём ТОЛЬКО из затрат, без изменений ----
    def merge_voronka_costs_preserve_impressions(
        df_voronka: pd.DataFrame,
        df_costs: pd.DataFrame,
        df_prices: pd.DataFrame,
        *,
        left_key_candidates=('Ozon ID','OZON ID','OZON_ID','Артикул OZ','Артикул'),
        right_key='SKU',
        spend_candidates=('Расход, ₽','Расход, руб','Расход, Р','Расход'),
        type_col_candidates=('ТипАктивности','Тип активности','Раздел', 'Тип продвижения'),
        preserve_cols=('Показы, всего',),   # инварианты по левым метрикам
        add_tail=True                        # добавлять «хвост» неприсоединившихся затрат
    ) -> pd.DataFrame:

        # --- копии + даты ---
        df_v = df_voronka.copy()
        df_z = df_costs.copy()
        df_p = df_prices.copy()

        for dframe in (df_v, df_z) + ((df_p,) if df_p is not None else ()):
            if dframe is not None and 'Дата' in dframe.columns:
                dframe['Дата'] = norm_date(dframe['Дата'])

        # --- контроль инвариантов для левых метрик ---
        _num = lambda s: pd.to_numeric(s, errors='coerce')
        baseline = {c: (_num(df_v[c]).sum() if c in df_v.columns else None) for c in preserve_cols}

        # --- выбор ключей слева/справа ---
        left_key = next((c for c in left_key_candidates if c in df_v.columns), None)
        if left_key is None:
            raise KeyError(f"Во воронке нет ни одного ключа из {left_key_candidates}")

        if right_key not in df_z.columns:
            raise KeyError(f"В затратах нет колонки {right_key}")

        spend_col = next((c for c in spend_candidates if c in df_z.columns), None)
        if spend_col is None:
            raise KeyError(f"В затратах нет денежной колонки из {spend_candidates}")

        # --- подготовка затрат: ключ и деньги ---
        df_z[right_key] = norm_code(df_z[right_key])
        df_z['__KEY__'] = norm_code_alnum(df_z[right_key])
        df_z[spend_col] = to_money(df_z[spend_col]).astype('float64')

        # правый столбец типа (берём как есть, БЕЗ нормализаций)
        right_type_col = next((c for c in type_col_candidates if c in df_z.columns), None)

        # дополнительные правые метрики, если есть
        extra_metrics = [c for c in ['Показы','Клики','Заказы, шт','Продажи, ₽'] if c in df_z.columns]
        z_metrics = [spend_col] + extra_metrics

        # агрегирование по (Дата, __KEY__) + тип из затрат как first (как есть)
        if right_type_col is not None:
            agg_dict = {m: 'sum' for m in z_metrics}
            agg_dict[right_type_col] = 'first'  # важное: тип из затрат, без изменений
            z_agg_full = (df_z.groupby(['Дата','__KEY__'], as_index=False)
                            .agg(agg_dict))
            # переименуем правый тип в целевое имя
            if right_type_col != 'ТипАктивности':
                z_agg_full = z_agg_full.rename(columns={right_type_col: 'ТипАктивности'})
        else:
            z_agg_full = (df_z.groupby(['Дата','__KEY__'], as_index=False)[z_metrics]
                            .sum(min_count=1))

        total_costs = z_agg_full[spend_col].sum()

        # --- левый нормализованный ключ ---
        df_v['__KEY__'] = norm_code_alnum(norm_code(df_v[left_key]))
        df_v['IS_TAIL'] = 0

        # --- основной LEFT m:1 join ---
        df = df_v.merge(z_agg_full, on=['Дата','__KEY__'], how='left', validate='m:1')

        # если слева уже был 'ТипАктивности', мы его полностью ЗАМЕНЯЕМ на тип из затрат
        if 'ТипАктивности' in df.columns and right_type_col is None:
            # в затратах типа нет — тогда оставим как было слева
            pass
        else:
            # гарантируем, что колонка называется ровно 'ТипАктивности' и она пришла из затрат
            if right_type_col is not None and 'ТипАктивности' not in df.columns:
                # если правый тип имел другое имя и мы переименовали выше — уже ок
                # здесь просто убедимся, что колонка есть
                pass
            # Если и слева, и справа есть 'ТипАктивности' (редкий кейс),
            # берем ровно правую версию: удалим левую и переименуем правую.
            if right_type_col is not None and 'ТипАктивности_x' in df.columns and 'ТипАктивности_y' in df.columns:
                # _x — из воронки, _y — из затрат; оставляем _y
                df.drop(columns=['ТипАктивности_x'], inplace=True, errors='ignore')
                df.rename(columns={'ТипАктивности_y': 'ТипАктивности'}, inplace=True)

        # --- инварианты по левым метрикам ---
        for c in preserve_cols:
            if c in df.columns and baseline[c] is not None:
                after = _num(df[c]).sum()
                msg = "[OK]" if np.isclose(after, baseline[c], rtol=1e-9, atol=1e-6) else "[WARN]"
                print(f"{msg} Инвариант '{c}': до={baseline[c]:,.2f} | после={after:,.2f}")

        # --- добавляем «хвост» из затрат (без изменений типов!) ---
        tail_rows = 0
        if add_tail:
            left_keys = df[['Дата','__KEY__']].drop_duplicates()
            right_cols_only = [c for c in z_agg_full.columns if c not in ('Дата','__KEY__')]
            miss = (z_agg_full.merge(left_keys, on=['Дата','__KEY__'], how='left', indicator=True)
                            .loc[lambda x: x['_merge']=='left_only', ['Дата','__KEY__'] + right_cols_only])
            if not miss.empty:
                tail = miss.copy()
                # каркас: заполним недостающие левые поля NaN
                for col in df.columns:
                    if col not in tail.columns:
                        tail[col] = np.nan
                tail['IS_TAIL'] = 1
                tail[left_key] = tail['__KEY__']
                tail = tail[df.columns]
                df = pd.concat([df, tail], ignore_index=True)
                tail_rows = len(tail)

        # --- прайс (m:1), если нужен ---
        if df_p is not None and {'Дата','Артикул'}.issubset(df.columns) and {'Дата','Артикул'}.issubset(df_p.columns):
            df_p['Артикул'] = norm_code(df_p['Артикул'])
            df = df.merge(df_p, on=['Дата','Артикул'], how='left', validate='m:1')

        # --- финальные сверки по расходам ---
        final_costs = pd.to_numeric(df[spend_col], errors='coerce').fillna(0).sum()
        delta = final_costs - total_costs
        print(f"[CHECK] Расходы: в файле = {total_costs:,.2f} ₽ | в результате = {final_costs:,.2f} ₽ | Δ = {delta:,.2f} ₽")
        if tail_rows:
            print(f"[INFO] Добавлен хвост (IS_TAIL=1): {tail_rows} строк")

        # --- уборка служебных ---
        df.drop(columns=['__KEY__','IS_TAIL'], inplace=True, errors='ignore')

        # --- опциональное переименование правых метрик в «рекламные …» ---
        df.rename(columns={
            "Продажи, ₽": "Рекламные заказано на сумму",
            "Показы": "Рекламные показы",
            "Клики": "Рекламные показы на карточке товара",
            "Заказы, шт": "Рекламные заказано товаров"
        }, inplace=True, errors="ignore")

        return df


    # 12. Получаем дополнительную информацию по затратам
    # Путь до папки
    folder_path_weeks = os.path.join(FOLDER_PATH, "Затраты", "Озон. Затраты из Аналитики New Format")

    # Собираем все .xlsx файлы
    files = glob(os.path.join(folder_path_weeks, "*.xlsx"))

    df_list = []
    df_list_union = []
    for file in files:
        # --- достаём дату из названия файла ---
        filename = os.path.basename(file)  # например: "Аналитика продвижения_16.09.2025.xlsx"
        date_str = filename.split("_")[-1].replace(".xlsx", "")  # "16.09.2025"
        date_parsed = (pd.to_datetime(date_str, format="%d.%m.%Y") - timedelta(days=1)).strftime("%Y-%m-%d")

        # читаем, пропуская первую строку
        df_tmp = pd.read_excel(file, engine='calamine', skiprows=1)
        df_tmp_union = pd.read_excel(file, sheet_name='Union', engine='calamine', skiprows=1)

        # оставляем только нужные колонки
        cols_keep = ["SKU", "ID кампании", "Инструмент", "Место размещения"]
        cols_keep_union = ["SKU в продвижении", "SKU из объединенной карточки", "Продажи, ₽", "Заказы, шт"]
        df_tmp = df_tmp[cols_keep]
        df_tmp_union = df_tmp_union[cols_keep_union]

        # добавляем колонку "Дата"
        df_tmp["Дата"] = date_parsed
        df_tmp_union["Дата"] = date_parsed

        df_list.append(df_tmp)
        df_list_union.append(df_tmp_union)

    # объединяем все файлы
    df_all = pd.concat(df_list, ignore_index=True)
    df_all_union = pd.concat(df_list_union, ignore_index=True)
    df_all.rename(columns={'SKU': "Артикул OZ"},inplace=True)
    df_all_union.rename(columns={'SKU в продвижении': "Артикул OZ"},inplace=True)

    # Функция для строгой нормализации
    def normalize_key_column(s: pd.Series, col_name: str) -> pd.Series:
        """Строгая нормализация с логированием"""
        original = s.copy()
        
        result = (s.astype(str)
                .str.replace('\u00A0', '', regex=False)  # неразрывный пробел
                .str.replace(' ', '', regex=False)        # обычный пробел
                .str.replace('\t', '', regex=False)       # табуляция
                .str.strip()
                .str.upper())
        
        changed = (original.astype(str) != result).sum()
        print(f"  Колонка '{col_name}': изменено {changed} значений")
        
        return result

    # Применяем ко ВСЕМ ключевым колонкам
    print("Нормализация df_zatraty:")
    df_zatraty['Дата'] = pd.to_datetime(df_zatraty['Дата'], errors='coerce').dt.normalize()
    df_zatraty['SKU'] = normalize_key_column(df_zatraty['SKU'], 'SKU в затратах')
    df_zatraty['ID кампании'] = normalize_key_column(df_zatraty['ID кампании'], 'ID в затратах')

    print("\nНормализация df_all:")
    df_all['Дата'] = pd.to_datetime(df_all['Дата'], errors='coerce').dt.normalize()
    df_all['Артикул OZ'] = normalize_key_column(df_all['Артикул OZ'], 'SKU в df_all')
    df_all['ID кампании'] = normalize_key_column(df_all['ID кампании'], 'ID в df_all')

    # # Нормализуем ключи (точно так же, как в merge_voronka_costs_preserve_impressions)
    # df_all['Дата'] = pd.to_datetime(df_all['Дата'], errors='coerce').dt.normalize()
    df_all['Артикул OZ'] = (df_all['Артикул OZ'].astype(str)
                    .str.replace('\u00A0', '', regex=False)
                    .str.strip()
                    .str.upper())
    df_all_union['Артикул OZ'] = (df_all_union['Артикул OZ'].astype(str)
                    .str.replace('\u00A0', '', regex=False)
                    .str.strip()
                    .str.upper())
    df_zatraty['SKU'] = (df_zatraty['SKU'].astype(str)
                    .str.replace('\u00A0', '', regex=False)
                    .str.strip()
                    .str.upper())

    # # ===== ШАГ 2: Обогащаем df_zatraty информацией из df_all =====
    # # ВАЖНО: делаем это ДО вызова merge_voronka_costs_preserve_impressions

    # Нормализуем дату в df_zatraty
    df_zatraty['Дата'] = pd.to_datetime(df_zatraty['Дата'], errors='coerce').dt.normalize()
    df_all['Дата'] = pd.to_datetime(df_all['Дата'], errors='coerce').dt.normalize()

    # Присоединяем инструменты по (SKU, ID кампании, Дата)
    df_zatraty_enriched = df_zatraty.rename(columns={'SKU': "Артикул OZ"}).merge(
        df_all[['Дата', 'Артикул OZ', 'ID кампании', 'Инструмент', 'Место размещения']],
        on=['Дата', 'Артикул OZ', 'ID кампании'],
        how='left',
        validate='m:1'  # каждая строка затрат должна найти максимум одну строку в df_all
    )

    # ===== ШАГ 3 [v2]: Создаём ТипАктивности В ЗАТРАТАХ по новым 4 категориям =====
    # Источники классификации:
    #   1. Инструмент = "Оплата за заказ: выбранные товары"      → "Оплата за заказ: выбранные товары"
    #   2. Инструмент = "Оплата за клик" + Место = "Поиск"        → "Оплата за клик: поиск"
    #   3. Инструмент = "Оплата за клик" + любое другое/NaN       → "Оплата за клик: поиск и рекомендации"
    #   4. Расход = 0/NaN                                          → "Органика" (последним, перебивает платные с нулём)
    df_zatraty_enriched['ТипАктивности'] = np.nan

    # 1) Оплата за заказ: выбранные товары
    mask_zakaz = (df_zatraty_enriched['Инструмент'] == 'Оплата за заказ: выбранные товары')
    df_zatraty_enriched.loc[mask_zakaz, 'ТипАктивности'] = 'Оплата за заказ: выбранные товары'

    # 2) Оплата за клик: поиск
    mask_klik_search = (
        (df_zatraty_enriched['Инструмент'] == 'Оплата за клик') &
        (df_zatraty_enriched['Место размещения'] == 'Поиск')
    )
    df_zatraty_enriched.loc[mask_klik_search, 'ТипАктивности'] = 'Оплата за клик: поиск'

    # 3) Оплата за клик: поиск и рекомендации (включая случаи с пустым/иным Местом размещения)
    mask_klik_recs = (
        (df_zatraty_enriched['Инструмент'] == 'Оплата за клик') & ~mask_klik_search
    )
    df_zatraty_enriched.loc[mask_klik_recs, 'ТипАктивности'] = 'Оплата за клик: поиск и рекомендации'

    # 4) Органика — нулевой/отсутствующий расход (применяется ПОСЛЕДНИМ, перебивает платные с нулём)
    spend_col = next((c for c in ['Расход, ₽','Расход, руб','Расход, Р','Расход']
                    if c in df_zatraty_enriched.columns), None)
    if spend_col:
        df_zatraty_enriched[spend_col] = pd.to_numeric(
            df_zatraty_enriched[spend_col].astype(str)
            .str.replace('\u00A0', '', regex=False)
            .str.replace(' ', '', regex=False)
            .str.replace('₽', '', regex=False)
            .str.replace(',', '.', regex=False),
            errors='coerce'
        )
        mask_organic = (df_zatraty_enriched[spend_col] == 0) | (df_zatraty_enriched[spend_col].isna())
        df_zatraty_enriched.loc[mask_organic, 'ТипАктивности'] = 'Органика'


    print("=" * 80)
    print("ПРАВИЛЬНЫЙ МЕРЖ БЕЗ ЛИШНЕГО МАППИНГА")
    print("=" * 80)

    # ===== ШАГ 1: Подготовка воронки =====
    print("\nШАГ 1: Подготовка воронки")
    print("-" * 80)

    df_voronka_clean = df_voronka.copy()

    # Убедимся, что Ozon ID есть
    if 'Ozon ID' not in df_voronka_clean.columns:
        raise ValueError("❌ В воронке нет колонки 'Ozon ID'!")

    # Нормализуем Ozon ID - приводим к строке и убираем лишнее
    df_voronka_clean['Ozon ID'] = (
        df_voronka_clean['Ozon ID']
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r'\.0$', '', regex=True)  # убираем .0 если есть
    )

    print(f"✓ Размер воронки: {len(df_voronka_clean):,} строк")
    print(f"✓ Заполненность Ozon ID: {df_voronka_clean['Ozon ID'].notna().sum():,} ({100*df_voronka_clean['Ozon ID'].notna().sum()/len(df_voronka_clean):.1f}%)")
    print(f"✓ Примеры Ozon ID: {df_voronka_clean['Ozon ID'].head(5).tolist()}")

    # ===== ШАГ 2: Подготовка затрат =====
    print("\nШАГ 2: Подготовка затрат")
    print("-" * 80)

    df_zatraty_clean = df_zatraty_enriched.copy()

    # Нормализуем Артикул OZ
    df_zatraty_clean['Артикул OZ'] = (
        df_zatraty_clean['Артикул OZ']
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r'\.0$', '', regex=True)
    )

    print(f"✓ Размер затрат: {len(df_zatraty_clean):,} строк")
    print(f"✓ Уникальных Артикул OZ: {df_zatraty_clean['Артикул OZ'].nunique():,}")
    print(f"✓ Примеры Артикул OZ: {df_zatraty_clean['Артикул OZ'].head(5).tolist()}")

    # ===== ШАГ 3: Проверка пересечения =====
    print("\nШАГ 3: Проверка пересечения ключей")
    print("-" * 80)

    voronka_ids = set(df_voronka_clean['Ozon ID'].dropna())
    zatraty_ids = set(df_zatraty_clean['Артикул OZ'].dropna())

    intersection = voronka_ids & zatraty_ids

    print(f"Уникальных ID в воронке: {len(voronka_ids):,}")
    print(f"Уникальных ID в затратах: {len(zatraty_ids):,}")
    print(f"Пересечение: {len(intersection):,} ({100*len(intersection)/len(voronka_ids) if voronka_ids else 0:.1f}% от воронки)")

    if len(intersection) == 0:
        print("\n⚠️ КРИТИЧЕСКАЯ ПРОБЛЕМА: НЕТ ПЕРЕСЕЧЕНИЙ!")
        print("Возможные причины:")
        print("  1. Разные периоды данных в воронке и затратах")
        print("  2. Разные площадки (один файл WB, другой Ozon)")
        print("  3. Проблема с ID - один Ozon ID, другой SKU")
        
        print("\nПримеры ID из воронки:")
        print(list(voronka_ids)[:10])
        print("\nПримеры ID из затрат:")
        print(list(zatraty_ids)[:10])
        
        raise ValueError("Невозможно продолжить без пересечений!")

    # ===== ШАГ 4: Мерж через функцию =====
    print("\nШАГ 4: Финальный мерж")
    print("-" * 80)

    # Переименуем Ozon ID -> Артикул OZ для единообразия
    df_voronka_clean = df_voronka_clean.rename(columns={'Ozon ID': 'Артикул OZ'})

    df_prices_exists = 'df_prices' in globals()

    df_funnel_final = merge_voronka_costs_preserve_impressions(
        df_voronka=df_voronka_clean,
        df_costs=df_zatraty_clean,
        df_prices=df_prices,# if df_prices_exists else None,
        preserve_cols=('Показы, всего',),
        left_key_candidates=('Артикул OZ',),    # ← теперь одинаковое название!
        right_key='Артикул OZ',                 # ← и тут тоже!
        add_tail=False,
        type_col_candidates=('ТипАктивности',)
    )

    print("\n✅ Мерж завершён!")

    # ===== ШАГ 5: Заполняем органику =====
    print("\nШАГ 5: Заполнение органического трафика")
    print("-" * 80)

    spend_col = next((c for c in ['Расход, ₽','Расход, руб'] if c in df_funnel_final.columns), None)
    shows_col = next((c for c in ['Показы, всего', 'Показы'] if c in df_funnel_final.columns), None)

    if spend_col and shows_col:
        # Нормализуем колонки
        df_funnel_final[spend_col] = pd.to_numeric(df_funnel_final[spend_col], errors='coerce').fillna(0)
        df_funnel_final[shows_col] = pd.to_numeric(df_funnel_final[shows_col], errors='coerce').fillna(0)
        
        # Маска для органики: нет типа + расход=0 + есть показы
        mask_organic = (
            df_funnel_final['ТипАктивности'].isna() &
            (df_funnel_final[spend_col] == 0) &
            (df_funnel_final[shows_col] > 0)
        )
        
        organic_count_before = mask_organic.sum()
        df_funnel_final.loc[mask_organic, 'ТипАктивности'] = 'Органика'
        
        print(f"✓ Заполнено 'Органика' для {organic_count_before:,} строк")

    # ===== ШАГ 6: Финальная статистика =====
    print("\n" + "=" * 80)
    print("ФИНАЛЬНАЯ СТАТИСТИКА")
    print("=" * 80)

    print(f"\nРазмер результата: {len(df_funnel_final):,} строк")

    print("\nРаспределение ТипАктивности:")
    type_counts = df_funnel_final['ТипАктивности'].value_counts(dropna=False)
    print(type_counts)

    nan_count = df_funnel_final['ТипАктивности'].isna().sum()
    nan_pct = 100 * nan_count / len(df_funnel_final)
    print(f"\nСтрок с NaN: {nan_count:,} ({nan_pct:.1f}%)")

    if nan_pct < 20:
        print("\n✅ УСПЕХ! Доля NaN < 20%")
    elif nan_pct < 50:
        print("\n⚠️ ЧАСТИЧНО РЕШЕНО: Доля NaN {:.1f}%".format(nan_pct))
    else:
        print("\n❌ ПРОБЛЕМА: Доля NaN всё ещё {:.1f}%".format(nan_pct))

    # Проверка расходов
    if spend_col:
        total_spend = df_funnel_final[spend_col].sum()
        original_spend = df_zatraty_clean[spend_col].sum()
        print(f"\nСумма расходов:")
        print(f"  - В затратах: {original_spend:,.2f} ₽")
        print(f"  - В результате: {total_spend:,.2f} ₽")
        print(f"  - Разница: {total_spend - original_spend:,.2f} ₽")

    print("\n" + "=" * 80)
    print("ГОТОВО! Результат в переменной df_funnel_final")
    print("=" * 80)

    # Сохраняем в df_funnel для удобства
    df_funnel = df_funnel_final
    print("\n✓ Также сохранено в df_funnel")

    df_union_reference = pd.merge(df_reference[['Артикул', 'Артикул OZ']], df_all_union, on="Артикул OZ", how='right')
    # df_union_reference[df_union_reference['Артикул'].isna()]
    df_union_reference

    df_union_reference['Продажи, ₽'] = pd.to_numeric(df_union_reference['Продажи, ₽'], errors='coerce')
    df_union_reference['Заказы, шт'] = pd.to_numeric(df_union_reference['Заказы, шт'], errors='coerce')

    df_union_agg = (
        df_union_reference
        .groupby(['Артикул OZ', 'Дата'], as_index=False)
        .agg({
            'Артикул': 'first',
            'SKU из объединенной карточки': 'first',
            'Продажи, ₽': 'sum',
            'Заказы, шт': 'sum',
        })
    )

    df_all=df_all.drop_duplicates(subset=['Дата','Артикул OZ'])



    # 13. Связать "Воронка" с "Справочник" БЕЗ дублирования
    # import numpy as np
    # --- утилита: выбрать колонку расхода (на всякий случай) ---
    SPEND_CANDIDATES = ['Расход, ₽', 'Расход, руб', 'Расход, Р', 'Расход']
    def _pick_spend_col(df: pd.DataFrame) -> str:
        for c in SPEND_CANDIDATES:
            if c in df.columns:
                return c
        raise KeyError(f"Не найдена колонка расхода среди: {SPEND_CANDIDATES}")
    # 10. Связать "Воронка" с "Справочник" БЕЗ дублирования
    try:
        print("Начинаем создавать таблицу ВоронкаСправочник...")
        start_time = time.time()

        # --- 0) Валидация входа ---
        required_columns = ["Дата", "Артикул"]
        for col in required_columns:
            if col not in df_funnel.columns:
                raise ValueError(f"Отсутствует столбец '{col}' в df_funnel.")

        # --- 1) Нормализуем ключ (точно так же в обеих таблицах) ---
        df_funnel = df_funnel.copy()
        df_reference = df_reference.copy()

        df_funnel["Артикул"] = (df_funnel["Артикул"].fillna('')
                                                .astype(str).str.strip().str.upper()
                                                .str[:8])
        df_reference["Артикул"] = (df_reference["Артикул"].fillna('')
                                                        .astype(str).str.strip().str.upper()
                                                        .str[:8])

        # Приведём дату в единый формат (без времени)
        # df_funnel["Дата"] = pd.to_datetime(df_funnel["Дата"], errors="coerce").dt.normalize()

        # --- 2) Оставляем только нужные поля справочника ---
        reference_columns = [
            "Артикул", "Артикул OZ", "Наименование", "Коллекция", "Бренд", "Сезон", "Направление",
            "Розничный отдел", "Модель", "Группа", "Бизнес-группа", "Техсегмент",
            "Байер", "Две последние коллекции", "Основной артикул", "Себестоимость с НДС",
            "Процент выкупа", "НДС", "Ответственный за группу", "Группа для отчетов"
        ]
        # оставим только реально существующие колонки
        reference_columns = [c for c in reference_columns if c in df_reference.columns]
        df_reference_filtered = df_reference[["Артикул"] + [c for c in reference_columns if c != "Артикул"]].copy()

        # --- 3) Делаем справочник УНИКАЛЬНЫМ по Артикулу (one-row-per-Артикул) ---
        # Если есть дубликаты одного артикула — берём первую строку (можно заменить на приоритетное правило)
        dups = df_reference_filtered["Артикул"].duplicated(keep=False).sum()
        if dups:
            print(f"[INFO] В справочнике обнаружены дубликаты по 'Артикул' (после .str[:8]): {dups} строк.")
        ref_unique = (df_reference_filtered
                    .sort_values(["Артикул"])
                    .drop_duplicates(subset=["Артикул"], keep="first")
                    .reset_index(drop=True))

        # sanity: строго уникально
        assert not ref_unique["Артикул"].duplicated().any(), "ref_unique всё ещё содержит дубликаты Артикул"

        # --- 4) Контроль инвариантов расхода ДО merge ---
        try:
            spend_col = _pick_spend_col(df_funnel)
        except KeyError:
            spend_col = None

        if spend_col:
            before_total = pd.to_numeric(df_funnel[spend_col], errors='coerce').sum()
            before_by_date = (df_funnel.groupby("Дата", as_index=False)[spend_col]
                                    .sum(min_count=1)
                                    .rename(columns={spend_col: "Расход_до"}))
        else:
            print("[INFO] В df_funnel нет колонки расхода — инварианты по расходу не проверяем.")

        del ref_unique['Артикул OZ']
        del ref_unique['Модель']
        # --- 5) LEFT-merge строго m:1 (исключает размножение строк) ---
        df_funnel_reference = pd.merge(
            df_funnel,
            ref_unique,
            on="Артикул",
            how="left",
            validate="m:1",     # если справа снова появятся дубликаты — упадёт сразу
            indicator=False
        )

        # Ничего НЕ дропаем из df_funnel_reference! (drop_duplicates ломает суммы)

        # --- 6) Контроль ПОСЛЕ merge (сумма не должна измениться ни по датам, ни итого) ---
        if spend_col:
            after_total = pd.to_numeric(df_funnel_reference[spend_col], errors='coerce').sum()
            after_by_date = (df_funnel_reference.groupby("Дата", as_index=False)[spend_col]
                                            .sum(min_count=1)
                                            .rename(columns={spend_col: "Расход_после"}))
            check = before_by_date.merge(after_by_date, on="Дата", how="outer").fillna(0)
            drift = check.loc[~np.isclose(check["Расход_до"], check["Расход_после"], rtol=1e-9, atol=1e-6)]

            print(f"[CHECK] Общая сумма расхода: до={before_total:,.2f} | после={after_total:,.2f}")
            if len(drift):
                print("[WARN] Обнаружены расхождения по датам (первые 10):")
                print(drift.head(10))
                # Если тут расхождения — значит есть иная проблема (например, обрезка .str[:8] склеила разные артикула)
        else:
            after_total = None

        # --- 7) Формат вывода даты (если нужен именно date без времени) ---
        # df_funnel_reference["Дата"] = pd.to_datetime(df_funnel_reference["Дата"]).dt.date

        # Итог
        print("Первые 5 строк таблицы ВоронкаСправочник:")
        print(df_funnel_reference.head())

        elapsed_time = time.time() - start_time
        print(f"Таблица ВоронкаСправочник успешно создана. Время выполнения: {elapsed_time:.2f} c.")

    except Exception as e:
        print(f"Ошибка при создании таблицы ВоронкаСправочник: {e}")

    df_funnel_reference['Расход, ₽'] = pd.to_numeric(
        df_funnel_reference['Расход, ₽']
            .astype(str)
            .str.replace('\u00A0', '', regex=False)  # NBSP
            .str.replace(' ',      '', regex=False)  # обычные пробелы
            .str.replace('₽',      '', regex=False)
            .str.replace(',',      '.', regex=False),  # ВАЖНО: str.replace, не replace
        errors='coerce'
    )
    df_union_agg['Дата'] = pd.to_datetime(df_union_agg['Дата'])
    df_funnel_reference = df_funnel_reference.merge(
        df_union_agg,
        on=["Дата", "Артикул OZ"],
        how="left",
        validate="m:1"
    )
    df_funnel_reference.drop(columns=['Артикул_y'], inplace=True)
    df_funnel_reference.rename(columns={'Артикул_x':'Артикул'}, inplace=True)

    df_funnel_reference.rename(columns={'Продажи, ₽':'Ассоциированные заказы, руб', 'Заказы, шт':'Ассоциированные заказы, шт'}, inplace=True)




    # 14. Создание широкой таблицы по Типу активности
    funnel_columns = ['Дата', 'Артикул', 'Артикул OZ', 'ТипАктивности', 'Показы, всего', 'Показы на карточке товара',
        'Показы в поиске и каталоге', 'Позиция в поиске и каталоге',
        'В корзину, всего', 'Заказано товаров', 'Отменено товаров',
        'Доставлено товаров', 'Возвращено товаров', 'Заказано на сумму',
        'В корзину из карточки товара', 'Выкупили ШТ', 
        'Расход, ₽', 'Рекламные заказано на сумму',
        'Рекламные заказано товаров', 'Рекламные показы',
        'Рекламные показы на карточке товара', 'Цена'
        ]
    funnel_columns_all = [
        'Дата', 'Артикул', 'ТипАктивности', 'Показы, всего',
        'Показы на карточке товара', 'Показы в поиске и каталоге',
        'Позиция в поиске и каталоге', 'В корзину, всего',
        'Заказано товаров', 'Отменено товаров', 'Доставлено товаров',
        'Возвращено товаров', 'Заказано на сумму',
        'В корзину из карточки товара', 'Выкупили ШТ',
        'Расход, ₽', 'Рекламные заказано на сумму',
        'Рекламные заказано товаров', 'Рекламные показы',
        'Рекламные показы на карточке товара', 'Цена',
        'Артикул OZ', 'Наименование', 'Коллекция', 'Бренд', 'Сезон',
        'Направление', 'Розничный отдел', 'Модель', 'Группа',
        'Бизнес-группа', 'Техсегмент', 'Байер', 'Две последние коллекции',
        'Основной артикул', 'Себестоимость с НДС', 'Процент выкупа', 'НДС',
        'Ответственный за группу', 'Группа для отчетов',
        'ID кампании', 'Инструмент', 'Место размещения'
    ]
    funnel_columns_widing = ['Дата', 'Артикул', 'Артикул OZ', 'ТипАктивности', 'Показы, всего', 'Показы на карточке товара',
        'Показы в поиске и каталоге', 'Позиция в поиске и каталоге',
        'В корзину, всего', 'Заказано товаров', 'Отменено товаров',
        'Доставлено товаров', 'Возвращено товаров', 'Заказано на сумму',
        'В корзину из карточки товара', 'Выкупили ШТ', 
        'Расход, ₽', 'Рекламные заказано на сумму',
        'Рекламные заказано товаров', 'Рекламные показы',
        'Рекламные показы на карточке товара', 'Цена'
        ]

    # import numpy as np
    # import pandas as pd

    def build_funnel_wide(
        df_raw: pd.DataFrame,
        funnel_columns: list,
        # [v2] 4 новых типа активности:
        all_types=(
            'Оплата за клик: поиск и рекомендации',
            'Оплата за клик: поиск',
            'Оплата за заказ: выбранные товары',
            'Органика',
        ),
        infer_organic_by_zero_spend=False,
        spend_col='Расход, ₽',
        extra_agg='first',          # 'first' | 'join'
        extra_join_sep=' | ',
        check_spend_invariance=True,
        atol=1e-6, rtol=1e-9
    ):
        """
        Склеивает строки по (Дата, Артикул), раскладывает метрики по типам активности,
        добавляет ИТОГИ, которые считаются напрямую из исходника по ключу (Дата, Артикул).
        Благодаря этому 'Расход, ₽' (и прочие итоги) сохраняют исходные значения.
        """

        # ---- 0) Исходная выборка для расчётов (только нужные колонки) ----
        cols_present = [c for c in funnel_columns if c in df_raw.columns]
        df = df_raw.loc[:, cols_present].copy()

        # ---- 1) Нормализуем тип активности ----
        type_col = 'ТипАктивности'
        # [v2] Старый replace 'ТОП' → 'Вывод в топ' больше не нужен.
        # df[type_col] = df[type_col].replace({'ТОП': 'Вывод в топ'})
        if infer_organic_by_zero_spend and spend_col in df.columns:
            m0 = pd.to_numeric(df[spend_col], errors='coerce').fillna(0).eq(0)
            df.loc[m0, type_col] = 'Органика'

        # ---- 2) Ключи/метрики и приведение типов ----
        key_cols = ['Дата', 'Артикул']
        met_start = funnel_columns.index(type_col) + 1
        metric_cols = [c for c in funnel_columns[met_start:] if c in df.columns]

        # аккуратно приводим метрики
        for c in metric_cols:
            if c == spend_col:
                df[c] = pd.to_numeric(df[c], errors='coerce').astype('float64')   # спенд в float64
            else:
                df[c] = pd.to_numeric(df[c], errors='coerce').astype('float32')

        # ---- 3) ИТОГИ БЕЗ ПРЕФИКСОВ (ИСТИНА) по (Дата, Артикул) ----
        totals_df = (df.groupby(key_cols, as_index=False)[metric_cols]
                    .sum(min_count=1))   # если где-то все NaN, останется NaN; это корректно

        # ---- 4) Префиксные метрики по типам ----
        g = (df.groupby(key_cols + [type_col], as_index=False)[metric_cols]
            .sum(min_count=1))

        # Дополним all_types тем, что реально встретилось
        types_present = g[type_col].dropna().unique().tolist()
        all_types = list(dict.fromkeys(list(all_types) + [t for t in types_present if t not in all_types]))

        # Полная база ключей = все пары (Дата, Артикул), которые встречаются в исходнике
        base = (df[key_cols].drop_duplicates()
                        .set_index(key_cols)
                        .sort_index())

        # Сформируем блоки префиксных метрик и флаги наличия типов
        metric_blocks, flag_blocks = [], []
        for t in all_types:
            sub = g[g[type_col] == t].set_index(key_cols)

            if sub.empty:
                # пустой тип → нули на всю базу
                sub_metrics = pd.DataFrame(
                    0.0, index=base.index,
                    columns=[f'{t}_{m}' for m in metric_cols],
                    dtype='float32'
                )
            else:
                sub_metrics = (sub[metric_cols]
                            .rename(columns={m: f'{t}_{m}' for m in metric_cols})
                            .reindex(base.index, fill_value=0.0))

                # типы данных: спенд оставляем float64
                for col in sub_metrics.columns:
                    if col.endswith(spend_col):
                        sub_metrics[col] = sub_metrics[col].astype('float64')
                    else:
                        sub_metrics[col] = sub_metrics[col].astype('float32')

            metric_blocks.append(sub_metrics)

            # бинарный флаг присутствия типа (на уровне ключа)
            flag = pd.Series(1, index=sub.index, name=t) if not sub.empty else pd.Series(0, index=base.index, name=t)
            flag_blocks.append(flag.reindex(base.index, fill_value=0).astype('int8'))

        metrics_block = pd.concat(metric_blocks, axis=1)
        flags_block   = pd.concat(flag_blocks, axis=1)

        # ---- 5) СБОРКА CORE: ключи + флаги + префиксные метрики + ИТОГИ ИЗ totals_df ----
        core = pd.concat(
            [
                base.reset_index(),
                flags_block.reset_index(drop=True),
                metrics_block.reset_index(drop=True)
            ],
            axis=1
        )

        # присоединяем ИТОГИ (истина) строго m:1
        core = core.merge(totals_df, on=key_cols, how='left', validate='m:1')
        
        # ---- 6) ДОП. колонки из df_raw (НЕ участвуют в расчётах) ----
        exclude = set(key_cols + metric_cols)
        extra_cols = [c for c in df_raw.columns if c not in exclude]
        if extra_cols:
            if extra_agg == 'first':
                dims_block = (df_raw[key_cols + extra_cols]
                                .sort_values(key_cols)
                                .groupby(key_cols, as_index=False)
                                .first())
            elif extra_agg == 'join':
                def _join_unique(s):
                    v = pd.unique(s.dropna().astype(str))
                    return extra_join_sep.join(v) if len(v) else np.nan
                dims_block = (df_raw[key_cols + extra_cols]
                                .groupby(key_cols, as_index=False)
                                .agg({c: _join_unique for c in extra_cols}))
            else:
                raise ValueError("extra_agg должен быть 'first' или 'join'")

            out = core.merge(dims_block, on=key_cols, how='left', validate='m:1')
        else:
            out = core

        # ---- 7) Проверка инварианта для 'Расход, ₽' (опционально) ----
        if check_spend_invariance and (spend_col in totals_df.columns):
            base_sp = (totals_df.groupby(key_cols, as_index=False)[spend_col].sum(min_count=1)
                                .rename(columns={spend_col: '__base__'}))
            after_sp = (out.groupby(key_cols, as_index=False)[spend_col].sum(min_count=1)
                            .rename(columns={spend_col: '__after__'}))
            chk = base_sp.merge(after_sp, on=key_cols, how='outer').fillna(0)
            bad = chk.loc[~np.isclose(chk['__base__'], chk['__after__'], rtol=rtol, atol=atol)]
            if not bad.empty:
                print("[WARN] Инвариант по 'Расход, ₽' нарушен для некоторых ключей (первые 10):")
                print(bad.head(10))

        # ---- 8) Порядок колонок: ключи → доп.колонки → флаги → префиксные метрики → ИТОГИ ----
        ordered = []
        ordered += key_cols
        ordered += type_col
        ordered += [c for c in df_raw.columns if (c in out.columns and c not in key_cols and c not in (metric_cols))]
        ordered += [t for t in all_types if t in out.columns]

        for m in metric_cols:
            # префиксные
            ordered += [f'{t}_{m}' for t in all_types if f'{t}_{m}' in out.columns]
        # ИТОГИ (без префикса) — в самом конце блоком в исходном порядке
        ordered += [m for m in metric_cols if m in out.columns]

        out = out[[c for c in ordered if c in out.columns]].copy()
        return out

    # [v2] Динамически собираем порядок колонок под новые 4 типа активности.
    # Раньше тут был статический список под старые префиксы (Вывод в топ / Трафарет / Оплата за заказ / Органика).
    NEW_TYPES = (
        'Оплата за клик: поиск и рекомендации',
        'Оплата за клик: поиск',
        'Оплата за заказ: выбранные товары',
        'Органика',
    )
    # Базовые имена метрик (порядок), которые разворачиваются по типам
    _metric_order = [c for c in funnel_columns_widing
                     if c not in ('Дата', 'Артикул', 'Артикул OZ', 'ТипАктивности')]
    # Дополнительные размерности из df_funnel_reference, попавшие через extra_cols в build_funnel_wide
    _key_set = {'Дата', 'Артикул', 'Артикул OZ', 'ТипАктивности'}
    _extra_dims = [c for c in df_funnel_reference.columns
                   if c not in _key_set and c not in _metric_order]

    final_after_widing_columns = (
        ['Дата', 'Артикул', 'Артикул OZ', 'ТипАктивности']
        + _extra_dims
        + list(NEW_TYPES)                                                       # бинарные флаги типов
        + [f'{t}_{m}' for m in _metric_order for t in NEW_TYPES]                 # префиксные метрики
        + _metric_order                                                          # итоги без префикса
    )

    out = build_funnel_wide(df_raw=df_funnel_reference, funnel_columns=funnel_columns_widing)
    # Оставляем только реально присутствующие колонки (защита от опечаток / отсутствия типа в данных)
    final_after_widing_columns = [c for c in final_after_widing_columns if c in out.columns]
    out = out[final_after_widing_columns]
    df_funnel_reference = out

    # 15. Связать "ВоронкаСправочник" с "Остатки с дистрибуцией"
    try:
        print("Начинаем создавать таблицу ДБбезПризнаков...")
        start_time = time.time()  # Запускаем таймер
        df_stock_with_distribution['Дата'] = pd.to_datetime(df_stock_with_distribution['Дата'])
        df_final_db = pd.merge(df_funnel_reference, df_stock_with_distribution, left_on=["Дата", "Артикул"], right_on=["Дата", "Артикул"], how="left")
        df_final_db = format_date_column(df_final_db, 'Дата')

        # Вывод первых 5 строк
        print("Первые 5 строк таблицы ДБбезПризнаков:")
        print(df_final_db.head())

        elapsed_time = time.time() - start_time  # Вычисляем затраченное время
        print(f"Таблица ДБбезПризнаков успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при создании таблицы ДБбезПризнаков: {e}")

    # 16. Получить данные из файла !!!_Признаки для артикула и даты для Озон
    try:
        print("Начинаем получать данные для Признаков...")
        start_time = time.time()  # Запускаем таймер
        file_path_features = os.path.join(FOLDER_PATH_FEATURES, "!!!_Признаки для артикула и даты для Озон.xlsx")
        if os.path.exists(file_path_features):
            df_item_features = pd.read_excel(file_path_features, sheet_name="Признаки для артикула", dtype=str, engine="calamine")
            df_date_features = pd.read_excel(file_path_features, sheet_name="Признаки для дат", dtype={0: "datetime64[ns]", **{i: str for i in range(1, 6)}}, engine="calamine")

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
            print("Файл '!!!_Признаки для артикула и даты для Озон.xlsx' не найден.")
    except Exception as e:
        print(f"Ошибка при получении данных для Признаков: {e}")


    # 17. Связать "ДБбезПризнаков" с "Признаки для артикула"
    try:
        print("Начинаем создавать таблицу ДБсПризнакамиАртикула...")
        start_time = time.time()  # Запускаем таймер
        df_final_db_item_features = pd.merge(df_final_db, df_item_features, left_on=["Артикул"], right_on=["Артикул"], how="left")
        df_final_db_item_features = format_date_column(df_final_db_item_features, 'Дата')

        # Вывод первых 5 строк
        print("Первые 5 строк таблицы ДБсПризнакамиАртикула:")
        print(df_final_db_item_features.head())

        elapsed_time = time.time() - start_time  # Вычисляем затраченное время
        print(f"Таблица ДБсПризнакамиАртикула успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при создании таблицы ДБсПризнакамиАртикула: {e}")



    # 18. Связать "ДБсПризнакамиАртикула" с "Признаки для дат"
    try:
        print("Начинаем создавать таблицу ДБсПризнаками...")
        start_time = time.time()
        df_final_db_all_features = pd.merge(df_final_db_item_features, df_date_features, on="Дата", how="left")
        df_final_db_all_features = format_date_column(df_final_db_all_features, 'Дата')

        print("Первые 5 строк таблицы ДБсПризнаками:")
        print(df_final_db_all_features.head())

        # Сохранение финальной таблицы
        # df_final_db_all_features.to_csv(os.path.join(FOLDER_PATH, "ДБсПризнаками_Ozon_New.csv"), index=False)

        elapsed_time = time.time() - start_time
        print(f"Таблица ДБсПризнаками успешно создана. Время выполнения: {format_elapsed_time(elapsed_time)}")
    except Exception as e:
        print(f"Ошибка при создании таблицы ДБсПризнаками: {e}")

    # 19. [v2] Дополнительная\финальная правка ТипАктивности под новые 4 типа.
    # Платные типы по убыванию приоритета (для красивого combo «А/Б», если строка попала под несколько):
    type_order_paid = [
        'Оплата за клик: поиск и рекомендации',
        'Оплата за клик: поиск',
        'Оплата за заказ: выбранные товары',
    ]
    paid_cols = [c for c in type_order_paid if c in df_final_db_all_features.columns]

    # Матрица флагов платных типов
    flags = df_final_db_all_features[paid_cols].fillna(0).astype('uint8').to_numpy()
    labels = np.array(paid_cols, dtype=object)

    # Подписи для платных комбинаций
    combo = ['/'.join(labels[row.astype(bool)]) if row.any() else '' for row in flags]
    df_final_db_all_features['ТипАктивности'] = combo

    # Только органика — подставим "Органика"
    if 'Органика' in df_final_db_all_features.columns:
        only_org = df_final_db_all_features['Органика'].fillna(0).astype('uint8').eq(1) & (flags.sum(axis=1) == 0)
        df_final_db_all_features.loc[only_org, 'ТипАктивности'] = 'Органика'

    # Пустые — в "Органика"
    df_final_db_all_features['ТипАктивности'] = df_final_db_all_features['ТипАктивности'].replace('', 'Органика')

    # Если ярлык "Органика", но расход > 0 — это не органика, а самая частая платная категория.
    SPEND_CANDIDATES = ['Расход, ₽', 'Расход, руб', 'Расход, Р', 'Расход']
    spend_col = next((c for c in SPEND_CANDIDATES if c in df_final_db_all_features.columns), None)
    if spend_col is None:
        raise KeyError(f"Не найдена колонка расхода среди: {SPEND_CANDIDATES}")

    spend = pd.to_numeric(df_final_db_all_features[spend_col], errors='coerce')
    is_organic = df_final_db_all_features['ТипАктивности'].astype('string').str.strip().eq('Органика')
    has_spend = spend.fillna(0).gt(0)
    mask = is_organic & has_spend
    df_final_db_all_features.loc[mask, 'ТипАктивности'] = 'Оплата за клик: поиск и рекомендации'


    # 20. Получить данные по склейкам из SQL
    # === SQL СЦЕПКИ ОЗОН ===
    sql = """
    SELECT scepka.[id]
        ,scepka.[offer_id]
        ,scepka.[product_id]
        ,sku.fbo_sku as [Артикул OZ]
        ,scepka.[group_value] as [Текущая склейка]
        ,sku.[article]
        ,scepka.[updated_at] as [Дата Обновления]
    FROM [DBReport].[mp].[ozon_scepka] scepka
    JOIN [DBReport].[mp].[ozon_sku] sku 
    ON  scepka.[product_id] = sku.[product_id] 
    and sku.actual = 1
    """
    df_links = pd.read_sql(sql, engine)
    df_links['Артикул OZ'] = df_links['Артикул OZ'].astype(str)
    df_links.to_excel(os.path.join(FOLDER_PATH, f"Склейки товаров\\OZ\\{df_links['Дата Обновления'].iloc[0].strftime('%d.%m.%Y')}_Склейка Товаров_OZ.xlsx"))
    df_final_db_all_features = pd.merge(df_final_db_all_features, df_links[["Артикул OZ", "Текущая склейка"]], how='left', on='Артикул OZ')


    print('Shape df_final_db_all_features: ' + str(df_final_db_all_features.shape))

    # =====================================================================
    # 21 [v2]. Переименование колонок согласно "Описание столбцов дашборда ОЗ.xlsx"
    # =====================================================================
    # База: для метрик берём из строк итогов (без префикса) и из ключей/измерений.
    # Префиксные колонки получают новый префикс (один из 4 типов) + новое тело метрики.
    NEW_TYPES_RENAME = (
        'Оплата за клик: поиск и рекомендации',
        'Оплата за клик: поиск',
        'Оплата за заказ: выбранные товары',
        'Органика',
    )

    # Базовые тела метрик: старое имя метрики -> новое
    METRIC_BODY_MAP = {
        'Показы, всего':                       'Количество показов товара',
        'Показы на карточке товара':           'Просмотры карточки товара',
        'Показы в поиске и каталоге':          'Показы в поиске и каталоге',
        'Позиция в поиске и каталоге':         'Средняя позиция в поиске и каталоге',
        'В корзину, всего':                    'Добавлений в корзину всего',
        'Заказано товаров':                    'Количество заказанных товаров',
        'Отменено товаров':                    'Количество отменённых товаров',
        'Доставлено товаров':                  'Количество доставленных товаров',
        'Возвращено товаров':                  'Количество возвращённых товаров',
        'Заказано на сумму':                   'Сумма заказов',
        'В корзину из карточки товара':        'Добавлений в корзину из карточки товара',
        'Выкупили ШТ':                         'Количество выкупленных товаров',
        'Расход, ₽':                           'Расходы на рекламу',
        'Рекламные заказано на сумму':         'Сумма заказов от рекламы',
        'Рекламные заказано товаров':          'Количество заказанных товаров от рекламы',
        'Рекламные показы':                    'Количество рекламных показов',
        'Рекламные показы на карточке товара': 'Количество кликов по рекламе',
        'Цена':                                'Розничная цена товара',
    }

    # Прямое переименование (ключи / измерения / справочник / признаки / служебные)
    DIRECT_RENAME = {
        'Дата':                          'Дата отчёта',
        # 'Артикул' и 'Артикул OZ' оставляем без переименования (по требованию)
        'ТипАктивности':                 'Тип рекламной активности',
        'Наименование':                  'Наименование товара',
        'Коллекция':                     'Коллекция',
        'Бренд':                         'Бренд товара',
        'Сезон':                         'Сезон',
        'Направление':                   'Направление',
        'Розничный отдел':               'Розничный отдел',
        'Модель':                        'Модель',
        'Группа':                        'Товарная группа',
        'Бизнес-группа':                 'Бизнес-группа',
        'Техсегмент':                    'Технологический сегмент',
        'Байер':                         'Байер',
        'Две последние коллекции':       'Две последние коллекции',
        'Основной артикул':              'Основной артикул',
        'Себестоимость с НДС':           'Себестоимость товара с НДС',
        'Процент выкупа':                'Процент выкупа',
        'НДС':                           'Ставка НДС',
        'Ответственный за группу':       'Ответственный за товарную группу',
        'Группа для отчетов':            'Группа для отчётов',
        'SKU из объединенной карточки':  'Идентификатор объединённой карточки на Озон',
        'Ассоциированные заказы, руб':   'Сумма ассоциированных заказов',
        'Ассоциированные заказы, шт':    'Количество ассоциированных заказов',
        'Остаток Агрегатора':            'Остаток товара на складе агрегатора',
        'Дистрибуция':                   'Дистрибуция размеров',
        'Признак Артикула 1':            'Признак товара 1',
        'Признак Артикула 2':            'Признак товара 2',
        'Признак Артикула 3':            'Признак товара 3',
        'Признак Артикула 4':            'Признак товара 4',
        'Признак Артикула 5':            'Признак товара 5',
        'Признак Даты 1':                'Признак даты 1',
        'Признак Даты 2':                'Признак даты 2',
        'Признак Даты 3':                'Признак даты 3',
        'Признак Даты 4':                'Признак даты 4',
        'Признак Даты 5':                'Признак даты 5',
        'Текущая склейка':               'Идентификатор текущей склейки товара',
    }
    # Итоги без префикса = переименование тела метрики
    DIRECT_RENAME.update(METRIC_BODY_MAP)
    # Бинарные флаги типов
    FLAG_RENAME = {
        'Оплата за клик: поиск и рекомендации': 'Признак активности Оплата за клик: поиск и рекомендации',
        'Оплата за клик: поиск':                'Признак активности Оплата за клик: поиск',
        'Оплата за заказ: выбранные товары':    'Признак активности Оплата за заказ: выбранные товары',
        'Органика':                             'Признак органического трафика',
    }

    def _build_full_rename_map(columns):
        rename = {}
        for col in columns:
            if col in DIRECT_RENAME:
                rename[col] = DIRECT_RENAME[col]
                continue
            if col in FLAG_RENAME:
                rename[col] = FLAG_RENAME[col]
                continue
            # Префиксные: <тип>_<метрика>
            for t in NEW_TYPES_RENAME:
                pref = t + '_'
                if col.startswith(pref):
                    body = col[len(pref):]
                    new_body = METRIC_BODY_MAP.get(body, body)
                    rename[col] = f'{t}_{new_body}'
                    break
        return rename

    rename_map = _build_full_rename_map(df_final_db_all_features.columns.tolist())
    df_final_db_all_features = df_final_db_all_features.rename(columns=rename_map)

    # Защита от дубликатов
    cols = list(df_final_db_all_features.columns)
    if len(cols) != len(set(cols)):
        from collections import Counter
        dup = [c for c, n in Counter(cols).items() if n > 1]
        raise RuntimeError(f"[v2] После переименования возникли дубликаты колонок: {dup}")

    print(f"[v2] Колонок после переименования: {len(cols)}; дубликатов: 0")

    # 22 [v2]. Сохранить df_final_db_all_features в ТЕСТОВЫЙ csv (не перетирая боевой)
    table = pa.Table.from_pandas(df_final_db_all_features)
    csv.write_csv(table, os.path.join(FOLDER_PATH, "ДБсПризнаками_Ozon.csv"))
    print(f"[v2] Сохранён тестовый CSV: {os.path.join(FOLDER_PATH, 'ДБсПризнаками_Ozon.csv')}")


    # 22. Обновляем и сохраняем Excel-файл
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

    # [v2] Тестовый запуск: НЕ обновляем боевой "Показы и затраты ОЗ_2.0.xlsx".
    # Достаточно созданного выше Тест_ДБсПризнаками_Ozon.csv для проверки.
    print("[v2] Тестовый прогон: обновление боевого Excel-файла пропущено.")
    # return  # ранний выход из assemble(); ниже — старый код, оставлен закомментированным.

    try:  # [v2] dead code (оставлен, чтобы не терять историю; не выполняется из-за return выше)
        print("Подготовка данных для ДБ завершена.")
        # input("Начать обновление файлов ДБ? Для подтверждения нажмите Enter...")
        print("Начинаем обновлять файл 'Показы и затраты ОЗ_2.0.xlsx'...")
        start_time = time.time()  # Запускаем таймер

        # Путь к исходному файлу
        file_path_shows_expenses = os.path.join(FOLDER_PATH_FOR_DB, "Показы и затраты ОЗ_2.0.xlsx")

        if os.path.exists(file_path_shows_expenses):
            # Создаем новое имя файла с текущей датой без года
            current_month_day = time.strftime("%d.%m")  # Текущая дата в формате ДД.ММ
            new_file_name = f"Показы и затраты ОЗ_2.0 {current_month_day}.xlsx"
            new_file_path = os.path.join(FOLDER_PATH_FEATURES, new_file_name)

            # Путь для сохранения в дополнительную папку FOLDER_PATH_DUDL
            dudl_file_path = os.path.join(FOLDER_PATH_DUDL, new_file_name)

            # Удаляем старые файлы из FOLDER_PATH_DUDL
            try:
                if os.path.exists(FOLDER_PATH_DUDL):
                    for filename in os.listdir(FOLDER_PATH_DUDL):
                        # Ищем файлы с шаблоном "Показы и затраты ОЗ_2.0 DD.MM.xlsx"
                        match = re.match(r"Показы и затраты ОЗ_2\.0 (\d{2}\.\d{2})\.xlsx", filename)
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
                        match = re.match(r"Показы и затраты ОЗ_2\.0 (\d{2}\.\d{2})\.xlsx", filename)
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
            print("Файл 'Показы и затраты ОЗ_2.0.xlsx_Test' не найден.")
    except Exception as e:
        print(f"Ошибка при обработке файла 'Показы и затраты ОЗ_2.0.xlsx_Test': {e}")

if __name__ == "__main__":
    assemble()
