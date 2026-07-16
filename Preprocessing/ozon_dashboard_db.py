# -*- coding: utf-8 -*-
"""
ozon_dashboard_db.py
====================
Загрузка последнего дня из CSV-файла ДБсПризнаками_Ozon.csv в PostgreSQL.

База:   analytics
Схема:  work
Таблица: work.ozon_dashboard_daily

Источник CSV:
    \\kari.local\...\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!\ДБсПризнаками_Ozon.csv

Подключение к PostgreSQL — через SSH-туннель (localhost:15432).
Учётные данные читаются из config/postgres_config.json.

Модуль спроектирован так же, как ozon_union_db.py: функция
try_upload_latest_to_postgres() никогда не бросает исключение — если
сервер/туннель недоступен, дашборд собирается дальше.
"""

import os
import sys
import logging

import pandas as pd

logger = logging.getLogger("ozon_dashboard")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CSV_PATH = (
    r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров"
    r"\Дашбоард по рекламным кампаниям"
    r"\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!"
    r"\ДБсПризнаками_Ozon.csv"
)

TABLE_SCHEMA = "work"
TABLE_NAME = "ozon_dashboard_daily"
FULL_TABLE = f"{TABLE_SCHEMA}.{TABLE_NAME}"
DATE_COLUMN = "Дата отчёта"

PG_NAMEDATALEN = 63

_PREFIX_SHORT = {
    "Оплата за клик: поиск и рекомендации": "КликПР",
    "Оплата за клик: поиск": "КликП",
    "Оплата за заказ: выбранные товары": "Заказ",
    "Органика": "Орг",
}

_STANDALONE_SHORT = {
    "Признак активности Оплата за клик: поиск и рекомендации": "Флаг_КликПР",
    "Признак активности Оплата за клик: поиск": "Флаг_КликП",
    "Признак активности Оплата за заказ: выбранные товары": "Флаг_Заказ",
    "Идентификатор объединённой карточки на Озон": "ID_объед_карточки",
    "Количество ассоциированных заказов": "Ассоц_заказы_шт",
    "Ответственный за товарную группу": "Ответственный",
    "Себестоимость товара с НДС": "Себестоимость_НДС",
    "Остаток товара на складе агрегатора": "Остаток_агрегатор",
    "Идентификатор текущей склейки товара": "ID_склейки",
    "Средняя позиция в поиске и каталоге": "Ср_позиция",
    "Добавлений в корзину из карточки товара": "Корзина_карточка",
    "Количество заказанных товаров от рекламы": "Рекл_заказы_шт",
}

_METRIC_SHORT = {
    "Количество показов товара": "Показы",
    "Просмотры карточки товара": "Просмотры",
    "Показы в поиске и каталоге": "Поиск_каталог",
    "Средняя позиция в поиске и каталоге": "Ср_позиция",
    "Добавлений в корзину всего": "Корзина",
    "Количество заказанных товаров": "Заказы_шт",
    "Количество отменённых товаров": "Отмены_шт",
    "Количество доставленных товаров": "Доставки_шт",
    "Количество возвращённых товаров": "Возвраты_шт",
    "Сумма заказов": "Заказы_руб",
    "Добавлений в корзину из карточки товара": "Корзина_карт",
    "Количество выкупленных товаров": "Выкуп_шт",
    "Расходы на рекламу": "Расход",
    "Сумма заказов от рекламы": "Рекл_руб",
    "Количество заказанных товаров от рекламы": "Рекл_заказы",
    "Количество рекламных показов": "Рекл_показы",
    "Количество кликов по рекламе": "Рекл_клики",
    "Розничная цена товара": "Цена",
}


def _shorten_columns(columns):
    """
    Укорачивает имена колонок, не влезающие в PG_NAMEDATALEN (63 байта).
    Сокращает и префикс типа активности, и название метрики.
    """
    result = {}
    for col in columns:
        if len(col.encode("utf-8")) <= PG_NAMEDATALEN:
            continue

        if col in _STANDALONE_SHORT:
            result[col] = _STANDALONE_SHORT[col]
            continue

        matched = False
        for long_prefix, short_prefix in _PREFIX_SHORT.items():
            sep = long_prefix + "_"
            if col.startswith(sep):
                metric = col[len(sep):]
                short_metric = _METRIC_SHORT.get(metric, metric)
                result[col] = f"{short_prefix}_{short_metric}"
                matched = True
                break
        if matched:
            continue

        result[col] = col.encode("utf-8")[:PG_NAMEDATALEN].decode("utf-8", errors="ignore")

    seen = set(c for c in columns if c not in result)
    for old, new in list(result.items()):
        if new in seen:
            i = 2
            base = new[:20] if len(new.encode("utf-8")) > PG_NAMEDATALEN - 3 else new
            while f"{base}_{i}" in seen:
                i += 1
            result[old] = f"{base}_{i}"
            new = result[old]
        seen.add(new)

    return result


def _safe_print(msg):
    try:
        print(msg)
    except Exception:
        try:
            enc = sys.stdout.encoding or "utf-8"
            print(msg.encode(enc, errors="replace").decode(enc, errors="replace"))
        except Exception:
            pass


def _get_engine(config_path=None):
    from ozon_union_db import load_db_config, get_engine
    cfg = load_db_config(config_path)
    return get_engine(cfg)


def read_last_day(csv_path=CSV_PATH):
    """
    Читает CSV и возвращает DataFrame только с данными последнего дня.
    """
    df = pd.read_csv(csv_path, encoding="utf-8", low_memory=False)

    if DATE_COLUMN not in df.columns:
        raise ValueError(f"В CSV нет колонки '{DATE_COLUMN}'. Колонки: {list(df.columns[:10])}...")

    df[DATE_COLUMN] = pd.to_datetime(df[DATE_COLUMN], format="%Y-%m-%d", errors="coerce")
    max_date = df[DATE_COLUMN].max()

    if pd.isna(max_date):
        raise ValueError("Не удалось определить максимальную дату в CSV.")

    df_last = df[df[DATE_COLUMN] == max_date].copy()
    df_last[DATE_COLUMN] = df_last[DATE_COLUMN].dt.date

    logger.info("CSV прочитан: %d строк всего, %d строк за %s", len(df), len(df_last), max_date.date())
    return df_last, max_date.date()


def _table_exists(conn, schema, table):
    from sqlalchemy import text
    result = conn.execute(
        text(
            "SELECT EXISTS ("
            "  SELECT 1 FROM information_schema.tables"
            "  WHERE table_schema = :schema AND table_name = :table"
            ")"
        ),
        {"schema": schema, "table": table},
    )
    return result.scalar()


def upload_to_postgres(df, report_date, engine, batch_size=200):
    """
    Идемпотентная загрузка: DELETE по дате + INSERT.
    При первом запуске таблица создаётся автоматически через to_sql.
    """
    from sqlalchemy import text

    if df.empty:
        logger.warning("Пустой DataFrame — нечего грузить.")
        return 0

    rename_map = _shorten_columns(df.columns.tolist())
    if rename_map:
        df = df.rename(columns=rename_map)
        logger.info("Укорочено %d колонок для PostgreSQL", len(rename_map))

    with engine.begin() as conn:
        if _table_exists(conn, TABLE_SCHEMA, TABLE_NAME):
            conn.execute(
                text(f'DELETE FROM {FULL_TABLE} WHERE "{DATE_COLUMN}" = :dt'),
                {"dt": report_date},
            )

        df.to_sql(
            TABLE_NAME,
            conn,
            schema=TABLE_SCHEMA,
            if_exists="append",
            index=False,
            chunksize=batch_size,
            method="multi",
        )

    logger.info("Загружено %d строк за %s в %s", len(df), report_date, FULL_TABLE)
    return len(df)


def try_upload_latest_to_postgres(config_path=None, csv_path=CSV_PATH):
    """
    БЕЗОПАСНАЯ обёртка для интеграции в пайплайн.
    НИКОГДА не бросает исключение: при любой проблеме печатает
    предупреждение и возвращает False.
    """
    try:
        df_last, report_date = read_last_day(csv_path)
        engine = _get_engine(config_path)
        try:
            rows = upload_to_postgres(df_last, report_date, engine)
        finally:
            engine.dispose()

        msg = f"✅ PostgreSQL: загружено {rows} строк за {report_date} в {FULL_TABLE}"
        _safe_print(msg)
        logger.info(msg)
        return True
    except Exception as e:
        msg = f"⚠️ Загрузка дашборда в PostgreSQL пропущена (не критично): {e}"
        _safe_print(msg)
        logger.warning(msg)
        return False


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ok = try_upload_latest_to_postgres()
    raise SystemExit(0 if ok else 1)
