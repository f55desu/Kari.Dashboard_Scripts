# -*- coding: utf-8 -*-
"""
ozon_union_db.py
================
Загрузка листа "Union" из Excel-файлов "Аналитика продвижения" в PostgreSQL.

База:   analytics
Схема:  work
Таблица: work.ozon_promo_union

Источник Excel:
    \\kari.local\...\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!\Затраты\
        Озон. Затраты из Аналитики New Format\Аналитика продвижения_DD.MM.YYYY.xlsx

Структура листа "Union":
    строка 0  — "Период: DD.MM.YYYY - DD.MM.YYYY"   (дата отчёта)
    строка 1  — заголовки
    строка 2+ — данные

Подключение к PostgreSQL — через SSH-туннель (localhost:15432), который должен быть
поднят отдельно (pgadmin_connect.bat или автозапуск analytics_tunnel.vbs).
См. Analytics_PostgreSQL_Guide.md.

Учётные данные читаются из config/postgres_config.json (рядом со скриптом).

Модуль НИЧЕГО не делает при импорте и спроектирован так, чтобы шаг загрузки в БД
был ОПЦИОНАЛЬНЫМ: функция try_upload_latest_to_postgres() никогда не бросает
исключение — если сервер/туннель недоступен, остальной дашборд собирается дальше.
"""

import os
import re
import sys
import json
import glob
import logging
from datetime import datetime, timedelta
from urllib.parse import quote_plus

import pandas as pd

logger = logging.getLogger("ozon_union")


def _safe_print(msg):
    """print, который не падает на консолях без UTF-8 (cp1251 и т.п.)."""
    try:
        print(msg)
    except Exception:
        try:
            enc = sys.stdout.encoding or "utf-8"
            print(msg.encode(enc, errors="replace").decode(enc, errors="replace"))
        except Exception:
            pass

# --------------------------------------------------------------------------- #
#  Константы
# --------------------------------------------------------------------------- #
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CONFIG_PATH = os.path.join(BASE_DIR, "config", "postgres_config.json")

SOURCE_FOLDER = (
    r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров"
    r"\Дашбоард по рекламным кампаниям"
    r"\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!\Затраты"
    r"\Озон. Затраты из Аналитики New Format"
)

SHEET_NAME = "Union"
TABLE_SCHEMA = "work"
TABLE_NAME = "ozon_promo_union"
FULL_TABLE = f"{TABLE_SCHEMA}.{TABLE_NAME}"

# Имя файла: "Аналитика продвижения_DD.MM.YYYY.xlsx"
FILE_PATTERN = r"Аналитика продвижения_(\d{2}\.\d{2}\.\d{4})\.xlsx"

# Заголовок Excel  ->  колонка БД
COLUMN_MAP = {
    "SKU в продвижении": "sku_promo",
    "Название товара в продвижении": "name_promo",
    "SKU из объединенной карточки": "sku_card",
    "Название товара из объединенной карточки": "name_card",
    "Продажи, ₽": "sales_rub",
    "Заказы, шт": "orders_qty",
}

# Порядок колонок в таблице (без ingested_at — он DEFAULT now())
DB_COLUMNS = [
    "report_date",
    "sku_promo",
    "name_promo",
    "sku_card",
    "name_card",
    "sales_rub",
    "orders_qty",
    "source_file",
]

DDL_STATEMENTS = [
    f"""
    CREATE TABLE IF NOT EXISTS {FULL_TABLE} (
        report_date  DATE        NOT NULL,
        sku_promo    BIGINT,
        name_promo   TEXT,
        sku_card     BIGINT,
        name_card    TEXT,
        sales_rub    NUMERIC(14,2),
        orders_qty   INTEGER,
        source_file  TEXT,
        ingested_at  TIMESTAMPTZ NOT NULL DEFAULT now()
    )
    """,
    f"CREATE INDEX IF NOT EXISTS ix_ozon_promo_union_date ON {FULL_TABLE} (report_date)",
    f"CREATE INDEX IF NOT EXISTS ix_ozon_promo_union_sku  ON {FULL_TABLE} (sku_promo)",
]

DEFAULT_BATCH_SIZE = 5000


# --------------------------------------------------------------------------- #
#  Конфигурация / подключение
# --------------------------------------------------------------------------- #
def load_db_config(config_path=None):
    """
    Читает параметры подключения. Приоритет:
      1) переменные окружения PG_HOST / PG_PORT / PG_DB / PG_USER / PG_PASSWORD
         (перекрывают значения из файла);
      2) JSON-файл config/postgres_config.json.

    Возвращает dict: host, port, dbname, user, password, connect_timeout.
    """
    cfg = {
        "host": "localhost",
        "port": 15432,
        "dbname": "analytics",
        "user": None,
        "password": None,
        "connect_timeout": 10,
    }

    path = config_path or DEFAULT_CONFIG_PATH
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as fh:
            file_cfg = json.load(fh)
        cfg.update({k: v for k, v in file_cfg.items() if v is not None})

    # Переменные окружения имеют приоритет
    env_map = {
        "PG_HOST": "host",
        "PG_PORT": "port",
        "PG_DB": "dbname",
        "PG_USER": "user",
        "PG_PASSWORD": "password",
    }
    for env_key, cfg_key in env_map.items():
        if os.environ.get(env_key):
            cfg[cfg_key] = os.environ[env_key]

    cfg["port"] = int(cfg["port"])
    cfg["connect_timeout"] = int(cfg["connect_timeout"])

    if not cfg["user"] or not cfg["password"]:
        raise RuntimeError(
            "Не заданы user/password для PostgreSQL. "
            f"Заполните {path} или переменные окружения PG_USER / PG_PASSWORD."
        )
    return cfg


def get_engine(cfg=None, config_path=None):
    """Создаёт SQLAlchemy engine. Требует sqlalchemy + psycopg2."""
    from sqlalchemy import create_engine

    if cfg is None:
        cfg = load_db_config(config_path)

    url = (
        "postgresql+psycopg2://"
        f"{quote_plus(str(cfg['user']))}:{quote_plus(str(cfg['password']))}"
        f"@{cfg['host']}:{cfg['port']}/{cfg['dbname']}"
    )
    return create_engine(
        url,
        connect_args={"connect_timeout": cfg["connect_timeout"]},
        pool_pre_ping=True,
    )


# --------------------------------------------------------------------------- #
#  Чтение Excel
# --------------------------------------------------------------------------- #
def _parse_date_str(s):
    return datetime.strptime(s.strip(), "%d.%m.%Y").date()


def parse_report_date(period_cell, filename):
    """
    Дата отчёта = первая дата из ячейки "Период: DD.MM.YYYY - ...".
    Фолбэк: дата из имени файла минус 1 день (файл за DD.MM датируется предыдущим днём).
    """
    if isinstance(period_cell, str):
        m = re.search(r"(\d{2}\.\d{2}\.\d{4})", period_cell)
        if m:
            return _parse_date_str(m.group(1))
    m = re.search(FILE_PATTERN, os.path.basename(filename))
    if m:
        return _parse_date_str(m.group(1)) - timedelta(days=1)
    raise ValueError(f"Не удалось определить report_date для файла: {filename}")


def _to_int64(series):
    return pd.to_numeric(series, errors="coerce").astype("Int64")


def _to_numeric(series):
    # В отчётах пропуски иногда обозначены '-'
    s = series.replace("-", pd.NA)
    return pd.to_numeric(s, errors="coerce")


def read_union_file(path):
    """
    Читает лист "Union" одного файла и возвращает очищенный DataFrame
    с колонками DB_COLUMNS, готовый к загрузке.
    """
    filename = os.path.basename(path)

    # Строка 0 — период
    period_cell = pd.read_excel(
        path, sheet_name=SHEET_NAME, header=None, nrows=1
    ).iloc[0, 0]
    report_date = parse_report_date(period_cell, filename)

    # Заголовок — строка с индексом 1
    df = pd.read_excel(path, sheet_name=SHEET_NAME, header=1)

    missing = [c for c in COLUMN_MAP if c not in df.columns]
    if missing:
        raise ValueError(f"{filename}: в листе '{SHEET_NAME}' нет колонок: {missing}")

    df = df[list(COLUMN_MAP.keys())].rename(columns=COLUMN_MAP)

    # Типизация
    df["sku_promo"] = _to_int64(df["sku_promo"])
    df["sku_card"] = _to_int64(df["sku_card"])
    df["sales_rub"] = _to_numeric(df["sales_rub"]).round(2)
    df["orders_qty"] = _to_int64(df["orders_qty"])
    df["name_promo"] = df["name_promo"].astype("string")
    df["name_card"] = df["name_card"].astype("string")

    # Убираем полностью пустые строки (служебные/итоговые хвосты)
    df = df.dropna(
        how="all",
        subset=["sku_promo", "name_promo", "sku_card", "name_card",
                "sales_rub", "orders_qty"],
    )

    df.insert(0, "report_date", report_date)
    df["source_file"] = filename

    return df[DB_COLUMNS].reset_index(drop=True)


def list_source_files(folder=SOURCE_FOLDER):
    files = [
        f for f in glob.glob(os.path.join(folder, "*.xlsx"))
        if re.search(FILE_PATTERN, os.path.basename(f))
    ]
    # Сортируем по дате из имени файла
    def _key(f):
        m = re.search(FILE_PATTERN, os.path.basename(f))
        return _parse_date_str(m.group(1)) if m else datetime.min.date()
    return sorted(files, key=_key)


def get_latest_source_file(folder=SOURCE_FOLDER):
    files = list_source_files(folder)
    if not files:
        raise FileNotFoundError(f"Нет файлов по шаблону в папке: {folder}")
    return files[-1]


# --------------------------------------------------------------------------- #
#  Запись в PostgreSQL
# --------------------------------------------------------------------------- #
def ensure_table(engine):
    """Создаёт таблицу и индексы, если их ещё нет."""
    from sqlalchemy import text
    with engine.begin() as conn:
        for stmt in DDL_STATEMENTS:
            conn.execute(text(stmt))


def load_dataframe(engine, df, batch_size=DEFAULT_BATCH_SIZE, replace_dates=True):
    """
    Грузит DataFrame в work.ozon_promo_union.

    replace_dates=True (идемпотентно): в одной транзакции удаляет все строки
    с report_date, встречающимися в df, и вставляет df заново. Повторный прогон
    того же дня не плодит дубли.

    Вставка идёт батчами (chunksize) методом multi-INSERT.
    """
    from sqlalchemy import text

    if df.empty:
        logger.warning("Пустой DataFrame — нечего грузить.")
        return 0

    ensure_table(engine)

    dates = sorted(pd.to_datetime(df["report_date"]).dt.date.unique().tolist())

    with engine.begin() as conn:
        if replace_dates:
            conn.execute(
                text(f"DELETE FROM {FULL_TABLE} WHERE report_date = ANY(:dates)"),
                {"dates": dates},
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
    logger.info("Загружено %d строк за даты %s", len(df), dates)
    return len(df)


# --------------------------------------------------------------------------- #
#  Высокоуровневые операции
# --------------------------------------------------------------------------- #
def upload_latest_to_postgres(config_path=None, folder=SOURCE_FOLDER,
                              batch_size=DEFAULT_BATCH_SIZE):
    """
    Грузит САМЫЙ свежий файл из папки-источника (delete+insert по его дате).
    Бросает исключения наверх — для использования вне дашборда.
    Возвращает (rows, report_date).
    """
    cfg = load_db_config(config_path)
    latest = get_latest_source_file(folder)
    df = read_union_file(latest)
    report_date = df["report_date"].iloc[0]

    engine = get_engine(cfg)
    try:
        rows = load_dataframe(engine, df, batch_size=batch_size, replace_dates=True)
    finally:
        engine.dispose()
    return rows, report_date


def try_upload_latest_to_postgres(config_path=None, folder=SOURCE_FOLDER):
    """
    БЕЗОПАСНАЯ обёртка для интеграции в дашборд.
    НИКОГДА не бросает исключение: если PostgreSQL/туннель недоступен,
    отсутствуют зависимости или нет конфига — печатает предупреждение и
    возвращает False, чтобы сборка дашборда продолжилась.
    """
    try:
        rows, report_date = upload_latest_to_postgres(config_path=config_path,
                                                       folder=folder)
        msg = f"✅ PostgreSQL: загружено {rows} строк за {report_date} в {FULL_TABLE}"
        _safe_print(msg)
        logger.info(msg)
        return True
    except Exception as e:  # noqa: BLE001 — намеренно ловим всё
        msg = f"⚠️ Загрузка в PostgreSQL пропущена (не критично): {e}"
        _safe_print(msg)
        logger.warning(msg)
        return False


if __name__ == "__main__":
    # Ручной прогон: грузит самый свежий файл.
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ok = try_upload_latest_to_postgres()
    raise SystemExit(0 if ok else 1)
