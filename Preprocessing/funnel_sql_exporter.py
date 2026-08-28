# -*- coding: utf-8 -*-
"""
funnel_sql_exporter.py
=======================
Выгрузка воронки продаж и затрат из SQL Server (cl01sql) в Excel и PostgreSQL.

Источники (SQL Server, база DBReport):
    [mp].[ozon_sales_funnel]
    [mp].[wb_sales_funnel_lk]
    [mp].[wb_marketing]

PostgreSQL (analytics, через SSH-туннель localhost:15432):
    work.ozon_sales_funnel
    work.wb_sales_funnel_lk
    work.wb_marketing

Использование:
    python funnel_sql_exporter.py --oz                # Ozon воронка: последний день
    python funnel_sql_exporter.py --wb                # WB воронка: последний день
    python funnel_sql_exporter.py --wb-costs          # WB затраты: последний день
    python funnel_sql_exporter.py --all               # Все таблицы
    python funnel_sql_exporter.py --oz --sync         # Ozon: догрузить недостающие дни
    python funnel_sql_exporter.py --wb-costs --full   # WB затраты: полная перезаливка
    python funnel_sql_exporter.py --oz --no-excel     # Только в PG, без Excel
    python funnel_sql_exporter.py --oz --no-pg        # Только Excel, без PG
"""

import os
import sys
import logging
from datetime import date, datetime, timedelta
from glob import glob

import pandas as pd
import pyodbc

logger = logging.getLogger("funnel_sql_exporter")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ── SQL Server ──────────────────────────────────────────────────────────
MSSQL_SERVER = "cl01sql"
MSSQL_DATABASE = "DBReport"
MSSQL_CONN_STR = (
    f"DRIVER={{SQL Server}};"
    f"SERVER={MSSQL_SERVER};"
    f"DATABASE={MSSQL_DATABASE};"
    f"Trusted_Connection=yes;"
)

# ── Ozon ────────────────────────────────────────────────────────────────
OZ_MSSQL_TABLE = "[mp].[ozon_sales_funnel]"
OZ_DATE_COLUMN = "Дата"
OZ_EXCEL_DIR = (
    r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров"
    r"\Дашбоард по рекламным кампаниям"
    r"\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!"
    r"\ВЫГРУЗКА воронка Озон"
)
OZ_EXCEL_SHEET = "Sheet1"
OZ_PG_SCHEMA = "work"
OZ_PG_TABLE = "ozon_sales_funnel"
OZ_PG_FULL = f"{OZ_PG_SCHEMA}.{OZ_PG_TABLE}"

# ── WB Воронка ──────────────────────────────────────────────────────────
WB_MSSQL_TABLE = "[mp].[wb_sales_funnel_lk]"
WB_DATE_COLUMN = "Дата"
WB_EXCEL_DIR = (
    r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров"
    r"\Дашбоард по рекламным кампаниям"
    r"\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!"
    r"\ВЫГРУЗКА воронка ВБ"
)
WB_EXCEL_SHEET = "Товары"
WB_PG_SCHEMA = "work"
WB_PG_TABLE = "wb_sales_funnel_lk"
WB_PG_FULL = f"{WB_PG_SCHEMA}.{WB_PG_TABLE}"

# ── WB Затраты (маркетинг) ─────────────────────────────────────────────
WB_COSTS_MSSQL_TABLE = "[mp].[wb_marketing]"
WB_COSTS_DATE_COLUMN = "Дата"
WB_COSTS_EXCEL_DIR = (
    r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров"
    r"\Дашбоард по рекламным кампаниям"
    r"\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!"
    r"\Затраты\Затраты ВБ"
)
WB_COSTS_EXCEL_SHEET = "Затраты ВБ"
WB_COSTS_PG_SCHEMA = "work"
WB_COSTS_PG_TABLE = "wb_marketing"
WB_COSTS_PG_FULL = f"{WB_COSTS_PG_SCHEMA}.{WB_COSTS_PG_TABLE}"

# ── Ozon Затраты (источник: Excel, не SQL Server) ──────────────────────
OZ_COSTS_EXCEL_SOURCE_DIR = (
    r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров"
    r"\Дашбоард по рекламным кампаниям"
    r"\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!"
    r"\Затраты\Озон. Затраты из Аналитики New Format"
)
OZ_COSTS_DATE_COLUMN = "Дата"
OZ_COSTS_EXCEL_SHEET = "Statistics"
OZ_COSTS_PG_SCHEMA = "work"
OZ_COSTS_PG_TABLE = "ozon_costs_statistics"
OZ_COSTS_PG_FULL = f"{OZ_COSTS_PG_SCHEMA}.{OZ_COSTS_PG_TABLE}"


# ═══════════════════════════════════════════════════════════════════════
#  Утилиты
# ═══════════════════════════════════════════════════════════════════════

def _safe_print(msg):
    try:
        print(msg)
    except Exception:
        try:
            enc = sys.stdout.encoding or "utf-8"
            print(msg.encode(enc, errors="replace").decode(enc, errors="replace"))
        except Exception:
            pass


def _get_pg_engine(config_path=None):
    from ozon_union_db import load_db_config, get_engine
    cfg = load_db_config(config_path)
    return get_engine(cfg)


def _pg_table_exists(conn, schema, table):
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


def _ensure_columns(conn, schema, table, df):
    from sqlalchemy import text
    type_map = {"object": "TEXT", "int64": "BIGINT", "float64": "DOUBLE PRECISION",
                "datetime64[ns]": "TIMESTAMP", "bool": "BOOLEAN"}
    existing = conn.execute(
        text(
            "SELECT column_name FROM information_schema.columns "
            "WHERE table_schema = :schema AND table_name = :table"
        ),
        {"schema": schema, "table": table},
    )
    existing_cols = {row[0] for row in existing}
    for col in df.columns:
        if col not in existing_cols:
            pg_type = type_map.get(str(df[col].dtype), "TEXT")
            conn.execute(text(
                f'ALTER TABLE {schema}.{table} ADD COLUMN IF NOT EXISTS "{col}" {pg_type}'
            ))
            print(f"  PostgreSQL: добавлена колонка \"{col}\" ({pg_type})")


# ═══════════════════════════════════════════════════════════════════════
#  Чтение из SQL Server
# ═══════════════════════════════════════════════════════════════════════

def read_from_mssql(mssql_table, date_column):
    """Читает всю таблицу из SQL Server и возвращает DataFrame."""
    conn = pyodbc.connect(MSSQL_CONN_STR)
    try:
        query = f"SELECT * FROM {mssql_table}"
        df = pd.read_sql(query, conn)
    finally:
        conn.close()

    if date_column in df.columns:
        df[date_column] = pd.to_datetime(df[date_column], errors="coerce")

    logger.info("SQL Server %s: прочитано %d строк, %d колонок", mssql_table, len(df), len(df.columns))
    return df


def read_last_day_from_mssql(mssql_table, date_column, target_date=None):
    """Читает данные за указанную дату из SQL Server.
    По умолчанию target_date = вчера (today - 1), т.к. данные за текущий день часто неполные."""
    if target_date is None:
        target_date = date.today() - timedelta(days=1)
    conn = pyodbc.connect(MSSQL_CONN_STR)
    try:
        query = f"SELECT * FROM {mssql_table} WHERE CAST([{date_column}] AS DATE) = '{target_date}'"
        df = pd.read_sql(query, conn)
    finally:
        conn.close()

    if date_column in df.columns:
        df[date_column] = pd.to_datetime(df[date_column], errors="coerce")

    report_date = target_date
    logger.info("SQL Server %s: %d строк за %s", mssql_table, len(df), report_date)
    return df, report_date


def read_dates_from_mssql(mssql_table, date_column):
    """Возвращает множество всех дат в таблице SQL Server."""
    conn = pyodbc.connect(MSSQL_CONN_STR)
    try:
        query = f"SELECT DISTINCT [{date_column}] FROM {mssql_table}"
        df = pd.read_sql(query, conn)
    finally:
        conn.close()
    return set(pd.to_datetime(df[date_column], errors="coerce").dropna().dt.date)


def read_days_from_mssql(mssql_table, date_column, dates):
    """Читает из SQL Server только указанные даты."""
    if not dates:
        return pd.DataFrame()
    conn = pyodbc.connect(MSSQL_CONN_STR)
    try:
        placeholders = ", ".join(f"'{d}'" for d in sorted(dates))
        query = f"SELECT * FROM {mssql_table} WHERE [{date_column}] IN ({placeholders})"
        df = pd.read_sql(query, conn)
    finally:
        conn.close()

    if date_column in df.columns:
        df[date_column] = pd.to_datetime(df[date_column], errors="coerce")

    logger.info("SQL Server %s: %d строк за %d дат", mssql_table, len(df), len(dates))
    return df


# ═══════════════════════════════════════════════════════════════════════
#  Сохранение в Excel
# ═══════════════════════════════════════════════════════════════════════

EXCEL_MAX_ROWS = 1_048_576

def save_to_excel(df, excel_dir, sheet_name, date_column):
    """Сохраняет DataFrame в Excel-файл(ы) с датой в имени.
    Если строк больше лимита Excel — разбивает по дням."""
    if df.empty:
        _safe_print("Пустой DataFrame — Excel не создан.")
        return None

    os.makedirs(excel_dir, exist_ok=True)

    if len(df) > EXCEL_MAX_ROWS and date_column in df.columns:
        _safe_print(f"DataFrame ({len(df):,} строк) превышает лимит Excel — сохраняю по дням")
        dates = sorted(df[date_column].dropna().unique())
        paths = []
        for dt in dates:
            df_day = df[df[date_column] == dt].copy()
            date_str = pd.Timestamp(dt).strftime("%Y-%m-%d")
            filename = f"funnel_{date_str}.xlsx"
            filepath = os.path.join(excel_dir, filename)
            df_day[date_column] = pd.to_datetime(df_day[date_column]).dt.date
            df_day.to_excel(filepath, sheet_name=sheet_name, index=False)
            paths.append(filepath)
        _safe_print(f"Excel: сохранено {len(paths)} файлов по дням в {excel_dir}")
        return paths

    max_date = df[date_column].max() if date_column in df.columns else None
    if pd.notna(max_date):
        date_str = pd.Timestamp(max_date).strftime("%Y-%m-%d")
    else:
        date_str = date.today().isoformat()

    filename = f"funnel_{date_str}.xlsx"
    filepath = os.path.join(excel_dir, filename)

    df_out = df.copy()
    if date_column in df_out.columns:
        df_out[date_column] = pd.to_datetime(df_out[date_column]).dt.date

    df_out.to_excel(filepath, sheet_name=sheet_name, index=False)
    _safe_print(f"Excel сохранён: {filepath} ({len(df_out)} строк, лист «{sheet_name}»)")
    logger.info("Excel: %s, %d строк", filepath, len(df_out))
    return filepath


# ═══════════════════════════════════════════════════════════════════════
#  Загрузка в PostgreSQL
# ═══════════════════════════════════════════════════════════════════════

def upload_to_postgres(df, report_date, date_column, pg_schema, pg_table, engine, batch_size=500):
    """Идемпотентная загрузка: DELETE по дате + INSERT."""
    from sqlalchemy import text

    full_table = f"{pg_schema}.{pg_table}"

    if df.empty:
        logger.warning("Пустой DataFrame — нечего грузить.")
        return 0

    df_up = df.copy()
    if date_column in df_up.columns:
        df_up[date_column] = pd.to_datetime(df_up[date_column]).dt.date

    with engine.begin() as conn:
        if _pg_table_exists(conn, pg_schema, pg_table):
            conn.execute(
                text(f'DELETE FROM {full_table} WHERE "{date_column}" = :dt'),
                {"dt": report_date},
            )
            _ensure_columns(conn, pg_schema, pg_table, df_up)

        df_up.to_sql(
            pg_table,
            conn,
            schema=pg_schema,
            if_exists="append",
            index=False,
            chunksize=batch_size,
            method="multi",
        )

    logger.info("PG: загружено %d строк за %s в %s", len(df_up), report_date, full_table)
    return len(df_up)


def _get_pg_existing_dates(engine, pg_schema, pg_table, date_column):
    from sqlalchemy import text
    full_table = f"{pg_schema}.{pg_table}"
    with engine.connect() as conn:
        if not _pg_table_exists(conn, pg_schema, pg_table):
            return set()
        rows = conn.execute(
            text(f'SELECT DISTINCT "{date_column}" FROM {full_table}')
        ).fetchall()
    return {row[0] for row in rows}


# ═══════════════════════════════════════════════════════════════════════
#  Чтение затрат Ozon из Excel
# ═══════════════════════════════════════════════════════════════════════

def _parse_oz_costs_date(filepath):
    """Извлекает дату данных из имени файла 'Аналитика продвижения_DD.MM.YYYY.xlsx'.
    Дата данных = дата в имени файла минус 1 день."""
    basename = os.path.basename(filepath)
    date_str = basename.split("_")[-1].replace(".xlsx", "")
    file_date = datetime.strptime(date_str, "%d.%m.%Y")
    return (file_date - timedelta(days=1)).date()


def read_oz_costs_from_excel(excel_dir, sheet_name="Statistics", cutoff_date=None, upper_date=None):
    """Читает Excel-файлы затрат Ozon из папки, фильтруя по дате.
    upper_date — верхняя граница (включительно), по умолчанию = вчера."""
    if upper_date is None:
        upper_date = date.today() - timedelta(days=1)

    pattern = os.path.join(excel_dir, "Аналитика продвижения_*.xlsx")
    files = glob(pattern)
    if not files:
        _safe_print(f"Нет файлов по шаблону: {pattern}")
        return pd.DataFrame()

    dfs = []
    for f in files:
        try:
            data_date = _parse_oz_costs_date(f)
        except (ValueError, IndexError):
            _safe_print(f"  Пропуск (не удалось определить дату): {os.path.basename(f)}")
            continue

        if cutoff_date and data_date < cutoff_date:
            continue
        if data_date > upper_date:
            continue

        df = pd.read_excel(f, sheet_name=sheet_name, engine="calamine", skiprows=1)
        df[OZ_COSTS_DATE_COLUMN] = pd.Timestamp(data_date)
        dfs.append(df)

    if not dfs:
        return pd.DataFrame()

    result = pd.concat(dfs, ignore_index=True)
    _safe_print(f"Прочитано {len(result)} строк из {len(dfs)} Excel-файлов (до {upper_date})")
    return result


def read_latest_oz_costs_from_excel(excel_dir, sheet_name="Statistics", target_date=None):
    """Читает Excel-файл затрат Ozon за указанную дату.
    По умолчанию target_date = вчера (today - 1).
    Если файла за эту дату нет — берёт ближайший предыдущий."""
    if target_date is None:
        target_date = date.today() - timedelta(days=1)

    pattern = os.path.join(excel_dir, "Аналитика продвижения_*.xlsx")
    files = glob(pattern)
    if not files:
        _safe_print(f"Нет файлов по шаблону: {pattern}")
        return pd.DataFrame(), None

    dated_files = []
    for f in files:
        try:
            data_date = _parse_oz_costs_date(f)
            dated_files.append((data_date, f))
        except (ValueError, IndexError):
            continue

    if not dated_files:
        return pd.DataFrame(), None

    # Ищем файл за target_date или ближайший предыдущий
    dated_files.sort(reverse=True)
    chosen_date, chosen_file = None, None
    for d, f in dated_files:
        if d <= target_date:
            chosen_date, chosen_file = d, f
            break

    if chosen_file is None:
        _safe_print(f"Нет файлов затрат Ozon за {target_date} или ранее.")
        return pd.DataFrame(), None

    df = pd.read_excel(chosen_file, sheet_name=sheet_name, engine="calamine", skiprows=1)
    df[OZ_COSTS_DATE_COLUMN] = pd.Timestamp(chosen_date)

    _safe_print(f"Прочитано {len(df)} строк за {chosen_date} из {os.path.basename(chosen_file)}")
    return df, chosen_date


# ═══════════════════════════════════════════════════════════════════════
#  Обогащение OZ-воронки колонкой «Артикул» из Excel
# ═══════════════════════════════════════════════════════════════════════

def _build_ozon_id_artikul_map(excel_dir=OZ_EXCEL_DIR, max_files=5):
    """Строит маппинг Ozon ID → Артикул из последних Excel-файлов воронки.
    Маппинг стабилен между файлами (проверено), поэтому берём несколько
    последних файлов для максимального покрытия."""
    import re
    pattern = os.path.join(excel_dir, "analytics_report_*.xlsx")
    files = sorted(glob(pattern))
    if not files:
        _safe_print(f"Нет Excel-файлов воронки для маппинга: {pattern}")
        return {}

    files = files[-max_files:]
    mapping = {}
    for f in files:
        try:
            df = pd.read_excel(f, engine="calamine", usecols=["Ozon ID", "Артикул"])
            for oid, art in zip(df["Ozon ID"].astype(str), df["Артикул"].astype(str)):
                if oid not in mapping:
                    mapping[oid] = art
        except Exception as e:
            _safe_print(f"  Пропуск {os.path.basename(f)}: {e}")

    _safe_print(f"Маппинг Ozon ID → Артикул: {len(mapping):,} записей из {len(files)} файлов")
    return mapping


def _enrich_oz_with_artikul(df, ozon_id_col="Ozon ID"):
    """Добавляет колонку 'Артикул' в DataFrame OZ-воронки на основе маппинга из Excel."""
    if "Артикул" in df.columns:
        return df
    if ozon_id_col not in df.columns:
        return df

    mapping = _build_ozon_id_artikul_map()
    if not mapping:
        return df

    df["Артикул"] = df[ozon_id_col].astype(str).map(mapping)
    matched = df["Артикул"].notna().sum()
    total = len(df)
    _safe_print(f"Обогащение Артикул: {matched:,}/{total:,} строк ({matched/total*100:.1f}%)")
    return df


# ═══════════════════════════════════════════════════════════════════════
#  Основные сценарии
# ═══════════════════════════════════════════════════════════════════════

def process_latest(mssql_table, date_column, excel_dir, sheet_name,
                   pg_schema, pg_table, skip_excel=False, skip_pg=False,
                   enrich_artikul=False):
    """Последний день: SQL Server → Excel + PostgreSQL."""
    df, report_date = read_last_day_from_mssql(mssql_table, date_column)
    if df.empty:
        _safe_print(f"Нет данных в {mssql_table}.")
        return False

    _safe_print(f"Прочитано {len(df)} строк за {report_date} из {mssql_table}")

    if enrich_artikul:
        df = _enrich_oz_with_artikul(df)

    if not skip_excel:
        save_to_excel(df, excel_dir, sheet_name, date_column)

    if not skip_pg:
        engine = _get_pg_engine()
        try:
            rows = upload_to_postgres(df, report_date, date_column, pg_schema, pg_table, engine)
            _safe_print(f"PostgreSQL: загружено {rows} строк за {report_date} в {pg_schema}.{pg_table}")
        finally:
            engine.dispose()

    return True


def process_sync(mssql_table, date_column, excel_dir, sheet_name,
                 pg_schema, pg_table, skip_excel=False, skip_pg=False,
                 enrich_artikul=False):
    """Догрузка недостающих дней: SQL Server → Excel + PostgreSQL (не позднее вчера)."""
    yesterday = date.today() - timedelta(days=1)
    mssql_dates = {d for d in read_dates_from_mssql(mssql_table, date_column) if d <= yesterday}
    _safe_print(f"В SQL Server {len(mssql_dates)} дат (до {yesterday}).")

    if skip_pg:
        missing = sorted(mssql_dates)
    else:
        engine = _get_pg_engine()
        try:
            pg_dates = _get_pg_existing_dates(engine, pg_schema, pg_table, date_column)
            missing = sorted(mssql_dates - pg_dates)
            _safe_print(f"В PostgreSQL {len(pg_dates)} дат, недостаёт {len(missing)}.")
        except Exception:
            engine.dispose()
            raise

    if not missing:
        _safe_print("Все даты уже загружены, догрузка не требуется.")
        return 0

    df = read_days_from_mssql(mssql_table, date_column, missing)

    if enrich_artikul:
        df = _enrich_oz_with_artikul(df)

    if not skip_excel:
        save_to_excel(df, excel_dir, sheet_name, date_column)

    if not skip_pg:
        total_rows = 0
        for i, dt in enumerate(missing, 1):
            df_day = df[df[date_column].dt.date == dt].copy()
            rows = upload_to_postgres(df_day, dt, date_column, pg_schema, pg_table, engine)
            total_rows += rows
            _safe_print(f"  [{i}/{len(missing)}] {dt} — {rows} строк")
        engine.dispose()
        _safe_print(f"Догрузка завершена: {total_rows} строк за {len(missing)} дат")
        return total_rows

    return len(df)


def process_full(mssql_table, date_column, excel_dir, sheet_name,
                 pg_schema, pg_table, skip_excel=False, skip_pg=False,
                 enrich_artikul=False):
    """Полная перезаливка: SQL Server → Excel + PostgreSQL (DROP + INSERT), не позднее вчера."""
    yesterday = date.today() - timedelta(days=1)
    df = read_from_mssql(mssql_table, date_column)
    if df.empty:
        _safe_print(f"Нет данных в {mssql_table}.")
        return 0

    if date_column in df.columns:
        before = len(df)
        df = df[df[date_column].dt.date <= yesterday].copy()
        if len(df) < before:
            _safe_print(f"Отфильтровано {before - len(df)} строк за даты после {yesterday}")

    _safe_print(f"Прочитано {len(df)} строк из {mssql_table} (до {yesterday})")

    if enrich_artikul:
        df = _enrich_oz_with_artikul(df)

    if not skip_excel:
        save_to_excel(df, excel_dir, sheet_name, date_column)

    if not skip_pg:
        from sqlalchemy import text
        full_table = f"{pg_schema}.{pg_table}"
        engine = _get_pg_engine()
        try:
            with engine.begin() as conn:
                if _pg_table_exists(conn, pg_schema, pg_table):
                    conn.execute(text(f"DROP TABLE {full_table}"))
                    _safe_print(f"Таблица {full_table} удалена.")

            dates = sorted(df[date_column].dropna().unique())
            total_rows = 0
            for i, dt in enumerate(dates, 1):
                df_day = df[df[date_column] == dt].copy()
                rows = upload_to_postgres(df_day, pd.Timestamp(dt).date(), date_column,
                                          pg_schema, pg_table, engine)
                total_rows += rows
                _safe_print(f"  [{i}/{len(dates)}] {pd.Timestamp(dt).date()} — {rows} строк")

            _safe_print(f"Полная загрузка: {total_rows} строк за {len(dates)} дат в {full_table}")
            return total_rows
        finally:
            engine.dispose()

    return len(df)


RECENT_LOOKBACK_DAYS = 45

def process_recent(mssql_table, date_column, excel_dir, sheet_name,
                   pg_schema, pg_table, skip_excel=False, skip_pg=False,
                   lookback_days=RECENT_LOOKBACK_DAYS, enrich_artikul=False):
    """Перезаливка последних N дней: DELETE из PG за период + INSERT свежих данных из SQL Server.
    Верхняя граница — вчера (today - 1), т.к. данные за текущий день часто неполные."""
    from sqlalchemy import text

    yesterday = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
    cutoff = (pd.Timestamp.today() - pd.DateOffset(days=lookback_days)).strftime("%Y-%m-%d")
    _safe_print(f"Режим recent: перезаливка данных с {cutoff} по {yesterday} ({lookback_days} дней)")

    conn = pyodbc.connect(MSSQL_CONN_STR)
    try:
        query = f"SELECT * FROM {mssql_table} WHERE [{date_column}] >= '{cutoff}' AND CAST([{date_column}] AS DATE) <= '{yesterday}'"
        df = pd.read_sql(query, conn)
    finally:
        conn.close()

    if date_column in df.columns:
        df[date_column] = pd.to_datetime(df[date_column], errors="coerce")

    if df.empty:
        _safe_print(f"Нет данных в {mssql_table} с {cutoff}.")
        return 0

    _safe_print(f"SQL Server: прочитано {len(df)} строк с {cutoff}")

    if enrich_artikul:
        df = _enrich_oz_with_artikul(df)

    if not skip_excel:
        save_to_excel(df, excel_dir, sheet_name, date_column)

    if not skip_pg:
        full_table = f"{pg_schema}.{pg_table}"
        engine = _get_pg_engine()
        try:
            with engine.begin() as conn:
                if _pg_table_exists(conn, pg_schema, pg_table):
                    deleted = conn.execute(
                        text(f'DELETE FROM {full_table} WHERE "{date_column}" >= :dt'),
                        {"dt": cutoff},
                    )
                    _safe_print(f"PostgreSQL: удалено строк за период с {cutoff} из {full_table}")

            dates = sorted(df[date_column].dropna().dt.date.unique())
            total_rows = 0
            for i, dt in enumerate(dates, 1):
                df_day = df[df[date_column].dt.date == dt].copy()
                rows = upload_to_postgres(df_day, dt, date_column, pg_schema, pg_table, engine)
                total_rows += rows
                _safe_print(f"  [{i}/{len(dates)}] {dt} — {rows} строк")

            _safe_print(f"Перезаливка завершена: {total_rows} строк за {len(dates)} дат в {full_table}")
            return total_rows
        finally:
            engine.dispose()

    return len(df)


def process_latest_oz_costs(skip_pg=False):
    """Последний день затрат Ozon (не позднее вчера): Excel → PostgreSQL."""
    yesterday = date.today() - timedelta(days=1)
    df, report_date = read_latest_oz_costs_from_excel(
        OZ_COSTS_EXCEL_SOURCE_DIR, OZ_COSTS_EXCEL_SHEET, target_date=yesterday
    )
    if df.empty:
        _safe_print("Нет данных по затратам Ozon.")
        return False

    if not skip_pg:
        engine = _get_pg_engine()
        try:
            rows = upload_to_postgres(
                df, report_date, OZ_COSTS_DATE_COLUMN,
                OZ_COSTS_PG_SCHEMA, OZ_COSTS_PG_TABLE, engine
            )
            _safe_print(f"PostgreSQL: загружено {rows} строк за {report_date} в {OZ_COSTS_PG_FULL}")
        finally:
            engine.dispose()

    return True


def process_recent_oz_costs(skip_pg=False, lookback_days=RECENT_LOOKBACK_DAYS):
    """Перезаливка последних N дней затрат Ozon: DELETE из PG + INSERT из Excel."""
    from sqlalchemy import text

    yesterday = date.today() - timedelta(days=1)
    cutoff = (yesterday - timedelta(days=lookback_days))
    _safe_print(f"Режим recent (oz-costs): перезаливка данных с {cutoff} по {yesterday} ({lookback_days} дней)")

    df = read_oz_costs_from_excel(
        OZ_COSTS_EXCEL_SOURCE_DIR, OZ_COSTS_EXCEL_SHEET,
        cutoff_date=cutoff, upper_date=yesterday
    )

    if df.empty:
        _safe_print(f"Нет файлов затрат Ozon с {cutoff}.")
        return 0

    if not skip_pg:
        engine = _get_pg_engine()
        try:
            with engine.begin() as conn:
                if _pg_table_exists(conn, OZ_COSTS_PG_SCHEMA, OZ_COSTS_PG_TABLE):
                    conn.execute(
                        text(f'DELETE FROM {OZ_COSTS_PG_FULL} WHERE "{OZ_COSTS_DATE_COLUMN}" >= :dt'),
                        {"dt": cutoff},
                    )
                    _safe_print(f"PostgreSQL: удалены данные с {cutoff} из {OZ_COSTS_PG_FULL}")

            dates = sorted(df[OZ_COSTS_DATE_COLUMN].dropna().dt.date.unique())
            total_rows = 0
            for i, dt in enumerate(dates, 1):
                df_day = df[df[OZ_COSTS_DATE_COLUMN].dt.date == dt].copy()
                rows = upload_to_postgres(
                    df_day, dt, OZ_COSTS_DATE_COLUMN,
                    OZ_COSTS_PG_SCHEMA, OZ_COSTS_PG_TABLE, engine
                )
                total_rows += rows
                _safe_print(f"  [{i}/{len(dates)}] {dt} — {rows} строк")

            _safe_print(f"Перезаливка завершена: {total_rows} строк за {len(dates)} дат в {OZ_COSTS_PG_FULL}")
            return total_rows
        finally:
            engine.dispose()

    return len(df)


def process_sync_oz_costs(skip_pg=False):
    """Догрузка недостающих дней затрат Ozon: Excel → PostgreSQL (только новые даты, не позднее вчера)."""
    yesterday = date.today() - timedelta(days=1)
    pattern = os.path.join(OZ_COSTS_EXCEL_SOURCE_DIR, "Аналитика продвижения_*.xlsx")
    files = glob(pattern)
    if not files:
        _safe_print(f"Нет файлов по шаблону: {pattern}")
        return 0

    excel_dates = {}
    for f in files:
        try:
            data_date = _parse_oz_costs_date(f)
            if data_date > yesterday:
                continue
            excel_dates[data_date] = f
        except (ValueError, IndexError):
            continue

    if not excel_dates:
        _safe_print("Не удалось определить даты ни в одном файле.")
        return 0

    _safe_print(f"В Excel найдено {len(excel_dates)} дат.")

    if skip_pg:
        missing = sorted(excel_dates.keys())
    else:
        engine = _get_pg_engine()
        try:
            pg_dates = _get_pg_existing_dates(
                engine, OZ_COSTS_PG_SCHEMA, OZ_COSTS_PG_TABLE, OZ_COSTS_DATE_COLUMN
            )
            pg_dates_normalized = {
                d.date() if hasattr(d, 'date') else d for d in pg_dates
            }
            missing = sorted(set(excel_dates.keys()) - pg_dates_normalized)
            _safe_print(f"В PostgreSQL {len(pg_dates)} дат, недостаёт {len(missing)}.")
        except Exception:
            engine.dispose()
            raise

    if not missing:
        _safe_print("Все даты уже загружены, догрузка не требуется.")
        return 0

    if not skip_pg:
        total_rows = 0
        for i, dt in enumerate(missing, 1):
            f = excel_dates[dt]
            df = pd.read_excel(f, sheet_name=OZ_COSTS_EXCEL_SHEET, engine="calamine", skiprows=1)
            df[OZ_COSTS_DATE_COLUMN] = pd.Timestamp(dt)
            rows = upload_to_postgres(
                df, dt, OZ_COSTS_DATE_COLUMN,
                OZ_COSTS_PG_SCHEMA, OZ_COSTS_PG_TABLE, engine
            )
            total_rows += rows
            _safe_print(f"  [{i}/{len(missing)}] {dt} — {rows} строк")
        engine.dispose()
        _safe_print(f"Догрузка завершена: {total_rows} строк за {len(missing)} дат в {OZ_COSTS_PG_FULL}")
        return total_rows

    return len(missing)


def try_process(marketplace, mode="latest", skip_excel=False, skip_pg=False):
    """
    Безопасная обёртка. Никогда не бросает исключение.
    marketplace: 'oz' | 'wb' | 'wb-costs'
    mode: 'latest' | 'sync' | 'full' | 'recent'
    """
    try:
        _safe_print(f"\n{'=' * 60}")
        _safe_print(f"  {marketplace.upper()} — режим: {mode}")
        _safe_print(f"{'=' * 60}")

        if marketplace == "oz-costs":
            if mode == "latest":
                process_latest_oz_costs(skip_pg=skip_pg)
            elif mode == "recent":
                process_recent_oz_costs(skip_pg=skip_pg)
            elif mode == "sync":
                process_sync_oz_costs(skip_pg=skip_pg)
            else:
                _safe_print(f"Режим '{mode}' не поддерживается для oz-costs (доступны: latest, recent, sync)")
                return False
            return True

        configs = {
            "oz": (OZ_MSSQL_TABLE, OZ_DATE_COLUMN, OZ_EXCEL_DIR, OZ_EXCEL_SHEET, OZ_PG_SCHEMA, OZ_PG_TABLE),
            "wb": (WB_MSSQL_TABLE, WB_DATE_COLUMN, WB_EXCEL_DIR, WB_EXCEL_SHEET, WB_PG_SCHEMA, WB_PG_TABLE),
            "wb-costs": (WB_COSTS_MSSQL_TABLE, WB_COSTS_DATE_COLUMN, WB_COSTS_EXCEL_DIR, WB_COSTS_EXCEL_SHEET, WB_COSTS_PG_SCHEMA, WB_COSTS_PG_TABLE),
        }
        cfg = configs[marketplace]

        dispatch = {
            "latest": process_latest,
            "sync": process_sync,
            "full": process_full,
            "recent": process_recent,
        }
        extra_kw = {}
        if marketplace == "oz":
            extra_kw["enrich_artikul"] = True
        dispatch[mode](*cfg, skip_excel=skip_excel, skip_pg=skip_pg, **extra_kw)
        return True
    except Exception as e:
        msg = f"Ошибка [{marketplace.upper()}]: {e}"
        _safe_print(msg)
        logger.warning(msg)
        return False


# ═══════════════════════════════════════════════════════════════════════
#  CLI
# ═══════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    import argparse
    parser = argparse.ArgumentParser(
        description="Выгрузка воронки продаж из SQL Server в Excel и PostgreSQL"
    )

    mp_group = parser.add_argument_group("Маркетплейс / таблица")
    mp_group.add_argument("--oz", action="store_true", help="Ozon воронка")
    mp_group.add_argument("--oz-costs", action="store_true", help="Ozon затраты (из Excel)")
    mp_group.add_argument("--wb", action="store_true", help="WB воронка")
    mp_group.add_argument("--wb-costs", action="store_true", help="WB затраты (wb_marketing)")
    mp_group.add_argument("--all", action="store_true", help="Все таблицы")

    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument("--full", action="store_true", help="Полная перезаливка (DROP + INSERT)")
    mode_group.add_argument("--sync", action="store_true", help="Догрузить недостающие дни")
    mode_group.add_argument("--recent", action="store_true", help="Перезалить последние 45 дней (DELETE + INSERT)")

    out_group = parser.add_argument_group("Выходы")
    out_group.add_argument("--no-excel", action="store_true", help="Не сохранять Excel")
    out_group.add_argument("--no-pg", action="store_true", help="Не загружать в PostgreSQL")

    args = parser.parse_args()

    if not args.oz and not args.oz_costs and not args.wb and not args.wb_costs and not args.all:
        parser.error("Укажите таблицу: --oz, --oz-costs, --wb, --wb-costs или --all")

    mode = "full" if args.full else ("sync" if args.sync else ("recent" if args.recent else "latest"))
    marketplaces = []
    if args.all:
        marketplaces = ["oz", "oz-costs", "wb", "wb-costs"]
    else:
        if args.oz:
            marketplaces.append("oz")
        if args.oz_costs:
            marketplaces.append("oz-costs")
        if args.wb:
            marketplaces.append("wb")
        if args.wb_costs:
            marketplaces.append("wb-costs")

    ok = True
    for mp in marketplaces:
        result = try_process(mp, mode=mode, skip_excel=args.no_excel, skip_pg=args.no_pg)
        ok = ok and result

    raise SystemExit(0 if ok else 1)
