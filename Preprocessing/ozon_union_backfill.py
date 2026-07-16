# -*- coding: utf-8 -*-
"""
ozon_union_backfill.py
======================
Первоначальная загрузка ВСЕХ Excel-файлов "Аналитика продвижения" (лист "Union")
из папки-источника в PostgreSQL: work.ozon_promo_union.

Запуск (туннель должен быть поднят, config заполнен):
    python ozon_union_backfill.py                 # все файлы, идемпотентно (delete+insert по дате)
    python ozon_union_backfill.py --truncate      # очистить таблицу перед загрузкой
    python ozon_union_backfill.py --dry-run        # прочитать и посчитать строки, без записи в БД
    python ozon_union_backfill.py --batch-size 10000

Идемпотентность: для каждого файла строки за его report_date сначала удаляются,
затем вставляются заново. Повторный прогон не плодит дубли.
"""

import sys
import io
import argparse
import logging

# Корректный вывод кириллицы/₽ в Windows-консоль
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")
except Exception:
    pass

import pandas as pd

import ozon_union_db as db

logger = logging.getLogger("ozon_union.backfill")


def parse_args():
    p = argparse.ArgumentParser(description="Backfill work.ozon_promo_union из Excel-файлов Union")
    p.add_argument("--config", default=None, help="путь к postgres_config.json")
    p.add_argument("--folder", default=db.SOURCE_FOLDER, help="папка с Excel-файлами")
    p.add_argument("--batch-size", type=int, default=db.DEFAULT_BATCH_SIZE)
    p.add_argument("--truncate", action="store_true",
                   help="TRUNCATE таблицы перед загрузкой (полная перезаливка)")
    p.add_argument("--dry-run", action="store_true",
                   help="только прочитать файлы и посчитать строки, без записи в БД")
    return p.parse_args()


def main():
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s")
    args = parse_args()

    files = db.list_source_files(args.folder)
    if not files:
        print(f"❌ Нет файлов по шаблону в папке: {args.folder}")
        return 1
    print(f"Найдено файлов: {len(files)}")

    # ---- DRY RUN: только чтение ----
    if args.dry_run:
        total = 0
        for i, f in enumerate(files, 1):
            try:
                d = db.read_union_file(f)
                total += len(d)
                print(f"[{i:>2}/{len(files)}] {d['report_date'].iloc[0]}  "
                      f"{len(d):>6} строк  | {db.os.path.basename(f)}")
            except Exception as e:
                print(f"[{i:>2}/{len(files)}] ОШИБКА чтения {db.os.path.basename(f)}: {e}")
        print(f"\nИТОГО строк (dry-run): {total}")
        return 0

    # ---- Реальная загрузка ----
    cfg = db.load_db_config(args.config)
    print(f"Подключение: {cfg['user']}@{cfg['host']}:{cfg['port']}/{cfg['dbname']}")
    engine = db.get_engine(cfg)

    try:
        db.ensure_table(engine)

        if args.truncate:
            from sqlalchemy import text
            with engine.begin() as conn:
                conn.execute(text(f"TRUNCATE TABLE {db.FULL_TABLE}"))
            print(f"🧹 Таблица {db.FULL_TABLE} очищена (TRUNCATE).")

        total = 0
        ok_files = 0
        for i, f in enumerate(files, 1):
            name = db.os.path.basename(f)
            try:
                d = db.read_union_file(f)
                rows = db.load_dataframe(engine, d,
                                         batch_size=args.batch_size,
                                         replace_dates=True)
                total += rows
                ok_files += 1
                print(f"[{i:>2}/{len(files)}] {d['report_date'].iloc[0]}  "
                      f"загружено {rows:>6} строк  | {name}")
            except Exception as e:
                print(f"[{i:>2}/{len(files)}] ❌ ОШИБКА {name}: {e}")

        # Контроль
        cnt = pd.read_sql(f"SELECT count(*) AS n FROM {db.FULL_TABLE}", engine)["n"].iloc[0]
        print(f"\n✅ Готово. Файлов успешно: {ok_files}/{len(files)}. "
              f"Загружено строк за прогон: {total}. Всего в таблице: {cnt}.")
    finally:
        engine.dispose()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
