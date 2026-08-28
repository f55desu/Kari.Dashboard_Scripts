# -*- coding: utf-8 -*-
"""
SQL_Exporter_Wrapper.py
========================
Выгрузка данных из SQL Server (DBReport) и Excel
в PostgreSQL (analytics) для всех таблиц WB и OZ.

Таблицы:
    [mp].[ozon_sales_funnel]   ->  work.ozon_sales_funnel     (OZ воронка, из SQL Server)
    Excel (New Format)         ->  work.ozon_costs_statistics  (OZ затраты, из Excel)
    [mp].[wb_sales_funnel_lk]  ->  work.wb_sales_funnel_lk    (WB воронка, из SQL Server)
    [mp].[wb_marketing]        ->  work.wb_marketing           (WB затраты, из SQL Server)

Режимы:
    recent  (по умолчанию) -- DELETE за последние 45 дней + INSERT свежих данных
    sync                   -- догрузить только недостающие даты
    latest                 -- только последний день
    full                   -- полная перезаливка (DROP + INSERT)

Excel-файлы не создаются по умолчанию.

Использование:
    python SQL_Exporter_Wrapper.py                # recent: WB воронка + затраты
    python SQL_Exporter_Wrapper.py --all          # recent: все таблицы
    python SQL_Exporter_Wrapper.py --sync         # sync: WB воронка + затраты
    python SQL_Exporter_Wrapper.py --sync --all   # sync: все таблицы
    python SQL_Exporter_Wrapper.py --latest --oz  # latest: OZ воронка (1 день)
"""

import sys
import logging

from Preprocessing.funnel_sql_exporter import try_process

def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )

    import argparse
    parser = argparse.ArgumentParser(
        description="Выгрузка данных из DBReport / Excel в PostgreSQL"
    )
    parser.add_argument("--all", action="store_true",
                        help="Все таблицы (OZ воронка + OZ затраты + WB воронка + WB затраты)")
    parser.add_argument("--oz", action="store_true", help="Ozon воронка")
    parser.add_argument("--oz-costs", action="store_true", help="Ozon затраты (из Excel)")
    parser.add_argument("--wb", action="store_true", help="WB воронка")
    parser.add_argument("--wb-costs", action="store_true", help="WB затраты")
    parser.add_argument("--with-excel", action="store_true",
                        help="Также сохранить Excel-файлы (по умолчанию не сохраняются)")

    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument("--sync", action="store_true",
                            help="Догрузить только недостающие даты")
    mode_group.add_argument("--latest", action="store_true",
                            help="Только последний день")
    mode_group.add_argument("--full", action="store_true",
                            help="Полная перезаливка (DROP + INSERT)")
    mode_group.add_argument("--recent", action="store_true",
                                help="Полная перезаливка последних 45 дней (DROP + INSERT)")

    args = parser.parse_args()

    if args.sync:
        mode = "sync"
    elif args.latest:
        mode = "latest"
    elif args.full:
        mode = "full"
    else:
        mode = "recent"

    if not args.all and not args.oz and not args.oz_costs and not args.wb and not args.wb_costs:
        tables = ["wb", "wb-costs"]
    elif args.all:
        tables = ["oz", "oz-costs", "wb", "wb-costs"]
    else:
        tables = []
        if args.oz:
            tables.append("oz")
        if args.oz_costs:
            tables.append("oz-costs")
        if args.wb:
            tables.append("wb")
        if args.wb_costs:
            tables.append("wb-costs")

    skip_excel = not args.with_excel
    ok = True

    for tbl in tables:
        result = try_process(tbl, mode=mode, skip_excel=skip_excel, skip_pg=False)
        ok = ok and result

    if ok:
        print(f"\nВсе таблицы успешно обработаны (режим: {mode}).")
    else:
        print("\nНекоторые таблицы не удалось обработать (см. ошибки выше).")

    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
