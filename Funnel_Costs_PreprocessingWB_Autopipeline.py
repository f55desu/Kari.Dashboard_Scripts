# -*- coding: utf-8 -*-
"""
Funnel_Costs_PreprocessingWB_Autopipeline.py
=============================================
Пайплайн препроцессинга WB для планировщика задач Windows.

Этапы:
    1. Скачивание отчётов WB (воронка + затраты)
    2. Запуск Dashboard_WB_Wrapper (препроцессинг Excel)
    3. Выгрузка последнего дня в PostgreSQL (wb воронка + wb затраты)

Запускается ПЕРЕД сборкой дашборда (DashboardAssemblyWB_AutoPipeline.py).
"""

import os
import sys
import ctypes
import subprocess

if os.name == 'nt':
    from ctypes import wintypes

    ENABLE_EXTENDED_FLAGS = 0x0080
    ENABLE_QUICK_EDIT_MODE = 0x0040
    STD_INPUT_HANDLE = -10

    h_input = ctypes.windll.kernel32.GetStdHandle(STD_INPUT_HANDLE)
    mode = ctypes.wintypes.DWORD()

    if ctypes.windll.kernel32.GetConsoleMode(h_input, ctypes.byref(mode)):
        new_mode = mode.value & ~ENABLE_QUICK_EDIT_MODE
        new_mode |= ENABLE_EXTENDED_FLAGS
        ctypes.windll.kernel32.SetConsoleMode(h_input, new_mode)

base_path = os.path.dirname(os.path.abspath(__file__))
os.chdir(base_path)
if base_path not in sys.path:
    sys.path.insert(0, base_path)
prep_path = os.path.join(base_path, 'Preprocessing')
if os.path.exists(prep_path) and prep_path not in sys.path:
    sys.path.insert(0, prep_path)

from datetime import date, datetime

import Preprocessing.Dashboard_WB_Wrapper as WB_Wrapper
import Preprocessing.Dashboard_WB_Wrapper_Tuesday as WB_Wrapper_Tuesday
from Preprocessing.wb_downloader_v2 import download_wb_report_v2
from Preprocessing.wb_campaign_downloader import download_wb_campaign_report
from Preprocessing.funnel_sql_exporter import try_process

import logging

logging.basicConfig(
    filename='preprocessing_wb.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)


def run_postgres_tunnel():
    subprocess.run([r'pgadmin_connect.bat'])


def main():
    logging.info("=" * 60)
    logging.info("Запуск Funnel_Costs_PreprocessingWB_Autopipeline")
    logging.info("=" * 60)

    # # --- Этап 1: Скачивание отчётов WB ---
    # try:
    #     download_wb_report_v2()
    #     download_wb_campaign_report()
    #     logging.info("Отчёты WB скачаны успешно")
    # except Exception as e:
    #     logging.error(f"Ошибка при скачивании отчётов WB: {e}")
    #     return 1

    # # --- Этап 2: Препроцессинг (Dashboard_WB_Wrapper) ---
    # if date.today().weekday() == 1:
    #     wrapper_name = "Dashboard_WB_Wrapper_Tuesday"
    #     wrapper_module = WB_Wrapper_Tuesday
    # else:
    #     wrapper_name = "Dashboard_WB_Wrapper"
    #     wrapper_module = WB_Wrapper

    # try:
    #     logging.info(f"Запуск {wrapper_name}...")
    #     wrapper_module.main()
    #     logging.info(f"{wrapper_name} выполнен успешно")
    # except Exception as e:
    #     logging.error(f"Ошибка в {wrapper_name}: {e}")
    #     return 1

    # --- Этап 3: Выгрузка последнего дня в PostgreSQL ---
    logging.info("Выгрузка WB данных в PostgreSQL (режим: latest)...")

    ok = True
    for tbl in ["wb", "wb-costs"]:
        try:
            result = try_process(tbl, mode="latest", skip_excel=True, skip_pg=False)
            if result:
                logging.info(f"  {tbl}: выгрузка в PG успешна")
            else:
                logging.warning(f"  {tbl}: ошибка выгрузки, пробуем поднять туннель PG...")
                run_postgres_tunnel()
                result = try_process(tbl, mode="latest", skip_excel=True, skip_pg=False)
                if result:
                    logging.info(f"  {tbl}: выгрузка в PG успешна (после поднятия туннеля)")
                else:
                    logging.error(f"  {tbl}: выгрузка в PG не удалась")
                    ok = False
        except Exception as e:
            logging.error(f"  {tbl}: исключение при выгрузке: {e}")
            ok = False

    if ok:
        logging.info("Препроцессинг WB завершён успешно")
    else:
        logging.warning("Препроцессинг WB завершён с ошибками")

    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
