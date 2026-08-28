# -*- coding: utf-8 -*-
"""
Funnel_Costs_PreprocessingOZ_Autopipeline.py
=============================================
Пайплайн препроцессинга Ozon для планировщика задач Windows.

Этапы:
    1. Скачивание отчётов Ozon (аналитика + реклама)
    2. Запуск Dashboard_OZON_Wrapper (препроцессинг Excel)
    3. Выгрузка последнего дня в PostgreSQL (oz воронка + oz затраты)

Запускается ПЕРЕД сборкой дашборда (DashboardAssemblyOZ_AutoPipeline.py).
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

from datetime import datetime

import Preprocessing.Dashboard_OZON_Wrapper as Ozon_Wrapper
from Preprocessing.ozon_ad_report_downloader import download_ozon_ad_report
from Preprocessing.ozon_analytics_downloader import download_ozon_analytics_report
from Preprocessing.funnel_sql_exporter import try_process

import logging

logging.basicConfig(
    filename='preprocessing_ozon.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)


def run_postgres_tunnel():
    subprocess.run([r'pgadmin_connect.bat'])


def main():
    logging.info("=" * 60)
    logging.info("Запуск Funnel_Costs_PreprocessingOZ_Autopipeline")
    logging.info("=" * 60)

    # --- Этап 1: Скачивание отчётов Ozon ---
    try:
        download_ozon_analytics_report()
        download_ozon_ad_report()
        logging.info("Отчёты Ozon скачаны успешно")
    except Exception as e:
        logging.error(f"Ошибка при скачивании отчётов Ozon: {e}")
        return 1

    # --- Этап 2: Препроцессинг (Dashboard_OZON_Wrapper) ---
    try:
        logging.info("Запуск Dashboard_OZON_Wrapper...")
        Ozon_Wrapper.main()
        logging.info("Dashboard_OZON_Wrapper выполнен успешно")
    except Exception as e:
        logging.error(f"Ошибка в Dashboard_OZON_Wrapper: {e}")
        return 1

    # --- Этап 3: Выгрузка последнего дня в PostgreSQL ---
    logging.info("Выгрузка Ozon данных в PostgreSQL (режим: latest)...")

    ok = True
    for tbl in ["oz", "oz-costs"]:
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
        logging.info("Препроцессинг Ozon завершён успешно")
    else:
        logging.warning("Препроцессинг Ozon завершён с ошибками")

    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
