import os
import sys
import ctypes
import subprocess

if os.name == 'nt':  # Только для Windows
    from ctypes import wintypes # Импортируем wintypes отдельно
    
    # Константы Windows API
    ENABLE_EXTENDED_FLAGS = 0x0080
    ENABLE_QUICK_EDIT_MODE = 0x0040
    STD_INPUT_HANDLE = -10

    h_input = ctypes.windll.kernel32.GetStdHandle(STD_INPUT_HANDLE)
    mode = ctypes.wintypes.DWORD()
    
    if ctypes.windll.kernel32.GetConsoleMode(h_input, ctypes.byref(mode)):
        # Отключаем QuickEdit Mode
        new_mode = mode.value & ~ENABLE_QUICK_EDIT_MODE
        # Важно также включить расширенные флаги
        new_mode |= ENABLE_EXTENDED_FLAGS
        ctypes.windll.kernel32.SetConsoleMode(h_input, new_mode)

# 1. Берем путь к папке, где лежит этот файл
base_path = os.path.dirname(os.path.abspath(__file__))
# 2. Принудительно переходим в эту папку
os.chdir(base_path)
# 3. Добавляем путь в системные пути поиска модулей
if base_path not in sys.path:
    sys.path.insert(0, base_path)
# 4. Для надежности: если папка Preprocessing внутри, добавим и её
prep_path = os.path.join(base_path, 'Preprocessing')
if os.path.exists(prep_path) and prep_path not in sys.path:
    sys.path.insert(0, prep_path)

from datetime import date, datetime, timedelta
import socket
import time

# Direct imports of wrapper scripts
import Preprocessing.Dashboard_OZON_Wrapper as Ozon_Wrapper
from Preprocessing.ozon_ad_report_downloader import download_ozon_ad_report
from Preprocessing.ozon_analytics_downloader import download_ozon_analytics_report
from Preprocessing.ozon_funnel_api_downloader import download_ozon_funnel_report
from DashboardDBUploaderOZ import try_upload_latest_to_postgres as try_upload_dashboard_to_postgres

import DashboardAssemblyOZ as DashboardAssemblyOZ

import logging

# Настройка логирования: файл, формат, уровень
logging.basicConfig(
    filename='assembly_ozon.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8' # Для корректного отображения кириллицы
)

output_dir=r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Taldykin"

def _is_port_open(host="localhost", port=15432, timeout=2):
    """Проверяет, слушается ли порт."""
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except (ConnectionRefusedError, OSError, TimeoutError):
        return False


def ensure_postgres_tunnel(max_wait=15):
    """Проверяет SSH-туннель к PostgreSQL. Если не поднят — запускает и ждёт."""
    if _is_port_open():
        logging.info("SSH-туннель PostgreSQL уже активен (порт 15432)")
        return True

    logging.info("SSH-туннель не обнаружен — запускаю...")
    print("SSH-туннель PostgreSQL не активен — запускаю...")

    ssh_exe = r"C:\Program Files\Git\usr\bin\ssh.exe"
    if not os.path.exists(ssh_exe):
        ssh_exe = r"C:\Users\i.taldykin\AppData\Local\Programs\Git\usr\bin\ssh.exe"

    subprocess.Popen(
        [ssh_exe, "-o", "StrictHostKeyChecking=no",
         "-N", "-L", "15432:localhost:25432", "cc@ru-mow-d1-sap41"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )

    for i in range(max_wait):
        time.sleep(1)
        if _is_port_open():
            logging.info(f"SSH-туннель поднят за {i + 1} сек.")
            print(f"SSH-туннель поднят за {i + 1} сек.")
            return True

    logging.warning(f"SSH-туннель не поднялся за {max_wait} сек.")
    print(f"⚠ SSH-туннель не поднялся за {max_wait} сек.")
    return False


def main():
    def run_postgres_tunnel():
        subprocess.run([r'pgadmin_connect.bat'])

    # try:
    #     download_ozon_analytics_report() # Deprecated scripted web downloading
    #     # download_ozon_funnel_report() # API Call function
    #     download_ozon_ad_report() # Deprecated scripted web downloading
    #     print("OZON Analytics and Ad Report downloaded successfully")
    #     logging.info("OZON Analytics and Ad Report downloaded successfully")
    # except Exception as e:
    #     print(f"Error in download_ozon_analytics_report or download_ozon_ad_report: {e}")
    #     logging.info(f"Error in download_ozon_analytics_report or download_ozon_ad_report: {e}")
    #     return False

    # wrapper_modules = {  # pyright: ignore[reportUnreachable]
    #         "wrapper": ("Dashboard_OZON_Wrapper.py", Ozon_Wrapper),
    #     }
        
    # wrapper_file, wrapper_module = wrapper_modules["wrapper"]
        
    # logging.info(f"▶️ Запуск скрипта: {wrapper_file}")
    # logging.info("─" * 60)
    
    # try:
    #     # Вызываем main() функцию напрямую
    #     logging.info("🔄 Выполнение скрипта обработки данных OZON...")
    #     wrapper_module.main()
    #     logging.info("✅ Скрипт выполнен успешно")
        
    #     logging.info("─" * 60)
    #     logging.info("🎉 ВСЕ ЭТАПЫ УСПЕШНО ВЫПОЛНЕНЫ!")
        
    #     # Записываем в лог-файл
    #     logging.info(f'{datetime.now()} - {wrapper_file} completed via GUI\n')
        
    # except Exception as e:
    #     error_message = f"❌ ОШИБКА: {str(e)}"
    #     logging.info(error_message)
    #     # Записываем ошибку в лог-файл
    #     logging.info(f'{datetime.now()} - ERROR in {wrapper_file}: {str(e)}\n')
    #     return False


    # Перед сборкой SQL-based дашборда убеждаемся, что SSH-туннель к PG поднят
    logging.info("▶️ Проверка SSH-туннеля PostgreSQL перед сборкой...")
    tunnel_ok = ensure_postgres_tunnel()
    if not tunnel_ok:
        logging.warning("⚠️ SSH-туннель не поднялся — пробуем через pgadmin_connect.bat...")
        run_postgres_tunnel()
        tunnel_ok = _is_port_open()

    if not tunnel_ok:
        logging.error("❌ PostgreSQL недоступен — сборка SQL-based дашборда невозможна")
        print("❌ PostgreSQL недоступен — сборка SQL-based дашборда невозможна")
        return False

    logging.info("▶️ Запуск скрипта: DashboardAssemblyOZ_SQLBased.py")
    logging.info("📦 Сборка дашборда OZON (SQL-based)...")
    logging.info("─" * 60)
    
    try:
        logging.info("🔄 Выполнение скрипта сборки дашборда OZON...")
        DashboardAssemblyOZ.assemble()
        logging.info("✅ Скрипт выполнен успешно")
        
        logging.info("─" * 60)
        logging.info("🎉 СБОРКА ДАШБОРДА ЗАВЕРШЕНА!")
        
        logging.info('DashboardAssemblyOZ_SQLBased.py completed via Pipeline')

        logging.info("▶️ Загрузка последнего дня дашборда в PostgreSQL...")
        try:
            ok = try_upload_dashboard_to_postgres()
            if ok:
                logging.info("✅ Загрузка дашборда в PostgreSQL завершена успешно")
            else:
                logging.warning("⚠️ Загрузка дашборда в PostgreSQL произошла с ошибкой. Поднимаем туннель...")
                ensure_postgres_tunnel()
                ok = try_upload_dashboard_to_postgres()
                if ok:
                    logging.info("✅ Загрузка дашборда в PostgreSQL завершена успешно")
                else:
                    logging.warning("⚠️ Загрузка дашборда в PostgreSQL пропущена (см. лог)")
        except Exception as db_err:
            logging.warning(f"⚠️ Загрузка дашборда в PostgreSQL пропущена: {db_err}")
        
        logging.info("Успех", "Скрипт успешно выполнен!")
        
    except Exception as e:
        error_message = f"❌ ОШИБКА: {str(e)}"
        logging.info(error_message)
        logging.info("Ошибка выполнения", str(e))
        
        # Записываем ошибку в лог-файл
        logging.info(f'{datetime.now()} - ERROR in DashboardAssemblyOZ.py: {str(e)}\n')
        return False

if __name__ == "__main__":
    main()