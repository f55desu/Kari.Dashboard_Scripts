"""
ozon_funnel_api_downloader.py — Выгружает «воронку продаж» Ozon через Seller API.

Аналог ozon_analytics_downloader.py (который тянет тот же отчёт браузером со
страницы Аналитика → Графики, кнопка «Скачать» → «Отчёт за период»), но без
браузера: данные забираются напрямую через Ozon Seller API и сохраняются в
xlsx-файл С ТЕМИ ЖЕ КОЛОНКАМИ и тем же шаблоном имени файла
(analytics_report_YYYY-MM-DD_HH_MM.xlsx), что и браузерный скрипт, — чтобы
быть drop-in заменой для Dashboard_OZON_Wrapper.py.

Колонки итогового файла (18, 1:1 с актуальной браузерной выгрузкой от 17.06.2026):
    Товары | Модель | Ozon ID | Артикул | День |
    Показы, всего | Показы на карточке товара | В корзину, всего | Заказано товаров |
    Показы в поиске и каталоге | Позиция в поиске и каталоге |
    В корзину из карточки товара | В корзину из поиска или каталога |
    Отменено товаров | Доставлено товаров | Возвращено товаров |
    Заказано на сумму | Конверсия в корзину, общая

Источники данных Ozon API:
    POST /v1/analytics/data        — метрики воронки, dimension = [sku, modelID]
                                     (даёт Товары, Ozon ID, Модель и 13 метрик,
                                      включая conv_tocart = «Конверсия в корзину»)
    POST /v3/product/info/list     — сопоставление sku → offer_id (колонка Артикул)
«День» = дата данных (target_date). «Позиция» и «Конверсия» пишутся текстом с
запятой ("24,86" / "0,17%"), как в браузерном экспорте.

Входные параметры — те же, что у браузерного скрипта:
    python ozon_funnel_api_downloader.py [--output-dir PATH] [--date YYYY-MM-DD]

  --date        дата отчёта (по умолчанию вчера). В отличие от браузерного
                скрипта здесь дата РЕАЛЬНО задаёт период выгрузки (date_from=date_to).
  --output-dir  куда сохранить файл (по умолчанию сетевая шара Taldykin).

Креды Ozon берутся из .env рядом со скриптом (OZON_CLIENT_ID / OZON_API_KEY).
"""
import sys
import os
import json
import time
import logging
from datetime import date, timedelta, datetime
from pathlib import Path

import requests
import pandas as pd

if sys.stdout is not None:
    sys.stdout.reconfigure(encoding="utf-8")
if sys.stderr is not None:
    sys.stderr.reconfigure(encoding="utf-8")

PROJECT_DIR = Path(__file__).parent

# Ключи API берём из .env рядом со скриптом (python-dotenv). Если пакета нет —
# просто работаем с уже выставленными переменными окружения.
try:
    from dotenv import load_dotenv
    load_dotenv(PROJECT_DIR / ".env")
except Exception:
    pass

downloads_folder = r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Taldykin"
DEFAULT_DOWNLOADS_DIR = Path(downloads_folder)

LOG_FILE = os.path.join(PROJECT_DIR, "ozon_funnel_api.log")

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'  # Для корректного отображения кириллицы
)
log = logging.getLogger(__name__)

# --- Ozon Seller API ---
API_BASE = "https://api-seller.ozon.ru"
ANALYTICS_URL = f"{API_BASE}/v1/analytics/data"
PRODUCT_INFO_LIST_URL = f"{API_BASE}/v3/product/info/list"

# Креды Seller API — из .env (OZON_CLIENT_ID / OZON_API_KEY). Единственный источник.
OZON_CLIENT_ID = os.environ.get("OZON_CLIENT_ID", "").strip()
OZON_API_KEY = os.environ.get("OZON_API_KEY", "").strip()

HEADERS = {
    "Client-Id": OZON_CLIENT_ID,
    "Api-Key": OZON_API_KEY,
    "Content-Type": "application/json",
}

# Лимит Ozon — 2 запроса/сек на клиента. Держим паузу с запасом + ретраи.
# 429 (rate limit) ретраим почти бесконечно с длинным бэкоффом: лимит общий
# на клиента, и если параллельно работает другой скрипт (ingestor/wrapper),
# 429 может сыпаться долго — сдаваться нельзя, иначе теряем всю выгрузку.
REQUEST_PAUSE = 0.7
MAX_RETRIES = 8            # для сетевых ошибок и 5xx
MAX_RETRIES_429 = 120      # отдельный, большой бюджет ретраев для 429
SERVER_ERROR_STATUS = {500, 502, 503, 504}

PAGE_LIMIT = 1000          # макс. строк за один вызов analytics/data
MAX_PAGES = 1000           # предохранитель от бесконечной пагинации
INFO_BATCH = 1000          # sku за один вызов product/info/list

# Колонки итогового файла — порядок и названия 1:1 с актуальной браузерной
# выгрузкой (18 колонок, формат с 17.06.2026: добавлены «День» и
# «Конверсия в корзину, общая», изменён порядок метрик).
OUTPUT_COLUMNS = [
    "Товары",
    "Модель",
    "Ozon ID",
    "Артикул",
    "День",
    "Показы, всего",
    "Показы на карточке товара",
    "В корзину, всего",
    "Заказано товаров",
    "Показы в поиске и каталоге",
    "Позиция в поиске и каталоге",
    "В корзину из карточки товара",
    "В корзину из поиска или каталога",
    "Отменено товаров",
    "Доставлено товаров",
    "Возвращено товаров",
    "Заказано на сумму",
    "Конверсия в корзину, общая",
]

# Метрики analytics/data. Порядок здесь = индексы m[...] в build_dataframe
# (НЕ совпадает с порядком колонок в файле — раскладку делает build_dataframe).
METRICS = [
    "hits_view",          # Показы, всего
    "hits_view_pdp",      # Показы на карточке товара
    "hits_view_search",   # Показы в поиске и каталоге
    "position_category",  # Позиция в поиске и каталоге
    "hits_tocart",        # В корзину, всего
    "ordered_units",      # Заказано товаров
    "cancellations",      # Отменено товаров
    "delivered_units",    # Доставлено товаров
    "returns",            # Возвращено товаров
    "revenue",            # Заказано на сумму
    "hits_tocart_pdp",    # В корзину из карточки товара
    "hits_tocart_search", # В корзину из поиска или каталога
    "conv_tocart",        # Конверсия в корзину, общая (в % уже от API)
]


def _retry_after(resp) -> float | None:
    """Парсит заголовок Retry-After (секунды), если Ozon его прислал."""
    val = resp.headers.get("Retry-After")
    if not val:
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def _post(url: str, payload: dict) -> dict:
    """POST с ретраями и паузой для соблюдения лимита 2 rps.

    429 (rate limit) ретраится с отдельным большим бюджетом MAX_RETRIES_429,
    сетевые ошибки и 5xx — с бюджетом MAX_RETRIES.
    """
    net_attempt = 0     # сетевые ошибки + 5xx
    rate_attempt = 0    # 429
    while True:
        try:
            resp = requests.post(url, headers=HEADERS, data=json.dumps(payload), timeout=90)
        except requests.RequestException as e:
            net_attempt += 1
            if net_attempt >= MAX_RETRIES:
                raise
            wait = min(60, 5 * net_attempt)
            log.warning("Сетевая ошибка (%s); попытка %d/%d, пауза %ds",
                        e, net_attempt, MAX_RETRIES, wait)
            time.sleep(wait)
            continue

        if resp.status_code == 429:
            rate_attempt += 1
            if rate_attempt >= MAX_RETRIES_429:
                log.error("429 не отпускает после %d попыток — сдаёмся", rate_attempt)
                resp.raise_for_status()
            # Экспоненциальный бэкофф с потолком 60s, либо Retry-After от сервера.
            wait = _retry_after(resp) or min(60, 1.5 * 2 ** min(rate_attempt - 1, 5))
            if rate_attempt == 1 or rate_attempt % 10 == 0:
                log.warning("HTTP 429 (rate limit); попытка %d/%d, пауза %.1fs",
                            rate_attempt, MAX_RETRIES_429, wait)
            time.sleep(wait)
            continue

        if resp.status_code in SERVER_ERROR_STATUS:
            net_attempt += 1
            if net_attempt >= MAX_RETRIES:
                resp.raise_for_status()
            wait = min(30, 3 * net_attempt)
            log.warning("HTTP %d (%s); попытка %d/%d, пауза %ds",
                        resp.status_code, resp.text[:120], net_attempt, MAX_RETRIES, wait)
            time.sleep(wait)
            continue

        resp.raise_for_status()
        time.sleep(REQUEST_PAUSE)
        return resp.json()


def fetch_funnel_rows(target_date: date) -> list:
    """Пагинированный забор воронки за день. dimension = [sku, modelID].

    Возвращает список dict с ключами: sku_id, sku_name, model_name, metrics(list).
    Дедуп по sku_id (на границах страниц возможны повторы из-за сортировки).
    """
    day = target_date.strftime("%Y-%m-%d")
    rows = {}
    offset = 0
    for page in range(MAX_PAGES):
        payload = {
            "date_from": day,
            "date_to": day,
            "metrics": METRICS,
            "dimension": ["sku", "modelID"],
            "filters": [],
            "sort": [{"key": "ordered_units", "order": "DESC"}],
            "limit": PAGE_LIMIT,
            "offset": offset,
        }
        data = (_post(ANALYTICS_URL, payload).get("result") or {}).get("data") or []
        if not data:
            break

        for item in data:
            dims = item.get("dimensions") or []
            if not dims:
                continue
            sku = dims[0] or {}
            model = dims[1] if len(dims) > 1 else {}
            sku_id = str(sku.get("id", "")).strip()
            if not sku_id or sku_id in rows:
                continue
            rows[sku_id] = {
                "sku_id": sku_id,
                "sku_name": sku.get("name") or "",
                "model_name": (model or {}).get("name") or "",
                "metrics": item.get("metrics") or [],
            }

        log.info("analytics/data: страница offset=%d, получено=%d, всего уникальных=%d",
                 offset, len(data), len(rows))
        if len(data) < PAGE_LIMIT:
            break
        offset += PAGE_LIMIT
    else:
        log.warning("Достигнут предохранитель MAX_PAGES=%d — возможно, выгружены не все строки", MAX_PAGES)

    return list(rows.values())


def fetch_offer_ids(sku_ids: list) -> dict:
    """sku → offer_id через product/info/list. Возвращает {sku_str: offer_id}."""
    mapping = {}
    sku_ints = []
    for s in sku_ids:
        try:
            sku_ints.append(int(s))
        except (TypeError, ValueError):
            continue

    for i in range(0, len(sku_ints), INFO_BATCH):
        chunk = sku_ints[i:i + INFO_BATCH]
        data = _post(PRODUCT_INFO_LIST_URL, {"sku": chunk})
        items = data.get("items") or (data.get("result") or {}).get("items") or []
        for it in items:
            offer_id = it.get("offer_id") or ""
            # Привязываем offer_id ко всем sku товара: верхнеуровневый + из sources
            top = it.get("sku")
            if top:
                mapping[str(top)] = offer_id
            for src in it.get("sources") or []:
                src_sku = src.get("sku")
                if src_sku:
                    mapping[str(src_sku)] = offer_id
        log.info("product/info/list: батч %d..%d, сопоставлено offer_id=%d",
                 i, i + len(chunk), len(mapping))

    return mapping


def _ru_decimal(value, suffix: str = "") -> str:
    """Число → текст с запятой и 2 знаками (как в браузерной выгрузке).

    Примеры: 24.86 → '24,86'; 0.17 + '%' → '0,17%'. None/нечисло → ''.
    """
    if not isinstance(value, (int, float)):
        return ""
    return f"{value:.2f}{suffix}".replace(".", ",")


def build_dataframe(rows: list, offer_map: dict, target_date: date) -> pd.DataFrame:
    """Собирает DataFrame со столбцами OUTPUT_COLUMNS в точном порядке (18 колонок)."""
    records = []
    n_metrics = len(METRICS)
    day_str = target_date.strftime("%Y-%m-%d")
    for r in rows:
        m = list(r["metrics"]) + [None] * (n_metrics - len(r["metrics"]))
        # Позиция (индекс 3) и Конверсия (индекс 12) — текст с запятой, как в оригинале.
        records.append({
            "Товары": r["sku_name"],
            "Модель": r["model_name"],
            "Ozon ID": r["sku_id"],
            "Артикул": offer_map.get(r["sku_id"], ""),
            "День": day_str,
            "Показы, всего": m[0],
            "Показы на карточке товара": m[1],
            "В корзину, всего": m[4],
            "Заказано товаров": m[5],
            "Показы в поиске и каталоге": m[2],
            "Позиция в поиске и каталоге": _ru_decimal(m[3]),
            "В корзину из карточки товара": m[10],
            "В корзину из поиска или каталога": m[11],
            "Отменено товаров": m[6],
            "Доставлено товаров": m[7],
            "Возвращено товаров": m[8],
            "Заказано на сумму": m[9],
            "Конверсия в корзину, общая": _ru_decimal(m[12], "%"),
        })

    df = pd.DataFrame(records, columns=OUTPUT_COLUMNS)
    return df


def download_ozon_funnel_report(target_date: date = None, output_dir: Path = None) -> bool:
    log.info("Starting Ozon Funnel API Downloader...")

    if not OZON_CLIENT_ID or not OZON_API_KEY:
        log.error("OZON_CLIENT_ID / OZON_API_KEY не заданы")
        print("[ERROR] OZON_CLIENT_ID / OZON_API_KEY не заданы")
        return False

    if target_date is None:
        target_date = date.today() - timedelta(days=1)

    DOWNLOADS_DIR = Path(output_dir) if output_dir else DEFAULT_DOWNLOADS_DIR
    if output_dir:
        DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
    else:
        # Fail-fast, если к сетевой шаре нет подключения (как в браузерном скрипте).
        try:
            DEFAULT_DOWNLOADS_DIR.stat()
        except Exception as e:
            print(f"[ERROR] Cannot access downloads folder: {downloads_folder}. {e}")
            return False

    log.info("Target date: %s", target_date)

    try:
        # 1) Воронка: метрики + Товары/Ozon ID/Модель
        rows = fetch_funnel_rows(target_date)
        log.info("Получено строк воронки: %d", len(rows))
        if not rows:
            log.error("API вернул 0 строк за %s — файл не создаётся", target_date)
            print(f"[ERROR] Нет данных воронки за {target_date}")
            return False

        # 2) Артикул (offer_id) по sku
        offer_map = fetch_offer_ids([r["sku_id"] for r in rows])
        log.info("Сопоставлено offer_id для %d sku", len(offer_map))

        # 3) Сборка таблицы и запись xlsx
        df = build_dataframe(rows, offer_map, target_date)

        ts = datetime.now()
        filename = f"analytics_report_{ts.strftime('%Y-%m-%d_%H_%M')}.xlsx"
        dest = DOWNLOADS_DIR / filename
        df.to_excel(dest, index=False)
        log.info("SUCCESS: сохранено %d строк в %s", len(df), dest)
        print(f"SUCCESS: saved {len(df)} rows to {dest}")
        return True

    except Exception as e:
        log.exception("Ошибка выгрузки воронки через API: %s", e)
        print(f"[ERROR] {e}")
        return False


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Download Ozon Sales Funnel report via Seller API")
    parser.add_argument("--date", type=str, default=None,
                        help="Target date YYYY-MM-DD (default: yesterday)")
    parser.add_argument("--output-dir", type=str, default=None,
                        help="Directory to save the xlsx file (default: network share Taldykin)")
    args = parser.parse_args()

    target = None
    if args.date:
        target = datetime.strptime(args.date, "%Y-%m-%d").date()

    success = download_ozon_funnel_report(target_date=target, output_dir=args.output_dir)
    sys.stdout.flush(); sys.stderr.flush()  # os._exit не сбрасывает буферы print()
    os._exit(0 if success else 1)
