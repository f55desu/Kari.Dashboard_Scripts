# Kari.Dashboard_Scripts

> **[English](#english)** | **[Русский](#русский)**

---

<a id="english"></a>

## English

Scripts for building and updating advertising campaign dashboards for marketplaces (Wildberries and Ozon).

### Project Structure

```
Kari.Dashboard_Scripts/
├── Preprocessing/                          # Data loading and preparation
│   ├── funnel_sql_exporter.py              # Funnel/costs export: MSSQL/Excel -> PostgreSQL
│   ├── ozon_funnel_api_downloader.py       # Ozon funnel via Seller API
│   ├── ozon_costs_api_downloader.py        # Ozon costs via Performance API
│   ├── ozon_analytics_downloader.py        # Ozon analytics
│   ├── ozon_ad_report_downloader.py        # Ozon ad reports
│   ├── ozon_union_db.py                    # PostgreSQL connection (config, engine)
│   ├── ozon_union_backfill.py              # Backfill work.ozon_promo_union
│   ├── ozon_dashboard_db.py                # PG utilities for Ozon
│   ├── wb_funnel_api_downloader.py         # WB funnel via nm-report API
│   ├── wb_costs_api_downloader.py          # WB costs via /adv/v1/upd API
│   ├── wb_downloader.py / wb_downloader_v2.py  # WB downloaders
│   ├── wb_campaign_downloader.py           # WB campaigns
│   ├── Dashboard_WB_Wrapper*.py            # WB launch wrappers (daily/Monday/Tuesday)
│   ├── Dashboard_OZON_Wrapper*.py          # Ozon launch wrappers (daily/Monday)
│   ├── Dashboard_WB_GUI.py                 # GUI for WB Wrapper
│   ├── Dashboard_OZON_GUI.py               # GUI for Ozon Wrapper
│   └── Concat_Utility.py                   # File concatenation utility
│
├── DashboardAssemblyOZ.py                  # Ozon dashboard assembly (original, from Excel)
├── DashboardAssemblyOZ_SQLBased.py         # Ozon dashboard assembly (from PostgreSQL)
├── DashboardAssemblyOZ_AutoPipeline.py     # Auto-pipeline: tunnel + Ozon assembly
├── DashboardAssemblyOZ_GUI.py              # GUI for Ozon assembly
├── DashboardAssemblyWB.py                  # WB dashboard assembly
├── DashboardAssemblyWB_AutoPipeline.py     # Auto-pipeline: tunnel + WB assembly
├── DashboardAssemblyWB_GUI.py              # GUI for WB assembly
├── DashboardDBUploaderOZ.py                # Upload assembled Ozon dashboard to DB
├── DashboardDBUploaderWB.py                # Upload assembled WB dashboard to DB
├── Funnel_Costs_PreprocessingOZ_Autopipeline.py  # Ozon funnel/costs preprocessing
├── Funnel_Costs_PreprocessingWB_Autopipeline.py  # WB funnel/costs preprocessing
├── SQL_Exporter_Wrapper.py                 # Wrapper for funnel_sql_exporter
│
├── analytics_tunnel.vbs                    # SSH tunnel to PostgreSQL
├── pgadmin_connect.bat                     # Connect to PG via pgAdmin
├── build_gui_apps.bat                      # Build GUI apps into .exe
└── README_BUILD.md                         # .exe build instructions
```

### Quick Start

#### 1. Prerequisites

- Python 3.9+
- Access to SQL Server `cl01sql` (databases `DBReport`, `DBPartners`)
- SSH tunnel to PostgreSQL (localhost:15432)
- Network access to `\\kari.local\public\all\Analytics\`

Install dependencies:
```bash
pip install pandas numpy sqlalchemy pyodbc openpyxl xlwings pywin32 pyarrow psycopg2-binary python-dotenv
```

#### 2. Setting Up the SSH Tunnel to PostgreSQL

Before running scripts that interact with PostgreSQL, open the SSH tunnel:

```bash
# Option 1: via VBS script
analytics_tunnel.vbs

# Option 2: via batch file
pgadmin_connect.bat
```

The tunnel forwards PostgreSQL to `localhost:15432`.

#### 3. Running Dashboard Assembly

**Via GUI (recommended):**
- `DashboardAssemblyOZ_GUI.py` -- Ozon
- `DashboardAssemblyWB_GUI.py` -- Wildberries

**Via auto-pipeline (automatically checks tunnel):**
```bash
python DashboardAssemblyOZ_AutoPipeline.py
python DashboardAssemblyWB_AutoPipeline.py
```

**Directly:**
```bash
python DashboardAssemblyOZ.py
python DashboardAssemblyWB.py
```

### Data Sources

| Source | Description | PG Table |
|--------|------------|----------|
| MSSQL `mp.ozon_sales_funnel` | Ozon sales funnel (14 columns, Ozon ID) | `work.ozon_sales_funnel` |
| Excel `analytics_report_*.xlsx` | Ozon funnel from dashboard (Ozon ID + Article) | -- |
| Excel `Аналитика продвижения_*.xlsx` | Ozon costs (Statistics sheet) | `work.ozon_costs_statistics` |
| PG `work.ozon_promo_union` | Combined Ozon promo data | -- |
| MSSQL `DBPartners.dbo.WblmRep*` | Prices, stock, distribution | -- |
| Excel `Справочник.xlsx` | Article reference (brand, group, season, etc.) | -- |

### Data Preprocessing

#### funnel_sql_exporter.py

Central data exporter. Supported modes:

```bash
# Export Ozon funnel from MSSQL to PG
python funnel_sql_exporter.py --oz-funnel

# Export Ozon costs from Excel to PG
python funnel_sql_exporter.py --oz-costs

# Export WB funnel
python funnel_sql_exporter.py --wb-funnel
```

#### API Downloaders

API downloaders require a `.env` file in the `Preprocessing/` folder:

```env
OZON_CLIENT_ID=...
OZON_API_KEY=...
OZON_PERFORMANCE_CLIENT_ID=...
OZON_PERFORMANCE_CLIENT_SECRET=...
WB_API_TOKEN=...
WB_ADVERT_TOKEN=...
```

### Building GUI Applications (.exe)

Detailed instructions in [README_BUILD.md](README_BUILD.md).

Quick build of all GUI apps:
```bash
build_gui_apps.bat
```

Manual build of a single app:
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=app_OZ.ico --name="Dashboard_OZ_Assembly" DashboardAssemblyOZ_GUI.py
```

Built .exe files will appear in `apps/` (or `dist/` for manual builds).

### Ozon Assembly Versions

| Script | Funnel Source | Costs Source | Status |
|--------|--------------|-------------|--------|
| `DashboardAssemblyOZ.py` | Excel files | CSV/Excel | Production |
| `DashboardAssemblyOZ_SQLBased.py` | PostgreSQL | PostgreSQL | In development |
| `DashboardAssemblyOZ_SQLBased_Experimental.py` | PostgreSQL (early costs join) | PostgreSQL | Experimental |

### Troubleshooting

#### SSH tunnel won't connect
Make sure VPN is active and the analytics server is reachable. Restart `analytics_tunnel.vbs`.

#### "column ... does not exist" in PostgreSQL
When new columns appear in Excel files, `funnel_sql_exporter.py` automatically adds missing columns via `ALTER TABLE ... ADD COLUMN IF NOT EXISTS`.

#### Number mismatch between original and SQLBased
The SQL version groups by Ozon ID (~209K/day), while the original groups by article prefix (~54K/day). The experimental version solves this with an early costs join before groupby.

---

<a id="русский"></a>

## Русский

Скрипты для сборки и обновления дашбордов по рекламным кампаниям маркетплейсов (Wildberries и Ozon).

### Структура проекта

```
Kari.Dashboard_Scripts/
├── Preprocessing/                          # Загрузка и подготовка данных
│   ├── funnel_sql_exporter.py              # Экспорт воронки/затрат: MSSQL/Excel -> PostgreSQL
│   ├── ozon_funnel_api_downloader.py       # Воронка Ozon через Seller API
│   ├── ozon_costs_api_downloader.py        # Затраты Ozon через Performance API
│   ├── ozon_analytics_downloader.py        # Аналитика Ozon
│   ├── ozon_ad_report_downloader.py        # Рекламные отчёты Ozon
│   ├── ozon_union_db.py                    # Подключение к PostgreSQL (конфиг, engine)
│   ├── ozon_union_backfill.py              # Бэкфилл work.ozon_promo_union
│   ├── ozon_dashboard_db.py                # Утилиты PG для Ozon
│   ├── wb_funnel_api_downloader.py         # Воронка WB через API nm-report
│   ├── wb_costs_api_downloader.py          # Затраты WB через API /adv/v1/upd
│   ├── wb_downloader.py / wb_downloader_v2.py  # Загрузчики WB
│   ├── wb_campaign_downloader.py           # Кампании WB
│   ├── Dashboard_WB_Wrapper*.py            # Обёртки запуска WB (ежедневно/понедельник/вторник)
│   ├── Dashboard_OZON_Wrapper*.py          # Обёртки запуска Ozon (ежедневно/понедельник)
│   ├── Dashboard_WB_GUI.py                 # GUI для WB Wrapper
│   ├── Dashboard_OZON_GUI.py               # GUI для Ozon Wrapper
│   └── Concat_Utility.py                   # Утилита объединения файлов
│
├── DashboardAssemblyOZ.py                  # Сборка дашборда Ozon (оригинал, из Excel)
├── DashboardAssemblyOZ_SQLBased.py         # Сборка дашборда Ozon (из PostgreSQL)
├── DashboardAssemblyOZ_AutoPipeline.py     # Автопайплайн: tunnel + сборка Ozon
├── DashboardAssemblyOZ_GUI.py              # GUI запуска сборки Ozon
├── DashboardAssemblyWB.py                  # Сборка дашборда WB
├── DashboardAssemblyWB_AutoPipeline.py     # Автопайплайн: tunnel + сборка WB
├── DashboardAssemblyWB_GUI.py              # GUI запуска сборки WB
├── DashboardDBUploaderOZ.py                # Загрузка собранного дашборда Ozon в БД
├── DashboardDBUploaderWB.py                # Загрузка собранного дашборда WB в БД
├── Funnel_Costs_PreprocessingOZ_Autopipeline.py  # Предобработка воронки/затрат Ozon
├── Funnel_Costs_PreprocessingWB_Autopipeline.py  # Предобработка воронки/затрат WB
├── SQL_Exporter_Wrapper.py                 # Обёртка для funnel_sql_exporter
│
├── analytics_tunnel.vbs                    # Запуск SSH-тоннеля к PostgreSQL
├── pgadmin_connect.bat                     # Подключение к PG через pgAdmin
├── build_gui_apps.bat                      # Сборка GUI-приложений в .exe
└── README_BUILD.md                         # Инструкция по сборке .exe
```

### Быстрый старт

#### 1. Предварительные требования

- Python 3.9+
- Доступ к SQL Server `cl01sql` (база `DBReport`, `DBPartners`)
- SSH-тоннель к PostgreSQL (localhost:15432)
- Сетевой доступ к `\\kari.local\public\all\Analytics\`

Установка зависимостей:
```bash
pip install pandas numpy sqlalchemy pyodbc openpyxl xlwings pywin32 pyarrow psycopg2-binary python-dotenv
```

#### 2. Настройка SSH-тоннеля к PostgreSQL

Перед запуском скриптов, работающих с PostgreSQL, необходимо открыть SSH-тоннель:

```bash
# Вариант 1: через скрипт
analytics_tunnel.vbs

# Вариант 2: через bat-файл
pgadmin_connect.bat
```

Тоннель пробрасывает PostgreSQL на `localhost:15432`.

#### 3. Запуск сборки дашборда

**Через GUI (рекомендуется):**
- `DashboardAssemblyOZ_GUI.py` -- Ozon
- `DashboardAssemblyWB_GUI.py` -- Wildberries

**Через автопайплайн (автоматически проверяет тоннель):**
```bash
python DashboardAssemblyOZ_AutoPipeline.py
python DashboardAssemblyWB_AutoPipeline.py
```

**Напрямую:**
```bash
python DashboardAssemblyOZ.py
python DashboardAssemblyWB.py
```

### Источники данных

| Источник | Описание | Таблица PG |
|----------|----------|------------|
| MSSQL `mp.ozon_sales_funnel` | Воронка продаж Ozon (14 колонок, Ozon ID) | `work.ozon_sales_funnel` |
| Excel `analytics_report_*.xlsx` | Воронка Ozon из ЛК (Ozon ID + Артикул) | -- |
| Excel `Аналитика продвижения_*.xlsx` | Затраты Ozon (лист Statistics) | `work.ozon_costs_statistics` |
| PG `work.ozon_promo_union` | Объединённые промо-данные Ozon | -- |
| MSSQL `DBPartners.dbo.WblmRep*` | Цены, остатки, дистрибуция | -- |
| Excel `Справочник.xlsx` | Справочник артикулов (бренд, группа, сезон и др.) | -- |

### Предобработка данных (Preprocessing)

#### funnel_sql_exporter.py

Центральный экспортёр данных. Поддерживает режимы:

```bash
# Экспорт воронки Ozon из MSSQL в PG
python funnel_sql_exporter.py --oz-funnel

# Экспорт затрат Ozon из Excel в PG
python funnel_sql_exporter.py --oz-costs

# Экспорт воронки WB
python funnel_sql_exporter.py --wb-funnel
```

#### API-загрузчики

Для работы API-загрузчиков необходим файл `.env` в папке `Preprocessing/`:

```env
OZON_CLIENT_ID=...
OZON_API_KEY=...
OZON_PERFORMANCE_CLIENT_ID=...
OZON_PERFORMANCE_CLIENT_SECRET=...
WB_API_TOKEN=...
WB_ADVERT_TOKEN=...
```

### Сборка GUI-приложений (.exe)

Подробная инструкция в [README_BUILD.md](README_BUILD.md).

Быстрая сборка всех GUI-приложений:
```bash
build_gui_apps.bat
```

Ручная сборка одного приложения:
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=app_OZ.ico --name="Dashboard_OZ_Assembly" DashboardAssemblyOZ_GUI.py
```

Готовые .exe появятся в папке `apps/` (или `dist/` при ручной сборке).

### Версии сборки Ozon

| Скрипт | Источник воронки | Источник затрат | Статус |
|--------|-----------------|-----------------|--------|
| `DashboardAssemblyOZ.py` | Excel-файлы | CSV/Excel | Продакшен |
| `DashboardAssemblyOZ_SQLBased.py` | PostgreSQL | PostgreSQL | В разработке |
| `DashboardAssemblyOZ_SQLBased_Experimental.py` | PostgreSQL (с ранним джойном затрат) | PostgreSQL | Экспериментальный |

### Устранение проблем

#### SSH-тоннель не подключается
Убедитесь, что VPN активен и есть доступ к серверу аналитики. Перезапустите `analytics_tunnel.vbs`.

#### "column ... does not exist" в PostgreSQL
При добавлении новых колонок в Excel-файлы, `funnel_sql_exporter.py` автоматически добавляет недостающие колонки через `ALTER TABLE ... ADD COLUMN IF NOT EXISTS`.

#### Расхождение чисел между оригиналом и SQLBased
SQL-версия группирует по Ozon ID (~209K/день), а оригинал по Артикул-prefix (~54K/день). Экспериментальная версия решает это через ранний джойн затрат до groupby.
