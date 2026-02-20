import os
import pandas as pd
from sqlalchemy import create_engine
import pyodbc
import win32com.client as win32
import time  # Для измерения времени выполнения
import shutil
import re


# FOLDER_PATH = os.path.normpath(r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!")
# df = pd.read_csv(os.path.join(FOLDER_PATH, "ДБсПризнаками.csv"))
# # print(df.head(20))
# # print(df.columns)

# print(df[df['Дата']=="30."])

# Функция для подключения к SQL Server с аутентификацией Windows
def connect_to_sql(server, database):
    connection_string = (
        f"mssql+pyodbc://{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server"
        "&trusted_connection=yes"
    )
    engine = create_engine(connection_string)
    return engine

SQL_SERVER = "cl01sql"
SQL_DATABASE_DBREPORT = "DBReport"
SQL_DATABASE_DBPARTNERS = "DBPartners"

engine = connect_to_sql(SQL_SERVER, SQL_DATABASE_DBPARTNERS)

 # 10. Получить данные таблицы с SQL (Цены)
# === 3. SQL ЦЕНЫ ===
sql = f"""
select 'OZ' as AGREGATOR, DT, ITEMID, PRICE
from [DBPartners].[dbo].[WblmRepPriceDiscountOzReport]
where dt >= '{"2026-01-01"}'
"""
df_prices = pd.read_sql(sql, engine)

df_prices = df_prices.groupby(["DT", "ITEMID"], as_index=False).agg({"PRICE": "max"})
df_prices = df_prices.rename(columns={"DT": "Дата", "ITEMID": "Артикул", "PRICE": "Цена"})

print(df_prices)