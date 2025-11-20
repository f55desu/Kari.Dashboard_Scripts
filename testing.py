import os
import pandas as pd
from sqlalchemy import create_engine
import pyodbc
import win32com.client as win32
import time  # Для измерения времени выполнения
import shutil
import re


FOLDER_PATH = os.path.normpath(r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!")
df = pd.read_csv(os.path.join(FOLDER_PATH, "ДБсПризнаками.csv"))
# print(df.head(20))
# print(df.columns)

print(df[df['Дата']=="30."])