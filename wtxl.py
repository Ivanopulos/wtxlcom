import pandas  # +openpyxl
import zipfile
import os
df = pandas.read_excel("1.xlsx")
tm = df.columns.get_loc("Имя_документа")
pathword = df.iloc[0][tm]
pathzip = pathword[:-4] + "zip"

mem = 0
memcol = 0
isch = 0
usch = 0
found = ""

try:
    os.remove(pathzip)
except:
    asd = 1