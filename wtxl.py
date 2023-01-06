import pandas  # +openpyxl
import zipfile
import os
df = pandas.read_excel("1.xlsx")
tm = df.columns.get_loc("Имя_документа")
pathword = df.iloc[0][tm] #адрес ворда
tm = df.columns.get_loc("Значение")
ndog = df.iloc[0][tm]#номер договора
tm = df.columns.get_loc("Экземпляров")
ekz = df.iloc[0][tm]#количество экземпларов
pathzip = pathword[:-4] + "zip"
try:
    os.remove(pathzip)
except:
    asd = 1
os.rename(pathword, pathzip)
for y in range(ekz):
    isch = 0
    usch = 0
    fantasy_zip = zipfile.ZipFile(pathzip)  # extract zip (+need rename docx to zip +need raname vise versa
    fantasy_zip.extractall("/B")
    fantasy_zip.close()
    with open("/B/word/document.xml", 'r', encoding='utf-8') as f:  # save before chenge
        get_all = f.readlines()
    print("file opened")

    ga2 = get_all
    with open("/B/word/document.xml", 'w', encoding='utf-8') as f:  # look for { and chenge it
        for i in ga2:         # STARTS THE NUMBERING FROM 1 (by default it begins with 0)
            usch = len(i)-1
            for u in i:
                if ga2[isch][usch:usch+3] == "{1}":
                    ga2[isch] = ga2[isch][:usch] + str(ndog) + ga2[isch][usch + 3:]
                usch = usch - 1
            isch = isch + 1
        f.writelines(ga2)
    print(ndog)
    try:
        os.remove("B.zip")
    except:
        asd = 1
    fantasy_zip = zipfile.ZipFile("B.zip", 'w')
    for folder, subfolders, files in os.walk("/B"):
        for file in files:
            fantasy_zip.write(os.path.join(folder, file), os.path.relpath(os.path.join(folder, file), "/B"))
    fantasy_zip.close()  # transform it to zip
    try:
        os.remove(str(ndog) + ".docx")
    except:
        asd = 1
    os.rename("B.zip", str(ndog) + ".docx")
    ndog = ndog+1
os.rename("A.zip", "A.docx")
