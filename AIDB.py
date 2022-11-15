# d_ папка и файлы детализации
# s_ словари
import numpy as np
import pandas as pd
import os
from tkinter import filedialog
import traceback
import re
import warnings
warnings.simplefilter("ignore")  # dftmp = pd.read_excel(d_path+filename)
def p():
    #print(traceback.extract_stack())  # Полная информация о стеке вызовов
    ns=traceback.extract_stack()[-2].lineno  # Только номер строки, откуда была вызвана функция
    s=0
    f = open('AIDB.py', 'r')
    line = f.readline()
    while line:
        #print (line),
        line = f.readline()
        s = s+1
        if s == ns-2:
            stroka = line
    f.close()
    stroka=re.search(r'\S*\s', stroka)[0]
    stroka=stroka[:-1]
    exec("print("+stroka+")")
    #sss = "sss1"
    #locals()[sss] = 5000  # globals()[sss] = 5000
    #print(sss1)
    #print(eval("2+2"))
def toxls12():  # соединение в один дф
    d_df = pd.DataFrame()                         # дф для сборки файлов
    d_path = filedialog.askopenfilename()         # произвольный файл из папки с базой(только целевые должны быть), а база едс на шаг выше
    d_path = re.search(r".*\/", d_path).group(0)  # из пути файла - в путь папки
    for root, dirs, files in os.walk(d_path):     # поиск всех файлов и сбор в один
        for filename in files:
            print(d_path+filename)
            dftmp = pd.read_excel(d_path+filename)
            print(dftmp)
            d_df = pd.concat([d_df, dftmp], axis=0)
    d_df.to_excel('./d_df.xlsx')                  # save
    print('xls out')
    d_df.to_pickle("./d_df.pkl ")
    print('pkl out')
    return d_df
if os.path.isfile("./d_df.pkl"):
    d_df = pd.read_pickle("./d_df.pkl ")
else:
    d_df = toxls12()
# очистка от лишнего
d_df = pd.read_pickle("./d_df.pkl ")
d_df.columns = d_df.iloc[0]                       # в заголовок первую теперь уже строку
d_df = d_df.iloc[1:]                              # удалить строку переехавшую в заголовок
d_df['ОМСУ'] = d_df['ОМСУ'].str.replace("г. о. ", "")
d_df['Исполнитель'] = d_df['Исполнитель'].str.replace(" *\d", "", regex=True)  # удаление двойников с цифрами
d_df = d_df.query('ОМСУ == "Химки" or ОМСУ == "Лобня" or ОМСУ == "Кашира" or ОМСУ == "Солнечногорск" or ОМСУ == "Чехов"')  # зачистка го(и артефактов свода/лишние заголовки)
s_ddf = pd.read_excel("./Словарь_управляшек.xlsx ", index_col=0)
s_uk = list(s_ddf['Наименование УК'])
print(s_ddf)
d_df = d_df[d_df.Исполнитель.isin(s_uk)]
s_uk1 = list(s_ddf['Статус'])
d_df = d_df[d_df.Статус.isin(s_uk1)]
#d_df = d_df.isin({'Исполнитель': s_uk})
#df.isin({'Исполнитель': [0, 3]})
print(d_df)
# d_df['оставляем'] = 0
# d_df['оставляем'] = np.where(d_df['ОМСУ'] == "Химки", d_df['оставляем']+1, d_df['оставляем'])
#print(d_df)
d_df.to_excel('./chk.xlsx')

