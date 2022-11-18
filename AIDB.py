# d_ папка и файлы детализации
# s_ словари
import numpy as np
import pandas as pd
import os
from tkinter import filedialog
import traceback
import regex as re
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
def och(d_df):                                            # очистка го и ук
    d_df = pd.read_pickle("./d_df.pkl ")
    d_df.columns = d_df.iloc[0]                       # в заголовок первую теперь уже строку
    d_df = d_df.iloc[1:]                              # удалить строку переехавшую в заголовок
    d_df['ОМСУ'] = d_df['ОМСУ'].str.replace("г. о. ", "")# удаление го для приведения к единому виду с базой
    d_df['Исполнитель'] = d_df['Исполнитель'].str.replace(" *\d", "", regex=True)  # удаление двойников с цифрами типа МО 55555 с пробелом
    d_df = d_df.query('ОМСУ == "Химки" or ОМСУ == "Лобня" or ОМСУ == "Кашира" or ОМСУ == "Солнечногорск" or ОМСУ == "Чехов"')  # зачистка го(и артефактов свода/лишние заголовки)
    s_ddf = pd.read_excel("./Словарь_управляшек.xlsx ", index_col=0)
    s_uk = list(s_ddf['Наименование УК'])
    d_df = d_df[d_df.Исполнитель.isin(s_uk)]
    s_uk1 = list(s_ddf['Статус'])
    d_df = d_df[d_df.Статус.isin(s_uk1)]
    #d_df = d_df[~d_df.Описание.str.startswith('Номер Добродела')]
    return d_df
if os.path.isfile("./chk.xlsx"):                                                                      # если уже очищено
    d_df = pd.read_excel("./chk.xlsx")
    print(1)
else:
    if not os.path.isfile("./d_df.pkl"):                                                # если ничего вообще не делалось
        d_df = toxls12()
        print(2)
    else:
        d_df = pd.read_pickle("./d_df.pkl ")                                          # если уже объединено но не чищено
        print(3)
    d_df = och(d_df)                                                                                                    # очистка
#kdf=d_df
d_df["корд"] = d_df['Описание'].str.extract(r"(5[456][.,]\d{3,8}\D{1,10}3[678][.,]\d{3,8})+")                           # формирование столбцов с координатами
d_df["Описание"] = d_df['Описание'].str.replace(r"5[456][.,]\d{2,8}\D{1,10}3[678][.,]\d{2,8}", "", regex=True)
d_df['Описание'] = d_df['Описание'].str.replace(r"\d{5,}", "", regex=True)                                              # стирает более чем 3знач числа
d_df['Описание'] = d_df['Описание'].str.replace(r"([0123][0123456789]\D{1,2}[01][0123456789]\D{1,2}20{,1}2{,1}2)+", "", regex=True)  # чистим даты из текста
d_df = d_df.loc[(d_df['Описание'].str.len() > 40)|(d_df['корд'].str.len() > 2)]                                         # строка длинее 40
d_df['Описание'] = d_df['Описание'].str.replace(r"Вопрос 1", "zzz", regex=True)                                         # заменяем цифру от фраззы чтоб не мешала удалять по признаку цифры
d_df = d_df[~d_df['Описание'].str.contains(r"zzz\D{,2}1{,1}.{,1}$", regex=True)]                                        # Вопрос 1 в конце как признак битости строки
d_df['Описание'] = d_df['Описание'].str.replace(r"202\d", "", regex=True)#7
d_df['Описание'] = d_df['Описание'].str.replace(r"\d{,9}[,.]{,1}\d{1,9}\D{,2}рубл[яе]", "", regex=True)
d_df['Описание'] = d_df['Описание'].str.replace(r"\d{,9}[,.]{,1}\d{1,9}\D{,2}копе[йие]", "", regex=True)#1!!!
d_df['Описание'] = d_df['Описание'].str.replace(r"\d{,2}-{,1}й{,1} {,1}под[ьъ]езда{,1} {,1}\d{,2}", "", regex=True) #10
d_df['Описание'] = d_df['Описание'].str.replace(r"\d{1,3} ?%", "", regex=True)
d_df['Описание'] = d_df['Описание'].str.replace(r"\d{1,3} ?суто?ки?", "", regex=True)
d_df['Описание'] = d_df['Описание'].str.replace(r"\d{1,3} ?этаже?", "", regex=True)

d_df['Описание'] = d_df['Описание'].str.replace(r"\d[еягй]?-?[еягй]?о?", "\d", regex=True)

d_df = d_df[d_df['Описание'].str.contains(r'\d{1,}', regex=True)|(d_df['корд'].str.len() > 2)]                          # есть хотябы одна цифра


print(d_df['Описание'])
d_df.to_excel('./chk1.xlsx')
#df.isin({'Исполнитель': [0, 3]})


