# d_ папка и файлы детализации d_df d_path
# s_ словари s_uk/s_uk1 s_ddf/s_adf s_adf1
# s_ddf словарь управляшек <01/мм*3/гггг> need to renew
# c_ результаты сведения
# путь по файлам:
#   запрос в одой папке входящих, d_df.pkl/xlsx, chk.xlsx, chk1.xlsx
import time
import numpy as np
import pandas as pd #pip install xlsxwriter
import os
from tkinter import filedialog
import regex as re
import warnings
#warnings.simplefilter("ignore")  # dftmp = pd.read_excel(d_path+filename)
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
def chk1(d_df):
    d_df["корд"] = d_df['Описание'].str.extract(r"(5[456][\.,]\d{4,8}\D{1,10}3[678][\.,]\d{4,8})+")  # формирование столбцов с координатами
    d_df["корд"] = d_df["корд"].str.replace(r"(\d\d)[\.,](\d{4,8})\D{1,10}(3[678])[\.,](\d{4,8})", r"[\1.\2,\3.\4]", regex=True)  # стандарт написания
    d_df["Описание"] = d_df['Описание'].str.replace(r"5[456][\.,]\d{2,8}\D{1,10}3[678][\.,]\d{2,8}", "", regex=True)  # удаление координат
    d_df['Описание'] = d_df['Описание'].str.replace(r'[\t\r\n]', "", regex=True)  # дублирует '[^0-9а-яА-Я \.,/ёЁ]' мож и не надо уже
    # d_df['Описание'] = d_df['Описание'].str.replace(r'\-', "/", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'[^0-9а-яА-Я \.,/ёЁ\\№\-]', "", regex=True)  # а-яА-Я ёЁйЙ.,\-
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d{2,}[/\.\-]\d{2,}[/\. \-]\d{2,}', "", regex=True)  # всякие номера типа ххх-хх-хх не захв 8 (498) 000-00-00
    d_df['Описание'] = d_df['Описание'].str.replace(r'[87]? ?\(\d{3,}\)', "", regex=True)  # остатки ном тел но походу ничего не дает
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d{4,}', "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ раз[ \.,]', "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d?[,\.\-]?\d+ месяц', "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d\d[\.\-:]00", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\dх", "", regex=True)
    d_df = d_df.loc[(d_df['Описание'].str.len() > 40) | (d_df['корд'].str.len() > 2)]  # строка длинее 40
    d_df['Описание'] = d_df['Описание'].str.replace(r"Вопрос 1", "zzz", regex=True)  # заменяем цифру от фраззы чтоб не мешала удалять по признаку цифры
    d_df = d_df[~d_df['Описание'].str.contains(r"zzz\D{,2}1{,1}\.{,1}$", regex=True)]  # Вопрос 1 в конце как признак битости строки
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d{,9}[,\.]{,1}\d{1,9}\D{,2}рубл[яе][й ,\.]", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d{,9}[,\.]{,1}\d{1,9}\D{,2}копе[йие]", "", regex=True)  # 1!!!
    d_df['Описание'] = d_df['Описание'].str.replace(r"(\d)[еягй]?\-?[еягй]?о?", r"\1", regex=True)  # 9-й в 9
    d_df['Описание'] = d_df['Описание'].str.replace(r" \d{,2}\-{,1}й{,1}г?о? {,1}[пП]од[ьъ]езда{,1}е{,1}о?м? {,1}\d{,2}", "", regex=True)  # 10
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d{1,3} ?%", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d{1,3} ?суто?ки?", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d{1,3} ?[эЭ]таже?", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"[012]?[0-9]:[012]?[0-9]", "", regex=True)  # time6968614-1
    d_df['Описание'] = d_df['Описание'].str.replace(r"[эЭ]таж ?\d{1,3}", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"[оО]ценка \d", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"[рР]ассмотрение \d", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"[вВ]опрос \d", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d[хм]", "", regex=True)
    d_df = d_df[d_df['Описание'].str.contains(r'\d{1,}', regex=True) | (d_df['корд'].str.len() > 2)]  # есть хотябы одна цифра
    d_df = d_df.drop_duplicates(subset=['Номер ЕЦУР'])  # 50 если описание то 100
    print(d_df['Описание'])
    d_df.to_excel('./chk1.xlsx')
    return d_df
def rasp(x):
    global s_adf
    global t
    s_adf1 = s_adf[(s_adf.ОМСУ == x["ОМСУ"])&(s_adf.УК == x["Исполнитель"])]
    s_adf1['есхожд'] = 0
    s_adf1['улица'] = s_adf1['улица'].str.replace(r'.+, ', "", regex=True)
    #print(s_adf1['ул'])
    u = s_adf1.columns.get_loc("улица")
    d = s_adf1.columns.get_loc('дом')
    e = s_adf1.columns.get_loc('есхожд')
    a = x["Описание"]
    for i in range(1, len(s_adf1)):
        b = s_adf1.iloc[i, u]
        c = s_adf1.iloc[i, d]
        if a.find(b) > 1 and a.find(c) > 1: # and (x["Описание"].find(str(s_adf1.iloc[i]["дом"]))>1
            s_adf1.iloc[i, e] = 1
            #print(s_adf1.iloc[i, e])
    #tmp=s_adf1['есхожд'].sum()
    #print(tmp)
    t=t+1
    if int(t/100)==t/100:
        print(t)
    if s_adf1['есхожд'].sum() == 1:
        itg = s_adf1[s_adf1["есхожд"] == 1]["Адрес в образце"]
        return itg
if not os.path.isfile("./d_df.pkl"):
    d_df = toxls12()  # сбор
    d_df.to_excel('./d_df.xlsx')
    print('xls out')
    d_df.to_pickle("./d_df.pkl ")
    print('pkl out')
    d_df = och(d_df)  # фильтр
    d_df.to_excel('./chk.xlsx', engine='xlsxwriter')
    print('chk out')
if not os.path.isfile("./chk1.xlsx"):
    d_df = pd.read_excel("./chk.xlsx")  # очистка от битых
    d_df = chk1(d_df)
    print('chk1 out')
s_adf = pd.read_excel("./адрес стандарт.xlsx")
s_adf["ОМСУ"] = s_adf["ОМСУ"].str.replace(r' г о', "", regex=True)
try:
    if len(d_df) < 10:
        print("low len")
except:
    d_df = pd.read_excel("./chk1.xlsx")
    print("chk1 eated")
t = 0
d_df["распознано"] = d_df.apply(rasp, axis=1)
d_df.to_excel("./chk2.xlsx")
print(d_df["распознано"])

#pd.options.display.max_colwidth = 1000
#print(str(d_df[d_df["Номер ЕЦУР"] == "7500157-2"]["Описание"]).encode("UTF-8"))7258977-1
#print(d_df[d_df["Номер ЕЦУР"] == "7258977-1"]["Описание"])
#print(b'\xb2\xd0\xb0 -5910223 \xd0\x9e\xd1'.decode("utf-8", "ignore"))
