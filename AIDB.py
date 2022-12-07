# d_ папка и файлы детализации d_df d_path
# s_ словари s_uk/s_uk1 s_ddf/s_adf s_adf1
# s_ddf словарь управляшек <01/мм*3/гггг> need to renew
# c_ результаты сведения
# путь по файлам:
#   запрос в одой папке входящих,
#   d_df.pkl/xlsx(итог сбора 10)
#   chk.xlsx(итог фильтров го и коллекции ук по ./Словарь_управляшек.xlsx )
#   chk1.xlsx(итог, определение, удаление мусорных, очистка цифр)
#   chk2.xlsx(распознание по единственному совпадению)
#   адрес_стандарт.xlsx(база возможной комбинации ул.дом.омсу.ук.написание в базе координат)
#   адреса.xlsx(коллекция координат)
import time
import numpy as np
import pandas as pd #pip install xlsxwriter
import os
from tkinter import filedialog
import regex as re
import warnings
warnings.simplefilter("ignore")  # dftmp = pd.read_excel(d_path+filename)
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
    s_ddf = pd.read_excel("./Словарь_управляшек.xlsx", index_col=0)
    s_uk = list(s_ddf['Наименование УК'])
    d_df = d_df[d_df.Исполнитель.isin(s_uk)]
    s_uk1 = list(s_ddf['Статус'])
    d_df = d_df[d_df.Статус.isin(s_uk1)]
    #d_df = d_df[~d_df.Описание.str.startswith('Номер Добродела')]
    return d_df
def chk1(d_df):  # очистка цифр
    d_df['Описание_старое'] = d_df['Описание']
    d_df["корд"] = d_df['Описание'].str.extract(r"(5[456][\.,]\d{4,8}\D{1,10}3[678][\.,]\d{4,8})+")  # формирование столбцов с координатами
    d_df["корд"] = d_df["корд"].str.replace(r"(\d\d)[\.,](\d{4,8})\D{1,10}(3[678])[\.,](\d{4,8})", r"[\1.\2,\3.\4]", regex=True)  # стандарт написания
    d_df["Описание"] = d_df['Описание'].str.replace(r"5[456][\.,]\d{2,8}\D{1,10}3[678][\.,]\d{2,8}", "", regex=True)  # удаление координат
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d\d\.\d\d\.\d{2,4}', "", regex=True) #даты правильного формата
    d_df['Описание'] = d_df['Описание'].str.replace(r'\S+@\S+', "", regex=True) #майлы
    d_df['Описание'] = d_df['Описание'].str.replace(r'^номер \S+', "", regex=True, flags=re.IGNORECASE)  # пос flag=re.MULTILINE
    d_df['Описание'] = d_df['Описание'].str.replace(r'[\t\r\n]', " ", regex=True)  # дублирует '[^0-9а-яА-Я \.,/ёЁ]' мож и не надо уже
    # d_df['Описание'] = d_df['Описание'].str.replace(r'\-', "/", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'!', " ", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'[^0-9а-яА-Я \.,/ёЁ\\№\-]', "", regex=True)  # а-яА-Я ёЁйЙ.,\-
    d_df['Описание'] = d_df['Описание'].str.replace(r'ё', "е", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\\', "/", regex=True)  # /\
    d_df['Описание'] = d_df['Описание'].str.replace(r'\b[тТ]ел[\. е](?!й)\D{,6}[\d\- \)]{,20}', "", regex=True) #тел счет
    d_df['Описание'] = d_df['Описание'].str.replace(r'\(\d{3,4}\)[\d\- \)]{,20}', "", regex=True)#тел счет
    d_df['Описание'] = d_df['Описание'].str.replace(r'(?<!дом)(?<!домом)(?<!д.)[\d\- \)]{10,20}', "", regex=True)#меньше 9 нельзя, жрет 112-110
    d_df['Описание'] = d_df['Описание'].str.replace(r'\bсч[ёе]т[№ ]\D{,6}[\d\- \)]{,20}', "", regex=True)#тел счет
    d_df['Описание'] = d_df['Описание'].str.replace(r'(\w)\-(\w)', r"\1\2", regex=True)  # тире разрывающее слово
    d_df['Описание'] = d_df['Описание'].str.replace(r'[87]?\-? ?\(\d{2,}\)?', "", regex=True)  # остатки ном тел но походу ничего не дает
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ ?года?', "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ век', "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ час', "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'с \d+ по \d+', "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+[^\.,]?фз ?№?\d*', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d?[,\.\-]?\d+ ?лет(?! октября)', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ января', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ февраля', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ апреля', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ июня', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ июля', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ августа', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ сентября', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ октября', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ ноября', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ декабря', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d?[012345679] марта', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d?[012345678] мая', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d{1,2}\.\d{2}', "", regex=True)                                  # 01.30
    d_df['Описание'] = d_df['Описание'].str.replace(r"\квартир[ае]? №?\d+", "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d+ квартир[ае]", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'кв\.? ?№?\d+', "", regex=True, flags=re.IGNORECASE)  # квартира
    d_df['Описание'] = d_df['Описание'].str.replace(r'[\b\d+ ]кв\.', "", regex=True, flags=re.IGNORECASE)  # квартира
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d{4,}', "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'[34567890]\d{2}', "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"под[ьъ]езда?е?о?м?\s?\d{1,2}(?! ?д)", "", regex=True, flags=re.IGNORECASE)  # 10
    d_df['Описание'] = d_df['Описание'].str.replace(r"(?<!\d)(?<!дом)(?<!дома)(?<!домом)(?<!№)(?<!д\.)[ /]?[1]?\d\-?й?г?о? ?под[ьъ]езда?е?о?м?", "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ домов', r"", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'(\d+)\-\d+', r"\1", regex=True, flags=re.IGNORECASE)              # может жрать корпуса формата а-б
    d_df['Описание'] = d_df['Описание'].str.replace(r"офис ?№?\d+", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"обращение №?\d+", "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d+ ?\d* ?тыс[ \.я][ч]?", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ раз[ \.,а]', "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d* стать[яи] \d*', "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ метр[ ао]в?', "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d+ м2[ \.,]', "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r'\d?[,\.\-]?\d+ месяц', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r'месяца \d(?! дом)(?! д\.)', "", regex=True, flags=re.IGNORECASE)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d?\d ?[\.\-:] ?00", "", regex=True)#23:00
    d_df['Описание'] = d_df['Описание'].str.replace(r"\dх", "", regex=True)
    d_df = d_df.loc[(d_df['Описание'].str.len() > 20) | (d_df['корд'].str.len() > 2)]  # строка длинее 40
    d_df['Описание'] = d_df['Описание'].str.replace(r"Вопрос 1", "zzz", regex=True)  # заменяем цифру от фраззы чтоб не мешала удалять по признаку цифры
    d_df = d_df[~d_df['Описание'].str.contains(r"zzz\D{,2}1{,1}\.{,1}$", regex=True)]  # Вопрос 1 в конце как признак битости строки
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d{,9}[,\. ]{,1}\d{1,9}\D{,2}руб[\.л][яе]?[й ,\.]?", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d+ ?(т.)?р/\.{,2}", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d{,9}[,\. ]{,1}\d{1,9}\D{,2}коп[\.е][йие]?", "", regex=True)  # 1!!!
    d_df['Описание'] = d_df['Описание'].str.replace(r"(\d)[еягй]?\-?[еягй]?о?", r"\1", regex=True)                      # 9-й в 9 мая и тп
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d микрорайон", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d{1,3} ?%", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d{1,2} ?градуса?о?в?х?", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d{1,3} ?суто?ки?", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d+ ?\-?й? ?д[не][яне][ьй]?\b", "", regex=True)                     # дней
    d_df['Описание'] = d_df['Описание'].str.replace(r"[012]?[0-9]:[012]?[0-9]", "", regex=True)  # time6968614-1
    d_df['Описание'] = d_df['Описание'].str.replace(r"[эЭ]таж ?\d{1,2}", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d{1,2} ?[эЭ]таже?", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"[оО]ценка \d", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"[рР]ассмотрение \d", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"[вВ]опрос \d", "", regex=True)
    d_df['Описание'] = d_df['Описание'].str.replace(r"\d[хм]", "", regex=True)#м - спорно
    d_df['Описание'] = d_df['Описание'].str.replace(r"\D0\d*", "", regex=True)  # м - спорно
    d_df = d_df[d_df['Описание'].str.contains(r'\d{1,}', regex=True) | (d_df['корд'].str.len() > 2)]  # есть хотябы одна цифра
    d_df = d_df.drop_duplicates(subset=['Номер ЕЦУР'])  # 50 если описание то 100
    print(d_df['Описание'])
    d_df.to_excel('./chk1.xlsx')
    return d_df
def rasp(x):
    global s_adf
    global t
    if x["Исполнитель"] == "Непосредственное управление":
        s_adf1 = s_adf[(s_adf.ОМСУ == x["ОМСУ"])]
        s_adf1 = s_adf1.drop_duplicates(["ОМСУ", "дом", "улица"])
    else:
        s_adf1 = s_adf[(s_adf.ОМСУ == x["ОМСУ"])&(s_adf.УК == x["Исполнитель"])]
    s_adf1['есхожд'] = 0#!!!!!!!!!!!!!!!!!убрать за функцию в с_адф
    s_adf1['улица'] = s_adf1['улица'].str.replace(r'.+, ', "", regex=True)  # сначала с краю от запятой
    a = str(x["Описание"]).lower()  # входящий на функцию
    d = s_adf1.columns.get_loc('дом')
    e = s_adf1.columns.get_loc('есхожд')  # столбец для суммы
    u = s_adf1.columns.get_loc("улица")
    for i in range(0, len(s_adf1)):
        b = "z" + s_adf1.iloc[i, u].lower() + "z"  # ул
        c = "#" + str(s_adf1.iloc[i, d]).lower() + '#'  # дом обособлено для поиска границ номера
        f = a.find(b)
        if f >= 0 and a.find(c) >= 0:  # ИМЕЕТ ДУБЛЕРА, АЛЯРМА and (x["Описание"].find(str(s_adf1.iloc[i]["дом"]))>1
            s_adf1.iloc[i, e] = 1
    t = t+1
    if int(t/100) == t/100:
        print(t)
    if s_adf1['есхожд'].sum() == 1:
        return s_adf1[s_adf1["есхожд"] == 1]["Адрес в образце"].values[0]
    elif s_adf1['есхожд'].sum() == 0 and not x["Исполнитель"] == "Непосредственное управление":
        s_adf1 = s_adf[(s_adf.ОМСУ == x["ОМСУ"])]
        s_adf1 = s_adf1.drop_duplicates(["ОМСУ", "дом", "улица"])
        s_adf1['есхожд'] = 0
        s_adf1['улица'] = s_adf1['улица'].str.replace(r'.+, ', "", regex=True)  # сначала с краю от запятой
        for i in range(0, len(s_adf1)):
            b = "z" + s_adf1.iloc[i, u].lower() + "z"  # ул
            c = "#" + str(s_adf1.iloc[i, d]).lower() + '#'  # дом обособлено для поиска границ номера
            f = a.find(b)
            # if x["Номер ЕЦУР"] == "7658034-1" and c == "#13#":
            #     print(a.find(c))
            #     print(b)
            #     print(c)
            #     print(d)
            #     print(e)
            #     print(f)
            if f >= 0 and a.find(c) >= 0:
                s_adf1.iloc[i, e] = 1
        if s_adf1['есхожд'].sum() == 1:
            g = str(s_adf1[s_adf1["есхожд"] == 1]["Адрес в образце"].values[0]) + "КОСЯК УК"
            return g
        elif s_adf1['есхожд'].sum() > 1:
            return "СТАЛО НЕПОНЯТНО"
        else:
            return "НЕ НАЙДЕНО"
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
    d_df = pd.read_excel("./chk.xlsx")
    d_df = chk1(d_df)  # выявление битых
    print('chk1 out')
s_adf = pd.read_excel("./адрес стандарт.xlsx")
s_adf["ОМСУ"] = s_adf["ОМСУ"].str.replace(r' г о', "", regex=True)
s_adf["улица"] = s_adf["улица"].str.replace(r' \d+', "", regex=True)
s_adf["улица"] = s_adf["улица"].str.replace(r'\d+ ', "", regex=True)
s_adf["улица"] = s_adf["улица"].str.replace(r'\-?\d+', "", regex=True)
s_adf["дом"] = s_adf["дом"].str.replace(r'/', "к", regex=True)
try:
    if len(d_df) < 10:
        print("low len")
except:
    d_df = pd.read_excel("./chk1.xlsx")
    print("chk1 eated")
t = 0
d_df['Описание'] = d_df['Описание'].str.replace(r'/', "к", regex=True)
d_df['Описание'] = d_df['Описание'].str.replace(r',? ?корп[\.]?у?с?[ао]?м? ?', "к", regex=True)
d_df['Описание'] = d_df['Описание'].str.replace(r'(\d+),? ?стр(оение)?\.? ?(\d+)', r"\1к\3", regex=True)
d_df['Описание'] = d_df['Описание'].str.replace(r'(\d)[ ,\.]{,2}(к)[ \.]{,2}(\d{1,2})', r"\1\2\3", regex=True)
d_df['Описание'] = d_df['Описание'].str.replace(r'московская обла?с?т?ь?', "", regex=True, flags=re.IGNORECASE)
d_df['Описание'] = d_df['Описание'].str.replace(r'(\d+ )в ', r"\1", regex=True, flags=re.IGNORECASE)
d_df['Описание'] = d_df['Описание'].str.replace(r'(?<!коловская )(?<!Чехов )(?<!завода )(?<!квартал )(\d+)(?! лет)(?!0 лет)(?! мая)(?! март)(?! кв)(?! сев)(?! дачн)(?! желез)(?! чапа)(?! ур)(?! пионер)(?! первом)(?! гвард) ?\-?(/?[абфкв]{,2}(?![а-я]))? ?(\d*)', r"#\1\2\3#", regex=True, flags=re.IGNORECASE)
d_df['Описание'] = d_df['Описание'].str.replace(r'77 ?ж\b', "#77ж#", regex=True, flags=re.IGNORECASE)
d_df['Описание'] = d_df['Описание'].str.replace(r'([а-яА-Я]{3,})', r"z\1z", regex=True, flags=re.IGNORECASE)
d_df["распознано"] = d_df.apply(rasp, axis=1)
d_df.to_excel("./chk2.xlsx")

#pd.options.display.max_colwidth = 1000
#print(str(d_df[d_df["Номер ЕЦУР"] == "7500157-2"]["Описание"]).encode("UTF-8"))7258977-1
#print(d_df[d_df["Номер ЕЦУР"] == "7258977-1"]["Описание"])
#print(b'\xb2\xd0\xb0 -5910223 \xd0\x9e\xd1'.decode("utf-8", "ignore"))

