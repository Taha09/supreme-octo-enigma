# -*- coding: utf-8 -*-
"""
Created on Sat Mar  6 15:50:27 2021

@author: Adminpyh
"""

import datetime
import itertools
import operator
import os
import sys
import tkinter.filedialog as fd
import zipfile
from collections import OrderedDict, defaultdict
from datetime import timedelta
from io import BytesIO
from itertools import combinations
from math import isnan
from tkinter import *
from tkinter import messagebox, ttk

import numpy as np
import pandas as pd
import tabula
from PyPDF2 import PdfFileReader
from tkcalendar import Calendar, DateEntry

mid = {'UK': 'RAMBSPGBP',
       'Norway': 'RAMBSPNORNOK',
       'Austria': 'RAMBSPAUTEUR',
       'Belgium': 'RAMBSPBLXEUR',
       'France': 'RAMBSPFRAEUR',
       'Germany': 'RAMBSPGEREUR',
       'Italy': 'RAMBSPITAEUR',
       'Netherland': 'RAMBSPNLEUR',
       'Portugal': 'RAMBSPPOREUR',
       'Spain': 'RAMBSPESPEUR',
       'Canada': 'RAMBSPCAD',
       'Sweden': 'RAMBSPSWESEK',
       'Denmark': 'RAMBSPDKDKK',
       'Switzerland': 'RAMBSPCHF',
       'Turkey': 'RAMBSPTURTRY',
       'USA': 'RAMARCUSAUSD'}


mid1 = {'UK': 22424124,
        'Norway': 22425694,
        'Austria': 22426744,
        'Belgium': 22427554,
        'France': 22427974,
        'Germany': 22428024,
        'Italy': 22428104,
        'Netherland': 22433534,
        'Portugal': 22433874,
        'Spain': 22434184,
        'Canada': 22434844,
        'Sweden': 22434924,
        'Denmark': 22436464,
        'Switzerland': 22436964,
        'Turkey': 22434264,
        'USA': 22434424
        }
mid2 = {'UK': 'Cpt 00112 00010006145',
        'Norway': 'Cpt 00112 00010006115',
        'Austria': 'Cpt 00112 00010006114',
        'Belgium': 'Cpt 00112 00010006102',
        'France': 'Cpt 00112 00010006108',
        'Germany': 'Cpt 00112 00010006107',
        'Italy': 'Cpt 00112 00010006104',
        'Netherland': 'Cpt 00112 00010006105',
        'Portugal': 'Cpt 00112 00010006106',
        'Spain': 'Cpt 00112 00010006103',
        'Canada': 'Cpt 00112 00010006108',
        'Sweden': 'Cpt 00112 00010006110',
        'Denmark': 'Cpt 00112 00010006111',
        'Switzerland': 'Cpt 00112 00010006147',
        'Turkey': 'Cpt 00112 00010006112',
        'USA': 'Cpt 00112 00010006146'}

mid3 = {'RAMBSPGBP': 'GB',
        'RAMBSPNORNOK': 'NO',
        'RAMBSPAUTEUR': 'AT',
        'RAMBSPBLXEUR': 'BE',
        'RAMBSPFRAEUR': 'FR',
        'RAMBSPGEREUR': 'DE',
        'RAMBSPITAEUR': 'IT',
        'RAMBSPNLEUR': 'NL',
        'RAMBSPPOREUR': 'PT',
        'RAMBSPESPEUR': 'ES',
        'RAMBSPCAD': 'CA',
        'RAMBSPSWESEK': 'SE',
        'RAMBSPDKDKK': 'DK',
        'RAMBSPCHF': 'CH',
        'RAMBSPTURTRY': 'TR',
        'RAMARCUSAUSD': 'US'}

mid4 = {
    'AT': 9402712103,
    'BE': 9412709966,
    'CA': 9592767651,
    'CH': 9582799789,
    'DE': 9502700586,
    'DK': 9472712108,
    'ES': 9562710475,
    'FR': 9492713391,
    'GB': 9422701359,
    'IT': 9522700350,
    'MA': 9592785505,
    'NL': 9532710795,
    'NO': 9542712104,
    'PT': 9552700296,
    'SE': 9572712107,
    'US': 9452798150
}

countries = list(mid.keys())
countries_nbr = list(mid1.keys())


root = Tk()
root.title('Amex Auto')
root.geometry('{}x{}'.format(640, 640))
root.resizable(0, 0)

# create all of the main containers
header = Frame(root, bg='#3B25BD', width=550, height=40, padx=3, pady=3)
center_frame_big = LabelFrame(root, text='Configuration', font=(
    'Tahoma', 11), bg='#FCFCFC', width=550, height=50, bd=1, relief=SOLID)
btm_frame = LabelFrame(root, text='Chargement des données', font=(
    'Tahoma', 11), bg='#FCFCFC', width=550, height=50, bd=1, relief=SOLID)
submit_frame = Frame(root, bg='#FCFCFC', width=550, height=60, padx=3, pady=3)
# progressbar_frame = Frame(root, bg='#FCFCFC', width=550, height=60, padx=3, pady=3)


header.grid(row=0, sticky="ew")
center_frame_big.grid(padx=10, row=1, sticky="ew")
btm_frame.grid(padx=10, row=2, sticky="ew")
submit_frame.grid(padx=10, row=3, sticky="ew")
#progressbar_frame.grid(padx=10,row=4, sticky="ew")


label_solution = Label(header, text='BSP Réconciliation',
                       fg='white', bg='#3B25BD', font=('Tahoma', 14))

label_solution.place(relx=0.5, rely=0.5, anchor=CENTER)

center_frame_big.grid_rowconfigure(0, weight=1)
center_frame_big.grid_columnconfigure(1, weight=1)


left_frame = LabelFrame(center_frame_big, text='Date', font=(
    'Tahoma', 11), bg='white', width=275, height=40, pady=30, bd=1, relief=SOLID)


"""cal = Calendar(left_frame, selectmode = 'day', 
               year = 2020, month = 5, 
               day = 22) """

cal_start = DateEntry(left_frame, width=30, height=100, background='darkblue',
                      foreground='white', borderwidth=2, font=('Tahoma', 11), date_pattern='dd/mm/y')


cal_end = DateEntry(left_frame, width=30, height=100, background='darkblue',
                    foreground='white', borderwidth=2, font=('Tahoma', 11), date_pattern='dd/mm/y')


datestart_label = Label(left_frame, text='Date début',
                        bg='white', font=('Tahoma', 11))
datefin_label = Label(left_frame, text='Date fin',
                      bg='white', font=('Tahoma', 11))

left_frame.grid(padx=20, pady=20, column=0, sticky="nsew")
datestart_label.grid_configure(padx=10)
datestart_label.grid(row=0, column=0, sticky="w")
datefin_label.grid_configure(padx=10)
datefin_label.grid(row=2, column=0, sticky="w")


cal_start.grid_configure(padx=10, pady=5)
cal_start.grid(row=1, column=0, sticky="w")
cal_end.grid_configure(padx=10, pady=5)
cal_end.grid_configure(row=3, column=0, sticky="w")


right_frame = LabelFrame(center_frame_big, text='Type', font=(
    'Tahoma', 11), bg='white', width=275, height=40, pady=30, bd=1, relief=SOLID)

pays_label = Label(right_frame, text='Pays', bg='white',
                   pady=5, font=('Tahoma', 11))
comboPays = ttk.Combobox(right_frame,
                         values=countries, font=('Tahoma', 11))


# layout the widgets in the right frame

right_frame.grid(padx=20, pady=20, row=0, column=1, sticky="nsew")
comboPays.grid_configure(padx=10, pady=5)
pays_label.grid(column=0, row=1)
comboPays.grid(column=1, row=1)
comboPays.current(1)


#filez = fd.askopenfilenames(parent=btm_frame, title='Choose a file')
#filez.grid(row=4, column=1)


label_file_banque = Label(btm_frame, text='Banque',
                          bg='white', pady=5, font=('Tahoma', 11))
label_file_bsp = Label(btm_frame, text='BSP', bg='white',
                       pady=5, font=('Tahoma', 11))
label_file_rejet = Label(btm_frame, text='Rejet',
                         bg='white', pady=5, font=('Tahoma', 11))
label_file_taux = Label(btm_frame, text='taux',
                        bg='white', pady=5, font=('Tahoma', 11))
label_zip_path = Label(btm_frame, text='Zip Path', bg='white', pady=5, font=('Tahoma', 11))

entry_file_banque = Entry(btm_frame, background="white",
                          width=50, bd=3, font=('Tahoma', 11))
entry_file_bsp = Entry(btm_frame, background="white",
                       width=50, bd=3, font=('Tahoma', 11))
entry_file_rejet = Entry(
    btm_frame, width=50, background="white", bd=3, font=('Tahoma', 11))
entry_file_taux = Entry(btm_frame, width=50,
                        background="white", bd=3, font=('Tahoma', 11))
entry_zip_path = Entry(btm_frame, background="white",
                          width=50, bd=3, font=('Tahoma', 11))


def set_text(text, input_file):
    input_file.delete(0, END)
    input_file.insert(0, text)
    return

def openZipPath():
    filepath = fd.askdirectory()
    set_text(filepath, entry_zip_path)

def openFileBanque():
    filepath = fd.askopenfilename(
        title="Banque Excel File",
        filetypes=[
            ("Excel files", "*.xlsm")])
    set_text(filepath, entry_file_banque)


def openFilerejet():
    filepath = fd.askopenfilename(
        title="Rejet Excel File",
        filetypes=[
            ("Excel files", "*.xls")])
    set_text(filepath, entry_file_rejet)


def openFileTaux():
    filepath = fd.askopenfilename(
        title="Taux File",
        filetypes=[
            ("Excel files", "*.xlsx")])

    set_text(filepath, entry_file_taux)


button_file_banque = Button(
    btm_frame, text="Charger", command=openFileBanque, font=('Tahoma', 11))
button_file_rejet = Button(btm_frame, text="Charger",
                           command=openFilerejet, font=('Tahoma', 11))
button_file_taux = Button(btm_frame, text="Charger",
                          command=openFileTaux, font=('Tahoma', 11))
button_zip_path = Button(
    btm_frame, text="Charger", command=openZipPath, font=('Tahoma', 11))

entry_file_banque.grid_configure(padx=10, pady=10)
entry_file_rejet.grid_configure(padx=10, pady=10)
entry_file_taux.grid_configure(padx=10, pady=10)
entry_zip_path.grid_configure(padx=10, pady=10)


label_file_banque.grid_configure(padx=10, pady=5)
label_file_rejet.grid_configure(padx=10, pady=10)
label_file_taux.grid_configure(padx=10, pady=10)
label_zip_path.grid_configure(padx=10, pady=5)



button_file_banque.grid_configure(padx=10, pady=10)
button_file_rejet.grid_configure(padx=10, pady=10)
button_file_taux.grid_configure(padx=10, pady=10)
button_zip_path.grid_configure(padx=10, pady=10)



label_file_banque.grid(row=0, column=0, sticky="nsew")
entry_file_banque.grid(row=0, column=1, sticky="nsew")
button_file_banque.grid(row=0, column=2, sticky="nsew")


label_file_rejet.grid(row=2, column=0, sticky="nsew")
entry_file_rejet.grid(row=2, column=1, sticky="nsew")
button_file_rejet.grid(row=2, column=2, sticky="nsew")


label_file_taux.grid(row=3, column=0, sticky="nsew")
entry_file_taux.grid(row=3, column=1, sticky="nsew")
button_file_taux.grid(row=3, column=2, sticky="nsew")

label_zip_path.grid(row=4, column=0, sticky="nsew")
entry_zip_path.grid(row=4, column=1, sticky="nsew")
button_zip_path.grid(row=4, column=2, sticky="nsew")

T =[]
def submit():
    global T

    # print(dir_bsp)
    # print(file_worldpay)
    file_banque = entry_file_banque.get()
    zippath = entry_zip_path.get()
    file_rejet = entry_file_rejet.get()
    file_taux = entry_file_taux.get()
    # print(file_banque)
    marche = comboPays.get()
    marche1 = comboPays.get()
    # print(marche)
    marche_nbr = mid1[marche]
    banque_num_compte = mid2[marche]
    marche = mid[marche]
    marche_pdf_bsp = (mid3[marche])

    date_rec_debut = datetime.datetime.strptime(cal_start.get(), '%d/%m/%Y')

    date_rec_fin = datetime.datetime.strptime(cal_end.get(), '%d/%m/%Y')

    delta = date_rec_fin - date_rec_debut

    banque = pd.read_excel(file_banque, skiprows=4,
                           sheet_name=banque_num_compte)
    banque = banque[(banque['Date'].notnull()) & (
        banque['Date'] != 'Liste de vos comptes')]
    banque['Date'] = pd.to_datetime(banque['Date'])
    banque = banque[banque['Libellé'].str.contains('AMEX')]
    if marche == 'RAMBSPCAD':
        banque = banque[banque['Libellé'].str.contains('9592767651')]
    elif marche == 'RAMBSPFRAEUR':
        banque = banque[banque['Libellé'].str.contains('9492713391')]
    dic_banque = defaultdict(list)
    for c, i in zip(banque.Date, banque['Crédit']):
        dic_banque[c].append(i)

    marche_rejet = mid4[marche_pdf_bsp]
    try:
        rejet = pd.read_excel(file_rejet, skiprows=9)
        rejet = rejet[rejet['Payee merchant number'] == marche_rejet]
        rejet = rejet.sort_values(by=['Settlement date'])
        #rejet = rejet[(rejet['Post Date']>= date_rec_debut - timedelta(5)) & (rejet['Post Date']<= date_rec_fin + timedelta(5)) ]
        list_rejet = list(rejet['Settlement amount'])

    except:
        pass

    try:
        taux = pd.read_excel(file_taux)
        taux = taux[['De', 'en', 'Taux', 'Déb.valid.']]


# taux['Déb.valid.']=pd.to_datetime(taux['Déb.valid.'],format='%d.%m.%Y').strftime("%Y-%m-%d")

        taux['Déb.valid.'] = taux['Déb.valid.'].str.replace(
            '.', '-', regex=True)
        taux2 = taux[(taux['De'] == 'CAD') & (taux['en'] == 'EUR')]
        taux2['Déb.valid.'] = pd.to_datetime(
            taux['Déb.valid.'], format='%d-%m-%Y')

        param = "/"

    except:
        pass

    bsp_dic_global = []

    for i in range(delta.days + 1):
        day = date_rec_debut + timedelta(days=i)

        if len(str(day.month)) == 1:
            day_month = '0' + str(day.month)
        else:
            day_month = str(day.month)
        year_day = str(day.year)
        if len(str(day.day)) == 1:
            day_day = '0' + str(day.day)
        else:
            day_day = str(day.day)
        try:
            #path_zip = "//casnas01/RAM_ARCHIVAGE/DP_BSPLINK/bsp "+year_day+"/bsp"+year_day+ marche_pdf_bsp.lower() +"/"+ year_day + marche_pdf_bsp.lower()+ day_month + "/"+ marche_pdf_bsp + "az1470_" + year_day+ day_month+ day_day + "_Airline_Daily.zip"
            path_zip = zippath + "/" + marche_pdf_bsp + \
                "az1470_" + year_day + day_month + day_day + "_Airline_Daily.zip"
            with zipfile.ZipFile(path_zip, 'r') as z:

                l = [string for string in z.namelist() if "PCAIDLYSUM" in string]
                if len(l) == 1:
                    file_bsp = str(l[0])
                else:
                    print('file not found or too many files')

                try:
                    a = datetime.datetime.strptime(file_bsp.split('.')[
                                                   0].split('_')[-1], '%y%m%d')
                except:
                    try:
                        a = datetime.datetime.strptime(file_bsp.split('.')[
                                                       0].split('_')[-1], '%Y%m%d')
                    except:
                        a = datetime.datetime.strptime(file_bsp.split('.')[
                                                       0].split('_')[-2], '%m%d').replace(year=2021)
        except:
            #path_zip = "//casnas01/RAM_ARCHIVAGE/DP_BSPLINK/bsp "+year_day+"/bsp"+year_day+ marche_pdf_bsp.lower() +"/"+ year_day + marche_pdf_bsp.lower()+ day_month + "/"+ marche_pdf_bsp + "az1470_" + year_day+ day_month+ day_day +"_"+year_day+ day_month+ day_day + "_Airline_Daily.zip"
            path_zip = zippath + "/" + marche_pdf_bsp + \
                "az1470_" + year_day + day_month + day_day + "_Airline_Daily.zip"

        with zipfile.ZipFile(path_zip, 'r') as z:

            l = [string for string in z.namelist() if "PCAIDLYSUM" in string]
            if len(l) == 1:
                file_bsp = str(l[0])
            else:
                print('file not found or too many files')

            try:
                a = datetime.datetime.strptime(file_bsp.split(
                    '.')[0].split('_')[-1], '%y%m%d')
            except:
                try:
                    a = datetime.datetime.strptime(file_bsp.split(
                        '.', maxsplit=1)[0].split('_')[-1], '%Y%m%d')
                except:
                    a = datetime.datetime.strptime(file_bsp.split('.')[
                                                   0].split('_')[-2], '%m%d').replace(year=2021)

            try:
                bsp = tabula.read_pdf(BytesIO(z.read(file_bsp)), pages='all', area=(
                    282.261, 6.728, 554.534, 824.6))

                new = bsp[0]

                T = []
                new_header = new.iloc[0]
                new = new[1:]
                new.columns = new_header

                new = new[new['CARD TYPE'].notna()]

                try:
                    T = []
                    new_visa = new[new['CARD TYPE'].str.contains(
                        'AX American Express')]

                    new_visa = new_visa.dropna(axis=1)
                    tst_columns = list(new_visa.columns)
                    lst_temp = []
                    for i in range(len(tst_columns)):
                        col = tst_columns[i]

                        if 'ISSUES' in str(col):
                            lst_temp.append(col)
                        if 'REFUNDS' in str(col):
                            lst_temp.append(col)

                    new_visa = new_visa[lst_temp]

                    a1 = new_visa.iloc[0, 0].split('  ')

                    if len(a1) != 2:
                        a1 = new_visa.iloc[0, 0].split(' ')

                    a2 = new_visa.iloc[0, 1].split('  ')
                    if len(a2) != 2:
                        a2 = new_visa.iloc[0, 1].split(' ')

                    visa = a1 + a2

                    for i in visa:
                        if i == '':
                            visa.remove(i)

                    if len(visa) == 3:
                        if visa[2] == '0':
                            visa.append(0)
                        if visa[0] == '0':
                            visa.insert(1, 0)
                    visa = [float(str(x).replace(',', '')) for x in visa]

                    visa = [visa[1], visa[3]]

                    T.extend(visa)

                except:
                    pass

            except:
                visa = [0, 0]

            try:

                bsp = tabula.read_pdf(BytesIO(z.read(file_bsp)), pages='all', area=(
                    102.497, 11.984, 547.176, 829.857))
                try:
                    for i in range(1, len(bsp)):
                        try:
                            new = bsp[i]

                            new = new[new[new.columns[0]].notna()]

                            new_visa = new[new[new.columns[0]].str.contains(
                                'AX American Express')]

                            #new_visa = new_visa.dropna(axis=1)
                            tst_columns = list(new_visa.columns)
                            new_visa.reset_index(drop=True, inplace=True)
                            visa = [new_visa['VALUE'].loc[0],
                                    new_visa['VALUE.1'].loc[0]]

                            T = []
                            visa = [0 if isnan(float(str(x).replace(',', ''))) else float(
                                str(x).replace(',', '')) for x in visa]

                            new = new[new[new.columns[0]].notna()]

                        except:
                            pass
                        try:

                            new = bsp[i]
                            new = new[new[new.columns[0]].notna()]

                            # new_visa = new[new[new.columns[0]].str.contains('AX American Express')]

                            new_visa = new[new[new.columns[0]].str.contains(
                                'AX American Express')]

                            new_visa = new_visa.dropna(axis=1)
                            tst_columns = list(new_visa.columns)
                            lst_temp = []
                            for m in range(len(tst_columns)):
                                col = tst_columns[m]

                                if 'ISSUES' in col or col == 'VALUE':
                                    lst_temp.append(col)
                                if 'REFUNDS' in col or col == 'VALUE.1':
                                    lst_temp.append(col)

                            new_visa = new_visa[[
                                '---------------ISSUES---------------', '---------------REFUNDS---------------']]

                            if new_visa.shape[0] == 0:
                                visa = [0, 0]
                            elif new_visa.shape[0] == 1:

                                if new_visa.columns[0] == 'VALUE':
                                    visa = [
                                        float(str(new_visa.iloc[0, 0]).replace(',', '')), 0]
                                elif new_visa.columns[0] == 'VALUE.1':
                                    visa = [
                                        0, float(str(new_visa.iloc[0, 1]).replace(',', ''))]
                            if new_visa.shape[0] == 1 and new_visa.shape[1] == 2:

                                visa = (
                                    new_visa.iloc[0, 0] + '  ' + new_visa.iloc[0, 1]).split('  ')

                            a1 = new_visa.iloc[0, 0].split('  ')
                            if len(a1) != 2:
                                a1 = new_visa.iloc[0, 0].split(' ')
                            a2 = new_visa.iloc[0, 1].split('  ')
                            if len(a2) != 2:
                                a2 = new_visa.iloc[0, 1].split(' ')
                            visa = a1 + a2
                            for x in visa:
                                if x == '':
                                    visa.remove(x)

                            if len(visa) == 3:
                                if visa[2] == '0':
                                    visa.append(0)
                                if visa[0] == '0':
                                    visa.insert(1, 0)
                            visa = [float(str(x).replace(',', ''))
                                    for x in visa]
                            visa = [visa[1], visa[3]]

                            T.extend(visa)

                        except:
                            pass

                except:
                    pass

            except:
                visa = [0, 0]

        try:
            del new
        except:
            pass
        try:
            del bsp
        except:
            pass

        if marche1 == "Canada":
            taux = pd.read_excel(file_taux)
            taux = taux[['De', 'en', 'Taux', 'Déb.valid.']]
            # taux['Déb.valid.']=pd.to_datetime(taux['Déb.valid.'],format='%d.%m.%Y').strftime("%Y-%m-%d")
            # taux['Déb.valid.']=taux['Déb.valid.'].str.replace('.','-',regex=True)

            # taux['Déb.valid.'].apply(lambda x: x.strftime('%d%m%Y'))
            taux2 = taux[(taux['De'] == 'CAD') & (taux['en'] == 'EUR')]
            taux3 = taux[(taux['De'] == 'MAD') & (taux['en'] == 'EUR')]

            taux2['Déb.valid.'] = pd.to_datetime(
                taux['Déb.valid.'], format='%d-%m-%Y')
            taux3['Déb.valid.'] = pd.to_datetime(
                taux['Déb.valid.'], format='%d-%m-%Y')

            # print(taux2)
            # taux2['Déb.valid.']=pd.to_datetime(taux['Déb.valid.'])
            # taux2.loc[pd.to_datetime(taux['Déb.valid.'],format='%d-%m-%Y'),'Déb.valid.']
            pd.to_datetime(taux2['Déb.valid.'], errors='coerce').dt.normalize()
            pd.to_datetime(taux3['Déb.valid.'], errors='coerce').dt.normalize()
            # taux3['Déb.valid.']=pd.to_datetime(taux['Déb.valid.'],format='%d-%m-%Y')
            taux4 = taux2[['Déb.valid.', 'Taux']]
            taux4.set_index('Déb.valid.', inplace=True)
            taux5 = taux3[['Déb.valid.', 'Taux']]
            taux5.set_index('Déb.valid.', inplace=True)
            param = "/"
            i = 0
            # print(list_date)
            dic_taux_MAD = taux4.to_dict()
            dic_taux_EUR = taux5.to_dict()
            # print(dic_taux_MAD)
            # print(dic_taux_EUR)
            subDate = a.strftime("%Y")+'-' + \
                a.strftime("%m")+'-'+a.strftime("%d")
            for i, row in taux4.iterrows():

                if i.strftime("%Y-%m-%d") == subDate:

                    val_taux = row['Taux']

                    T = np.multiply(T, round(val_taux, 2))

            # for i,row in taux5.iterrows():

            #     if i.strftime("%Y-%m")==subDate:
            #         val_taux=row['Taux']
            #         val_taux=float(val_taux[2:].replace(",","."))
            #         print(val_taux)
            #         T=np.divide(T,val_taux)

        if len(T) == 0:
            T = [0, 0]
            row_vi_encaissement = ['AX', 'ENCAISSEMENT', a.date(), T[0]]
            row_vi_remboursement = ['AX', 'REMBOURSEMENT', a.date(), T[1]]
            bsp_dic_global.append(row_vi_encaissement)
            bsp_dic_global.append(row_vi_remboursement)
            del (T)

        elif (len(T) == 2):
            # print(T)
            row_vi_encaissement = ['AX', 'ENCAISSEMENT', a.date(), T[0]]
            row_vi_remboursement = ['AX', 'REMBOURSEMENT', a.date(), T[1]]
            bsp_dic_global.append(row_vi_encaissement)
            bsp_dic_global.append(row_vi_remboursement)
            del (T)

        else:

            # T=list(OrderedDict.fromkeys(T))
            # print(T)
            # T=list(map(str,T))
            # for i in T:
            #     if i=="0":
            #         T.remove(i)
            # print(T)
            # T=[float(x) for x in T]

            # if (len(T)==3 and T[1]==0.0):
            #     row_vi_encaissement = ['AX','ENCAISSEMENT', a.date() ,T[0]]
            #     row_vi_remboursement = ['AX','REMBOURSEMENT', a.date(), T[1]]
            #     row_vi_encaissement1 = ['AX','ENCAISSEMENT', a.date() ,T[2]]
            #     row_vi_remboursement1 = ['AX','REMBOURSEMENT', a.date(), T[1]]

            # elif (len(T)==3 and T[0]==0.0 and T[2]<0):
            #     row_vi_encaissement = ['AX','ENCAISSEMENT', a.date() ,T[0]]
            #     row_vi_remboursement = ['AX','REMBOURSEMENT', a.date(), T[1]]
            #     row_vi_encaissement1 = ['AX','ENCAISSEMENT', a.date() ,T[0]]
            #     row_vi_remboursement1 = ['AX','REMBOURSEMENT', a.date(), T[2]]
            # elif (len(T)==3 and T[0]==0.0 and T[2]>0):
            #     row_vi_encaissement = ['AX','ENCAISSEMENT', a.date() ,T[0]]
            #     row_vi_remboursement = ['AX','REMBOURSEMENT', a.date(), T[1]]
            #     row_vi_encaissement1 = ['AX','ENCAISSEMENT', a.date() ,T[2]]
            #     row_vi_remboursement1 = ['AX','REMBOURSEMENT', a.date(), T[0]]
            # elif (len(T)==3 and T[2]==0.0):
            #     row_vi_encaissement = ['AX','ENCAISSEMENT', a.date() ,T[0]]
            #     row_vi_remboursement = ['AX','REMBOURSEMENT', a.date(), T[1]]

            # elif (len(T)==2):
            #     print("on va inserer T")
            #     print(T)

            #     row_vi_encaissement = ['AX','ENCAISSEMENT', a.date() ,T[0]]
            #     row_vi_remboursement = ['AX','REMBOURSEMENT', a.date(), T[1]]

            # print(T)
            # print("wewe")
            row_vi_encaissement = ['AX', 'ENCAISSEMENT', a.date(), T[0]]
            row_vi_remboursement = ['AX', 'REMBOURSEMENT', a.date(), T[1]]
            row_vi_encaissement1 = ['AX', 'ENCAISSEMENT', a.date(), T[2]]
            row_vi_remboursement1 = ['AX', 'REMBOURSEMENT', a.date(), T[3]]

            # print("insertion de T")
            # print(a.date())
            # print(T)
            try:
                bsp_dic_global.append(row_vi_encaissement)
                bsp_dic_global.append(row_vi_remboursement)
                bsp_dic_global.append(row_vi_encaissement1)
                bsp_dic_global.append(row_vi_remboursement1)
            except:
                pass

        
        # print(T)

    cols_output_bsp = ['CARTE', 'TYPE', 'DATE', 'AMEX']
    output_bsp = pd.DataFrame(bsp_dic_global, columns=cols_output_bsp)
    output_bsp['DATE'] = pd.to_datetime(output_bsp['DATE'])

    dic_wp = (output_bsp.groupby('DATE')['AMEX'].sum() * 97.25 / 100).to_dict()

    def output_lst(date_nouveau,  fees, rejet):
        lst_fees.append(fees)
        lst_dates_rec.append(date_nouveau)
        lst_encaissement.append(
            dic_banque[date + datetime.timedelta(days=j)][0])
        lst_dates_enc.append(date + datetime.timedelta(days=j))
        lst_rejet_output.append(rejet)
        return lst_fees, lst_dates_rec, lst_encaissement, lst_dates_enc, lst_rejet_output

    def sum_rejet(x):
        try:
            return sum(x)
        except:
            return x

    w = 0
    fees_list = [-0.01, 0.01, 0.02, -0.02, 0.03, 0.04, -0.03, -0.04]

    lst_fees = []
    lst_dates_rec = []
    lst_encaissement = []
    lst_dates_enc = []
    lst_rejet_output = []

    temp_lst_dates = []
    temp = list(dic_wp)

    maxi = []
    for date in list(dic_wp):

        now_wp = dic_wp[date]

        temp_lst_dates = []

        for t in range(90):

            temp = list(dic_wp)
            try:
                temp_lst_dates.append(temp[temp.index(date) + t])
                maxi.append(now_wp)
                for j in range(20, 120):
                    try:

                        try:
                            dic_banque[date + datetime.timedelta(days=j)]
                            if round(now_wp, 2) == w + dic_banque[date + datetime.timedelta(days=j)][0]:
                                #del dic_wp[date]
                                #del dic_banque[date]
                                print("uh oh", date, 'ok', date +
                                      datetime.timedelta(days=j))
                                for temp_date in temp_lst_dates:
                                    lst_fees, lst_dates_rec, lst_encaissement, lst_dates_enc, lst_rejet_output = output_lst(
                                        temp_date, 0.0, 0.0)
                                break
                        except:
                            pass
                        try:
                            for fees in fees_list:

                                if round(now_wp - fees, 2) == w + dic_banque[date + datetime.timedelta(days=j)][0]:

                                    for temp_date in temp_lst_dates:
                                        lst_fees, lst_dates_rec, lst_encaissement, lst_dates_enc, lst_rejet_output = output_lst(
                                            temp_date, fees, 0.0)
                        except:
                            pass
                        try:

                            for e, f in combinations(range(len(list_rejet) + 1), 2):
                                i = list_rejet[e:f]
                                p = round(sum(i), 2)

                                try:

                                    if round(now_wp, 2) + p - dic_banque[date + datetime.timedelta(days=j)][0] == 0:
                                        #del dic_wp[date]
                                        #del dic_banque[date]

                                        list_rejet = list_rejet[f:]

                                        for temp_date in temp_lst_dates:

                                            lst_fees, lst_dates_rec, lst_encaissement, lst_dates_enc, lst_rejet_output = output_lst(
                                                temp_date, 0.0, i)

                                        #print('okey', date,'ok',date +datetime.timedelta(days=j))
                                        break
                                except:
                                    pass
                                try:

                                    if round(now_wp, 2) - p - w - dic_banque[date + datetime.timedelta(days=j)][0] == 0:
                                        #del dic_wp[date]
                                        #del dic_banque[date]

                                        list_rejet = list_rejet[f:]
                                        for temp_date in temp_lst_dates:
                                            lst_fees, lst_dates_rec, lst_encaissement, lst_dates_enc, lst_rejet_output = output_lst(
                                                temp_date, 0.0, i)

                                        #print('okey', date,'ok',date +datetime.timedelta(days=j))
                                        break
                                except:
                                    pass
                                try:
                                    for fees in fees_list:
                                        if round(now_wp, 2) - fees - p == w + dic_banque[date + datetime.timedelta(days=j)][0]:

                                            #del dic_wp[date]
                                            #del dic_banque[date]
                                            list_rejet = list_rejet[f:]
                                            for temp_date in temp_lst_dates:
                                                lst_fees, lst_dates_rec, lst_encaissement, lst_dates_enc, lst_rejet_output = output_lst(
                                                    temp_date, fees, i)

                                            #print('okey', date,'ok',date +datetime.timedelta(days=j))
                                            break
                                except:
                                    pass
                                try:
                                    for fees in fees_list:
                                        if round(now_wp, 2) - fees + p == w + dic_banque[date + datetime.timedelta(days=j)][0]:
                                            #del dic_wp[date]
                                            #del dic_banque[date]
                                            list_rejet = list_rejet[f:]
                                            for temp_date in temp_lst_dates:
                                                lst_fees, lst_dates_rec, lst_encaissement, lst_dates_enc, lst_rejet_output = output_lst(
                                                    temp_date, fees, i)

                                            #print('okey', date,'ok',date +datetime.timedelta(days=j))
                                            break
                                except:
                                    pass
                                try:
                                    for fees in fees_list:
                                        if round(now_wp, 2) + fees - p == w + dic_banque[date + datetime.timedelta(days=j)][0]:

                                            #del dic_wp[date]
                                            #del dic_banque[date]
                                            list_rejet = list_rejet[f:]
                                            for temp_date in temp_lst_dates:
                                                lst_fees, lst_dates_rec, lst_encaissement, lst_dates_enc, lst_rejet_output = output_lst(
                                                    temp_date, fees, i)

                                            #print('okey', date,'ok',date +datetime.timedelta(days=j))
                                            break
                                except:
                                    pass
                                try:
                                    for fees in fees_list:
                                        if round(now_wp, 2) + fees + p == w + dic_banque[date + datetime.timedelta(days=j)][0]:
                                            #del dic_wp[date]
                                            #del dic_banque[date]
                                            list_rejet = list_rejet[f:]
                                            for temp_date in temp_lst_dates:
                                                lst_fees, lst_dates_rec, lst_encaissement, lst_dates_enc, lst_rejet_output = output_lst(
                                                    temp_date, fees, i)

                                            #print('okey', date,'ok',date +datetime.timedelta(days=j))
                                            break
                                except:
                                    pass

                        except:
                            pass

                    except:
                        pass
                now_wp = now_wp + dic_wp[temp[temp.index(date) + t + 1]]

            except:
                pass

    output_ecart2 = pd.DataFrame(list(zip(lst_dates_rec, lst_encaissement, lst_dates_enc, lst_rejet_output, lst_fees)),
                                 columns=['DATE', 'Montant Encaissé AMEX USD', "Date de règlement", "REJET", "Ecart AMEX"])

    output_ecart2['DATE'] = pd.to_datetime(output_ecart2['DATE'])
    output3 = output_bsp.merge(output_ecart2, how='outer', on='DATE')

    output3['Commission AMEX'] = output3['AMEX'] * 2.75 / 100
    output3['VENTES NET AMEX'] = output3['AMEX'] - output3['Commission AMEX']

    output3.rename(columns={'DATE': 'DATE VENTES/RBT',
                   'CARTE': 'Type CARTE'}, inplace=True)

    output6 = output3[(output3['DATE VENTES/RBT'] >= date_rec_debut)
                      & (output3['DATE VENTES/RBT'] <= date_rec_fin)]

    output6.to_excel('AX_auto_'+marche + '_'+date_rec_debut.strftime("%d-%m-%Y") +
                     '_' + date_rec_fin.strftime("%d-%m-%Y")+'.xlsx', index=False)
    messagebox.showinfo('Info', 'Process completed!')


button_submit = Button(submit_frame, text="Executer",
                       font=('Tahoma', 11), command=submit)


button_submit.place(relx=0.5, rely=0.5, anchor=CENTER)
root.mainloop()
