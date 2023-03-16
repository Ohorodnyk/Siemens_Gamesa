import os
from fnmatch import fnmatch
import subprocess
import tkinter
from datetime import time

import camelot
import tabula
import pandas as pd
import numpy as np
import PyPDF2
import fitz
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import csv

# import xlsxwrite
# from camelot import read_pdf
import os
# from tabula import read_pdf
# from tabulate import tabulate
import pandas as pd
import io
import cv2
from PIL import Image

from tkinter.filedialog import askopenfilename
messagebox.showinfo("showinfo", "Choose folder");
foldername = tkinter.filedialog.askdirectory()
messagebox.showinfo("showinfo", "Choose folder to save result files");
foldername_to_save = tkinter.filedialog.askdirectory()
messagebox.showinfo("showinfo", "Choose folder to save fotos");
foldername_to_save_foto = tkinter.filedialog.askdirectory()
pattern = "*.pdf"
df_Description = pd.read_excel('D:/Covizmo2/Knime/test_pdf/3properties.xlsx', sheet_name="Description")
df_BREAK_REASON = pd.read_excel('D:/Covizmo2//Knime/test_pdf/3properties.xlsx', sheet_name="BREAK REASON")
df_WorkToDo = pd.read_excel('D:/Covizmo2/Knime//test_pdf/3properties.xlsx', sheet_name="WorkToDo")

for path, subdirs, files in os.walk(foldername):
    for name in files:
        if fnmatch(name, pattern):
            path_to_save = os.path.join(path, name)
            print(path_to_save)
            print(subdirs)
            print(path)
            file2 = name[:-4]
            start = path_to_save.find('2015_pdfs\\')# розділювач для створення структури папок
            print(start)
            # path_to_save2 = os.path.join(foldername_to_save, a, file2)
            try:
                if start != -1:
                    a = path[start+5:]
                    # a.replace('\\','/')
                    print('a')
                    print(a)
                    print("//n")
                    print('norm path')
                    path_to_save2 = os.path.join(foldername_to_save, a, file2)
                    print(path_to_save2)
                    os.makedirs(path_to_save2)
                    # pdf_file = PyPDF2.PdfFileReader(name)
                    # print('pdf',pdf_file)

            except Exception:
                pass
            try:
                    pdf_file = PyPDF2.PdfFileReader(path_to_save)
                    print(pdf_file)

                    # Name of input file
                    # a = os.path.basename(file);

                    # messagebox.showinfo("showinfo", "Choose folder to save result files");
                    # foldername = tkinter.filedialog.askdirectory()
                    # directory = "final"
                    # path = os.path.join(foldername+"/", file2)
                    # os.mkdir(path)

                    # ------------- read tables and save CSV, XSLX
                    # Read a PDF File
                    df = tabula.read_pdf(path_to_save, pages='all')[0]

                    # output = subprocess.check_output("ping -c 2 -W 2 1.1.1.1; exit 0", stderr=subprocess.STDOUT, shell=True)
                    # convert PDF into CSV
                    tabula.convert_into(path_to_save, path_to_save2 + "/test.csv", output_format="csv",
                                        pages='all')  # -- save csv file
                    try:
                        subprocess.check_output("ping -c 2 -W 2 1.1.1.1; exit 0", stderr=subprocess.STDOUT, shell=True)
                    except subprocess.CalledProcessError as e:
                        print(e.returncode)
                        print(e.output)
                    # print(df)

                    col_names = ['No_CODE', 'ZONE_SHELL_ROOT', 'LENGTH_TRANS', 'empty1', 'val1', 'val2', 'val3', 'val4',
                                 'val5', 'val6', 'val7', 'val8', 'val9','val10','val11']
                    tbl = pd.read_csv(path_to_save2 + '/test.csv', sep=',', names=col_names, encoding="ISO-8859-1")

                    # General data of inspection
                    info = tbl[
                        (tbl['No_CODE'] == "WIND FARM:") | (tbl['No_CODE'] == "PARQUE:") | (tbl['No_CODE'] == "FL") | (
                                    tbl['No_CODE'] == "UT")
                        | (tbl['No_CODE'] == "WTG:") | (tbl['No_CODE'] == "AEROGENERADOR:") | (
                                    tbl['No_CODE'] == "CLIENT:") | (tbl['No_CODE'] == "CLIENTE:")
                        | (tbl['No_CODE'] == "WO No") | (tbl['No_CODE'] == "No OT") | (tbl['No_CODE'] == "SC") | (
                                    tbl['No_CODE'] == "SC RESPONSIBLE")
                        | (tbl['No_CODE'] == "RESPONSABLE SC") | (tbl['No_CODE'] == "SUPERVISOR") | (
                                    tbl['No_CODE'] == "WTG MODEL") | (tbl['No_CODE'] == "MODELO MAQUINA")
                        | (tbl['No_CODE'] == "COUNTRY") | (tbl['No_CODE'] == "PAIS") | (
                                    tbl['No_CODE'] == "PRODUCTION (kWh)") | (tbl['No_CODE'] == "PRODUCCION (kWh)")
                        | (tbl['No_CODE'] == "LINE HOURS OK") | (tbl['No_CODE'] == "HORAS LINEA OK") | (
                                    tbl['No_CODE'] == "WTG HOURS OK") | (tbl['No_CODE'] == "HORAS TURBINA OK")]
                    info = info[['No_CODE', 'ZONE_SHELL_ROOT']]
                    info = info.rename(columns={'No_CODE': 'ColName', 'ZONE_SHELL_ROOT': 'Value'})
                    print("\nInfo:")
                    print(info)

                    filtr = tbl[(tbl['No_CODE'] == "BLADE No") | (tbl['No_CODE'] == "No PALA")]
                    filtr = filtr[['No_CODE', 'ZONE_SHELL_ROOT']]
                    filtr = filtr.rename(columns={'No_CODE': 'ColName', 'ZONE_SHELL_ROOT': 'BladeNo'})
                    print("\nBlades:")
                    filtr = filtr.drop(['ColName'], axis=1, inplace=False)

                    df4 = pd.DataFrame()
                    df4['No'] = filtr['BladeNo']
                    df44 = pd.DataFrame()
                    df44 = pd.DataFrame({'N2': ['A', 'B', 'С']})

                    print(filtr)

                    # number of blades
                    count_blade = filtr.shape[0]

                    df_dates = tbl[(tbl['LENGTH_TRANS'].astype(str) == "START DATE") | (
                                tbl['LENGTH_TRANS'].astype(str) == "FECHA INICIO") |
                                   (tbl['LENGTH_TRANS'].astype(str) == "END DATE") | (
                                               tbl['LENGTH_TRANS'].astype(str) == "FECHA FIN")]
                    print(df_dates)
                    df_dates = df_dates[['LENGTH_TRANS', 'empty1']]

                    df_dates = df_dates.rename(columns={'LENGTH_TRANS': 'ColName', 'empty1': 'Data'})
                    print("\nStare date and End date:")
                    print(df_dates)

                    # delete unnecessary data
                    rows_with_nan = [index for index, row in tbl.iterrows() if row.isnull().all()]
                    shape = tbl.shape

                    data = tbl

                    # """
                    # creating column with all info about damage
                    data['NewRow'] = ""
                    data['NewRow'] = np.where((data['No_CODE'].str[0] == "0") & (data['No_CODE'].str.len() == 2),
                                              data['No_CODE'].astype(str) + " " +
                                              data['ZONE_SHELL_ROOT'].astype(str) + " " + data['LENGTH_TRANS'].astype(
                                                  str) + " " + data['empty1'].astype(str) + " " + data['val1'].astype(
                                                  str) + " " + data['val2'].astype(str) + " " + data['val3'].astype(
                                                  str) + " " + data['val4'].astype(str), "Empty")
                    data['NewRow'] = np.where((data['No_CODE'].str[0] == "0") & (data['No_CODE'].str.len() != 2),
                                              data['No_CODE'].astype(str) + " " + data['ZONE_SHELL_ROOT'].astype(
                                                  str) + " " + data['LENGTH_TRANS'].astype(str) + " " + data[
                                                  'empty1'].astype(str), data['NewRow'])
                    # data['Work'] = np.where((data['No_CODE'].str[0] == "0") & (data['No_CODE'].str.len() == 2)
                    data['WorkToDo'] = ""
                    data['WorkToDo'] = np.where((data['No_CODE'].str[0] == "0") & (data['No_CODE'].str.len() == 2),
                                                data['val6'].astype(str), "Empty")
                    data['WorkToDo'] = np.where((data['No_CODE'].str[0] == "0") & (data['No_CODE'].str.len() != 2) & (
                                data['val2'].astype(str) == "nan"), data['val3'].astype(str), data['WorkToDo'])
                    data['WorkToDo'] = np.where((data['No_CODE'].str[0] == "0") & (data['No_CODE'].str.len() == 2) & (
                                data['val2'].astype(str) != "nan") & (data['WorkToDo'].astype(str) == "Empty"),
                                                data['val2'].astype(str), data['WorkToDo'])

                    blades = filtr['BladeNo'].to_numpy()
                    # print("\nAll blades:")
                    # print(blades)

                    print("\nDam Data:")
                    print(data)

                    data.to_excel(path + '/DamData-2.xlsx', index=None,
                                  header=True)  # creating file with necessary data only

                    # add missed last row
                    data = data.append({'No_CODE': "DAMAGE", 'ZONE_SHELL_ROOT': "POSITION (mm)", 'NewRow': "Empty"},
                                       ignore_index=True)
                    # data['No_CODE'] = data['No_CODE'].astype(str)
                    data = data[(data['No_CODE'].astype(str) != "No") & (data['No_CODE'].astype(str) != "nan") & (
                                data['No_CODE'].astype(str) != "NaN")]
                    print('data')
                    data.reset_index(drop=True)
                    data.index = pd.RangeIndex(len(data.index))
                    print(data)

                    data = data[(data.No_CODE == '1') | (data.No_CODE == '2') | (data.No_CODE == '3') | (
                                data.No_CODE == '4') | (data.No_CODE == '5') | (data.No_CODE == '6') | (
                                            data.No_CODE == '7') | (data.No_CODE == '8') | (data.No_CODE == '9')]

                    # data.drop(data[data['ZONE_SHELL_ROOT'] == np.nan].index, inplace = True)
                    data = data[data['ZONE_SHELL_ROOT'].notnull()]
                    data = data.drop(['WorkToDo'], axis=1, inplace=False)
                    data = data.rename(
                        columns={'No_CODE': 'No', 'ZONE_SHELL_ROOT': 'Code', 'val3': 'Description', 'val5': 'WorkToDo',
                                 'val7': 'BREAK REASON'})
                    data.loc[(data['LENGTH_TRANS'] == 'X'), ('LENGTH_TRANS')] = '0 X'
                    data['Zone'] = data['LENGTH_TRANS'].str.split(' ').str.get(0)
                    data['1_Shell'] = data['LENGTH_TRANS'].str.split(' ').str.get(1)
                    data['1_Shell'] = data['1_Shell'].replace("X", 'CS', regex=True)
                    data['empty1'] = data['empty1'].replace("X", '/CI', regex=True)
                    data['1_Shell'] = data['1_Shell'].replace("x", 'CS', regex=True)
                    data['empty1'] = data['empty1'].replace("x", '/CI', regex=True)
                    data['Shell'] = data['1_Shell'].astype(str) + data['empty1'].astype(str)
                    data['Shell'] = data['Shell'].replace("nan/", '', regex=True)
                    data['Shell'] = data['Shell'].replace("/nan", '', regex=True)
                    data['Shell'] = data['Shell'].replace("nan", '', regex=True)
                    data['Root'] = data['val1'].str.split(' ').str.get(0)
                    data['LE'] = data['val1'].str.split(' ').str.get(1)
                    data['Length'] = data['val1'].str.split(' ').str.get(1)
                    data['Trans'] = data['val1'].str.split(' ').str.get(1)
                    aa = df_Description['Description'].astype(str).to_list()
                    print('a')
                    print(aa)
                    bb = df_BREAK_REASON['BREAK REASON'].astype(str).to_list()
                    cc = df_WorkToDo['WorkToDo'].astype(str).to_list()
                    data['WorkToDo'] = data['WorkToDo'].astype(str)
                    data['ColumnA'] = data[data.columns[1:]].apply(
                        lambda x: ','.join(x.dropna().astype(str)),
                        axis=1
                    )


                    def matcherN_Description(x):
                        for i in aa:
                            if i.lower() in x.lower():
                                return i
                        else:
                            return np.nan


                    data['Description'] = data['ColumnA'].apply(matcherN_Description)


                    def matcherN_WorkToDo(x):
                        for i in bb:
                            if i.lower() in x.lower():
                                return i
                        else:
                            return np.nan


                    data['BREAK REASON'] = data['ColumnA'].apply(matcherN_WorkToDo)


                    def matcherN_BREAK_REASON(x):
                        for i in cc:
                            if i.lower() in x.lower():
                                return i
                        else:
                            return np.nan


                    data['WorkToDO'] = data['ColumnA'].apply(matcherN_BREAK_REASON)

                    data = data[
                        ['No', 'Code', 'Zone', 'Shell', 'Root', 'LE', 'Length', 'Trans', 'Description', 'WorkToDo',
                         'BREAK REASON']]
                    data['BladeNo'] = ''

                    df4.reset_index(drop=True, inplace=True)
                    data.reset_index(drop=True, inplace=True)
                    # data = data.reindex(['BladeNo','No', 'Code', 'Zone', 'Shell', 'Root', 'LE', 'Length', 'Trans','Description', 'WorkToDo','BREAK REASON'], axis=1)
                    data.astype(str)
                    print(data)
                    print(df4)
                    data['BladeNo'] = df4.loc[(data['No'] == '1').cumsum() - 1, 'No'].set_axis(data.index, axis=0)
                    data['Blade'] = df44.loc[(data['No'] == '1').cumsum() - 1, 'N2'].set_axis(data.index, axis=0)

                    data.to_excel(path + '/Data-2.xlsx', index=None, header=True)
                    # indx1 = [None] * 5
                    # count_blade.astype(str)
                    # for i in count_blade:

                    # Новий спосіб для блейдів
                    # indx1 = data[data['ZONE_SHELL_ROOT'].str[0:2]  == '1'].index.values
                    # index1 = indx1[:count_blade+1]
                    # print(index1)
                    # #print(indx1)
                    # #indx2 = data[data['No_CODE'].str[0:2]  == '01'].index.values
                    # #print(indx2)
                    # # for i in range(len(indx2)):
                    # #     if indx2[i] > 1:
                    # #         indx2[i] = indx2[i]-1
                    # print("\n")
                    #
                    #
                    # # creating column with BladeNo - search indexes of parts
                    # arr_num=[0]*(len(index1))
                    # arr_num[0]=0
                    # for i in range(len(index1)-1):
                    #       arr_num[i+1]=(index1[i+1])-index1[i]
                    #       print(arr_num)
                    #
                    # for i in range(len(arr_num)-1):
                    #       arr_num[i+1]=arr_num[i]+arr_num[i+1]
                    #       print("arr-num2")
                    #       print(arr_num)

                    # creating column with BladeNo - filling data
                    print("\n")

                    # for i in range(len(damInfo)):
                    #       for j in range(len(arr_num)-1):
                    #             if i>=arr_num[j] and i<arr_num[j+1]:
                    #                   damInfo.loc[i,'BladeNo'] = blades[j]
                    # WT
                    # indx1 = data[data['ZONE_SHELL_ROOT'].str[0:2]  == 'WT'].index.values
                    # index1 = indx1[:count_blade+1]
                    # print(index1)
                    # #print(indx1)
                    # #indx2 = data[data['No_CODE'].str[0:2]  == '01'].index.values
                    # #print(indx2)
                    # # for i in range(len(indx2)):
                    # #     if indx2[i] > 1:
                    # #         indx2[i] = indx2[i]-1
                    # print("\n")
                    #
                    #
                    # # creating column with BladeNo - search indexes of parts
                    # arr_num=[0]*(len(index1))
                    # arr_num[0]=0
                    # for i in range(len(index1)-1):
                    #       arr_num[i+1]=(index1[i+1])-index1[i]
                    #       print(arr_num)
                    #
                    # for i in range(len(arr_num)-1):
                    #       arr_num[i+1]=arr_num[i]+arr_num[i+1]
                    #       print("arr-num2")
                    #       print(arr_num)
                    #
                    #
                    # # creating column with BladeNo - filling data
                    # print("\n")
                    #
                    # damInfo.index = pd.RangeIndex(len(damInfo.index))
                    # for i in range(len(damInfo)):
                    #       for j in range(len(arr_num)-1):
                    #             if i>=arr_num[j] and i<arr_num[j+1]:
                    #                   damInfo.loc[i,'BladeNo'] = blades[j]
                    # WT
                    print("daminfo1")
                    # adding columns with general data
                    WIND_FARM = info.loc[
                        ((info['ColName'] == "WIND FARM:") | (info['ColName'] == "PARQUE:")).idxmax(), 'Value']
                    FL = info.loc[((info['ColName'] == "FL") | (info['ColName'] == "UT")).idxmax(), 'Value']
                    WTG = info.loc[
                        ((info['ColName'] == "WTG:") | (info['ColName'] == "AEROGENERADOR:")).idxmax(), 'Value']
                    CLIENT = info.loc[
                        ((info['ColName'] == "CLIENT:") | (info['ColName'] == "CLIENTE:")).idxmax(), 'Value']
                    WO_No = info.loc[((info['ColName'] == "WO No") | (info['ColName'] == "No OT")).idxmax(), 'Value']
                    SC = info.loc[(info['ColName'] == "SC").idxmax(), 'Value']
                    SC_RESPONSIBLE = info.loc[((info['ColName'] == "SC RESPONSIBLE") | (
                                info['ColName'] == "RESPONSABLE SC")).idxmax(), 'Value']
                    SUPERVISOR = info.loc[(info['ColName'] == "SUPERVISOR").idxmax(), 'Value']
                    WTG_MODEL = info.loc[
                        ((info['ColName'] == "WTG MODEL") | (info['ColName'] == "MODELO MAQUINA")).idxmax(), 'Value']
                    COUNTRY = info.loc[((info['ColName'] == "COUNTRY") | (info['ColName'] == "PAIS")).idxmax(), 'Value']
                    PRODUCTION_kWh = info.loc[((info['ColName'] == "PRODUCTION (kWh)") | (
                                info['ColName'] == "PRODUCCION (kWh)")).idxmax(), 'Value']
                    LINE_HOURS_OK = info.loc[((info['ColName'] == "LINE HOURS OK") | (
                                info['ColName'] == "HORAS LINEA OK")).idxmax(), 'Value']
                    WTG_HOURS_OK = info.loc[((info['ColName'] == "WTG HOURS OK") | (
                                info['ColName'] == "HORAS TURBINA OK")).idxmax(), 'Value']

                    StartDate = df_dates.loc[(df_dates['ColName'] == "START DATE").idxmax(), 'Data']
                    EndDate = df_dates.loc[(df_dates['ColName'] == "END DATE").idxmax(), 'Data']

                    data['WIND_FARM'] = WIND_FARM
                    data['FL'] = FL
                    data['WTG'] = WTG
                    data['CLIENT'] = CLIENT
                    data['WO_No'] = WO_No
                    data['SC'] = SC
                    data['SC_RESPONSIBLE'] = SC_RESPONSIBLE
                    data['SUPERVISOR'] = SUPERVISOR
                    data['WTG_MODEL'] = WTG_MODEL
                    data['COUNTRY'] = COUNTRY
                    data['PRODUCTION_kWh'] = PRODUCTION_kWh
                    data['LINE_HOURS_OK'] = LINE_HOURS_OK
                    data['WTG_HOURS_OK'] = WTG_HOURS_OK

                    data['StartDate'] = StartDate
                    data['EndDate'] = EndDate
                    data['ImgPath'] = ''
                    # spletting "Info" Column into properties
                    # damInfo[['No', 'Code', 'Zone', 'Shell', 'Root', 'LE', 'Length', 'Trans']] = damInfo['Info'].str.split(" ",n=8, expand=True)
                    # damInfo['Info']=damInfo['Info'].str.replace("  "," ",regex=True)
                    # damInfo['No'] = damInfo['Info'].str.split(' ').str.get(0)
                    # damInfo['Code'] = damInfo['Info'].str.split(' ').str.get(1)
                    # damInfo['Zone'] = damInfo['Info'].str.split(' ').str.get(2)
                    # damInfo['Shell'] = damInfo['Info'].str.split(' ').str.get(3)
                    # damInfo['Root'] = damInfo['Info'].str.split(' ').str.get(4)
                    # damInfo['LE'] = damInfo['Info'].str.split(' ').str.get(5)
                    # damInfo['Length'] = damInfo['Info'].str.split(' ').str.get(6)
                    # damInfo['Trans'] = damInfo['Info'].str.split(' ').str.get(7)
                    # damInfo['ROTURA'] = damInfo['Info'].str.split(' ').str.get(8)
                    # damInfo['ROTURA2'] = damInfo['Info'].str.split(' ').str.get(9)
                    damInfo = data[
                        ['WIND_FARM', 'FL', 'WTG', 'CLIENT', 'WO_No', 'SC', 'SC_RESPONSIBLE', 'SUPERVISOR', 'WTG_MODEL',
                         'COUNTRY', 'PRODUCTION_kWh', 'LINE_HOURS_OK', 'WTG_HOURS_OK', 'StartDate', 'EndDate','Blade',
                         'BladeNo', 'ImgPath', 'No', 'Code', 'Zone', 'Shell', 'Root', 'LE', 'Length', 'Trans',
                         'Description', 'WorkToDo', 'BREAK REASON']]
                    damInfo = damInfo.loc[damInfo['Code'] != '0.00']
                    damInfo.reset_index(drop=True, inplace=True)
                    damInfo['PRODUCTION_kWh'] = damInfo['PRODUCTION_kWh'].replace(np.nan, '0', regex=True)
                    damInfo['LINE_HOURS_OK'] = damInfo['LINE_HOURS_OK'].replace(np.nan, '0', regex=True)
                    damInfo['WTG_HOURS_OK'] = damInfo['WTG_HOURS_OK'].replace(np.nan, '0', regex=True)
                    damInfo['COUNTRY'] = damInfo['COUNTRY'].replace(np.nan, '0', regex=True)
                    damInfo['WTG_MODEL'] = damInfo['WTG_MODEL'].replace(np.nan, '0', regex=True)
                    damInfo['SUPERVISOR'] = damInfo['SUPERVISOR'].replace(np.nan, '0', regex=True)
                    damInfo['SC_RESPONSIBLE'] = damInfo['SC_RESPONSIBLE'].replace(np.nan, '0', regex=True)
                    print(damInfo)
                    # """

                    #  --------- read and save IMAGES
                    num = 1
                    idx = 0
                    doc = fitz.open(path_to_save)
                    for i in range(len(doc)):
                        for img in doc.get_page_images(i):
                            xref = img[0]
                            pix = fitz.Pixmap(doc, xref)
                            if (pix.n < 5) & (i > 3):  # this is GRAY or RGB
                                # if ( i%2==0 ):
                                pix.save(foldername_to_save_foto + "/p%s-%s.jpg" % (
                                i, xref))  # -- save all images into folder
                                im = cv2.imread(foldername_to_save_foto + "/p" + str(i) + "-" + str(xref) + ".jpg")
                                img = Image.open(foldername_to_save_foto + "/p" + str(i) + "-" + str(xref) + ".jpg")
                                height = img.height
                                if ((img.height) > 300) and (
                                        (img.height) < 1250):  # -- save final images - only with height larger then 70
                                    pix.save(path_to_save2 + "/p%s-%s.jpg" % (i, num))
                                    if (num % 2 != 0):  # to write path for the first image
                                        damInfo.loc[idx]['ImgPath'] = path_to_save2 + "/p%s-%s.jpg" % (i, num)
                                        idx = idx + 1
                                        num = num + 1
                                    else:  # to write path for the second image
                                        damInfo.loc[idx - 1]['ImgPath'] = damInfo.loc[idx - 1][
                                                                              'ImgPath'] + ", " + path_to_save2 + "/p%s-%s.jpg" % (
                                                                          i, num)
                                        num = num + 1
                            else:  # CMYK: convert to RGB first
                                pix1 = fitz.Pixmap(fitz.csRGB, pix)
                                # pix1.save(foldername_to_save_foto+"/p%s-%s.png" % (i, xref))
                                pix1 = None
                            pix = None

                    # creating final file

                    print(damInfo)
                    # num_rows11 = damInfo.shape[0]
                    # a = int(num_rows11/2)
                    # a = int(num_rows11)
                    #
                    # # a.astype(str)
                    # damInfo22 = damInfo.tail(a)
                    # damInfo = damInfo[:arr_num[count_blade]]
                    print("\nInfo about damages:")
                    print(damInfo)
                    # damInfo = damInfo[:index1[count_blade]]
                    # damInfo.to_excel(path + '/DamagesInfo.xlsx', index=None,
                    #                  header=True)  # creating file with necessary data only
                    # df111 = damInfo22.drop(['BladeNo','ImgPath'],axis = 1)
                    # damInfo = pd.DataFrame()
                    # damInfo['ImgPath'] = ''
                    # damInfo['ImgPath2'] = damInfo['ImgPath']
                    # damInfo22['BladeNo2'] = damInfo['BladeNo']
                    # df4 = pd.DataFrame(damInfo[['BladeNo','ImgPath']])
                    # df4.reset_index(drop=True, inplace=True)
                    # df5 = df4.rename(columns = {'BladeNo' : 'BladeNo1', 'ImgPath' : 'ImgPath1'}, inplace = True)
                    # damInfo.reset_index(drop=True, inplace=True)

                    # damInfo22['qw'] = df5['ImgPath1']
                    # df1111 = damInfo22.join(df5)

                    # df1111 = damInfo22.join(df5 , left_index=True, columns=['column_new_1','column_new_3'])
                    # result = pd.concat([df111, df], axis=1)
                    # df111.to_excel(path+'/DamagesInfo11.xlsx',index = None, header=True)
                    damInfo['ImgPath'] = damInfo['ImgPath'].astype(str)

                    # damInfo['ImgPath2'] = damInfo['ImgPath2'].astype(str)
                    # a = "D:/Covizmo2/camesa/Script_Gamesa/"
                    # writer = pd.ExcelWriter(path+"result_script_check.xlsx", engine="xlsxwriter")
                    damInfo['path2'] = damInfo['ImgPath'].str.split(',').str[-1]
                    damInfo['path1'] = damInfo['ImgPath'].str.split(',').str[0]
                    damInfo['number1'] = damInfo['path1'].str.split('/').str[-1]
                    damInfo['number2'] = damInfo['path2'].str.split('/').str[-1]
                    # print(damInfo22['Оренда'])
                    damInfo["img_path1"] = None
                    damInfo["img_path2"] = None
                    damInfo["img_path1"] = (
                        '=LEFT(CELL("filename"),FIND("[",CELL("filename"))-1)&AC:AC'
                    )
                    damInfo["img_path2"] = (
                        '=LEFT(CELL("filename"),FIND("[",CELL("filename"))-1)&AD:AD'
                    )
                    damInfo = damInfo.drop(['path2','path1','ImgPath'], axis=1, inplace=False)
                    # damInfo22['asa'] = damInfo22['product'].copy()
                    # damInfo22.to_excel(writer, index=False)

                    # writer.save()
                    print(damInfo)
                    # os.remove(foldername +"/test.csv")
                    damInfo.to_excel(path_to_save2+'/' +file2+ '.xlsx', index=None, header=True)
            except Exception:
                os.rename(path_to_save2, path_to_save2 + "- YOU NEED TO MAKE IT MANUALLY")
                pass

            # try:
            #     pdf_file = PyPDF2.PdfFileReader(path_to_save)
            #     print('pdf', pdf_file)
            #     df = tabula.read_pdf(path_to_save, pages='all')[0]
            #
            #     # output = subprocess.check_output("ping -c 2 -W 2 1.1.1.1; exit 0", stderr=subprocess.STDOUT, shell=True)
            #     # convert PDF into CSV
            #     tabula.convert_into(path_to_save, path_to_save2 + "/test.csv", output_format="csv",
            #                         pages='all')  # -- save csv file
            # except Exception:
            #     pass
            # pdf_file = PyPDF2.PdfFileReader(name)
            # print('pdf', pdf_file)
            # file2 = name[:-4]
            # path2 = os.path.join(foldername_to_save + "/" + file2)
            # print(path2)
            # print(file2)
            # os.makedirs(path2)