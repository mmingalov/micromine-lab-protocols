import xlrd
import pandas as pd
import re
# import sys
pd.options.display.max_columns = 100
from tkinter import *
# from tkinter import Menu, messagebox,simpledialog,ttk


class LabProtocol:
    dict = {}
    def __init__(self,dict1):
        self.dict = dict1

def writeCSV(lp : LabProtocol,path_,sep_):
    f = open(path_,'w',encoding='utf-8')
    # v=lp.dict['lab_job_no'][0]
    t = 'Номер работы:' + sep_ + str(lp.dict['lab_job_no'][0]) + '\n'
    f.write(t)
    t = 'Номер отправки:' + sep_ + str(lp.dict['despatch_id'][0]) + '\n'
    f.write(t)
    t = 'Дата получения:' + sep_ + str(lp.dict['receipt_date'][0]) + '\n'
    f.write(t)
    t = 'Метод:' + sep_ + str(lp.dict['lab_method'][0]) + '\n'
    f.write(t)
    t = 'Элемент:' + sep_ + str(lp.dict['lab_element'][0]) + '\n'
    f.write(t)
    t = 'Единицы измерения:' + sep_ + str(lp.dict['lab_units'][0]) + '\n'
    f.write(t)
    t = 'Нижний предел:' + sep_ + sep_ + '\n'
    f.write(t)
    t = 'Верхний предел:' + sep_ +sep_ + '\n'
    f.write(t)

    # for i in lp.dict['result_range_selected']:
    #     t = i['sample_tag'] + ',' + str(i['result']) + '\n'
    #     f.write(t)
    f.close()

    # b = lp.dict['result_range_selected'].copy().drop(['lab_tag'],axis=1)
    b = lp.dict['result_range_selected'].copy()
    # b = b[b.columns['sample_tag','']]
    b.to_csv(path_, sep=sep_, mode = 'a',header=False,index=False)

def parseExcelFile(path1, lab_id_, lab_method_):
    # every LAB_ID uses its parsing scheme cause has different protocol structure
    if (lab_id_ == 'Алдан') & (lab_method_ == 'НСАМ № 505-Х'): #лаборатория "Алдан", пробирный анализ
        FR = 18
        c_lab_tag = 1       #далее 6
        r_lab_job_no = 6
        c_lab_job_no = 0
        r_lab_date = 6
        c_lab_date = 8
        r_despatch = 9
        c_despatch = 2
        r_element = 14
        c_element = 3
        sheetName = 'Лист2'
        lab_method=[]
        lab_method.append(lab_method_)
        lab_units=[]
        lab_units.append('G/T')

        # book = p.get_book(file_name=path1)
        book = xlrd.open_workbook(filename=path1)
        sheet = book.sheet_by_name(sheetName)

        df_result_range = pd.read_excel(
                    path1,  # name of excel sheet
                    sheet_name=sheetName,
                    header=None,
                    index_col=None,
                    dtype= {c_lab_tag: object, (c_lab_tag+5): object,(c_element): object,(c_element+5): object},
                    skiprows=range(0, FR),  # list of rows you want to omit at the beginning
                )
        dict = {
            'despatch_id': re.findall('\w\-\w+',sheet.cell_value(r_despatch,c_despatch)),
            'lab_id': lab_id_,
            'lab_job_no': re.findall('\w\-\w+',sheet.cell_value(r_lab_job_no,c_lab_job_no)),
            'receipt_date': re.findall('[\d\.]+',sheet.cell_value(r_lab_date,c_lab_date)),
            'lab_date': re.findall('[\d\.]+',sheet.cell_value(r_lab_date,c_lab_date)),
            'lab_method':lab_method,
            'lab_element':re.findall('Ag|Au',sheet.cell_value(r_element,c_element)),
            'lab_units':lab_units
            # 'result_range': df_result_range
        }
        ELEMENT = dict['lab_element'][0]
        # опираясь на формат протокола, отберем нужные столбцы, откинув все лишнее
        df = df_result_range.copy().drop([0,4,5,9],axis=1)

        # сделаем конкатенацию в единый DF
        df1 = df[df.columns[0:3]]
        df1.columns = ['lab_tag', 'sample_tag', ELEMENT]
        df2 = df[df.columns[3:6]]
        df2.columns = ['lab_tag', 'sample_tag', ELEMENT]
        df3 = pd.concat([df1,df2], axis=0, ignore_index=True)

        # обработаем пропуски и НПО
        df3.loc[df3[ELEMENT] == 'НПО',ELEMENT] = 'LDL'
        df3.loc[df3[ELEMENT].isnull(), ELEMENT] = 'NA'

        # откинем промежуточную строку-шапку, где в заголовке столбца lab_tag стоит 2 или 7
        # откинем все пустые по полю lab_tag
        dict['result_range_selected'] = df3.query('(lab_tag != [2,7]) & (lab_tag == lab_tag)').reset_index(drop=True)
    elif (lab_id_ == 'Алдан') & (lab_method_ == 'НСАМ № 130-с'):  # лаборатория "Алдан", атомно-абсорбционный анализ
        FR = 18
        c_sample_tag = 1  # далее 6
        r_lab_job_no = 5
        c_lab_job_no = 1
        r_lab_date = 5
        c_lab_date = 7
        r_despatch = 8
        c_despatch = 2
        r_element = 14
        c_element = 3
        sheetName = 'Лист1'
        lab_method = []
        lab_method.append(lab_method_)
        lab_units = []
        lab_units.append('G/T')

        # book = p.get_book(file_name=path1)
        book = xlrd.open_workbook(filename=path1)
        sheet = book.sheet_by_name(sheetName)

        df_result_range = pd.read_excel(
            path1,  # name of excel sheet
            sheet_name=sheetName,
            header=None,
            index_col=None,
            dtype={c_sample_tag: object, (c_element): object},
            skiprows=range(0, FR),  # list of rows you want to omit at the beginning
        )
        dict = {
            'despatch_id': re.findall('[0-9,.]+', sheet.cell_value(r_despatch, c_despatch)),
            'lab_id': lab_id_,
            'lab_job_no': re.findall('\d+', sheet.cell_value(r_lab_job_no, c_lab_job_no)),
            'receipt_date': re.findall('[\d\.]+', sheet.cell_value(r_lab_date, c_lab_date)),
            'lab_date': re.findall('[\d\.]+', sheet.cell_value(r_lab_date, c_lab_date)),
            'lab_method': lab_method,
            'lab_element': re.findall('Ag|Au', sheet.cell_value(r_element, c_element)),
            'lab_units': lab_units,
            'result_range': df_result_range
        }
        ELEMENT = dict['lab_element'][0]
        # опираясь на формат протокола, отберем нужные столбцы, откинув все лишнее
        df = df_result_range[df_result_range.columns[0:9]].drop(columns=[0,2, 4, 5,7], axis = 1)
        # dict['result_range_selected'] = df

        # сделаем конкатенацию в единый DF
        df1 = df[df.columns[0:2]]
        df1.columns = ['sample_tag', ELEMENT]
        df2 = df[df.columns[2:5]]
        df2.columns = ['sample_tag', ELEMENT]
        df3 = pd.concat([df1, df2], axis=0, ignore_index=True)

        # обработаем НПО
        df3.loc[df3[ELEMENT] == 'НПО', ELEMENT] = 'LDL'
        # df3.loc[df3[ELEMENT].isnull(), ELEMENT] = 'NA'

        # откинем промежуточную строку-шапку, где в заголовке столбца sample_tag стоит 2 или 6
        # откинем все пустые по полю ELEMENT
        dict['result_range_selected'] = df3.query('(Ag == Ag) & (sample_tag != [2,6]) & (sample_tag == sample_tag)').reset_index(drop=True)

    elif (lab_id_ == 'Рябиновое') & (lab_method_ == 'НСАМ № 505-Х'): #лаборатория "Алдан", пробирный анализ
        FR = 19
        c_lab_tag = 1       #далее 6
        r_lab_job_no = 7#6
        r_lab_date = 7#6
        r_despatch_date = 10
        r_despatch = 10  # 9
        r_element = 15  # 14

        c_lab_job_no = 7#0
        c_despatch_date = 3
        c_lab_date = 9#8
        c_despatch = 2
        c_element = 3

        sheetName = 'Лист1'
        lab_method=[]
        lab_method.append(lab_method_)
        lab_units=[]
        lab_units.append('G/T')

        # book = p.get_book(file_name=path1)
        book = xlrd.open_workbook(filename=path1)
        sheet = book.sheet_by_name(sheetName)

        df_result_range = pd.read_excel(
                    path1,  # name of excel sheet
                    sheet_name=sheetName,
                    header=None,
                    index_col=None,
                    dtype= {c_lab_tag: object, (c_lab_tag+5): object,(c_element): object,(c_element+5): object},
                    skiprows=range(0, FR),  # list of rows you want to omit at the beginning
                )
        v_despatch = sheet.cell_value(r_despatch,c_despatch)
        if v_despatch.is_integer(): #на самом деле он тут float
             v_despatch = str(int(v_despatch))
        v_lab_job_no = sheet.cell_value(r_lab_job_no,c_lab_job_no)
        if v_lab_job_no.is_integer(): #на самом деле он тут float
             v_lab_job_no = str(int(v_lab_job_no))
        dict = {
            'despatch_id': re.findall('[\w\-\d]+',v_despatch), #int(sheet.cell_value(r_despatch,c_despatch)), #
            'lab_id': lab_id_,
            'lab_job_no': re.findall('[\w\-\d]+',v_lab_job_no), #re.findall('\w\-\w+',sheet.cell_value(r_lab_job_no,c_lab_job_no)),
            'receipt_date': re.findall('[\d\.]+',sheet.cell_value(r_despatch_date,c_despatch_date)),
            'lab_date': re.findall('[\d\.]+',sheet.cell_value(r_lab_date,c_lab_date)),
            'lab_method':lab_method,
            'lab_element':re.findall('Ag|Au',sheet.cell_value(r_element,c_element)),
            'lab_units':lab_units
            # 'result_range': df_result_range
        }
        ELEMENT = dict['lab_element'][0]
        # опираясь на формат протокола, отберем нужные столбцы, откинув все лишнее
        df = df_result_range.copy().drop([0,4,5,9],axis=1)

        # сделаем конкатенацию в единый DF
        df1 = df[df.columns[0:3]]
        df1.columns = ['lab_tag', 'sample_tag', ELEMENT]
        df2 = df[df.columns[3:6]]
        df2.columns = ['lab_tag', 'sample_tag', ELEMENT]
        df3 = pd.concat([df1,df2], axis=0, ignore_index=True)

        # обработаем пропуски и НПО
        df3.loc[df3[ELEMENT] == 'НПО',ELEMENT] = 'LDL'
        df3.loc[df3[ELEMENT].isnull(), ELEMENT] = 'NA'

        # откинем промежуточную строку-шапку, где в заголовке столбца lab_tag стоит 2 или 7
        # откинем все пустые по полю lab_tag
        dict['result_range_selected'] = df3.query('(lab_tag != [2,7]) & (lab_tag == lab_tag)').reset_index(drop=True)
    return dict

if __name__ == '__main__':
    # python lpc_main.py "D:\Cloud\Git\micromine\lpc\лаборатория алдан\Протокол Au пробирный.xls" "Алдан" "НСАМ № 505-Х"
    # python lpc_main.py "D:\Cloud\Git\micromine\lpc\лаборатория алдан\Протокол Ag атомно-абсорбционный.xlsx" "Алдан" "НСАМ № 130-с"
    print("Скрипт запущен с набором параметров:")
    print(sys.argv)

    # были переданы 3 параметра, как и требуется
    if len(sys.argv) == 4:
        SOURCE_FILE_PATH = sys.argv[1] # ".\lpc\лаборатория алдан\Протокол Au пробирный.xls" "D:\Cloud\Git\micromine\lpc\лаборатория алдан\Протокол Au пробирный.xls"
        LAB = sys.argv[2] #  "Алдан"
        METHOD = sys.argv[3] # "НСАМ № 505-Х"

        RESULT_FILE_PATH = SOURCE_FILE_PATH.replace('.xlsx', '.csv')
        RESULT_FILE_PATH = RESULT_FILE_PATH.replace('.xls', '.csv')

        dict = parseExcelFile(SOURCE_FILE_PATH, LAB, METHOD)
        C = LabProtocol(dict)
        writeCSV(C, RESULT_FILE_PATH, ';')
        sys.exit()
    else:
        # кол-во параметров не такое, как требуется
        print("Входные параметры некорректные. Будет выполнен демо-запуск.")

        SOURCE_FILE_PATH = ".\lpc\лаборатория алдан\Протокол Au пробирный.xls"
        SOURCE_FILE_PATH = ".\lpc\лаборатория рябиновое\Протокол Au пробирный_Рябиновое.xls"
        RESULT_FILE_PATH = SOURCE_FILE_PATH.replace('.xlsx','.csv')
        RESULT_FILE_PATH = RESULT_FILE_PATH.replace('.xls','.csv')
        LAB = 'Алдан'
        LAB = 'Рябиновое'
        METHOD = 'НСАМ № 505-Х'
        print()
        print('Тест:1, Лаборатория:', LAB, 'Метод:', METHOD)
        print('Файл:', SOURCE_FILE_PATH)
        dict = parseExcelFile(SOURCE_FILE_PATH,LAB,METHOD)
        print(dict['despatch_id'],dict['lab_job_no'],dict['lab_element'],dict['lab_method'],dict['lab_units'])

        C = LabProtocol(dict)
        writeCSV(C,RESULT_FILE_PATH,';')

        # ----------------------------------------------------------------------------------
        # входные параметры
        SOURCE_FILE_PATH = ".\lpc\лаборатория алдан\Протокол Ag атомно-абсорбционный.xlsx"
        RESULT_FILE_PATH = SOURCE_FILE_PATH.replace('.xlsx','.csv')
        RESULT_FILE_PATH = RESULT_FILE_PATH.replace('.xls','.csv')

        LAB = 'Алдан'
        METHOD = 'НСАМ № 130-с'
        print()
        print('Тест:2, Лаборатория:', LAB, 'Метод:', METHOD)
        print('Файл:', SOURCE_FILE_PATH)
        dict = parseExcelFile(SOURCE_FILE_PATH,LAB,METHOD)
        print(dict['despatch_id'],dict['lab_job_no'],dict['lab_element'],dict['lab_method'],dict['lab_units'])

        C = LabProtocol(dict)
        writeCSV(C,RESULT_FILE_PATH,';')

        # print (sys.path)
