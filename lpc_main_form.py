import xlrd
import pandas as pd
pd.options.display.max_columns = 100
# import re
# import sys
from tkinter import *
from tkinter.ttk import *
import tkinter.messagebox as msgbox
import tkinter.filedialog as dialog
from datetime import datetime, timedelta

def xldate_to_datetime(xldate):
   tempDate = datetime(1900, 1, 1)
   deltaDays =timedelta(days=int(xldate)-2)
   TheTime = (tempDate + deltaDays )
   return TheTime.strftime("%d.%m.%Y")

# in Terminal launch:
# pyinstaller --onefile --windowed --distpath "C:\Geobank" --workpath "C:\Geobank" --hidden-import pyexcel_xls --hidden-import pyexcel_io --hidden-import pyexcel lpc_main_form.py
# pyinstaller --onefile --windowed --distpath "C:\Geobank" --workpath "C:\Geobank" lpc_main_form.py
class App:
    def __init__(self, root):
        """Создание интерфейса"""
        self.root = root
        self.root.title("СЕЛИГДАР: Форматировать файл получения")
        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

        self.create_variables()
        self.create_widgets()
        self.create_bindings()
        self.grid_widgets()

        self.root.update()

        # self.root.wm_iconbitmap("C:\\Program Files (x86)\\Micromine\\Geobank\\9\\Geobank.ico")

        wscreen = self.root.winfo_screenwidth()
        hscreen = self.root.winfo_screenheight()
        wroot, hroot = tuple(int(_) for _ in self.root.geometry().split('+')[0].split('x'))
        self.root.geometry("+%d+%d" % ((int(wscreen/2 - wroot/2), int(hscreen/2 - hroot/2))))
        self.root.focus_force()

    def create_bindings(self):
        self.root.bind("<Return>", lambda e: self.check())
        self.root.bind("<Escape>", lambda e: self.close_app())
        self.ent_inputPath.bind("<Double-Button-1>", lambda e: self.browse_open_file())
        # self.ent_outputPath.bind("<Double-Button-1>", lambda e: self.browse_save_file())

    def create_variables(self):
        self.var_laboratories = ['Рябиновое'] #['Алдан', 'Рябиновое']
        self.var_methods = ["НСАМ № 505-Х", "НСАМ № 130-с", "НСАМ № 392"]

    def create_widgets(self):
        self.lfr_input = LabelFrame(self.root, text = "Ввод")
        self.lbl_inputPath = Label(self.lfr_input, text = "Файл:")
        self.ent_inputPath = Entry(self.lfr_input, width = 80)
        self.ent_inputPath.insert("end", "D:\\Cloud\\Git\\micromine\\lpc\\лаборатория рябиновое\\Протокол Ag атомно-абсорбционный[AA Ag].xlsx")
        # D:\\Cloud\\Git\\micromine\\lpc\\лаборатория рябиновое\\Протокол №П-1130 Рябиновый РМЮ-42477-РМЮ-42542[Пробирный Au].xls   505
        # D:\\Cloud\\Git\\micromine\\lpc\\лаборатория рябиновое\\Протокол ЭРСА № 285 ТГ (21402-21450)[РСА Au].xls   392
        # D:\\Cloud\\Git\\micromine\\lpc\\лаборатория рябиновое\\Протокол Ag атомно-абсорбционный[AA Ag].xlsx    130-с

        self.but_inputBrowse = Button(self.lfr_input, text = "...", width = 3, command = self.browse_open_file)

        self.lbl_inputLab = Label(self.lfr_input, text = "Лаборатория:")
        self.cmb_inputLab = Combobox(self.lfr_input, width = 35, state = "readonly")
        self.cmb_inputLab["values"] = self.var_laboratories
        self.cmb_inputLab.current(0)

        self.lbl_inputMethod = Label(self.lfr_input, text = "Метод:")
        self.cmb_inputMethod = Combobox(self.lfr_input, width = 35, state = "readonly")
        self.cmb_inputMethod["values"] = self.var_methods
        self.cmb_inputMethod.current(0)

        self.but_ok = Button(self.root, text = "OK", width = 13, command = self.check)
        self.but_cancel = Button(self.root, text = "Отмена", width = 13, command = self.root.destroy)

    def grid_widgets(self):
        self.lfr_input.grid(row = 0, rowspan = 20, column = 0, padx = 5, pady = 3, sticky = "nsew")
        self.lbl_inputPath.grid(row = 0, column = 0, padx = 5, pady = 3, sticky = "e")
        self.ent_inputPath.grid(row = 0, column = 1, padx = 5, pady = 3, sticky = "ew")
        self.but_inputBrowse.grid(row = 0, column = 2, padx = 5, pady = 3, sticky = "ew")
        self.lbl_inputLab.grid(row = 1, column = 0, padx = 5, pady = 3, sticky = "e")
        self.cmb_inputLab.grid(row = 1, column = 1, padx = 5, pady = 3, sticky = "ew")

        self.lbl_inputMethod.grid(row=2, column=0, padx=5, pady=3, sticky="e")
        self.cmb_inputMethod.grid(row=2, column=1, padx=5, pady=3, sticky="ew")
        self.but_ok.grid(row = 7, column = 1, padx = 5, pady = 3, sticky = "ew")
        self.but_cancel.grid(row = 8, column = 1, padx = 5, pady = 3, sticky = "ew")

    def browse_open_file(self):
        path = dialog.askopenfilename(filetypes = [("Файлы Excel", "*.xls;*.xlsx")]).replace("/", "\\")
        if not path:	return
        self.ent_inputPath.delete(0, "end")
        self.ent_inputPath.insert("end", path)
        self.root.focus_force()

    def check(self):
        try:
            self.run()
        except Exception as e:
            msgbox.showerror("Ошибка", "Не удалось форматировать файл. " + str(e))
            return

    def run(self):
        SOURCE_FILE_PATH = self.ent_inputPath.get()
        LAB = self.cmb_inputLab.get()
        METHOD = self.cmb_inputMethod.get()

        RESULT_FILE_PATH = SOURCE_FILE_PATH.replace('.xlsx', '.csv')
        RESULT_FILE_PATH = RESULT_FILE_PATH.replace('.xls', '.csv')

        dict = parseExcelFile(SOURCE_FILE_PATH, LAB, METHOD)
        C = LabProtocol(dict)
        writeCSV(C, RESULT_FILE_PATH, ';')
        sys.exit()



    def close_app(self):
        self.root.destroy()

class LabProtocol:
    dict = {}
    def __init__(self,dict1):
        self.dict = dict1

def writeCSV(lp : LabProtocol,path_,sep_):
    f = open(path_,'w',encoding='utf-8')
    #если значение ключа в виде списка, то применяем индекс. Регулярные выражения подразумевают список на выходе. Если их не используем -- то убираем индексы
    t = 'Номер работы:' + sep_ + str(lp.dict['lab_job_no']) + '\n'
    f.write(t)
    t = 'Номер отправки:' + sep_ + str(lp.dict['despatch_id']) + '\n'
    f.write(t)
    t = 'Дата получения:' + sep_ + str(lp.dict['receipt_date']) + '\n'
    f.write(t)
    t = 'Метод:' + sep_ + str(lp.dict['lab_method'][0]) + '\n'
    f.write(t)
    t = 'Элемент:' + sep_ + str(lp.dict['lab_element']) + '\n'
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
    if (lab_id_ == 'Рябиновое') & (lab_method_ == 'НСАМ № 505-Х'): #лаборатория "Рябиновое", пробирный анализ
        FR = 19
        c_lab_tag = 1  # далее 6
        r_lab_job_no = 7  # 6
        r_lab_date = 7  # 6
        r_despatch_date = 10
        r_despatch = 10  # 9
        r_element = 15  # 14

        c_lab_job_no = 7  # 0
        c_despatch_date = 4
        c_lab_date = 9  # 8
        c_despatch = 2
        c_element = 3

        sheetName = 'Лист2'
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
            dtype={c_lab_tag: object, (c_lab_tag + 5): object, (c_element): object, (c_element + 5): object},
            skiprows=range(0, FR),  # list of rows you want to omit at the beginning
        )
        v_despatch = sheet.cell_value(r_despatch, c_despatch)
        if v_despatch.is_integer():  # на самом деле он тут float
            v_despatch = str(int(v_despatch))
        v_lab_job_no = sheet.cell_value(r_lab_job_no, c_lab_job_no)
        if v_lab_job_no.is_integer():  # на самом деле он тут float
            v_lab_job_no = str(int(v_lab_job_no))
        dict = {
            #'despatch_id': re.findall('[\w\-\d]+', v_despatch),  # int(sheet.cell_value(r_despatch,c_despatch)), #
            'despatch_id': v_despatch,

            'lab_id': lab_id_,
            #'lab_job_no': re.findall('[\w\-\d]+', v_lab_job_no),
            'lab_job_no': v_lab_job_no,

            #'receipt_date': re.findall('[\d\.]+', sheet.cell_value(r_despatch_date, c_despatch_date)),
            'receipt_date': xldate_to_datetime(sheet.cell_value(r_despatch_date, c_despatch_date)),

            #'lab_date': re.findall('[\d\.]+', sheet.cell_value(r_lab_date, c_lab_date)),
            'lab_date': xldate_to_datetime(sheet.cell_value(r_lab_date, c_lab_date)),

            'lab_method': lab_method,

            #'lab_element': re.findall('Ag|Au', sheet.cell_value(r_element, c_element)),
            'lab_element': 'Au',

            'lab_units': lab_units
            # 'result_range': df_result_range
        }
        ELEMENT = dict['lab_element'][0]
        # опираясь на формат протокола, отберем нужные столбцы, откинув все лишнее
        df = df_result_range.copy().drop([0, 4, 5, 9], axis=1)

        # сделаем конкатенацию в единый DF
        df1 = df[df.columns[0:3]]
        df1.columns = ['lab_tag', 'sample_tag', ELEMENT]
        df2 = df[df.columns[3:6]]
        df2.columns = ['lab_tag', 'sample_tag', ELEMENT]
        df3 = pd.concat([df1, df2], axis=0, ignore_index=True)

        # обработаем пропуски и НПО
        df3.loc[df3[ELEMENT] == 'НПО', ELEMENT] = 'LDL'
        df3.loc[df3[ELEMENT].isnull(), ELEMENT] = 'NA'

        # откинем промежуточную строку-шапку, где в заголовке столбца lab_tag стоит 2 или 7
        # откинем все пустые по полю lab_tag
        dict['result_range_selected'] = df3.query('(lab_tag != [2,7]) & (lab_tag == lab_tag)').reset_index(drop=True)
    elif (lab_id_ == 'Рябиновое') & (lab_method_ == 'НСАМ № 392'): #лаборатория "Рябиновое", экстракционный рентгено-спектральный анализ
        FR = 20
        c_lab_tag = 1  # далее 6
        r_lab_job_no = 7  # 6
        r_lab_date = 7  # 6
        r_despatch_date = 10
        r_despatch = 10  # 9
        r_element = 15  # 14

        c_lab_job_no = 7  # 0
        c_despatch_date = 3
        c_lab_date = 9  # 8
        c_despatch = 2
        c_element = 3

        sheetName = 'Лист2'
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
            dtype={c_lab_tag: object, (c_lab_tag + 5): object, (c_element): object, (c_element + 5): object},
            skiprows=range(0, FR),  # list of rows you want to omit at the beginning
        )
        v_despatch = sheet.cell_value(r_despatch, c_despatch)
        if v_despatch.is_integer():  # на самом деле он тут float
            v_despatch = str(int(v_despatch))
        v_lab_job_no = sheet.cell_value(r_lab_job_no, c_lab_job_no)
        if v_lab_job_no.is_integer():  # на самом деле он тут float
            v_lab_job_no = str(int(v_lab_job_no))
        dict = {
            #'despatch_id': re.findall('[\w\-\d]+', v_despatch),  # int(sheet.cell_value(r_despatch,c_despatch)), #
            'despatch_id': v_despatch,

            'lab_id': lab_id_,
            #'lab_job_no': re.findall('[\w\-\d]+', v_lab_job_no),
            'lab_job_no': v_lab_job_no,

            #'receipt_date': re.findall('[\d\.]+', sheet.cell_value(r_despatch_date, c_despatch_date)),
            'receipt_date': xldate_to_datetime(sheet.cell_value(r_despatch_date, c_despatch_date)),

            #'lab_date': re.findall('[\d\.]+', sheet.cell_value(r_lab_date, c_lab_date)),
            'lab_date': xldate_to_datetime(sheet.cell_value(r_lab_date, c_lab_date)),

            'lab_method': lab_method,

            #'lab_element': re.findall('Ag|Au', sheet.cell_value(r_element, c_element)),
            'lab_element': 'Au',

            'lab_units': lab_units
            # 'result_range': df_result_range
        }
        ELEMENT = dict['lab_element'][0]
        # опираясь на формат протокола, отберем нужные столбцы, откинув все лишнее
        df = df_result_range.copy().drop([0, 4, 5, 9], axis=1)

        # сделаем конкатенацию в единый DF
        df1 = df[df.columns[0:3]]
        df1.columns = ['lab_tag', 'sample_tag', ELEMENT]
        df2 = df[df.columns[3:6]]
        df2.columns = ['lab_tag', 'sample_tag', ELEMENT]
        df3 = pd.concat([df1, df2], axis=0, ignore_index=True)

        # обработаем пропуски и НПО
        df3.loc[df3[ELEMENT] == 'НПО', ELEMENT] = 'LDL'
        df3.loc[df3[ELEMENT].isnull(), ELEMENT] = 'NA'

        # откинем промежуточную строку-шапку, где в заголовке столбца lab_tag стоит 2 или 7
        # откинем все пустые по полю lab_tag
        dict['result_range_selected'] = df3.query('(lab_tag != [2,7]) & (lab_tag == lab_tag)').reset_index(drop=True)
    elif (lab_id_ == 'Рябиновое') & (lab_method_ == 'НСАМ № 130-с'):  # лаборатория "Рябиновое", атомно-абсорбционный анализ
        FR = 19
        c_lab_tag = 1  # далее 6
        r_lab_job_no = 7  # 6
        r_lab_date = 7  # 6
        r_despatch_date = 7
        r_despatch = 10  # 9
        r_element = 15  # 14

        c_lab_job_no = 5  # 0
        c_despatch_date = 6
        c_lab_date = 6  # 8
        c_despatch = 1
        c_element = 2

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
            dtype={c_lab_tag: object, (c_lab_tag + 5): object, (c_element): object, (c_element + 5): object},
            skiprows=range(0, FR),  # list of rows you want to omit at the beginning
        )
        v_despatch = sheet.cell_value(r_despatch, c_despatch)
        v_receipt_date = sheet.cell_value(r_despatch_date, c_despatch_date)
        v_receipt_date = v_receipt_date.replace(' ','')
        v_receipt_date = v_receipt_date.replace('от', '')
        try:
            if v_despatch.is_integer():  # на самом деле он тут float
                v_despatch = str(int(v_despatch))
        except:
            None

        v_lab_job_no = sheet.cell_value(r_lab_job_no, c_lab_job_no)

        try:
            if v_lab_job_no.is_integer():  # на самом деле он тут float
                v_lab_job_no = str(int(v_lab_job_no))
        except:
            None

        dict = {
            # 'despatch_id': re.findall('[\w\-\d]+', v_despatch),  # int(sheet.cell_value(r_despatch,c_despatch)), #
            'despatch_id': v_despatch,

            'lab_id': lab_id_,
            # 'lab_job_no': re.findall('[\w\-\d]+', v_lab_job_no),
            'lab_job_no': v_lab_job_no,

            # 'receipt_date': re.findall('[\d\.]+', sheet.cell_value(r_despatch_date, c_despatch_date)),
            'receipt_date': v_receipt_date, #xldate_to_datetime(sheet.cell_value(r_despatch_date, c_despatch_date)),

            # 'lab_date': re.findall('[\d\.]+', sheet.cell_value(r_lab_date, c_lab_date)),
            'lab_date': v_receipt_date, #xldate_to_datetime(sheet.cell_value(r_lab_date, c_lab_date)),

            'lab_method': lab_method,

            # 'lab_element': re.findall('Ag|Au', sheet.cell_value(r_element, c_element)),
            'lab_element': 'Ag',

            'lab_units': lab_units
            # 'result_range': df_result_range
        }
        ELEMENT = dict['lab_element']
        # опираясь на формат протокола, отберем нужные столбцы, откинув все лишнее
        df = df_result_range.copy().drop([3, 7], axis=1)

        # сделаем конкатенацию в единый DF
        df1 = df[df.columns[0:3]]
        df1.columns = ['lab_tag', 'sample_tag', ELEMENT]
        df2 = df[df.columns[3:6]]
        df2.columns = ['lab_tag', 'sample_tag', ELEMENT]
        df3 = pd.concat([df1, df2], axis=0, ignore_index=True)

        # обработаем пропуски и НПО
        df3.loc[df3[ELEMENT] == 'НПО', ELEMENT] = 'LDL'
        df3.loc[df3[ELEMENT].isnull(), ELEMENT] = 'NA'

        # откинем промежуточную строку-шапку, где в заголовке столбца sample_tag стоит 2 или 6
        # откинем все пустые (NA) по полю Ag
        dict['result_range_selected'] = df3.query('(sample_tag != [2,6]) & (Ag != "NA")').reset_index(drop=True)

    return dict

if __name__ == '__main__':
    root = Tk()
    try:
        app = App(root)
        root.mainloop()
    except Exception as e:
        msgbox.showerror("Ошибка", str(e))
        root.destroy()
        