import os
from time import sleep
import win32com.client
import threading
from pythoncom import CoInitializeEx as pythoncomCoInitializeEx
from PyQt5 import QtCore, QtWidgets
import sys
import traceback

from okno_ui import Ui_Form
from  vxv_tnnc_SQL_Pyton import Sql
from version import ver
os.system('CLS')

app = QtWidgets.QApplication(sys.argv)
Form = QtWidgets.QWidget()
ui = Ui_Form()
ui.setupUi(Form)
Form.show()

_translate = QtCore.QCoreApplication.translate
Title = 'CollectorExcel v. 1.0' + str(ver)
Form.setWindowTitle(_translate("Form", Title))

'''Чистим "plainTextEdit" для отображения текста по умолчанию'''
ui.plainTextEdit.clear()
'''Обертка функции в потопк (декоратор)'''
def thread(my_func):
    def wrapper():
        threading.Thread(target=my_func, daemon=True).start()
    return wrapper

def colorBar(progBar, color):
    # progBar.setStyleSheet("QProgressBar::chunk {background-color: rgb(170, 170, 170); margin: 2px;}")
    progBar.setStyleSheet("QProgressBar::chunk {background-color: rgb("f"{color[0]}, {color[1]}, {color[2]}); margin: 2px;""}")

def Book():
    pythoncomCoInitializeEx(0)
    Excel = win32com.client.Dispatch("Excel.Application")
    # Excel.Visible = 0
    wb = Excel.ActiveWorkbook   # Получаем доступ к активной книге
    return wb

def importdataCode(sheet, StartRow, StartColl, EndRow, EndColl):
    '''Собираем список из 1ой колонки'''
    vals = sheet.Range(sheet.Cells(StartRow, StartColl), sheet.Cells(EndRow, EndColl)).value
    vals = [vals[i][x] for i in range(len(vals)) for x in range(len(vals[i]))]
    return vals

def loadListPoisk(fail):
    '''Открываем файл с кодировкой для чтения'''
    f = open(fail, 'r', encoding='utf-8')
    '''Читаем весь файл целиком как текст'''
    # text = f.read()
    '''Читаем файл и разбивает строку на подстроки в зависимости от разделителя'''
    text = f.read().split("\n") 
    return text

# @thread
def GO():
    ListPoiskSistem = loadListPoisk("Sistema.ini")
    ListPoisk = loadListPoisk("TypeMTR.ini")
    # directory = r"C:\vxvproj\tnnc-Excel\collectorExcel\collectorApp\temp"
    directory = str(ui.plainTextEdit.toPlainText())[8:]
    print(f"directory = {directory}")

    sig.signal_label.emit("Закрытие файлов" + " .." * 1)
    SH_Excel = win32com.client.Dispatch("Excel.Application")
    SH_wb = SH_Excel.Workbooks.Open(os.getcwd() + "\Сборная ведомость.xltx")
    SH_Excel.Visible = 1
    SH_sheet = SH_wb.Worksheets("Свод")
    SH_StartRow = 3
    SH_StartRow_1 = SH_StartRow
    SH_StartColl = 5
    SH_EndColl = 14
    SH_EndRow = SH_sheet.UsedRange.Rows.Count + SH_StartRow

    '''Удаляем строки со сдвигом вверх'''
    SH_sheet.Rows(f"{SH_StartRow}:{SH_EndRow}").Delete(1)

    sig.signal_label.emit("Закрытие файлов" + " .." * 2)
    Excel = win32com.client.Dispatch("Excel.Application")
    wbList = []
    direct = os.listdir(directory)
    countfailList = []
    
    for filename in direct:
        fff = os.path.join(directory, filename)
        if os.path.isfile(fff) and ".xls" in filename:
            countfailList.append(1)
    countfail = len(countfailList)
    sig.signal_label.emit("Закрытие файлов" + " .." * 3)
    
    for filename in direct:
        fff = os.path.join(directory, filename)
        if os.path.isfile(fff) and ".xls" in filename:
            textLable = filename
            
            nomerfail = direct.index(filename) + 1
            sig.signal_label.emit(f"Обработка файла {nomerfail} / {countfail}: {textLable}")
            
            wb = Excel.Workbooks.Open(fff)
            wb.Activate()
            wbList.append(wb)
            sheet = wb.Worksheets("Ввод")
            
            StartRow = 2
            '''от нижнего края вверх до нижней крайней заполненной ячейки'''
            EndRow = sheet.Cells(sheet.Rows.Count, 3).End(3).Row
            countRow = EndRow - StartRow + 1
            StartColl = 1
            EndColl = 10
            
            celscopy = sheet.Range(sheet.Cells(StartRow, StartColl), sheet.Cells(EndRow, EndColl))
            celscopy.Copy()
            
            SH_wb.Activate()
            SH_EndRow = SH_StartRow + countRow - 1
            SH_sheet.Range(SH_sheet.Cells(SH_StartRow, SH_StartColl), SH_sheet.Cells(SH_EndRow, SH_EndColl)).Activate()
            SH_sheet.Paste()

            sheet = wb.Worksheets("Штамп")
            val = sheet.Range("H2").value
            # val = sheet.Range("H2").Formula
            nom = val.find("-Р-")
            # nom = val.index("-Р-")
            shifr = val[:nom]

            SH_sheet.Range(SH_sheet.Cells(SH_StartRow, 1), SH_sheet.Cells(SH_EndRow, 2)).value = [shifr, val] * countRow

            cel = importdataCode(SH_sheet, SH_StartRow, 7, SH_EndRow, 7)
            sistColl = []
            searchList = []
            defaultval = None
            for i in cel:
                if i != None:
                    for g in ListPoiskSistem:
                        if g in i:
                            defaultval = g
                        else:
                            pass
                    sistColl.append((defaultval, ))

                    xxx = None
                    for g in ListPoisk:
                        if g in i:
                            xxx = g
                        else:
                            pass
                    searchList.append((xxx, ))
                else:
                    sistColl.append((defaultval, ))
                    searchList.append((None, ))
            
            SH_sheet.Range(SH_sheet.Cells(SH_StartRow, 3), SH_sheet.Cells(SH_EndRow, 3)).value = sistColl
            SH_sheet.Range(SH_sheet.Cells(SH_StartRow, 4), SH_sheet.Cells(SH_EndRow, 4)).value = searchList

            SH_StartRow = SH_EndRow + 1

            # nomerfail = direct.index(filename) + 1
            proc = round(nomerfail / countfail * 100)
            sig.signal_Probar.emit(proc)
            # sig.signal_label.emit(f"Обработка файла {nomerfail} / {countfail}: {textLable}")

    SH_EndRow = SH_sheet.Cells(SH_sheet.Rows.Count, 7).End(3).Row
    cel = SH_sheet.Range(SH_sheet.Cells(1, 1), SH_sheet.Cells(SH_EndRow, SH_EndColl))
    cel.Borders.Weight = 2
    SH_sheet.Cells(SH_EndRow + 1, SH_EndColl).Activate()
    sig.signal_label.emit("Закрытие файлов")
    sleep(1)
    for i in wbList:
        try:
            i.Activate()
            # i.close
            # i.Close()
            i.Close(False)
            sig.signal_label.emit("Закрытие файлов" + " .." * wbList.index(i))
        except:
            pass
    sig.signal_bool.emit(False)
    # SH_Excel.Visible = 1
    # SH_wb.Save()
    # SH_Excel.Quit()

    print("-----------------------------------")

class Signals(QtCore.QObject):
    signal_Probar = QtCore.pyqtSignal(int)
    signal_label = QtCore.pyqtSignal(str)
    signal_err = QtCore.pyqtSignal(str)
    signal_bool = QtCore.pyqtSignal(bool)
    signal_color = QtCore.pyqtSignal(list)

    def __init__(self, parent=None):
        QtCore.QThread.__init__(self, parent)
        self.signal_Probar.connect(self.on_change_Probar,QtCore.Qt.QueuedConnection)
        self.signal_label.connect(self.on_change_label,QtCore.Qt.QueuedConnection)
        self.signal_err.connect(self.on_change_err,QtCore.Qt.QueuedConnection)
        self.signal_bool.connect(self.on_change_bool,QtCore.Qt.QueuedConnection)
        self.signal_color.connect(self.on_change_color,QtCore.Qt.QueuedConnection)

    '''Отправляем сигналы в элементы окна'''
    def on_change_Probar(self, s):
        ui.progressBar_1.setValue(s)
    def on_change_label(self, s):
        ui.label.setText(s)
    def on_change_err(self, s):
        QtWidgets.QMessageBox.information(Form, 'Excel не отвечает...', s)
    def on_change_color(self, s):
        colorBar(ui.progressBar_1, color = s)
    def on_change_bool(self, s):
        ui.pushButton.setDisabled(s)

sig = Signals()

@thread
def start():
    sig.signal_Probar.emit(0)
    try:
        Sql("CollectorExcel")
        sig.signal_bool.emit(True)
        sig.signal_Probar.emit(0)
        sig.signal_color.emit([100, 150, 150])
        GO()
    except Exception as e:
        # print(str(traceback.print_tb))
        errortext = traceback.format_exc()
        text = f"Сбор таблиц не выполнен, повторите попытку \n\n{errortext}"
        sig.signal_err.emit(text)
    sig.signal_bool.emit(False)
    sig.signal_color.emit([170, 170, 170])
    sig.signal_Probar.emit(100)
    sig.signal_label.emit("Выполнено")

def openSistema():
    '''открытие файла как при двойном клике'''
    os.startfile('Sistema.ini')

def openTypeMTR():
    '''открытие файла как при двойном клике'''
    os.startfile('TypeMTR.ini')

ui.pushButton.clicked.connect(start)
ui.pushButton_2.clicked.connect(openSistema)
ui.pushButton_3.clicked.connect(openTypeMTR)

# xxx = "Закрытие файлов" + " .." * 5
# print(f"xxx = {xxx}")
if __name__ == "__main__":
    # start()
    sys.exit(app.exec_())
