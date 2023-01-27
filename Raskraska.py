# import os
from this import d
from time import sleep
import win32com.client
from collections import Counter

# from VBAExcel import *
# import VBAExcel

# from rich import print
# from rich import inspect
# from prettytable import PrettyTable

def EndIndexRowCol(sheet):
    # EndRow, EndCol = EndIndexRowCol(sheet)
    '''Определяем позиции первой и последней ячейки'''
    UsedRange = sheet.UsedRange
    # '''Количество занимаемых таблицей строк'''
    count_row = UsedRange.Rows.Count
    # '''Количество занимаемых таблицей колонок'''
    count_col = UsedRange.Columns.Count
    # '''Номер первой занимаемой строчки'''
    StartRow = UsedRange.Row
    # '''Номер первой занимаемой колонки'''
    StartCol = UsedRange.Column
    # '''Номер последней занимаемой строчки'''
    EndRow = StartRow + count_row - 1
    # '''Номер последней занимаемой колонки'''
    EndCol = StartCol + count_col - 1
    return EndRow, EndCol

'''Округление'''
def NFt(cells, okrug):
    try:
        cells.NumberFormat = okrug
    except:
        cells.NumberFormat = okrug.replace('.', ',')

def GO(sig):
    Excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    Excel.Visible = 1
    wb = Excel.ActiveWorkbook

    try:
        sheet = wb.Worksheets("Свод")
        sheet.Activate()
    except:
        text = f"В открытом файле Excel не обнаружен лист \"Свод\""
        sig.signal_err.emit(text)
        return

    sheet = wb.Worksheets("Свод")
    EndRow, EndCol = EndIndexRowCol(sheet)
    EndRow += 1
    StartRow = 3
    StartCol = 5

    sheet.Copy(After=wb.Worksheets[wb.Worksheets.Count])

    sleep(0.5)
    sig.signal_Probar.emit(5)

    sheet = wb.Worksheets[wb.Worksheets.Count]
    sheet.Activate()

    sheet.Activate()
    cel = sheet.Range("A3")
    cel.Activate()
    # sheet.Paste()
    sleep(0.5)
    sheet.Range(sheet.Columns(4), sheet.Columns(5)).Insert()
    sheet.Columns(3).ColumnWidth = 0
    sheet.Columns(4).ColumnWidth = 12
    sheet.Columns(5).ColumnWidth = 25
    sheet.Columns(6).ColumnWidth = 56
    sheet.Columns(6).HorizontalAlignment = 1
    sheet.Range("F2").HorizontalAlignment = 3
    sleep(0.5)
    sheet.Range("D2:I2").Formula = data = ["Поставка", "Тип МТР", "Техническая характеристика", "Поз.", "", "Наименование"]
    # sheet.Range("A1").Select()

    sig.signal_Probar.emit(10)
    StartCol = 7

    '''Убираем текст из ячеек (Поз.спец.XXX)'''
    Items = sheet.Range(sheet.Cells(StartRow, 9), sheet.Cells(EndRow, 9))
    Items.Replace( What="Поз.спец.* ", Replacement="")


    def cellsSelect(colList):
        col1, col2, col3, col4, col5, col6, col7, col8, col9 = colList
        # cel_1 = sheet.Range(sheet.Cells(StartRow, col1), sheet.Cells(EndRow, col1))
        cel_2 = sheet.Range(sheet.Cells(StartRow, col2), sheet.Cells(EndRow, col2))
        cel_3 = sheet.Range(sheet.Cells(StartRow, col3), sheet.Cells(EndRow, col3))
        cel_4 = sheet.Range(sheet.Cells(StartRow, col4), sheet.Cells(EndRow, col4))
        cel_5 = sheet.Range(sheet.Cells(StartRow, col5), sheet.Cells(EndRow, col5))
        # cel_6 = sheet.Range(sheet.Cells(StartRow, col6), sheet.Cells(EndRow, col6))
        cel_7 = sheet.Range(sheet.Cells(StartRow, col7), sheet.Cells(EndRow, col7))
        # cel_8 = sheet.Range(sheet.Cells(StartRow, col8), sheet.Cells(EndRow, col8))
        cel_9 = sheet.Range(sheet.Cells(StartRow, col9), sheet.Cells(EndRow, col9))
        # return cel_1, cel_2, cel_3, cel_4, cel_5, cel_6, cel_7, cel_8, cel_9
        return cel_2, cel_3, cel_4, cel_5, cel_7, cel_9
    
    sleep(5)
    colList = 7, 9, 10, 11, 12, 13, 14, 15, 16
    # cel_1, cel_2, cel_3, cel_4, cel_5, cel_6, cel_7, cel_8, cel_9 = cellsSelect(colList)
    cel_2, cel_3, cel_4, cel_5, cel_7, cel_9 = cellsSelect(colList)

    textNameList = [i.Value for i in cel_2]
    BoldNameList = [i.Font.Bold for i in cel_2]
    indexRow = [i for i in range(StartRow, StartRow + len(textNameList))]
    
    CountList = [i.Value for i in cel_7]

    postavka = ['Заказчик','Подрядчик']

    sig.signal_Probar.emit(15)

    text = ''
    NameCell = ''
    postavkaList = []
    NameCellList = []
    RowXXX = []
   
    for i in range(len(textNameList)):
        if textNameList[i] != None and BoldNameList[i] == False:
            if CountList[i] == None:
                Name = textNameList[i].strip()
                if NameCell != '':
                    NameCell += '\n' + Name
                else:
                    NameCell += Name
        for g in postavka:
            if textNameList[i] != None and g.lower() in textNameList[i].lower():
                text = g
        if textNameList[i] != None:
            if CountList[i] != None:
                NameCellList.append(NameCell)
                RowXXX.append(indexRow[i])
                postavkaList.append(text)
                if CountList[i+1] == None:
                    NameCell = ''


    Tip = []
    xxx = ''
    for i in range(len(textNameList)):
        if textNameList[i] != None and BoldNameList[i] == True and BoldNameList[i+1] == False:
            xxx = textNameList[i]
        if CountList[i] != None:
            Tip.append(xxx)

    '''Функция копирует значения ячейки вниз по группе МТР по строчкам с количеством'''
    def dataCopyNext(NameCollumn, nomerCol):
        DN = []
        xxx = ''
        for i in range(len(NameCollumn)):
            if BoldNameList[i] == True:
                xxx = ''
            if NameCollumn[i] != None:
                xxx = NameCollumn[i]
            if CountList[i] != None:
                DN.append(xxx)
        for i in range(len(RowXXX)):
            sheet.Cells(RowXXX[i], nomerCol).Value = DN[i]

    docCells = [i.Value for i in cel_3]
    dataCopyNext(docCells, 10)
    
    sig.signal_Probar.emit(20)

    '''Функция переносит значения ячейки до ближайшей строчки вниз с количеством или оставляет свое значение в этой строчке'''
    def dataDown(NameCollumn, nomerCol):
        DN = []
        xxx = ''
        for i in range(len(NameCollumn)):
            if NameCollumn[i] != None:
                xxx = NameCollumn[i]
            if CountList[i] != None:
                DN.append(xxx)
                xxx = ''
        for i in range(len(RowXXX)):
            sheet.Cells(RowXXX[i], nomerCol).Value = DN[i]
    
    CodProdCells = [i.Value for i in cel_4]
    dataDown(CodProdCells, 11)            
    Col12Cells = [i.Value for i in cel_5]
    dataDown(Col12Cells, 12)            
    PrimechCells = [i.Value for i in cel_9]
    dataDown(PrimechCells, 16)

    print('postavkaList = ', len(postavkaList))
    print('NameCellList = ', len(NameCellList))
    print('RowXXX = ', len(RowXXX))
    # postavkaList = postavkaList + ['yyyyyyyyy'] * 2
    print('postavkaList = ', len(postavkaList))


    data = [[postavkaList[i], Tip[i], NameCellList[i]] for i in range(len(RowXXX))]
    
    sig.signal_Probar.emit(25)

    for i in range(len(RowXXX)):
        sheet.Range(sheet.Cells(RowXXX[i], 4), sheet.Cells(RowXXX[i], 6)).Value = data[i]

    sleep(1)
    




    '''Удаляем не нужные строки, считая удаленные строки'''
    EndRow, EndCol = EndIndexRowCol(sheet)
    ColDel = sheet.Range(sheet.Cells(StartRow, 14), sheet.Cells(EndRow, 14))
    delta = len(ColDel) - len(RowXXX)
    nomerdelete = 0
    for i in range(len(ColDel)):
        while ColDel[i].Value == None:
            if nomerdelete >= delta:
                break
            else:
                ColDel[i].EntireRow.Delete()
                nomerdelete += 1
    
    sig.signal_Probar.emit(40)

    '''Удаляем зачеркнутые строчки'''
    EndRow, EndCol = EndIndexRowCol(sheet)            
    ColDel = sheet.Range(sheet.Cells(StartRow, 9), sheet.Cells(EndRow, 9))
    
    # ColDel.Replace( What="Поз.спец.* ", Replacement="")
    
    delta = sum([1 if i.Font.Strikethrough == True else 0 for i in ColDel])
    nomerdelete = 0
    for i in range(len(ColDel)):
        while ColDel[i].Font.Strikethrough == True:
            if nomerdelete >= delta:
                break
            else:
                ColDel[i].EntireRow.Delete()
                nomerdelete += 1

    sig.signal_Probar.emit(45)                
    sleep(0.5)
        

    EndRow, EndCol = EndIndexRowCol(sheet)
    array1 = sheet.Range(sheet.Cells(StartRow, 6), sheet.Cells(EndRow, 6))
    array2 = sheet.Range(sheet.Cells(StartRow, 9), sheet.Cells(EndRow, 9))
    array3 = sheet.Range(sheet.Cells(StartRow, 14), sheet.Cells(EndRow, 14))
    countelem = [i.Value for i in array3]
    rownomer = [i for i in range(StartRow, EndRow + 1)]
    arrayList = [f'{array1[i].Value}\n{array2[i].Value}' for i in range(1, len(array1) + 1)]
    countarray = Counter(arrayList)
    
    '''Считаем количество повторящихся строк'''
    dict_keys = [i for i in countarray.keys()]
    dict_values = [i for i in countarray.values()]

    '''Список повторяющихся строк'''
    keysres = []
    for i in range(len(dict_values)):
        if dict_values[i] > 1:
            keysres.append(dict_keys[i])

    sig.signal_Probar.emit(50)

    def vvod(xxx):
        try:
            xxx = round(float(xxx), 2)
        except:
            pass
        if xxx != '': 
            try:
                xxx = xxx.replace(',', '.')
            except:
                pass
        else: 
            xxx = 0
        try:
            xxx = round(float(xxx), 2)
        except:
            pass
        return xxx

    '''Собираем данные из повторяющихся строк'''
    schet = []
    schet_2 = []
    schet_3 = []
    schet_4 = []
    for g in keysres:
        xxx = 0.0
        xx2 = []
        xx3 = []
        xx4 = []
        for i in range(len(arrayList)):
            if g in arrayList[i]:
                aaa = vvod(countelem[i])
                try:
                    if isinstance(aaa, str):
                        if '/' in aaa:
                            aaa = aaa.split('/')[0]
                        if '+' in aaa:
                            aaa = aaa.split('+')
                            aaa = [float(r) for r in aaa]
                            aaa = sum(aaa)
                        aaa = float(aaa)
                    xxx += aaa
                except:
                    text = f'Значение в колонке "Количество" в строке {rownomer[i]}\n должно быть в числовом формате'
                    sig.signal_err.emit(text)
                    return

                xx2.append(sheet.Cells(rownomer[i], 1).Value)
                xx3.append(sheet.Cells(rownomer[i], 2).Value)
                eee = sheet.Cells(rownomer[i], 7).Value
                # xx4.append('' if eee == None else eee)
                if eee != None:
                    xx4.append(str(eee))
                    

        schet.append(xxx)
        schet_2.append(", ".join(xx2))
        schet_3.append(", ".join(xx3))
        schet_4.append(", ".join(xx4))

    sig.signal_Probar.emit(60)
    
    '''Округляем колонку с количеством'''
    NFt(sheet.Range(sheet.Cells(3, 14), sheet.Cells(EndRow, 14)), "0.0")
    '''Удаляем из Ед.изм. (колонка 13) значения после / '''
    cels = sheet.Range(sheet.Cells(3, 13), sheet.Cells(EndRow, 13))
    edizm = [str(i.Value) for i in cels]
    edizmres = []
    for i in edizm:
        if '/' in i:
            i = i.split('/')[0]
        edizmres.append((i,))
    cels.Value = edizmres
    sleep(0.1)

    '''Схлапываем повторяющиеся позиции'''
    for g in range(len(keysres)):
        ggg = 0
        nomerdelete = 0
        for i in range(len(arrayList)):
            if keysres[g] in arrayList[i] and ggg == 0:
                ggg = 1
                cel = sheet.Cells(rownomer[i], 14)
                cel.Value = (schet[g],)
                cel.WrapText = True
                
                cel = sheet.Cells(rownomer[i], 1)
                cel.Value = (schet_2[g],)
                cel.WrapText = True
                
                cel = sheet.Cells(rownomer[i], 2)
                cel.Value = (schet_3[g],)
                cel.WrapText = True
                
                cel = sheet.Cells(rownomer[i], 7)
                cel.Value = (schet_4[g],)
                cel.WrapText = True
                # sleep(0.1)
                continue
            
            if keysres[g] in arrayList[i] and ggg == 1:
                sheet.Cells(rownomer[i], 14).Value = ("Del",)

    sig.signal_Probar.emit(80)
    sleep(0.5)

    '''Удаляем строчки'''
    EndRow, EndCol = EndIndexRowCol(sheet)            
    ColDel = sheet.Range(sheet.Cells(StartRow, 14), sheet.Cells(EndRow, 14))
    delta = sum([1 if i.Value == "Del" else 0 for i in ColDel])
    for i in range(len(ColDel)):
        while ColDel[i].Value == "Del":
            if nomerdelete >= delta:
                break
            else:
                ColDel[i].EntireRow.Delete()
                nomerdelete += 1
    sleep(0.5)

    # '''Сортируем строчки по ключу колонки E'''
    # sheet.Sort.SortFields.Clear
    # sheet.Sort.SortFields.Add(Key=sheet.Range("E3"))
    # sor = sheet.Sort
    # sor.SetRange(sheet.Range(sheet.Cells(3, 1), sheet.Cells(EndRow, EndCol)))
    # sleep(0.5)
    # sor.Apply()
    
    
    
    sig.signal_Probar.emit(100)
    return None


if __name__ == "__main__":
    import sys
    from CollectorExcel import  app, sig
    GO(sig)
    sys.exit(app.exec_())

