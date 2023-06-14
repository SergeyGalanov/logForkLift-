# coding: utf-8
from __future__ import unicode_literals
#from apso_utils import msgbox
import sys, datetime
from datetime import date


def sheetChanged(args=None):
    #Подключение к книге
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    #Получение активной вкладки
    sheet = model.CurrentController.ActiveSheet
    #Получение текущего выделенного фрагмента
    oSelection = model.getCurrentSelection()
    #Получение адреса текущего фрагмента
    
    #Обработка появления ошибки появляющейся при изменении группы ячеек при активном фильтре, когда ячейки не идут по порядку. Возможно есть другой способ получить текущее выделение. Добавил обрабутку, чтоб не смущала пользователя до тех пор пока не найду решение.
    try:
        oArea = oSelection.getRangeAddress()
        #print(oArea)
    except:
        print("Error: Change selection when filter is activated")
        return
    
    if  oArea.StartRow != oArea.EndRow or oArea.StartColumn != oArea.EndColumn or oArea.StartRow == 0:
        print('range')
        return
        
    #print(oSelection.getString()) #Получили содержимое активной ячейки
    
    #Автоматическое доавление даты при выборе номера машины. Если номер выбран, то в соответствующую ячейку колонки А ставим дату и время. Если удалили данныев ячейке, то удалили и дату.
    if sheet.getCellByPosition(oArea.StartColumn,0).String == 'Машина №' or sheet.getCellByPosition(oArea.StartColumn,0).String == 'Тип вмешательства':
        if len(oSelection.getString())>0:
           
            #перевод даты и времени в десятичное число используя функцию convert_date_to_excel_ordinal
            curent_time_decimal = convert_date_to_excel_ordinal()
            #Внесение даты в десятичном формате в ячейку с форматом дата "DD.MM.YYYY HH:MM:SS"
            sheet.getCellByPosition(0,oArea.StartRow).setValue(curent_time_decimal)
            txtFormula = '=IF(LEN(I{0})>0;1;IF(AND(LEN(L{0})=0;Len(J{0})=2);2;3))'.format(oArea.StartRow+1)
            # =IF(LEN(B108)<>0;VLOOKUP(B108;Погрузчики.$D$2:$F$54;3;0);"")
            sheet.getCellByPosition(2,oArea.StartRow).setFormula('=IF(LEN(B{0})<>0;VLOOKUP(B{0};Погрузчики.$D$2:$F$54;3;0);"")'.format(oArea.StartRow+1)) 
            sheet.getCellByPosition(17,oArea.StartRow).setFormula('=IF(LEN(I{0})>0;1;IF(AND(LEN(L{0})=0;Len(J{0})=2);2;3))'.format(oArea.StartRow+1)) 
            
        else:
            #Очистка ячейки, если произошло удаление номера из ячейки
            sheet.getCellByPosition(0,oArea.StartRow).clearContents(2)
            sheet.getCellByPosition(2,oArea.StartRow).clearContents(16)
            sheet.getCellByPosition(17,oArea.StartRow).clearContents(16)
            
    #Отметка о выполнении 
    if sheet.getCellByPosition(oArea.StartColumn,0).String == 'Отметка о выполнении':
        if len(oSelection.getString())>0:
           
            #перевод даты и времени в десятичное число используя функцию convert_date_to_excel_ordinal
            curent_time_decimal = convert_date_to_excel_ordinal()
        
            #Внесение даты в десятичном формате в ячейку с форматом дата "DD.MM.YYYY HH:MM:SS"
            sheet.getCellByPosition(10,oArea.StartRow).setValue(curent_time_decimal)
        else:
            #Очистка ячейки, если произошло удаление номера из ячейки
            sheet.getCellByPosition(10,oArea.StartRow).clearContents(2)    
 

def convert_date_to_excel_ordinal():
    
    #Получение текущей даты и времени
    current_time = datetime.datetime.now() 
    #Перевод время в секунды
    current_time_seconds = current_time.hour*3600+current_time.minute*60+current_time.second
    
    # Specifying offset value i.e.,
    # the date value for the date of 1900-01-00
    offset = 693594
    current = date(current_time.year, current_time.month, current_time.day)
 
    # Calling the toordinal() function to get
    # the excel serial date number in the form
    # of date values
    n = current.toordinal()
    return float(n - offset)+(float(current_time_seconds) / 86400)
    