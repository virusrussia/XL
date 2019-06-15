'''
Created on 12 мая 2019 г.

@author: alex
'''

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import os
from config import *
from fileinput import close
from config import *
from services import *
import logging


pathOfFile=os.path.dirname(__file__)
path=os.path.split(pathOfFile)[0]
sourcePath=os.path.join(path, SOURCEDIR)
print (sourcePath)

MRF={}

#Более новая штатка
WBNwer = load_workbook(os.path.join(sourcePath, "1.xlsx"), data_only=True)
print ("Загружено!")
#Более старая штатка
WBOlder = load_workbook(os.path.join(sourcePath, "2.xlsx"), data_only=True)
print ("Загружено!")
WSNwer = WBNwer.worksheets[0]
WSOlder = WBOlder.worksheets[0]


#Находим номер столбцов с табельными номерами
tabNwerNumber = getColumnID(WSNwer, columnName=EMPLOYERID)
tabOlderNumber = getColumnID(WSOlder, columnName=EMPLOYERID)
print ("Нашли столбцы с табельными")
tabOlderFOT=getColumnID(WSOlder, columnName=FOT)

#Получаем кортежи с табельными номерами. Устанавливаем min_row=2, что бы не взять заголовок таблицы.
rNew=tuple(WSNwer.iter_cols(min_col=tabNwerNumber, max_col=tabNwerNumber,min_row=2))
rOld=tuple(WSOlder.iter_cols(min_col=tabOlderNumber, max_col=tabOlderNumber))


#Если Табельный номер из новой штатки без учета количества назначений присутствует в старой штатке, то запоминаем его для последюующего вывода.
for i in rNew[0]:
    print (i)
    for y in rOld[0]:
        if (i.value!="") and (str(i.value).split("-")[0] == str(y.value).split("-")[0]):
            MRFName = WSNwer.cell(row=i.row, column=getColumnID(WSNwer, columnName="1")).value
            
            if MRFName not in MRF.keys():
                MRF[MRFName] = Workbook()
                MRF[MRFName].create_sheet("Лист 1", 0)               
                copyTableHeader(MRF[MRFName].worksheets[0], WSNwer)               
                print (MRF)
                
            maxRow=MRF[MRFName].worksheets[0].max_row+1
            MRF[MRFName].worksheets[0].insert_rows(maxRow)           

                
            for w in WSNwer.iter_rows(min_row=i.row, max_row=i.row):
                for z in range(1, len(w)+1):
                    copyCell(MRF[MRFName].worksheets[0].cell(row=maxRow, column=z), WSNwer.cell(row=i.row, column=z))
                
                
                copyCell(MRF[MRFName].worksheets[0].cell(row=maxRow, column=len(w)+1), WSOlder.cell(row=y.row, column=tabOlderFOT))
                
                
                MRF[MRFName].worksheets[0].cell(row=maxRow, column=len(w)+2).value=f"={get_column_letter(len(w))}{maxRow}/{(get_column_letter(len(w)+1))}{maxRow}-1"
for i in MRF:
    MRF[i].save(os.path.join(sourcePath, f"{i}.xlsx")) 
      
WBNwer.close()
WBOlder.close()