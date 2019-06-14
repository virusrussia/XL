'''
Created on 12 мая 2019 г.

@author: alex
'''

from openpyxl import Workbook, load_workbook
import os
from config import *
from fileinput import close

def getColumnEmployerID (WB, columnName=EMPLOYERID):
   "Ищем номер столбца с заданным именем"
   for i in range(1, WSNewer.max_column+1):
    if WSNewer.cell(row=1, column=i).value == columnName:
        return i
        break
 

pathOfFile=os.path.dirname(__file__)
path=os.path.split(pathOfFile)[0]
sourcePath=os.path.join(path, SOURCEDIR)
print (sourcePath)

WBNewer = load_workbook(os.path.join(sourcePath, "1.xlsx"))
WBOlder = load_workbook(os.path.join(sourcePath, "2.xlsx"))

WSNewer=WBNewer.worksheets[0]
WSOlder=WBOlder.worksheets[0]

tabNewr = getColumnEmployerID(WSNewer)
tabOlder = getColumnEmployerID(WSOlder)

r=tuple(WSNewer.iter_cols(min_col=tabNewr, max_col=tabNewr))

for i in r[0]:
    print (i.value.split('-')[0])

WBNewer.close()