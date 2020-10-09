import pandas as pd
import win32com.client

xlApp = win32com.client.Dispatch("Excel.Application")

filename,password = r"your path", 'password'
excel = xlApp.Workbooks.Open(filename, False, True, None, Password=password)
excel = excel.Worksheets("Sheet1")

col=1
while excel.Cells(1, col).Value is not None:
    col +=1

row=1
while excel.Cells(row, 1).Value is not None:
    row +=1
    
title=excel.Range(excel.Cells(1,1),excel.Cells(1,col-1)).Value
data=excel.Range(excel.Cells(2,1),excel.Cells(row-1,col-1)).Value

df = pd.DataFrame(data,columns=title[0])
