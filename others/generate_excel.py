import win32com.client
import os

xlsApp = win32com.client.Dispatch("Excel.Application")
xlsApp.Visible = 0                  #顯示 Excel
xlsBook = xlsApp.Workbooks.Add()    #新增一工作簿
xlsSheet = xlsBook.Worksheets(1)  #新增的工作簿預設含三個工作表
xlsSheet.Cells(1,1).Value=100
xlsBook.SaveAs('C:\\Users\\jayhsieh\\Desktop\\text.xlsx')
xlsBook.Close(SaveChanges=1)
xlsApp.Quit()
del xlsApp
