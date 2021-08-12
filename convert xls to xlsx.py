import win32com.client as win32
import pathlib 
import os

path = pathlib.Path()

for pass_obj in path.iterdir():
    if pass_obj.match("*.xls"):
        fname = str(pass_obj.resolve())
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(fname)

        movefolder_name = "Convert"
        if not os.path.exists(movefolder_name):
            os.mkdir(movefolder_name)

        wb.SaveAs(str(pass_obj) + "x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
        wb.Close()                               #FileFormat = 56 is for .xls extension
        excel.Application.Quit()