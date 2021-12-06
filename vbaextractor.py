from win32com.client import Dispatch
#Code used to extract the VBA code from an Excel file, if the Macro is an AutoRun this code might run the function, if not then it successfully extracts the code for analysis. 
wbpath = 'C:\\test.xlsm'
xl = Dispatch("Excel.Application")
xl.Visible = 1
wb = xl.Workbooks.Open(wbpath)
vbcode = wb.VBProject.VBComponents(1).CodeModule
print (vbcode.Lines(1, vbcode.CountOfLines))
