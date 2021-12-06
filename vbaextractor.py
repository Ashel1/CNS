from win32com.client import Dispatch
#Code used to extract the VBA code from an Excel file, if the Macro is an AutoRun this code might run the function, if not then it successfully extracts the code for analysis. 
#All the VBA codes are designed to delete a folder (testfolder) in C:\, the delete file VBA deletes the test.txt file that must be created within the testfolder in teh C drive 
wbpath = 'C:\\test.xlsm'
xl = Dispatch("Excel.Application")
xl.Visible = 1
wb = xl.Workbooks.Open(wbpath)
vbcode = wb.VBProject.VBComponents(1).CodeModule
print (vbcode.Lines(1, vbcode.CountOfLines))
