Dim objExcel
Dim objWB


Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True


Set objWB = objExcel.Workbooks.Open("C:\Users\55169\Desktop\datacoin.xlsx")
WScript.Sleep 10000


objExcel.Worksheets("Datacoin").Select
WScript.Sleep 5000


objWB.Save
objWB.Close True
objExcel.Quit