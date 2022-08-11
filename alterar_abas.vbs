Dim objExcel
Dim objWB

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWB = objExcel.Workbooks.Open("caminho_do_arquivo")
WScript.Sleep 10000

objExcel.Worksheets("nome_da_aba").Select
WScript.Sleep 5000

objWB.Save
objWB.Close True
objExcel.Quit
