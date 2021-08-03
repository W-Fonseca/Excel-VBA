'Procura um valor em uma planilha 

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open("C:\Users\wellington.fonseca\Desktop\teste.xlsx")
Set objWorksheet = objWorkbook.Worksheets("Planilha1")

Const xlByRows = 1
Const xlPrevious = 2

Set objExcel = GetObject(,"Excel.Application")
LastRow = objExcel.ActiveSheet.Cells.Find("juvenal", , , , xlByRows, xlPrevious).Address

MsgBox LastRow
