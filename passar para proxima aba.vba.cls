VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub passar para proxima aba()

cont = 1

Do Until cont > Sheets.Count
Sheets(cont).Select

cont = cont + 1
Loop

End Sub

