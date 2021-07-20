Sub abriremoutroformato()

Dim doc As Object
Dim wa As Object
Set wa = CreateObject("word.application")
Set doc = wa.Documents.Open("C:\Users\wellington.fonseca\Desktop\lista-a-incluidos-09-07-2021.pdf", False, Format:="PDF Files")
wa.Selection.WholeStory
wa.Selection.Copy
Worksheets("planilha1").Select
Range("A1").Select
ActiveSheet.Paste

End Sub
