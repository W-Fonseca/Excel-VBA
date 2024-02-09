Sub oi()

Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetscriptingEngine
Set Connection = App.Children(0)
Set Session = Connection.Children(0)

   Range("A1").Value = "SapObject.ContainerType"
    Range("B1").Value = "SapObject.Id"
    Range("C1").Value = "SapObject.Name"
    Range("D1").Value = "SapObject.Parent"
    Range("E1").Value = "SapObject.Type"
    Range("F1").Value = "SapObject.TypeAsNumber"
    Range("G1").Value = "SapObject.AccLabelCollection"
    Range("H1").Value = "SapObject.AccText"
    Range("I1").Value = "SapObject.AccTextOnRequest"
    Range("J1").Value = "SapObject.AccTooltip"
    Range("K1").Value = "SapObject.Changeable"
    Range("L1").Value = "SapObject.DefaultTooltip"
    Range("M1").Value = "SapObject.Height"
    Range("N1").Value = "SapObject.IconName"
    Range("O1").Value = "SapObject.IsSymbolFont"
    Range("P1").Value = "SapObject.Left"
    Range("Q1").Value = "SapObject.Modified"
    Range("R1").Value = "SapObject.ParentFrame"
    Range("S1").Value = "SapObject.ScreenLeft"
    Range("T1").Value = "SapObject.ScreenTop"
    Range("U1").Value = "SapObject.Text"
    Range("V1").Value = "SapObject.Tooltip"
    Range("W1").Value = "SapObject.Top"
    Range("X1").Value = "SapObject.Width"

i = 1

For Each SapObject In Session.Children
   Call AdicionarSapObjectNaTabela(SapObject, i)
Next

End Sub

Sub AdicionarSapObjectNaTabela(SapObject, i)
On Error Resume Next
validador = ""
validador = SapObject.ContainerType
If validaor > "" Then
i = i + 1
Range("A" & i) = SapObject.ContainerType
Range("B" & i) = SapObject.ID
Range("C" & i) = SapObject.Name
Range("D" & i) = SapObject.Parent
Range("E" & i) = SapObject.Type
Range("F" & i) = SapObject.TypeAsNumber
Range("G" & i) = SapObject.AccLabelCollection
Range("H" & i) = SapObject.AccText
Range("I" & i) = SapObject.AccTextOnRequest
Range("J" & i) = SapObject.AccTooltip
Range("K" & i) = SapObject.Changeable
Range("L" & i) = SapObject.DefaultTooltip
Range("M" & i) = SapObject.Height
Range("N" & i) = SapObject.IconName
Range("O" & i) = SapObject.IsSymbolFont
Range("P" & i) = SapObject.Left
Range("Q" & i) = SapObject.Modified
Range("R" & i) = SapObject.ParentFrame
Range("S" & i) = SapObject.ScreenLeft
Range("T" & i) = SapObject.ScreenTop
Range("U" & i) = SapObject.Text
Range("V" & i) = SapObject.Tooltip
Range("W" & i) = SapObject.Top
Range("X" & i) = SapObject.Widtho
End If
Debug.Print SapObject.Text
On Error GoTo 0
On Error GoTo SAIDA2
If SapObject.Children.Count > 0 Then
        For Each SapOb In SapObject.Children
           Call AdicionarSapObjectNaTabela(SapOb, i)
        Next
    End If
SAIDA2:
End Sub
