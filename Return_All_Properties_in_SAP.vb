


Sub oi()

Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetscriptingEngine
Set Connection = App.Children(0)
Set Session = Connection.Children(0)

i = 0

For Each SapObject In Session.Children
   Call AdicionarSapObjectNaTabela(SapObject, i)
Next

End Sub

Sub AdicionarSapObjectNaTabela(SapObject, i)
On Error Resume Next
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
Range("X" & i) = SapObject.Width


Debug.Print SapObject.Text
If SapObject.Children.Count > 0 Then
        For Each SapOb In SapObject.Children
           Call AdicionarSapObjectNaTabela(SapOb, i)
        Next
    End If
End Sub
