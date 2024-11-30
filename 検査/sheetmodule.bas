Private Sub Worksheet_Change(ByVal Target As Range)
If Target.row = 2 Or Target.row = 1 Then
    Exit Sub
End If
      Module.UpdateActiveRowFromViewRow (Target)
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Module.ShowActiveRowToViewRow
End Sub
