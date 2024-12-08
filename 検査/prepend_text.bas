'j1からm1のセルに入力された文字列を、j2からm2のセルの文字列の先頭にそれぞれ追記する
Sub prepend_text()
    Dim ws As Worksheet
    Set ws = activesheet
    For j = 10 To 13
        Set cell = ws.Cells(1, j)
        Set cell2 = ws.Cells(2, j)
        cell2.Value = prepend_text(cell, cell2)
    Next j
End Sub

'引数で指定したセルの値を、２つ目の引数に指定したセルの先頭に追記する
function prepend_text(cell as range, cell2 as range) as string
    prepend_text = cell2.value & cell.value
end function

