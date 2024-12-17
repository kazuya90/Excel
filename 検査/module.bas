Option Explicit
' アクティブ行の値を2行目にコピーする
Sub ShowActiveRowToViewRow()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    '表示する行番号を指定する
    Dim view_row As Long
    view_row = 2
    
    '選択行番号
    Dim active_row As Long
    active_row = activeCell.Row

    Dim col As Long
    For col = 1 To 26 ' Columns A to Z
    '1,2行目は除外
    If (activeCell.Row > 4) And (activeCell.Row < 50005) Then
        ws.Cells(view_row, col).Value = ws.Cells(active_row, col).Value
        '選択行番号をA1セルに表示する
        ws.Cells(1, 1).Value = active_row
    End If
    Next col
    

End Sub

' 表示行の値を指定した列にコピーする
Sub UpdateActiveRowFromViewRow()
  '反映する列の番号を配列で指定
  Dim cols As Variant
  cols = Array(10, 11, 12, 13, 14)
  'A1セルの反映先行番号を取得
  Dim active_row As Long
  active_row = Cells(1, 1).Value
  '反映元の行番号を指定
  Dim view_row As Long
  view_row = 2
  'アクティブシートを取得
  Dim ws As Worksheet
  Set ws = ActiveSheet
  '指定した列に値をコピー
  Dim col As Variant
  For Each col In cols
    ws.Cells(active_row, col).Value = ws.Cells(view_row, col).Value
  Next col
End Sub
'j1からm1のセルに入力された文字列を、j2からm2のセルの文字列の先頭にそれぞれ追記する
Sub PrependViewRow()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim prepend_cell As Range
    Dim view_cell As Range
    Dim col As Long
    For col = 10 To 13
        Set prepend_cell = ws.Cells(1, col)
        Set view_cell = ws.Cells(2, col)
        If (prepend_cell <> "") Then
            view_cell.Value = prepend_text(prepend_cell, view_cell)
        End If
    Next col
End Sub

'引数で指定したセルの値を、２つ目の引数に指定したセルの先頭に追記する
Function prepend_text(cell As Range, cell2 As Range) As String
  'cell2の値が空の場合は、cellの値をそのまま返す
  If cell2.Value = "" Then
      prepend_text = cell.Value
  Else
      prepend_text = cell.Value & vbCrLf & cell2.Value
  End If
End Function
