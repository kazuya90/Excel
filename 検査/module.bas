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
    active_row = activeCell.row

    Dim col As Long
    For col = 1 To 26 ' Columns A to Z
    '1,2行目は除外
    If (activeCell.row > 4) And (activeCell.row < 50005) Then
        ws.Cells(view_row, col).value = ws.Cells(active_row, col).value
        '選択行番号をA1セルに表示する
        ws.Cells(1, 1).value = active_row
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
  active_row = Cells(1, 1).value
  '反映元の行番号を指定
  Dim view_row As Long
  view_row = 2
  'アクティブシートを取得
  Dim ws As Worksheet
  Set ws = ActiveSheet
  '指定した列に値をコピー
  Dim col As Variant
  For Each col In cols
    ws.Cells(active_row, col).value = ws.Cells(view_row, col).value
  Next col
End Sub
