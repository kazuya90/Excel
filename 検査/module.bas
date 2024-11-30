Option Explicit
' アクティブ行の値を2行目にコピーする
Sub ShowActiveRowToViewRow()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    '表示する行番号を指定する
    Dim view_row As Integer
    view_row = 2
    
    '選択行番号
    Dim active_row As Integer
    active_row = activeCell.Row

    Dim col As Integer
    For col = 1 To 26 ' Columns A to Z
    '1,2行目は除外
    If activeCell.Row <> 1 And activeCell.Row <> 2 Then
        ws.Cells(view_row, col).Value = ws.Cells(active_row, col).Value
    End If
    Next col
    
    '選択行番号をA1セルに表示する
    ws.Cells(1, 1).Value = active_row
End Sub

' 表示行の値を指定した列にコピーする
Sub UpdateActiveRowFromViewRow()
  '反映する列の番号を配列で指定
  Dim cols As Variant
  cols = Array(7, 8, 9,10)
  'A1セルの反映先行番号を取得
  Dim active_row As Integer
  active_row = Cells(1, 1).Value
  '反映元の行番号を指定
  Dim view_row As Integer
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
