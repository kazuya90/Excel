Option Explicit

Dim wsInspection As Worksheet
Dim wsDev As Worksheet
'反映先の列番号_検査結果から備考まで
Dim cols As variant
'反映先の列番号_チェック項目まで
Dim start_check_col As long
'反映先の行番号
Dim active_row As Long
'反映元の行番号
Dim view_row As Long
'チェック項目
Dim rangeCheck As range

' 表示行の値を指定した列にコピーする
Sub UpdateActiveRowFromViewRow()
  '画面がちらつかないようにする
  Application.ScreenUpdating = False
  'チェック範囲
  Set rangeCheck = wsInspection.range("K2:K11")
  '反映する列の番号を配列で指定
  cols = Array(6, 7, 8, 9, 10)
  start_check_col = 11
  '開発用シートにある反映先行番号を取得
  active_row = Sheets("開発用").Cells(2, 2).value
  '反映元の行番号を指定
  view_row = 2
  '検査シートを取得
  Set wsInspection = Sheets("検査")
  '指定した列に値をコピー
  Dim col As Variant
  For Each col In cols
    wsInspection.Cells(active_row, col).value = wsInspection.Cells(view_row, col).value
  Next col

  'チェック項目をrangecheckにコピー
  Dim i As As Long
  i = 0
  For Each cell In rangeCheck
    If cell.value <> "" Then
      wsInspection.Cells(active_row, start_check_col + i).value = cell.value
      i = i + 1
    End If
    '画面を更新
    Application.ScreenUpdating = True
End Sub
