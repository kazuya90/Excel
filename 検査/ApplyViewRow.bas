Option Explicit

' アクティブ行の値を表示用の行にコピーする
Sub ShowActiveRowToViewRow()
  ' 画面更新を停止
  Application.ScreenUpdating = False

  Dim ws As Worksheet
  Set ws = ActiveSheet

  '表示する行番号を指定する
  Dim view_row As Long
  view_row = 2

  '選択行番号
  Dim active_row As Long
  active_row = ActiveCell.Row

  Dim col As Long
  For col = 2 To 10 ' Columns B To J
    '記入用の行以外は除外
    If (ActiveCell.Row > 13) And (ActiveCell.Row < 50005) Then
      If (col < 5) Then
        ws.Cells(view_row - 1, col + 5).value = ws.Cells(active_row, col).value
      Elseif (col = 5) Then
        '行番号は反映しない

      Else
        ws.Cells(view_row, col).value = ws.Cells(active_row, col).value
      End If
    End If
  Next col


  ' 画面更新を停止
  Application.ScreenUpdating = True
End Sub

' アクティブ行の値を表示用の行にコピーする
Sub ShowActiveRowToViewRow_Compare()
  ' 画面更新を停止
  Application.ScreenUpdating = False

  Dim ws As Worksheet
  Set ws = ActiveSheet

  '表示する行番号を指定する
  Dim view_row As Long
  view_row = 2

  '選択行番号
  Dim active_row As Long
  active_row = ActiveCell.Row

  Dim col As Long
  For col = 1 To 26 ' Columns B To J
    '記入用の行以外は除外
    If (ActiveCell.Row > 4) And (ActiveCell.Row < 50005) Then
      ws.Cells(view_row, col).value = ws.Cells(active_row, col).value
    End If
  Next col

  ' 画面更新を停止
  Application.ScreenUpdating = True
End Sub

'チェックリストへ値を反映
Sub UpdateCheckList()
  '対象行の✓を表示・編集用へコピーする
  '対象行の
  g_i_targetCheckRange
  '表示・編集用
  g_CheckRange
  '対象行の値を表示・編集用にコピー
  Dim i As Long
  i = 0
  For Each cell In g_i_targetCheckRange
    g_CheckRange.value = cell.value
    i = i + 1
  Next cell

End Sub