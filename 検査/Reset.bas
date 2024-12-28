'rangeの値を他のrangeにコピーする関数
'引数:列番号
Sub CopyRange(target_col As Integer)
  '開発用シートにある表示・編集対象の行番号を取得
  Dim ws As Worksheet
  Set ws = Sheets("開発用")
  Dim target_row As Long
  target_row = ws.Range("B2").Value

  Dim wsFrom As worksheet
  Set wsFrom = Sheets("全検査結果一覧")
  Dim rngFrom As Range
  Set rngFrom = wsFrom.cells(target_row-3, target_col+1)

  'コピー先のシート「検査」を取得
  Dim wsTo As Worksheet
  Set wsTo = Sheets("検査")
  'コピー先のrangeを指定
  Dim rngTo As Range
  Set rngTo = wsTo.cells(target_row, target_col)

  rngTo.Value = rngFrom.Value
End Sub