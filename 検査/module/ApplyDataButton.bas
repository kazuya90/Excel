'rangeの値を他のrangeにコピーする関数
'引数：コピー元range,コピー先range

Sub CopyRange()
  '「全検査結果一覧」という名前のシートを取得
  Dim ws As Worksheet
  Set ws = Sheets("全検査結果一覧")
  'コピー元のrangeを指定
  Dim rngFrom As Range
  Set rngFrom = ws.Range("G2:J50004")

  'コピー先のシート「検査」を取得
  Dim wsTo As Worksheet
  Set wsTo = Sheets("検査")
  'コピー先のrangeを指定
  Dim rngTo As Range
  Set rngTo = wsTo.Range("F5;I50007")
  rngTo.Value = rngFrom.Value
End Sub