'検査シート
'開発用シート
Dim wsInspection As Worksheet
Dim wsDev As Worksheet

Dim rangeCheck As range
Dim rangeComment As range
Dim rangeTarget As range
Dim rangeAmmend As range

Dim commentCell As range
Dim targetCell As range
Dim ammendCell As range

dim targetRow as range

Sub main()
  Application.ScreenUpdating = False
  Set wsInspection = Sheets("検査")
  Set wsDev = Sheets("開発用")

  Set rangeCheck = wsInspection.range("K2:K11")
  Set rangeComment = wsDev.range("G2:G11")
  Set rangeTarget = wsDev.range("H2:H11")
  Set rangeAmmend = wsDev.range("I2:I11")

  Set commentCell = wsInspection.range("G2")
  Set targetCell = wsInspection.range("H2")
  Set ammendCell = wsInspection.range("I2")

  set targetRow = wsDev.range("B2")

  Dim comment As String
  Dim target As String
  Dim ammend As String

  comment = commentCell.value
  target = targetCell.value
  ammend = ammendCell.value

  comment = CheckListMapper(rangeComment) & comment
  'target = CheckListMapper(rangeTarget) & target
  ammend = CheckListMapper(rangeAmmend) & ammend


  commentCell.value = RemoveLastCrLf(comment)
  'targetCell.value = target
  ammendCell.value = RemoveLastCrLf(ammend)

  Debug.Print comment
  Debug.Print "test"
  Application.ScreenUpdating = True

End Sub

'rangeCheck?と対応する引数rangeの値を改行でつなげて返す
Function CheckListMapper(Byval range As range) As String
  Dim cell As range
  Dim result As String

  'rangeオブジェクトを配列に変換
  Dim rangeArray As Variant
  rangeArray = RangeToArray(range)

  Dim i As Integer
  i = 0

  For Each cell In rangeCheck
    If cell.value <> "" Then
      'i=0でresultに値がある場合は改行を追加
      If i = 0 And result <> "" Then
        result = result & vbCrLf
      End If
      result = result & rangeArray(i) & vbCrLf
    End If
    i = i + 1
  Next cell

  CheckListMapper = result
End Function

'rangeオブジェクトを配列に変換
Function RangeToArray(Byval range As range) As Variant
  Dim cell As range
  Dim result() As Variant
  Dim i As Integer
  i = 0
  For Each cell In range
    ReDim Preserve result(i)
    result(i) = cell.value
    i = i + 1
  Next cell
  RangeToArray = result
End Function

'末尾に改行があった場合は削除
'引数：文字列
Function RemoveLastCrLf(Byval str As String) As String
  Dim lastChar As String
  lastChar = Right(str, 2)
  If lastChar = vbCrLf Then
    str = Left(str, Len(str) - 2)
  End If
  RemoveLastCrLf = str
End Function