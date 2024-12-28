'比較用シートへ比較元のシートを反映
Sub ApplyCompare()
  '比較元のシートを選択するセル
  Dim CompareCell As Range
  Set CompareCell = Sheets("比較").Range("B3")

  '数式を反映するセル
  Dim FormulaCell As Range
  Set FormulaCell = Sheets("比較").Range("A5")

  'CompareCellの値を取得
  Dim CompareSheetName As String
  CompareSheetName = CompareCell.Value
  'FrmulaCellへ数式を反映
  FormulaCell.Formula = "=_" & CompareSheetName
End Sub