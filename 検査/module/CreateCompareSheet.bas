'比較シートを作成する
Sub createCompareSheet()
  Dim ws As Worksheet
  Set ws = Sheets("全検査結果一覧")

  Dim sheetandtable_name As String
  sheetandtable_name = getNow()
  '新しいシートを作成
  Sheets.Add after:=Sheets(Sheets.Count)

  '新しいシートの名前を変更
  '同じシート名がある場合はそのシートを削除
  If SheetExists(sheetandtable_name) Then
    Application.DisplayAlerts = False
    Sheets(sheetandtable_name).Delete
    Application.DisplayAlerts = True
  End If
  ActiveSheet.Name = sheetandtable_name

  '全検査結果一覧テーブルをコピー
  ws.UsedRange.Copy
  ActiveSheet.Paste

  '貼り付けたテーブルの名前を変更
  ActiveSheet.ListObjects(1).Name = sheetandtable_name

  '全検査結果一覧以外のクエリを削除
  KeepSpecificQueries

  setTableName

End Sub

'現在の日時を取得する
Function getNow() As String
  getNow = Format(Now, "_mmdd_hhmm")
End Function

Sub KeepSpecificQueries()
  Dim wb As Workbook
  Dim pq As WorkbookQuery
  Dim keepList As Object
  Dim queryNames As Collection
  Dim queryName As Variant

  ' 対象のブックを設定（現在のブック）
  Set wb = ThisWorkbook

  ' 残したいクエリ名のリストを定義（大文字小文字を区別）
  Set keepList = CreateObject("Scripting.Dictionary")
  keepList.Add "全検査結果一覧", True

  ' クエリ一覧を保存してループ（コレクションを利用して安全に処理）
  Set queryNames = New Collection
  For Each pq In wb.Queries
    queryNames.Add pq.Name
  Next pq

  ' 削除処理（ループの中で削除を安全に行う）
  For Each queryName In queryNames
    If Not keepList.exists(queryName) Then
      wb.Queries(queryName).Delete
    End If
  Next queryName

End Sub

'開発用シートのD2セルにテーブル名を入力する
Sub setTableName()
  Dim ws As Worksheet
  Set ws = Sheets("開発用")

  '格納用の変数を定義
  Dim tableNameList As Collection
  Set tableNameList = New Collection

  Dim tbr As Variant

  'worksheetをforeachで回す
  Dim wsName As String
  For Each ws In ThisWorkbook.Worksheets
    'listobjectがある場合はテーブル名を格納
    For Each tbr In ws.ListObjects
      tableNameList.Add tbr.Name
    Next tbr
  Next ws

  'テーブル名を開発用シートに入力
  Dim i As Integer
  For i = 1 To tableNameList.Count
    ws.Cells(i + 1, 4).Value = tableNameList(i)
  Next i

End Sub

Function SheetExists(sheetName As String) As Boolean
  Dim ws As Worksheet

  ' シートをループで確認
  On Error Resume Next
  Set ws = ThisWorkbook.Sheets(sheetName)
  On Error Goto 0

    ' シートが見つかった場合は True、見つからない場合は False
    SheetExists = Not ws Is Nothing
End Function

