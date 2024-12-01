' クエリを実行してデータをロードする
Sub QueryAndLoadData()
    ' クエリ名を指定
    Dim query_name As String
    query_name = "更新"
    'テーブル名を指定
    Dim table_name As String
    table_name = "更新"
    'シート名を指定

    ' クエリが存在するか確認
    Dim qry As WorkbookQuery
    Set qry = GetQueryByName(query_name)
    'クエリを実行
    qry.Refresh
End Sub
Sub CopyToNewSheet()

    Dim sheet_name As String
    sheet_name = "DATA"

    ' 新規シートを作成、すでにシートが存在する場合は末尾に数字を付加
    Dim ws As Worksheet
    Set ws = CreateSheet(sheet_name)

    ' テーブルを取得
    Dim sourceTable As ListObject
    Set sourceTable = GetTableByName("更新")

    ' sourceTable のデータを新規シートにコピー
    sourceTable.Range.Copy Destination:=ws.Cells(1, 1)

    '新規シートへ貼り付けたテーブル名を変更
    ws.ListObjects(1).Name = sheet_name

End Sub

Function GetQueryByName(queryName As String) As WorkbookQuery
    Dim qry As WorkbookQuery
    On Error Resume Next
    Set qry = ActiveWorkbook.Queries(queryName)
    On Error GoTo 0
    Set GetQueryByName = qry
End Function

Function GetTableByName(tableName As String) As ListObject
    Dim ws As Worksheet
    Dim tbl As ListObject
    On Error Resume Next
    For Each ws In ActiveWorkbook.Sheets
        Set tbl = ws.ListObjects(tableName)
        If Not tbl Is Nothing Then
            Set GetTableByName = tbl
            Exit Function
        End If
    Next ws
    On Error GoTo 0
    Set GetTableByName = Nothing
End Function

Function CreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    Dim i As Integer
    i = 1
    Do
        On Error Resume Next
        ws.Name = sheetName & i
        If Err.Number = 0 Then
            Exit Do
        End If
        On Error GoTo 0
        i = i + 1
    Loop
    Set CreateSheet = ws
End Function


