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
' クエリを実行してデータをロードする
Sub QueryAndLoadData()
    ' クエリ名を指定
    Dim query_name As String
    query_name = "更新"
    'テーブル名を指定
    Dim table_name As String
    table_name = "更新"
    'シート名を指定
    Dim sheet_name As String
    sheet_name = "DATA" 

    ' クエリが存在するか確認
    Dim qry As Object
    On Error Resume Next
    Set qry = ActiveWorkbook.Queries(query_name)
    On Error GoTo 0
    If qry Is Nothing Then
        MsgBox "クエリ '" & query_name & "' が存在しません。"
        Exit Sub
    End If

    ' クエリを実行（リフレッシュ）
    qry.Refresh
    ' クエリの更新を待つ
    Do While qry.Refreshing
        DoEvents
    Loop

    ' 新規シートを作成、すでにシートが存在する場合は末尾に数字を付加
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))

    Dim i As Integer
    i = 1
    Do
        On Error Resume Next
        ws.Name = sheet_name & i
        If Err.Number = 0 Then
            Exit Do
        End If
        On Error GoTo 0
        i = i + 1
    Loop

    ' テーブルを取得
    Dim sourceTable As ListObject
    Set sourceTable = GetTableByName("更新")

    ' sourceTable のデータを新規シートにコピー
    sourceTable.Range.Copy Destination:=ws.Cells(1, 1)

End Sub
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