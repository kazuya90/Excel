Option Explicit
'検査シート
Public g_InspectionSheet As Worksheet
Public g_devSheet As Worksheet

Public g_checkRange As range

Sub setGlobal()
  Set g_InspectionSheet = Sheets("検査")
  Set g_devSheet = Sheets("開発用")
  Set checkRange = g_InspectionSheet.range("K2:K11")
End Sub
