'クリックされたセルに特定の値を入力する
Sub inspection_check(Byval target As range)
  'targetがg_checkRangeに含まれているか
  If intersect(target, checkRange) is nothing Then
   Exit Sub
  End If
  '何らかの値がある場合は空白
  If target.value <> "" Then
    target.value = ""
  Else
    target.value = "●"
  End If
end Sub