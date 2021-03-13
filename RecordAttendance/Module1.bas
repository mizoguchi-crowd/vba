Attribute VB_Name = "Module1"
Private Const 日付列 As String = "A"
Private Const 出勤時刻列 As String = "B"
Private Const 退勤時刻列 As String = "C"

Private Sub Shukkin_Click()
  Call 打刻処理(出勤時刻列)
End Sub

Private Sub Taikin_Click()
  Call 打刻処理(退勤時刻列)
End Sub

Private Sub 打刻処理(打刻列)
  Call 日付記録
  Dim 打刻行 As String
  打刻行 = 打刻行取得()
  Dim 打刻区分 As String
  打刻区分 = 打刻区分取得(打刻列)
  
  If (多重打刻チェック(打刻行, 打刻列)) Then
    MsgBox ("本日はすでに" + 打刻区分 + "時間を打刻済のため、記録しませんでした")
    Exit Sub
  Else
    Sheet1.Cells(打刻行, 打刻列) = 現在時刻取得()
    MsgBox (打刻区分 + "時間を打刻しました")
  End If
End Sub

Private Function 現在時刻取得() As String
  現在時刻取得 = Format(Now, "hh:mm:ss")
End Function

Private Sub 日付記録()
  Dim 当日日付 As String
  当日日付 = Format(Now, "YYYY/MM/DD")
  Dim 日付列最終行の日付 As String
  日付列最終行の日付 = Sheet1.Cells(Rows.Count, 日付列).End(xlUp)
  
  If (日付列最終行の日付 = 当日日付) Then
    '何もしない
  Else
    Sheet1.Cells(Sheet1.Cells(Rows.Count, 日付列).End(xlUp).Row + 1, 日付列) = 当日日付
  End If

End Sub

Private Function 打刻行取得() As String
  打刻行取得 = Sheet1.Cells(Rows.Count, 日付列).End(xlUp).Row
End Function

Private Function 多重打刻チェック(打刻行, 打刻列) As Boolean
  If (Sheet1.Cells(打刻行, 打刻列) = "") Then
    多重打刻チェック = False
  Else
    多重打刻チェック = True
  End If
End Function

Private Function 打刻区分取得(打刻列) As String
  If (打刻列 = 出勤時刻列) Then
    打刻区分取得 = "出勤"
  ElseIf (打刻列 = 退勤打刻列) Then
    打刻区分取得 = "退勤"
  End If
  
End Function
