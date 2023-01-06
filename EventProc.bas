
 
Public Sub 正方形長方形1_Click()
On Error GoTo 正方形長方形1_Click_Err

    '画面の更新をオフにする
    Application.ScreenUpdating = False
    
    '勤務表出力
    If OutputWorkShedule() = False Then
        Call MsgBox(WORK_SHEDULE_OUTPUT_FAILED, vbInformation, CONFIRM)
        GoTo 正方形長方形1_Click_Exit
    End If
    
正方形長方形1_Click_Err:

正方形長方形1_Click_Exit:
    '画面の更新を再開する
    Application.ScreenUpdating = True
End Sub

