

'【概要】祝日かどうか
Public Function IsHoliday(ByVal strDate As String, _
                       ByVal arrHoliday As Variant) As Boolean
On Error GoTo IsHoliday_Err
    
    IsHoliday = False
    
    '祝日配列に存在する場合、祝日という判定になる
    If InArr(arrHoliday, strDate) = True Then
        IsHoliday = True
    End If
    
IsHoliday_Err:

IsHoliday_Exit:

End Function


'【概要】日ごとの色を取得
Public Function GetDayColor(ByVal strDate As String, _
                            ByVal arrHoliday As Variant) As Long
On Error GoTo GetDayColor_Err
    
    '土曜日の場合
    If Weekday(strDate) = vbSaturday Then
        GetDayColor = RGB(157, 204, 224)
        GoTo GetDayColor_Exit
    End If
    
    '日曜日,祝日の場合
    If Weekday(strDate) = vbSunday Or _
        IsHoliday(strDate, arrHoliday) = True Then
        GetDayColor = RGB(250, 219, 218)
        GoTo GetDayColor_Exit
    End If
    
    '平日の場合
    GetDayColor = RGB(255, 255, 255)
       
GetDayColor_Err:

GetDayColor_Exit:

End Function


'【概要】ヘッダ出力
Public Function OutputHeader(ByVal objWb As Excel.Workbook, _
                        ByVal strSheetName As String, _
                        ByVal lngYear As Long, _
                        ByVal lngMonth As Long, _
                        ByVal lngHeaderRow As Long, _
                        ByVal arrHoliday As Variant) As Boolean
On Error GoTo OutputHeader_Err
    
    OutputHeader = False
    
    Dim lngHolidayRow As Long
    Dim lngOutputCol As Long
    Dim lngDayCount As Long
    Dim lngDayColor As Long
    Dim lngWeekdayRow As Long
    Dim strDate As String
    
    '曜日列
    lngWeekdayRow = lngHeaderRow + 1
    
    With objWb.Worksheets(strSheetName)
        '年を出力
        .Cells(lngHeaderRow, 1).Value = CStr(lngYear) & "年"
        '氏名を出力
        .Cells(lngWeekdayRow, 1).Value = "氏名"
        '罫線を引く
        If BorderAround(objWb, strSheetName, lngHeaderRow, 1, lngHeaderRow, 1) = False Then
            GoTo OutputHeader_Exit
        End If
        '罫線で囲む
        If BorderAround(objWb, strSheetName, lngWeekdayRow, 1, lngWeekdayRow, 1) = False Then
            GoTo OutputHeader_Exit
        End If
        '月を出力
        .Cells(lngHeaderRow, 2).Value = CStr(lngMonth) & "月"
        '役割を出力
        .Cells(lngWeekdayRow, 2).Value = "役割"
        '罫線を引く
        If BorderAround(objWb, strSheetName, lngHeaderRow, 2, lngHeaderRow, 2) = False Then
            GoTo OutputHeader_Exit
        End If
        '罫線で囲む
        If BorderAround(objWb, strSheetName, lngWeekdayRow, 2, lngWeekdayRow, 2) = False Then
            GoTo OutputHeader_Exit
        End If
        '30回繰り返す
        For lngDayCount = 1 To 30
            '日
            strDate = CStr(lngYear) & "/" & Format(CStr(lngMonth), "00") & "/" & Format(CStr(lngDayCount), "00")
            '色
            lngDayColor = GetDayColor(strDate, arrHoliday)
            '列
            lngOutputCol = lngDayCount + 2
            'TODO：同じコードを複数回書いている
            '罫線で囲む
            If BorderAround(objWb, strSheetName, lngHeaderRow, lngOutputCol, lngHeaderRow, lngOutputCol) = False Then
                GoTo OutputHeader_Exit
            End If
            '色を塗る
            If DrawColor(objWb, strSheetName, lngHeaderRow, lngOutputCol, lngHeaderRow, lngOutputCol, lngDayColor) = False Then
                GoTo OutputHeader_Exit
            End If
            '日付を出力
            .Cells(lngHeaderRow, lngOutputCol) = CStr(lngDayCount)
            '罫線で囲む
            If BorderAround(objWb, strSheetName, lngWeekdayRow, lngOutputCol, lngWeekdayRow, lngOutputCol) = False Then
                GoTo OutputHeader_Exit
            End If
            '色を塗る
            If DrawColor(objWb, strSheetName, lngWeekdayRow, lngOutputCol, lngWeekdayRow, lngOutputCol, lngDayColor) = False Then
                GoTo OutputHeader_Exit
            End If
            '日付を出力
            .Cells(lngWeekdayRow, lngOutputCol) = GetWeekDay(strDate)
        Next lngDayCount
    End With
            
    OutputHeader = True
    
OutputHeader_Err:

OutputHeader_Exit:

End Function


'【概要】複数列を囲む
Public Function BorderAroundCols(ByVal objWb As Excel.Workbook, _
                        ByVal strSheetName As String, _
                        ByVal lngStartCol As Long, _
                        ByVal lngEndCol As Long, _
                        ByVal lngStartRow As Long, _
                        ByVal lngEndRow As Long) As Boolean
On Error GoTo BorderAroundCols_Err
    
    BorderAroundCols = False
   
    Dim lngCurrentCol As Long
    
    '最初の列から終わりの列まで繰り返す
    For lngCurrentCol = lngStartCol To lngEndCol
        '列を囲む
        If BorderAroundRow(objWb, strSheetName, lngCurrentCol, lngStartRow, lngEndRow) = False Then
            GoTo BorderAroundCols_Exit
        End If
    Next lngCurrentCol
            
    BorderAroundCols = True
    
BorderAroundCols_Err:

BorderAroundCols_Exit:

End Function


'【概要】複数列を囲む
Public Function OutputWorkShedule() As Boolean
On Error GoTo OutputWorkShedule_Err
    
    OutputWorkShedule = False
    
    Dim lngYear As Long
    Dim lngMonth As Long
    Dim lngHeaderRow As Long
    Dim strSheetName As String
    Dim objWb As Excel.Workbook
    Dim lngStartCol As Long
    Dim lngEndCol As Long
    Dim lngStartRow As Long
    Dim lngEndRow As Long
    Dim arrName() As Variant
    Dim arrJob() As Variant

    '氏名
    arrName = GetRowData(NAME_COL, DATA_SHEET_NAME)
    '役割
    arrJob = GetRowData(JOB_COL, DATA_SHEET_NAME)
    '祝日
    arrHoliday = GetRowData(HOLIDAY_COL, HOLIDAY_SHEET_NAME)
        
    'ブック追加
    Set objWb = Workbooks.Add
    
    'シート名をデータにする
    ActiveSheet.Name = DATA_SHEET_NAME
    'シート名設定
    strSheetName = ActiveSheet.Name
    
    '年
    lngYear = 2022
    '月
    lngMonth = 5
    'ヘッダ行
    lngHeaderRow = 4
    '開始列
    lngStartCol = 1
    '終了列
    lngEndCol = 32
    '開始行
    lngStartRow = 6
    '終了行
    'TODO：　改善の余地あり
    lngEndRow = lngHeaderRow + UBound(arrName) + 2
    
    'ヘッダ作成
    If OutputHeader(objWb, strSheetName, lngYear, lngMonth, lngHeaderRow, arrHoliday) = False Then
        GoTo OutputWorkShedule_Exit
    End If
    
    '列描画
    If BorderAroundCols(objWb, strSheetName, lngStartCol, lngEndCol, lngStartRow, lngEndRow) = False Then
        GoTo OutputWorkShedule_Exit
    End If
    
    '氏名を出力
    If ArrToRow(objWb, strSheetName, arrName, lngStartRow, NAME_COL) = False Then
        GoTo OutputWorkShedule_Exit
    End If
    
    '役割を出力
    If ArrToRow(objWb, strSheetName, arrJob, lngStartRow, JOB_COL) = False Then
        GoTo OutputWorkShedule_Exit
    End If
            
    OutputWorkShedule = True
    
OutputWorkShedule_Err:

OutputWorkShedule_Exit:
    Set objWb = Nothing
End Function
