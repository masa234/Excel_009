'定数
Public Const DATA_SHEET_NAME = "DATA"
Public Const HOLIDAY_SHEET_NAME = "祝日"
Public Const CONFIRM = "確認"
Public Const NAME_COL = 1
Public Const JOB_COL = 2
Public Const HOLIDAY_COL = 2
'メッセージ
Public Const WORK_SHEDULE_OUTPUT_FAILED = "シフト表の出力に失敗しました。"


'【概要】配列内に存在するか？
Public Function InArr(ByVal arrSearch As Variant, _
                        ByVal strSearch As String) As Boolean
On Error GoTo InArr_Err
    
    InArr = False
    
    Dim lngArrIdx As Long
    
    '配列の最初から終端まで繰り返す
    For lngArrIdx = 0 To UBound(arrSearch)
        '配列の要素番号が検索値の場合、True
        If arrSearch(lngArrIdx) = strSearch Then
            InArr = True
        End If
    Next lngArrIdx
       
InArr_Err:

InArr_Exit:

End Function


'【概要】列を罫線で囲む
'TODO:　変数名改善の余地あり
Public Function BorderAroundRow(ByVal objWb As Excel.Workbook, _
                        ByVal strSheetName As String, _
                        ByVal lngDataCol As Long, _
                        ByVal lngStartRow As Long, _
                        ByVal lngEndRow As Long) As Boolean
On Error GoTo BorderAroundRow_Err
    
    BorderAroundRow = False
    
    Dim lngCurrentRow As Long
    
    '最初の行から終わりの行まで繰り返す
    For lngCurrentRow = lngStartRow To lngEndRow
        With objWb.Worksheets(strSheetName)
            '罫線を引く
            .Range(.Cells(lngCurrentRow, lngDataCol), .Cells(lngCurrentRow, lngDataCol)).BorderAround ColorIndex:=vbBlack, Weight:=xlThick
        End With
    Next
       
    BorderAroundRow = True
       
BorderAroundRow_Err:

BorderAroundRow_Exit:

End Function


'【概要】罫線で囲む
Public Function BorderAround(ByVal objWb As Excel.Workbook, _
                        ByVal strSheetName As String, _
                        ByVal lngStartRow As Long, _
                        ByVal lngStartCol As Long, _
                        ByVal lngEndRow As Long, _
                        ByVal lngEndCol As Long) As Boolean
On Error GoTo BorderAround_Err
    
    BorderAround = False
    
    With objWb.Worksheets(strSheetName)
        '罫線を引く
        .Range(.Cells(lngStartRow, lngStartCol), .Cells(lngEndRow, lngEndCol)).BorderAround ColorIndex:=vbBlack, Weight:=xlThick
    End With
       
    BorderAround = True
       
BorderAround_Err:

BorderAround_Exit:

End Function


'【概要】色を塗る
Public Function DrawColor(ByVal objWb As Excel.Workbook, _
                        ByVal strSheetName As String, _
                        ByVal lngStartRow As Long, _
                        ByVal lngStartCol As Long, _
                        ByVal lngEndRow As Long, _
                        ByVal lngEndCol As Long, _
                        ByVal lngDayColor As Long) As Boolean
On Error GoTo DrawColor_Err
    
    DrawColor = False
    
    With objWb.Worksheets(strSheetName)
        '色を設定する
        .Range(.Cells(lngStartRow, lngStartCol), .Cells(lngEndRow, lngEndCol)).Interior.Color = lngDayColor
    End With
       
    DrawColor = True
       
DrawColor_Err:

DrawColor_Exit:

End Function


'【概要】曜日を取得
Public Function GetWeekDay(ByVal strDate As String) As String
On Error GoTo GetWeekDay_Err
    
    '曜日毎に値を返却する
    Select Case Weekday(strDate)
    Case vbSunday
        GetWeekDay = "日"
    Case vbMonday
        GetWeekDay = "月"
    Case vbTuesday
        GetWeekDay = "火"
    Case vbWednesday
        GetWeekDay = "水"
    Case vbThursday
        GetWeekDay = "木"
    Case vbFriday
        GetWeekDay = "金"
    Case vbSaturday
        GetWeekDay = "土"
    End Select

       
GetWeekDay_Err:

GetWeekDay_Exit:

End Function


'【概要】列情報を配列にする
Public Function GetRowData(ByVal lngDataCol As Long, _
                            ByVal strSheetName As String) As Variant
On Error GoTo GetRowData_Err
        
    Dim lngLastRow As Long
    Dim lngCurrentRow As Long
    Dim lngArrIdx As Long
    Dim arrRet() As Variant
    
    With ThisWorkbook.Worksheets(strSheetName)
       '最終行取得
       lngLastRow = .Cells(1, 1).End(xlDown).Row
        '要素番号を初期化
        lngArrIdx = 0
        '最終行まで繰り返す
        For lngCurrentRow = 1 To lngLastRow
            '配列再宣言
            ReDim Preserve arrRet(lngArrIdx)
            '配列格納
            arrRet(lngArrIdx) = .Cells(lngCurrentRow, lngDataCol).Value
            '配列の要素番号を1つ進める
            lngArrIdx = lngArrIdx + 1
        Next lngCurrentRow
    End With
    
    GetRowData = arrRet
       
GetRowData_Err:

GetRowData_Exit:

End Function


'【概要】配列を列にする
Public Function ArrToRow(ByVal objWb As Excel.Workbook, _
                            ByVal strSheetName As String, _
                            ByVal arrOutput As Variant, _
                            ByVal lngStartRow As Long, _
                            ByVal lngOutputCol As Long) As Boolean
On Error GoTo ArrToRow_Err
    
    ArrToRow = False
        
    Dim lngCurrentRow As Long
    Dim lngArrIdx As Long
    
    '初期化
    lngCurrentRow = lngStartRow
    
    '配列の最初から終端まで繰り返す
    For lngArrIdx = 0 To UBound(arrOutput)
        '出力
        objWb.Worksheets(strSheetName).Cells(lngCurrentRow, lngOutputCol).Value = arrOutput(lngArrIdx)
        '行をカウントアップ
        lngCurrentRow = lngCurrentRow + 1
    Next lngArrIdx
    
    ArrToRow = True
       
ArrToRow_Err:

ArrToRow_Exit:

End Function

