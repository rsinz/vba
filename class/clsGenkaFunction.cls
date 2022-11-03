VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGenkaFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 最終更新日：2016/02/16  ＠しら

Private Const MAX_RETIO As Single = 99.99
Private flow6500(1 To 3) As Variant, flow5000(1 To 3) As Variant, flow4000(1 To 3) As Variant
Private initFmdCompleted As Boolean



Private Sub Class_Initialize()

End Sub


'///////// index //////////

' ---------------------------------------------------------------
' 週間関連
' ---------------------------------------------------------------
' FirstShift:                指定日付の一直が何シフトかを返す
' ScWeekNumber:    指定日付がSC週単位の何週目かを返す

' ---------------------------------------------------------------
' 調和FMD関連
' ---------------------------------------------------------------
' initializeFmd:            調和FMD関連の初期値セット｡
' CountSameLot:        流し順配列の中に､いくつ同じロットが並んでいるかを返す｡上から順に検索をかける
' GetContainRatio:     スライスごとの拡散割合を返す




'///////// module //////////

' ---------------------------------------------------------------
' 週間関連
' ---------------------------------------------------------------
'<summary>
'  指定日付の一直が何シフトかを返す
'
'　dt：指定日付
'</summary>
Function FirstShift(dt As Date) As String

    Const baseDate As Date = #3/31/2014#            ' Aシフトが一直で、かつ月曜日の日を基準にした
    Dim pastDay As Long: pastDay = Abs(dt - baseDate)
    Dim pastWeek As Long: pastWeek = Application.WorksheetFunction.RoundDown(pastDay / 7, 0)
    
    Select Case pastWeek Mod 3
        Case 0: FirstShift = "A"
        Case 1: FirstShift = IIf(dt > baseDate, "B", "C")
        Case 2: FirstShift = IIf(dt > baseDate, "C", "B")
    End Select
    
End Function

'<summary>
'  指定日付がSC週単位の何週目かを返す
'
'　dt：指定日付
'</summary>
Function ScWeekNumber(dt As Date) As Integer

    Dim BeginOfMonth As Date: BeginOfMonth = DateAdd("d", -Day(dt) + 1, dt)
    
    Do While format(BeginOfMonth, "aaa") <> "日"        ' 月初めの日曜日を調べる
        BeginOfMonth = BeginOfMonth + 1
    Loop
    
    If BeginOfMonth <= dt Then
        ScWeekNumber = Application.WorksheetFunction.RoundDown((dt - BeginOfMonth) / 7, 0) + 1
    Else
        ScWeekNumber = ScWeekNumber(BeginOfMonth - 7)
    End If

End Function


' ---------------------------------------------------------------
' 調和FMD関連
' ---------------------------------------------------------------
'<summary>
'　調和FMD関連の初期値セット。
'</summary>
Private Sub initializeFmd()

    initFmdCompleted = True
    Dim retioArr(1 To 3, 1 To 5) As Variant      ' 行：ロット　列：スライス番号

    ' 流量6500kg/h
    retioArr(1, 1) = 44.06
    retioArr(1, 2) = 59.88
    retioArr(1, 3) = 70.91
    retioArr(1, 4) = 70.82
    retioArr(1, 5) = 62.56
    flow6500(1) = retioArr

    retioArr(1, 1) = 44.06
    retioArr(1, 2) = 59.88
    retioArr(1, 3) = 72.81
    retioArr(1, 4) = 82.53
    retioArr(1, 5) = 89.57
    retioArr(2, 1) = 94.7
    retioArr(2, 2) = 98.3
    retioArr(2, 3) = 98.09
    retioArr(2, 4) = 88.27
    retioArr(2, 5) = 72.97
    flow6500(2) = retioArr

    retioArr(1, 1) = 44.06
    retioArr(1, 2) = 59.88
    retioArr(1, 3) = 72.81
    retioArr(1, 4) = 82.53
    retioArr(1, 5) = 89.57
    retioArr(2, 1) = 94.7
    retioArr(2, 2) = 98.3
    retioArr(2, 3) = MAX_RETIO
    retioArr(2, 4) = MAX_RETIO
    retioArr(2, 5) = MAX_RETIO
    retioArr(3, 1) = MAX_RETIO
    retioArr(3, 2) = MAX_RETIO
    retioArr(3, 3) = 98.09
    retioArr(3, 4) = 88.27
    retioArr(3, 5) = 72.97
    flow6500(3) = retioArr

    ' 流量5000kg/h
    retioArr(1, 1) = 47.1
    retioArr(1, 2) = 67.34
    retioArr(1, 3) = 81.77
    retioArr(1, 4) = 85.69
    retioArr(1, 5) = 73.18
    flow5000(1) = retioArr

    retioArr(1, 1) = 47.1
    retioArr(1, 2) = 67.34
    retioArr(1, 3) = 81.77
    retioArr(1, 4) = 91.13
    retioArr(1, 5) = 97.18
    retioArr(2, 1) = MAX_RETIO
    retioArr(2, 2) = MAX_RETIO
    retioArr(2, 3) = MAX_RETIO
    retioArr(2, 4) = 94.54
    retioArr(2, 5) = 75.99
    flow5000(2) = retioArr

    retioArr(1, 1) = 47.1
    retioArr(1, 2) = 67.34
    retioArr(1, 3) = 81.77
    retioArr(1, 4) = 91.13
    retioArr(1, 5) = 97.18
    retioArr(2, 1) = MAX_RETIO
    retioArr(2, 2) = MAX_RETIO
    retioArr(2, 3) = MAX_RETIO
    retioArr(2, 4) = MAX_RETIO
    retioArr(2, 5) = MAX_RETIO
    retioArr(3, 1) = MAX_RETIO
    retioArr(3, 2) = MAX_RETIO
    retioArr(3, 3) = MAX_RETIO
    retioArr(3, 4) = 94.54
    retioArr(3, 5) = 75.99
    flow5000(3) = retioArr
    
    ' 流量4000kg/h
    retioArr(1, 1) = 48.6
    retioArr(1, 2) = 70.71
    retioArr(1, 3) = 85.36
    retioArr(1, 4) = 91.18
    retioArr(1, 5) = 76.84
    flow4000(1) = retioArr

    retioArr(1, 1) = 48.6
    retioArr(1, 2) = 70.71
    retioArr(1, 3) = 85.36
    retioArr(1, 4) = 94.3
    retioArr(1, 5) = 99.36
    retioArr(2, 1) = MAX_RETIO
    retioArr(2, 2) = MAX_RETIO
    retioArr(2, 3) = MAX_RETIO
    retioArr(2, 4) = 96.89
    retioArr(2, 5) = 77.48
    flow4000(2) = retioArr

    retioArr(1, 1) = 48.6
    retioArr(1, 2) = 70.71
    retioArr(1, 3) = 85.36
    retioArr(1, 4) = 94.3
    retioArr(1, 5) = 99.36
    retioArr(2, 1) = MAX_RETIO
    retioArr(2, 2) = MAX_RETIO
    retioArr(2, 3) = MAX_RETIO
    retioArr(2, 4) = MAX_RETIO
    retioArr(2, 5) = MAX_RETIO
    retioArr(3, 1) = MAX_RETIO
    retioArr(3, 2) = MAX_RETIO
    retioArr(3, 3) = MAX_RETIO
    retioArr(3, 4) = 96.89
    retioArr(3, 5) = 77.48
    flow4000(3) = retioArr
    
End Sub

'<summary>[private]
'  配列の次元数を返す
'
'　whatArray：次元数を調べる配列
'</summary>
Private Function GetArrayElementNumber(whatArray As Variant) As Integer
    On Error Resume Next
    Do While Err.Number = 0
        i = i + 1
        TempData = UBound(whatArray, i)
    Loop
    On Error GoTo 0
    GetArrayElementNumber = i - 1
End Function

'<summary>
'  流し順配列の中に、いくつ同じロットが並んでいるかを返す。上から順に検索をかける
'
'　whatArray：検索する配列。一次元配列
'  searchString：検索する文字列
'</summary>
Function CountSameLot(whatArray As Variant, searchString As String, Optional firstRow As Long = 0) As Long
    
    If GetArrayElementNumber(whatArray) > 1 Then MsgBox "CountSameLot: 引数whatArrayは一次元配列で指定してください": Exit Function

    Dim counter As Integer: counter = 0
    Dim onFlow As Boolean: onFlow = False
    
    For i = firstRow To UBound(whatArray)
        If CStr(whatArray(i)) <> searchString And onFlow Then
            Exit For
        ElseIf CStr(whatArray(i)) = searchString Then
            counter = counter + 1: onFlow = True
        End If
    Next i
    
    CountSameLot = counter
End Function

'<summary>
'  スライスごとの拡散割合を返す
'
'　flowSP：原料流量の目標値
'　lotCount：何ロット連続しているか
'</summary>
Function GetContainRatio(flowSP As Integer, lotCount As Integer)

    If initFmdCompleted = False Then initializeFmd
    If lotCount < 1 Then MsgBox "GetContainRatio: 連続ロット数は1以上を指定してください": Exit Function

    Dim difArr As Variant
    Dim returnArr As Variant: ReDim returnArr(1 To lotCount, 1 To 5)    ' 行：ロット　列：スライス番号
    

    If flowSP = 6500 Then
        If lotCount = 1 Then
            difArr = flow6500(1)
        ElseIf lotCount = 2 Then
            difArr = flow6500(2)
        Else
            difArr = flow6500(3)
        End If
    ElseIf flowSP = 5000 Then
        If lotCount = 1 Then
            difArr = flow5000(1)
        ElseIf lotCount = 2 Then
            difArr = flow5000(2)
        Else
            difArr = flow5000(3)
        End If
    ElseIf flowSP = 4500 Then
        If lotCount = 1 Then
            difArr = flow4000(1)
        ElseIf lotCount = 2 Then
            difArr = flow4000(2)
        Else
            difArr = flow4000(3)
        End If
    Else
        MsgBox "GetContainRatio: 目標流量が不明です": Exit Function
    End If

    For i = 1 To UBound(returnArr, 1)
        For j = 1 To UBound(returnArr, 2)
            If (lotCount > 3) And (i - 3 >= 0) And (i - lotCount <> 0) Then     ' 4ロット以上連続していた場合、頭3ロットとケツ1ロット以外の残り全てはMAX_RETIOになる
                returnArr(i, j) = MAX_RETIO
            ElseIf (lotCount > 3) And (i - lotCount = 0) Then           ' 4ロット以上連続していた場合のケツの処理
                returnArr(i, j) = difArr(3, j)
            Else
                returnArr(i, j) = difArr(i, j)
            End If
        Next j
    Next i

    GetContainRatio = returnArr

End Function
































