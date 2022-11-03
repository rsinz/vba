VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTimeEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private value_ As Date
Private clock_ As Date

Public Enum DateType
    Year_ = 0
    Month_
    Day_
    Hour_
    Minute_
    Second_
End Enum

'///////// let property //////////
Public Property Let value(v As Variant)
Attribute value.VB_UserMemId = 0
    value_ = CDate(v)
End Property

Public Property Let SetByYYYYMMDD(s As String)
    value_ = DateSerial(Left(s, 4), Mid(s, 5, 2), Right(s, 2))
End Property

'///////// get property //////////

' 現在セットされている日付をそのまま返す
Public Property Get value() As Variant
    value = value_
End Property

' 指定のフォーマットに変換して文字列で返す
Public Property Get ToString(StringFormat As String) As String
    ToString = format(value_, StringFormat)
End Property

' よく使うフォーマット形式にして返す
Public Property Get ToUsualString() As String
    ToUsualString = ToString("yyyy/mm/dd(aaa) hh:mm")
End Property

' よく使う保存ファイル名にして返す
Public Property Get NowStamp() As String
    NowStamp = format(Now, "yyyymmdd_hhmmss")
End Property

' 日付の部分だけ抜き出して返す
Public Property Get ToDate(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    ToDate = DateSerial(Year(dt_), Month(dt_), Day(dt_))
End Property

' 時刻の部分だけ抜き出して返す
Public Property Get ToTime(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    ToTime = DateSerial(Hour(dt_), Minute(dt_), Second(dt_))
End Property

' 曜日に変換して返す
Public Property Get ToWeekday(Optional dt As Variant = Empty) As String
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    ToWeekday = WeekdayName(Weekday(dt_), True)
End Property

' 翌月を返す
Public Property Get NextMonth(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    NextMonth = Add(Month_, 1, dt_)
End Property

' 前月を返す
Public Property Get PreviousMonth(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    PreviousMonth = Add(Month_, -1, dt_)
End Property

' 翌週を返す
Public Property Get NextWeek(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    NextWeek = Add(Day_, 7, dt_)
End Property

' 先週を返す
Public Property Get PreviousWeek(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    PreviousWeek = Add(Day_, -7, dt_)
End Property

' 閏年判定
Public Property Get IsLeapYear(Optional dt As Variant = Empty) As Boolean
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    IsLeapYear = IIf(Day(DateSerial(Year(dt_), 2, 29)) = 29, True, False)       ' 2/29をDay関数に突っ込んで29が返ってきたら閏年。1が返ってきたら閏年じゃない。
End Property


'///////// functions //////////

' BeginOfMonth:          指定された日付の月の最初の日付を返す
' EndOfMonth:             指定された日付の月の最後の日付を返す
' BeginOfWeek:           指定された日付の週の最初の日付を返す
' EndOfWeek:              指定された日付の週の最後の日付を返す
' Add:                          日付の加減算
' Between:                 ２つの日付の間隔
' FirstShift:                  一直(早番)が何シフトかを返す
' ScWeekNumber:      SC週単位の何週目かを返す。日曜日で切り替えるようにしています。



'<summary>[Private]
'  DateType を文字列に変換する。DateAdd用
'
'  d：列挙型DateType から選ぶ
'</summary>
Private Function ToDateChar(d As DateType) As String
    Select Case d
        Case Year_: ToDateChar = "yyyy"
        Case Month_: ToDateChar = "m"
        Case Day_: ToDateChar = "d"
        Case Hour_: ToDateChar = "h"
        Case Minute_: ToDateChar = "n"
        Case Second_: ToDateChar = "s"
    End Select
End Function

'<summary>
'  指定された日付の月の最初の日付を返す
'
'  [省略可] dt：基準となる日付
'</summary>
Function BeginOfMonth(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    BeginOfMonth = DateAdd("d", -Day(dt_) + 1, dt_)
End Function

'<summary>
'  指定された日付の月の最後の日付を返す
'
'  [省略可] dt：基準となる日付
'</summary>
Function EndOfMonth(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    EndOfMonth = DateAdd("d", -1, BeginOfMonth(DateAdd("m", 1, dt_)))
End Function

'<summary>
'  指定された日付の週の最初の日付を返す
'
'  [省略可] dt：基準となる日付
'　[省略可] FirstDayOfWeek：週始めの曜日。デフォルトで日曜
'</summary>
Function BeginOfWeek(Optional dt As Variant = Empty, Optional FirstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    BeginOfWeek = DateAdd("d", 1 - Weekday(dt_, FirstDayOfWeek), dt_)
End Function

'<summary>
'  指定された日付の週の最後の日付を返す
'
'  [省略可] dt：基準となる日付
'　[省略可] FirstDayOfWeek：週始めの曜日。デフォルトで日曜
'</summary>
Function EndOfWeek(Optional dt As Variant = Empty, Optional FirstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    EndOfWeek = DateAdd("d", 7 - Weekday(dt_, FirstDayOfWeek), dt_)
End Function

'<summary>
'  日付の加減算
'
'  type_ : 日とか秒とか計算単位を指定
'  v : どれだけ加減算するか
'  [省略可] dt：基準となる日付
'</summary>
Function Add(type_ As DateType, v As Double, Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    Add = DateAdd(ToDateChar(type_), v, dt_)
End Function

'<summary>
'  ２つの日付の間隔
'
'  dt1 : １つ目の日付
'  [省略可] dt2：２つ目の日付
'</summary>
Function Between(dt1 As Variant, Optional dt2 As Variant = Empty) As Double
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt2), value_, CDate(dt2))
    Between = Abs(dt_ - CDate(dt1))
End Function

'<summary>
'  一直(早番)が何シフトかを返す
'
'  [省略可] dt：基準となる日付
'</summary>
Function FirstShift(Optional dt As Variant = Empty) As String

    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    Const baseDate As Date = #1/1/1900#                     ' Aシフトが一直で、かつ月曜日の日を基準にした
    
    If dt_ < baseDate Then Debug.Print "DateEx.FirstShift : 引数が不正です": Exit Function
    
    FirstShift = Choose((WorksheetFunction.RoundDown(Abs(dt_ - baseDate) / 7, 0)) Mod 3 + 1, _
                                    "A", "B", "C")
End Function

'<summary>
'  SC週単位の何週目かを返す。日曜日で切り替えるようにしています。
'
'  [省略可] dt：基準となる日付
'</summary>
Function ScWeekNumber(Optional dt As Variant = Empty) As Integer

    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    Dim dayOfFirst As Date: dayOfFirst = BeginOfMonth(dt_)
    
    Do While format(dayOfFirst, "aaa") <> "日"        ' 月初めの日曜日を調べる
        dayOfFirst = dayOfFirst + 1
    Loop
    
    If dayOfFirst <= dt_ Then
        ScWeekNumber = Application.WorksheetFunction.RoundDown((dt_ - dayOfFirst) / 7, 0) + 1
    Else
        ScWeekNumber = ScWeekNumber(dayOfFirst - 7)
    End If

End Function

' <summary>
' 引数の秒を、時間:分:秒の形式に変換
'
' sec : 変換する秒数
' </summary>
Public Function ToHourMinuteSecond(sec As Long) As String

    With Application.WorksheetFunction
        Dim h As Long: h = .RoundDown(sec / (24 * 60), 0)
        Dim m As Long: m = .RoundDown((sec - (h * 24 * 60)) / 60, 0)
        Dim s As Long: s = .RoundDown(sec - (h * 24 * 60) - (m * 60), 0)
    End With

    ToHourMinuteSecond = h & ":" & m & ":" & s

End Function












