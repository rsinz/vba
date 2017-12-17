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

Public Enum DateType
    Year_ = 0
    Month_
    Day_
    Hour_
    Minute_
    Second_
End Enum

'///////// let property //////////
Public Property Let Value(v As Variant)
Attribute Value.VB_UserMemId = 0
    value_ = CDate(v)
End Property

Public Property Let SetByYYYYMMDD(s As String)
    value_ = DateSerial(Left(s, 4), Mid(s, 5, 2), Right(s, 2))
End Property

'///////// get property //////////

' ���݃Z�b�g����Ă�����t�����̂܂ܕԂ�
Public Property Get Value() As Variant
    Value = value_
End Property

' �w��̃t�H�[�}�b�g�ɕϊ����ĕ�����ŕԂ�
Public Property Get ToString(StringFormat As String) As String
    ToString = format(value_, StringFormat)
End Property

' �悭�g���t�H�[�}�b�g�`���ɂ��ĕԂ�
Public Property Get ToUsualString() As String
    ToUsualString = ToString("yyyy/mm/dd(aaa) hh:mm")
End Property

' �悭�g���ۑ��t�@�C�����ɂ��ĕԂ�
Public Property Get NowStamp() As String
    NowStamp = format(Now, "yyyymmdd_hhmmss")
End Property

' ���t�̕������������o���ĕԂ�
Public Property Get ToDate(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    ToDate = DateSerial(Year(dt_), Month(dt_), Day(dt_))
End Property

' �����̕������������o���ĕԂ�
Public Property Get ToTime(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    ToTime = DateSerial(Hour(dt_), Minute(dt_), Second(dt_))
End Property

' �j���ɕϊ����ĕԂ�
Public Property Get ToWeekday(Optional dt As Variant = Empty) As String
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    ToWeekday = WeekdayName(Weekday(dt_), True)
End Property

' ������Ԃ�
Public Property Get NextMonth(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    NextMonth = Add(Month_, 1, dt_)
End Property

' �O����Ԃ�
Public Property Get PreviousMonth(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    PreviousMonth = Add(Month_, -1, dt_)
End Property

' ���T��Ԃ�
Public Property Get NextWeek(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    NextWeek = Add(Day_, 7, dt_)
End Property

' ��T��Ԃ�
Public Property Get PreviousWeek(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    PreviousWeek = Add(Day_, -7, dt_)
End Property

' �[�N����
Public Property Get IsLeapYear(Optional dt As Variant = Empty) As Boolean
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    IsLeapYear = IIf(Day(DateSerial(Year(dt_), 2, 29)) = 29, True, False)       ' 2/29��Day�֐��ɓ˂������29���Ԃ��Ă�����[�N�B1���Ԃ��Ă�����[�N����Ȃ��B
End Property


'///////// functions //////////

' BeginOfMonth:          �w�肳�ꂽ���t�̌��̍ŏ��̓��t��Ԃ�
' EndOfMonth:             �w�肳�ꂽ���t�̌��̍Ō�̓��t��Ԃ�
' BeginOfWeek:           �w�肳�ꂽ���t�̏T�̍ŏ��̓��t��Ԃ�
' EndOfWeek:              �w�肳�ꂽ���t�̏T�̍Ō�̓��t��Ԃ�
' Add:                          ���t�̉����Z
' Between:                 �Q�̓��t�̊Ԋu



'<summary>[Private]
'  DateType �𕶎���ɕϊ�����BDateAdd�p
'
'  d�F�񋓌^DateType ����I��
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
'  �w�肳�ꂽ���t�̌��̍ŏ��̓��t��Ԃ�
'
'  [�ȗ���] dt�F��ƂȂ���t
'</summary>
Function BeginOfMonth(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    BeginOfMonth = DateAdd("d", -Day(dt_) + 1, dt_)
End Function

'<summary>
'  �w�肳�ꂽ���t�̌��̍Ō�̓��t��Ԃ�
'
'  [�ȗ���] dt�F��ƂȂ���t
'</summary>
Function EndOfMonth(Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    EndOfMonth = DateAdd("d", -1, BeginOfMonth(DateAdd("m", 1, dt_)))
End Function

'<summary>
'  �w�肳�ꂽ���t�̏T�̍ŏ��̓��t��Ԃ�
'
'  [�ȗ���] dt�F��ƂȂ���t
'�@[�ȗ���] FirstDayOfWeek�F�T�n�߂̗j���B�f�t�H���g�œ��j
'</summary>
Function BeginOfWeek(Optional dt As Variant = Empty, Optional FirstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    BeginOfWeek = DateAdd("d", 1 - Weekday(dt_, FirstDayOfWeek), dt_)
End Function

'<summary>
'  �w�肳�ꂽ���t�̏T�̍Ō�̓��t��Ԃ�
'
'  [�ȗ���] dt�F��ƂȂ���t
'�@[�ȗ���] FirstDayOfWeek�F�T�n�߂̗j���B�f�t�H���g�œ��j
'</summary>
Function EndOfWeek(Optional dt As Variant = Empty, Optional FirstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    EndOfWeek = DateAdd("d", 7 - Weekday(dt_, FirstDayOfWeek), dt_)
End Function

'<summary>
'  ���t�̉����Z
'
'  type_ : ���Ƃ��b�Ƃ��v�Z�P�ʂ��w��
'  v : �ǂꂾ�������Z���邩
'  [�ȗ���] dt�F��ƂȂ���t
'</summary>
Function Add(type_ As DateType, v As Double, Optional dt As Variant = Empty) As Date
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt), value_, CDate(dt))
    Add = DateAdd(ToDateChar(type_), v, dt_)
End Function

'<summary>
'  �Q�̓��t�̊Ԋu
'
'  dt1 : �P�ڂ̓��t
'  [�ȗ���] dt2�F�Q�ڂ̓��t
'</summary>
Function Between(dt1 As Variant, Optional dt2 As Variant = Empty) As Double
    Dim dt_ As Date: dt_ = IIf(IsEmpty(dt2), value_, CDate(dt2))
    Between = Abs(dt_ - CDate(dt1))
End Function











