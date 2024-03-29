VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



Sub Cyclic(procedure As String, sec As Long)

    Call Application.Run(procedure)
    Call Application.OnTime(Now + timeValue(ToHourMinuteSecond(sec)), "'Cyclic " & "" & procedure & """ & ", " & """ & sec & """ & '")

End Sub


' <summary>
' 引数の秒を、時間:分:秒の形式に変換
'
' sec : 変換する秒数
' </summary>
Private Function ToHourMinuteSecond(sec As Long) As String

    With Application.WorksheetFunction
        Dim h As Long: h = .RoundDown(sec / (24 * 60), 0)
        Dim m As Long: m = .RoundDown((sec - (h * 24 * 60)) / 60, 0)
        Dim s As Long: s = .RoundDown(sec - (h * 24 * 60) - (m * 60), 0)
    End With

    ToHourMinuteSecond = h & ":" & m & ":" & s

End Function
