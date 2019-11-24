VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBattery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare PtrSafe Function GetSystemPowerStatus Lib "kernel32.dll" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long

Private Type SYSTEM_POWER_STATUS
    ACLineStatus As Byte
    BatteryFlag As Byte
    BatteryLifePercent As Byte
    Reserved1 As Byte
    BatteryLifeTime As Long
    BatteryFullLifeTime As Long
End Type

Private sps As SYSTEM_POWER_STATUS


'///////// get property //////////

' 電源に接続されているか
Public Property Get ACLineStatus() As String

    GetSystemPowerStatus sps
    
    Select Case sps.ACLineStatus
        Case 0: ACLineStatus = "Off_Line"
        Case 1: ACLineStatus = "On_Line"
        Case Else: ACLineStatus = "不明"
    End Select
    
End Property

' バッテリー容量をざっくり知りたい時
Public Property Get BatterFlag() As String

    GetSystemPowerStatus sps
    
    Dim flags As String
    If (sps.BatteryFlag And 1) <> 0 Then flags = flags & "高い　"       ' 66%を超える
    If (sps.BatteryFlag And 2) <> 0 Then flags = flags & "低い　"       ' 33%未満
    If (sps.BatteryFlag And 4) <> 0 Then flags = flags & "危険　"       ' 5%未満
    If (sps.BatteryFlag And 8) <> 0 Then flags = flags & "チャージ中　" ' 充電中
    If (sps.BatteryFlag And 128) <> 0 Then flags = flags & "バッテリ無し　"
    If sps.BatteryFlag = 255 Then flags = flags & "不明　"
    
    BatterFlag = flags
    
End Property

' バッテリー容量の残パーセンテージ
Public Property Get BatteryLifePercent() As Long
    GetSystemPowerStatus sps
    BatteryLifePercent = sps.BatteryLifePercent
End Property

' バッテリー容量の残時間
Public Property Get BatteryLifeTime() As String

    GetSystemPowerStatus sps
    
    Dim life As Long: life = sps.BatteryLifeTime
    BatteryLifeTime = IIf(life = -1, "不明", ToHourMinuteSecond(life))
    
End Property

' フルチャージからの経過時間
Public Property Get BatteryFullLifeTime() As String
    GetSystemPowerStatus sps
    BatteryFullLifeTime = IIf(sps.BatteryFullLifeTime = -1, "不明", ToHourMinuteSecond(sps.BatteryFullLifeTime))
End Property

' 電源に接続されていない場合、TRUE
Public Property Get IsOffLine() As Boolean
    IsOffLine = (InStr(ACLineStatus, "Off_Line") <> 0)
End Property

' 電源に接続されている場合、TRUE
Public Property Get IsOnLine() As Boolean
    IsOnLine = (InStr(ACLineStatus, "On_Line") <> 0)
End Property

' 電源容量が5%以下の場合、TRUE
Public Property Get IsDenger() As Boolean
    IsDenger = (InStr(BatterFlag, "危険") <> 0)
End Property

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















