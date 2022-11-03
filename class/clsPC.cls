VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' BatteryModuleとセットで使ってね
Private sps As SYSTEM_POWER_STATUS
Public Event BatteryIsDenger()
Private Computer As Object


Private Sub Class_Initialize()
    Set Computer = CreateObject("WScript.Network")
End Sub

'///////// get property //////////

Public Property Get Domain() As String
    Domain = Computer.Domain
End Property

Public Property Get UserName() As String
    UserName = Computer.UserName
End Property

Public Property Get ComputerName() As String
    ComputerName = Computer.ComputerName
End Property

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
    BatteryFullLifeTime = IIf(sps.BatteryFullLifeTime = -1, "不明", sps.BatteryFullLifeTime & "　秒")
End Property

' 電源に接続されていない場合、TRUE
Public Property Get IsOffLine() As Boolean
    IsOffLine = IIf(InStr(ACLineStatus, "Off_Line") <> 0, True, False)
End Property

' 電源に接続されている場合、TRUE
Public Property Get IsOnLine() As Boolean
    IsOnLine = IIf(InStr(ACLineStatus, "On_Line") <> 0, True, False)
End Property

' 電源容量が5%以下の場合、TRUE
Public Property Get IsDenger() As Boolean
    IsDenger = IIf(InStr(BatterFlag, "危険") <> 0, True, False)
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










