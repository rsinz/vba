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

' �d���ɐڑ�����Ă��邩
Public Property Get ACLineStatus() As String

    GetSystemPowerStatus sps
    
    Select Case sps.ACLineStatus
        Case 0: ACLineStatus = "Off_Line"
        Case 1: ACLineStatus = "On_Line"
        Case Else: ACLineStatus = "�s��"
    End Select
    
End Property

' �o�b�e���[�e�ʂ���������m�肽����
Public Property Get BatterFlag() As String

    GetSystemPowerStatus sps
    
    Dim flags As String
    If (sps.BatteryFlag And 1) <> 0 Then flags = flags & "�����@"       ' 66%�𒴂���
    If (sps.BatteryFlag And 2) <> 0 Then flags = flags & "�Ⴂ�@"       ' 33%����
    If (sps.BatteryFlag And 4) <> 0 Then flags = flags & "�댯�@"       ' 5%����
    If (sps.BatteryFlag And 8) <> 0 Then flags = flags & "�`���[�W���@" ' �[�d��
    If (sps.BatteryFlag And 128) <> 0 Then flags = flags & "�o�b�e�������@"
    If sps.BatteryFlag = 255 Then flags = flags & "�s���@"
    
    BatterFlag = flags
    
End Property

' �o�b�e���[�e�ʂ̎c�p�[�Z���e�[�W
Public Property Get BatteryLifePercent() As Long
    GetSystemPowerStatus sps
    BatteryLifePercent = sps.BatteryLifePercent
End Property

' �o�b�e���[�e�ʂ̎c����
Public Property Get BatteryLifeTime() As String

    GetSystemPowerStatus sps
    
    Dim life As Long: life = sps.BatteryLifeTime
    BatteryLifeTime = IIf(life = -1, "�s��", ToHourMinuteSecond(life))
    
End Property

' �t���`���[�W����̌o�ߎ���
Public Property Get BatteryFullLifeTime() As String
    GetSystemPowerStatus sps
    BatteryFullLifeTime = IIf(sps.BatteryFullLifeTime = -1, "�s��", ToHourMinuteSecond(sps.BatteryFullLifeTime))
End Property

' �d���ɐڑ�����Ă��Ȃ��ꍇ�ATRUE
Public Property Get IsOffLine() As Boolean
    IsOffLine = (InStr(ACLineStatus, "Off_Line") <> 0)
End Property

' �d���ɐڑ�����Ă���ꍇ�ATRUE
Public Property Get IsOnLine() As Boolean
    IsOnLine = (InStr(ACLineStatus, "On_Line") <> 0)
End Property

' �d���e�ʂ�5%�ȉ��̏ꍇ�ATRUE
Public Property Get IsDenger() As Boolean
    IsDenger = (InStr(BatterFlag, "�댯") <> 0)
End Property

' <summary>
' �����̕b���A����:��:�b�̌`���ɕϊ�
'
' sec : �ϊ�����b��
' </summary>
Public Function ToHourMinuteSecond(sec As Long) As String

    With Application.WorksheetFunction
        Dim h As Long: h = .RoundDown(sec / (24 * 60), 0)
        Dim m As Long: m = .RoundDown((sec - (h * 24 * 60)) / 60, 0)
        Dim s As Long: s = .RoundDown(sec - (h * 24 * 60) - (m * 60), 0)
    End With

    ToHourMinuteSecond = h & ":" & m & ":" & s

End Function














