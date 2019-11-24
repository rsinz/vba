Attribute VB_Name = "BatteryModule"

' clsBattery�ƃZ�b�g�Ŏg���Ă�


Private battery As New clsBattery
Private cycle As Date      ' (Application.ontime) �ݒ莞�ԋL������
Private usercycle As Long       ' (Application.ontime) ���[�U�[�w��̎������ԋL���̂���

' <summary>
' �d���̃��j�^�[���J�n����
'
' [�ȗ���] cycleSecond : �Ď������@/�f�t�H���g�F180�b
' </summary>
Public Sub StartMoniter(Optional cycleSecond As Long = 180)
    usercycle = cycleSecond
    Call MoniterLoop
End Sub

' <summary>
' �d���̃��j�^�[���~����
' </summary>
Public Sub StopMoniter()
    Call Application.OnTime(EarliestTime:=cycle, _
                                           procedure:="MoniterLoop", _
                                           Schedule:=False)
End Sub

' <summary>[private]
' Application.OnTime�𗘗p���Ē莞���Ƀ��\�b�h�����s����
' </summary>
Private Sub MoniterLoop()

    cycle = Now + timeValue(battery.ToHourMinuteSecond(usercycle))
    Call Application.OnTime(EarliestTime:=cycle, _
                                           procedure:="MoniterLoop")
    Call Moniter

End Sub

' <summary>[private]
' ��������\�b�h�̒��g
' </summary>
Private Sub Moniter()       ' todo : �����ɃC�x���g������
    Debug.Print format(Now, "hh:mm:ss")
    If battery.IsDenger Then
        Debug.Print "�o�b�e���[��5%�ȉ��ɂȂ��Ă��܂��B"
    End If
End Sub









