VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJwComApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' �ŏI�X�V���F2022/08/22  ������

Public jwCom As Object

Private Sub Class_Initialize()
'    Set jwCom = CreateObject("JwComApi.ocx")
    Set jwCom = New JwComApi
End Sub

Private Sub Class_Terminate()
    Set jwCom = Nothing
End Sub


'///////// index //////////
' OpenJw:           JoyWatcher �̃l�b�g���[�N�~�h���E�G�A�Ɛڑ�����
' CloseJw:           JoyWatcher �̃l�b�g���[�N�~�h���E�G�A�Ƃ̐ڑ�����������
' SetTags:           �^�O�̈ꊇ�o�^
' ReadValue:      �o�^�σ^�O�̒l��P�̂œǂݍ���
' ReadValues:     �o�^�σ^�O�̒l���ꊇ�œǂݍ���
' WriteValue:      �o�^�σ^�O�̒l��P�̂ŏ�������
' WriteValues:    �o�^�σ^�O�̒l���ꊇ�ŏ�������
' ResetAllTags:   �o�^�σ^�O��S�ă��Z�b�g����
' StartTimer:      �w�莞�Ԃ̃^�C�}�[�C�x���g���J�n����
' StopTimer:       �w�莞�Ԃ̃^�C�}�[�C�x���g���~����
' ViewDialog:     �^�O�I���_�C�A���O�̕\��


'///////// module //////////

'<summary>
'  JoyWatcher �̃l�b�g���[�N�~�h���E�G�A�Ɛڑ�����
'</summary>
Sub OpenJw()
    jwCom.Open
End Sub

'<summary>
'  JoyWatcher �̃l�b�g���[�N�~�h���E�G�A�Ƃ̐ڑ�����������
'</summary>
Sub CloseJw()
    jwCom.Close
End Sub

'<summary>
'  �^�O�̈ꊇ�o�^
'
' list : �o�^����^�O���X�g vbCrLf�ŋ�؂�
' ��) "LOCAL$D.D0001$VALUE" + vbCrLf + "LOCAL$D.D0002$VALUE" + vbCrLf
'</summary>
Sub SetTags(list As Variant)

    If VarType(list) = vbString Then
        
        Call ResetAllTags
        jwCom.PLCTagName(0) = Replace(Replace(list, "[", ""), "]", "")
    
    ElseIf IsArray(list) Then
    
        Call ResetAllTags
        
        For i = LBound(list) To UBound(list)
            jwCom.PLCTagName(i - LBound(list)) = Replace(Replace(list(i), "[", ""), "]", "")
        Next i
    
    End If
    
End Sub

'<summary>
'  �o�^�σ^�O�̒l��P�̂œǂݍ���
'
' tagIndex : �o�^�σ^�O���X�g�̂O����n�܂�Y����
'</summary>
Function ReadValue(tagIndex As Long) As Variant
    
    If tagIndex < 0 Or tagIndex > jwCom.PLCTagCount - 1 Then Exit Function
    If jwCom.PLCTagCount = 0 Then Exit Function
    
    If jwCom.Active = False Then Call OpenJw
    
    ReadValue = jwCom.PLCTagValue(tagIndex)
    
End Function

'<summary>
'  �o�^�σ^�O�̒l���ꊇ�œǂݍ���
'</summary>
Function ReadValues(Optional list As Variant = Empty) As Variant
    
    If IsEmpty(list) Then
        If jwCom.PLCTagCount = 0 Then Exit Function
    Else
        Call SetTags(list)
    End If
        
    If jwCom.Active = False Then Call OpenJw
    
    Dim tagVal As Variant: ReDim tagVal(0 To jwCom.PLCTagCount - 1) As Variant
    For i = 0 To jwCom.PLCTagCount - 1
        tagVal(i) = jwCom.PLCTagValue(i)
    Next i
        
    ReadValues = tagVal
    
End Function

'<summary>
'  �o�^�σ^�O�̒l��P�̂ŏ�������
'
' tagIndex : �o�^�σ^�O���X�g�̂O����n�܂�Y����
' val : �������ޒl
'</summary>
Sub WriteValue(tagIndex As Long, Val As Variant)
    
    If tagIndex < 0 Or tagIndex > jwCom.PLCTagCount - 1 Then Exit Sub
    If jwCom.PLCTagCount = 0 Then Exit Sub
    
    If jwCom.Active = False Then Call OpenJw
    
    jwCom.PLCTagValue(tagIndex) = Val
    
End Sub

'<summary>
'  �o�^�σ^�O�̒l���ꊇ�ŏ�������
'
' val : �������ޒl(�z��) PLCTagValue�Ɠ����Y�����ɏ�������
'</summary>
Sub WriteValues(Val As Variant, Optional list As Variant = Empty)
            
    If IsEmpty(list) Then
        If jwCom.PLCTagCount = 0 Then Exit Sub
    Else
        Call SetTags(list)
    End If
    
    If jwCom.Active = False Then Call OpenJw
    
    For i = 0 To jwCom.PLCTagCount - 1
        jwCom.PLCTagValue(i) = Val(i + LBound(Val))
    Next i
        
End Sub

'<summary>
'  �o�^�σ^�O��S�ă��Z�b�g����
'</summary>
Sub ResetAllTags()
    jwCom.TagList = ""
End Sub

'<summary>
'  �w�莞�Ԃ̃^�C�}�[�C�x���g���J�n����
'
' timeValue : �C�x���g�����Ԋu�@�P��[msec]
'</summary>
Sub StartTimer(timeValue As Long)
    jwCom.TimerInterval = timeValue
    jwCom.TimerEnabled = True
End Sub

'<summary>
'  �w�莞�Ԃ̃^�C�}�[�C�x���g���~����
'</summary>
Sub StopTimer()
    jwCom.TimerEnabled = False
End Sub

'<summary>
'  �^�O�I���_�C�A���O�̕\���@�Ԃ�l�F�^�O���X�g
'�@���@�_�C�A���O��OK�{�^�����������ƂŁA�^�O���X�g�ɓo�^�����
'�@���@�O��l�͈����p���Ȃ����߁A�^�O���X�g�͖��񃊃Z�b�g�����
'</summary>
Function ViewDialog() As Variant
    jwCom.TagSelectDialog
    ViewDialog = jwCom.TagList
End Function

'///////// events //////////

'<summary>
'  �^�C�}�[�C�x���g�̓��e�@StartTimer�����s���邱�Ƃŏ���������
'</summary>
Private Sub JwCom_Timer()

    ' to do : write your code.
    
End Sub
















