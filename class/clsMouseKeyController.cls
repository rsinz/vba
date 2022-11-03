VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMouseKeyController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' �ŏI�X�V���F2018/10/19  ��sira


'///////// memo //////////


'///////// property list //////////
' ManualEnable :        [Let] �L�[�E�}�E�X������󂯕t���邩�@�����s�I�����ɕK��TRUE�ɂ���


'///////// module list //////////
' ClickLeft:                �}�E�X�ŉ�ʂ̔C�ӂ̈ʒu�����N���b�N
' ClickLeftDouble:      �}�E�X�ŉ�ʂ̔C�ӂ̈ʒu����W�N���b�N
' ClickRight:              �}�E�X�ŉ�ʂ̔C�ӂ̈ʒu���E�N���b�N
' ClickLeftDragDrop:  �h���b�O& �h���b�v
' MousePosition :    �y�s�N�Z���l�����p�z�}�E�X�ŉ�ʂ��N���b�N�����Ƃ��̈ʒu���A�C�~�f�B�G�C�g�E�B���h�E�ɕ\��
' KeyIn:                    �C�ӃL�[����
' InputStr:                �ʃA�v���ɕ�������
' CtrlC:                     Ctrl + C (�R�s�[)
' CtrlV:                     Ctrl v(�\��t��)
' PrintScreenTo:        �v�����g�X�N���[��
' Iterator:                 �L�[�J��Ԃ�����
' SetClipText:          �N���b�v�{�[�h�Ƀe�L�X�g��ݒ�
' GetClipText:          �N���b�v�{�[�h�̃e�L�X�g�擾
' ClearClipBoard:     �N���b�v�{�[�h�N���A
' Sleeper:                �w�莞�ԑ҂�




' �}�E�X����API
Private Declare Sub mouse_event Lib "user32" ( _
    ByVal dwFlags As MouseEventFlag, _
    Optional ByVal dx As Long = 0, _
    Optional ByVal dy As Long = 0, _
    Optional ByVal dwDate As Long = 0, _
    Optional ByVal dwExtraInfo As Long = 0 _
)       ' dwDate : �z�C�[���̈ړ�
Private Enum MouseEventFlag
    mef_MOVE = &H1              '  mouse move
    mef_LEFTDOWN = &H2      '  left button down
    mef_LEFTUP = &H4            '  left button up
    mef_RIGHTDOWN = &H8    '  right button down
    mef_RIGHTUP = &H10       '  right button up
    mef_MIDDLEDOWN = &H20 '  middle button down
    mef_MIDDLEUP = &H40      '  middle button up
    mef_WHEEL = &H800         ' mouse wheel
    mef_ABSOLUTE = &H8000& '  absolute move
End Enum

' �}�E�X�ʒu�w��API
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long

' �}�E�X�ʒu�擾API
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As stPoint) As Long
Private Type stPoint
    X As Long
    y As Long
End Type


' �L�[�{�[�h����API
Private Declare PtrSafe Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, _
    Optional ByVal bScan As Byte = 0, _
    Optional ByVal dwFlags As Long = 0, _
    Optional ByVal dwExtraInfo As Long = 0 _
)

' �}�E�X�E�L�[��Ԏ擾API
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Long

' �ҋ@�pAPI
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

' �L�[�R�[�h�񋓌^�@vbkey���ƑΉ��ł��Ȃ��L�[�����邱�ƂƁA�C���e���Z���X��^���������R�[�f�B���O�ɗD��������
Public Enum vbKeysEnum
    F1 = 112: F2 = 113: F3 = 114: F4 = 115: F5 = 116: F6 = 117: F7 = 118: F8 = 119: F9 = 120: F10 = 121: F11 = 122: F12 = 123: F13 = 124: F14 = 125: F15 = 126: F16 = 127
    num0 = 48: num1 = 49: num2 = 50: num3 = 51: num4 = 52: num5 = 53: num6 = 54: num7 = 55: num8 = 56: num9 = 57
    ten0 = 96: ten1 = 97: ten2 = 98: ten3 = 99: ten4 = 100: ten5 = 101: ten6 = 102: ten7 = 103: ten8 = 104: ten9 = 105
    MultiplyTenkey = 106: PlusTenkey = 107: EnterTenkey = 108: MinusTenkey = 109: PeriodTenkey = 110: SlashTenkey = 111
    A_ = 65: B_ = 66: c_ = 67: D_ = 68: E_ = 69: F_ = 70: G_ = 71: H_ = 72: I_ = 73: J_ = 74: K_ = 75: l_ = 76: M_ = 77
    n_ = 78: O_ = 79: p_ = 80: Q_ = 81: r_ = 82: s_ = 83: T_ = 84: U_ = 85: V_ = 86: W_ = 87: X_ = 88: Y_ = 89: Z_ = 90
    BackSpace = 8: Tab_ = 9: Enter = 13: Shift = 16: Ctrl = 17: Alt = 18: Pause = 19: Esc = 27: Conversion = 28: NoConversion = 29: Space = 32
    PgUp = 33: PgDown = 34: Home = 36: Left = 37: Up = 38: Right = 39: Down = 40: ins = 45: Del = 46: Win = 91: Popup = 93: NumLock = 144
    Colon = 186: SemiColon = 187: Comma = 188: Minus = 189: Period = 190: Slash = 191: Atmark = 192: QuotationStart = 219: Yenmark = 220: QuotationEnd = 221: Yamamark = 222: BackSlash = 226
End Enum

' �R���X�^���g�ϐ�
Private Const KEYEVENT_UP = &H2

' �����ϐ�



' ---------------------------------------------------------------------

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Application.Interactive = True
End Sub


' [Let] �L�[�E�}�E�X������󂯕t���邩�@�����s�I�����ɕK��TRUE�ɂ���
Public Property Let ManualEnable(b As Boolean)
    Application.Interactive = b
End Property


'<summary>
'  �}�E�X�ŉ�ʂ̔C�ӂ̈ʒu�����N���b�N
'</summary>
Public Sub ClickLeft(leftPos As Long, topPos As Long)
    Call SetCursorPos(leftPos, topPos)      ' ��΍��W�܂ňړ�
    Call mouse_event(mef_LEFTDOWN)
    Call mouse_event(mef_LEFTUP)
End Sub

'<summary>
'  �}�E�X�ŉ�ʂ̔C�ӂ̈ʒu����W�N���b�N
'</summary>
Public Sub ClickLeftDouble(X As Long, y As Long)
    Call SetCursorPos(X, y)      ' ��΍��W�܂ňړ�
    Call mouse_event(mef_LEFTDOWN)
    Call mouse_event(mef_LEFTUP)
    Call mouse_event(mef_LEFTDOWN)
    Call mouse_event(mef_LEFTUP)
End Sub

'<summary>
'  �}�E�X�ŉ�ʂ̔C�ӂ̈ʒu���E�N���b�N
'</summary>
Public Sub ClickRight(leftPos As Long, topPos As Long)
    Call SetCursorPos(leftPos, topPos)      ' ��΍��W�܂ňړ�
    Call mouse_event(mef_RIGHTDOWN)
    Call mouse_event(mef_RIGHTUP)
End Sub

'<summary>
'  �h���b�O���h���b�v
'</summary>
Public Sub ClickLeftDragDrop(X As Long, y As Long, X_ As Long, Y_ As Long)
    
    Call SetCursorPos(X, y)      ' ��΍��W�܂ňړ�
    Call mouse_event(mef_LEFTDOWN)
    
    Dim step_ As Long: step_ = IIf(X < X_, 1, -1)
    For i = X To X_ Step step_      ' 1�s�N�Z�����������Ȃ��ƃ_���݂���
        Call SetCursorPos(i, Y_)
    Next
    
    Call mouse_event(mef_LEFTUP)

End Sub

'<summary>
'  �}�E�X�z�C�[������       �����ȋ���
'</summary>
'Public Sub Wheel(distance As Long)
'    Call mouse_event(mef_WHEEL, , , distance)
'End Sub

'<summary>
'  �y�s�N�Z���l�����p�z�}�E�X�ŉ�ʂ��N���b�N�����Ƃ��̈ʒu���A�C�~�f�B�G�C�g�E�B���h�E�ɕ\��
'  ��Escape�L�[�Ń��[�v�𔲂��Ă��������B
'</summary>
Public Sub MousePosition(Optional X As Long = 0, Optional y As Long = 0)

    Const MSTATE_LBUTTON = &H1 '[LeftClick]
    Const MSTATE_RBUTTON = &H2 '[RightClick]
    Dim p_ As stPoint
    
    Do
        If GetAsyncKeyState(MSTATE_LBUTTON) <> 0 Then
            p_ = OutputMousePosition(p_, "ClickLeft", X, y)
        ElseIf GetAsyncKeyState(MSTATE_RBUTTON) <> 0 Then
            p_ = OutputMousePosition(p_, "ClickRight", X, y)
        End If
        
        DoEvents
    Loop Until GetAsyncKeyState(vbKeyEscape)
    
End Sub
Private Function OutputMousePosition(p_ As stPoint, comment As String, Optional X As Long = 0, Optional y As Long = 0) As stPoint       ' �Œ�ʒu�ł̘A���o�͖h�~
    
    Dim p_mem As stPoint: p_mem = p_
    Call GetCursorPos(p_)
    
    If Not ((p_.X = p_mem.X) And (p_.y = p_mem.y)) Then Debug.Print "Call ." & comment & "(" & p_.X - X & ", " & p_.y - y & ")"
    OutputMousePosition = p_

End Function

'<summary>
'  �C�ӃL�[���� :Enter = vbKeyReturn
'</summary>
Sub KeyIn(vbkey As vbKeysEnum)
    Call keybd_event(vbkey)
    Call keybd_event(vbkey, , KEYEVENT_UP)
    Call Sleep(10)
End Sub

'<summary>
'  �����L�[����(��������)
'
' keys : ���͂������L�[��Array()�Ŏw��
'</summary>
Sub KeysIn(keys As Variant)

    For i = LBound(keys) To UBound(keys)
        Call keybd_event(keys(i))
    Next i
    Call Sleep(10)
    
    For i = LBound(keys) To UBound(keys)
        Call keybd_event(keys(i), , KEYEVENT_UP)
    Next i
    Call Sleep(10)
    
End Sub

'<summary>
'  �����L�[����(��������)
'
' keys : ���͂������L�[��Array()�Ŏw��
'</summary>
Sub KeysInPram(key1 As vbKeysEnum, Optional key2 As vbKeysEnum = -1, Optional key3 As vbKeysEnum = -1, Optional key4 As vbKeysEnum = -1)

    Call keybd_event(key1)
    If key2 <> -1 Then Call keybd_event(key2)
    If key3 <> -1 Then Call keybd_event(key3)
    If key4 <> -1 Then Call keybd_event(key4)
    
    Call keybd_event(key1, , KEYEVENT_UP)
    If key2 <> -1 Then Call keybd_event(key2, , KEYEVENT_UP)
    If key3 <> -1 Then Call keybd_event(key3, , KEYEVENT_UP)
    If key4 <> -1 Then Call keybd_event(key4, , KEYEVENT_UP)
    
    Call Sleep(10)
End Sub

'<summary>
'  �ʃA�v���ɕ�������
'</summary>
Sub InputStr(s As String)
    Call SetClipText(s)
    DoEvents
    Call Sleep(50)
    Call CtrlV
    Call Sleep(10)
    DoEvents
End Sub

'<summary>
'  Ctrl + C (�R�s�[)
'</summary>
Sub CtrlC()
    Call keybd_event(vbKeyControl)
    Call keybd_event(vbKeyC)
    Call keybd_event(vbKeyControl, , KEYEVENT_UP)
    Call keybd_event(vbKeyC, , KEYEVENT_UP)
    DoEvents
End Sub

'<summary>
'  Ctrl + V (�\��t��)
'</summary>
Sub CtrlV()
    Call keybd_event(vbKeyControl)
    Call keybd_event(vbKeyV)
    Call keybd_event(vbKeyControl, , KEYEVENT_UP)
    Call keybd_event(vbKeyV, , KEYEVENT_UP)
    DoEvents
End Sub

'<summary>
'  �v�����g�X�N���[��
'</summary>
Sub PrintScreenTo(r As range)
    Call keybd_event(vbKeyMenu)     ' Alt
    Call keybd_event(vbKeySnapshot)
    Call keybd_event(vbKeyMenu, , KEYEVENT_UP)
    Call keybd_event(vbKeySnapshot, , KEYEVENT_UP)
    DoEvents
    r.PasteSpecial
End Sub

'<summary>
'  vbkey �J��Ԃ�����
'</summary>
Sub Iterator(vbkey As vbKeysEnum, time As Integer)

    If time <= 0 Then Exit Sub
    
    For i = 1 To time
        Call keybd_event(vbkey)
        Call keybd_event(vbkey, , KEYEVENT_UP)
        Call Sleep(10)
    Next i

End Sub

'<summary>
'  �N���b�v�{�[�h�Ƀe�L�X�g��ݒ�
'
'�@str�F�ݒ肷�镶����
'</summary>
Sub SetClipText(str As String)
'�N���b�v�{�[�h�ɕ�������i�[
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .text = str
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With
End Sub

'<summary>
'  �N���b�v�{�[�h�̃e�L�X�g�擾
'</summary>
Function GetClipText()
    With New DataObject
        .GetFromClipboard
        GetClipText = .GetText
    End With
End Function

'<summary>
'  �N���b�v�{�[�h�N���A
'</summary>
Sub ClearClipBoard()
    With New DataObject
        .Clear
    End With
End Sub

'<summary>
'  �w�莞�ԑ҂�
'</summary>
Sub Sleeper(seconds As Long)
    Call Sleep(seconds * 1000)
    DoEvents
End Sub







