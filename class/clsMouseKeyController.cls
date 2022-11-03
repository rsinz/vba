VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMouseKeyController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' 最終更新日：2018/10/19  ＠sira


'///////// memo //////////


'///////// property list //////////
' ManualEnable :        [Let] キー・マウス操作を受け付けるか　※実行終了時に必ずTRUEにする


'///////// module list //////////
' ClickLeft:                マウスで画面の任意の位置を左クリック
' ClickLeftDouble:      マウスで画面の任意の位置を左Wクリック
' ClickRight:              マウスで画面の任意の位置を右クリック
' ClickLeftDragDrop:  ドラッグ& ドロップ
' MousePosition :    【ピクセル値調査用】マウスで画面をクリックしたときの位置を、イミディエイトウィンドウに表示
' KeyIn:                    任意キー入力
' InputStr:                別アプリに文字入力
' CtrlC:                     Ctrl + C (コピー)
' CtrlV:                     Ctrl v(貼り付け)
' PrintScreenTo:        プリントスクリーン
' Iterator:                 キー繰り返し入力
' SetClipText:          クリップボードにテキストを設定
' GetClipText:          クリップボードのテキスト取得
' ClearClipBoard:     クリップボードクリア
' Sleeper:                指定時間待ち




' マウス操作API
Private Declare Sub mouse_event Lib "user32" ( _
    ByVal dwFlags As MouseEventFlag, _
    Optional ByVal dx As Long = 0, _
    Optional ByVal dy As Long = 0, _
    Optional ByVal dwDate As Long = 0, _
    Optional ByVal dwExtraInfo As Long = 0 _
)       ' dwDate : ホイールの移動
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

' マウス位置指定API
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long

' マウス位置取得API
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As stPoint) As Long
Private Type stPoint
    X As Long
    y As Long
End Type


' キーボード操作API
Private Declare PtrSafe Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, _
    Optional ByVal bScan As Byte = 0, _
    Optional ByVal dwFlags As Long = 0, _
    Optional ByVal dwExtraInfo As Long = 0 _
)

' マウス・キー状態取得API
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Long

' 待機用API
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

' キーコード列挙型　vbkeyだと対応できないキーがあることと、インテリセンスを与えた方がコーディングに優しいため
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

' コンスタント変数
Private Const KEYEVENT_UP = &H2

' 内部変数



' ---------------------------------------------------------------------

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Application.Interactive = True
End Sub


' [Let] キー・マウス操作を受け付けるか　※実行終了時に必ずTRUEにする
Public Property Let ManualEnable(b As Boolean)
    Application.Interactive = b
End Property


'<summary>
'  マウスで画面の任意の位置を左クリック
'</summary>
Public Sub ClickLeft(leftPos As Long, topPos As Long)
    Call SetCursorPos(leftPos, topPos)      ' 絶対座標まで移動
    Call mouse_event(mef_LEFTDOWN)
    Call mouse_event(mef_LEFTUP)
End Sub

'<summary>
'  マウスで画面の任意の位置を左Wクリック
'</summary>
Public Sub ClickLeftDouble(X As Long, y As Long)
    Call SetCursorPos(X, y)      ' 絶対座標まで移動
    Call mouse_event(mef_LEFTDOWN)
    Call mouse_event(mef_LEFTUP)
    Call mouse_event(mef_LEFTDOWN)
    Call mouse_event(mef_LEFTUP)
End Sub

'<summary>
'  マウスで画面の任意の位置を右クリック
'</summary>
Public Sub ClickRight(leftPos As Long, topPos As Long)
    Call SetCursorPos(leftPos, topPos)      ' 絶対座標まで移動
    Call mouse_event(mef_RIGHTDOWN)
    Call mouse_event(mef_RIGHTUP)
End Sub

'<summary>
'  ドラッグ＆ドロップ
'</summary>
Public Sub ClickLeftDragDrop(X As Long, y As Long, X_ As Long, Y_ As Long)
    
    Call SetCursorPos(X, y)      ' 絶対座標まで移動
    Call mouse_event(mef_LEFTDOWN)
    
    Dim step_ As Long: step_ = IIf(X < X_, 1, -1)
    For i = X To X_ Step step_      ' 1ピクセルずつ動かさないとダメみたい
        Call SetCursorPos(i, Y_)
    Next
    
    Call mouse_event(mef_LEFTUP)

End Sub

'<summary>
'  マウスホイール操作       微妙な挙動
'</summary>
'Public Sub Wheel(distance As Long)
'    Call mouse_event(mef_WHEEL, , , distance)
'End Sub

'<summary>
'  【ピクセル値調査用】マウスで画面をクリックしたときの位置を、イミディエイトウィンドウに表示
'  ※Escapeキーでループを抜けてください。
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
Private Function OutputMousePosition(p_ As stPoint, comment As String, Optional X As Long = 0, Optional y As Long = 0) As stPoint       ' 固定位置での連続出力防止
    
    Dim p_mem As stPoint: p_mem = p_
    Call GetCursorPos(p_)
    
    If Not ((p_.X = p_mem.X) And (p_.y = p_mem.y)) Then Debug.Print "Call ." & comment & "(" & p_.X - X & ", " & p_.y - y & ")"
    OutputMousePosition = p_

End Function

'<summary>
'  任意キー入力 :Enter = vbKeyReturn
'</summary>
Sub KeyIn(vbkey As vbKeysEnum)
    Call keybd_event(vbkey)
    Call keybd_event(vbkey, , KEYEVENT_UP)
    Call Sleep(10)
End Sub

'<summary>
'  複数キー入力(同時押し)
'
' keys : 入力したいキーをArray()で指定
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
'  複数キー入力(同時押し)
'
' keys : 入力したいキーをArray()で指定
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
'  別アプリに文字入力
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
'  Ctrl + C (コピー)
'</summary>
Sub CtrlC()
    Call keybd_event(vbKeyControl)
    Call keybd_event(vbKeyC)
    Call keybd_event(vbKeyControl, , KEYEVENT_UP)
    Call keybd_event(vbKeyC, , KEYEVENT_UP)
    DoEvents
End Sub

'<summary>
'  Ctrl + V (貼り付け)
'</summary>
Sub CtrlV()
    Call keybd_event(vbKeyControl)
    Call keybd_event(vbKeyV)
    Call keybd_event(vbKeyControl, , KEYEVENT_UP)
    Call keybd_event(vbKeyV, , KEYEVENT_UP)
    DoEvents
End Sub

'<summary>
'  プリントスクリーン
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
'  vbkey 繰り返し入力
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
'  クリップボードにテキストを設定
'
'　str：設定する文字列
'</summary>
Sub SetClipText(str As String)
'クリップボードに文字列を格納
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .text = str
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With
End Sub

'<summary>
'  クリップボードのテキスト取得
'</summary>
Function GetClipText()
    With New DataObject
        .GetFromClipboard
        GetClipText = .GetText
    End With
End Function

'<summary>
'  クリップボードクリア
'</summary>
Sub ClearClipBoard()
    With New DataObject
        .Clear
    End With
End Sub

'<summary>
'  指定時間待ち
'</summary>
Sub Sleeper(seconds As Long)
    Call Sleep(seconds * 1000)
    DoEvents
End Sub








