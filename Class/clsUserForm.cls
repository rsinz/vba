VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 最終更新日：2018/09/20  ＠rsinz


' この辺はForm Windowの枠消し&ウィンドウドラックのため
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Sub ReleaseCapture Lib "user32.dll" ()
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc.dll" (ByVal IAcessible As Object, ByRef hWnd As Long) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Const GWL_STYLE = (-16&)
Const GWL_EXSTYLE = (-20&)
Const WS_CAPTION = &HC00000
Const WS_EX_DLGMODALFRAME = &H1&

' マウスカーソルを変更するためのもの
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const IDC_ARROW = 32512&        '矢印カーソル
Private Const IDC_HAND = 32649&           'ハンドカーソル
Private Const IDC_IBEAM = 32513&          'Iカーソル
Private Const IDC_WAIT = 32514&            '砂時計カーソル
Private Const IDC_SIZEALL = 32646&        '4方向の矢印カーソル
Private mCursorDefault As Long, mCursorHand As Long, mCursorI As Long, mCursorWait As Long, mCursorCross As Long

' 待ち時間処理のため
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' 透過処理のため
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal classname As Any, ByVal winname As Any) As Long
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd&, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Const WS_EX_TOOLWINDOW = &H80
Const WS_EX_LAYERED = &H80000
Const LWA_COLORKEY = 1
Const LWA_ALPHA = 2



'///////// enumerated type //////////
Public Enum CursorType          ' ChangeCursorで使用。
    Default = 0: Hand = 1: Ibeam = 2: Wait = 3: Cross = 4
End Enum

Public Enum FormColors          ' GetColorで使用。
    WHITE = 0
    GRAY01 = 1
    GRAY02 = 2
    GRAY03 = 3
    GRAY04 = 4
    RED = 5
    
    LightRED = 10
    LightGREEN = 11
    LightBLUE = 12
    LightBROUN = 13
    LightPURPLE = 14
    
    DarkRED = 20
    DarkGREEN = 21
    DarkBLUE = 22
    DarkBROUN = 23
    
    ORANGE = 30
    YELLOW = 31
    MINT = 32
    ExcelGREEN = 33
    
    BLACK = 40
End Enum

Public Enum SpCharType          ' GetCharactorで使用。
    CheckBoxOff = 0
    CheckBoxOn = 1
    CheckBoxX = 2
    OptionOff = 3
    OptionOn = 4
    
'    ISeting = 10                       'Win8,10用
'    ICamera = 11
'    ISave = 12
'    IDelete = 13
'    ICalender = 14
'    IShere = 15
'    IStart = 16
'    IPause = 17
'    IStop = 18
    ICherryBlossom = 19
    ICircle = 20
    ICopyRight = 21
    IArrowUp = 22
    IArrowDown = 23
    IArrowRight = 24
    IArrowLeft = 25
End Enum



'///////// module list //////////

' NonTitleBar:            ユーザーフォームのタイトルバー非表示
' FormDrag:               ユーザーフォームのタイトルバー以外でドラッグ可能にする
' ChangeCursor:        カーソルの形状変更
' ToShortString:         完全パス文字列を省略して返す｡2つ上の親ディレクトリまで表示
' GetColor:                 色データの取得
' ToDarkColor:           引数に指定した色を濃くして返す
' GetCharactor:         特殊文字の取得
' BeepEx:                   効果音を鳴らす
' FadeIn:                   Frameのフェードイン
' FadeOut:                 Frameのフェードアウト
' ThisScreenWidth:     画面の幅
' ThisScreenHeight:    画面の高さ
' ViewDesktop:           デスクトップを表示する
' Delay :                    待ち時間をつくる



'///////// module //////////

Private Sub Class_Initialize()
    mCursorDefault = LoadCursor(0&, IDC_ARROW)   'マウスポインターを標準の矢印の形にするためのマウスカーソルを取得
    mCursorHand = LoadCursor(0&, IDC_HAND)         '手の形
    mCursorI = LoadCursor(0&, IDC_IBEAM)                'I の形
    mCursorWait = LoadCursor(0&, IDC_WAIT)            '砂時計の形
    mCursorCross = LoadCursor(0&, IDC_SIZEALL)      '上下左右矢印の形
End Sub

'<summary>
'  ユーザーフォームのタイトルバー非表示　返り値：0=失敗 0<>成功 変更前のウィンドウスタイルの値
'
'  objName: ユーザーフォームの名前　"MainForm"とか
'</summary>
Function NonTitleBar(objName As String) As Long

    Dim wnd As Long, formHeight As Double
    Dim uForm As Object
    
    For Each uf In VBA.UserForms
        If uf.name = objName Then Set uForm = uf: Exit For
    Next
    
    formHeight = uForm.InsideHeight
    WindowFromAccessibleObject uForm, wnd
    
    SetWindowLong wnd, GWL_EXSTYLE, GetWindowLong(wnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME
    NonTitleBar = SetWindowLong(wnd, GWL_STYLE, GetWindowLong(wnd, GWL_STYLE) And Not WS_CAPTION)
    
    DrawMenuBar wnd
    uForm.Height = uForm.Height - uForm.InsideHeight + formHeight
End Function

'<summary>
'  ユーザーフォームのタイトルバー以外でドラッグ可能にする
'
'  objName: ユーザーフォームの名前　"MainForm"とか
'  Button: MouseMoveイベントのButtonをそのまま渡す
'</summary>
Sub FormDrag(objName As String, ByVal Button As Integer)

    Dim hWnd As Long
    Dim uForm As Object
    
    For Each uf In VBA.UserForms
        If uf.name = objName Then Set uForm = uf: Exit For
    Next
    
    If Button = 1 Then
        WindowFromAccessibleObject uForm, hWnd
        ReleaseCapture
        Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If

End Sub

'<summary>
'  フォームの透過度を設定する
'
'  objName: ユーザーフォームの名前　"MainForm"とか
'  [省略可] alpha: 透過度　初期値: 192
'</summary>
Sub OnLayer(objName As String, Optional alpha As Long = 192)

    hWnd& = FindWindow("ThunderDFrame", objName)
    If hWnd& <> 0& Then
        SetWindowLong hWnd&, GWL_EXSTYLE, GetWindowLong(hWnd&, GWL_EXSTYLE) Or WS_EX_LAYERED
        SetLayeredWindowAttributes hWnd&, 0, alpha, LWA_ALPHA
    End If

End Sub

'<summary>
'  フォームをフェードインさせる
'
'  objName: ユーザーフォームの名前　"MainForm"とか
'</summary>
Sub FadeIn(objName As String)

    Dim i As Long
    For i = 0 To 255 Step 17
        Call OnLayer(objName, i)
        DoEvents
        Sleep 1
    Next
    
End Sub

'<summary>
'  フォームをフェードアウト
'
'  objName: ユーザーフォームの名前　"MainForm"とか
'</summary>
Sub FadeOut(objName As String)

    Dim i As Long
    For i = 255 To 0 Step -17
        Call OnLayer(objName, i)
        DoEvents
        Sleep 1
    Next
        
End Sub

'<summary>
'  カーソルの形状変更
'
'  cType: カーソルタイプ　列挙型から選んで
'</summary>
Function ChangeCursor(cType As CursorType) As Long
    Select Case cType
        Case CursorType.Default: Call SetCursor(mCursorDefault)
        Case CursorType.Hand:    Call SetCursor(mCursorHand)
        Case CursorType.Ibeam:  Call SetCursor(mCursorI)
        Case CursorType.Wait:     Call SetCursor(mCursorWait)
        Case CursorType.Cross:    Call SetCursor(mCursorCross)
    End Select
End Function

'<summary>
'  完全パス文字列を省略して返す。2つ上の親ディレクトリまで表示
'
'  path: 完全パス
'</summary>
Function ToShortString(ByVal path As String)

    Dim splited As Variant: splited = Split(path, "\")
    
    If UBound(splited) > 2 Then
        path = "..."
        For i = UBound(splited) - 2 To UBound(splited)
            path = path & "\" & splited(i)
        Next i
    End If
    
    ToShortString = path
End Function

'<summary>
'  色データの取得
'
'  colors: 色　列挙型から選んで
'</summary>
Function GetColor(colors As FormColors) As Long
    Select Case colors
        Case FormColors.WHITE:                  GetColor = RGB(255, 255, 255)
        Case FormColors.GRAY01:                GetColor = RGB(50, 50, 50) 'RGB(64, 64, 64)
        Case FormColors.GRAY02:                GetColor = RGB(192, 192, 192)
        Case FormColors.GRAY03:                GetColor = RGB(224, 224, 224)
        Case FormColors.GRAY04:                GetColor = RGB(238, 238, 238)
        Case FormColors.RED:                      GetColor = RGB(230, 0, 51)
        
        Case FormColors.LightRED:              GetColor = RGB(208, 87, 107)
        Case FormColors.LightGREEN:          GetColor = RGB(108, 187, 90)
        Case FormColors.LightBLUE:             GetColor = RGB(39, 146, 195)
        Case FormColors.LightBROUN:         GetColor = RGB(141, 100, 73)
        Case FormColors.LightPURPLE:         GetColor = RGB(145, 92, 139)
        
        Case FormColors.DarkRED:              GetColor = RGB(164, 55, 58)
        Case FormColors.DarkGREEN:          GetColor = RGB(0, 110, 84)
        Case FormColors.DarkBLUE:             GetColor = RGB(50, 66, 107)
        Case FormColors.DarkBROUN:          GetColor = RGB(99, 74, 71)
        
        Case FormColors.ORANGE:               GetColor = RGB(238, 121, 72)
        Case FormColors.YELLOW:                GetColor = RGB(199, 179, 112)
        Case FormColors.MINT:                    GetColor = RGB(0, 172, 154)
        Case FormColors.ExcelGREEN:         GetColor = RGB(34, 116, 71)
        
        Case FormColors.BLACK:                  GetColor = RGB(60, 60, 60)
    End Select
End Function

'<summary>
'  引数に指定した色を濃くして返す
'
'  colors: 色
'</summary>
Function ToDarkColor(color As Long) As Long

    Const range As Integer = 30
    Dim RED As Integer:        RED = (color And CLng("&H0000FF")) / (2 ^ 0)
    Dim green As Integer:    green = (color And CLng("&H00FF00")) / (2 ^ 8)
    Dim blue As Integer:       blue = (color And CLng("&HFF0000")) / (2 ^ 16)

    RED = IIf(RED - range >= 0, RED - range, 0)
    green = IIf(green - range >= 0, green - range, 0)
    blue = IIf(blue - range >= 0, blue - range, 0)
    
    ToDarkColor = RGB(RED, green, blue)
End Function

'<summary>
'  特殊文字の取得
'
'  chars: 文字　列挙型から選んで
'</summary>
Function GetCharactor(chars As SpCharType) As String
    Select Case chars
        Case SpCharType.CheckBoxOff:     GetCharactor = ChrW(&H2610)
        Case SpCharType.CheckBoxOn:     GetCharactor = ChrW(&H2611)
        Case SpCharType.CheckBoxX:       GetCharactor = ChrW(&H2612)
        Case SpCharType.OptionOff:          GetCharactor = ChrW(&H25CB)
        Case SpCharType.OptionOn:          GetCharactor = ChrW(&H25C9)

'        Case SpCharType.ISeting:              GetCharactor = ChrW(&HE115)
'        Case SpCharType.ICamera:            GetCharactor = ChrW(&HE114)
'        Case SpCharType.ISave:                GetCharactor = ChrW(&HE105)
'        Case SpCharType.IDelete:              GetCharactor = ChrW(&HE107)
'        Case SpCharType.ICalender:          GetCharactor = ChrW(&HE163)
'        Case SpCharType.IShere:               GetCharactor = ChrW(&HE2FD)
'        Case SpCharType.IStart:                GetCharactor = ChrW(&HE102)
'        Case SpCharType.IPause:               GetCharactor = ChrW(&HE103)
'        Case SpCharType.IStop:                 GetCharactor = ChrW(&HE004)
        Case SpCharType.ICherryBlossom: GetCharactor = ChrW(&H273F)
        Case SpCharType.ICircle:                GetCharactor = ChrW(&H26AB)
        Case SpCharType.ICopyRight:         GetCharactor = ChrW(&HA9)
        Case SpCharType.IArrowUp:            GetCharactor = ChrW(&H25B2)
        Case SpCharType.IArrowDown:       GetCharactor = ChrW(&H25BC)
        Case SpCharType.IArrowRight:        GetCharactor = ChrW(&H25B6)
        Case SpCharType.IArrowLeft:          GetCharactor = ChrW(&H25C0)
        
    End Select
End Function

'<summary>
'  画面の幅
'</summary>
Function ThisScreenWidth() As Long
'    Windows(ThisWorkbook.Name).WindowState = xlMaximized           ' win10/Excel2013 環境だと上２行
'    ThisScreenWidth2 = Windows(ThisWorkbook.Name).Width
    Application.WindowState = xlMaximized                                            ' win7/Excel2010 環境だと下２行
    ThisScreenWidth2 = Application.Width
End Function

'<summary>
'  画面の高さ
'</summary>
Function ThisScreenHeight() As Long
'    Windows(ThisWorkbook.Name).WindowState = xlMaximized           ' win10/Excel2013 環境だと上２行
'    ThisScreenHeight2 = Windows(ThisWorkbook.Name).Height
    Application.WindowState = xlMaximized                                            ' win7/Excel2010 環境だと下２行
    ThisScreenHeight2 = Application.Height
End Function

'<summary>
'  デスクトップを表示する
'</summary>
Sub ViewDesktop()
    Dim shell As Object: Set shell = CreateObject("Shell.Application")
    shell.MinimizeAll
    Set shell = Nothing
End Sub

'<summary>
'  待ち時間をつくる
'</summary>
Sub Delay(msec As Long)
    Sleep msec
End Sub













