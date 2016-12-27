VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


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


Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const IDC_ARROW = 32512&
Private Const IDC_HAND = 32649&
Private Const IDC_IBEAM = 32513&
Private Const IDC_WAIT = 32514&
Private Const IDC_SIZEALL = 32646&
Private mCursorDefault As Long, mCursorHand As Long, mCursorI As Long, mCursorWait As Long, mCursorCross As Long


'///////// enumerated type //////////
Public Enum CursorType
    Default = 0: Hand = 1: Ibeam = 2: Wait = 3: Cross = 4
End Enum

Public Enum FormColors
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
    
    BLACK = 40
End Enum

'///////// module //////////

Private Sub Class_Initialize()
    mCursorDefault = LoadCursor(0&, IDC_ARROW)
    mCursorHand = LoadCursor(0&, IDC_HAND)
    mCursorI = LoadCursor(0&, IDC_IBEAM)
    mCursorWait = LoadCursor(0&, IDC_WAIT)
    mCursorCross = LoadCursor(0&, IDC_SIZEALL)
End Sub

Function NonTitleBar(objName As String) As Long

    Dim wnd As Long, formHeight As Double
    Dim uForm As Object
    
    For Each uf In VBA.UserForms
        If uf.Name = objName Then Set uForm = uf: Exit For
    Next
    
    formHeight = uForm.InsideHeight
    WindowFromAccessibleObject uForm, wnd
    
    SetWindowLong wnd, GWL_EXSTYLE, GetWindowLong(wnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME
    NonTitleBar = SetWindowLong(wnd, GWL_STYLE, GetWindowLong(wnd, GWL_STYLE) And Not WS_CAPTION)
    
    DrawMenuBar wnd
    uForm.Height = uForm.Height - uForm.InsideHeight + formHeight
End Function

Sub FormDrag(objName As String, ByVal Button As Integer)

    Dim hWnd As Long
    Dim uForm As Object
    
    For Each uf In VBA.UserForms
        If uf.Name = objName Then Set uForm = uf: Exit For
    Next
    
    If Button = 1 Then
        WindowFromAccessibleObject uForm, hWnd
        ReleaseCapture
        Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If

End Sub

Function ChangeCursor(cType As CursorType) As Long
    Select Case cType
        Case CursorType.Default: Call SetCursor(mCursorDefault)
        Case CursorType.Hand:    Call SetCursor(mCursorHand)
        Case CursorType.Ibeam:  Call SetCursor(mCursorI)
        Case CursorType.Wait:     Call SetCursor(mCursorWait)
        Case CursorType.Cross:    Call SetCursor(mCursorCross)
    End Select
End Function

Function GetColor(colors As FormColors) As Long
    Select Case colors
        Case FormColors.WHITE:                  GetColor = RGB(255, 255, 255)
        Case FormColors.GRAY01:                GetColor = RGB(50, 50, 50)
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
        
        Case FormColors.BLACK:                  GetColor = RGB(60, 60, 60)
    End Select
End Function













