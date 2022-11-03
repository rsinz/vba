VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWindowController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


' ウィンドウ最大化用API
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwindow As Long, ByVal cmdshow As Long) As Long
Public Enum WindowSize
    Normal = 1
    Minimized
    Maximized
End Enum

' ウィンドウアクティブ用API
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


'アプリケーション ウィンドウクラス名
'Access                       OMain
'Excel                          XLMAIN
'Word                         OpusApp
'PowerPoint                 PP10FrameClass
'Internet Explorer 　　  IEFrame
'Visual Basic Editor 　　wndclass_desked_gsk
'エクスプローラ            ExploreWClass
'メモ帳                         Notepad
'ワードパッド               WordPadClass
'電卓                           SciCalc
'ペイント                    MSPaintApp


'<summary>
'  指定されたウィンドウ名がいくつ実行中か
'</summary>
Function ProcessCount(name As String) As Long

    Dim count_ As Long

    With CreateObject("WbemScripting.SWbemLocator").ConnectServer
        For Each obj_ In .ExecQuery("Select * From Win32_Process")
            If obj_.Caption Like "*" & name & "*" Then count_ = count_ + 1
        Next
    End With

    ProcessCount = count_
End Function

'<summary>
'  指定されたウィンドウ名のハンドル
'</summary>
Function WindowHandle(name As String) As Long

    With CreateObject("WbemScripting.SWbemLocator").ConnectServer
        For Each obj_ In .ExecQuery("Select * From Win32_Process")
            If obj_.Caption Like "*" & name & "*" Then WindowHandle = obj_.Handle: Exit Function
        Next
    End With

End Function

'<summary>
'  指定ウィンドウのアクティブ化。フォーカスも移る
'</summary>
Sub ActivateWindow(name As String)
    
    Dim hwnd As Long, ret As Long
    If Len(name) Then hwnd = FindWindow(vbNullString, name)
    res = SetForegroundWindow(hwnd)
'    DoEvents

End Sub

'<summary>
'  指定ウィンドウのアクティブ化。フォーカスも移る。曖昧な名前でもOK
'</summary>
Sub ActivateWindowNameLike(name As String)
    
    With CreateObject("WbemScripting.SWbemLocator").ConnectServer
        For Each obj_ In .ExecQuery("Select * From Win32_Process")
            If obj_.Caption Like "*" & name & "*" Then res = SetForegroundWindow(obj_.Handle): Exit Sub
        Next
    End With
    
End Sub

'<summary>
'  指定ウィンドウの最大・最小化
'</summary>
Sub ChangeWindowSize(name As String, winSize As WindowSize)
    
    Dim hwnd As Long, ret As Long
    If Len(name) Then hwnd = FindWindow(vbNullString, name)
    ret = ShowWindow(hwnd, winSize)
'    DoEvents

End Sub





















