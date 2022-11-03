VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJwComApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 最終更新日：2022/08/22  ＠しら

Public jwCom As Object

Private Sub Class_Initialize()
'    Set jwCom = CreateObject("JwComApi.ocx")
    Set jwCom = New JwComApi
End Sub

Private Sub Class_Terminate()
    Set jwCom = Nothing
End Sub


'///////// index //////////
' OpenJw:           JoyWatcher のネットワークミドルウエアと接続する
' CloseJw:           JoyWatcher のネットワークミドルウエアとの接続を解除する
' SetTags:           タグの一括登録
' ReadValue:      登録済タグの値を単体で読み込み
' ReadValues:     登録済タグの値を一括で読み込み
' WriteValue:      登録済タグの値を単体で書き込み
' WriteValues:    登録済タグの値を一括で書き込み
' ResetAllTags:   登録済タグを全てリセットする
' StartTimer:      指定時間のタイマーイベントを開始する
' StopTimer:       指定時間のタイマーイベントを停止する
' ViewDialog:     タグ選択ダイアログの表示


'///////// module //////////

'<summary>
'  JoyWatcher のネットワークミドルウエアと接続する
'</summary>
Sub OpenJw()
    jwCom.Open
End Sub

'<summary>
'  JoyWatcher のネットワークミドルウエアとの接続を解除する
'</summary>
Sub CloseJw()
    jwCom.Close
End Sub

'<summary>
'  タグの一括登録
'
' list : 登録するタグリスト vbCrLfで区切る
' 例) "LOCAL$D.D0001$VALUE" + vbCrLf + "LOCAL$D.D0002$VALUE" + vbCrLf
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
'  登録済タグの値を単体で読み込み
'
' tagIndex : 登録済タグリストの０から始まる添え字
'</summary>
Function ReadValue(tagIndex As Long) As Variant
    
    If tagIndex < 0 Or tagIndex > jwCom.PLCTagCount - 1 Then Exit Function
    If jwCom.PLCTagCount = 0 Then Exit Function
    
    If jwCom.Active = False Then Call OpenJw
    
    ReadValue = jwCom.PLCTagValue(tagIndex)
    
End Function

'<summary>
'  登録済タグの値を一括で読み込み
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
'  登録済タグの値を単体で書き込み
'
' tagIndex : 登録済タグリストの０から始まる添え字
' val : 書き込む値
'</summary>
Sub WriteValue(tagIndex As Long, Val As Variant)
    
    If tagIndex < 0 Or tagIndex > jwCom.PLCTagCount - 1 Then Exit Sub
    If jwCom.PLCTagCount = 0 Then Exit Sub
    
    If jwCom.Active = False Then Call OpenJw
    
    jwCom.PLCTagValue(tagIndex) = Val
    
End Sub

'<summary>
'  登録済タグの値を一括で書き込み
'
' val : 書き込む値(配列) PLCTagValueと同じ添え字に書き込み
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
'  登録済タグを全てリセットする
'</summary>
Sub ResetAllTags()
    jwCom.TagList = ""
End Sub

'<summary>
'  指定時間のタイマーイベントを開始する
'
' timeValue : イベント発生間隔　単位[msec]
'</summary>
Sub StartTimer(timeValue As Long)
    jwCom.TimerInterval = timeValue
    jwCom.TimerEnabled = True
End Sub

'<summary>
'  指定時間のタイマーイベントを停止する
'</summary>
Sub StopTimer()
    jwCom.TimerEnabled = False
End Sub

'<summary>
'  タグ選択ダイアログの表示　返り値：タグリスト
'　※　ダイアログのOKボタンを押すことで、タグリストに登録される
'　※　前回値は引き継がないため、タグリストは毎回リセットされる
'</summary>
Function ViewDialog() As Variant
    jwCom.TagSelectDialog
    ViewDialog = jwCom.TagList
End Function

'///////// events //////////

'<summary>
'  タイマーイベントの内容　StartTimerを実行することで処理が走る
'</summary>
Private Sub JwCom_Timer()

    ' to do : write your code.
    
End Sub

















