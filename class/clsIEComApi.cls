VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIEComApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 最終更新日：2020/03/11  ＠しら

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private ie As Object

Public Enum ElementType
    None = 0
    ID = 1
    name = 2
    tagName = 3
    classname = 4
End Enum

Public Enum FamilyType
    parents = 1
    Children = 2
End Enum

Sub Class_Initialize()
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
End Sub


'///////// class property //////////

Property Let Visible(X As Boolean)
    ie.Visible = X
End Property

Property Get Visible() As Boolean
    Visible = ie.Visible
End Property

Property Get Document() As Boolean
    Document = ie.Document
End Property

'///////// index //////////

' GotoUrl:                          指定URLを開く
' GotoUrlNoWait:              指定URLを開く｡waitIEを挟むと無限ループするようなページ対策
' Quit:                               IEを閉じる
' ElementById:                IDから要素を取得｡
' ElementByName:          名前から要素を取得｡
' ElementByTag:              タグ名から要素を取得｡
' ElementByClass:            クラス名から要素を取得｡無かったらNULL｡
' SearchAtKeywords:        指定された要素の中からキーワードを含む要素を返す｡
' ElementByKeyword:        指定されたキーワードを含む要素を返す｡GetElementBy～とSearchAtKeywordsのラッピング
' ParentsElement:             指定された要素の親関係にある要素を返す
' ChildrenElement:            指定された要素の子関係にある要素を返す
' InputVal:                        プロパティを直接指定して値入力
' ClickSubmit:                   プロパティを直接指定してクリック
' RunJs:                            JavaScriptを実行
' SetCheckBoxState:        チェックボックスの状態をセット
' SetSelectBox:                セレクトボックス(コンボボックスもしくはリストボックス)を文言ベースで選択
' SetRadioBotton:            ラジオボタンを値ベースで選択
' WaitIE:                           IEの応答を待つ
' GetVersion:                    IEのバージョン取得
' DisableConfirm:             confirm()呼び出し時に確認ダイアログを表示させない



'///////// module //////////

'<summary>
'  指定URLを開く
'</summary>
Sub GotoUrl(URL As String)
    ie.Navigate URL
    WaitIE
End Sub

'<summary>
'  指定URLを開く。waitIEを挟むと無限ループするようなページ対策
'</summary>
Sub GotoUrlNoWait(URL As String)
    ie.Navigate URL
End Sub

'<summary>
'  IEを閉じる
'</summary>
Sub Quit()
    ie.Quit
End Sub

'<summary>
'  IDから要素を取得。無かったらNULL。複数あったら一番最初にでてきた要素を返す
'  ※旧IEのElementByIdはnameも参照する
'
'  id：要素のID
'</summary>
Function ElementById(ID As String)
    Set ElementById = ie.Document.ElementById(ID)
End Function

'<summary>
'  名前から要素を取得。無かったらNULL。
'
'  name：取得する要素の名前
'</summary>
Function ElementByName(name As String)
    Set ElementByName = ie.Document.ElementByName(name)
End Function

'<summary>
'  タグ名から要素を取得。無かったらNULL。
'
'  tagType：取得する要素種類　例："a" →<a>要素を全部取得
'</summary>
Function ElementByTag(tagType As String)
    Set ElementByTag = ie.Document.ElementByTagName(tagType)
End Function

'<summary>
'  クラス名から要素を取得。無かったらNULL。
'
'  className：取得するクラス名
'</summary>
Function ElementByClass(classname As String)
    Set ElementByClass = ie.Document.ElementByClassName(classname)
End Function

'<summary>
'  指定された要素の中からキーワードを含む要素を返す。
'
'  elements：要素
'  keywords：HTML内にkeywordsを含んでる(完全一致している)要素を返す
'</summary>
Function SearchAtKeywords(elements As Object, keywords As String) As Object
    For Each obj In elements
        If InStr(obj.outerHTML, keywords) > 0 Then
            Set SearchAtKeywords = obj: Exit For
        End If
    Next
End Function

'<summary>
'  指定されたキーワードを含む要素を返す。GetElementBy～とSearchAtKeywordsのラッピング
'
'  elmType：要素の種類　ElementTypeから選んで
'  property：要素の値　ID名とかNameとか
'  [省略可]keywords：HTML内にkeywordsを含んでる(完全一致している)要素を返す
'</summary>
Function ElementByKeyword(elmType As ElementType, property As String, Optional keywords As String = "") As Object

    Dim elements As Object
    
    Select Case elmType
        Case ElementType.ID
            Set elements = ElementById(property)
        Case ElementType.name
            Set elements = ElementByName(property)
        Case ElementType.tagName
            Set elements = ElementByTag(property)
        Case ElementType.classname
            Set elements = ElementByClass(property)
        Case Else
            Set elements = ie.Document.DocumentElement.all
    End Select

    Set ElementByKeyword = SearchAtKeywords(elements, keywords)
End Function

'<summary>
'  指定された要素の親関係にある要素を返す
'
'  elements：対象の要素
'  [省略可]hierarchy：階層。　例：2→2つ上の親要素を返す
'</summary>
Function ParentsElement(elements As Object, Optional hierarchy As Integer = 1) As Object

    Dim parentElements As Object: Set parentElements = elements.parentElement
    If hierarchy > 1 Then
        For i = 1 To hierarchy - 1
            Set parentElements = parentElements.parentElement
        Next i
    End If
    Set ParentsElement = parentElements
    
End Function

'<summary>
'  指定された要素の子関係にある要素を返す
'
'  elements：対象の要素
'  [省略可]hierarchy：階層。　例：2→2つ下の子要素を返す
'</summary>
Function ChildrenElement(elements As Object, Optional hierarchy As Integer = 1) As Object

    Dim childElements As Object: Set childElements = elements.Children
    If hierarchy > 1 Then
        For i = 1 To hierarchy - 1
            Set childElements = childElements.Children
        Next i
    End If
    Set ChildrenElement = childElements
    
End Function

'<summary>
'  プロパティを直接指定して値入力
'
'  element：HTML内の要素
'  val：入力する値
'</summary>
Sub SetInputVal(element As Object, Val)
    element.value = Val
    WaitIE
End Sub

'<summary>
'  プロパティを直接指定してクリック
'
'  element：HTML内の要素
'</summary>
Sub ClickSubmit(element As Object)
    element.Click
    WaitIE
End Sub

'<summary>
'  JavaScriptを実行
'
'  jsCode：JavaScript
'</summary>
Sub RunJs(jsCode As String)
    ie.Navigate "JavaScript:" & jsCode
End Sub

'<summary>
'  チェックボックスの状態をセット
'
'  element：HTML内の要素
'  [省略可]isChecked：チェックを入れるか外すか　デフォルト：入れる
'</summary>
Sub SetCheckBoxState(element As Object, Optional isChecked As Boolean = True)
    If Not (element.Checked = isChecked) Then element.Click
    WaitIE
End Sub

'<summary>
'  セレクトボックス(コンボボックスもしくはリストボックス)を文言ベースで選択
'
'  element：HTML内の要素
'  label：選択する文字列
'</summary>
Sub SetSelectBox(element As Object, label As String)
    
    If Len(label) < 1 Then Exit Sub
    
    With element
        For i = 0 To .options.Length - 1
            If .options(i).innerText = label Then
                .options(i).Selected = True: Exit For
            End If
        Next i
    End With
    
End Sub

'<summary>
'  ラジオボタンを値ベースで選択
'
'  element：HTML内の要素
'  value：選択する要素のvalue
'</summary>
Sub SetRadioBotton(element As Object, value As String)

    If Len(value) < 1 Then Exit Sub

    Dim i: For i = 0 To element.Length - 1
        If element(i).value = value Then
            element(i).Click
            Sleep 100
            Exit For
        End If
    Next i

End Sub

'<summary>
'  IEの応答を待つ
'</summary>
Sub WaitIE()
    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop
    Sleep 100
End Sub

'<summary>
'  IEの応答を待つ
'</summary>
Sub WaitUntilReady()
    Do Until ie.ReadyState < 4
        DoEvents
    Loop
    Sleep 100
End Sub

'<summary>
'  IEの応答を待つ
'</summary>
Function WaitUntilDisplayOf(target As String, etype As ElementType, Optional timeout As Integer = 15) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim element_ As Object
    
    Do Until Not element_ Is Null
        Set element_ = ElementByKeyword(etype, target)
        Sleep 100
        
        i = i + 1
        If i > timeout * 10 Then Call Err.Raise(10000, , "要求タイムアウト")
    Loop
    
    WaitUntilDisplayOf = True
    GoTo Finally
    
ErrHandler:
    Call MsgBox(Err.Description)

Finally:
    Set element_ = Nothing
    Sleep 100
    
End Function

'<summary>
'  IEのバージョン取得
'</summary>
Function GetVersion() As Integer
    Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
    GetVersion = Fix(Val(FS.GetFileVersion(ie.FullName)))
End Function

'<summary>
'  confirm()呼び出し時に確認ダイアログを表示させない
'</summary>
Sub DisableConfirm()

    Dim script: Set script = ie.Document.createElement("SCRIPT")
    script.Type = "text/javascript"
    script.Text = "function confirm() { return true; }"
    
    Call ie.Document.body.appendChild(script)
End Sub

















