VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private list_ As New Collection
Private dictionary As Scripting.dictionary



'///////// let property //////////
Public Property Let value(v As Variant, key As String)
    If dictionary.Exists(key) Then
        
    dictionary(key) = v
End Property

'///////// get property //////////
Public Property Get value(key As Variant) As Variant
Attribute value.VB_UserMemId = 0

'    If (TypeName(key) = "String") Or (TypeName(key) = "Integer") Or (TypeName(key) = "Long") Then Value = list_(key)
    Select Case TypeName(key)
        Case "String", "Integer", "Long": value = list_(key)
        Case Else: value = Empty
    End Select

'    If index = -1 Then
'        Set Value = list_
'    Else
'        Value = list_(index)
'    End If
End Property

' 要素(item)が何番目に入っているか
Public Property Get IndexOf(v As Variant) As Long
    
    Dim index As Long: index = -1

    If list_.Count <> 0 Then
        For i = 1 To list_.Count
            If list_(i) = v Then index = i: Exit For
        Next i
    End If

    IndexOf = index
End Property

' 要素数を返す
Public Property Get Count() As Long
    Count = list_.Count
End Property

' リスト内に指定値が存在するか判定する
Public Property Get Exist(v As Variant) As Boolean
    Exist = (IndexOf(v) <> -1)
End Property

' 要素が入っているか判定する
Public Property Get IsEmpty_() As Boolean
    IsEmpty_ = (list_.Count = 0)
End Property

'///////// functions //////////

Private Sub Class_Initialize()
    Set dictionary = CreateObject("Scripting.Dictionary")
End Sub

'<summary>
'  リストの初期化
'</summary>
Public Sub SetList(ParamArray list() As Variant)
    For Each v In list
        list_.Add (v)
    Next
End Sub

'<summary>
'  リスト内の要素をFor each で全部取れるようにする
'</summary>
Public Function NewEnum() As Variant ' IEnumVARIANT : OLEAutomation参照時のみ
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = list_.[_NewEnum]
End Function

'<summary>
'  要素の追加
'</summary>
Public Sub Add(v As Variant)
    list_.Add (v)
End Sub

'<summary>
'  要素の削除
'</summary>
Public Sub Remove(v As Variant)
    list_.Remove (IndexOf(v))
End Sub

'<summary>
'  全要素の削除
'</summary>
Public Sub RemoveAll()
    For i = 1 To list_.Count
        list_.Remove (1)
    Next i
End Sub














