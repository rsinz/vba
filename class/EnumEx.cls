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

' �v�f(item)�����Ԗڂɓ����Ă��邩
Public Property Get IndexOf(v As Variant) As Long
    
    Dim index As Long: index = -1

    If list_.Count <> 0 Then
        For i = 1 To list_.Count
            If list_(i) = v Then index = i: Exit For
        Next i
    End If

    IndexOf = index
End Property

' �v�f����Ԃ�
Public Property Get Count() As Long
    Count = list_.Count
End Property

' ���X�g���Ɏw��l�����݂��邩���肷��
Public Property Get Exist(v As Variant) As Boolean
    Exist = (IndexOf(v) <> -1)
End Property

' �v�f�������Ă��邩���肷��
Public Property Get IsEmpty_() As Boolean
    IsEmpty_ = (list_.Count = 0)
End Property

'///////// functions //////////

Private Sub Class_Initialize()
    Set dictionary = CreateObject("Scripting.Dictionary")
End Sub

'<summary>
'  ���X�g�̏�����
'</summary>
Public Sub SetList(ParamArray list() As Variant)
    For Each v In list
        list_.Add (v)
    Next
End Sub

'<summary>
'  ���X�g���̗v�f��For each �őS������悤�ɂ���
'</summary>
Public Function NewEnum() As Variant ' IEnumVARIANT : OLEAutomation�Q�Ǝ��̂�
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = list_.[_NewEnum]
End Function

'<summary>
'  �v�f�̒ǉ�
'</summary>
Public Sub Add(v As Variant)
    list_.Add (v)
End Sub

'<summary>
'  �v�f�̍폜
'</summary>
Public Sub Remove(v As Variant)
    list_.Remove (IndexOf(v))
End Sub

'<summary>
'  �S�v�f�̍폜
'</summary>
Public Sub RemoveAll()
    For i = 1 To list_.Count
        list_.Remove (1)
    Next i
End Sub













