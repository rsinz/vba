VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 最終更新日：2018/02/21  ＠しら


'///////// memo //////////
'

'///////// module list //////////




Private table_ As ListObject


'///////// class property //////////
Public Property Let value(t As ListObject)
    Set table_ = t
End Property

Public Property Get value() As ListObject
Attribute value.VB_UserMemId = 0
    Set value = table_
End Property





'///////// module //////////

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub

'<summary>[private]
'
'</summary>
















