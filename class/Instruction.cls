VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MenuCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private RootMenu As CommandBarPopup
Private EventHandlers As Collection
Private MenuMacroComponentFullName As String

Public Sub Init(tag As String, rootCaption As String, vbc As VBComponent)

    Call Me.RemoveMenu(tag)
    Set EventHandlers = New Collection
    Dim VBEMenuBar As CommandBar: Set VBEMenuBar = Application.VBE.CommandBars(1)
    
    With New FileSystemObject
        MenuMacroComponentFullName _
            = "'" & .GetFileName(vbc.Collection.Parent.fileName) & "'!" & vbc.name
    End With
    
    Set RootMenu = VBEMenuBar.Controls.Add(Type:=msoControlPopup)
    RootMenu.Caption = rootCaption
    RootMenu.tag = tag
End Sub

Public Sub AddPopMenu(pName As String, shortcut As String, border As Boolean)

    Dim SubMenu As CommandBarControl: Set SubMenu = RootMenu.Controls.Add(Type:=msoControlPopup)
        
    With SubMenu
        .Caption = Replace(pName, "_", "") & "(&" & shortcut & ")"
        .BeginGroup = border
    End With
    
End Sub

Public Sub AddChildMenu(pName As String, shortcut As String, border As Boolean)

    Dim sp As Variant: sp = Split(pName, "_")

    Dim index As Integer: Dim popName As String
    For i = 1 To RootMenu.Controls.Count        ' 親メニューのインデックスを探ります
        popName = RootMenu.Controls(i).Caption
        If Left(popName, InStr(popName, "(&") - 1) = sp(0) Then index = RootMenu.Controls(i).index
    Next i
    If index = 0 Then Exit Sub

    Dim PopMenu As CommandBarControl: Set PopMenu = RootMenu.Controls(index)
    Dim SubMenu As CommandBarControl: Set SubMenu = PopMenu.Controls.Add
    With SubMenu
        .Caption = sp(1) & "(&" & shortcut & ")"
        .BeginGroup = border
        .OnAction = MenuMacroComponentFullName & "." & pName
    End With
    
    With New EventHandler
        Set .MenuEvent = Application.VBE.events.CommandBarEvents(SubMenu)
        EventHandlers.Add .Self
    End With
    
End Sub

Public Sub AddSubMenu(pName As String, shortcut As String, border As Boolean)

    Dim SubMenu As CommandBarControl: Set SubMenu = RootMenu.Controls.Add
        
    With SubMenu
        .Caption = pName & "(&" & shortcut & ")"
        .BeginGroup = border
        .OnAction = MenuMacroComponentFullName & "." & pName
    End With

    With New EventHandler
        Set .MenuEvent = Application.VBE.events.CommandBarEvents(SubMenu)
        EventHandlers.Add .Self
    End With
    
End Sub

Public Sub RemoveMenu(MenuTag As String)
    'RootMenu.Deleteとする代わりに、わざわざMenuTagで検索して消すのは、
    '前回の異常終了で残ってしまったメニューも片づけるため。
    If MenuTag = "" Then Exit Sub
    Dim MyMenu As CommandBarControl: Set MyMenu _
        = Application.VBE.CommandBars.FindControl(tag:=MenuTag)
    Do Until MyMenu Is Nothing
        MyMenu.Delete
        Set MyMenu = Application.VBE.CommandBars.FindControl(tag:=MenuTag)
    Loop
    Set EventHandlers = Nothing
End Sub

