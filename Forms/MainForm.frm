VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Template Form"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14190
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Const AppVersion As String = "ver. 1.0"
Private clsForm As New clsUserForm
Private THEME As Long

'**********************************
' property
'**********************************
Property Let SetProgressMsg(msg As String)
    Me.progMsg.Caption = msg
    DoEvents
End Property

'**********************************
'user form
'**********************************
Private Sub UserForm_Initialize()
    
    On Error GoTo ErrHandler
    
    Static initCompleted As Boolean

    If initCompleted = False Then
    
        initCompleted = True
        
        THEME = clsForm.GetColor(DarkBLUE)         ' Choose theme colors
        clsForm.NonTitleBar Me.Name                      ' Set Flat style

        Call initFormSetting
        Call NormalizeSet
                
    End If
    
    GoTo Finally

ErrHandler:
    Call MsgBox(Err.Description, , "例外が発生しました。")

Finally:
End Sub

Private Sub UserForm_Terminate()
    Application.WindowState = xlMaximized
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call NormalizeSet
End Sub

'**********************************
'sub routine
'**********************************
Private Sub initFormSetting()

    Me.BorderColor = THEME
    
    Me.labelTitle.Top = 1
    Me.labelTitle.Left = 1
    Me.labelTitle.Width = Me.Width - 3
    Me.labelTitle.BackColor = THEME
    
    Me.btnClose.Top = 1
    Me.btnClose.Left = Me.labelTitle.Width - Me.btnClose.Width + 1
    
    Me.btnHelp.Top = 1
    Me.btnHelp.Left = Me.btnClose.Left - Me.btnHelp.Width - 5
    
    Me.version = AppVersion
    
    Me.progMsg.ForeColor = THEME
    
End Sub

Private Sub NormalizeSet()

    Me.btnClose.BackColor = THEME
    Me.btnClose.ForeColor = clsForm.GetColor(WHITE)
    
    Me.btnHelp.BackColor = THEME
    Me.btnHelp.ForeColor = clsForm.GetColor(WHITE)
    
    Me.btnExcute.SpecialEffect = fmSpecialEffectEtched
    
End Sub

'**********************************
'top label
'**********************************
Private Sub labelTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call NormalizeSet
    clsForm.ChangeCursor Cross
End Sub

Private Sub labelTitle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    clsForm.FormDrag Me.Name, Button
End Sub

'**********************************
'close button
'**********************************
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnClose.BackColor = clsForm.GetColor(RED)
    Me.btnClose.ForeColor = clsForm.GetColor(WHITE)
    clsForm.ChangeCursor Hand
End Sub

'**********************************
'help button
'**********************************
Private Sub btnHelp_Click()
    ' write your code.
End Sub

Private Sub btnHelp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnHelp.BackColor = clsForm.GetColor(WHITE)
    Me.btnHelp.ForeColor = THEME
    clsForm.ChangeCursor Hand
End Sub

'**********************************
'excute button
'**********************************
Private Sub btnExcute_Click()
    Me.btnExcute.SpecialEffect = fmSpecialEffectBump
End Sub

Private Sub btnExcute_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    clsForm.ChangeCursor Hand
End Sub

'**********************************
'bottom label
'**********************************
Private Sub labelBottom_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call NormalizeSet
End Sub










