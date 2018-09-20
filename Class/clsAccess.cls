VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 2018/09/05 @rsinz

Private ado_ As Object

'///////// module list //////////

' OpenDB:       Access�Ɛڑ�����
' CloseDB:       Access�Ƃ̐ڑ���؂�
' Download:    �����Ŏw�肳�ꂽAccess�e�[�u�����_�E�����[�h����z��ɕϊ��



'///////// module //////////

Private Sub Class_Initialize()
    Set ado_ = CreateObject("ADODB.Connection")
End Sub

Private Sub Class_Terminate()
    CloseDB
    Set ado_ = Nothing
End Sub

'<summary>
'  ADODB.Connection��open����B
'
'�@Access�t�@�C���̊��S�p�X�FAccess�t�@�C���̃t���p�X
'</summary>
Sub OpenDB(ByVal Access�t�@�C���̊��S�p�X As String)
    
    If (ado_.State And 1) = 1 Then CloseDB     ' 1 = adStateOpen
    ado_.PROVIDER = "Microsoft.Ace.OLEDB.12.0"
    ado_.Open Access�t�@�C���̊��S�p�X
    
End Sub

'<summary>
'  ADODB.Connection��close����B
'</summary>
Sub CloseDB()

    ' 0 = adStateClosed;
    ' 2 = adStateConnecting;
    ' 4 = adStateExecuting;
    ' 8 = adStateFetching;
    If (ado_.State = 0) Or _
        ((ado_.State And 2) = 2) Or _
        ((ado_.State And 4) = 4) Or _
        ((ado_.State And 8) = 8) Then Exit Sub
    
    ado_.Close

End Sub

'<summary>
'�@�����Ŏw�肳�ꂽAccess�e�[�u�����_�E�����[�h���A�z��ɕϊ��B�G���[�R�[�h Null
'
'�@�e�[�u�����F�f�[�^�x�[�X���̃e�[�u����
'�@[�ȗ���]���o�����F�p�����[�^�z��BArray�֐��Ŏw��  Access�N�G���̃p�����[�^�ƃ����N����B�@�K��lNull
'</summary>
Function Download(ByVal �e�[�u���� As String, Optional ���o���� As Variant = Null) As Variant

    On Error GoTo ErrorHandler
    
    If (ado_.State And 1) <> 1 Then Err.Raise Number:=10000, Description:="DB�Ɛڑ����Ă��܂���BOpen���\�b�h�����s���Ă��������B"
    
    ' -- Access ����_�E�����[�h���܂�
    With CreateObject("ADODB.Command")
        .ActiveConnection = ado_
        .CommandText = �e�[�u����
        
        Dim arr As Variant
        If IsNull(���o����) Then
            arr = .Execute().GetRows()
        Else
            arr = .Execute(Parameters:=���o����).GetRows()
        End If
    End With
    
    ' -- Excel �ɓ\��t���₷���悤�A�s������ւ��܂�
    If HasNull(arr) Then arr = ReplaceNullTo("", arr)
    Download = WorksheetFunction.Transpose(arr)
    If ArrayRank(Download) = 1 Then Download = To2D(Download)       ' �z�񂪈ꎟ���̏ꍇ�A�񎟌��ɕϊ����Ă����ƌ�Ŋy
    
    GoTo Finally

ErrorHandler:
    Download = Empty
Finally:
End Function

'<summary>
'�@Private] �z���NULL�����邩���肷��BNULL�������Transpose�ŃG���[���o��̂ŁA����𗘗p����
'
'�@���肷��z��F���̂܂��
'</summary>
Private Function HasNull(���肷��z�� As Variant) As Boolean

    On Error GoTo TruePattern
    Dim v As Variant: v = WorksheetFunction.Transpose(���肷��z��)
    
    HasNull = False
    Exit Function

TruePattern:
    HasNull = True
End Function

'<summary>
'�@private] �z�����Null�������l�ɒu���B
'
'�@�u����̒l�F�u����̒l
'�@�z��F�Ώۂ̔z��
'</summary>
Private Function ReplaceNullTo(ByVal �u����̒l As Variant, ByVal �z�� As Variant) As Variant
    
    For i = LBound(�z��, 1) To UBound(�z��, 1)
        For j = LBound(�z��, 2) To UBound(�z��, 2)
            If IsNull(�z��(i, j)) Or IsEmpty(�z��(i, j)) Then �z��(i, j) = �u����̒l
        Next j
    Next i

    ReplaceNullTo = �z��

End Function

'<summary>
'  private] �z��̎�������Ԃ��B�C���f�b�N�X�O�̎�����UBound����ƃG���[���o��̂ŁA����𗘗p
'
'�@�z��F�������𒲂ׂ�z��
'</summary>
Private Function ArrayRank(�z�� As Variant) As Integer

    On Error Resume Next
    
    Do While Err.Number = 0
        i = i + 1
        errTest = UBound(�z��, i)
    Loop
    
    On Error GoTo 0
    ArrayRank = i - 1
End Function

'<summary>
'  private] 1�����z���2�����z��ɂ���B
'  1�����z���2�����z�񂪍�������ꍇ�Ɉ�l�̏������s�������Ƃ��Ƃ��Ɏg��
'
'  �z��F�ϊ��Ώ�
'</summary>
Private Function To2D(�z�� As Variant) As Variant

    Dim arr As Variant: ReDim arr(1 To 1, 1 To UBound(�z��))
    
    For i = 1 To UBound(�z��)
        If IsObject(�z��(i)) Then
            Set arr(1, i) = �z��(i)
        Else
            arr(1, i) = �z��(i)
        End If
    Next i

    To2D = arr
End Function





















