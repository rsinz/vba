VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSQLServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 2022/10/29 @rsinz

Private ado_ As Object

'///////// module list //////////

' OpenDB:      SQLServer�Ɛڑ�����
' CloseDB:      SQLServer�Ƃ̐ڑ���؂�
' GetRows:     �����Ŏw�肳�ꂽSQLServer�e�[�u�����_�E�����[�h����z��ɕϊ��
' Insert:         �z��C���T�[�g
' Excute:        SQL���s


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
'�@SQLServer�t�@�C���̊��S�p�X�FSQLServer�t�@�C���̃t���p�X
'</summary>
Sub OpenDB(server As String, db As String, user As String, pass As String)
    
    If (ado_.State And 1) = 1 Then CloseDB     ' 1 = adStateOpen
    
    ado_.Open "Driver={SQL Server}; " & _
                        "server=" & server & ";" & _
                        "database=" & db & ";" & _
                        "uid=" & user & ";" & _
                        "pwd=" & pass & ";"
    
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
'�@�����Ŏw�肳�ꂽSQLServer�e�[�u�����_�E�����[�h���A�z��ɕϊ��B�G���[�R�[�h Null
'
'�@query_ : SQL�����w��
'</summary>
Function GetRows(query_ As String) As Variant

    On Error GoTo ErrorHandler
    
    If (ado_.State And 1) <> 1 Then Err.Raise Number:=10000, Description:="DB�Ɛڑ����Ă��܂���BOpen���\�b�h�����s���Ă��������B"
    
    ' -- SQLServer ����_�E�����[�h���܂�
    arr = ado_.Execute(query_).GetRows()
    
    ' -- Excel �ň����₷���悤�ɁA�s������ւ��܂�
    If HasNull(arr) Then arr = ReplaceNullTo("", arr)
    GetRows = WorksheetFunction.Transpose(arr)
    If ArrayRank(GetRows) = 1 Then GetRows = To2D(GetRows)       ' �z�񂪈ꎟ���̏ꍇ�A�񎟌��ɕϊ����Ă����ƌ�Ŋy
    
    GoTo Finally

ErrorHandler:
    GetRows = Empty
Finally:
End Function

'<summary>
'�@�z��C���T�[�g
'
'�@arr : �z��
'   table : �e�[�u����
'</summary>
Sub Insert(arr As Variant, table As String)
 
    Dim query_: query_ = "INSERT INTO " & table & " VALUES "
    
    For i = LBound(arr) To UBound(arr)
        Dim line_: line_ = ""
        
        For j = LBound(arr, 2) To UBound(arr, 2)
            If IsEmpty(arr(i, j)) Or arr(i, j) = "NULL" Then
                line_ = line_ & "NULL" & IIf(j = UBound(arr, 2), "", ",")
            Else
                line_ = line_ & "'" & arr(i, j) & "'" & IIf(j = UBound(arr, 2), "", ",")
            End If
        Next j
        query_ = query_ & "(" & line_ & ")" & ","
    
        If i Mod 100 = 0 Then
            Call ado_.Execute(Left(query_, Len(query_) - 1) & ";")
            query_ = "INSERT INTO " & table & " VALUES "
        End If
    
    Next i
    
    Call ado_.Execute(Left(query_, Len(query_) - 1) & ";")
    
End Sub

'<summary>
'�@SQL���s
'</summary>
Sub Excute(query_ As String)
    Call ado_.Execute(query_)
End Sub

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





















