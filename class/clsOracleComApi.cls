VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOracleComApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' �ŏI�X�V���F2018/09/20  ��rsinz

Dim adoCom As Object, adoRs As Object

'///////// index //////////
' OpenDB:           OracleDB�Ɛڑ�����
' CloseDB:           OracleDB�Ƃ̐ڑ���؂�
' Mode:                �V�X�e�����[�h�擾
' Download:         sql ���s(�Ԃ�l����)
' Commit:            sql ���s(�Ԃ�l�Ȃ�)
' Header:             �w��e�[�u���̃w�b�_�[��z��ŕԂ�
' ColumnTypeOf:  �w��e�[�u���̃w�b�_�[�̃f�[�^�^��Ԃ�
' ToSqlFormat:     SQL�p�̏����ɕϊ�����
' Exist:                �f�[�^�̑��݃`�F�b�N



'///////// module //////////

Private Sub Class_Initialize()
    Set adoCom = CreateObject("ADODB.connection")
    Set adoRs = CreateObject("ADODB.Recordset")
End Sub

Private Sub Class_Terminate()
    CloseDB
    Set adoRs = Nothing
    Set adoCom = Nothing
End Sub

'<summary>
'  OracleDB�Ɛڑ�����
'
'  [�ȗ���]�\�[�X�� : �f�[�^�\�[�X�� jtdb�ł��H
'  [�ȗ���]ID : �ڑ��pID
'  [�ȗ���]�p�X���[�h : �ڑ��p�p�X���[�h
'</summary>
Sub OpenDB(Optional �\�[�X�� As String = "NPAS_R", Optional ID As String = "pacs", Optional �p�X���[�h As String = "pacs")
    adoCom.Open "DSN=" & �\�[�X�� & ";UID=" & ID & ";PWD=" & �p�X���[�h & ";"
End Sub

'<summary>
'  OracleDB�Ƃ̐ڑ���؂�
'</summary>
Sub CloseDB()
    
    ' 0 = adStateClosed;
    ' 2 = adStateConnecting;
    ' 4 = adStateExecuting;
    ' 8 = adStateFetching;
    If (adoCom.State = 0) Or _
        ((adoCom.State And 2) = 2) Or _
        ((adoCom.State And 4) = 4) Or _
        ((adoCom.State And 8) = 8) Then Exit Sub
        
    adoCom.Close
End Sub

'<summary>
'  �G���[�܂Ƃ߁@0:NoError  ����ȊO:Error
'</summary>
Private Function ErrCheck() As Integer

    If (adoCom.State And 1) <> 1 Then      ' 1 = adStateOpen
        MsgBox "DB�Ɛڑ����Ă��܂���BOpenDB���\�b�h�����s���Ă��������B", , "clsOracleComApi"
        ErrCheck = -1
        Exit Function
    End If

End Function

'<summary>
'  �V�X�e�����[�h�擾
'</summary>
Function Mode() As Variant

    If ErrCheck <> 0 Then Exit Function
    
    Dim sql As String:
    sql = "SELECT * " & _
                "FROM ���[�h���e�[�u�� " & _
                "ORDER BY �^�p���[�h"
    
    Set adoRs = adoCom.Execute(sql)
    Mode = adoRs.GetRows()
    
End Function

'<summary>
'  �_�E�����[�h
'
'  SQL : SQL�\��
'</summary>
Function Download(sql As String) As Variant

    If ErrCheck <> 0 Then Exit Function

    Set adoRs = adoCom.Execute(sql)
    Download = adoRs.GetRows()

End Function

'<summary>
'  SQL ���s
'
'  SQL : SQL�\��
'</summary>
Sub Commit(sql As String)

    On Error GoTo ErrorHandler

    If ErrCheck <> 0 Then Exit Sub
    adoCom.BeginTrans

    adoCom.Execute (sql)
    adoCom.CommitTrans
    
    GoTo Finally
ErrorHandler:
    adoCom.RollbackTrans
Finally:

End Sub


'<summary>
'  �w��e�[�u���̃w�b�_�[��z��ŕԂ�
'
'  �e�[�u���� : �܂�܃e�[�u����
'</summary>
Function Header(�e�[�u���� As String) As Variant

    If ErrCheck <> 0 Then Exit Function

    Dim sql As String
    sql = "SELECT *" & _
               " FROM " & �e�[�u����

    Set adoRs = adoCom.Execute(sql)
    
    Dim h As Variant: ReDim h(0 To adoRs.fields.Count - 1)
    For i = 0 To adoRs.fields.Count - 1
        h(i) = adoRs.fields(i).name
    Next i

    Header = h
End Function

'<summary>
'  �w��e�[�u���̃w�b�_�[�̃f�[�^�^��Ԃ�
'
'  �e�[�u���� : �܂�܃e�[�u����
'</summary>
Function ColumnTypeOf(�e�[�u���� As String) As Variant

    If ErrCheck <> 0 Then Exit Function

    Dim sql As String
    sql = "SELECT DATA_TYPE " & _
            "FROM USER_TAB_COLUMNS " & _
            "WHERE TABLE_NAME = '" & �e�[�u���� & "' " & _
            "ORDER BY COLUMN_ID"

    
    Set adoRs = adoCom.Execute(sql)
    ColumnTypeOf = adoRs.GetRows()

End Function

'<summary>
'  SQL�p�̏����ɕϊ�����
'
'  �ϊ�����l : �����Ɏw�肳�ꂽ�l�̏�����ϊ����܂�
'  �f�[�^�^ : ���̌^�Ɋ�Â��ĕϊ����܂�
'</summary>
Function ToSqlFormat(�ϊ�����l As Variant, �f�[�^�^ As String) As String

    If IsEmpty(�ϊ�����l) Then ToSqlFormat = "NULL": Exit Function
        
    Select Case �f�[�^�^
        Case "NUMBER": ToSqlFormat = CStr(�ϊ�����l)
        Case "VARCHAR2": ToSqlFormat = "'" & CStr(�ϊ�����l) & "'"
        Case "DATE": ToSqlFormat = "TO_DATE('" & CStr(�ϊ�����l) & "','YYYY/MM/DD HH24:MI:SS')"
        Case Else: ToSqlFormat = CStr(�ϊ�����l)
    End Select
    
End Function

'<summary>
'  �f�[�^�̑��݃`�F�b�N
'
'  �e�[�u���� : �܂�܃e�[�u����
'  [�ȗ���] where : ������
'</summary>
Function Exist(�e�[�u���� As String, Optional where As String = "") As Boolean

    If ErrCheck <> 0 Then Exit Function

    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT * " & _
            "FROM " & �e�[�u���� & _
            IIf(where = "", "", " WHERE " & where)
    
    Set adoRs = adoCom.Execute(sql)
    Exist = Not IsNull(adoRs)
    
    GoTo Finally
ErrorHandler:
    Call Err.Raise(10000, "clsOracleComApi.Exist", "SQL�\���ɊԈႢ������܂�")
Finally:
    
End Function














