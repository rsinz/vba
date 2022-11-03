VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 202210/29 : sira


Private log_dir As String

'///////// module list //////////
' WriteLog : msg���e�L�X�g�t�@�C���ɕۑ��B�t�@�C���͓��t���Ƃɐ����B


'///////// module //////////

Private Sub Class_Initialize()
    log_dir = ThisWorkbook.path & "\LOG"
End Sub

Private Sub Class_Terminate()
End Sub

' [Let] �f�t�H���g�v���p�e�B
Public Property Let FolderPath(p As String)
Attribute FolderPath.VB_UserMemId = 0
    p = IIf(Right(p, 1) = "\", Left(p, Len(p) - 1), p)
    log_dir = p
End Property

' [Get] �f�t�H���g�v���p�e�B
Public Property Get FolderPath() As String
    FolderPath = log_dir
End Property

'<summary>
'  private] LOG�p�f�B���N�g�����݊m�F�B������������
'</summary>
Private Sub MakeDirectry()
    If Dir(log_dir, vbDirectory) = "" Then Call MkDir(log_dir)
End Sub

'<summary>
'  msg���e�L�X�g�t�@�C���ɕۑ��B�t�@�C���͓��t���Ƃɐ����B
'
'  msg: ���O���b�Z�[�W
'</summary>
Public Sub WriteText(msg As String)

    Call MakeDirectry
    Dim date_: date_ = format(Now, "yyyy-mm-dd")
        
    With CreateObject("Scripting.FileSystemObject")
        Dim file_: Set file_ = .OpenTextFile(log_dir & "\" & date_ & "_log.txt", 8, True)
    End With
    
    Dim time_: time_ = format(Now, "hh:mm:ss")
    Call file_.writeline(date_ & " " & time_ & vbNewLine & msg & vbNewLine)

    file_.Close
    Set file_ = Nothing
    
End Sub

Public Sub RemoveText(orverDays As Integer)

    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(log_dir).Files
            If (DateDiff("d", f.DateLastModified, Date) > orverDays) And (.GetExtensionName(f) = "txt") Then Call .DeleteFile(f)
        Next
    End With
    
End Sub


'<summary>
'  �w��̃Z���̍ŉ�����msg��ǉ�
'
'  msg: ���O���b�Z�[�W
'  r: ���O�\�̊�Z��
'</summary>
Public Sub WriteCell(msg As String, r As Range)

    With r.Parent
        Dim bottm: bottm = .Cells(.Rows.Count, r.column).End(xlUp).row + 1
        .Cells(bottm, r.column) = format(Now, "yyyy/mm/dd hh:mm:ss")
        .Cells(bottm, r.column + 1) = msg
    End With
    
End Sub








