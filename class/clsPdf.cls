VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' PDFDesigner Tools API
Private Declare PtrSafe Function GetPDFText Lib "pdftool.dll" (ByVal OpenFileName As String, ByVal SaveFileName As String) As Long


' <summary>
' PDF�t�@�C�����e�L�X�g�t�@�C���ɕϊ�
'  �߂�l  : 1:���� -1:���s -2:PDF�t�@�C�����Í�������Ă�
'
' pdfpath : PDF�t�@�C���t���p�X
' txtpath : txt�t�@�C���t���p�X
' </summary>
Public Function PDF2T(ByVal pdfpath As String, ByVal txtpath As String) As Long

    Call ChDir(ThisWorkbook.path)
    Dim tmp As String: tmp = Replace(pdfpath, ".pdf", ".tmp", compare:=vbTextCompare)
    Call ChangePDFVersion(pdfpath, tmp)    ' ����PDF�t�@�C�����R�s�[���ăo�[�W������1.4�`���ɕύX����
    
    PDF2T = GetPDFText(tmp, txtpath)
    Kill (tmp)
End Function

' <summary>
' PDF�̃o�[�W������1.4�`���ɂ���(�R�s�[�����t�@�C����ҏW)
'
' pdfpath : ��PDF�t�@�C���t���p�X
' tmppath : 1.4�`���ɂ�����̕ۑ���t���p�X
' </summary>
Private Sub ChangePDFVersion(ByVal pdfpath As String, ByVal tmppath As String)
    
    Dim Stream() As Byte: ReDim Stream(FileLen(pdfpath) - 1)
    
    ' PDF�ǂݍ���
    Dim FileNo As Integer: FileNo = FreeFile
    
    Open pdfpath For Binary As #FileNo
        Get #FileNo, , Stream
        Stream(5) = "&H31":   Stream(7) = "&H34"        ' PDF�̌`����1.4�ɂ���
    Close #FileNo
    
    ' PDF�o��
    FileNo = FreeFile
    Open tmppath For Binary Access Write As #FileNo
        Put #FileNo, , Stream
    Close #FileNo
    
End Sub

' <summary>
' �t�H���_���SPDF�e�L�X�g��
'
' [�ȗ���] InitFolderPath : �����̃t�H���_�p�X
' </summary>
Public Sub AllPDF2T(Optional InitFolderPath As String = "")

    '�t�H���_�p�X��I�����t�@�I���_���̃t�@�C���ꗗ���擾
    Dim target As String
    
    If InitFolderPath = "" Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            If .Show = True Then
                target = .SelectedItems(1)
            End If
        End With
    
    Else
        target = InitFolderPath
        
    End If
    
    Dim fopath As String: fopath = Dir(target & "\*.pdf")
    
    '�t�H���_���ɂ���PDF�t�@�C�����e�L�X�g�t�@�C���ɕϊ�
    Do While fopath <> ""
        Call PDF2T(fopath, Replace(fopath, "pdf", "txt"))
        fopath = Dir()
    Loop

End Sub
























