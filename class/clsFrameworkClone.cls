VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFrameworkClone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' �ŏI�X�V���F2016/06/23  ������

Private DOClipboard
Private RegExp As Object
Private mBuffer As Object, mIndex As Long

#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongPtr
    Private StartTime As LongPtr
    Private EndTime As LongPtr
    Private AccumulatedTime As LongPtr
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private StartTime As Long
    Private EndTime As Long
    Private AccumulatedTime As Long
#End If

Public Enum RegexOption
    IgnoreCase = &H80000000
    MultiLine = &H80000001
End Enum

Private Sub Class_Initialize()
    Set DOClipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    Set mBuffer = CreateObject("Scripting.Dictionary")
    mIndex = 0
    
End Sub

Private Sub Class_Terminate()
    Set DOClipboard = Nothing
    Set mBuffer = Nothing
End Sub


'///////// class property //////////

Property Get Elapsed()      ' stopwatch�ł̌o�ߎ���
    Elapsed = TimeSerial( _
                    AccumulatedTime / 1000 / 60 / 60, _
                    AccumulatedTime / 1000 / 60, _
                    AccumulatedTime / 1000)
End Property

Property Get ElapsedMilliseconds()
    ElapsedMilliseconds = AccumulatedTime
End Property


'///////// index //////////

' ---------------------------------------------------------------
' .NET Framework System.Windows.Forms.Clipboard�̃N���[��
' ---------------------------------------------------------------
' GetClipText:          �N���b�v�{�[�h�̃e�L�X�g�擾
' SetClipText:          �N���b�v�{�[�h�Ƀe�L�X�g��ݒ�
' ClearClipBoard:    �N���b�v�{�[�h�N���A

' ---------------------------------------------------------------
' .NET Framework System.IO.Directory�̃N���[��
' ---------------------------------------------------------------
' CreateDirectory:          �t�H���_�쐬
' DeleteDirectory:          �t�H���_�폜
' CopyDirectory:             �t�H���_�R�s�[
' ExistsDirectory:            �t�H���_�̑��݃`�F�b�N
' GetCurrentDirectory:   ���݃t�H�[�J�X����Ă���K�w��Ԃ�
' MoveDirectory:            �t�H���_�ړ�
' SetCurrentDirectory:   �w��p�X���t�H�[�J�X����
' GetFolderPath:            �t�@�C���̎Q�Ɛ�擾��Q�ƃ_�C�A���O�Ń��[�U�[�ɑI�����Ă��炤original
' SearchFile:                  �w��t�H���_�ɑ΂��ă��C���h�J�[�h�t�@�C���̌����������ΐ��K�t�@�C������Ԃ�original

' ---------------------------------------------------------------
' .NET Framework System.IO.File�̃N���[��
' ---------------------------------------------------------------
' AppendAllText:        �w�肵����������t�@�C���ɒǉ���t�@�C�����܂����݂��Ȃ��ꍇ�ͤ�t�@�C�����쐬�
' CopyFile:                 �t�@�C���̃R�s�[
' DeleteFile:               �t�@�C���̍폜
' ExistsFile:                 �t�@�C���̑��݃`�F�b�N
' MoveFile:                 �t�@�C���ړ�
' ReadAllText:            �t�@�C���ǂݍ��ݡ�ǂݍ��񂾃f�[�^�͕�����ŕԂ�
' ReadAllLines:           �t�@�C���ǂݍ��ݡ��������s���Ƃɔz��ŕԂ�
' WriteAllText:            �t�@�C���������ݡ��C�ɑS������
' WriteAllLines:          �t�@�C���������ݡ�������ݕ�������s���Ƃɔz��Ŏw�肵�����Ƃ���

' ---------------------------------------------------------------
' .NET Framework System.Text.RegularExpressions.Regex �̃N���[��
' ---------------------------------------------------------------
' InitRegex:            Regex�̏���������K�\�����g�������Ƃ��� (���[�U�[���͂̑Ó����`�F�b�N�Ƃ�)
' IsMatch:               �w�蕶����Ƀp�^�[�����܂܂�Ă��邩�ǂ���
' Match:                 �w�蕶������Ƀp�^�[�����܂܂�Ă������v�����ŏ��̕����񂾂��Ԃ�
' Matches:              �w�蕶������Ƀp�^�[�����܂܂�Ă�����item�R���N�V�����Ƃ��ĕԂ��
' Replace:              �w�蕶������Ƀp�^�[�����܂܂�Ă�����w�蕶����ɒu�������ĕԂ�

' ---------------------------------------------------------------
' .NET Framework System.Diagnostics.Stopwatch �̃N���[��
' ---------------------------------------------------------------
'StartTimer:        �^�C�}�[�̃X�^�[�g
'StopTimer:         �^�C�}�[�̃X�g�b�v �Ԃ�l�̓v���p�e�B����擾����
'ResetTimer:       �^�C�}�[�̌o�ߎ��ԃ��Z�b�g
'RestartTimer:    �o�ߎ��Ԃ����Z�b�g���¤�V���Ƀ^�C�}�[�X�^�[�g

' ---------------------------------------------------------------
' .NET Framework System.Text.StringBuilder�̃N���[��
' ---------------------------------------------------------------
' Append:               �������镶�����ݒ肷��
' AppendLine:        �������镶����+���s��ݒ肷��
' ClearDictionary:   �ݒ肵�Ă���������N���A����
' ToString:              �ݒ肵���������S�Č������ĕԂ�



'///////// module //////////

' ---------------------------------------------------------------
' .NET Framework System.Windows.Forms.Clipboard�̃N���[��
' ---------------------------------------------------------------

'<summary>
'  �N���b�v�{�[�h�̃e�L�X�g�擾
'</summary>
Function GetClipText()
    With DOClipboard
        .GetFromClipboard
        GetClipText = .GetClipText
    End With
End Function

'<summary>
'  �N���b�v�{�[�h�Ƀe�L�X�g��ݒ�
'
'�@str�F�ݒ肷�镶����
'</summary>
Sub SetClipText(str As String)
    With DOClipboard
        .SetText str
        .PutInClipboard
    End With
End Sub

'<summary>
'  �N���b�v�{�[�h�N���A
'</summary>
Sub ClearClipBoard()
    With DOClipboard
        .Clear
    End With
End Sub

' ---------------------------------------------------------------
' .NET Framework System.IO.Directory�̃N���[��
' ---------------------------------------------------------------

'<summary>
'  �t�H���_�쐬
'
'�@path�F�쐬����ꏊ�B���S�p�X�Ŏw��
'</summary>
Public Sub CreateDirectory(path As String)
    Call MkDir(path)
End Sub

'<summary>
'  �t�H���_�폜
'
'�@path�F�폜����t�H���_�̏ꏊ�B���S�p�X�Ŏw��
'  [�ȗ���]recursive�FTrue�œǂݎ���p�̑��������t�H���_���폜�B�f�t�H���g�FFalse
'</summary>
Public Sub DeleteDirectory(path As String, Optional recursive As Boolean = False)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.DeleteFolder(path, recursive)
End Sub

'<summary>
'  �t�H���_�R�s�[
'
'�@sourcePath�F�R�s�[���̃p�X
'�@destinationPath�F�R�s�[��̃p�X
'  [�ȗ���]recursive�FTrue�œǂݎ���p�̑��������t�H���_���R�s�[�ΏۂɁB�f�t�H���g�FFalse
'</summary>
Public Sub CopyDirectory(sourcePath As String, destinationPath As String, Optional recursive As Boolean = False)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.CopyFolder(sourcePath, destinationPath, recursive)
End Sub

'<summary>
'  �t�H���_�̑��݃`�F�b�N
'
'�@path�F�`�F�b�N����t�H���_�p�X
'</summary>
Public Function ExistsDirectory(path As String)
    ExistsDirectory = IIf(Dir(path, vbDirectory) = "", False, True)
End Function

'<summary>
'  ���݃t�H�[�J�X����Ă���K�w��Ԃ�
'</summary>
Public Function GetCurrentDirectory()
    GetCurrentDirectory = CurDir
End Function

'Public Function GetDirectories(path As String)
'    Err.Raise 513, "", "NotImplementedError"
'End Function
'
'Public Function GetFiles(path As String, Optional searchPattern As String = "*")
'    Err.Raise 513, "", "NotImplementedError"
'End Function

'<summary>
'  �t�H���_�ړ�
'
'�@sourcePath�F�ړ����̃p�X
'�@destinationPath�F�ړ���̃p�X
'</summary>
Public Sub MoveDirectory(sourcePath As String, destinationPath As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.moveFolder(sourcePath, destinationPath)
End Sub

'<summary>
'  �w��p�X���t�H�[�J�X����
'
'�@path�F�t�H�[�J�X����p�X
'</summary>
Public Sub SetCurrentDirectory(path As String)
    Call ChDrive(path)
    Call ChDir(path)
End Sub

'<summary>
'  �t�@�C���̎Q�Ɛ�擾�B�Q�ƃ_�C�A���O�Ń��[�U�[�ɑI�����Ă��炤
'</summary>
Function GetFolderPath() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            GetFolderPath = .SelectedItems(1)
        End If
    End With
End Function

'<summary>
'�@�w��t�H���_�ɑ΂��ă��C���h�J�[�h�t�@�C���̌����B������ΐ��K�t�@�C������Ԃ�
'
'�@folderPath�F�`�F�b�N�Ώۃt�H���_�̃p�X
'  filePath : �����t�@�C����
'</summary>
Function SearchFile(folderPath As String, filePath As String) As Variant
    SearchFile = Dir(folderPath & filePath, vbNormal)
End Function

' ---------------------------------------------------------------
' .NET Framework System.IO.File�̃N���[��
' ---------------------------------------------------------------

'<summary>
'  �w�肵����������t�@�C���ɒǉ��B�t�@�C�����܂����݂��Ȃ��ꍇ�́A�t�@�C�����쐬�B
'
'�@path�F�t�@�C���̊��S�p�X
'  str�F�ǉ����镶����
'</summary>
Public Sub AppendAllText(path As String, str As String)
    On Error GoTo SkipW
    Dim fileNum As Integer: fileNum = FreeFile
    
    Open path For Append As #fileNum
    Print #fileNum, str
SkipW:
    Close #fileNum
End Sub

'<summary>
'  �t�@�C���̃R�s�[
'
'�@from_�F�R�s�[���@���O���܂񂾃p�X
'  to_�F�R�s�[��@���O���܂񂾃p�X
'  [�ȗ���]overwrite�F�㏑�����邩�B�f�t�H���g�FFalse
'</summary>
Public Sub CopyFile(from_ As String, To_ As String, Optional overwrite As Boolean = False)
    If Not ExistsFile(To_) Or (ExistsFile(To_) And overwrite) Then
        Call FileCopy(from_, To_)
    End If
End Sub

'<summary>
'  �t�@�C���̍폜
'
'�@path�F�폜����t�@�C���̃p�X
'</summary>
Public Sub DeleteFile(path As String)
    Call Kill(path)
End Sub

'<summary>
'  �t�@�C���̑��݃`�F�b�N
'
'�@path�F�`�F�b�N����t�@�C���̃p�X
'</summary>
Public Function ExistsFile(path As String)
    With CreateObject("Scripting.FileSystemObject")
        ExistsFile = IIf(.FileExists(path), True, False)
    End With
End Function

'<summary>
'  �t�@�C���ړ�
'
'�@from_�F�ړ����@���O���܂񂾃p�X
'  to_�F�ړ���@���O���܂񂾃p�X
'</summary>
Public Sub MoveFile(from_ As String, To_ As String)
    With CreateObject("Scripting.FileSystemObject")
        Call .MoveFile(from_, To_)
    End With
End Sub

'<summary>
'  �t�@�C���ǂݍ��݁B�ǂݍ��񂾃f�[�^�͕�����ŕԂ�
'
'�@fileName�F�t�@�C�������܂񂾃p�X
'</summary>
Public Function ReadAllText(fileName As String) As String
    Dim fso As Object, buf As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    With fso.GetFile(fileName).OpenAsTextStream
        buf = .ReadAll
        .Close
    End With
    Set fso = Nothing
    
    ReadAllText = buf
End Function

'<summary>
'  �t�@�C���ǂݍ��݁B��������s���Ƃɔz��ŕԂ� �������ɕςȋ󔒕���������񂾂��ǒN��������
'
'�@fileName�F�t�@�C�������܂񂾃p�X
'</summary>
Public Function ReadAllLines(fileName As String) As Variant
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim data() As String: ReDim data(0)
    Dim i As Integer: i = 0
    
    With fso.GetFile(fileName).OpenAsTextStream
        Do Until .AtEndOfStream
            data(i) = .ReadLine
            i = i + 1
            ReDim Preserve data(i)
        Loop
        
        .Close
    End With
    Set fso = Nothing
    
    ReadAllLines = data
End Function

'<summary>
'  �t�@�C���������݁B��C�ɑS������
'
'�@path�F�t�@�C�������܂񂾃p�X
'  contents�F�������݂���������
'</summary>
Public Sub WriteAllText(path As String, contents As String)
    On Error GoTo SkipW
    Dim fileNum As Integer: fileNum = FreeFile
    
    Open path For Output As #fileNum
    Print #fileNum, contents
SkipW:
    Close #fileNum
End Sub

'<summary>
'  �t�@�C���������݁B�������ݕ�������s���Ƃɔz��Ŏw�肵�����Ƃ���
'
'�@path�F�t�@�C�������܂񂾃p�X
'  contents�F�������݂���������
'</summary>
Public Sub WriteAllLines(path As String, contents() As String)
    On Error GoTo SkipW
    Dim fileNum As Integer: fileNum = FreeFile
    
    Open path For Output As #fileNum
    Dim c: For Each c In contents
        Print #fileNum, c
    Next
SkipW:
    Close #fileNum
End Sub


' ---------------------------------------------------------------
' .NET Framework System.Text.RegularExpressions.Regex �̃N���[��
' ---------------------------------------------------------------

'<summary>
'  Regex�̏������B���K�\�����g�������Ƃ���(���[�U�[���͂̑Ó����`�F�b�N�Ƃ�)
'
'�@pattern�F����������p�^�[��
'  options�F�I�v�V����
'                   RegexOption.IgnoreCase�F�啶������������ʂ���
'                   RegexOption.MultiLine�F�����s��������
'
'�@****  �p�^�[�����  �Q�l�Fhttps://msdn.microsoft.com/ja-jp/library/ms974570.aspx
'�@�p���� �p���Ɛ����ɕ����ǂ���Ƀ}�b�`���܂��
'�@\n       ���s�Ƀ}�b�`���܂��B
'�@\f        �t�H�[�� �t�B�[�h�Ƀ}�b�`���܂��B
'�@\r       �L�����b�W ���^�[���Ƀ}�b�`���܂��B
'�@\t       �����^�u�Ƀ}�b�`���܂��B
'�@\v       �����^�u�Ƀ}�b�`���܂��B
'�@\?       ? �Ƀ}�b�`���܂��B
'�@\*       * �Ƀ}�b�`���܂��B
'�@\+      + �Ƀ}�b�`���܂��B
'�@\.       . �Ƀ}�b�`���܂��B
'�@\|       | �Ƀ}�b�`���܂��B
'�@\{       { �Ƀ}�b�`���܂��B
'�@\}       } �Ƀ}�b�`���܂��B
'�@\\       \ �Ƀ}�b�`���܂��B
'�@\[       [ �Ƀ}�b�`���܂��B
'�@\]       ] �Ƀ}�b�`���܂��B
'�@\(       ( �Ƀ}�b�`���܂��B
'�@\)       ) �Ƀ}�b�`���܂��B
'�@[xyz]  �����Z�b�g�Ɋ܂܂�Ă���C�ӂ�1�����Ƀ}�b�`���܂��
'�@\w      �P��Ɏg�p�����C�ӂ̕����Ƀ}�b�`���܂��B[a-zA-Z_0-9]�Ɠ����B
'�@\W     �P��Ɏg�p����镶���ȊO�̔C�ӂ̕����Ƀ}�b�`���܂��B[^a-zA-Z_0-9]�Ɠ����B
'�@\d      �C�ӂ̐����Ƀ}�b�`���܂��B[0-9]�Ɠ����B
'�@\D      �C�ӂ̐����ȊO�̕����Ƀ}�b�`���܂��B[^0-9]�Ɠ����B
'�@\s       �C�ӂ̃X�y�[�X�����Ƀ}�b�`���܂��B[ \t\r\n\v\f]�Ɠ����B
'�@\S      �C�ӂ̔�X�y�[�X�����Ƀ}�b�`���܂��B[^ \t\r\n\v\f]�Ɠ����B
'</summary>
Public Sub InitRegex(pattern As String, Optional options As String = "")
    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.pattern = pattern
    RegExp.IgnoreCase = IIf(InStr(options, "IgnoreCase"), True, False)
    RegExp.MultiLine = IIf(InStr(options, "MultiLine"), True, False)
End Sub

'<summary>
'  �w�蕶����Ƀp�^�[�����܂܂�Ă��邩�ǂ���
'
'�@input_�F�`�F�b�N���镶����
'  [�ȗ���]startAt�F�������ڂ���`�F�b�N���邩
'</summary>
Public Function IsMatch(input_ As String, Optional startAt As Integer = 1) As Boolean
    IsMatch = RegExp.test(Mid(input_, startAt))
End Function

'<summary>
'  �w�蕶������Ƀp�^�[�����܂܂�Ă������v�����ŏ��̕����񂾂��Ԃ�
'
'�@input_�F�`�F�b�N���镶����
'  [�ȗ���]startAt�F�������ڂ���`�F�b�N���邩
'</summary>
Public Function Match(input_ As String, Optional startAt As Integer = 1)
    RegExp.Global = False
    Set Match = RegExp.Execute(Mid(input_, startAt))
End Function

'<summary>
'  �w�蕶������Ƀp�^�[�����܂܂�Ă�����item�R���N�V�����Ƃ��ĕԂ��B
'
'�@input_�F�`�F�b�N���镶����
'  [�ȗ���]startAt�F�������ڂ���`�F�b�N���邩
'</summary>
Public Function Matches(input_ As String, Optional startAt As Integer = 1)
    RegExp.Global = True
    Set Matches = RegExp.Execute(Mid(input_, startAt))
End Function

'<summary>
'  �w�蕶������Ƀp�^�[�����܂܂�Ă�����w�蕶����ɒu�������ĕԂ�
'
'�@input_�F�`�F�b�N���镶����
'  replacement�F�u��������̕�����
'</summary>
Public Function Replace(input_ As String, replacement As String)
    Replace = RegExp.Replace(input_, replacement)
End Function


' ---------------------------------------------------------------
' .NET Framework System.Diagnostics.Stopwatch �̃N���[��
' ---------------------------------------------------------------

'<summary>
'  �^�C�}�[�̃X�^�[�g
'</summary>
Public Sub StartTimer()
    StartTime = GetTickCount()
End Sub

'<summary>
'  �^�C�}�[�̃X�g�b�v�@�Ԃ�l�̓v���p�e�B����擾����
'</summary>
Public Sub StopTimer()
    EndTime = GetTickCount()
    AccumulatedTime = AccumulatedTime + EndTime - StartTime
End Sub

'<summary>
'  �^�C�}�[�̌o�ߎ��ԃ��Z�b�g
'</summary>
Public Sub ResetTimer()
    AccumulatedTime = 0
End Sub

'<summary>
'  �o�ߎ��Ԃ����Z�b�g���A�V���Ƀ^�C�}�[�X�^�[�g
'</summary>
Public Sub RestartTimer()
    StartTime = GetTickCount()
    AccumulatedTime = 0
End Sub


' ---------------------------------------------------------------
' .NET Framework System.Text.StringBuilder�̃N���[��
' Append�ł������������ݒ肵�AToString�őS�Č����������̂�Ԃ�
' ---------------------------------------------------------------

'<summary>
'  �������镶�����ݒ肷��
'
'�@value�F�ݒ肷�镶����
'</summary>
Public Sub Append(value As String)
    mIndex = mIndex + 1
    Call mBuffer.Add(CStr(mIndex), value)
End Sub

'<summary>
'  �������镶����{���s��ݒ肷��
'
'�@[�ȗ���]value�F�ݒ肷�镶����
'</summary>
Public Sub AppendLine(Optional value As String = "")
    mIndex = mIndex + 1
    Call mBuffer.Add(CStr(mIndex), value & vbCrLf)
End Sub

'<summary>
'  �ݒ肵�Ă���������N���A����
'</summary>
Public Sub ClearDictionary()
    mIndex = 0
    mBuffer.RemoveAll
End Sub

'<summary>
'  �ݒ肵���������S�Č������ĕԂ�
'</summary>
Public Function ToString() As String
    ToString = Join(mBuffer.items(), "")
End Function
























