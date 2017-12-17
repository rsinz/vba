VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private value_ As Variant
Private Clipboard As Object
Private RegExp As Object


Private Sub Class_Initialize()
    Set Clipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
End Sub

'///////// let property //////////
Public Property Let Value(v As Variant)
Attribute Value.VB_UserMemId = 0
    value_ = v
End Property

' �����ێ��l�Ɍ���
Public Property Let Plus(s As Variant)
    value_ = value_ & CStr(s)
End Property

' �����ێ��l�ɉ��s���Č���
Public Property Let PlusAndLine(s As Variant)
    value_ = value_ & vbNewLine & CStr(s)
End Property

'///////// get property //////////
Public Property Get Value() As Variant
    Value = CStr(value_)
End Property

' Integer�^�ɂ��ĕԂ�
Public Property Get ToInt() As Integer
    ToInt = CInt(value_)
End Property

' Long�^�ɂ��ĕԂ�
Public Property Get ToLong() As Long
    ToLong = CLng(value_)
End Property

' �啶�� -> �������ɂ��ĕԂ�
Public Property Get ToLower() As String
    ToLower = StrConv(value_, vbLowerCase)
End Property

' ������-> �啶�� �ɂ��ĕԂ�
Public Property Get ToUpper() As String
    ToUpper = StrConv(value_, vbUpperCase)
End Property

' �S�p-> ���p �ɂ��ĕԂ�
Public Property Get ToNarrow() As String
    ToNarrow = StrConv(value_, vbNarrow)
End Property

' ���p-> �S�p �ɂ��ĕԂ�
Public Property Get ToWide() As String
    ToWide = StrConv(value_, vbWide)
End Property

' String�^���ǂ�������
Public Property Get IsString() As Boolean
    IsString = (Not IsArray(value_) And Not IsDate(value_) And Not IsEmpty(value_) And Not IsEmpty(value_) And Not IsError(value_) And _
                    Not IsJagg(value_) And Not IsMissing(value_) And Not IsNull(value_) And Not IsNumeric(value_) And Not IsObject(value_))
'    IsString = (TypeName(value_) = "String")        ' �������̏ꍇ�� "1"�Ƃ��ł�������Ƃ��ĔF������
End Property

' �󕶎����ǂ�������
Public Property Get IsEmpty_() As Boolean
    IsEmpty_ = (value_ = vbNullString)
End Property

' ������̒�����Ԃ�
Public Property Get Length() As Variant
    Length = Len(value_)
End Property

'///////// functions //////////

' Mid_:                 �w�肳�ꂽ�����ȍ~(�ȑO)�̕�����𔲂��o��
' Insert:               ������̑}��
' InStr_:               �w�蕶����̍ŏ�(�ō�)�̈ʒu�BStringEx�𒼐�InStr()�ɓ����ƃo�O�邽�߁B
' InStrLast:          �w�蕶����̍Ō�(�ŉE)�̈ʒu�B
' CountOf:           �w�蕶���񂪂����܂܂�Ă��邩
' GetClipboard:    �N���b�v�{�[�h�̃e�L�X�g�擾
' SetClipboard:    �N���b�v�{�[�h�Ƀe�L�X�g��ݒ�
' ToText:              �e�L�X�g�t�@�C���ɕ�����ǉ�
' Log:                   Log�e�L�X�g�t�@�C���ɏo�͡Log�t�@�C����������������
' RegexInit:         Regex�̏���������K�\�����g�������Ƃ��� (���[�U�[���͂̑Ó����`�F�b�N�Ƃ�)
' RegexIsMatch:  �w�蕶����Ƀp�^�[�����܂܂�Ă��邩�ǂ���
' RegexMatch:     �w�蕶������Ƀp�^�[�����܂܂�Ă������v�����ŏ��̕����񂾂��Ԃ�
' RegexMatches:  �w�蕶������Ƀp�^�[�����܂܂�Ă�����item�R���N�V�����Ƃ��ĕԂ��
' RegexReplace:  �w�蕶������Ƀp�^�[�����܂܂�Ă�����w�蕶����ɒu�������ĕԂ�



'<summary>
'  �w�肳�ꂽ�����ȍ~(�ȑO)�̕�����𔲂��o��
'
'  pointer�F�����Ɏw�肳�ꂽ�����ȍ~(�ȑO)�������o���ΏۂɂȂ�
'�@[�ȗ���] isAfter�F�f�t�H���g�ňȍ~�𔲂��o���Bfalse/ pointerStr���O�𔲂��o��
'  [�ȗ���] s�F�Ώە�����B�������甲���o��
'</summary>
Public Function Mid_(pointer As String, Optional isAfter As Boolean = True, Optional s As String = "") As String

    Dim s_ As String: s_ = IIf(s = "", value_, s)
    
    If isAfter Then
        Mid_ = Mid(s_, InStr(s_, pointer) + 1)
    Else
        Mid_ = Mid(s_, 1, InStr(s_, pointer) - 1)
    End If

End Function

'<summary>
'  ������̑}��
'
'  index�F�}���ʒu�B��Findex = 3. insert_ = "C". s = "ABDE". -> Insert = "ABCDE"
'�@insert_�F�}�����镶����
'  [�ȗ���] s�F�Ώە�����B����ɑ}������
'</summary>
Public Function Insert(index As Integer, insert_ As String, Optional s As String = "") As String

    Dim s_ As String: s_ = IIf(s = "", value_, s)
    Dim LString As String: LString = Left(s_, index - 1)
    Dim RString As String: RString = Right(s_, index - 1)
    
    Insert = LString & insert_ & RString
End Function

'<summary>
'  �w�蕶����̍ŏ�(�ō�)�̈ʒu�BStringEx�𒼐�InStr()�ɓ����ƃo�O�邽�߁B
'
'  pointer�F�����Ɏw�肳�ꂽ�����������ΏۂɂȂ�
'  [�ȗ���] s�F�Ώە�����B����ɑ}������
'</summary>
Public Function InStr_(pointer As String, Optional s As String = "") As Long
    Dim s_ As String: s_ = IIf(s = "", value_, s)
    InStr_ = InStr(s_, pointer)
End Function

'<summary>
'  �w�蕶����̍Ō�(�ŉE)�̈ʒu�B
'
'  pointer�F�����Ɏw�肳�ꂽ�����������ΏۂɂȂ�
'  [�ȗ���] s�F�Ώە�����B����ɑ}������
'</summary>
Public Function InStrLast(pointer As String, Optional s As String = "") As Long
    Dim s_ As String: s_ = IIf(s = "", value_, s)
    InStrLast = Len(s_) - InStr(StrReverse(s_), pointer) + 1
End Function

'<summary>
'  �w�蕶���񂪂����܂܂�Ă��邩
'
'  pointer�F�����Ɏw�肳�ꂽ�����������ΏۂɂȂ�
'  [�ȗ���] s�F�Ώە�����B����ɑ}������
'</summary>
Public Function CountOf(pointer As String, Optional s As String = "") As Long
    CountOf = UBound(Split(IIf(s = "", value_, s), pointer))
End Function

'<summary>
'  �N���b�v�{�[�h�̃e�L�X�g�擾
'</summary>
Public Function GetClipboard() As String
    GetClipboard = Clipboard.GetText
End Function

'<summary>
'  �N���b�v�{�[�h�Ƀe�L�X�g��ݒ�
'
'�@[�ȗ���] s�F�ݒ肷�镶����
'</summary>
Public Sub SetClipboard(Optional s As String = "")
    Clipboard.SetText (IIf(s = "", value_, s))
End Sub

'<summary>
'  �e�L�X�g�t�@�C���ɕ�����ǉ�
'
'�@path�F�t�@�C���̊��S�p�X
'  [�ȗ���] s�F�ǉ����镶����
'</summary>
Public Sub ToText(path As String, Optional s As String = "")

    On Error GoTo SkipW
    
    Dim fileNum As Integer: fileNum = FreeFile
    Open path For Append As #fileNum
    
    Dim s_ As String: s_ = IIf(s = "", value_, s)
    Print #fileNum, s_
SkipW:
    Close #fileNum
End Sub

'<summary>
'  Log�e�L�X�g�t�@�C���ɏo�́BLog�t�@�C����������������
'
'  [�ȗ���] s�F�c������������
'</summary>
Public Sub Log(Optional s As String = "")

    Dim path As String: path = ThisWorkbook.path & "\Log.txt"
        
    With CreateObject("Scripting.FileSystemObject")
        If .FileExists(path) = False Then .CreateTextFile (path)
    End With
    
    Dim s_ As String: s_ = IIf(s = "", value_, s)
    Call ToText(path, format(Now, "yyyy/mm/dd hh:mm:ss") & "," & s_)
    
End Sub

'<summary>
'  Regex�̏������B���K�\�����g�������Ƃ���(���[�U�[���͂̑Ó����`�F�b�N�Ƃ�)
'
' pattern�F����������p�^�[��
'�@[�ȗ���] �啶������������ʂ���F�f�t�H���g�ő啶������������ʂ��Ȃ�
'�@[�ȗ���] �����s��������F�f�t�H���g�ŕ����s�������Ȃ�
'
'�@****  �p�^�[�����  �Q�l�Fhttps://msdn.microsoft.com/ja-jp/library/ms974570.aspx
'�@�p���� �p���Ɛ����ɕ����ǂ���Ƀ}�b�`���܂��
'�@\n       ���s�Ƀ}�b�`���܂��B
'�@\f        �t�H�[�� �t�B�[�h�Ƀ}�b�`���܂��B
'�@\r       �L�����b�W ���^�[���Ƀ}�b�`���܂��B
'�@\t       �����^�u�Ƀ}�b�`���܂��B
'�@\v       �����^�u�Ƀ}�b�`���܂��B
'�@\\       \ �Ƀ}�b�`���܂��B
'�@[xyz]  �����Z�b�g�Ɋ܂܂�Ă���C�ӂ�1�����Ƀ}�b�`���܂��
'�@\w      �P��Ɏg�p�����C�ӂ̕����Ƀ}�b�`���܂��B[a-zA-Z_0-9]�Ɠ����B
'�@\W     �P��Ɏg�p����镶���ȊO�̔C�ӂ̕����Ƀ}�b�`���܂��B[^a-zA-Z_0-9]�Ɠ����B
'�@\d      �C�ӂ̐����Ƀ}�b�`���܂��B[0-9]�Ɠ����B
'�@\D      �C�ӂ̐����ȊO�̕����Ƀ}�b�`���܂��B[^0-9]�Ɠ����B
'�@\s       �C�ӂ̃X�y�[�X�����Ƀ}�b�`���܂��B[ \t\r\n\v\f]�Ɠ����B
'�@\S      �C�ӂ̔�X�y�[�X�����Ƀ}�b�`���܂��B[^ \t\r\n\v\f]�Ɠ����B
'</summary>
Public Sub RegexInit(pattern As String, Optional �啶������������ʂ��� As Boolean = False, Optional �����s�������� As Boolean = False)
    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.pattern = pattern
    RegExp.IgnoreCase = �啶������������ʂ���
    RegExp.MultiLine = �����s��������
End Sub

'<summary>
'  �w�蕶����Ƀp�^�[�����܂܂�Ă��邩�ǂ���
'
'  [�ȗ���] startAt�F�������ڂ���`�F�b�N���邩
'�@[�ȗ���] s_�F�`�F�b�N���镶����
'</summary>
Public Function RegexIsMatch(Optional startAt As Integer = 1, Optional s As String = "") As Boolean
    Dim s_ As String: s_ = IIf(s = "", value_, s)
    IsMatch = RegExp.test(Mid(s_, startAt))
End Function

'<summary>
'  �w�蕶������Ƀp�^�[�����܂܂�Ă������v�����ŏ��̕����񂾂��Ԃ�
'
'  [�ȗ���] startAt�F�������ڂ���`�F�b�N���邩
'�@[�ȗ���] s_�F�`�F�b�N���镶����
'</summary>
Public Function RegexMatch(Optional startAt As Integer = 1, Optional s As String = "")
    Dim s_ As String: s_ = IIf(s = "", value_, s)
    RegExp.Global = False
    Set Match = RegExp.Execute(Mid(s_, startAt))
End Function

'<summary>
'  �w�蕶������Ƀp�^�[�����܂܂�Ă�����item�R���N�V�����Ƃ��ĕԂ��B
'
'  [�ȗ���] startAt�F�������ڂ���`�F�b�N���邩
'�@[�ȗ���] s_�F�`�F�b�N���镶����
'</summary>
Public Function RegexMatches(Optional startAt As Integer = 1, Optional s As String = "")
    Dim s_ As String: s_ = IIf(s = "", value_, s)
    RegExp.Global = True
    Set Matches = RegExp.Execute(Mid(s_, startAt))
End Function

'<summary>
'  �w�蕶������Ƀp�^�[�����܂܂�Ă�����w�蕶����ɒu�������ĕԂ�
'
'  replacement�F�u��������̕�����
'�@[�ȗ���] s_�F�`�F�b�N���镶����
'</summary>
Public Function RegexReplace(replacement As String, Optional s_ As String = "")
    Dim s_ As String: s_ = IIf(s = "", value_, s)
    Replace = RegExp.Replace(s_, replacement)
End Function













