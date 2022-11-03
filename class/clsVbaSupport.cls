VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVbaSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' �ŏI�X�V���F2016/07/16  ������

' �O���A�v���P�[�V�����Ƀ��b�Z�[�W�𑗂�
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const WM_CLOSE As Long = &H10
          
'///////// enumerated type //////////


'///////// class property //////////



'///////// module list //////////

' ARound:                    ���l�̊ۂߍ���
' MidInString:               �w�肳�ꂽ�����ȍ~/�ȑO�̕������Ԃ�
' BeginOfMonth:          �w�肳�ꂽ���t�̌��̍ŏ��̓��t��Ԃ�
' EndOfMonth:             �w�肳�ꂽ���t�̌��̍Ō�̓��t��Ԃ�
' BeginOfWeek:           �w�肳�ꂽ���t�̏T�̍ŏ��̓��t��Ԃ�
' EndOfWeek:              �w�肳�ꂽ���t�̏T�̍Ō�̓��t��Ԃ�
' BinaryToDecimal:      2�i����10�i���ϊ�
' DecimalToBinary:      10�i����2�i���ϊ�
' HexToDecimal:          16�i����10�i���ϊ�
' DecimalToHex:          10�i����16�i���ϊ�
' EvalVBS:                    VBScript�ŏ����ꂽ�\�������s����
' EvalJS:                       JavaScript�ŏ����ꂽ�\�������s����



'<summary>
'  ���l�̊ۂߍ���
'
'  num�F�ۂߍ��ݑΏۂ̐��l
'�@[�ȗ���]  digits�F�ۂߍ��݌���+1�B�f�t�H���g�ŏ����_�ȉ�1�ʂ��ۂߍ��ށB�}�C�i�X�w��Ő��������ۂߍ��ނ̂��L��
'�@[�ȗ���]  rndup�F�ۂߍ���臒l�B�f�t�H���g�Ŏl�̌ܓ��̐U�镑�� ��F8��8�ȏ�؂�グ�B����ȊO�؂�̂�
'</summary>
Function ARound( _
    ByVal num As Variant, Optional ByVal digits As Integer = 0, Optional rndup As Integer = 5 _
    ) As Variant
    
On Error GoTo Escape

    If Not IsNumeric(num) Then Err.Raise 13
    If Not (1 <= rndup And rndup <= 10) Then Err.Raise 5
    
    num = CDec(num)
    Dim n As Variant: n = CDec(10 ^ Abs(digits))
    Dim z As Variant: z = CDec(Sgn(num) * 0.1 * (10 - rndup))
    If digits >= 0 Then
        ARound = Fix(num * n + z) / n
    Else
        ARound = Fix(num / n + z) * n
    End If
    
Escape:
End Function

'<summary>
'  �w�肳�ꂽ�����ȍ~(�ȑO)�̕�����𔲂��o��
'
'  targetStr�F�Ώە�����B�������甲���o��
'  pointerStr�F�����Ɏw�肳�ꂽ�����ȍ~(�ȑO)�������o���ΏۂɂȂ�
'�@[�ȗ���]  isAfter�F�f�t�H���g�ňȍ~�𔲂��o���Bfalse/ pointerStr���O�𔲂��o��
'</summary>
Function MidInString(targetStr As String, pointerStr As String, Optional isAfter As Boolean = True)

    If isAfter Then
        MidInString = Mid(targetStr, InStr(targetStr, pointerStr) + 1)
    Else
        MidInString = Mid(targetStr, 1, InStr(targetStr, pointerStr) - 1)
    End If

End Function

'<summary>
'  �w�肳�ꂽ���t�̌��̍ŏ��̓��t��Ԃ�
'
'  dt�F��ƂȂ���t
'</summary>
Function BeginOfMonth(ByVal dt As Date) As Date
    BeginOfMonth = DateAdd("d", -Day(dt) + 1, dt)
End Function

'<summary>
'  �w�肳�ꂽ���t�̌��̍Ō�̓��t��Ԃ�
'
'  dt�F��ƂȂ���t
'</summary>
Function EndOfMonth(ByVal dt As Date) As Date
    EndOfMonth = DateAdd("d", -1, BeginOfMonth(DateAdd("m", 1, dt)))
End Function

'<summary>
'  �w�肳�ꂽ���t�̏T�̍ŏ��̓��t��Ԃ�
'
'  dt�F��ƂȂ���t
'�@[�ȗ���]  fstDayOfWeek�F�T�n�߂̗j���B�f�t�H���g�œ��j
'</summary>
Function BeginOfWeek(ByVal dt As Date, Optional fstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    BeginOfWeek = DateAdd("d", 1 - Weekday(dt, fstDayOfWeek), dt)
End Function

'<summary>
'  �w�肳�ꂽ���t�̏T�̍Ō�̓��t��Ԃ�
'
'  dt�F��ƂȂ���t
'�@[�ȗ���]  fstDayOfWeek�F�T�n�߂̗j���B�f�t�H���g�œ��j
'</summary>
Public Function EndOfWeek(ByVal dt As Date, Optional fstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    EndOfWeek = DateAdd("d", 7 - Weekday(dt, fstDayOfWeek), dt)
End Function

'<summary>
'  2�i����10�i���ϊ�
'
'  flgs�F�ϊ�����2�i���@��F(1,0,0)�ł� (True,Flase,False)�ł� ("1","0","0") �ł����ʂ͂S
'</summary>
Function BinaryToDecimal(ParamArray flgs() As Variant) As Long
    BitFlag = 0
    Dim ub As Long: ub = UBound(flgs)
    
    Dim i As Long
    For i = 0 To ub
        BitFlag = BitFlag + Abs(flgs(i)) * 2 ^ (ub - i)
    Next
End Function

'<summary>
'  10�i����2�i���ϊ�
'
'  dec�F�ϊ�����10�i��
'</summary>
Function DecimalToBinary(dec As Long) As Variant


    Dim bitCnt As Integer: bitCnt = 0
    Dim strData As String: strData = ""
    Dim strArray As Variant

    Do Until (dec < 2 ^ bitCnt)
        If (dec And 2 ^ bitCnt) <> 0 Then
            strData = "1" & strData
        Else
            strData = "0" & strData
        End If

        bitCnt = bitCnt + 1
    Loop

    ReDim strArray(Len(strData) - 1) As String
    For i = 1 To Len(strData)
        strArray(i - 1) = Mid(strData, i, 1)
    Next i
    DecimalToBinary = strArray

End Function

'<summary>
'  16�i����10�i���ϊ�
'
'  hexa�F�ϊ�����16�i��
'</summary>
Function HexToDecimal(hexa As String) As Long
    HexToDecimal = CLng("&H" & hexa)
End Function

'<summary>
'  10�i����16�i���ϊ�
'
'  dec�F�ϊ�����10�i��
'</summary>
Function DecimalToHex(dec As Long) As String
    DecimalToHex = Hex(dec)
End Function

'<summary>
'  VBScript�ŏ����ꂽ�\�������s����
'
'  vbsExpr�FVBScript�\��
'</summary>
Function EvalVBS(ByVal vbsExpr As String) As String
    EvalVBS = EvalScript(vbsExpr, "VBScript")
End Function

'<summary>
'  JavaScript�ŏ����ꂽ�\�������s����
'
'  jsExpr�FJavaScript�\��
'</summary>
Function EvalJS(ByVal jsExpr As String) As String
    EvalJS = EvalScript(jsExpr, "JScript")
End Function

'<summary>
'  VBA�ȊO�̌���ŏ����ꂽ�v���O���������s����
'
'  expr�F�\��
'  lang�F����̎��
'</summary>
Private Function EvalScript(ByVal expr As String, ByVal lang As String) As String
    Dim Doc As Object: Set Doc = CreateObject("HtmlFile")
    Doc.parentWindow.execScript "document.write(" & expr & ")", lang
    If Not Doc.body Is Nothing Then EvalScript = Doc.body.innerHTML
End Function

'<summary>
'  �O���A�v���P�[�V�������J��
'
'  path : �A�v���P�[�V�����̃p�X�B
'  name : �A�v���P�[�V�����̖��O�B"�d��"�Ƃ�"Form1"�Ƃ�
'</summary>
Sub OpenExe(path As String, name As String)
    Set ws = CreateObject("Wscript.Shell")
    ws.Run path & "\" & name & ".exe", 1
End Sub

'<summary>
'  �O���A�v���P�[�V���������
'
'  name : �A�v���P�[�V�����̖��O�B"�d��"�Ƃ�"Form1"�Ƃ�
'</summary>
Sub CloseExe(name As String)
    Dim hwnd As Long, ret As Long

    If Len(name) Then hwnd = FindWindow(vbNullString, name)
    ret = SendMessage(hwnd, WM_CLOSE, 0&, 0&)
End Sub











