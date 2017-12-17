VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IntegerEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private value_ As Variant


'///////// let property //////////
Public Property Let Value(v As Variant)
Attribute Value.VB_UserMemId = 0
    value_ = v
End Property

' �����ێ��l�ɉ��Z
Public Property Let Plus(d As Double)
    value_ = value_ + d
End Property

' �����ێ��l�Ɍ��Z
Public Property Let Minus(d As Double)
    value_ = value_ - d
End Property

'///////// get property //////////

' ���݃Z�b�g����Ă���l�����̂܂ܕԂ�
Public Property Get Value() As Variant
    Value = value_
End Property

' integer�^�ɂ��ĕԂ�
Public Property Get ToInt() As Variant
    ToInt = CInt(value_)
End Property

' long�^�ɂ��ĕԂ�
Public Property Get ToLong() As Variant
    ToLong = CLng(value_)
End Property

' string�^�ɂ��ĕԂ�
Public Property Get ToString() As Variant
    ToString = CStr(value_)
End Property

' ���R�������� (�����_��}�C�i�X�l�A�[�������e���Ȃ�)
Public Property Get IsNaturalNumber(Optional num As Variant = Empty) As Variant
    Dim num_ As Variant: num_ = IIf(IsEmpty(num), value_, num)
    IsNaturalNumber = (IsNumeric(num_) And (InStr(num_, ".") = 0) And (InStr(num_, ",") = 0) And (num_ > 0))
End Property

'///////// functions //////////

' Plus1:                        �C���N�������g
' Minus1:                     �f�N�������g
' ARound:                    ���l�̊ۂߍ���
' BinaryToDecimal:      2�i����10�i���ϊ�
' DecimalToBinary:      10�i����2�i���ϊ�
' HexToDecimal:          16�i����10�i���ϊ�
' DecimalToHex:          10�i����16�i���ϊ�


'<summary>
'  �C���N�������g
'</summary>
Function Plus1() As Variant
    value_ = value_ + 1
    Plus1 = value_
End Function

'<summary>
'  �f�N�������g
'</summary>
Function Minus1() As Variant
    value_ = value_ - 1
    Minus1 = value_
End Function

'<summary>
'  ���l�̊ۂߍ���
'
'  num�F�ۂߍ��ݑΏۂ̐��l
'�@[�ȗ���]  digits�F�ۂߍ��ݏ��������B�f�t�H���g�ŏ����_�ȉ�1�ʂ��ۂߍ��ށB�}�C�i�X�w��Ő��������ۂߍ��ނ̂��L��
'�@[�ȗ���]  rndup�F�ۂߍ���臒l�B�f�t�H���g�Ŏl�̌ܓ��̐U�镑�� ��F8��8�ȏ�؂�グ�B����ȊO�؂�̂�
'</summary>
Function Round_(Optional ByVal digits As Integer = 0, Optional rndup As Integer = 5, Optional num As Variant = Empty) As Variant
    
On Error GoTo Escape

    Dim num_ As Variant: num_ = IIf(IsEmpty(num), value_, num)
    If Not IsNumeric(num_) Then Err.Raise 13
    If Not (1 <= rndup And rndup <= 10) Then Err.Raise 5
    
    num_ = CDec(num_)
    Dim n As Variant: n = CDec(10 ^ Abs(digits))
    Dim z As Variant: z = CDec(Sgn(num_) * 0.1 * (10 - rndup))
    If digits >= 1 Then
        Round_ = Fix(num_ * n + z) / n
    Else
        Round_ = Fix(num_ / n + z) * n
    End If
    
Escape:
End Function

'<summary>
'  2�i����10�i���ϊ�
'
'  [�ȗ���]  bits�F�ϊ�����2�i��
'</summary>
Function BinaryToDecimal(Optional bits As Variant = Empty) As Long

    Dim bits_ As Variant: bits_ = IIf(IsEmpty(bits), value_, bits)
    If IsNumeric(bits_) = False Then Exit Function
    
    BinaryToDecimal = WorksheetFunction.Bin2Dec(bits_)
End Function

'<summary>
'  10�i����2�i���ϊ�
'
'  [�ȗ���]  dec�F�ϊ�����10�i��
'</summary>
Function DecimalToBinary(Optional dec As Variant = Empty) As String

    Dim dec_ As Variant: dec_ = IIf(IsEmpty(dec), value_, dec)
    If IsNumeric(dec_) = False Then Exit Function
    
    DecimalToBinary = WorksheetFunction.Dec2Bin(dec_)
End Function

'<summary>
'  16�i����10�i���ϊ�
'
'  [�ȗ���]  hexa�F�ϊ�����16�i��
'</summary>
Function HexToDecimal(Optional hexa As Variant = Empty) As Long
    Dim hexa_ As Variant: hexa_ = IIf(IsEmpty(hexa), value_, hexa)
    HexToDecimal = CLng("&H" & hexa_)
End Function

'<summary>
'  10�i����16�i���ϊ�
'
'  [�ȗ���]  dec�F�ϊ�����10�i��
'</summary>
Function DecimalToHex(Optional dec As Variant = Empty) As String
    
    Dim dec_ As Variant: dec_ = IIf(IsEmpty(dec), value_, dec)
    If IsNumeric(dec_) = False Then Exit Function
    
    DecimalToHex = Hex(dec)
End Function
















