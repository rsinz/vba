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

' 内部保持値に加算
Public Property Let Plus(d As Double)
    value_ = value_ + d
End Property

' 内部保持値に減算
Public Property Let Minus(d As Double)
    value_ = value_ - d
End Property

'///////// get property //////////

' 現在セットされている値をそのまま返す
Public Property Get Value() As Variant
    Value = value_
End Property

' integer型にして返す
Public Property Get ToInt() As Variant
    ToInt = CInt(value_)
End Property

' long型にして返す
Public Property Get ToLong() As Variant
    ToLong = CLng(value_)
End Property

' string型にして返す
Public Property Get ToString() As Variant
    ToString = CStr(value_)
End Property

' 自然数か判定 (小数点やマイナス値、ゼロを許容しない)
Public Property Get IsNaturalNumber(Optional num As Variant = Empty) As Variant
    Dim num_ As Variant: num_ = IIf(IsEmpty(num), value_, num)
    IsNaturalNumber = (IsNumeric(num_) And (InStr(num_, ".") = 0) And (InStr(num_, ",") = 0) And (num_ > 0))
End Property

'///////// functions //////////

' Plus1:                        インクリメント
' Minus1:                     デクリメント
' ARound:                    数値の丸め込み
' BinaryToDecimal:      2進数→10進数変換
' DecimalToBinary:      10進数→2進数変換
' HexToDecimal:          16進数→10進数変換
' DecimalToHex:          10進数→16進数変換


'<summary>
'  インクリメント
'</summary>
Function Plus1() As Variant
    value_ = value_ + 1
    Plus1 = value_
End Function

'<summary>
'  デクリメント
'</summary>
Function Minus1() As Variant
    value_ = value_ - 1
    Minus1 = value_
End Function

'<summary>
'  数値の丸め込み
'
'  num：丸め込み対象の数値
'　[省略可]  digits：丸め込み小数桁数。デフォルトで小数点以下1位を丸め込む。マイナス指定で整数桁を丸め込むのも有り
'　[省略可]  rndup：丸め込む閾値。デフォルトで四捨五入の振る舞い 例：8→8以上切り上げ。それ以外切り捨て
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
'  2進数→10進数変換
'
'  [省略可]  bits：変換する2進数
'</summary>
Function BinaryToDecimal(Optional bits As Variant = Empty) As Long

    Dim bits_ As Variant: bits_ = IIf(IsEmpty(bits), value_, bits)
    If IsNumeric(bits_) = False Then Exit Function
    
    BinaryToDecimal = WorksheetFunction.Bin2Dec(bits_)
End Function

'<summary>
'  10進数→2進数変換
'
'  [省略可]  dec：変換する10進数
'</summary>
Function DecimalToBinary(Optional dec As Variant = Empty) As String

    Dim dec_ As Variant: dec_ = IIf(IsEmpty(dec), value_, dec)
    If IsNumeric(dec_) = False Then Exit Function
    
    DecimalToBinary = WorksheetFunction.Dec2Bin(dec_)
End Function

'<summary>
'  16進数→10進数変換
'
'  [省略可]  hexa：変換する16進数
'</summary>
Function HexToDecimal(Optional hexa As Variant = Empty) As Long
    Dim hexa_ As Variant: hexa_ = IIf(IsEmpty(hexa), value_, hexa)
    HexToDecimal = CLng("&H" & hexa_)
End Function

'<summary>
'  10進数→16進数変換
'
'  [省略可]  dec：変換する10進数
'</summary>
Function DecimalToHex(Optional dec As Variant = Empty) As String
    
    Dim dec_ As Variant: dec_ = IIf(IsEmpty(dec), value_, dec)
    If IsNumeric(dec_) = False Then Exit Function
    
    DecimalToHex = Hex(dec)
End Function

















