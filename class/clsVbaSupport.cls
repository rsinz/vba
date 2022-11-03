VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVbaSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 最終更新日：2016/07/16  ＠しら

' 外部アプリケーションにメッセージを送る
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const WM_CLOSE As Long = &H10
          
'///////// enumerated type //////////


'///////// class property //////////



'///////// module list //////////

' ARound:                    数値の丸め込み
' MidInString:               指定された文字以降/以前の文字列を返す
' BeginOfMonth:          指定された日付の月の最初の日付を返す
' EndOfMonth:             指定された日付の月の最後の日付を返す
' BeginOfWeek:           指定された日付の週の最初の日付を返す
' EndOfWeek:              指定された日付の週の最後の日付を返す
' BinaryToDecimal:      2進数→10進数変換
' DecimalToBinary:      10進数→2進数変換
' HexToDecimal:          16進数→10進数変換
' DecimalToHex:          10進数→16進数変換
' EvalVBS:                    VBScriptで書かれた構文を実行する
' EvalJS:                       JavaScriptで書かれた構文を実行する



'<summary>
'  数値の丸め込み
'
'  num：丸め込み対象の数値
'　[省略可]  digits：丸め込み桁数+1。デフォルトで小数点以下1位を丸め込む。マイナス指定で整数桁を丸め込むのも有り
'　[省略可]  rndup：丸め込む閾値。デフォルトで四捨五入の振る舞い 例：8→8以上切り上げ。それ以外切り捨て
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
'  指定された文字以降(以前)の文字列を抜き出す
'
'  targetStr：対象文字列。ここから抜き出す
'  pointerStr：ここに指定された文字以降(以前)が抜き出し対象になる
'　[省略可]  isAfter：デフォルトで以降を抜き出す。false/ pointerStrより前を抜き出す
'</summary>
Function MidInString(targetStr As String, pointerStr As String, Optional isAfter As Boolean = True)

    If isAfter Then
        MidInString = Mid(targetStr, InStr(targetStr, pointerStr) + 1)
    Else
        MidInString = Mid(targetStr, 1, InStr(targetStr, pointerStr) - 1)
    End If

End Function

'<summary>
'  指定された日付の月の最初の日付を返す
'
'  dt：基準となる日付
'</summary>
Function BeginOfMonth(ByVal dt As Date) As Date
    BeginOfMonth = DateAdd("d", -Day(dt) + 1, dt)
End Function

'<summary>
'  指定された日付の月の最後の日付を返す
'
'  dt：基準となる日付
'</summary>
Function EndOfMonth(ByVal dt As Date) As Date
    EndOfMonth = DateAdd("d", -1, BeginOfMonth(DateAdd("m", 1, dt)))
End Function

'<summary>
'  指定された日付の週の最初の日付を返す
'
'  dt：基準となる日付
'　[省略可]  fstDayOfWeek：週始めの曜日。デフォルトで日曜
'</summary>
Function BeginOfWeek(ByVal dt As Date, Optional fstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    BeginOfWeek = DateAdd("d", 1 - Weekday(dt, fstDayOfWeek), dt)
End Function

'<summary>
'  指定された日付の週の最後の日付を返す
'
'  dt：基準となる日付
'　[省略可]  fstDayOfWeek：週始めの曜日。デフォルトで日曜
'</summary>
Public Function EndOfWeek(ByVal dt As Date, Optional fstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    EndOfWeek = DateAdd("d", 7 - Weekday(dt, fstDayOfWeek), dt)
End Function

'<summary>
'  2進数→10進数変換
'
'  flgs：変換する2進数　例：(1,0,0)でも (True,Flase,False)でも ("1","0","0") でも結果は４
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
'  10進数→2進数変換
'
'  dec：変換する10進数
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
'  16進数→10進数変換
'
'  hexa：変換する16進数
'</summary>
Function HexToDecimal(hexa As String) As Long
    HexToDecimal = CLng("&H" & hexa)
End Function

'<summary>
'  10進数→16進数変換
'
'  dec：変換する10進数
'</summary>
Function DecimalToHex(dec As Long) As String
    DecimalToHex = Hex(dec)
End Function

'<summary>
'  VBScriptで書かれた構文を実行する
'
'  vbsExpr：VBScript構文
'</summary>
Function EvalVBS(ByVal vbsExpr As String) As String
    EvalVBS = EvalScript(vbsExpr, "VBScript")
End Function

'<summary>
'  JavaScriptで書かれた構文を実行する
'
'  jsExpr：JavaScript構文
'</summary>
Function EvalJS(ByVal jsExpr As String) As String
    EvalJS = EvalScript(jsExpr, "JScript")
End Function

'<summary>
'  VBA以外の言語で書かれたプログラムを実行する
'
'  expr：構文
'  lang：言語の種類
'</summary>
Private Function EvalScript(ByVal expr As String, ByVal lang As String) As String
    Dim Doc As Object: Set Doc = CreateObject("HtmlFile")
    Doc.parentWindow.execScript "document.write(" & expr & ")", lang
    If Not Doc.body Is Nothing Then EvalScript = Doc.body.innerHTML
End Function

'<summary>
'  外部アプリケーションを開く
'
'  path : アプリケーションのパス。
'  name : アプリケーションの名前。"電卓"とか"Form1"とか
'</summary>
Sub OpenExe(path As String, name As String)
    Set ws = CreateObject("Wscript.Shell")
    ws.Run path & "\" & name & ".exe", 1
End Sub

'<summary>
'  外部アプリケーションを閉じる
'
'  name : アプリケーションの名前。"電卓"とか"Form1"とか
'</summary>
Sub CloseExe(name As String)
    Dim hwnd As Long, ret As Long

    If Len(name) Then hwnd = FindWindow(vbNullString, name)
    ret = SendMessage(hwnd, WM_CLOSE, 0&, 0&)
End Sub












