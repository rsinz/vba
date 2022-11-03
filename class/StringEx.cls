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
Public Property Let value(v As Variant)
Attribute value.VB_UserMemId = 0
    value_ = v
End Property

' 内部保持値に結合
Public Property Let Plus(s As Variant)
    value_ = value_ & CStr(s)
End Property

' 内部保持値に改行して結合
Public Property Let PlusAndLine(s As Variant)
    value_ = value_ & vbNewLine & CStr(s)
End Property

'///////// get property //////////
Public Property Get value() As Variant
    value = CStr(value_)
End Property

' Integer型にして返す
Public Property Get ToInt() As Integer
    ToInt = CInt(value_)
End Property

' Long型にして返す
Public Property Get ToLong() As Long
    ToLong = CLng(value_)
End Property

' 大文字 -> 小文字にして返す
Public Property Get ToLower() As String
    ToLower = StrConv(value_, vbLowerCase)
End Property

' 小文字-> 大文字 にして返す
Public Property Get ToUpper() As String
    ToUpper = StrConv(value_, vbUpperCase)
End Property

' 全角-> 半角 にして返す
Public Property Get ToNarrow() As String
    ToNarrow = StrConv(value_, vbNarrow)
End Property

' 半角-> 全角 にして返す
Public Property Get ToWide() As String
    ToWide = StrConv(value_, vbWide)
End Property

' String型かどうか判定
Public Property Get IsString() As Boolean
    IsString = (Not IsArray(value_) And Not IsDate(value_) And Not IsEmpty(value_) And Not IsEmpty(value_) And Not IsError(value_) And _
                    Not IsJagg(value_) And Not IsMissing(value_) And Not IsNull(value_) And Not IsNumeric(value_) And Not IsObject(value_))
'    IsString = (TypeName(value_) = "String")        ' こっちの場合は "1"とかでも文字列として認識する
End Property

' 空文字かどうか判定
Public Property Get IsEmpty_() As Boolean
    IsEmpty_ = (value_ = vbNullString)
End Property

' 文字列の長さを返す
Public Property Get Length() As Variant
    Length = Len(value_)
End Property

'///////// functions //////////

' Mid_:                 指定された文字以降(以前)の文字列を抜き出す
' Insert:               文字列の挿入
' InStr_:               指定文字列の最初(最左)の位置。StringExを直接InStr()に入れるとバグるため。
' InStrLast:          指定文字列の最後(最右)の位置。
' CountOf:           指定文字列がいくつ含まれているか
' GetClipboard:    クリップボードのテキスト取得
' SetClipboard:    クリップボードにテキストを設定
' ToText:              テキストファイルに文字列追加
' Log:                   Logテキストファイルに出力｡Logファイルが無かったら作る
' RegexInit:         Regexの初期化｡正規表現を使いたいときに (ユーザー入力の妥当性チェックとか)
' RegexIsMatch:  指定文字列にパターンが含まれているかどうか
' RegexMatch:     指定文字列内にパターンが含まれていたら一致した最初の文字列だけ返す
' RegexMatches:  指定文字列内にパターンが含まれていたらitemコレクションとして返す｡
' RegexReplace:  指定文字列内にパターンが含まれていたら指定文字列に置き換えて返す



'<summary>
'  指定された文字以降(以前)の文字列を抜き出す
'
'  pointer：ここに指定された文字以降(以前)が抜き出し対象になる
'　[省略可] isAfter：デフォルトで以降を抜き出す。false/ pointerStrより前を抜き出す
'  [省略可] s：対象文字列。ここから抜き出す
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
'  文字列の挿入
'
'  index：挿入位置。例：index = 3. insert_ = "C". s = "ABDE". -> Insert = "ABCDE"
'　insert_：挿入する文字列
'  [省略可] s：対象文字列。これに挿入する
'</summary>
Public Function Insert(index As Integer, insert_ As String, Optional s As String = "") As String

    Dim s_ As String: s_ = IIf(s = "", value_, s)
    Dim LString As String: LString = Left(s_, index - 1)
    Dim RString As String: RString = Right(s_, index - 1)
    
    Insert = LString & insert_ & RString
End Function

'<summary>
'  指定文字列の最初(最左)の位置。StringExを直接InStr()に入れるとバグるため。
'
'  pointer：ここに指定された文字が検索対象になる
'  [省略可] s：対象文字列。これに挿入する
'</summary>
Public Function InStr_(pointer As String, Optional s As String = "") As Long
    Dim s_ As String: s_ = IIf(s = "", value_, s)
    InStr_ = InStr(s_, pointer)
End Function

'<summary>
'  指定文字列の最後(最右)の位置。
'
'  pointer：ここに指定された文字が検索対象になる
'  [省略可] s：対象文字列。これに挿入する
'</summary>
Public Function InStrLast(pointer As String, Optional s As String = "") As Long
    Dim s_ As String: s_ = IIf(s = "", value_, s)
    InStrLast = Len(s_) - InStr(StrReverse(s_), pointer) + 1
End Function

'<summary>
'  指定文字列がいくつ含まれているか
'
'  pointer：ここに指定された文字が検索対象になる
'  [省略可] s：対象文字列。これに挿入する
'</summary>
Public Function CountOf(pointer As String, Optional s As String = "") As Long
    CountOf = UBound(Split(IIf(s = "", value_, s), pointer))
End Function

'<summary>
'  クリップボードのテキスト取得
'</summary>
Public Function GetClipboard() As String
    GetClipboard = Clipboard.GetText
End Function

'<summary>
'  クリップボードにテキストを設定
'
'　[省略可] s：設定する文字列
'</summary>
Public Sub SetClipboard(Optional s As String = "")
    Clipboard.SetText (IIf(s = "", value_, s))
End Sub

'<summary>
'  テキストファイルに文字列追加
'
'　path：ファイルの完全パス
'  [省略可] s：追加する文字列
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
'  Logテキストファイルに出力。Logファイルが無かったら作る
'
'  [省略可] s：残したい文字列
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
'  Regexの初期化。正規表現を使いたいときに(ユーザー入力の妥当性チェックとか)
'
' pattern：検索文字列パターン
'　[省略可] 大文字小文字を区別する：デフォルトで大文字小文字を区別しない
'　[省略可] 複数行検索する：デフォルトで複数行検索しない
'
'　****  パターン一例  参考：https://msdn.microsoft.com/ja-jp/library/ms974570.aspx
'　英数字 英字と数字に文字どおりにマッチします｡
'　\n       改行にマッチします。
'　\f        フォーム フィードにマッチします。
'　\r       キャリッジ リターンにマッチします。
'　\t       水平タブにマッチします。
'　\v       垂直タブにマッチします。
'　\\       \ にマッチします。
'　[xyz]  文字セットに含まれている任意の1文字にマッチします｡
'　\w      単語に使用される任意の文字にマッチします。[a-zA-Z_0-9]と等価。
'　\W     単語に使用される文字以外の任意の文字にマッチします。[^a-zA-Z_0-9]と等価。
'　\d      任意の数字にマッチします。[0-9]と等価。
'　\D      任意の数字以外の文字にマッチします。[^0-9]と等価。
'　\s       任意のスペース文字にマッチします。[ \t\r\n\v\f]と等価。
'　\S      任意の非スペース文字にマッチします。[^ \t\r\n\v\f]と等価。
'</summary>
Public Sub RegexInit(pattern As String, Optional 大文字小文字を区別する As Boolean = False, Optional 複数行検索する As Boolean = False)
    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.pattern = pattern
    RegExp.IgnoreCase = 大文字小文字を区別する
    RegExp.MultiLine = 複数行検索する
End Sub

'<summary>
'  指定文字列にパターンが含まれているかどうか
'
'  [省略可] startAt：何文字目からチェックするか
'　[省略可] s_：チェックする文字列
'</summary>
Public Function RegexIsMatch(Optional startAt As Integer = 1, Optional s As String = "") As Boolean
    Dim s_ As String: s_ = IIf(s = "", value_, s)
    IsMatch = RegExp.test(Mid(s_, startAt))
End Function

'<summary>
'  指定文字列内にパターンが含まれていたら一致した最初の文字列だけ返す
'
'  [省略可] startAt：何文字目からチェックするか
'　[省略可] s_：チェックする文字列
'</summary>
Public Function RegexMatch(Optional startAt As Integer = 1, Optional s As String = "")
    Dim s_ As String: s_ = IIf(s = "", value_, s)
    RegExp.Global = False
    Set Match = RegExp.Execute(Mid(s_, startAt))
End Function

'<summary>
'  指定文字列内にパターンが含まれていたらitemコレクションとして返す。
'
'  [省略可] startAt：何文字目からチェックするか
'　[省略可] s_：チェックする文字列
'</summary>
Public Function RegexMatches(Optional startAt As Integer = 1, Optional s As String = "")
    Dim s_ As String: s_ = IIf(s = "", value_, s)
    RegExp.Global = True
    Set Matches = RegExp.Execute(Mid(s_, startAt))
End Function

'<summary>
'  指定文字列内にパターンが含まれていたら指定文字列に置き換えて返す
'
'  replacement：置き換え後の文字列
'　[省略可] s_：チェックする文字列
'</summary>
Public Function RegexReplace(replacement As String, Optional s_ As String = "")
    Dim s_ As String: s_ = IIf(s = "", value_, s)
    Replace = RegExp.Replace(s_, replacement)
End Function














