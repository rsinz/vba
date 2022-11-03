VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFrameworkClone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 最終更新日：2016/06/23  ＠しら

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

Property Get Elapsed()      ' stopwatchでの経過時間
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
' .NET Framework System.Windows.Forms.Clipboardのクローン
' ---------------------------------------------------------------
' GetClipText:          クリップボードのテキスト取得
' SetClipText:          クリップボードにテキストを設定
' ClearClipBoard:    クリップボードクリア

' ---------------------------------------------------------------
' .NET Framework System.IO.Directoryのクローン
' ---------------------------------------------------------------
' CreateDirectory:          フォルダ作成
' DeleteDirectory:          フォルダ削除
' CopyDirectory:             フォルダコピー
' ExistsDirectory:            フォルダの存在チェック
' GetCurrentDirectory:   現在フォーカスされている階層を返す
' MoveDirectory:            フォルダ移動
' SetCurrentDirectory:   指定パスをフォーカスする
' GetFolderPath:            ファイルの参照先取得｡参照ダイアログでユーザーに選択してもらうoriginal
' SearchFile:                  指定フォルダに対してワイルドカードファイルの検索｡見つかれば正規ファイル名を返すoriginal

' ---------------------------------------------------------------
' .NET Framework System.IO.Fileのクローン
' ---------------------------------------------------------------
' AppendAllText:        指定した文字列をファイルに追加｡ファイルがまだ存在しない場合は､ファイルを作成｡
' CopyFile:                 ファイルのコピー
' DeleteFile:               ファイルの削除
' ExistsFile:                 ファイルの存在チェック
' MoveFile:                 ファイル移動
' ReadAllText:            ファイル読み込み｡読み込んだデータは文字列で返す
' ReadAllLines:           ファイル読み込み｡文字列を行ごとに配列で返す
' WriteAllText:            ファイル書き込み｡一気に全部書く
' WriteAllLines:          ファイル書き込み｡書き込み文字列を行ごとに配列で指定したいときに

' ---------------------------------------------------------------
' .NET Framework System.Text.RegularExpressions.Regex のクローン
' ---------------------------------------------------------------
' InitRegex:            Regexの初期化｡正規表現を使いたいときに (ユーザー入力の妥当性チェックとか)
' IsMatch:               指定文字列にパターンが含まれているかどうか
' Match:                 指定文字列内にパターンが含まれていたら一致した最初の文字列だけ返す
' Matches:              指定文字列内にパターンが含まれていたらitemコレクションとして返す｡
' Replace:              指定文字列内にパターンが含まれていたら指定文字列に置き換えて返す

' ---------------------------------------------------------------
' .NET Framework System.Diagnostics.Stopwatch のクローン
' ---------------------------------------------------------------
'StartTimer:        タイマーのスタート
'StopTimer:         タイマーのストップ 返り値はプロパティから取得して
'ResetTimer:       タイマーの経過時間リセット
'RestartTimer:    経過時間をリセットしつつ､新たにタイマースタート

' ---------------------------------------------------------------
' .NET Framework System.Text.StringBuilderのクローン
' ---------------------------------------------------------------
' Append:               結合する文字列を設定する
' AppendLine:        結合する文字列+改行を設定する
' ClearDictionary:   設定してた文字列をクリアする
' ToString:              設定した文字列を全て結合して返す



'///////// module //////////

' ---------------------------------------------------------------
' .NET Framework System.Windows.Forms.Clipboardのクローン
' ---------------------------------------------------------------

'<summary>
'  クリップボードのテキスト取得
'</summary>
Function GetClipText()
    With DOClipboard
        .GetFromClipboard
        GetClipText = .GetClipText
    End With
End Function

'<summary>
'  クリップボードにテキストを設定
'
'　str：設定する文字列
'</summary>
Sub SetClipText(str As String)
    With DOClipboard
        .SetText str
        .PutInClipboard
    End With
End Sub

'<summary>
'  クリップボードクリア
'</summary>
Sub ClearClipBoard()
    With DOClipboard
        .Clear
    End With
End Sub

' ---------------------------------------------------------------
' .NET Framework System.IO.Directoryのクローン
' ---------------------------------------------------------------

'<summary>
'  フォルダ作成
'
'　path：作成する場所。完全パスで指定
'</summary>
Public Sub CreateDirectory(path As String)
    Call MkDir(path)
End Sub

'<summary>
'  フォルダ削除
'
'　path：削除するフォルダの場所。完全パスで指定
'  [省略可]recursive：Trueで読み取り専用の属性を持つフォルダも削除。デフォルト：False
'</summary>
Public Sub DeleteDirectory(path As String, Optional recursive As Boolean = False)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.DeleteFolder(path, recursive)
End Sub

'<summary>
'  フォルダコピー
'
'　sourcePath：コピー元のパス
'　destinationPath：コピー先のパス
'  [省略可]recursive：Trueで読み取り専用の属性を持つフォルダもコピー対象に。デフォルト：False
'</summary>
Public Sub CopyDirectory(sourcePath As String, destinationPath As String, Optional recursive As Boolean = False)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.CopyFolder(sourcePath, destinationPath, recursive)
End Sub

'<summary>
'  フォルダの存在チェック
'
'　path：チェックするフォルダパス
'</summary>
Public Function ExistsDirectory(path As String)
    ExistsDirectory = IIf(Dir(path, vbDirectory) = "", False, True)
End Function

'<summary>
'  現在フォーカスされている階層を返す
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
'  フォルダ移動
'
'　sourcePath：移動元のパス
'　destinationPath：移動先のパス
'</summary>
Public Sub MoveDirectory(sourcePath As String, destinationPath As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.moveFolder(sourcePath, destinationPath)
End Sub

'<summary>
'  指定パスをフォーカスする
'
'　path：フォーカスするパス
'</summary>
Public Sub SetCurrentDirectory(path As String)
    Call ChDrive(path)
    Call ChDir(path)
End Sub

'<summary>
'  ファイルの参照先取得。参照ダイアログでユーザーに選択してもらう
'</summary>
Function GetFolderPath() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            GetFolderPath = .SelectedItems(1)
        End If
    End With
End Function

'<summary>
'　指定フォルダに対してワイルドカードファイルの検索。見つかれば正規ファイル名を返す
'
'　folderPath：チェック対象フォルダのパス
'  filePath : 検索ファイル名
'</summary>
Function SearchFile(folderPath As String, filePath As String) As Variant
    SearchFile = Dir(folderPath & filePath, vbNormal)
End Function

' ---------------------------------------------------------------
' .NET Framework System.IO.Fileのクローン
' ---------------------------------------------------------------

'<summary>
'  指定した文字列をファイルに追加。ファイルがまだ存在しない場合は、ファイルを作成。
'
'　path：ファイルの完全パス
'  str：追加する文字列
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
'  ファイルのコピー
'
'　from_：コピー元　名前を含んだパス
'  to_：コピー先　名前を含んだパス
'  [省略可]overwrite：上書きするか。デフォルト：False
'</summary>
Public Sub CopyFile(from_ As String, To_ As String, Optional overwrite As Boolean = False)
    If Not ExistsFile(To_) Or (ExistsFile(To_) And overwrite) Then
        Call FileCopy(from_, To_)
    End If
End Sub

'<summary>
'  ファイルの削除
'
'　path：削除するファイルのパス
'</summary>
Public Sub DeleteFile(path As String)
    Call Kill(path)
End Sub

'<summary>
'  ファイルの存在チェック
'
'　path：チェックするファイルのパス
'</summary>
Public Function ExistsFile(path As String)
    With CreateObject("Scripting.FileSystemObject")
        ExistsFile = IIf(.FileExists(path), True, False)
    End With
End Function

'<summary>
'  ファイル移動
'
'　from_：移動元　名前を含んだパス
'  to_：移動先　名前を含んだパス
'</summary>
Public Sub MoveFile(from_ As String, To_ As String)
    With CreateObject("Scripting.FileSystemObject")
        Call .MoveFile(from_, To_)
    End With
End Sub

'<summary>
'  ファイル読み込み。読み込んだデータは文字列で返す
'
'　fileName：ファイル名を含んだパス
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
'  ファイル読み込み。文字列を行ごとに配列で返す ※末尾に変な空白文字が入るんだけど誰か直して
'
'　fileName：ファイル名を含んだパス
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
'  ファイル書き込み。一気に全部書く
'
'　path：ファイル名を含んだパス
'  contents：書き込みたい文字列
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
'  ファイル書き込み。書き込み文字列を行ごとに配列で指定したいときに
'
'　path：ファイル名を含んだパス
'  contents：書き込みたい文字列
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
' .NET Framework System.Text.RegularExpressions.Regex のクローン
' ---------------------------------------------------------------

'<summary>
'  Regexの初期化。正規表現を使いたいときに(ユーザー入力の妥当性チェックとか)
'
'　pattern：検索文字列パターン
'  options：オプション
'                   RegexOption.IgnoreCase：大文字小文字を区別する
'                   RegexOption.MultiLine：複数行検索する
'
'　****  パターン一例  参考：https://msdn.microsoft.com/ja-jp/library/ms974570.aspx
'　英数字 英字と数字に文字どおりにマッチします｡
'　\n       改行にマッチします。
'　\f        フォーム フィードにマッチします。
'　\r       キャリッジ リターンにマッチします。
'　\t       水平タブにマッチします。
'　\v       垂直タブにマッチします。
'　\?       ? にマッチします。
'　\*       * にマッチします。
'　\+      + にマッチします。
'　\.       . にマッチします。
'　\|       | にマッチします。
'　\{       { にマッチします。
'　\}       } にマッチします。
'　\\       \ にマッチします。
'　\[       [ にマッチします。
'　\]       ] にマッチします。
'　\(       ( にマッチします。
'　\)       ) にマッチします。
'　[xyz]  文字セットに含まれている任意の1文字にマッチします｡
'　\w      単語に使用される任意の文字にマッチします。[a-zA-Z_0-9]と等価。
'　\W     単語に使用される文字以外の任意の文字にマッチします。[^a-zA-Z_0-9]と等価。
'　\d      任意の数字にマッチします。[0-9]と等価。
'　\D      任意の数字以外の文字にマッチします。[^0-9]と等価。
'　\s       任意のスペース文字にマッチします。[ \t\r\n\v\f]と等価。
'　\S      任意の非スペース文字にマッチします。[^ \t\r\n\v\f]と等価。
'</summary>
Public Sub InitRegex(pattern As String, Optional options As String = "")
    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.pattern = pattern
    RegExp.IgnoreCase = IIf(InStr(options, "IgnoreCase"), True, False)
    RegExp.MultiLine = IIf(InStr(options, "MultiLine"), True, False)
End Sub

'<summary>
'  指定文字列にパターンが含まれているかどうか
'
'　input_：チェックする文字列
'  [省略可]startAt：何文字目からチェックするか
'</summary>
Public Function IsMatch(input_ As String, Optional startAt As Integer = 1) As Boolean
    IsMatch = RegExp.test(Mid(input_, startAt))
End Function

'<summary>
'  指定文字列内にパターンが含まれていたら一致した最初の文字列だけ返す
'
'　input_：チェックする文字列
'  [省略可]startAt：何文字目からチェックするか
'</summary>
Public Function Match(input_ As String, Optional startAt As Integer = 1)
    RegExp.Global = False
    Set Match = RegExp.Execute(Mid(input_, startAt))
End Function

'<summary>
'  指定文字列内にパターンが含まれていたらitemコレクションとして返す。
'
'　input_：チェックする文字列
'  [省略可]startAt：何文字目からチェックするか
'</summary>
Public Function Matches(input_ As String, Optional startAt As Integer = 1)
    RegExp.Global = True
    Set Matches = RegExp.Execute(Mid(input_, startAt))
End Function

'<summary>
'  指定文字列内にパターンが含まれていたら指定文字列に置き換えて返す
'
'　input_：チェックする文字列
'  replacement：置き換え後の文字列
'</summary>
Public Function Replace(input_ As String, replacement As String)
    Replace = RegExp.Replace(input_, replacement)
End Function


' ---------------------------------------------------------------
' .NET Framework System.Diagnostics.Stopwatch のクローン
' ---------------------------------------------------------------

'<summary>
'  タイマーのスタート
'</summary>
Public Sub StartTimer()
    StartTime = GetTickCount()
End Sub

'<summary>
'  タイマーのストップ　返り値はプロパティから取得して
'</summary>
Public Sub StopTimer()
    EndTime = GetTickCount()
    AccumulatedTime = AccumulatedTime + EndTime - StartTime
End Sub

'<summary>
'  タイマーの経過時間リセット
'</summary>
Public Sub ResetTimer()
    AccumulatedTime = 0
End Sub

'<summary>
'  経過時間をリセットしつつ、新たにタイマースタート
'</summary>
Public Sub RestartTimer()
    StartTime = GetTickCount()
    AccumulatedTime = 0
End Sub


' ---------------------------------------------------------------
' .NET Framework System.Text.StringBuilderのクローン
' Appendでいくつか文字列を設定しつつ、ToStringで全て結合したものを返す
' ---------------------------------------------------------------

'<summary>
'  結合する文字列を設定する
'
'　value：設定する文字列
'</summary>
Public Sub Append(value As String)
    mIndex = mIndex + 1
    Call mBuffer.Add(CStr(mIndex), value)
End Sub

'<summary>
'  結合する文字列＋改行を設定する
'
'　[省略可]value：設定する文字列
'</summary>
Public Sub AppendLine(Optional value As String = "")
    mIndex = mIndex + 1
    Call mBuffer.Add(CStr(mIndex), value & vbCrLf)
End Sub

'<summary>
'  設定してた文字列をクリアする
'</summary>
Public Sub ClearDictionary()
    mIndex = 0
    mBuffer.RemoveAll
End Sub

'<summary>
'  設定した文字列を全て結合して返す
'</summary>
Public Function ToString() As String
    ToString = Join(mBuffer.items(), "")
End Function

























