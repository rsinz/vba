VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 2018/09/05 @rsinz

Private ado_ As Object

'///////// module list //////////

' OpenDB:       Accessと接続する｡
' CloseDB:       Accessとの接続を切る｡
' Download:    引数で指定されたAccessテーブルをダウンロードし､配列に変換｡



'///////// module //////////

Private Sub Class_Initialize()
    Set ado_ = CreateObject("ADODB.Connection")
End Sub

Private Sub Class_Terminate()
    CloseDB
    Set ado_ = Nothing
End Sub

'<summary>
'  ADODB.Connectionをopenする。
'
'　Accessファイルの完全パス：Accessファイルのフルパス
'</summary>
Sub OpenDB(ByVal Accessファイルの完全パス As String)
    
    If (ado_.State And 1) = 1 Then CloseDB     ' 1 = adStateOpen
    ado_.PROVIDER = "Microsoft.Ace.OLEDB.12.0"
    ado_.Open Accessファイルの完全パス
    
End Sub

'<summary>
'  ADODB.Connectionをcloseする。
'</summary>
Sub CloseDB()

    ' 0 = adStateClosed;
    ' 2 = adStateConnecting;
    ' 4 = adStateExecuting;
    ' 8 = adStateFetching;
    If (ado_.State = 0) Or _
        ((ado_.State And 2) = 2) Or _
        ((ado_.State And 4) = 4) Or _
        ((ado_.State And 8) = 8) Then Exit Sub
    
    ado_.Close

End Sub

'<summary>
'　引数で指定されたAccessテーブルをダウンロードし、配列に変換。エラーコード Null
'
'　テーブル名：データベース側のテーブル名
'　[省略可]抽出条件：パラメータ配列。Array関数で指定  Accessクエリのパラメータとリンクする。　規定値Null
'</summary>
Function Download(ByVal テーブル名 As String, Optional 抽出条件 As Variant = Null) As Variant

    On Error GoTo ErrorHandler
    
    If (ado_.State And 1) <> 1 Then Err.Raise Number:=10000, Description:="DBと接続していません。Openメソッドを実行してください。"
    
    ' -- Access からダウンロードします
    With CreateObject("ADODB.Command")
        .ActiveConnection = ado_
        .CommandText = テーブル名
        
        Dim arr As Variant
        If IsNull(抽出条件) Then
            arr = .Execute().GetRows()
        Else
            arr = .Execute(Parameters:=抽出条件).GetRows()
        End If
    End With
    
    ' -- Excel に貼り付けやすいよう、行列を入れ替えます
    If HasNull(arr) Then arr = ReplaceNullTo("", arr)
    Download = WorksheetFunction.Transpose(arr)
    If ArrayRank(Download) = 1 Then Download = To2D(Download)       ' 配列が一次元の場合、二次元に変換しておくと後で楽
    
    GoTo Finally

ErrorHandler:
    Download = Empty
Finally:
End Function

'<summary>
'　Private] 配列にNULLがあるか判定する。NULLがあるとTransposeでエラーが出るので、それを利用する
'
'　判定する配列：そのまんま
'</summary>
Private Function HasNull(判定する配列 As Variant) As Boolean

    On Error GoTo TruePattern
    Dim v As Variant: v = WorksheetFunction.Transpose(判定する配列)
    
    HasNull = False
    Exit Function

TruePattern:
    HasNull = True
End Function

'<summary>
'　private] 配列内のNullを引数値に置換。
'
'　置換後の値：置換後の値
'　配列：対象の配列
'</summary>
Private Function ReplaceNullTo(ByVal 置換後の値 As Variant, ByVal 配列 As Variant) As Variant
    
    For i = LBound(配列, 1) To UBound(配列, 1)
        For j = LBound(配列, 2) To UBound(配列, 2)
            If IsNull(配列(i, j)) Or IsEmpty(配列(i, j)) Then 配列(i, j) = 置換後の値
        Next j
    Next i

    ReplaceNullTo = 配列

End Function

'<summary>
'  private] 配列の次元数を返す。インデックス外の次元をUBoundするとエラーが出るので、それを利用
'
'　配列：次元数を調べる配列
'</summary>
Private Function ArrayRank(配列 As Variant) As Integer

    On Error Resume Next
    
    Do While Err.Number = 0
        i = i + 1
        errTest = UBound(配列, i)
    Loop
    
    On Error GoTo 0
    ArrayRank = i - 1
End Function

'<summary>
'  private] 1次元配列を2次元配列にする。
'  1次元配列と2次元配列が混合する場合に一様の処理を行いたいときとかに使う
'
'  配列：変換対象
'</summary>
Private Function To2D(配列 As Variant) As Variant

    Dim arr As Variant: ReDim arr(1 To 1, 1 To UBound(配列))
    
    For i = 1 To UBound(配列)
        If IsObject(配列(i)) Then
            Set arr(1, i) = 配列(i)
        Else
            arr(1, i) = 配列(i)
        End If
    Next i

    To2D = arr
End Function






















