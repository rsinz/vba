VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSQLServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 2022/10/29 @rsinz

Private ado_ As Object

'///////// module list //////////

' OpenDB:      SQLServerと接続する｡
' CloseDB:      SQLServerとの接続を切る｡
' GetRows:     引数で指定されたSQLServerテーブルをダウンロードし､配列に変換｡
' Insert:         配列インサート
' Excute:        SQL発行


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
'　SQLServerファイルの完全パス：SQLServerファイルのフルパス
'</summary>
Sub OpenDB(server As String, db As String, user As String, pass As String)
    
    If (ado_.State And 1) = 1 Then CloseDB     ' 1 = adStateOpen
    
    ado_.Open "Driver={SQL Server}; " & _
                        "server=" & server & ";" & _
                        "database=" & db & ";" & _
                        "uid=" & user & ";" & _
                        "pwd=" & pass & ";"
    
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
'　引数で指定されたSQLServerテーブルをダウンロードし、配列に変換。エラーコード Null
'
'　query_ : SQL文を指定
'</summary>
Function GetRows(query_ As String) As Variant

    On Error GoTo ErrorHandler
    
    If (ado_.State And 1) <> 1 Then Err.Raise Number:=10000, Description:="DBと接続していません。Openメソッドを実行してください。"
    
    ' -- SQLServer からダウンロードします
    arr = ado_.Execute(query_).GetRows()
    
    ' -- Excel で扱いやすいように、行列を入れ替えます
    If HasNull(arr) Then arr = ReplaceNullTo("", arr)
    GetRows = WorksheetFunction.Transpose(arr)
    If ArrayRank(GetRows) = 1 Then GetRows = To2D(GetRows)       ' 配列が一次元の場合、二次元に変換しておくと後で楽
    
    GoTo Finally

ErrorHandler:
    GetRows = Empty
Finally:
End Function

'<summary>
'　配列インサート
'
'　arr : 配列
'   table : テーブル名
'</summary>
Sub Insert(arr As Variant, table As String)
 
    Dim query_: query_ = "INSERT INTO " & table & " VALUES "
    
    For i = LBound(arr) To UBound(arr)
        Dim line_: line_ = ""
        
        For j = LBound(arr, 2) To UBound(arr, 2)
            If IsEmpty(arr(i, j)) Or arr(i, j) = "NULL" Then
                line_ = line_ & "NULL" & IIf(j = UBound(arr, 2), "", ",")
            Else
                line_ = line_ & "'" & arr(i, j) & "'" & IIf(j = UBound(arr, 2), "", ",")
            End If
        Next j
        query_ = query_ & "(" & line_ & ")" & ","
    
        If i Mod 100 = 0 Then
            Call ado_.Execute(Left(query_, Len(query_) - 1) & ";")
            query_ = "INSERT INTO " & table & " VALUES "
        End If
    
    Next i
    
    Call ado_.Execute(Left(query_, Len(query_) - 1) & ";")
    
End Sub

'<summary>
'　SQL発行
'</summary>
Sub Excute(query_ As String)
    Call ado_.Execute(query_)
End Sub

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






















