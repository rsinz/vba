VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 最終更新日：2018/09/20  ＠rsinz

Dim adoCom As Object, adoRs As Object

'///////// index //////////
' OpenDB:           OracleDBと接続する
' CloseDB:           OracleDBとの接続を切る
' Mode:                システムモード取得
' Download:         sql 実行(返り値あり)
' Commit:            sql 実行(返り値なし)
' Header:             指定テーブルのヘッダーを配列で返す
' ColumnTypeOf:  指定テーブルのヘッダーのデータ型を返す
' ToSqlFormat:     SQL用の書式に変換する
' Exist:                データの存在チェック



'///////// module //////////

Private Sub Class_Initialize()
    Set adoCom = CreateObject("ADODB.connection")
    Set adoRs = CreateObject("ADODB.Recordset")
End Sub

Private Sub Class_Terminate()
    CloseDB
    Set adoRs = Nothing
    Set adoCom = Nothing
End Sub

'<summary>
'  OracleDBと接続する
'
'  ソース名 : データソース名
'  ID : 接続用ID
'  パスワード : 接続用パスワード
'</summary>
Sub OpenDB(ソース名 As String, ID As String, パスワード As String)
    adoCom.Open "DSN=" & ソース名 & ";UID=" & ID & ";PWD=" & パスワード & ";"
End Sub

'<summary>
'  OracleDBとの接続を切る
'</summary>
Sub CloseDB()
    
    ' 0 = adStateClosed;
    ' 2 = adStateConnecting;
    ' 4 = adStateExecuting;
    ' 8 = adStateFetching;
    If (adoCom.State = 0) Or _
        ((adoCom.State And 2) = 2) Or _
        ((adoCom.State And 4) = 4) Or _
        ((adoCom.State And 8) = 8) Then Exit Sub
        
    adoCom.Close
End Sub

'<summary>
'  エラーまとめ　0:NoError  それ以外:Error
'</summary>
Private Function ErrCheck() As Integer

    If (adoCom.State And 1) <> 1 Then      ' 1 = adStateOpen
        MsgBox "DBと接続していません。OpenDBメソッドを実行してください。", , "clsOracle"
        ErrCheck = -1
        Exit Function
    End If

End Function

'<summary>
'  システムモード取得
'</summary>
Function Mode() As Variant

    If ErrCheck <> 0 Then Exit Function
    
    Dim sql As String:
    sql = "SELECT * " & _
                "FROM モード情報テーブル " & _
                "ORDER BY 運用モード"
    
    Set adoRs = adoCom.Execute(sql)
    Mode = adoRs.GetRows()
    
End Function

'<summary>
'  ダウンロード
'
'  SQL : SQL構文
'</summary>
Function Download(sql As String) As Variant

    If ErrCheck <> 0 Then Exit Function

    Set adoRs = adoCom.Execute(sql)
    Download = adoRs.GetRows()

End Function

'<summary>
'  SQL 実行
'
'  SQL : SQL構文
'</summary>
Sub Commit(sql As String)

    On Error GoTo ErrorHandler

    If ErrCheck <> 0 Then Exit Sub
    adoCom.BeginTrans

    adoCom.Execute (sql)
    adoCom.CommitTrans
    
    GoTo Finally
ErrorHandler:
    adoCom.RollbackTrans
Finally:

End Sub


'<summary>
'  指定テーブルのヘッダーを配列で返す
'
'  テーブル名 : まんまテーブル名
'</summary>
Function Header(テーブル名 As String) As Variant

    If ErrCheck <> 0 Then Exit Function

    Dim sql As String
    sql = "SELECT *" & _
               " FROM " & テーブル名

    Set adoRs = adoCom.Execute(sql)
    
    Dim h As Variant: ReDim h(0 To adoRs.fields.Count - 1)
    For i = 0 To adoRs.fields.Count - 1
        h(i) = adoRs.fields(i).name
    Next i

    Header = h
End Function

'<summary>
'  指定テーブルのヘッダーのデータ型を返す
'
'  テーブル名 : まんまテーブル名
'</summary>
Function ColumnTypeOf(テーブル名 As String) As Variant

    If ErrCheck <> 0 Then Exit Function

    Dim sql As String
    sql = "SELECT DATA_TYPE " & _
            "FROM USER_TAB_COLUMNS " & _
            "WHERE TABLE_NAME = '" & テーブル名 & "' " & _
            "ORDER BY COLUMN_ID"

    
    Set adoRs = adoCom.Execute(sql)
    ColumnTypeOf = adoRs.GetRows()

End Function

'<summary>
'  SQL用の書式に変換する
'
'  変換する値 : ここに指定された値の書式を変換します
'  データ型 : この型に基づいて変換します
'</summary>
Function ToSqlFormat(変換する値 As Variant, データ型 As String) As String

    If IsEmpty(変換する値) Then ToSqlFormat = "NULL": Exit Function
        
    Select Case データ型
        Case "NUMBER": ToSqlFormat = CStr(変換する値)
        Case "VARCHAR2": ToSqlFormat = "'" & CStr(変換する値) & "'"
        Case "DATE": ToSqlFormat = "TO_DATE('" & CStr(変換する値) & "','YYYY/MM/DD HH24:MI:SS')"
        Case Else: ToSqlFormat = CStr(変換する値)
    End Select
    
End Function

'<summary>
'  データの存在チェック
'
'  テーブル名 : まんまテーブル名
'  [省略可] where : 条件文
'</summary>
Function Exist(テーブル名 As String, Optional where As String = "") As Boolean

    If ErrCheck <> 0 Then Exit Function

    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT * " & _
            "FROM " & テーブル名 & _
            IIf(where = "", "", " WHERE " & where)
    
    Set adoRs = adoCom.Execute(sql)
    Exist = Not IsNull(adoRs)
    
    GoTo Finally
ErrorHandler:
    Call Err.Raise(10000, "clsOracle.Exist", "SQL構文に間違いがあります")
Finally:
    
End Function

















