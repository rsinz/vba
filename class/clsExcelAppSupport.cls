VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelAppSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 最終更新日：2018/02/17  ＠しら


Private Declare PtrSafe Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As Long
Private Const HIGH_PRIORITY_CLASS = &H80
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const REALTIME_PRIORITY_CLASS = &H100
Private cn As Object, rs As Object



'///////// enumerated type //////////

Public Enum BoderIndex          ' AddBorderで使用。罫線を引く場所を指定するため
    biGrid = 0
    biAround = 1
    biBottom = 2
    biTop = 3
    biLeft = 4
    biRight = 5
    biInsideGrid = 6
    biInsideHorizontal = 7
    biInsideVertical = 8
    biCross = 9
    biLBottomToRTop = 10
    biLTopToRBottom = 11
End Enum

'///////// class property //////////



'///////// module list //////////

' Preheart:                      Excel アプリケーションの設定を処理優先にする｡マクロ実行前にやると､処理が早くなるかも｡
' Cooldown:                    Excel アプリケーションの設定を元に戻す｡
' Boost:                          Excel のプロセス優先度を高に設定
' Idle:                             Excel のプロセス優先度を通常に設定
' AddBook:                      Excelを開く
' CloseBook:                    Excelを閉じる　任意で保存も可
' StopAutoFunction:        Excelの自動機能を停止する
' ReStartAutoFunction:    Excelの自動機能を再開する
' ArrayToSheet:               配列をExcelシートに貼り付ける
' SetHyperlink:                Excelシートにハイパーリンクを設定する
' LastRow:                      Rangeを基準とした最下位行を返す
' LastCol:                        Rangeを基準とした最右列を返す
' ExistsSheet:                  シートの存在チェック
' AddSheet:                    シートの追加
' ----ここまで開発確認
' CopySheet:                   シートのコピー
' DeleteSheet                  シートの削除
' ChangeSheetName：     シートの名前変更
' ClearRowsValue:           範囲行の「値」削除
' ClearColumnsValue:      範囲列の「値」削除
' ParentName:                指定されたRange名があるシートの名前を返す
' Paint:                          セルの塗りつぶし
' AddBorder：                  罫線追加
' CreatePivot:                  ピボットテーブルの作成
' PivotContents:               ピボットテーブルのデータ部を配列に変換
' ShapeProperty:             オートシェイプ(図形とかテキストボックス)のプロパティ変更
' SetPrintArea:                 印刷範囲の設定
' PrintOut:                        印刷する
' Merge:                           複数範囲セル結合
' UnMerge:                      複数範囲セル結合解除
' FilterOn:                        指定Rangeに完全一致フィルターをかける
' FilterOff:                        指定シートのフィルターを解除
' Execute:                       SQLの実行




'///////// module //////////

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub

'<summary>
'　Excel アプリケーションの設定を処理優先にする。マクロ実行前にやると、処理が早くなるかも。
'</summary>
Sub Preheart()
    Call Boost
    Call StopAutoFunction
End Sub

'<summary>
'　Excel アプリケーションの設定を元に戻す。
'</summary>
Sub Cooldown()
    Call ReStartAutoFunction
    Call Idle
End Sub

'<summary>
'　Excel のプロセス優先度を高に設定
'</summary>
Sub Boost()
    Call SetPriorityClass(GetCurrentProcess, HIGH_PRIORITY_CLASS)
End Sub

'<summary>
'　Excel のプロセス優先度を通常に設定
'</summary>
Sub Idle()
    Call SetPriorityClass(GetCurrentProcess, NORMAL_PRIORITY_CLASS)
End Sub

'<summary>
'　Excelを開く。
'
'　[省略可]filePath：指定パスのBookをテンプレートとして開く。省略された場合、新規Bookを開く
'　[省略可]forcusAddedBook : 開いたBookをアクティブにするかどうか　True:アクティブにする
'</summary>
Function AddBook(Optional filePath As String = "", Optional forcusAddedBook As Boolean = True) As Workbook

    On Error GoTo ErrorHandler

    Dim parent_ As Workbook: Set parent_ = ThisWorkbook

    If filePath = "" Then       ' ファイルパスの指定が無ければ新規Bookを開く
        Set AddBook = Workbooks.Add
        
    Else
        If Dir(filePath, vbDirectory) = "" Then Err.Raise 10000, , "指定されたパスにファイルがありませんでした"
        
        If IsBookOpened(filePath) Then        ' Book が既に開かれていた場合
'
'            With CreateObject("Scripting.FileSystemObject")
'                Dim fileName_ As String
'                Call .GetFileName(filePath, fileName_)
'            End With
'
            Set AddBook = Workbooks(GetFileName(filePath))
            AddBook.Activate
            
        Else
            Set AddBook = Workbooks.Open(filePath)
            
        End If
        
    End If

    If forcusAddedBook = False Then parent_.Activate

    GoTo Finally
    
ErrorHandler:
    Debug.Print "clsExcelAppSupport.AddBook : " & Err.Description
    Set AddBook = Nothing
    
Finally:
End Function

Private Function IsBookOpened(filePath As String) As Boolean

    On Error Resume Next
    
    Open filePath For Append As #1
    Close #1
    
    IsBookOpened = (Err.Number > 0)
    
End Function

Private Function GetFileName(filePath As String) As String
    Dim s As Variant: s = Split(filePath, Application.PathSeparator)
    GetFileName = s(UBound(s))
End Function

'<summary>
'　Excelを閉じる。
'
'　bookName：閉じる対象のBookの名前
'　isSaveChanged：保存するかどうか
'　[省略可]　savePath：保存先のパス。デフォルト上書き保存
'　[省略可]　saveFormat：保存形式。デフォルト.xls
'</summary>
Sub CloseBook(bookName As String, isSaveChanged As Boolean, Optional savePath As String = "", Optional format As XlFileFormat = 0)

    On Error GoTo ErrorHandler

    If Not isSaveChanged Then
        Workbooks(bookName).Close       '保存しないで閉じる
    Else
            If savePath = "" Then
                Workbooks(bookName).Save    '上書き保存して閉じる
                Workbooks(bookName).Close
            Else
                    Dim split_path As Variant: split_path = Split(savePath, "\")
                    
                    If format = 0 Then
                        Workbooks(bookName).SaveAs fileName:=savePath   '任意の保存先に保存して閉じる
                        Workbooks(split_path(UBound(split_path))).Close
                    Else
                        Workbooks(bookName).SaveAs fileName:=savePath, fileFormat:=format  '任意の保存先に保存して閉じる(フォーマット指定)
                        Workbooks(split_path(UBound(split_path))).Close
                    End If
                    
            End If
    End If


    GoTo Finally

ErrorHandler:
    Debug.Print "clsExcelAppSupport.CloseBook : " & Err.Description

Finally:

End Sub


'<summary>
'　Excelの自動機能を停止する
'</summary>
Sub StopAutoFunction()

    With Application
'        .Visible = False                        '全体の表示を停止
        .WindowState = xlMinimized   'ブック最小化
        .DisplayAlerts = False              'アラートの表示を停止
        .statusBar = False                   'ステータスバーの表示更新を停止
        .ScreenUpdating = False         'スクリーンの描画を停止
        .EnableEvents = False             'イベントを一時停止
        .Calculation = xlManual           '計算を手動モードにする
    End With
    
End Sub

'<summary>
'　Excelの自動機能を再開する
'</summary>
Sub ReStartAutoFunction()

    With Application
        .Calculation = xlAutomatic      '計算を自動モードに戻す
        .EnableEvents = True              'イベントを再開
        .ScreenUpdating = True          'スクリーンの描画を再開
        .statusBar = True                    'ステータスバーの表示を再開
        .DisplayAlerts = True               'アラートの表示を再開
        .WindowState = xlMaximized  'ブック最大化
'        .Visible = True                         '全体の表示を回復
    End With

End Sub

'<summary>
'　配列をExcelシートに貼り付ける
'
'　baseRange：基点のRange
'　arr：貼り付ける配列
'</summary>
Sub ArrayToSheet(baseRange As range, arr As Variant)

    Dim ws As Worksheet: Set ws = Worksheets(baseRange.Parent.name)
    Dim wb As Workbook: Set wb = Workbooks(ws.Parent.name)
    Dim baseRow As Long: baseRow = baseRange.Row
    Dim baseColumn As Long: baseColumn = baseRange.column
    
    With wb.Worksheets(ws.name)
        .Activate
        .range(Cells(baseRow, baseColumn), Cells(baseRow + UBound(arr, 1) - LBound(arr, 1), baseColumn + UBound(arr, 2) - LBound(arr, 2))) = arr
    End With

End Sub

'<summary>
'　Excelシートにハイパーリンクを設定する
'
'　baseRange：貼り付けるRange
'  link：アドレス文字列
'</summary>
Sub SetHyperlink(baseRange As range, link As String)
    
    Dim ws As Worksheet: Set ws = Worksheets(baseRange.Parent.name)
    Dim wb As Workbook: Set wb = Workbooks(ws.Parent.name)
    
    With wb.Worksheets(ws.name)
        .Activate
        .Hyperlinks.Add rng, link
    End With
End Sub

'<summary>
'  baseRangeを基準とした最下位行を返す
'
'　baseRange：基準となるRange
'　[省略可]  isIgnoreBlank：空白行を無視するか,　True：空白行お構いなしで最終データ行を返す
'</summary>
Function LastRow(ByVal baseRange As range, Optional isIgnoreBlank As Boolean = False) As Long

    If isIgnoreBlank Then
        LastRow = baseRange.Worksheet.Cells(baseRange.Worksheet.Rows.Count, baseRange.column).End(xlUp).Row
    Else
        LastRow = baseRange.End(xlDown).Row
    End If

End Function

'<summary>
'  baseRangeを基準とした最右列を返す
'
'　baseRange：基準となるRange
'　[省略可]  isIgnoreBlank：空白行を無視するか,　True：空白行お構いなしで最終データ列を返す
'</summary>
Function LastCol(ByVal baseRange As range, Optional isIgnoreBlank As Boolean = False) As Long

    If isIgnoreBlank Then
        LastCol = baseRange.Worksheet.Cells(baseRange.Row, baseRange.Worksheet.Columns.Count).End(xlToLeft).column
    Else
        LastCol = baseRange.End(xlToRight).column
    End If

End Function

'<summary>
'  sheetの存在チェック
'
'　wsName：チェック対象シートの名前
'　[省略可]  wb：チェック対象Bookオブジェクト
'</summary>
Function ExistsSheet(ByVal wsName As String, Optional ByVal wb As Workbook = Nothing) As Boolean
    
    Dim ws As Worksheet
    
    If TypeName(wb) = "Nothing" Then Set wb = Application.ThisWorkbook
    
    For Each ws In wb.Sheets
        If ws.name = wsName Then ExistsSheet = True: Exit Function
    Next

End Function

'<summary>
'  [private] sheetの名前が重複しないように、末尾に数字を付けた名前を返す
'
'　wsName：追加しようとしてるシートの名前
'　[省略可] wb：追加するブックオブジェクト デフォルト：呼び出し元Book
'　[省略可] count：末尾の数字。再起用 デフォルト：0
'</summary>
Private Function SheetNameRaps(wsName As String, Optional wb As Workbook = Nothing, Optional Count As Integer = 0) As String

    If TypeName(wb) = "Nothing" Then Set wb = Application.ThisWorkbook
    
    Dim ws As Worksheet
    Dim name As String: name = IIf(Count = 0, wsName, wsName & "(" & CStr(Count) & ")")
    
    If ExistsSheet(name, wb) = True Then
        SheetNameRaps = SheetNameRaps(wsName, wb, Count + 1)
    Else
        SheetNameRaps = name
    End If

End Function

'<summary>
'  sheetの追加
'
'　[省略可] wsName：追加するシートの名前 デフォルト：自動
'　[省略可] wb：追加するブックオブジェクト デフォルト：呼び出し元Book
'</summary>
Function AddSheet(Optional ByVal wsName As String = "", Optional ByVal wb As Workbook = Nothing) As Worksheet

    On Error GoTo ErrorHandler
    
    If TypeName(wb) = "Nothing" Then Set wb = Application.ThisWorkbook
'    If ExistsSheet(wsName, wb) Then Err.Raise 10000, , "シートを追加しようとしましたが、既に同名のシートが存在します"
    
    Dim ws As Worksheet: Set ws = wb.Worksheets.Add()
    If wsName <> "" Then ws.name = SheetNameRaps(wsName, wb)
    Set AddSheet = ws
    
    GoTo Finally

ErrorHandler:
    Debug.Print "clsExcelAppSupport.AddSheet : " & Err.Description
    Set AddSheet = Nothing

Finally:
End Function

'<summary>
'  sheetのコピー
'
'  sourceSheetName：コピー元のシート名
'　afterWsName：コピー後のシート名
'　[省略可]  wb：追加するブックオブジェクト
'</summary>
Function CopySheet(ByVal SourceSheetName As String, ByVal afterWsName As String, Optional ByVal wb As Workbook = Nothing) As Worksheet

    On Error GoTo ErrorHandler
    
    If TypeName(wb) = "Nothing" Then Set wb = Application.ThisWorkbook
    
    If Not ExistsSheet(SourceSheetName, wb) Then Err.Raise 10000, , SourceSheetName & vbNewLine & "というシートが存在しません"
    If ExistsSheet(afterWsName, wb) Then Err.Raise 10000, , afterWsName & vbNewLine & "というシートは既に存在しています"
    
    wb.Worksheets(SourceSheetName).Copy after:=wb.Worksheets(SourceSheetName)
    wb.ActiveSheet.name = afterWsName
    
    Set CopySheet = wb.ActiveSheet
        
    GoTo Finally

ErrorHandler:
    MsgBox Err.Description, vbOKOnly + vbExclamation, "エラーが発生しました : clsExcelAppSupport.CopySheet"
    CopySheet = Nothing

Finally:
End Function

'<summary>
'  sheetの削除
'
'  wsName：削除するシート名
'　[省略可]  wb：削除するブックオブジェクト
'</summary>
Sub DeleteSheet(ByVal wsName As String, Optional ByVal wb As Workbook = Nothing)

    On Error GoTo ErrorHandler
    
    If TypeName(wb) = "Nothing" Then Set wb = Application.ThisWorkbook
    
    If Not ExistsSheet(wsName, wb) Then Err.Raise 10000, , wsName & vbNewLine & "というシートが存在しません"
    
    Worksheets(wsName).Delete
        
    GoTo Finally

ErrorHandler:
    MsgBox Err.Description, vbOKOnly + vbExclamation, "エラーが発生しました : clsExcelAppSupport.DeleteSheet"

Finally:
End Sub

'<summary>
'  sheetの名前変更
'
'  beforeName：変更するシート名
'　afterName：変更後のシート名
'</summary>
Sub ChangeSheetName(beforeName As String, afterName As String)

    If beforeName = "" Or afterName = "" Then MsgBox "ChangeSheetName:シート名を指定してください": Exit Sub
    If ExistsSheet(beforeName) = False Then MsgBox "ChangeSheetName:シートが存在しませんでした": Exit Sub

    Worksheets(beforeName).name = afterName

End Sub

'<summary>
'  firstRowNum～lastRowNumの範囲行の「値」削除
'
'  ws：対象シート
'  firstRowNum：削除開始行数
'　[省略可]  lastRowNum：削除終端行数。指定がなければアプリケーション上の最終行
'</summary>
Sub ClearRowsValue(ws As Worksheet, firstRowNum As Long, Optional lastRowNum As Long = 0)

    If lastRowNum = 0 Then lastRowNum = Rows.Count
    
    If (firstRowNum = 0) Or (firstRowNum > lastRowNum) Then Exit Sub
    
    ws.Activate
    ws.range(Rows(firstRowNum), Rows(lastRowNum)).ClearContents

End Sub

'<summary>
'  firstColumnNum～lastColumnNumの範囲列の「値」削除
'
'  ws：対象シート
'  firstColumnNum：削除開始列数
'　[省略可]  lastColumnNum：削除終端列数。指定がなければアプリケーション上の最終列
'</summary>
Sub ClearColumnsValue(ws As Worksheet, firstColumnNum As Long, Optional lastColumnNum As Long = 0)

    If lastColumnNum = 0 Then lastColumnNum = Columns.Count
    
    If (firstColumnNum = 0) Or (firstColumnNum > lastColumnNum) Then Exit Sub
    
    ws.Activate
    ws.range(Columns(firstColumnNum), Columns(lastColumnNum)).ClearContents

End Sub

'<summary>
'  指定されたRange名があるシートの名前を返す
'
'  rngName：Rangeの名前
'</summary>
Private Function ParentName(rngName As String)

    Dim sheetName As String
    
    For Each nm In ActiveWorkbook.Names
        If nm.name = rngName Then
            sheetName = MidInString(targetStr:=nm.value, pointerStr:="=")
            sheetName = MidInString(targetStr:=sheetName, pointerStr:="!", isAfter:=False)
            Exit For
        End If
    Next

    If sheetName = "" Then MsgBox rngName & " という名前のRangeが見つかりませんでした。" & vbNewLine & _
                                                        "数式 -> 名前の管理 を確認してください。"
    ParentName = sheetName
    
End Function

'<summary>
'  セルの塗りつぶし
'
'  targetRange：塗りつぶすRange
'　[省略可]  colors：色　デフォルト黄色
'</summary>
Sub Paint(targetRange As range, Optional colors As XlRgbColor = rgbYellow)
    targetRange.Interior.color = colors
End Sub

'<summary>
'  罫線追加
'
'  targetRange：罫線を書くRange
'　[省略可]  boderPlace：罫線を引く場所　デフォルト：周囲中央
'　[省略可]  style：実線点線破線とか　       デフォルト：実線
'　[省略可]  weights：太さ　                     デフォルト：極細
'　[省略可]  colors：色　                          デフォルト：黒
'</summary>
Sub AddBorder(targetRange As range, _
                        Optional boderPlace As BoderIndex = biGrid, _
                        Optional style As XlLineStyle = xlContinuous, _
                        Optional weights As XlBorderWeight = xlThin, _
                        Optional colors As XlRgbColor = rgbBlack)

    With targetRange
        Select Case boderPlace
            Case 0 To 2: GoTo biBottom
            Case 3: GoTo biTop
            Case 4: GoTo biLeft
            Case 5: GoTo biRight
            Case 6 To 7: GoTo biInsideHorizontal
            Case 8: GoTo biInsideVertical
            Case 9 To 10: GoTo biLBottomToRTop
            Case 11: GoTo biLTopToRBottom
            Case Else: MsgBox ("error:AddBorder" & vbNewLine & "BoderIndexの範囲外です"): Exit Sub
        End Select
        
biBottom:
        Call ModifyBorderProperty(.borders(xlEdgeBottom), style, weights, colors)
        If boderPlace = biBottom Then Exit Sub
biTop:
        Call ModifyBorderProperty(.borders(xlEdgeTop), style, weights, colors)
        If boderPlace = biTop Then Exit Sub
biLeft:
        Call ModifyBorderProperty(.borders(xlEdgeLeft), style, weights, colors)
        If boderPlace = biLeft Then Exit Sub
biRight:
        Call ModifyBorderProperty(.borders(xlEdgeRight), style, weights, colors)
        If boderPlace = biRight Or boderPlace = biAround Then Exit Sub
biInsideHorizontal:
        Call ModifyBorderProperty(.borders(xlInsideHorizontal), style, weights, colors)
        If boderPlace = biInsideHorizontal Then Exit Sub
biInsideVertical:
        Call ModifyBorderProperty(.borders(xlInsideVertical), style, weights, colors)
        If boderPlace = biInsideVertical Or boderPlace = biInsideGrid Or boderPlace = biGrid Then Exit Sub
biLBottomToRTop:
        Call ModifyBorderProperty(.borders(xlDiagonalUp), style, weights, colors)
        If boderPlace = biLBottomToRTop Then Exit Sub
biLTopToRBottom:
        Call ModifyBorderProperty(.borders(xlDiagonalDown), style, weights, colors)
    End With



End Sub

Private Sub ModifyBorderProperty(targetBorder As Object, _
                                                        style As XlLineStyle, _
                                                        weights As XlBorderWeight, _
                                                        colors As XlRgbColor)
    With targetBorder
        .color = colors: .Weight = weights: .LineStyle = style
    End With
End Sub

'<summary>
'  ピボットテーブルの作成　新しいシートを作ってそこにピボットテーブルを作成します
'
'  sourceRange：元データRange範囲
'  pivotName：ピボットテーブルの名前
'　[省略可]　rowField：テーブル内で行に設定する列名　例：Array("H1","H2")　→H1とH2を行に設定
'　[省略可]　colField：テーブル内で列に設定する列名　使い方は同上
'　[省略可]　dataField：テーブル内で値に設定する列名　使い方は同上
'</summary>
Sub CreatePivot(sourceRange As range, _
                            pivotName As String, _
                            Optional rowField As Variant = Empty, _
                            Optional colField As Variant = Empty, _
                            Optional dataField As Variant = Empty)

    If CheckPivotField(rowField) = False Then Exit Sub
    If CheckPivotField(colField) = False Then Exit Sub
    If CheckPivotField(dataField) = False Then Exit Sub

    AddSheet
    
    Dim pvt As PivotTable
    Set pvt = ActiveWorkbook.PivotCaches.Add( _
                                    sourceType:=xlDatabase, _
                                    SourceData:=sourceRange _
                                    ).CreatePivotTable(TableDestination:=range("A3"), tableName:=pivotName)

    With pvt
        If VarType(rowField) = vbArray + vbVariant Then
            For Each rowData In rowField
                .PivotFields(rowData).Orientation = xlRowField
            Next
        End If
        
        If VarType(colField) = vbArray + vbVariant Then
            For Each colData In colField
                .PivotFields(colData).Orientation = xlColumnField
            Next
        End If

        If VarType(dataField) = vbArray + vbVariant Then
            For Each oneData In dataField
                .PivotFields(oneData).Orientation = xlDataField
            Next
        End If
    End With

End Sub

Private Function CheckPivotField(fieldData As Variant) As Boolean

    If VarType(fieldData) <> vbEmpty And VarType(fieldData) <> vbArray + vbVariant Then
        MsgBox "CreatePivot: フィールド指定はArray型にしてください" & vbNewLine & "例：Array(""Header1"",""Header2"")"
        CheckPivotField = False
        Exit Function
    ElseIf VarType(fieldData) = vbArray + vbVariant Then
        For Each oneData In fieldData
            If VarType(oneData) <> vbString Then
                MsgBox "CreatePivot: フィールド要素は文字列にしてください" & vbNewLine & "例：Array(""Header1"",""Header2"")"
                CheckPivotField = False
                Exit Function
            End If
        Next
    End If
    
    CheckPivotField = True
End Function

'<summary>
'  ピボットテーブルのデータ部を配列に変換
'
'  pivotName：ピボットテーブルの名前
'　[省略可] wsName：ピボットがあるシート名　デフォルト：アクティブシート
'</summary>
Function PivotContents(pivotName As String, Optional ByVal wsName As String = "") As Variant

    If wsName = "" Then
        wsName = ActiveSheet.name
    ElseIf ExistsSheet(wsName) = False Then
        MsgBox "PivotContents:" & wsName & "シートが存在しませんでした"
        Exit Function
    End If
        
    Set DataArea = Worksheets(wsName).PivotTables(pivotName).TableRange1
  
    Set DataArea = DataArea.Offset(1, 0)
    Set DataArea = DataArea.Resize(DataArea.Rows.Count - 1)
  
    PivotContents = DataArea
End Function

'<summary>
'  オートシェイプ(図形とかテキストボックス)のプロパティ変更
'
'  shapeName：オートシェイプの名前
'　[省略可] wsName：オートシェイプがあるシート名　デフォルト：アクティブシート
'　[省略可] text：入力したい文字列
'　[省略可] backColor：変更後の背景色
'</summary>
Sub ShapeProperty(shapeName As String, Optional ByVal wsName As String = "", _
                                            Optional Text As String = "  ", Optional backColor As XlRgbColor = -1)

    If wsName = "" Then wsName = ActiveSheet.name
    If ExistsSheet(wsName) = False Then
        MsgBox "ShapeProperty: " & wsName & "シートが見つかりませんでした"
        Exit Sub
    End If

    Set obj = Worksheets(wsName).Shapes(shapeName)

    If Text <> "  " Then
        obj.OLEFormat.Object.Text = Text
    End If
    
    If backColor > 0 And backColor < 16777216 Then        ' 16777215 = rgbWhite でMax
        Dim RED As Integer: RED = backColor Mod 256
        Dim green As Integer: green = Int(backColor / 256) Mod 256
        Dim blue As Integer: blue = Int(backColor / 256 / 256)
        obj.Fill.ForeColor.RGB = RGB(RED, green, blue)
    End If

End Sub

'<summary>
'  印刷範囲の設定
'
'  rangeArea：印刷範囲を設定するRange範囲
'　[省略可] wsName：設定するシート名　デフォルト：アクティブシート
'</summary>
Sub SetPrintArea(rangeArea As range, Optional wsName As String = "")

    If wsName = "" Then wsName = ActiveSheet.name
    If ExistsSheet(wsName) = False Then
        MsgBox "SetPrintArea: " & wsName & "シートが見つかりませんでした"
        Exit Sub
    End If
    
    Worksheets(wsName).PageSetup.PrintArea = rangeArea.address
End Sub

'<summary>
'  印刷範囲のアドレス取得
'
'  ws：印刷範囲が設定されているシートオブジェクト
' [省略可] style : A1形式 or R1C1形式
'</summary>
Function PrintAddress(ws As Worksheet, Optional style As XlReferenceStyle = XlReferenceStyle.xlA1) As String
    PrintAddress = range(ws.PageSetup.PrintArea).address(, , style)
End Function

'<summary>
'  印刷範囲のレンジ取得　っていうか Range("Print_Area")でいんじゃね？
'
'  ws：印刷範囲が設定されているシートオブジェクト
'</summary>
Function PrintRange(ws As Worksheet) As range
    Set PrintRange = ws.range(PrintAddress(ws))
End Function

'<summary>
'  印刷する
'
'　[省略可] wsName：設定するシート名　デフォルト：アクティブシート
'　[省略可] isColouring：カラー印刷するか　デフォルト：白黒
'　[省略可] verticalOrHorizon：印刷の向き　デフォルト：縦
'　[省略可] paperSize：印刷紙の大きさ　デフォルト：A4
'　[省略可] paperCount：印刷部数　デフォルト：1部
'　[省略可] isPreview：印刷プレビューを表示するか　デフォルト：表示なしでそのまま印刷
'</summary>
Sub PrintOut(Optional wsName As String = "", Optional isColouring As Boolean = False, _
                    Optional verticalOrHorizon As XlPageOrientation = XlPageOrientation.xlPortrait, _
                    Optional paperSize As XlPaperSize = XlPaperSize.xlPaperA4, _
                    Optional paperCount As Integer = 1, _
                    Optional isPreview As Boolean = False)

    If wsName = "" Then wsName = ActiveSheet.name
    If ExistsSheet(wsName) = False Then
        MsgBox "PrintOut: " & wsName & "シートが見つかりませんでした"
        Exit Sub
    End If

    With Worksheets(wsName).PageSetup
        .BlackAndWhite = isColouring
        .Orientation = verticalOrHorizon
        .paperSize = paperSize
    End With

    If isPreview Then
        Worksheets(wsName).PrintPreview
    Else
        Worksheets(wsName).PrintOut Copies:=paperCount
    End If
    
End Sub

'<summary>
'  複数範囲セル結合
'
'  ranges：結合するRange範囲 複数指定可
'</summary>
Sub Merge(ParamArray ranges())
    For i = 0 To UBound(ranges)
        ranges(i).Merge
    Next i
End Sub

'<summary>
'  複数範囲セル結合解除
'
'  ranges：結合解除するRange範囲 複数指定可
'</summary>
Sub UnMerge(ParamArray ranges())
    For i = 0 To UBound(ranges)
        ranges(i).UnMerge
    Next i
End Sub

'<summary>
'  指定Rangeに完全一致フィルターをかける
'
'  ranges：フィルターをかけるヘッダーRange
'  [省略可]column：フィルターをかける列数
'  [省略可]cmd：条件
'</summary>
Sub FilterOn(ranges As range, Optional column As Long = 0, Optional cmd As Variant = "")
    ranges.AutoFilter Field:=column, Criteria1:=cmd
End Sub

'<summary>
'  指定シートのフィルターを解除
'
'  sheetName：解除するシートの名前
'</summary>
Sub FilterOff(sheetName As String)
    Worksheets(sheetName).AutoFilterMode = False
End Sub

'<summary>
'  baseRangeを基点としたレンジ範囲のアドレス文字列を返す
'
'  baseRange：基点のレンジ
'  [省略可] style : 形式。A1 か、R1C1。規定値 A1形式
'  [省略可] ignoreRowBlank : 空白無視。規定値 行方向の空白を無視するテーブル範囲
'  [省略可] ignoreColBlank : 空白無視。規定値 列方向の空白を無視するテーブル範囲
'</summary>
Function RangeAddress(baseRange As range, Optional style As XlReferenceStyle = XlReferenceStyle.xlA1, _
                                        Optional ignoreRowBlank As Boolean = False, Optional ignoreColBlank As Boolean = False) As String
    
    Dim myLastRow As Long: myLastRow = LastRow(baseRange, ignoreRowBlank)
    Dim myLastCol As Long: myLastCol = LastCol(baseRange, ignoreColBlank)
    
    RangeAddress = range(Cells(baseRange.Row, baseRange.column), Cells(myLastRow, myLastCol)).address(, , style)
    
End Function


'<summary>
'  SQLの実行

'  sql : SQL構文
'  [省略可] Book : ソースになるワークブック

' SQL例1 フィルター：SELECT F1,F3 FROM [Sheet1$B3:E20] WHERE F1 = 'Man' and (F2 >= 20 and F2 < 30) ORDER BY F4, F3 DESC
' SQL例2 リレー：SELECT * FROM [Sheet2$],[Sheet3$] WHERE [Sheet2$].F2 = [Sheet3$].F1 ORDER BY [Sheet2$].F1
' ※F○は列名
'</summary>
Function Execute(sql As String, Optional Book As Workbook = Nothing) As Variant

    On Error GoTo ErrorHandler
    
    ' -- レコードセット設定
    If Book Is Nothing Then Set Book = ThisWorkbook
    If cn Is Nothing Then Call OpenSource(Book)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open Source:=sql, ActiveConnection:=cn
    
    ' -- 配列変換
    Dim arr As Variant: arr = rs.GetRows()
    If HasNull(arr) Then Call ReplaceNullTo("", arr)        ' RecordsetにNULLが含まれてると、Transposeでエラーになりますので、空白に置き換えます。
    arr = WorksheetFunction.Transpose(arr)
    
    Execute = arr
    
    GoTo Finally
    
ErrorHandler:
    If rs.State = 0 Then GoTo Finally       ' 0 = Close状態
    If rs.EOF Then GoTo Finally
    MsgBox Err.Description, vbOKOnly + vbExclamation, "エラーが発生しました : clsExcelAppSupport.Execute"
Finally:
End Function

'<summary>
'  private] Excelデータソースと接続してみる

'  cn : ADODB.Connection オブジェクト
'  wb : ソースになるワークブック
'</summary>
Private Sub OpenSource(wb As Workbook)

    Const PROVIDER    As String = "PROVIDER=MSDASQL;"
    Const DRIVER        As String = "DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};"
    Const path            As String = "DBQ="
    
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = PROVIDER & DRIVER & path & wb.path & "\" & wb.name & ";"
    cn.Open
End Sub

'<summary>
'  Private]  ソースとの接続を閉じる。

'  cn : ADODB.Connection オブジェクト
'</summary>
Private Sub CloseSource(cn As Object)
    If (cn.State = 1) Then cn.Close
End Sub


'<summary>
'　Private] 配列にNULLがあるか判定する。NULLがあるとTransposeでエラーが出るので、それを利用する
'
'　判定する配列：そのまんま
'</summary>
Private Function HasNull(判定する配列 As Variant) As Boolean

    On Error GoTo ArrayHasNull
    Dim v As Variant: v = WorksheetFunction.Transpose(判定する配列)
    
    HasNull = False
    Exit Function

ArrayHasNull:
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













