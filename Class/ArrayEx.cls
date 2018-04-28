VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 最終更新日：2018/04/28  ＠rsinz
' メソッドチェーン対応　内部変数default_追加。メソッドの返り値をArrayEx(Me)に変更。
' Is1D/ Is2D/ IsJagg/ ArrayRank/ Exist / SearchIndex をプロパティ化

'///////// memo //////////
' オブジェクト内エラーはイミディエイトウィンドウに表示します
' 3次元以上の配列は推奨しません
' For each で回す際は.Value を使ってください。ウォッチウィンドウでクラッシュする事象を回避できなかったです。NewEnumの中で内部関数を使うとダメみたい


'///////// property list //////////
' value :                   [Let] デフォルトプロパティ
' SetTable :              [Let] 表形式になっているテーブルデータをセットします
' CurrentRegion :      [Let] 表形式になっているテーブルデータをセットします。SetTable よりお手軽だけど、使える状況に制限があります
' value :                   [Get] デフォルトプロパティ
' MinRow :               [Get] 行の最小インデックス
' MaxRow :              [Get] 行の最大インデックス
' MinColumn :          [Get] 列の最小インデックス
' MaxColumn :          [Get] 列の最大インデックス
' Count :                  [Get] 行または列の総数
' IsEmpty_ :            [Get] 内部変数が空だったらTRUE
' Is1D :                    [Get] 1 次元配列だったらTRUE
' Is2D :                    [Get] 2 次元配列だったらTRUE
' IsJagg :                 [Get] ジャグ配列だったらTRUE
' ArrayRank :           [Get] 配列の次元数
' Exist :                   [Get] 配列内での指定値存在チェック
' Average_ :            [Get] 平均値の算出
' Sum_ :                  [Get] 合計値の算出
' Median_ :              [Get] 中央値の算出
' Max_ :                  [Get] 最大値の算出
' Min_ :                   [Get] 最小値の算出
' Mode_ :                [Get] 最頻値の算出。最頻値は複数あるかもしれないので、返り値は配列にしています。
' StDev_ :               [Get] 標準偏差の算出
' CPK_ :                  [Get] CPKの算出
' SearchIndex :        [Get] 配列内の検索。指定列を順方向・行方向に一致をみる。返り値 値が一致した添え字


'///////// module list //////////
' Init :                   内部変数 array_を初期値に戻します。繰り返し ArrayEx を使う場合に呼び出します。
' To1D:                  2次元配列を1次元配列にする。
' To2D:                  1次元配列を2次元配列にする。
' FirstIndexTo:       最初のインデックスを引数 indexに強制する
' Remove :             配列内から指定値の要素ごと削除(配列が圧縮される可能性がある)。
' ReplaceNullTo:    配列内のNull/Empty を引数値に置換
' Filter_:                配列のフィルター結果を返す｡
' Append:              配列拡張
' Slicer:                  配列から指定の１行または１列を取り出す。返り値 1次元配列
' Select_:              指定された列の取り出し
' Union:                 配列の結合
' GroupBy:              配列内の行データ重複削除
' ToSheet:              Excelシートに配列データを貼り付ける
' Execute:              SQLの実行
' FROM:                  SQL.FROM句のヘルパー関数　新規シートを末尾に追加して、データベース代わりにする
' DebugPrint:         引数に型マークを付けてイミディエイトウィンドウに書き出す



'///////// variables //////////
Private default_ As Variant       ' メソッドを使用すると内部変数 array_を変更する仕様にしております(メソッドチェーンのため)ので、計算前に戻したいとき用に保持しています
Private array_ As Variant

'///////// class property //////////

' [Let] デフォルトプロパティ
Public Property Let value(v As Variant)
Attribute value.VB_UserMemId = 0
    default_ = v
    array_ = v
End Property

' [Let] 表形式になっているテーブルデータをセットします
' このプロパティのセットはちょっと特殊。 -> .SetTable(False, False) = range("A1")　みたいな感じで
Public Property Let SetTable(空白行を無視 As Boolean, 空白列を無視 As Boolean, 左上のレンジ As Range)
    
    Dim lastRow_ As Long: lastRow_ = LastRow(左上のレンジ, 空白行を無視)
    Dim lastCol_ As Long: lastCol_ = LastCol(左上のレンジ, 空白列を無視)
        
    With 左上のレンジ.Parent
        value = .Range(.Cells(左上のレンジ.Row, 左上のレンジ.column), .Cells(lastRow_, lastCol_))
    End With
    
End Property

' [Let] 表形式になっているテーブルデータをセットします。SetTable よりお手軽だけど、使える状況に制限があります
Public Property Let CurrentRegion(テーブル内レンジ As Range)

    Dim rng As Range: Set rng = テーブル内レンジ.CurrentRegion
    value = rng.Offset(1, 0).Resize(rng.Rows.Count - 1, rng.Columns.Count)   ' ヘッダー分の行を取り除きます

End Property

' [Get] デフォルトプロパティ
Public Property Get value() As Variant
    value = array_
End Property

' [Get] 行の最小インデックス
Public Property Get MinRow(Optional 配列 As Variant = Empty) As Variant     ' 配列の大きさによって、Integer にも Long にもしたいので Variant で定義してます

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    If IsArray(arr) = False Then Debug.Print "ArrayEx.MinRow:配列をセットしてください": Exit Property
    MinRow = LBound(arr, 1)
    
End Property

' [Get] 行の最大インデックス
Public Property Get MaxRow(Optional 配列 As Variant = Empty) As Variant

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    If IsArray(arr) = False Then Debug.Print "ArrayEx.MaxRow:配列をセットしてください": Exit Property
    MaxRow = UBound(arr, 1)
    
End Property

' [Get] 列の最小インデックス
Public Property Get MinColumn(Optional 配列 As Variant = Empty) As Variant

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    If IsArray(arr) = False Then Debug.Print "ArrayEx.MinColumn:配列をセットしてください": Exit Property
    MinColumn = LBound(arr, 2)
    
End Property

' [Get] 列の最大インデックス
Public Property Get MaxColumn(Optional 配列 As Variant = Empty) As Variant

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    If IsArray(arr) = False Then Debug.Print "ArrayEx.MaxColumn:配列をセットしてください": Exit Property
    MaxColumn = UBound(arr, 2)
    
End Property

' [Get] 行または列の総数
Public Property Get Count(Optional 行列 As XlRowCol = XlRowCol.xlRows, Optional 配列 As Variant = Empty) As Variant
    
    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    If IsArray(arr) = False Then Debug.Print "ArrayEx.Count:配列をセットしてください": Exit Property
    
    If 行列 = XlRowCol.xlRows Then
        Count = MaxRow(arr) - MinRow(arr) + 1
    Else
        Count = MaxColumn(arr) - MinColumn(arr) + 1
    End If

End Property

' [Get] 内部変数が空だったらTRUE
Public Property Get IsEmpty_() As Boolean
    IsEmpty_ = (IsEmpty(array_))
End Property

' [Get] 1 次元配列だったらTRUE
Public Property Get Is1D(Optional 配列 As Variant = Empty) As Boolean
    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    Is1D = IIf(ArrayRank(arr) = 1, True, False)
End Property

' [Get] 2 次元配列だったらTRUE
Public Property Get Is2D(Optional 配列 As Variant = Empty) As Boolean
    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    Is2D = IIf(ArrayRank(arr) = 2, True, False)
End Property

' [Get] ジャグ配列だったらTRUE
Function IsJagg(Optional 配列 As Variant = Empty) As Boolean

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    If IsArray(arr) = False Then Debug.Print "ArrayEx.IsJagg:配列をセットしてください": Exit Function
    If ArrayRank(arr) > 1 Then Exit Function
    
    Dim v1 As Variant, v2 As Variant
    IsJagg = False

    For Each v1 In arr
        If Not IsObject(v1) Then
            For Each v2 In v1
                If Not IsObject(v2) Then
                    IsJagg = True: Exit Function
                End If
            Next v2
        End If
    Next v1
    
End Function

' [Get] 配列の次元数
Public Property Get ArrayRank(Optional 配列 As Variant = Empty) As Integer

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    
    On Error Resume Next
    Do While Err.Number = 0
        i = i + 1
        errTest = UBound(arr, i)
    Loop
    On Error GoTo 0
    
    ArrayRank = i - 1
End Property

' [Get] 配列内での指定値存在チェック
Public Property Get Exist(ByVal 検索値 As Variant, Optional 検索列 As Long = 0, Optional 開始行 As Long = 0, _
                                    Optional 検索方法 As XlLookAt = XlLookAt.xlWhole, Optional 配列 = Empty) As Boolean
    
    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    
    If IsArray(arr) = False Then Debug.Print "ArrayEx.Exist: 検索対象は配列で指定してください": GoTo ExitHandler
    If (ArrayRank(arr) = 2) And ((検索列 < LBound(arr)) Or (検索列 > UBound(arr))) Then Debug.Print "ArrayEx.Exist: 2次元配列の場合、有効な検索列を指定してください": GoTo ExitHandler
    
    Exist = IIf(SearchIndex(検索値, 検索列, 開始行, 検索方法, arr) = -1, False, True)
ExitHandler:
End Property

' [Get] 平均値の算出
Public Property Get Average_(Optional 列番 As Long = -1, Optional 配列 As Variant = Empty) As Variant

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    
    Select Case ArrayRank(arr)
        Case 1: Average_ = WorksheetFunction.Average(arr)
        Case 2: Average_ = WorksheetFunction.Average(Slicer(列番, xlVerticalCoordinate, arr))
                    array_ = arr
    End Select
    
End Property

' [Get] 合計値の算出
Public Property Get Sum_(Optional 列番 As Long = -1, Optional 配列 As Variant = Empty) As Variant

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    
    Select Case ArrayRank(arr)
        Case 1: Sum_ = WorksheetFunction.Sum(arr)
        Case 2: Sum_ = WorksheetFunction.Sum(Slicer(列番, xlVerticalCoordinate, arr))
                    array_ = arr
    End Select

End Property

' [Get] 中央値の算出
Public Property Get Median_(Optional 列番 As Long = -1, Optional 配列 As Variant = Empty) As Variant

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    
    Select Case ArrayRank(arr)
        Case 1: Median_ = WorksheetFunction.Median(arr)
        Case 2: Median_ = WorksheetFunction.Median(Slicer(列番, xlVerticalCoordinate, arr))
                    array_ = arr
    End Select

End Property

' [Get] 最大値の算出
Public Property Get Max_(Optional 列番 As Long = -1, Optional 配列 As Variant = Empty) As Variant

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    
    Select Case ArrayRank(arr)
        Case 1: Max_ = WorksheetFunction.Max(arr)
        Case 2: Max_ = WorksheetFunction.Max(Slicer(列番, xlVerticalCoordinate, arr))
                    array_ = arr
    End Select

End Property

' [Get] 最小値の算出
Public Property Get Min_(Optional 列番 As Long = -1, Optional 配列 As Variant = Empty) As Variant

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    
    Select Case ArrayRank(arr)
        Case 1: Min_ = WorksheetFunction.Min(arr)
        Case 2: Min_ = WorksheetFunction.Min(Slicer(列番, xlVerticalCoordinate, arr))
                    array_ = arr
    End Select

End Property

' [Get] 最頻値の算出。最頻値は複数あるかもしれないので、返り値は配列にしています。
Public Property Get Mode_(Optional 列番 As Long = -1, Optional 配列 As Variant = Empty) As Variant

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    
    Select Case ArrayRank(arr)
        Case 1: Mode_ = WorksheetFunction.Mode_Mult(arr)
        Case 2: Mode_ = WorksheetFunction.Mode_Mult(Slicer(列番, xlVerticalCoordinate, arr))
                    array_ = arr
    End Select
    
End Property

' [Get] 標準偏差の算出
Public Property Get StDev_(Optional 列番 As Long = -1, Optional 配列 As Variant = Empty) As Variant

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    
    Select Case ArrayRank(arr)
        Case 1: StDev_ = WorksheetFunction.StDevP(arr)
        Case 2: StDev_ = WorksheetFunction.StDevP(Slicer(列番, xlVerticalCoordinate, arr))
                    array_ = arr
    End Select

End Property

' [Get] CPKの算出
Public Property Get CPK_(SP As Single, 管理幅 As Single, Optional 列番 As Long = -1, Optional 配列 As Variant = Empty) As Variant

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    
    Dim ave As Single: ave = Average_(列番, arr)
    Dim std As Single: std = StDev_(列番, arr)
    
    CPK_ = WorksheetFunction.Min(((SP + 管理幅) - ave), (ave - (SP - 管理幅))) / (3 * std)

End Property

'<summary>
'  [Get]配列内の検索。指定列を順方向・行方向に一致をみる
'　返り値 値が一致した"行の"添え字。無かったら[-1]。
'
'　検索値：検索値
'　[省略可]  検索列：検索する列番号
'　[省略可]  検索開始行：検索開始行番号　デフォルト：配列の一番上から
'　[省略可]  検索方法：完全一致か部分一致　デフォルト：完全一致
'</summary>
Public Property Get SearchIndex(ByVal 検索値 As Variant, Optional 検索列 As Long = 0, _
                                    Optional 検索開始行 As Long = 0, Optional 検索方法 As XlLookAt = XlLookAt.xlWhole, Optional 配列 As Variant = Empty) As Long
    
    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    Dim ret As Variant
    
    If IsArray(arr) = False Then Debug.Print "ArrayEx.SearchIndex:クラスに配列をセットしてください": GoTo ExitHandler
    
    If 検索開始行 = 0 Then 検索開始行 = LBound(arr)
    If (検索方法 = XlLookAt.xlPart) And Not IsObject(検索値) And Not IsNull(検索値) And Not IsEmpty(検索値) Then 検索値 = "*" & 検索値 & "*"
    
    Select Case ArrayRank(arr)
        Case 1: ret = Search1D(検索値, 検索開始行, arr)
        Case 2:
            If (検索列 < LBound(arr, 2)) Or (検索列 > UBound(arr, 2)) Then Debug.Print "ArrayEx.SearchIndex: 2次元配列の場合、有効な検索列を指定してください": GoTo ExitHandler
            ret = Search2D(検索値, 検索列, 検索開始行, arr)
    End Select

ExitHandler:
    SearchIndex = ret
End Property

Private Function Search1D(ByVal 検索値 As Variant, 検索開始行 As Long, 配列 As Variant) As Long
        
    For i = 検索開始行 To UBound(配列)
        If (IsObject(配列(i)) And IsObject(検索値)) Or _
            (IsNull(配列(i)) And IsNull(検索値)) Or _
            (IsEmpty(配列(i)) And IsEmpty(検索値)) Then
            
            If TypeName(配列(i)) = TypeName(検索値) Then Search1D = i: Exit Function

        ElseIf Not IsObject(配列(i)) And Not IsObject(検索値) Then
            If (配列(i) Like 検索値) Then Search1D = i: Exit Function
        End If
    Next i

    Search1D = -1       ' 検索値が見つからなかった時
End Function

Private Function Search2D(ByVal 検索値 As Variant, 検索列 As Long, 検索開始行 As Long, 配列 As Variant) As Long
        
        For i = 検索開始行 To UBound(配列)
            
            If (IsObject(配列(i, 検索列)) And IsObject(検索値)) Or _
                (IsNull(配列(i, 検索列)) And IsNull(検索値)) Or _
                (IsEmpty(配列(i, 検索列)) And IsEmpty(検索値)) Then
                
                If TypeName(配列(i, 検索列)) = TypeName(検索値) Then Search2D = i: Exit Function
 
            ElseIf Not IsObject(配列(i, 検索列)) And Not IsObject(検索値) Then
                If (配列(i, 検索列) Like 検索値) Then Search2D = i: Exit Function
            End If
            
        Next i

    Search2D = -1       ' 検索値が見つからなかった時
End Function


'///////// module //////////

'<summary>
'  内部変数 array_を初期値に戻します。繰り返し ArrayEx を使う場合に呼び出します。
'</summary>
Function Init() As ArrayEx
    array_ = default_
    Set Init = Me
End Function

'<summary>
'  2次元配列を1次元配列にする。
'
'　[省略可]  配列：対象の配列      デフォルト：ローカル配列
'</summary>
Function To1D(Optional 配列 As Variant = Empty) As ArrayEx

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    Dim ret As Variant
    
    If Is1D(arr) = True Then Debug.Print "ArrayEx.To1D:既に1次元配列です": GoTo ExitHandler
    If IsArray(arr) = False Then Debug.Print "ArrayEx.To1D:配列をセットしてください": GoTo ExitHandler

    ReDim ret(LBound(arr) To (UBound(arr) - LBound(arr) + 1) * (UBound(arr, 2) - LBound(arr, 2) + 1) + LBound(arr) - 1)
    Dim cnt As Long: cnt = LBound(ret)
    
    For i = LBound(arr) To UBound(arr)
        For j = LBound(arr, 2) To UBound(arr, 2)
            Call Assign(ret(cnt), arr(i, j))
            cnt = cnt + 1
        Next j
    Next i
    
ExitHandler:
    array_ = ret: Set To1D = Me
    
End Function

'<summary>
'  1次元配列を2次元配列にする。
'  1次元配列と2次元配列が混合する場合でも一様の処理を行いたいときとかに使う
'
'　[省略可]  行列方向：展開する方向。xlRows : (0,0),(1,0),(2,0)... 　xlColumns：(0,0),(0,1),(0,2)...　 デフォルト：xlRows
'　[省略可]  配列：対象の配列      デフォルト：ローカル配列
'</summary>
Function To2D(Optional 行列方向 As XlRowCol = XlRowCol.xlRows, Optional 配列 As Variant = Empty) As ArrayEx
Attribute To2D.VB_Description = "1次元配列を2次元配列に変換する"
    
    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
     Dim ret As Variant
     
    If Is2D(arr) = True Then Debug.Print "ArrayEx.To2D:既に2次元配列です": GoTo ExitHandler
    If IsArray(arr) = False Then Debug.Print "ArrayEx.To2D:配列をセットしてください": GoTo ExitHandler
    
    If 行列方向 = xlRows Then
        ReDim ret(LBound(arr) To UBound(arr), LBound(arr) To LBound(arr)) As Variant
        For i = LBound(arr) To UBound(arr)
            Assign ret(i, LBound(arr)), arr(i)
        Next i
    Else
        ReDim ret(LBound(arr) To LBound(arr), LBound(arr) To UBound(arr)) As Variant
        For i = LBound(arr) To UBound(arr)
            Assign ret(LBound(arr), i), arr(i - LBound(arr))
        Next i
    End If

ExitHandler:
    array_ = ret: Set To2D = Me
    
End Function

'<summary>
'  配列をコレクションに変換する
'
'　[省略可]  配列：対象の配列      デフォルト：ローカル配列
'</summary>
Function ToCollection(Optional 配列 As Variant = Empty) As Collection

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    If IsArray(arr) = False Then Debug.Print "ArrayEx.ToCollection:クラスに配列をセットしてください": Exit Function
    
    Select Case ArrayRank(arr)
        Case 1: Set ToCollection = ToCollection1D(arr)
        Case 2: Set ToCollection = ToCollection2D(arr)
    End Select
    
End Function

Private Function ToCollection1D(arr) As Collection
    
    Dim c As New Collection
    For i = LBound(arr) To UBound(arr)
        c.Add (arr(i))
    Next i
    
    Set ToCollection1D = c
End Function

Private Function ToCollection2D(arr) As Collection
    
    Dim c As New Collection: Dim i As Long
    For i = LBound(arr) To UBound(arr)
        c.Add (Slicer2D(i, xlHorizontalCoordinate, arr))
    Next i
    
    Set ToCollection2D = c
End Function

'<summary>
'  最初のインデックスを引数の indexに強制する
'
' index : 強制したい最初のインデックス値
'  [省略可] 配列：対象の配列      デフォルト：ローカル配列
'</summary>
Function FirstIndexTo(index As Long, Optional 配列 As Variant = Empty) As ArrayEx

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    Dim ret As Variant
    
    If IsArray(arr) = False Then Debug.Print "ArrayEx.FirstIndexTo:クラスに配列をセットしてください": GoTo ExitHandler
    
    Select Case ArrayRank(arr)
        Case 1: ret = FirstIndexTo1D(index, arr)
        Case 2: ret = FirstIndexTo2D(index, arr)
    End Select

ExitHandler:
    array_ = ret: Set FirstIndexTo = Me
    
End Function

Private Function FirstIndexTo1D(index As Long, arr As Variant) As Variant
    
    If LBound(arr, 1) = index Then
        FirstIndexTo1D = arr
    
    Else
        Dim ret As Variant: ReDim ret(index To index + UBound(arr) - LBound(arr))
        For i = index To UBound(ret, 1)
            Call Assign(ret(i), arr(i + LBound(arr) - LBound(ret)))
        Next i
        
        FirstIndexTo1D = ret
    End If

End Function

Private Function FirstIndexTo2D(index As Long, arr As Variant) As Variant

    If (LBound(arr, 1) = index) And (LBound(arr, 2) = index) Then
        FirstIndexTo2D = arr
    
    Else
        Dim ret As Variant: ReDim ret(index To index + UBound(arr) - LBound(arr), index To index + UBound(arr, 2) - LBound(arr, 2))
        For i = index To UBound(ret, 1)
            For j = index To UBound(ret, 2)
                Call Assign(ret(i, j), arr(i + LBound(arr) - LBound(ret), j + LBound(arr, 2) - LBound(ret, 2)))
            Next j
        Next i
        
        FirstIndexTo2D = ret
    End If

End Function

'<summary>
'  配列内から指定値の要素ごと削除(配列が圧縮される可能性がある)。Delete/ClearではなくRemove
'
'　削除する値：削除対象になるKey
'　[省略可]  列番号：２次元配列の場合の対象列
'</summary>
Function Remove(削除する値 As Variant, Optional 列番号 As Long = -1, Optional 配列 As Variant = Empty) As ArrayEx

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    Dim ret As Variant
    
    If IsArray(arr) = False Then Debug.Print "ArrayEx.Remove:クラスに配列をセットしてください": GoTo ExitHandler
    
    Select Case ArrayRank(arr)
        Case 1: ret = Remove1D(削除する値, arr)
        Case 2:
            If (列番号 < LBound(arr, 2)) Or (列番号 > UBound(arr, 2)) Then Debug.Print "ArrayEx.Remove: 2次元配列の場合、有効な検索列を指定してください": GoTo ExitHandler
            ret = Remove2D(削除する値, 列番号, arr)
    End Select

ExitHandler:
    array_ = ret: Set Remove = Me
    
End Function

Private Function Remove1D(削除する値 As Variant, 配列 As Variant) As Variant

    Dim ret As Variant: ReDim ret(LBound(配列) To LBound(配列)) As Variant
    
    For i = LBound(配列) To UBound(配列)
        If 配列(i) <> 削除する値 Then
            Call Assign(ret(UBound(ret)), 配列(i))
            ReDim Preserve ret(LBound(ret) To UBound(ret) + 1)
        End If
    Next i
    
    If (LBound(ret) = UBound(ret)) And (ret(LBound(ret)) = Empty) Then
        Remove1D = Empty
    Else
        ReDim Preserve ret(LBound(ret) To UBound(ret) - 1)
        Remove1D = ret
    End If
    
End Function

Private Function Remove2D(削除する値 As Variant, 列番号 As Long, 配列 As Variant) As Variant

    Dim ret As Variant
    Dim firstkiss As Boolean: firstkiss = False
    
    For i = LBound(配列) To UBound(配列)
        
        If 配列(i, 列番号) <> 削除する値 Then
            
            If firstkiss = False Then
                ReDim ret(LBound(配列) To LBound(配列), LBound(配列, 2) To UBound(配列, 2)) As Variant      ' 初回のみ
            Else
                ret = Append(xlRows, ret)      ' 2次元の行追加には ReDim Preserve が使えません
            End If
            
            firstkiss = True
            For j = LBound(配列, 2) To UBound(配列, 2)
                Call Assign(ret(UBound(ret), j), 配列(i, j))
            Next j
            
        End If
    Next i
    
    Remove2D = IIf(firstkiss = False, Empty, ret)
    
End Function

'<summary>
'　配列内のNull/Empty を引数値に置換。
'
'　置換後の値：置換後の値
'　[省略可] 配列：対象の配列
'</summary>
Function ReplaceNullTo(ByVal 置換後の値 As Variant, Optional 配列 As Variant = Empty) As ArrayEx
    
    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    Dim ret As Variant
    
    If IsArray(arr) = False Then
        If IsNull(arr) Or IsEmpty(arr) Then Call Assign(ret, 置換後の値)
    Else
        Select Case ArrayRank(arr)
            Case 1: ret = ReplaceNullTo1D(置換後の値, arr)
            Case 2: ret = ReplaceNullTo2D(置換後の値, arr)
        End Select
    End If
   
    array_ = ret: Set ReplaceNullTo = Me

End Function

Function ReplaceNullTo1D(ByVal 置換後の値 As Variant, arr As Variant) As Variant
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        If IsNull(arr(i)) Or IsEmpty(arr(i)) Then Call Assign(arr(i), 置換後の値)
    Next i
    ReplaceNullTo1D = arr

End Function

Function ReplaceNullTo2D(ByVal 置換後の値 As Variant, Optional arr As Variant = Empty) As Variant

    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            If IsNull(arr(i, j)) Or IsEmpty(arr(i, j)) Then Call Assign(arr(i, j), 置換後の値)
        Next j
    Next i
    ReplaceNullTo2D = arr

End Function

'<summary>
'  配列のフィルター結果を返す。
'
'　検索値：検索値
'　検索行：検索する列番号
'　[省略可]  検索方法：完全一致か部分一致　デフォルト：完全一致
'</summary>
Function Filter_(検索値 As Variant, Optional 検索列 As Long = 0, _
                             Optional 検索方法 As XlLookAt = XlLookAt.xlWhole, Optional 配列 = Empty) As ArrayEx

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    Dim ret As Variant
    
    If IsArray(arr) = False Then Debug.Print "ArrayEx.Filter_: クラスに配列をセットしてください": GoTo ExitHandler
    If (検索方法 = XlLookAt.xlPart) And Not IsObject(検索値) And Not IsNull(検索値) And Not IsEmpty(検索値) Then 検索値 = "*" & 検索値 & "*"
    
    Select Case ArrayRank(arr)
        Case 1: ret = FilterOn1D(検索値, arr)
        Case 2:
            If (検索列 < LBound(arr, 2)) Or (検索列 > UBound(arr, 2)) Then Debug.Print "ArrayEx.Filter_: 2次元配列の場合、有効な検索列を指定してください": GoTo ExitHandler
            ret = FilterOn2D(検索値, 検索列, arr)
    End Select

ExitHandler:
    array_ = ret: Set Filter_ = Me
End Function

Private Function FilterOn1D(検索値 As Variant, arr As Variant) As Variant

    Dim ret As Variant: ReDim ret(LBound(arr) To UBound(arr)) As Variant
    Dim rowCount As Long: rowCount = LBound(arr)        ' 返す配列の行
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)
        
        i = Search1D(検索値, i, arr)
        If i = -1 Then Exit For
        
        Assign ret(rowCount), array_(i)
        rowCount = rowCount + 1
        
    Next i
    
'    If IsEmpty(ret(LBound(ret))) Then Exit Function
    FilterOn1D = Remove1D(Empty, ret)
    
End Function

Private Function FilterOn2D(ByVal 検索値 As Variant, 検索列 As Long, arr As Variant) As Variant

    Dim ret As Variant: ReDim ret(LBound(arr) To UBound(arr), LBound(arr, 2) To UBound(arr, 2)) As Variant
    Dim rowCount As Long: rowCount = LBound(arr)        ' 代入先の行
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)
    
        i = Search2D(検索値, 検索列, i, arr)
        If i = -1 Then Exit For
        
        For j = LBound(arr, 2) To UBound(arr, 2)
            Assign ret(rowCount, j), arr(i, j)
        Next j
        rowCount = rowCount + 1
        
    Next i
    
'    If IsEmpty(ret(LBound(ret))) Then Exit Function
    FilterOn2D = Remove2D(Empty, 検索列, ret)
    
End Function

'<summary>
'  配列拡張
'
'  [省略可]  配列：拡張する配列
'  [省略可]  行列選択：行か列か　デフォルト：行
'</summary>
Function Append(Optional 行列選択 As XlRowCol = XlRowCol.xlRows, Optional 配列 As Variant = Empty) As ArrayEx

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    Dim ret As Variant
    
    If IsArray(arr) = False Then Debug.Print "ArrayEx.Append: 対象の配列を指定してください": GoTo ExitHandler

    Select Case ArrayRank(arr)
        Case 1:
            ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1) + 1)
            ret = arr
        Case 2:
            If 行列選択 = xlColumns Then
                ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2) + 1)       ' Preserveは最下位の次元のみ使える
                ret = arr
            Else
                ReDim ret(LBound(arr, 1) To UBound(arr, 1) + 1, LBound(arr, 2) To UBound(arr, 2))                         ' 2次元配列の行は地道に
                Call AssignArray(ret, arr)
            End If
    End Select
    
ExitHandler:
    array_ = ret: Set Append = Me
End Function

'<summary>
'  配列から指定の行/列を取り出す。返り値 1次元配列
'
'　行列番号：行または列の番号・添え字
'　[省略可]  行列選択：行か列か　デフォルト：列 ※表でみると縦1列
'  [省略可]  配列：対象の配列
'</summary>
Function Slicer(行列番号 As Long, Optional 行列選択 As XlPieSliceLocation = XlPieSliceLocation.xlVerticalCoordinate, Optional 配列 As Variant = Empty) As ArrayEx

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    Dim ret As Variant
    
    If IsArray(arr) = False Then Debug.Print "ArrayEx.Slicer: 対象の配列を指定してください": GoTo ExitHandler
    
    Select Case ArrayRank(arr)
        Case 1: ret = Slicer1D(行列番号, 行列選択, arr)
        Case 2:
            If 行列選択 = XlPieSliceLocation.xlVerticalCoordinate Then
                If (行列番号 < LBound(arr, 2)) Or (行列番号 > UBound(arr, 2)) Then Debug.Print "ArrayEx.Slicer: 2次元配列の場合、有効な検索列を指定してください": GoTo ExitHandler
            ElseIf 行列選択 = XlPieSliceLocation.xlHorizontalCoordinate Then
                If (行列番号 < LBound(arr, 1)) Or (行列番号 > UBound(arr, 1)) Then Debug.Print "ArrayEx.Slicer: 2次元配列の場合、有効な検索行を指定してください": GoTo ExitHandler
            End If
        
            ret = Slicer2D(行列番号, 行列選択, arr)
    End Select

ExitHandler:
    array_ = ret: Set Slicer = Me
End Function

Private Function Slicer1D(行番号 As Long, 行列選択 As XlPieSliceLocation, arr As Variant) As Variant
    
    If 行列選択 = XlPieSliceLocation.xlVerticalCoordinate Then
        Slicer1D = arr
    ElseIf 行列選択 = XlPieSliceLocation.xlHorizontalCoordinate Then
        Slicer1D = arr(IIf(行番号 < LBound(arr), LBound(arr), IIf(行番号 > UBound(arr), UBound(arr), 行番号)))
    End If
    
End Function

Private Function Slicer2D(行列番号 As Long, 行列選択 As XlPieSliceLocation, arr As Variant) As Variant

    Dim ret As Variant
    If ArrayRank(arr) = 1 Then arr = To2D(xlRows, arr)
    
    If 行列選択 = XlPieSliceLocation.xlVerticalCoordinate Then
        ReDim ret(LBound(arr, 1) To UBound(arr, 1))
        For i = LBound(arr, 1) To UBound(arr, 1)
            Assign ret(i), arr(i, 行列番号)
        Next i
        
    ElseIf 行列選択 = XlPieSliceLocation.xlHorizontalCoordinate Then
        ReDim ret(LBound(arr, 2) To UBound(arr, 2))
        For i = LBound(arr, 2) To UBound(arr, 2)
            Assign ret(i), arr(行列番号, i)
        Next i
    
    End If
    
    Slicer2D = ret

End Function

'<summary>
'　指定された列の取り出し
'
'　列番号：取り出す列番号　Array()関数で指定
'　[省略可] 配列：対象の配列
'</summary>
Function Select_(列番号 As Variant, Optional 配列 As Variant = Empty) As ArrayEx

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    Dim ret As Variant
    
    If IsArray(arr) = False Then Debug.Print "ArrayEx.Select_: 対象の配列を指定してください": GoTo ExitHandler
    If (IsArray(列番号) = False) And (IsNumeric(列番号) = False) Then Debug.Print "ArrayEx.Select_: 列番号を配列もしくは、数値で指定してください": GoTo ExitHandler
    
    If IsArray(列番号) = False Then
        ret = Slicer(CLng(列番号), xlVerticalCoordinate, arr)
    Else
        For i = LBound(列番号) To UBound(列番号)
            If (列番号(i) > UBound(arr, 2)) Or (列番号(i) < LBound(arr, 2)) Then Debug.Print "ArrayEx.Select_: インデックスが範囲外です": Exit For
            ret = Union(xlColumns, Slicer(CLng(列番号(i)), xlVerticalCoordinate, arr), ret)
        Next i
    End If
    
ExitHandler:
    array_ = FirstIndexTo(LBound(arr), ret): Set Select_ = Me
End Function

'<summary>
'　配列の結合
'
'　行列方向：結合する方向。Rowなら下方向に、Colなら右方向に追加したものを返す
'  結合する配列：結合する配列
'　[省略可] 結合される配列：この配列を基準に結合する
'</summary>
Function Union(行列方向 As XlRowCol, 結合する配列 As Variant, Optional 結合される配列 As Variant = Null) As ArrayEx

    Dim arr As Variant: arr = IIf(IsNull(結合される配列), me_, 結合される配列)
    Dim ret As Variant
    
    If IsEmpty(arr) Then
        ret = 結合する配列
        
    Else
        If (IsArray(arr) = False) Or (IsArray(結合する配列) = False) Then Debug.Print "ArrayEx.Union: 対象の配列を指定してください": GoTo ExitHandler
    
        ' arr と結合する配列がそれぞれ１次元か２次元かで処理を分けて記述するのが面倒すぎるので、どっちも一旦、２次元にします
        If Is1D(arr) Then arr = To2D(xlRows, arr)
        If Is1D(結合する配列) Then 結合する配列 = To2D(xlRows, 結合する配列)
        
        ret = Union2D(行列方向, FirstIndexTo2D(LBound(arr), 結合する配列), FirstIndexTo2D(LBound(arr), arr))
        If LBound(ret, 2) = UBound(ret, 2) Then ret = To1D(ret)
    End If

ExitHandler:
    array_ = ret: Set Union = Me
End Function

Private Function Union2D(行列方向 As XlRowCol, 結合する配列 As Variant, arr As Variant) As Variant

    Dim ret As Variant
    
    If 行列方向 = XlRowCol.xlRows Then
        
        ReDim ret(LBound(arr) To LBound(arr) + (UBound(arr) - LBound(arr) + 1) + (UBound(結合する配列) - LBound(結合する配列)), _
                         LBound(arr, 2) To IIf(UBound(arr, 2) > UBound(結合する配列, 2), UBound(arr, 2), UBound(結合する配列, 2))) As Variant
        Call AssignArray(ret, arr)

        For i = LBound(ret, 1) To UBound(ret, 1)
            If i <= UBound(結合する配列) Then
                For j = LBound(ret, 2) To UBound(ret, 2)
                    If j <= UBound(結合する配列, 2) Then Call Assign(ret(UBound(arr) - LBound(arr) + 1 + i, j), 結合する配列(i, j))
                Next j
            End If
        Next i

    Else
        
        ReDim ret(LBound(arr) To IIf(UBound(arr, 1) > UBound(結合する配列, 1), UBound(arr, 1), UBound(結合する配列, 1)), _
                         LBound(arr, 2) To LBound(arr, 2) + (UBound(arr, 2) - LBound(arr, 2) + 1) + (UBound(結合する配列, 2) - LBound(結合する配列, 2))) As Variant
        Call AssignArray(ret, arr)

        For i = LBound(ret, 2) To UBound(ret, 2)
            If i <= UBound(結合する配列, 2) Then
                For j = LBound(ret, 1) To UBound(ret, 1)
                    If j <= UBound(結合する配列) Then Call Assign(ret(j, UBound(arr, 2) - LBound(arr, 2) + 1 + i), 結合する配列(j, i))
                Next j
            End If
        Next i

    End If

    Union2D = ret
    
End Function

'<summary>
'　配列内の行データ重複削除。2次元の場合、行データがすべて一致している時のみ削除します
'
'　[省略可] 配列：重複削除対象の配列
'</summary>
Function GroupBy(Optional ByVal 配列 As Variant = Empty) As ArrayEx

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    Dim ret As Variant
    
    If IsArray(arr) = False Then Debug.Print "ArrayEx.GroupBy: 対象の配列を指定してください": GoTo ExitHandler
    
    Select Case ArrayRank(arr)
        Case 1: ret = GroupBy1D(arr)
        Case 2: ret = GroupBy2D(arr)
    End Select

ExitHandler:
    array_ = ret: Set GroupBy = Me
End Function

Private Function GroupBy1D(ByVal arr As Variant) As Variant

    Dim dic As Object
    Set dic = CreateObject("scripting.dictionary")
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        If dic.Exists(arr(i)) = False Then
            dic.Add arr(i), arr(i)
        End If
    Next i

    GroupBy1D = dic.items
    
End Function

Private Function GroupBy2D(ByVal arr As Variant) As Variant

    Dim dicObj As Object
    Set dicObj = CreateObject("scripting.dictionary")
    
    Dim key As Variant
    
    For i = LBound(arr, 1) To UBound(arr, 1)
    
        key = Join(Slicer(CLng(i), xlHorizontalCoordinate, arr), ",")
        If dicObj.Exists(key) = False Then
            dicObj.Add key, Slicer(CLng(i), xlHorizontalCoordinate, arr)
        End If
    
    Next i

    GroupBy2D = JaggTo2D(dicObj.items)
    
End Function

Private Function JaggTo2D(arr As Variant) As Variant

    Dim arr2D As Variant: ReDim arr2D(LBound(arr, 1) To UBound(arr, 1), LBound(arr(0), 1) To UBound(arr(0), 1))
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr(0), 1) To UBound(arr(0), 1)
            arr2D(i, j) = arr(i)(j)
        Next j
    Next i
    JaggTo2D = arr2D

End Function

'<summary>
'  Excelシートに配列データを貼り付ける

'  baseRange : 貼り付ける基準(左上)のレンジ
'　[省略可] 配列：対象の配列
'</summary>
Sub ToSheet(baseRange As Range, Optional 配列 As Variant = Empty)

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    
    If IsArray(arr) = False Then
        baseRange = arr
    Else
        Dim ws As Worksheet: Set ws = Worksheets(baseRange.Parent.name)
        Dim wb As Workbook: Set wb = Workbooks(ws.Parent.name)
        Dim baseRow As Long: baseRow = baseRange.Row
        Dim baseColumn As Long: baseColumn = baseRange.column
        
        With wb.Worksheets(ws.name): .Activate
            Select Case ArrayRank(arr)
                Case 1: .Range(Cells(baseRow, baseColumn), Cells(baseRow + UBound(arr, 1) - LBound(arr, 1), baseColumn)) = WorksheetFunction.Transpose(ReplaceNullTo("", arr))
                Case 2: .Range(Cells(baseRow, baseColumn), Cells(baseRow + UBound(arr, 1) - LBound(arr, 1), baseColumn + UBound(arr, 2) - LBound(arr, 2))) = arr
            End Select
        End With
    End If
    
End Sub

'<summary>
'  SQLの実行

'  sql : SQL構文
'  [省略可] Book : ソースになるワークブック

' SQL例1 フィルター：SELECT F1,F3 FROM [Sheet1$B3:E20] WHERE F1 = 'Man' and (F2 >= 20 and F2 < 30) ORDER BY F4, F3 DESC
' SQL例2 リレー：SELECT * FROM [Sheet2$],[Sheet3$] WHERE [Sheet2$].F2 = [Sheet3$].F1 ORDER BY [Sheet2$].F1
'</summary>
Function Execute(sql As String, Optional Book As Workbook = Nothing) As ArrayEx

    On Error GoTo ErrorHandler
    
    Dim ret As Variant
    If Book Is Nothing Then Set Book = ThisWorkbook
   
    ' 接続試行
    Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
    Call OpenSource(cn, Book)
    
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    rs.Open Source:=sql, ActiveConnection:=cn
    
    ret = rs.GetRows()
    
    GoTo Finally
    
ErrorHandler:
    If rs.EOF Then GoTo Finally
    Debug.Print "ArrayEx.Execute: " & Err.Description
    
Finally:
    Call CloseSource(cn)
    Set cn = Nothing: Set rs = Nothing
    array_ = ret: Set Execute = Me
End Function

'<summary>
'  private] Excelデータソースと接続してみる

'  cn : ADODB.Connection オブジェクト
'  wb : ソースになるワークブック
'</summary>
Private Sub OpenSource(cn As Object, wb As Workbook)

    Const PROVIDER    As String = "PROVIDER=MSDASQL;"
    Const DRIVER        As String = "DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};"
    Const path            As String = "DBQ="
    
    cn.ConnectionString = PROVIDER & DRIVER & path & wb.path & "\" & wb.name & ";"
    cn.Open
End Sub

'<summary>
'  Private]  ソースとの接続を閉じる。

'  cn : ADODB.Connection オブジェクト
'</summary>
Private Sub CloseSource(cn As Object)

    ' 0 = adStateClosed;
    ' 2 = adStateConnecting;
    ' 4 = adStateExecuting;
    ' 8 = adStateFetching;
    If (cn.State = 0) Or _
        ((cn.State And 2) = 2) Or _
        ((cn.State And 4) = 4) Or _
        ((cn.State And 8) = 8) Then Exit Sub
    
    cn.Close

End Sub

'<summary>
'  SQL.FROM句のヘルパー関数。　新規シートに配列を貼り付けて、範囲文字列を返す。

'  [省略可] header1D : ヘッダー(1次元配列)　デフォルト：F1,F2,F3・・・
'  [省略可] 配列 : ソースになる配列
'</summary>
Function FROM(Optional header1D As Variant = Empty, Optional 配列 As Variant = Empty) As String
    
    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    If IsArray(arr) = False Then Debug.Print "ArrayEx.FROM: 対象の配列を指定してください": Exit Function
    
    ' 新規シートの追加
'    With ThisWorkbook
'        Set Source = .Worksheets.Add(after:=.Worksheets(.Worksheets.Count))
'    End With
    
    With ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    
        ' ヘッダーを追記
        If IsEmpty(header1D) Then
            For i = LBound(arr, 2) To UBound(arr, 2)
                .Cells(1, i) = "F" & CStr(i)
            Next i
        ElseIf IsArray(header1D) And (ArrayRank(header1D) = 1) Then
            For i = 0 To UBound(header1D, 1) - LBound(header1D, 1)
                .Cells(1, i + 1) = header1D(LBound(header1D, 1) + i)
            Next i
        End If
        
        ' 配列を貼り付けて、範囲文字列を返す
        Call ToSheet(.Range("A2"), arr)
        FROM = " FROM [" & .name & "$] "
        Select Case ArrayRank(arr)
            Case 1:
                FROM = " FROM [" & .name & "$" & .Cells(1, 1).Address(False, False) & ":" & _
                            .Cells(UBound(arr) - LBound(arr) + 2).Address(False, False) & "] "

            Case 2:
                FROM = " FROM [" & .name & "$" & .Cells(1, 1).Address(False, False) & ":" & _
                            .Cells(UBound(arr, 1) - LBound(arr, 1) + 2, UBound(arr, 2) - LBound(arr, 2) + 1).Address(False, False) & "] "
        End Select
        
    End With

End Function


'<summary>
'  引数に型マークを付けてイミディエイトウィンドウに書き出す
'
'  [省略可] x：マークをつける変数。デフォルト：このクラスの内部変数arrを書き出す
'</summary>
Sub DebugPrint(Optional ByVal 配列 As Variant = Empty)

    Dim arr As Variant: arr = IIf(IsEmpty(配列), array_, 配列)
    Debug.Print (Dump(arr))

End Sub

Function Dump(ByVal x As Variant) As String

    Dim ty As String: ty = TypeName(x)
    
    If IsObject(x) Then Dump = ty: GoTo Escape
    Select Case ty
    Case "Boolean":     Dump = x
    Case "Integer":     Dump = x & "%"
    Case "Long":        Dump = x & "&"
    #If VBA7 And Win64 Then
    Case "LongLong":    Dump = x & "^"
    #End If
    Case "Single":      Dump = x & "!"
    Case "Double":      Dump = x & "#"
    Case "Currency":    Dump = x & "@"
    Case "Byte":        Dump = "CByte(" & x & ")"
    Case "Decimal":     Dump = "CDec(" & x & ")"
    Case "Date":
        Dim d As String, t As String
        If Abs(x) >= 1 Then d = Month(x) & "/" & Day(x) & "/" & Year(x)
        If Not IsInt(x) Then t = format(x, "h:nn:ss AM/PM")
        Dump = "#" & Trim(d & " " & t) & "#"
    Case "String"
        If StrPtr(x) = 0 Then
            Dump = "(vbNullString)"
        Else
            Dump = """" & x & """"
        End If
    Case "Empty", "Null", "Nothing"
        Dump = "(" & ty & ")"
    Case "Error"
        If IsMissing(x) Then
            Dump = "(Missing)"
        Else
            Dump = "CVErr(" & ReMatch(CStr(x), "\d+")(0) & ")"
        End If
    Case "ErrObject"
        Dump = "Err " & x.Number
    Case "Unknown"
        Dump = ty
    Case Else
        If Not IsArray(x) Then Dump = "": GoTo Escape
        
        Dim lb As Long: lb = LBound(x)
        Dim ub As Long: ub = UBound(x)
        Dim ar As Variant
        Dim i As Long
        
        If ArrayRank(x) = 1 Then         ' 1次元配列
            ReDim ar(lb To ub)
            For i = lb To ub: ar(i) = Dump(x(i)): Next
            Dump = "Array1D(" & vbNewLine & vbTab & Join(ar, ", " & vbNewLine & vbTab) & vbNewLine & ")"
        
        ElseIf ArrayRank(x) = 2 Then   ' 2次元配列
            Dim j As Long
            Dim jStr As String
            Dim lb2 As Long: lb2 = LBound(x, 2)
            Dim ub2 As Long: ub2 = UBound(x, 2)
            ReDim ar(lb2 To ub2)
            
            For i = lb To ub
                For j = lb2 To ub2: ar(j) = Dump(x(i, j)): Next
                jStr = jStr & vbNewLine & vbTab & Join(ar, ", " & vbTab)
                If i < ub Then jStr = jStr & ","
            Next i
            Dump = "Array2D(" & jStr & vbNewLine & ")"
            
        End If
    End Select
    
Escape:
End Function

Private Function IsInt(ByVal num As Variant) As Boolean
    If IsDate(num) Then num = CDbl(num)
    If Not IsNumeric(num) Then Err.Raise 13
    
    IsInt = num = Fix(num)
End Function

Private Function ReMatch( _
    ByVal expr As String, ByVal ptrnFind As String, _
    Optional ByVal iCase As Boolean = False _
    ) As Variant
    
    Dim ret As Variant: ret = Array()
    
    Dim regex As Object:  Set regex = CreateRegExp(ptrnFind, IIf(iCase, "i", ""))
    Dim ms As Object:     Set ms = regex.Execute(expr)
    If ms.Count < 1 Then: GoTo Ending
    
    Dim sms As Object:    Set sms = ms(0).SubMatches
    ReDim ret(sms.Count)
    
    ret(0) = ms.Item(0).value
    Dim i As Integer
    For i = 1 To UBound(ret): ret(i) = sms.Item(i - 1): Next
    
Ending:
    ReMatch = ret
End Function

'<summary>[private]
'  baseRangeを基準とした最下位行を返す
'
'　baseRange：基準となるRange
'　isIgnoreRowBlank：空白行を無視するか,　True：空白行お構いなしで最終データ行を返す
'</summary>
Private Function LastRow(ByVal baseRange As Range, isIgnoreRowBlank As Boolean) As Long

    With baseRange.Parent
        LastRow = IIf(isIgnoreRowBlank = True, _
                                .Cells(.Rows.Count, baseRange.column).End(xlUp).Row, _
                                baseRange.End(xlDown).Row)
    End With
    
    If LastRow < baseRange.Row Then LastRow = baseRange.Row

End Function

'<summary>[private]
'  baseRangeを基準とした最右列を返す
'
'　baseRange：基準となるRange
'　isIgnoreColumnBlank：空白列を無視するか,　True：空白行お構いなしで最終データ行を返す
'</summary>
Private Function LastCol(ByVal baseRange As Range, isIgnoreColumnBlank As Boolean) As Long

    With baseRange.Parent
        LastCol = IIf(isIgnoreColumnBlank = True, _
                            .Cells(baseRange.Row, .Columns.Count).End(xlToLeft).column, _
                            baseRange.End(xlToRight).column)
    End With
    
    If LastCol < baseRange.column Then LastCol = baseRange.column

End Function

'<summary>[private]
'  ただの代入関数。なんか、しょっちゅう使うので
'
'　代入先：代入される側
'　代入元：まんま代入元
'</summary>
Private Sub Assign(ByRef 代入先 As Variant, ByRef 代入元 As Variant)

    If IsObject(代入元) Then
        Set 代入先 = 代入元
    Else
        代入先 = 代入元
    End If

End Sub

'<summary>[private]
'  ただの代入関数。配列版
'
'　代入先：代入される側
'　代入元：まんま代入元
'</summary>
Private Sub AssignArray(ByRef 代入先 As Variant, ByRef 代入元 As Variant)

    If ArrayRank(代入元) = 1 Then
        For i = LBound(代入元) To UBound(代入元)
            Assign 代入先(i), 代入元(i)
        Next i
        
    ElseIf ArrayRank(代入元) = 2 Then
        For i = LBound(代入元) To UBound(代入元)
            For j = LBound(代入元, 2) To UBound(代入元, 2)
                Assign 代入先(i, j), 代入元(i, j)
            Next j
        Next i
        
    End If

End Sub













