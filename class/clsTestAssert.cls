VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestAssert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 最終更新日：2016/02/07  ＠しら


Private Const TestClassSuffix As String = "Test"
Private Const TestProcSuffix As String = "_Test"

Private Const AssertModule As String = "Assert"
Private Const GeneratedProc As String = "TestRunner"
Private Const CommentLineInGeneratedProc As Long = 1

Private Const ResultLineLen As Long = 76

Private xStartTime As Single
Private xEndTime As Single
Private xSuccSubCount As Long
Private xFailSubCount As Long

Private xAssertIx As Long
Private xFailMsgs As Collection

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private tm_startTime As Long

'///////// class property //////////

Private Property Get VBProject() As Object
    Dim app As Object: Set app = Application
    Select Case app.name
        Case "Microsoft Word":   Set VBProject = app.MacroContainer.VBProject
        Case "Microsoft Excel":  Set VBProject = app.ThisWorkbook.VBProject
        Case "Microsoft Access": Set VBProject = app.VBE.ActiveVBProject
        Case Else: Err.Raise 17
    End Select
End Property

Property Get StartTimeValue() As String
    StartTimeValue = tm_startTime
End Property

'///////// module list //////////

' TimerStart:           ミリ秒単位での現在時刻の記録
' TimerStop:           返り値: 経過時間
' TimeStamp:         イミディエイトウィンドウに時間を記録
' Wait:                    指定された時間処理を停止する｡誤差15msぐらい
' Dump:                  引数に型マークを付けて返す｡Debug.Print と組み合わせると便利かも
' RandomBetween:指定数値内(UpperBound～LowerBound)でランダムに数値を返す
' RandomArray:      ランダム値を格納した配列を返す


'///////// module //////////

'<summary>
'  ミリ秒単位での現在時刻の記録
'</summary>
Sub TimerStart()
  tm_startTime = GetTickCount()
End Sub

'<summary>
'  返り値：経過時間
'</summary>
Function TimerStop() As Long
  TimerStop = GetTickCount() - tm_startTime
End Function

'<summary>
'  イミディエイトウィンドウに時間を記録
'
'  [省略可] isReset：tm_startTimeをリセットするか
'</summary>
Sub TimeStamp(Optional title As String = "", Optional isReset As Boolean = False)
    Debug.Print title & TimerStop & " ms"
    If isReset Then TimerStart
End Sub

'<summary>
'  指定された時間処理を停止する。誤差15msぐらい
'
'　waitTime_ms：待ち時間。単位[msec]。15ms以上を設定した方が安牌
'</summary>
Sub Wait(waitTime_ms As Long)
    Sleep waitTime_ms
End Sub

'<summary>
'  引数に型マークを付けて返す。Debug.Print と組み合わせると便利かも
'
'  x：マークをつける変数
'</summary>
Function Dump(ByVal X As Variant) As String
    If IsObject(X) Then
        Dump = ToStr(X)
        GoTo Escape
    End If
    
    Dim ty As String: ty = TypeName(X)
    Select Case ty
    Case "Boolean":     Dump = X
    Case "Integer":     Dump = X & "%"
    Case "Long":        Dump = X & "&"
    #If VBA7 And Win64 Then
    Case "LongLong":    Dump = X & "^"
    #End If
    Case "Single":      Dump = X & "!"
    Case "Double":      Dump = X & "#"
    Case "Currency":    Dump = X & "@"
    Case "Byte":        Dump = "CByte(" & X & ")"
    Case "Decimal":     Dump = "CDec(" & X & ")"
    Case "Date":
        Dim d As String, t As String
        If Abs(X) >= 1 Then d = Month(X) & "/" & Day(X) & "/" & Year(X)
        If Not IsInt(X) Then t = format(X, "h:nn:ss AM/PM")
        Dump = "#" & Trim(d & " " & t) & "#"
    Case "String"
        If StrPtr(X) = 0 Then
            Dump = "(vbNullString)"
        Else
            Dump = """" & X & """"
        End If
    Case "Empty", "Null", "Nothing"
        Dump = "(" & ty & ")"
    Case "Error"
        If IsMissing(X) Then
            Dump = "(Missing)"
        Else
            Dump = "CVErr(" & ReMatch(CStr(X), "\d+")(0) & ")"
        End If
    Case "ErrObject"
        Dump = "Err " & X.Number
    Case "Unknown"
        Dump = ty
    Case Else
        If Not IsArray(X) Then
            Dump = ""
            GoTo Escape
        End If
        
        Dim rnk As Integer: rnk = ArrRank(X)
        If rnk = 1 Then
            Dim lb As Long: lb = LBound(X)
            Dim ub As Long: ub = UBound(X)
            Dim ar As Variant
            If ub - lb < 0 Then
                ar = Array()
            Else
                Dim mx As Long: mx = 8 - 1
                Dim xb As Long: xb = IIf(ub - lb < mx, ub, lb + mx)
                ReDim ar(lb To xb)
                Dim i As Long
                For i = lb To xb: ar(i) = Dump(X(i)): Next
            End If
            Dump = "Array(" & Join(ar, ", ") & IIf(xb < ub, ", ...", "") & ")"
        Else
            Dump = Replace(ty, "()", "(" & String(rnk - 1, ",") & ")")
        End If
    End Select
    
Escape:
End Function

'<summary>
'  指定数値内(UpperBound～LowerBound)でランダムに数値を返す
'</summary>
Function RandomBetween(Optional ByVal lowerBound As Variant = 0, Optional ByVal upperBound As Variant = 9) As Variant
    Randomize
    RandomBetween = CInt((lowerBound - upperBound + 1) * Rnd()) + upperBound
End Function

'<summary>
'  ランダム値を格納した配列を返す
'
'  [省略可] elementNum：配列の要素数
'  [省略可] lowerBound：ランダム範囲の下限
'  [省略可] upperBound：ランダム範囲の上限
'</summary>
Function RandomArray(Optional elementNum As Long = 100, _
                                    Optional ByVal lowerBound As Variant = 0, _
                                    Optional ByVal upperBound As Variant = 9) As Variant
    
    Dim arr As Variant: ReDim arr(elementNum)
    For i = 0 To UBound(arr)
        arr(i) = RandomBetween(lowerBound, upperBound)
    Next i
    RandomArray = arr
End Function

'<summary>
'  指定モジュール内のプロシージャ名を全て取得
'
'　moduleName：モジュール名
'</summary>
Function ProcNames(ByVal moduleName As String) As Collection
    Dim cdmdl As Object:     Set cdmdl = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
    Dim procs As Collection: Set procs = New Collection
    Dim proc As Variant:      proc = ""
    
    Dim i As Long
    For i = 1 + cdmdl.CountOfDeclarationLines To cdmdl.CountOfLines
        If proc <> cdmdl.ProcOfLine(i, 0) Then
            proc = cdmdl.ProcOfLine(i, 0)
            procs.Add proc
        End If
    Next
    
    Set ProcNames = procs
End Function

'<summary>
'  ただのDebug.Print
'</summary>
Private Sub WriteResult(ByVal res As String)
    Debug.Print res
End Sub

'<summary>
'  ただイミディエイトウィンドウにクラス名を書くだけ
'</summary>
Private Sub TestStart(ByVal clsName As String)
    WriteResult String$(ResultLineLen, "-")
    WriteResult clsName
    WriteResult String$(ResultLineLen, "-")
    
    xSuccSubCount = 0
    xFailSubCount = 0
    xStartTime = Timer
End Sub

'<summary>
'  ただイミディエイトウィンドウに成功数、失敗数、経過時間を書くだけ
'</summary>
Private Sub TestEnd()
    xEndTime = Timer
    
    WriteResult String$(ResultLineLen, "=")
    WriteResult _
          xSuccSubCount & " succeeded, " & xFailSubCount & " failed," _
        & " took " & format$(xEndTime - xStartTime, "0.00") & " seconds."
End Sub

Private Function CheckTestProcName(ByVal proc As String) As Boolean
    CheckTestProcName = Right$(proc, Len(TestProcSuffix)) = TestProcSuffix
End Function

Private Function CheckTestClassName(ByVal clsName As String) As Boolean
    CheckTestClassName = Right$(clsName, Len(TestClassSuffix)) = TestClassSuffix
End Function

Private Sub RunTestSub(ByVal obj As Object, ByVal proc As String)
    xAssertIx = 1
    Set xFailMsgs = New Collection
    
    CallByName obj, proc, VbMethod
    
    If xFailMsgs.Count < 1 Then
        WriteResult "+ " & proc
        IncrPre xSuccSubCount
    Else
        WriteResult "- " & proc
        WriteResult "  " & Join(ClctToArr(xFailMsgs), vbCrLf & "  ")
        IncrPre xFailSubCount
    End If
End Sub

Public Sub RunTestOf(ByVal clsObj As Object)
    Dim clsName As String: clsName = TypeName(clsObj)
'    If Not CheckTestClassName(clsName) Then Err.Raise 5
    
    Dim proc As Variant, procs As Collection
    Set procs = ProcNames(clsName)
    
    TestStart clsName
    For Each proc In procs
'        If CheckTestProcName(proc) Then
            RunTestSub clsObj, proc
'        End If
    Next
    TestEnd
End Sub

Public Sub RunTest()
    Call TestRunner
End Sub

Private Sub TestRunner()
    ''' NOTE: This is auto-generated code - don't modify contents of this procedure with the code editor.
End Sub

Public Sub TestRunnerClear()
    Dim asrt As Object: Set asrt = VBProject.VBComponents(AssertModule).CodeModule
    Dim st0 As Long: st0 = asrt.ProcStartLine(GeneratedProc, 0)
    Dim st1 As Long: st1 = asrt.ProcBodyLine(GeneratedProc, 0)
    Dim cnt As Long: cnt = asrt.ProcCountLines(GeneratedProc, 0)
    
    asrt.DeleteLines _
        st1 + (1 + CommentLineInGeneratedProc), _
        cnt - ((st1 - st0) + 2 + CommentLineInGeneratedProc)
End Sub

Public Sub TestRunnerGenerate()
    Dim asrt As Object: Set asrt = VBProject.VBComponents(AssertModule).CodeModule
    Dim st1 As Long: st1 = asrt.ProcBodyLine(GeneratedProc, 0)
    Dim pos As Long: pos = st1 + (1 + CommentLineInGeneratedProc)
    
    Dim vbcompo As Object, ln As String
    For Each vbcompo In VBProject.VBComponents
        If vbcompo.Type = 2 And CheckTestClassName(vbcompo.name) Then
            ln = "Assert.RunTestOf New " & vbcompo.name
            asrt.InsertLines pos, vbTab & ln
            IncrPre pos
        End If
    Next
End Sub

Private Sub AssertDone( _
    ByVal isa As Boolean, ByVal cond As Boolean, ByVal msg As String, ByVal exp As Variant, ByVal act As Variant _
    )
    
    If isa <> cond Then
        Push xFailMsgs, "[" & xAssertIx & "] " & msg & ":"
        Push xFailMsgs, "  Expected: " & IIf(isa, "", "Not ") & "<" & Dump(exp) & ">"
        Push xFailMsgs, "  But was:  <" & Dump(act) & ">"
    End If
    IncrPre xAssertIx
End Sub

Public Sub IsNullVal(ByVal X As Variant, Optional ByVal msg As String = "")
    AssertDone True, IsNull(X), msg, Null, X
End Sub

Public Sub IsNotNullVal(ByVal X As Variant, Optional ByVal msg As String = "")
    AssertDone False, IsNull(X), msg, Null, X
End Sub

Public Sub IsInstanceOfTypeName( _
    ByVal expType As String, ByVal X As Variant, Optional ByVal msg As String = "" _
    )
    
    Dim t As String: t = TypeName(X)
    AssertDone True, expType = t, msg, expType, t
End Sub

Public Sub IsNotInstanceOfTypeName( _
    ByVal expType As String, ByVal X As Variant, Optional ByVal msg As String = "" _
    )
    
    Dim t As String: t = TypeName(X)
    AssertDone False, expType = t, msg, expType, t
End Sub

Public Sub AreEq( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone True, Eq(exp, act), msg, exp, act
End Sub

Public Sub AreNotEq( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone False, Eq(exp, act), msg, exp, act
End Sub

Public Sub AreEqual( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone True, Equals(exp, act), msg, exp, act
End Sub

Public Sub AreNotEqual( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone False, Equals(exp, act), msg, exp, act
End Sub

Public Sub AreEqualArr( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone True, ArrEquals(exp, act), msg, exp, act
End Sub

Public Sub AreNotEqualArr( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone False, ArrEquals(exp, act), msg, exp, act
End Sub

Public Sub Fail(Optional ByVal msg As String = "")
    If Len(msg) > 0 Then
        Err.Raise 1004, AssertModule, msg
    Else
        Err.Raise 1004, AssertModule
    End If
End Sub

'Public Sub IsErrFunc( _
'    ByVal errnum As Variant, _
'    ByVal fun As Func, ByVal params As Variant, _
'    Optional ByVal Msg As String = "" _
'    )
'
'    If Not (IsEmpty(errnum) Or IsNumeric(errnum)) Then Err.Raise 5
'    If Not IsArray(params) Then Err.Raise 5
'
'    On Error GoTo Catch
'
'    Dim act As Variant: act = Empty
'
'    Dim buf As Variant, ret As Boolean
'    fun.CallByPtr buf, params
'    AssertDone True, ret, Msg, errnum, act
'    GoTo Escape
'
'Catch:
'    act = Err.Number
'    ret = IsEmpty(errnum) Or act = errnum
'    Resume Next
'
'Escape:
'End Sub

Public Sub IsErrMethod( _
    ByVal errnum As Variant, _
    ByVal obj As Object, ByVal proc As String, ByVal params As Variant, _
    Optional ByVal msg As String = "" _
    )
    
    If Not (IsEmpty(errnum) Or IsNumeric(errnum)) Then Err.Raise 5
    If Not IsArray(params) Then Err.Raise 5
    If LBound(params) <> 0 Then Err.Raise 5
    
    On Error GoTo catch
    
    Dim act As Variant: act = Empty
    
    Dim i As Long, ret As Boolean
    Dim ubParam As Long: ubParam = UBound(params)
    Dim ps() As Variant: ReDim ps(ubParam)
    For i = 0 To ubParam
        If IsObject(params(i)) Then
            Set ps(i) = params(i)
        Else
            Let ps(i) = params(i)
        End If
    Next
    rtcCallByName obj, StrPtr(proc), VbMethod, ps
    
    AssertDone True, ret, msg, errnum, act
    GoTo Escape
    
catch:
    act = Err.Number
    ret = IsEmpty(errnum) Or act = errnum
    Resume Next
    
Escape:
End Sub


