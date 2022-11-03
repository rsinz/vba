VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelAppSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' �ŏI�X�V���F2018/02/17  ������


Private Declare PtrSafe Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As Long
Private Const HIGH_PRIORITY_CLASS = &H80
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const REALTIME_PRIORITY_CLASS = &H100
Private cn As Object, rs As Object



'///////// enumerated type //////////

Public Enum BoderIndex          ' AddBorder�Ŏg�p�B�r���������ꏊ���w�肷�邽��
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

' Preheart:                      Excel �A�v���P�[�V�����̐ݒ�������D��ɂ��顃}�N�����s�O�ɂ��Ƥ�����������Ȃ邩��
' Cooldown:                    Excel �A�v���P�[�V�����̐ݒ�����ɖ߂��
' Boost:                          Excel �̃v���Z�X�D��x�����ɐݒ�
' Idle:                             Excel �̃v���Z�X�D��x��ʏ�ɐݒ�
' AddBook:                      Excel���J��
' CloseBook:                    Excel�����@�C�ӂŕۑ�����
' StopAutoFunction:        Excel�̎����@�\���~����
' ReStartAutoFunction:    Excel�̎����@�\���ĊJ����
' ArrayToSheet:               �z���Excel�V�[�g�ɓ\��t����
' SetHyperlink:                Excel�V�[�g�Ƀn�C�p�[�����N��ݒ肷��
' LastRow:                      Range����Ƃ����ŉ��ʍs��Ԃ�
' LastCol:                        Range����Ƃ����ŉE���Ԃ�
' ExistsSheet:                  �V�[�g�̑��݃`�F�b�N
' AddSheet:                    �V�[�g�̒ǉ�
' ----�����܂ŊJ���m�F
' CopySheet:                   �V�[�g�̃R�s�[
' DeleteSheet                  �V�[�g�̍폜
' ChangeSheetName�F     �V�[�g�̖��O�ύX
' ClearRowsValue:           �͈͍s�́u�l�v�폜
' ClearColumnsValue:      �͈͗�́u�l�v�폜
' ParentName:                �w�肳�ꂽRange��������V�[�g�̖��O��Ԃ�
' Paint:                          �Z���̓h��Ԃ�
' AddBorder�F                  �r���ǉ�
' CreatePivot:                  �s�{�b�g�e�[�u���̍쐬
' PivotContents:               �s�{�b�g�e�[�u���̃f�[�^����z��ɕϊ�
' ShapeProperty:             �I�[�g�V�F�C�v(�}�`�Ƃ��e�L�X�g�{�b�N�X)�̃v���p�e�B�ύX
' SetPrintArea:                 ����͈͂̐ݒ�
' PrintOut:                        �������
' Merge:                           �����͈̓Z������
' UnMerge:                      �����͈̓Z����������
' FilterOn:                        �w��Range�Ɋ��S��v�t�B���^�[��������
' FilterOff:                        �w��V�[�g�̃t�B���^�[������
' Execute:                       SQL�̎��s




'///////// module //////////

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub

'<summary>
'�@Excel �A�v���P�[�V�����̐ݒ�������D��ɂ���B�}�N�����s�O�ɂ��ƁA�����������Ȃ邩���B
'</summary>
Sub Preheart()
    Call Boost
    Call StopAutoFunction
End Sub

'<summary>
'�@Excel �A�v���P�[�V�����̐ݒ�����ɖ߂��B
'</summary>
Sub Cooldown()
    Call ReStartAutoFunction
    Call Idle
End Sub

'<summary>
'�@Excel �̃v���Z�X�D��x�����ɐݒ�
'</summary>
Sub Boost()
    Call SetPriorityClass(GetCurrentProcess, HIGH_PRIORITY_CLASS)
End Sub

'<summary>
'�@Excel �̃v���Z�X�D��x��ʏ�ɐݒ�
'</summary>
Sub Idle()
    Call SetPriorityClass(GetCurrentProcess, NORMAL_PRIORITY_CLASS)
End Sub

'<summary>
'�@Excel���J���B
'
'�@[�ȗ���]filePath�F�w��p�X��Book���e���v���[�g�Ƃ��ĊJ���B�ȗ����ꂽ�ꍇ�A�V�KBook���J��
'�@[�ȗ���]forcusAddedBook : �J����Book���A�N�e�B�u�ɂ��邩�ǂ����@True:�A�N�e�B�u�ɂ���
'</summary>
Function AddBook(Optional filePath As String = "", Optional forcusAddedBook As Boolean = True) As Workbook

    On Error GoTo ErrorHandler

    Dim parent_ As Workbook: Set parent_ = ThisWorkbook

    If filePath = "" Then       ' �t�@�C���p�X�̎w�肪������ΐV�KBook���J��
        Set AddBook = Workbooks.Add
        
    Else
        If Dir(filePath, vbDirectory) = "" Then Err.Raise 10000, , "�w�肳�ꂽ�p�X�Ƀt�@�C��������܂���ł���"
        
        If IsBookOpened(filePath) Then        ' Book �����ɊJ����Ă����ꍇ
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
'�@Excel�����B
'
'�@bookName�F����Ώۂ�Book�̖��O
'�@isSaveChanged�F�ۑ����邩�ǂ���
'�@[�ȗ���]�@savePath�F�ۑ���̃p�X�B�f�t�H���g�㏑���ۑ�
'�@[�ȗ���]�@saveFormat�F�ۑ��`���B�f�t�H���g.xls
'</summary>
Sub CloseBook(bookName As String, isSaveChanged As Boolean, Optional savePath As String = "", Optional format As XlFileFormat = 0)

    On Error GoTo ErrorHandler

    If Not isSaveChanged Then
        Workbooks(bookName).Close       '�ۑ����Ȃ��ŕ���
    Else
            If savePath = "" Then
                Workbooks(bookName).Save    '�㏑���ۑ����ĕ���
                Workbooks(bookName).Close
            Else
                    Dim split_path As Variant: split_path = Split(savePath, "\")
                    
                    If format = 0 Then
                        Workbooks(bookName).SaveAs fileName:=savePath   '�C�ӂ̕ۑ���ɕۑ����ĕ���
                        Workbooks(split_path(UBound(split_path))).Close
                    Else
                        Workbooks(bookName).SaveAs fileName:=savePath, fileFormat:=format  '�C�ӂ̕ۑ���ɕۑ����ĕ���(�t�H�[�}�b�g�w��)
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
'�@Excel�̎����@�\���~����
'</summary>
Sub StopAutoFunction()

    With Application
'        .Visible = False                        '�S�̂̕\�����~
        .WindowState = xlMinimized   '�u�b�N�ŏ���
        .DisplayAlerts = False              '�A���[�g�̕\�����~
        .statusBar = False                   '�X�e�[�^�X�o�[�̕\���X�V���~
        .ScreenUpdating = False         '�X�N���[���̕`����~
        .EnableEvents = False             '�C�x���g���ꎞ��~
        .Calculation = xlManual           '�v�Z���蓮���[�h�ɂ���
    End With
    
End Sub

'<summary>
'�@Excel�̎����@�\���ĊJ����
'</summary>
Sub ReStartAutoFunction()

    With Application
        .Calculation = xlAutomatic      '�v�Z���������[�h�ɖ߂�
        .EnableEvents = True              '�C�x���g���ĊJ
        .ScreenUpdating = True          '�X�N���[���̕`����ĊJ
        .statusBar = True                    '�X�e�[�^�X�o�[�̕\�����ĊJ
        .DisplayAlerts = True               '�A���[�g�̕\�����ĊJ
        .WindowState = xlMaximized  '�u�b�N�ő剻
'        .Visible = True                         '�S�̂̕\������
    End With

End Sub

'<summary>
'�@�z���Excel�V�[�g�ɓ\��t����
'
'�@baseRange�F��_��Range
'�@arr�F�\��t����z��
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
'�@Excel�V�[�g�Ƀn�C�p�[�����N��ݒ肷��
'
'�@baseRange�F�\��t����Range
'  link�F�A�h���X������
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
'  baseRange����Ƃ����ŉ��ʍs��Ԃ�
'
'�@baseRange�F��ƂȂ�Range
'�@[�ȗ���]  isIgnoreBlank�F�󔒍s�𖳎����邩,�@True�F�󔒍s���\���Ȃ��ōŏI�f�[�^�s��Ԃ�
'</summary>
Function LastRow(ByVal baseRange As range, Optional isIgnoreBlank As Boolean = False) As Long

    If isIgnoreBlank Then
        LastRow = baseRange.Worksheet.Cells(baseRange.Worksheet.Rows.Count, baseRange.column).End(xlUp).Row
    Else
        LastRow = baseRange.End(xlDown).Row
    End If

End Function

'<summary>
'  baseRange����Ƃ����ŉE���Ԃ�
'
'�@baseRange�F��ƂȂ�Range
'�@[�ȗ���]  isIgnoreBlank�F�󔒍s�𖳎����邩,�@True�F�󔒍s���\���Ȃ��ōŏI�f�[�^���Ԃ�
'</summary>
Function LastCol(ByVal baseRange As range, Optional isIgnoreBlank As Boolean = False) As Long

    If isIgnoreBlank Then
        LastCol = baseRange.Worksheet.Cells(baseRange.Row, baseRange.Worksheet.Columns.Count).End(xlToLeft).column
    Else
        LastCol = baseRange.End(xlToRight).column
    End If

End Function

'<summary>
'  sheet�̑��݃`�F�b�N
'
'�@wsName�F�`�F�b�N�ΏۃV�[�g�̖��O
'�@[�ȗ���]  wb�F�`�F�b�N�Ώ�Book�I�u�W�F�N�g
'</summary>
Function ExistsSheet(ByVal wsName As String, Optional ByVal wb As Workbook = Nothing) As Boolean
    
    Dim ws As Worksheet
    
    If TypeName(wb) = "Nothing" Then Set wb = Application.ThisWorkbook
    
    For Each ws In wb.Sheets
        If ws.name = wsName Then ExistsSheet = True: Exit Function
    Next

End Function

'<summary>
'  [private] sheet�̖��O���d�����Ȃ��悤�ɁA�����ɐ�����t�������O��Ԃ�
'
'�@wsName�F�ǉ����悤�Ƃ��Ă�V�[�g�̖��O
'�@[�ȗ���] wb�F�ǉ�����u�b�N�I�u�W�F�N�g �f�t�H���g�F�Ăяo����Book
'�@[�ȗ���] count�F�����̐����B�ċN�p �f�t�H���g�F0
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
'  sheet�̒ǉ�
'
'�@[�ȗ���] wsName�F�ǉ�����V�[�g�̖��O �f�t�H���g�F����
'�@[�ȗ���] wb�F�ǉ�����u�b�N�I�u�W�F�N�g �f�t�H���g�F�Ăяo����Book
'</summary>
Function AddSheet(Optional ByVal wsName As String = "", Optional ByVal wb As Workbook = Nothing) As Worksheet

    On Error GoTo ErrorHandler
    
    If TypeName(wb) = "Nothing" Then Set wb = Application.ThisWorkbook
'    If ExistsSheet(wsName, wb) Then Err.Raise 10000, , "�V�[�g��ǉ����悤�Ƃ��܂������A���ɓ����̃V�[�g�����݂��܂�"
    
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
'  sheet�̃R�s�[
'
'  sourceSheetName�F�R�s�[���̃V�[�g��
'�@afterWsName�F�R�s�[��̃V�[�g��
'�@[�ȗ���]  wb�F�ǉ�����u�b�N�I�u�W�F�N�g
'</summary>
Function CopySheet(ByVal SourceSheetName As String, ByVal afterWsName As String, Optional ByVal wb As Workbook = Nothing) As Worksheet

    On Error GoTo ErrorHandler
    
    If TypeName(wb) = "Nothing" Then Set wb = Application.ThisWorkbook
    
    If Not ExistsSheet(SourceSheetName, wb) Then Err.Raise 10000, , SourceSheetName & vbNewLine & "�Ƃ����V�[�g�����݂��܂���"
    If ExistsSheet(afterWsName, wb) Then Err.Raise 10000, , afterWsName & vbNewLine & "�Ƃ����V�[�g�͊��ɑ��݂��Ă��܂�"
    
    wb.Worksheets(SourceSheetName).Copy after:=wb.Worksheets(SourceSheetName)
    wb.ActiveSheet.name = afterWsName
    
    Set CopySheet = wb.ActiveSheet
        
    GoTo Finally

ErrorHandler:
    MsgBox Err.Description, vbOKOnly + vbExclamation, "�G���[���������܂��� : clsExcelAppSupport.CopySheet"
    CopySheet = Nothing

Finally:
End Function

'<summary>
'  sheet�̍폜
'
'  wsName�F�폜����V�[�g��
'�@[�ȗ���]  wb�F�폜����u�b�N�I�u�W�F�N�g
'</summary>
Sub DeleteSheet(ByVal wsName As String, Optional ByVal wb As Workbook = Nothing)

    On Error GoTo ErrorHandler
    
    If TypeName(wb) = "Nothing" Then Set wb = Application.ThisWorkbook
    
    If Not ExistsSheet(wsName, wb) Then Err.Raise 10000, , wsName & vbNewLine & "�Ƃ����V�[�g�����݂��܂���"
    
    Worksheets(wsName).Delete
        
    GoTo Finally

ErrorHandler:
    MsgBox Err.Description, vbOKOnly + vbExclamation, "�G���[���������܂��� : clsExcelAppSupport.DeleteSheet"

Finally:
End Sub

'<summary>
'  sheet�̖��O�ύX
'
'  beforeName�F�ύX����V�[�g��
'�@afterName�F�ύX��̃V�[�g��
'</summary>
Sub ChangeSheetName(beforeName As String, afterName As String)

    If beforeName = "" Or afterName = "" Then MsgBox "ChangeSheetName:�V�[�g�����w�肵�Ă�������": Exit Sub
    If ExistsSheet(beforeName) = False Then MsgBox "ChangeSheetName:�V�[�g�����݂��܂���ł���": Exit Sub

    Worksheets(beforeName).name = afterName

End Sub

'<summary>
'  firstRowNum�`lastRowNum�͈͍̔s�́u�l�v�폜
'
'  ws�F�ΏۃV�[�g
'  firstRowNum�F�폜�J�n�s��
'�@[�ȗ���]  lastRowNum�F�폜�I�[�s���B�w�肪�Ȃ���΃A�v���P�[�V������̍ŏI�s
'</summary>
Sub ClearRowsValue(ws As Worksheet, firstRowNum As Long, Optional lastRowNum As Long = 0)

    If lastRowNum = 0 Then lastRowNum = Rows.Count
    
    If (firstRowNum = 0) Or (firstRowNum > lastRowNum) Then Exit Sub
    
    ws.Activate
    ws.range(Rows(firstRowNum), Rows(lastRowNum)).ClearContents

End Sub

'<summary>
'  firstColumnNum�`lastColumnNum�͈̔͗�́u�l�v�폜
'
'  ws�F�ΏۃV�[�g
'  firstColumnNum�F�폜�J�n��
'�@[�ȗ���]  lastColumnNum�F�폜�I�[�񐔁B�w�肪�Ȃ���΃A�v���P�[�V������̍ŏI��
'</summary>
Sub ClearColumnsValue(ws As Worksheet, firstColumnNum As Long, Optional lastColumnNum As Long = 0)

    If lastColumnNum = 0 Then lastColumnNum = Columns.Count
    
    If (firstColumnNum = 0) Or (firstColumnNum > lastColumnNum) Then Exit Sub
    
    ws.Activate
    ws.range(Columns(firstColumnNum), Columns(lastColumnNum)).ClearContents

End Sub

'<summary>
'  �w�肳�ꂽRange��������V�[�g�̖��O��Ԃ�
'
'  rngName�FRange�̖��O
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

    If sheetName = "" Then MsgBox rngName & " �Ƃ������O��Range��������܂���ł����B" & vbNewLine & _
                                                        "���� -> ���O�̊Ǘ� ���m�F���Ă��������B"
    ParentName = sheetName
    
End Function

'<summary>
'  �Z���̓h��Ԃ�
'
'  targetRange�F�h��Ԃ�Range
'�@[�ȗ���]  colors�F�F�@�f�t�H���g���F
'</summary>
Sub Paint(targetRange As range, Optional colors As XlRgbColor = rgbYellow)
    targetRange.Interior.color = colors
End Sub

'<summary>
'  �r���ǉ�
'
'  targetRange�F�r��������Range
'�@[�ȗ���]  boderPlace�F�r���������ꏊ�@�f�t�H���g�F���͒���
'�@[�ȗ���]  style�F�����_���j���Ƃ��@       �f�t�H���g�F����
'�@[�ȗ���]  weights�F�����@                     �f�t�H���g�F�ɍ�
'�@[�ȗ���]  colors�F�F�@                          �f�t�H���g�F��
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
            Case Else: MsgBox ("error:AddBorder" & vbNewLine & "BoderIndex�͈̔͊O�ł�"): Exit Sub
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
'  �s�{�b�g�e�[�u���̍쐬�@�V�����V�[�g������Ă����Ƀs�{�b�g�e�[�u�����쐬���܂�
'
'  sourceRange�F���f�[�^Range�͈�
'  pivotName�F�s�{�b�g�e�[�u���̖��O
'�@[�ȗ���]�@rowField�F�e�[�u�����ōs�ɐݒ肷��񖼁@��FArray("H1","H2")�@��H1��H2���s�ɐݒ�
'�@[�ȗ���]�@colField�F�e�[�u�����ŗ�ɐݒ肷��񖼁@�g�����͓���
'�@[�ȗ���]�@dataField�F�e�[�u�����Œl�ɐݒ肷��񖼁@�g�����͓���
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
        MsgBox "CreatePivot: �t�B�[���h�w���Array�^�ɂ��Ă�������" & vbNewLine & "��FArray(""Header1"",""Header2"")"
        CheckPivotField = False
        Exit Function
    ElseIf VarType(fieldData) = vbArray + vbVariant Then
        For Each oneData In fieldData
            If VarType(oneData) <> vbString Then
                MsgBox "CreatePivot: �t�B�[���h�v�f�͕�����ɂ��Ă�������" & vbNewLine & "��FArray(""Header1"",""Header2"")"
                CheckPivotField = False
                Exit Function
            End If
        Next
    End If
    
    CheckPivotField = True
End Function

'<summary>
'  �s�{�b�g�e�[�u���̃f�[�^����z��ɕϊ�
'
'  pivotName�F�s�{�b�g�e�[�u���̖��O
'�@[�ȗ���] wsName�F�s�{�b�g������V�[�g���@�f�t�H���g�F�A�N�e�B�u�V�[�g
'</summary>
Function PivotContents(pivotName As String, Optional ByVal wsName As String = "") As Variant

    If wsName = "" Then
        wsName = ActiveSheet.name
    ElseIf ExistsSheet(wsName) = False Then
        MsgBox "PivotContents:" & wsName & "�V�[�g�����݂��܂���ł���"
        Exit Function
    End If
        
    Set DataArea = Worksheets(wsName).PivotTables(pivotName).TableRange1
  
    Set DataArea = DataArea.Offset(1, 0)
    Set DataArea = DataArea.Resize(DataArea.Rows.Count - 1)
  
    PivotContents = DataArea
End Function

'<summary>
'  �I�[�g�V�F�C�v(�}�`�Ƃ��e�L�X�g�{�b�N�X)�̃v���p�e�B�ύX
'
'  shapeName�F�I�[�g�V�F�C�v�̖��O
'�@[�ȗ���] wsName�F�I�[�g�V�F�C�v������V�[�g���@�f�t�H���g�F�A�N�e�B�u�V�[�g
'�@[�ȗ���] text�F���͂�����������
'�@[�ȗ���] backColor�F�ύX��̔w�i�F
'</summary>
Sub ShapeProperty(shapeName As String, Optional ByVal wsName As String = "", _
                                            Optional Text As String = "  ", Optional backColor As XlRgbColor = -1)

    If wsName = "" Then wsName = ActiveSheet.name
    If ExistsSheet(wsName) = False Then
        MsgBox "ShapeProperty: " & wsName & "�V�[�g��������܂���ł���"
        Exit Sub
    End If

    Set obj = Worksheets(wsName).Shapes(shapeName)

    If Text <> "  " Then
        obj.OLEFormat.Object.Text = Text
    End If
    
    If backColor > 0 And backColor < 16777216 Then        ' 16777215 = rgbWhite ��Max
        Dim RED As Integer: RED = backColor Mod 256
        Dim green As Integer: green = Int(backColor / 256) Mod 256
        Dim blue As Integer: blue = Int(backColor / 256 / 256)
        obj.Fill.ForeColor.RGB = RGB(RED, green, blue)
    End If

End Sub

'<summary>
'  ����͈͂̐ݒ�
'
'  rangeArea�F����͈͂�ݒ肷��Range�͈�
'�@[�ȗ���] wsName�F�ݒ肷��V�[�g���@�f�t�H���g�F�A�N�e�B�u�V�[�g
'</summary>
Sub SetPrintArea(rangeArea As range, Optional wsName As String = "")

    If wsName = "" Then wsName = ActiveSheet.name
    If ExistsSheet(wsName) = False Then
        MsgBox "SetPrintArea: " & wsName & "�V�[�g��������܂���ł���"
        Exit Sub
    End If
    
    Worksheets(wsName).PageSetup.PrintArea = rangeArea.address
End Sub

'<summary>
'  ����͈͂̃A�h���X�擾
'
'  ws�F����͈͂��ݒ肳��Ă���V�[�g�I�u�W�F�N�g
' [�ȗ���] style : A1�`�� or R1C1�`��
'</summary>
Function PrintAddress(ws As Worksheet, Optional style As XlReferenceStyle = XlReferenceStyle.xlA1) As String
    PrintAddress = range(ws.PageSetup.PrintArea).address(, , style)
End Function

'<summary>
'  ����͈͂̃����W�擾�@���Ă����� Range("Print_Area")�ł��񂶂�ˁH
'
'  ws�F����͈͂��ݒ肳��Ă���V�[�g�I�u�W�F�N�g
'</summary>
Function PrintRange(ws As Worksheet) As range
    Set PrintRange = ws.range(PrintAddress(ws))
End Function

'<summary>
'  �������
'
'�@[�ȗ���] wsName�F�ݒ肷��V�[�g���@�f�t�H���g�F�A�N�e�B�u�V�[�g
'�@[�ȗ���] isColouring�F�J���[������邩�@�f�t�H���g�F����
'�@[�ȗ���] verticalOrHorizon�F����̌����@�f�t�H���g�F�c
'�@[�ȗ���] paperSize�F������̑傫���@�f�t�H���g�FA4
'�@[�ȗ���] paperCount�F��������@�f�t�H���g�F1��
'�@[�ȗ���] isPreview�F����v���r���[��\�����邩�@�f�t�H���g�F�\���Ȃ��ł��̂܂܈��
'</summary>
Sub PrintOut(Optional wsName As String = "", Optional isColouring As Boolean = False, _
                    Optional verticalOrHorizon As XlPageOrientation = XlPageOrientation.xlPortrait, _
                    Optional paperSize As XlPaperSize = XlPaperSize.xlPaperA4, _
                    Optional paperCount As Integer = 1, _
                    Optional isPreview As Boolean = False)

    If wsName = "" Then wsName = ActiveSheet.name
    If ExistsSheet(wsName) = False Then
        MsgBox "PrintOut: " & wsName & "�V�[�g��������܂���ł���"
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
'  �����͈̓Z������
'
'  ranges�F��������Range�͈� �����w���
'</summary>
Sub Merge(ParamArray ranges())
    For i = 0 To UBound(ranges)
        ranges(i).Merge
    Next i
End Sub

'<summary>
'  �����͈̓Z����������
'
'  ranges�F������������Range�͈� �����w���
'</summary>
Sub UnMerge(ParamArray ranges())
    For i = 0 To UBound(ranges)
        ranges(i).UnMerge
    Next i
End Sub

'<summary>
'  �w��Range�Ɋ��S��v�t�B���^�[��������
'
'  ranges�F�t�B���^�[��������w�b�_�[Range
'  [�ȗ���]column�F�t�B���^�[���������
'  [�ȗ���]cmd�F����
'</summary>
Sub FilterOn(ranges As range, Optional column As Long = 0, Optional cmd As Variant = "")
    ranges.AutoFilter Field:=column, Criteria1:=cmd
End Sub

'<summary>
'  �w��V�[�g�̃t�B���^�[������
'
'  sheetName�F��������V�[�g�̖��O
'</summary>
Sub FilterOff(sheetName As String)
    Worksheets(sheetName).AutoFilterMode = False
End Sub

'<summary>
'  baseRange����_�Ƃ��������W�͈͂̃A�h���X�������Ԃ�
'
'  baseRange�F��_�̃����W
'  [�ȗ���] style : �`���BA1 ���AR1C1�B�K��l A1�`��
'  [�ȗ���] ignoreRowBlank : �󔒖����B�K��l �s�����̋󔒂𖳎�����e�[�u���͈�
'  [�ȗ���] ignoreColBlank : �󔒖����B�K��l ������̋󔒂𖳎�����e�[�u���͈�
'</summary>
Function RangeAddress(baseRange As range, Optional style As XlReferenceStyle = XlReferenceStyle.xlA1, _
                                        Optional ignoreRowBlank As Boolean = False, Optional ignoreColBlank As Boolean = False) As String
    
    Dim myLastRow As Long: myLastRow = LastRow(baseRange, ignoreRowBlank)
    Dim myLastCol As Long: myLastCol = LastCol(baseRange, ignoreColBlank)
    
    RangeAddress = range(Cells(baseRange.Row, baseRange.column), Cells(myLastRow, myLastCol)).address(, , style)
    
End Function


'<summary>
'  SQL�̎��s

'  sql : SQL�\��
'  [�ȗ���] Book : �\�[�X�ɂȂ郏�[�N�u�b�N

' SQL��1 �t�B���^�[�FSELECT F1,F3 FROM [Sheet1$B3:E20] WHERE F1 = 'Man' and (F2 >= 20 and F2 < 30) ORDER BY F4, F3 DESC
' SQL��2 �����[�FSELECT * FROM [Sheet2$],[Sheet3$] WHERE [Sheet2$].F2 = [Sheet3$].F1 ORDER BY [Sheet2$].F1
' ��F���͗�
'</summary>
Function Execute(sql As String, Optional Book As Workbook = Nothing) As Variant

    On Error GoTo ErrorHandler
    
    ' -- ���R�[�h�Z�b�g�ݒ�
    If Book Is Nothing Then Set Book = ThisWorkbook
    If cn Is Nothing Then Call OpenSource(Book)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open Source:=sql, ActiveConnection:=cn
    
    ' -- �z��ϊ�
    Dim arr As Variant: arr = rs.GetRows()
    If HasNull(arr) Then Call ReplaceNullTo("", arr)        ' Recordset��NULL���܂܂�Ă�ƁATranspose�ŃG���[�ɂȂ�܂��̂ŁA�󔒂ɒu�������܂��B
    arr = WorksheetFunction.Transpose(arr)
    
    Execute = arr
    
    GoTo Finally
    
ErrorHandler:
    If rs.State = 0 Then GoTo Finally       ' 0 = Close���
    If rs.EOF Then GoTo Finally
    MsgBox Err.Description, vbOKOnly + vbExclamation, "�G���[���������܂��� : clsExcelAppSupport.Execute"
Finally:
End Function

'<summary>
'  private] Excel�f�[�^�\�[�X�Ɛڑ����Ă݂�

'  cn : ADODB.Connection �I�u�W�F�N�g
'  wb : �\�[�X�ɂȂ郏�[�N�u�b�N
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
'  Private]  �\�[�X�Ƃ̐ڑ������B

'  cn : ADODB.Connection �I�u�W�F�N�g
'</summary>
Private Sub CloseSource(cn As Object)
    If (cn.State = 1) Then cn.Close
End Sub


'<summary>
'�@Private] �z���NULL�����邩���肷��BNULL�������Transpose�ŃG���[���o��̂ŁA����𗘗p����
'
'�@���肷��z��F���̂܂��
'</summary>
Private Function HasNull(���肷��z�� As Variant) As Boolean

    On Error GoTo ArrayHasNull
    Dim v As Variant: v = WorksheetFunction.Transpose(���肷��z��)
    
    HasNull = False
    Exit Function

ArrayHasNull:
    HasNull = True
End Function

'<summary>
'�@private] �z�����Null�������l�ɒu���B
'
'�@�u����̒l�F�u����̒l
'�@�z��F�Ώۂ̔z��
'</summary>
Private Function ReplaceNullTo(ByVal �u����̒l As Variant, ByVal �z�� As Variant) As Variant
    
    For i = LBound(�z��, 1) To UBound(�z��, 1)
        For j = LBound(�z��, 2) To UBound(�z��, 2)
            If IsNull(�z��(i, j)) Or IsEmpty(�z��(i, j)) Then �z��(i, j) = �u����̒l
        Next j
    Next i

    ReplaceNullTo = �z��

End Function












