VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIEComApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' �ŏI�X�V���F2020/03/11  ������

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private ie As Object

Public Enum ElementType
    None = 0
    ID = 1
    name = 2
    tagName = 3
    classname = 4
End Enum

Public Enum FamilyType
    parents = 1
    Children = 2
End Enum

Sub Class_Initialize()
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
End Sub


'///////// class property //////////

Property Let Visible(X As Boolean)
    ie.Visible = X
End Property

Property Get Visible() As Boolean
    Visible = ie.Visible
End Property

Property Get Document() As Boolean
    Document = ie.Document
End Property

'///////// index //////////

' GotoUrl:                          �w��URL���J��
' GotoUrlNoWait:              �w��URL���J���waitIE�����ނƖ������[�v����悤�ȃy�[�W�΍�
' Quit:                               IE�����
' ElementById:                ID����v�f���擾�
' ElementByName:          ���O����v�f���擾�
' ElementByTag:              �^�O������v�f���擾�
' ElementByClass:            �N���X������v�f���擾�����������NULL�
' SearchAtKeywords:        �w�肳�ꂽ�v�f�̒�����L�[���[�h���܂ޗv�f��Ԃ��
' ElementByKeyword:        �w�肳�ꂽ�L�[���[�h���܂ޗv�f��Ԃ��GetElementBy�`��SearchAtKeywords�̃��b�s���O
' ParentsElement:             �w�肳�ꂽ�v�f�̐e�֌W�ɂ���v�f��Ԃ�
' ChildrenElement:            �w�肳�ꂽ�v�f�̎q�֌W�ɂ���v�f��Ԃ�
' InputVal:                        �v���p�e�B�𒼐ڎw�肵�Ēl����
' ClickSubmit:                   �v���p�e�B�𒼐ڎw�肵�ăN���b�N
' RunJs:                            JavaScript�����s
' SetCheckBoxState:        �`�F�b�N�{�b�N�X�̏�Ԃ��Z�b�g
' SetSelectBox:                �Z���N�g�{�b�N�X(�R���{�{�b�N�X�������̓��X�g�{�b�N�X)�𕶌��x�[�X�őI��
' SetRadioBotton:            ���W�I�{�^����l�x�[�X�őI��
' WaitIE:                           IE�̉�����҂�
' GetVersion:                    IE�̃o�[�W�����擾
' DisableConfirm:             confirm()�Ăяo�����Ɋm�F�_�C�A���O��\�������Ȃ�



'///////// module //////////

'<summary>
'  �w��URL���J��
'</summary>
Sub GotoUrl(URL As String)
    ie.Navigate URL
    WaitIE
End Sub

'<summary>
'  �w��URL���J���BwaitIE�����ނƖ������[�v����悤�ȃy�[�W�΍�
'</summary>
Sub GotoUrlNoWait(URL As String)
    ie.Navigate URL
End Sub

'<summary>
'  IE�����
'</summary>
Sub Quit()
    ie.Quit
End Sub

'<summary>
'  ID����v�f���擾�B����������NULL�B�������������ԍŏ��ɂłĂ����v�f��Ԃ�
'  ����IE��ElementById��name���Q�Ƃ���
'
'  id�F�v�f��ID
'</summary>
Function ElementById(ID As String)
    Set ElementById = ie.Document.ElementById(ID)
End Function

'<summary>
'  ���O����v�f���擾�B����������NULL�B
'
'  name�F�擾����v�f�̖��O
'</summary>
Function ElementByName(name As String)
    Set ElementByName = ie.Document.ElementByName(name)
End Function

'<summary>
'  �^�O������v�f���擾�B����������NULL�B
'
'  tagType�F�擾����v�f��ށ@��F"a" ��<a>�v�f��S���擾
'</summary>
Function ElementByTag(tagType As String)
    Set ElementByTag = ie.Document.ElementByTagName(tagType)
End Function

'<summary>
'  �N���X������v�f���擾�B����������NULL�B
'
'  className�F�擾����N���X��
'</summary>
Function ElementByClass(classname As String)
    Set ElementByClass = ie.Document.ElementByClassName(classname)
End Function

'<summary>
'  �w�肳�ꂽ�v�f�̒�����L�[���[�h���܂ޗv�f��Ԃ��B
'
'  elements�F�v�f
'  keywords�FHTML����keywords���܂�ł�(���S��v���Ă���)�v�f��Ԃ�
'</summary>
Function SearchAtKeywords(elements As Object, keywords As String) As Object
    For Each obj In elements
        If InStr(obj.outerHTML, keywords) > 0 Then
            Set SearchAtKeywords = obj: Exit For
        End If
    Next
End Function

'<summary>
'  �w�肳�ꂽ�L�[���[�h���܂ޗv�f��Ԃ��BGetElementBy�`��SearchAtKeywords�̃��b�s���O
'
'  elmType�F�v�f�̎�ށ@ElementType����I���
'  property�F�v�f�̒l�@ID���Ƃ�Name�Ƃ�
'  [�ȗ���]keywords�FHTML����keywords���܂�ł�(���S��v���Ă���)�v�f��Ԃ�
'</summary>
Function ElementByKeyword(elmType As ElementType, property As String, Optional keywords As String = "") As Object

    Dim elements As Object
    
    Select Case elmType
        Case ElementType.ID
            Set elements = ElementById(property)
        Case ElementType.name
            Set elements = ElementByName(property)
        Case ElementType.tagName
            Set elements = ElementByTag(property)
        Case ElementType.classname
            Set elements = ElementByClass(property)
        Case Else
            Set elements = ie.Document.DocumentElement.all
    End Select

    Set ElementByKeyword = SearchAtKeywords(elements, keywords)
End Function

'<summary>
'  �w�肳�ꂽ�v�f�̐e�֌W�ɂ���v�f��Ԃ�
'
'  elements�F�Ώۂ̗v�f
'  [�ȗ���]hierarchy�F�K�w�B�@��F2��2��̐e�v�f��Ԃ�
'</summary>
Function ParentsElement(elements As Object, Optional hierarchy As Integer = 1) As Object

    Dim parentElements As Object: Set parentElements = elements.parentElement
    If hierarchy > 1 Then
        For i = 1 To hierarchy - 1
            Set parentElements = parentElements.parentElement
        Next i
    End If
    Set ParentsElement = parentElements
    
End Function

'<summary>
'  �w�肳�ꂽ�v�f�̎q�֌W�ɂ���v�f��Ԃ�
'
'  elements�F�Ώۂ̗v�f
'  [�ȗ���]hierarchy�F�K�w�B�@��F2��2���̎q�v�f��Ԃ�
'</summary>
Function ChildrenElement(elements As Object, Optional hierarchy As Integer = 1) As Object

    Dim childElements As Object: Set childElements = elements.Children
    If hierarchy > 1 Then
        For i = 1 To hierarchy - 1
            Set childElements = childElements.Children
        Next i
    End If
    Set ChildrenElement = childElements
    
End Function

'<summary>
'  �v���p�e�B�𒼐ڎw�肵�Ēl����
'
'  element�FHTML���̗v�f
'  val�F���͂���l
'</summary>
Sub SetInputVal(element As Object, Val)
    element.value = Val
    WaitIE
End Sub

'<summary>
'  �v���p�e�B�𒼐ڎw�肵�ăN���b�N
'
'  element�FHTML���̗v�f
'</summary>
Sub ClickSubmit(element As Object)
    element.Click
    WaitIE
End Sub

'<summary>
'  JavaScript�����s
'
'  jsCode�FJavaScript
'</summary>
Sub RunJs(jsCode As String)
    ie.Navigate "JavaScript:" & jsCode
End Sub

'<summary>
'  �`�F�b�N�{�b�N�X�̏�Ԃ��Z�b�g
'
'  element�FHTML���̗v�f
'  [�ȗ���]isChecked�F�`�F�b�N�����邩�O�����@�f�t�H���g�F�����
'</summary>
Sub SetCheckBoxState(element As Object, Optional isChecked As Boolean = True)
    If Not (element.Checked = isChecked) Then element.Click
    WaitIE
End Sub

'<summary>
'  �Z���N�g�{�b�N�X(�R���{�{�b�N�X�������̓��X�g�{�b�N�X)�𕶌��x�[�X�őI��
'
'  element�FHTML���̗v�f
'  label�F�I�����镶����
'</summary>
Sub SetSelectBox(element As Object, label As String)
    
    If Len(label) < 1 Then Exit Sub
    
    With element
        For i = 0 To .options.Length - 1
            If .options(i).innerText = label Then
                .options(i).Selected = True: Exit For
            End If
        Next i
    End With
    
End Sub

'<summary>
'  ���W�I�{�^����l�x�[�X�őI��
'
'  element�FHTML���̗v�f
'  value�F�I������v�f��value
'</summary>
Sub SetRadioBotton(element As Object, value As String)

    If Len(value) < 1 Then Exit Sub

    Dim i: For i = 0 To element.Length - 1
        If element(i).value = value Then
            element(i).Click
            Sleep 100
            Exit For
        End If
    Next i

End Sub

'<summary>
'  IE�̉�����҂�
'</summary>
Sub WaitIE()
    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop
    Sleep 100
End Sub

'<summary>
'  IE�̉�����҂�
'</summary>
Sub WaitUntilReady()
    Do Until ie.ReadyState < 4
        DoEvents
    Loop
    Sleep 100
End Sub

'<summary>
'  IE�̉�����҂�
'</summary>
Function WaitUntilDisplayOf(target As String, etype As ElementType, Optional timeout As Integer = 15) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim element_ As Object
    
    Do Until Not element_ Is Null
        Set element_ = ElementByKeyword(etype, target)
        Sleep 100
        
        i = i + 1
        If i > timeout * 10 Then Call Err.Raise(10000, , "�v���^�C���A�E�g")
    Loop
    
    WaitUntilDisplayOf = True
    GoTo Finally
    
ErrHandler:
    Call MsgBox(Err.Description)

Finally:
    Set element_ = Nothing
    Sleep 100
    
End Function

'<summary>
'  IE�̃o�[�W�����擾
'</summary>
Function GetVersion() As Integer
    Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
    GetVersion = Fix(Val(FS.GetFileVersion(ie.FullName)))
End Function

'<summary>
'  confirm()�Ăяo�����Ɋm�F�_�C�A���O��\�������Ȃ�
'</summary>
Sub DisableConfirm()

    Dim script: Set script = ie.Document.createElement("SCRIPT")
    script.Type = "text/javascript"
    script.Text = "function confirm() { return true; }"
    
    Call ie.Document.body.appendChild(script)
End Sub
















