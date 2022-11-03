VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' PDFDesigner Tools API
Private Declare PtrSafe Function GetPDFText Lib "pdftool.dll" (ByVal OpenFileName As String, ByVal SaveFileName As String) As Long


' <summary>
' PDFファイルをテキストファイルに変換
'  戻り値  : 1:成功 -1:失敗 -2:PDFファイルが暗号化されてる
'
' pdfpath : PDFファイルフルパス
' txtpath : txtファイルフルパス
' </summary>
Public Function PDF2T(ByVal pdfpath As String, ByVal txtpath As String) As Long

    Call ChDir(ThisWorkbook.path)
    Dim tmp As String: tmp = Replace(pdfpath, ".pdf", ".tmp", compare:=vbTextCompare)
    Call ChangePDFVersion(pdfpath, tmp)    ' 元のPDFファイルをコピーしてバージョンを1.4形式に変更する
    
    PDF2T = GetPDFText(tmp, txtpath)
    Kill (tmp)
End Function

' <summary>
' PDFのバージョンを1.4形式にする(コピーしたファイルを編集)
'
' pdfpath : 元PDFファイルフルパス
' tmppath : 1.4形式にした後の保存先フルパス
' </summary>
Private Sub ChangePDFVersion(ByVal pdfpath As String, ByVal tmppath As String)
    
    Dim Stream() As Byte: ReDim Stream(FileLen(pdfpath) - 1)
    
    ' PDF読み込み
    Dim FileNo As Integer: FileNo = FreeFile
    
    Open pdfpath For Binary As #FileNo
        Get #FileNo, , Stream
        Stream(5) = "&H31":   Stream(7) = "&H34"        ' PDFの形式を1.4にする
    Close #FileNo
    
    ' PDF出力
    FileNo = FreeFile
    Open tmppath For Binary Access Write As #FileNo
        Put #FileNo, , Stream
    Close #FileNo
    
End Sub

' <summary>
' フォルダ内全PDFテキスト化
'
' [省略可] InitFolderPath : 初期のフォルダパス
' </summary>
Public Sub AllPDF2T(Optional InitFolderPath As String = "")

    'フォルダパスを選択→ファオルダ内のファイル一覧を取得
    Dim target As String
    
    If InitFolderPath = "" Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            If .Show = True Then
                target = .SelectedItems(1)
            End If
        End With
    
    Else
        target = InitFolderPath
        
    End If
    
    Dim fopath As String: fopath = Dir(target & "\*.pdf")
    
    'フォルダ内にあるPDFファイルをテキストファイルに変換
    Do While fopath <> ""
        Call PDF2T(fopath, Replace(fopath, "pdf", "txt"))
        fopath = Dir()
    Loop

End Sub

























