VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



' Microsoft Outlook *.* Object を参照

Private Outlook As Outlook.Application 'Object

'///////// index //////////



'///////// module //////////
Private Sub Class_Initialize()
    Set Outlook = New Outlook.Application 'CreateObject("Outlook.Application")
End Sub

Private Sub Class_Terminate()
    Set Outlook = Nothing
End Sub

'<summary>
'  Outlook カレンダーの取得
'  [日付][時刻 件名@場所]の配列で返します。
'
' startDay : 取得する最初の日
' [省略可] endDay : 取得する最後の日。省略するとstartDay以降、全部の予定を取得します。
'</summary>
Public Function GetCalender(startDay As Date, Optional endDay As Date = 0) As Variant

    Dim stores_ As Stores: Set stores_ = Outlook.GetNamespace("MAPI").Stores
    
    Dim store_ As Store
    For Each store_ In stores_
        
        If store_.DisplayName = "インターネット予定表" Then
        
            Dim filter_ As String: filter_ = "[Start] >= '" & startDay & "'" & " AND " & "[End] <= '" & endDay & "'"
            Dim calender_ As items: Set calender_ = store_.Session.GetDefaultFolder(olFolderCalendar).items.Restrict(filter_)
            
            If calender_.Count <> 0 Then
            
                Dim items_ As Object: Set items_ = CreateObject("Scripting.Dictionary")
                Dim appoint_ As AppointmentItem
                
                For Each appoint_ In calender_
                    Dim descript_ As String: descript_ = format(appoint_.Start, "h:mm") & " " & appoint_.Subject & "@" & appoint_.Location
                    
                    If items_.Exists(format(appoint_.Start, "yyyy/m/d")) Then        ' 一旦dictionalyにつっ込みます
                        items_(format(appoint_.Start, "yyyy/m/d")) = items_(format(appoint_.Start, "yyyy/m/d")) & " / " & descript_
                    Else
                        Call items_.Add((format(appoint_.Start, "yyyy/m/d")), descript_)
                    End If
                Next
                
                GetCalender = Dict2Array(items_)
                Set items_ = Nothing: Set appoint_ = Nothing
            
            End If
            
            Set calender_ = Nothing
        End If
    Next

    Set stores_ = Nothing: Set store_ = Nothing
End Function

Private Function Dict2Array(dic As Object) As Variant
    
    Dim keys_ As Variant: keys_ = dic.Keys
    Dim items_ As Variant: items_ = dic.items
    Dim arr_ As Variant: ReDim arr_(0 To UBound(keys_), 0 To 1) As Variant
        
    For i = LBound(arr_) To UBound(arr_)
        arr_(i, 0) = keys_(i)
        arr_(i, 1) = items_(i)
    Next i
    
    Dict2Array = arr_
End Function

'<summary>
'  指定アドレスにメール送信
'
'</summary>
Sub SendTo(address As String, Optional msg As String = "", Optional title As String = "")

    Dim now_ As Date: now_ = Now
    
    With Outlook.CreateItem(olMailItem)
        .Subject = "【bot】" & IIf(title = "", ThisWorkbook.name, title) & "_" & format(now_, "yyyymmddhhmmss")
        .To = address
        .body = now_ & vbNewLine & _
                    "<file://" & ThisWorkbook.path & "\" & ThisWorkbook.name & ">" & vbNewLine & vbNewLine & _
                    "-- auto sending by vba --" & vbNewLine & _
                    msg
        .Save
        '.Send
    End With

End Sub




























