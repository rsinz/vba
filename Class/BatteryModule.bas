Attribute VB_Name = "BatteryModule"

' clsBatteryとセットで使ってね


Private battery As New clsBattery
Private cycle As Date      ' (Application.ontime) 設定時間記憶ため
Private usercycle As Long       ' (Application.ontime) ユーザー指定の周期時間記憶のため

' <summary>
' 電源のモニターを開始する
'
' [省略可] cycleSecond : 監視周期　/デフォルト：180秒
' </summary>
Public Sub StartMoniter(Optional cycleSecond As Long = 180)
    usercycle = cycleSecond
    Call MoniterLoop
End Sub

' <summary>
' 電源のモニターを停止する
' </summary>
Public Sub StopMoniter()
    Call Application.OnTime(EarliestTime:=cycle, _
                                           procedure:="MoniterLoop", _
                                           Schedule:=False)
End Sub

' <summary>[private]
' Application.OnTimeを利用して定時刻にメソッドを実行する
' </summary>
Private Sub MoniterLoop()

    cycle = Now + timeValue(battery.ToHourMinuteSecond(usercycle))
    Call Application.OnTime(EarliestTime:=cycle, _
                                           procedure:="MoniterLoop")
    Call Moniter

End Sub

' <summary>[private]
' 定周期メソッドの中身
' </summary>
Private Sub Moniter()       ' todo : ここにイベントを書く
    Debug.Print format(Now, "hh:mm:ss")
    If battery.IsDenger Then
        Debug.Print "バッテリーが5%以下になっています。"
    End If
End Sub









