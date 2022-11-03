VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (frequency As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (procTime As Currency) As Long

Private freqMem As Currency
Private timeMem As Double

Private Function GetMicroSecond() As Double

    Dim procTime As Currency
    Call QueryPerformanceCounter(procTime)
    
    GetMicroSecond = procTime / freqMem
    
End Function

Sub Start_()
    Call QueryPerformanceFrequency(freqMem)
    timeMem = GetMicroSecond
End Sub

Sub Lap_()
    Debug.Print GetMicroSecond - timeMem & "[sec]"
End Sub
