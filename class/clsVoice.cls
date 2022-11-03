VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Enum Language
    JP = 0
    EN
End Enum

Private Sub Class_Initialize()
End Sub

Public Sub Speak(msg As String, Optional lang As Language = Language.JP)

    With CreateObject("SAPI.SpVoice")
        Set .Voice = .GetVoices.Item(lang)
        .Speak (msg)
    End With

End Sub












