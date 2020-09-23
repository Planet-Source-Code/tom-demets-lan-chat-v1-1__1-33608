Attribute VB_Name = "Module2"
Private Const SND_APPLICATION = &H80
Private Const SND_ALIAS = &H10000
Private Const SND_ALIAS_ID = &H110000
Private Const SND_ASYNC = &H1
Private Const SND_FILENAME = &H20000
Private Const SND_LOOP = &H8
Private Const SND_MEMORY = &H4
Private Const SND_NODEFAULT = &H2
Private Const SND_NOSTOP = &H10
Private Const SND_NOWAIT = &H200
Private Const SND_PURGE = &H40
Private Const SND_RESOURCE = &H40004
Private Const SND_SYNC = &H0
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Sub UserOnline()
    PlaySound App.Path & "\online.WAV", ByVal 0&, SND_FILENAME Or SND_ASYNC
End Sub
Public Sub UserMSG()
    PlaySound App.Path & "\type.WAV", ByVal 0&, SND_FILENAME Or SND_ASYNC
End Sub

