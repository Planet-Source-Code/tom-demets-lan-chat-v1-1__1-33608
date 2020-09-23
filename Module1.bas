Attribute VB_Name = "Module1"
Option Explicit
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Dim ListOfActiveProcess() As PROCESSENTRY32
Public Function GetActiveProcess() As Long
Dim hSnapshot As Long
Dim tProcess As PROCESSENTRY32
Dim R As Long, I As Integer
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
If hSnapshot = 0 Then
    GetActiveProcess = 0
    Exit Function
End If
With tProcess
    .dwSize = Len(tProcess)
    R = ProcessFirst(hSnapshot, tProcess)
    ReDim Preserve ListOfActiveProcess(20)
    Do While R
        I = I + 1
    If I Mod 20 = 0 Then ReDim Preserve ListOfActiveProcess(I + 20)
        ListOfActiveProcess(I) = tProcess
        R = ProcessNext(hSnapshot, tProcess)
    Loop
End With
GetActiveProcess = I
Call CloseHandle(hSnapshot)
End Function
Public Function exePath(ByVal Index As Long) As String
exePath = ListOfActiveProcess(Index).szExeFile
End Function

