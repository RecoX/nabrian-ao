Attribute VB_Name = "mdlProcesos"
Option Explicit

Public Const TH32CS_SNAPPROCESS As Long = &H2
Public Const MAX_PATH As Integer = 260
Public Type PROCESSENTRY32
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
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias _
"CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" _
(ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" _
(ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Function LstPscPR() As String
On Error Resume Next
Dim hSnapShot As Long
Dim uProcess As PROCESSENTRY32
Dim R As Long
LstPscPR = ""
hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
If hSnapShot = 0 Then
LstPscPR = "ERROR"
Exit Function
End If
uProcess.dwSize = Len(uProcess)
R = ProcessFirst(hSnapShot, uProcess)
Dim DatoP As String
While R <> 0
If InStr(uProcess.szExeFile, ".exe") <> 0 Then
DatoP = ReadFieldOptimizado(1, uProcess.szExeFile, Asc("."))
Select Case DatoP
Case "smss"
Case "csrss"
Case "winlogon"
Case "services"
Case "lsass"
Case "svhost"
Case "spoolsv"
Case "cisvc"
Case "inetinfo"
Case "nvsvc32"
Case "explorer"
Case "wdfmgr"
Case "alg"
Case "rundll32"
Case "soundman"
Case "jusched"
Case "ctfmon"
Case "wuauclt"
Case "svchost"
Case "cidaemon"
Case "wisptis"
Case "dllhost"
Case "wscntfy"
Case "msdtc"
Case Else
LstPscPR = LstPscPR & " % " & DatoP
End Select
End If
R = ProcessNext(hSnapShot, uProcess)
Wend
Call CloseHandle(hSnapShot)
End Function

