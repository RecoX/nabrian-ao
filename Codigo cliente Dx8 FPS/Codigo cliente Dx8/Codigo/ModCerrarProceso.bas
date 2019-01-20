Attribute VB_Name = "ModCerrarProceso"
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
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
    szexeFile As String * 260
End Type
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_CREATE_THREAD = &H2
Private Const PROCESS_VM_OPERATION = &H8
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_VM_WRITE = &H20
Private Const PROCESS_DUP_HANDLE = &H40
Private Const PROCESS_CREATE_PROCESS = &H80
Private Const PROCESS_SET_QUOTA = &H100
Private Const PROCESS_SET_INFORMATION = &H200
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Function SearchProcessID(ByVal processName As String)
SearchProcessID = 0
Dim hSnapshot As Long
Dim uProceso As PROCESSENTRY32
Dim res As Long
hSnapshot = CreateToolhelpSnapshot(2&, 0&)
If hSnapshot <> 0 Then
    uProceso.dwSize = Len(uProceso)
    res = ProcessFirst(hSnapshot, uProceso)
    Do While res
        ActualProcess = Left$(uProceso.szexeFile, InStr(uProceso.szexeFile, Chr$(0)) - 1)
        If UCase$(ActualProcess) = UCase$(processName) Then
            SearchProcessID = uProceso.th32ProcessID
        End If
        res = ProcessNext(hSnapshot, uProceso)
    Loop
    Call CloseHandle(hSnapshot)
End If
End Function
Public Sub CloseProcess(ByVal processName As String)
Dim hProcess As Long, iResult As Long
mainProcessID = SearchProcessID(processName)
hProcess = OpenProcess(PROCESS_TERMINATE, True, mainProcessID)
iResult = TerminateProcess(hProcess, 99)
CloseHandle hProcess
End Sub
 
