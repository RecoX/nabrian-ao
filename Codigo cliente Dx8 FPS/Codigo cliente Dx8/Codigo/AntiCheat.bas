Attribute VB_Name = "AntiCheat"
Option Explicit
 
Private Declare Function OpenProcess Lib "kernel32" (ByVal _
dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long
 
Private Declare Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long
 
Private Declare Function TerminateProcess Lib "kernel32" _
(ByVal hProcess As Long, ByVal uExitCode As Long) As Long
 
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject _
As Long) As Long
 
Private Declare Function GetWindowThreadProcessId Lib "user32" _
(ByVal hWnd As Long, lpdwProcessId As Long) As Long
 
Private Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
 
Private Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
 
Const PROCESS_TERMINATE = &H1
Const PROCESS_QUERY_INFORMATION = &H400
Const STILL_ACTIVE = &H103
 
Public Sub CerrarProceso(TítuloVentana As String)
Dim hProceso As Long
Dim lEstado As Long
Dim idProc As Long
Dim winHwnd As Long
 
winHwnd = FindWindow(vbNullString, TítuloVentana)
If winHwnd = 0 Then
Debug.Print "El proceso no está abierto": Exit Sub
End If
Call GetWindowThreadProcessId(winHwnd, idProc)
 
' Obtenemos el handle al proceso
hProceso = OpenProcess(PROCESS_TERMINATE Or _
PROCESS_QUERY_INFORMATION, 0, idProc)
If hProceso <> 0 Then
' Comprobamos estado del proceso
GetExitCodeProcess hProceso, lEstado
If lEstado = STILL_ACTIVE Then
' Cerramos el proceso
If TerminateProcess(hProceso, 9) <> 0 Then
Debug.Print "Proceso cerrado"
Else
Debug.Print "No se pudo matar el proceso"
End If
End If
' Cerramos el handle asociado al proceso
CloseHandle hProceso
Else
Debug.Print "No se pudo tener acceso al proceso"
End If
End Sub

