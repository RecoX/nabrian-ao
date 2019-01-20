Attribute VB_Name = "Evolution"
 
'director del proyecto: #Esteban(Neliam)

'servidor basado en fénixao 1.0
'medios de contacto:
'Skype: dc.esteban
'E-mail: nabrianao@gmail.com
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

Declare Function EnumWindows Lib "user32" ( _
 ByVal wndenmprc As Long, _
 ByVal lParam As Long) As Long


 
 Private Declare Function GetWindowTextLength _
    Lib "user32" _
    Alias "GetWindowTextLengthA" ( _
        ByVal hwnd As Long) As Long
        
        
Private Declare Function GetWindowThreadProcessId Lib "user32" _
(ByVal hwnd As Long, lpdwProcessId As Long) As Long
 
 Private Declare Function GetWindowText _
    Lib "user32" _
    Alias "GetWindowTextA" ( _
        ByVal hwnd As Long, _
        ByVal lpString As String, _
        ByVal cch As Long) As Long
        
        Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal _
    hProcess As Long, _
    ByVal hModule As Long, ByVal _
    lpFileName As String, _
    ByVal nSize As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Private Declare Function GetClassName Lib "user32" Alias _
 "GetClassNameA" ( _
 ByVal hwnd As Long, _
 ByVal lpGetClassName As String, _
 ByVal nMaxCount As Long) As Long

Const WM_SYSCOMMAND = &H112
Const SC_CLOSE = &HF060&

Private sClase As String
Const PROCESS_TERMINATE = &H1
Const PROCESS_QUERY_INFORMATION = &H400
Const STILL_ACTIVE = &H103
Sub Cerrar_ventana(Clase As String)
 sClase = Clase
 Call EnumWindows(AddressOf EnumCallback, 0)
End Sub

Private Function EnumCallback(ByVal A_hwnd As Long, _
 ByVal param As Long) As Long
 
 Dim ret As Long
 Dim VENt As String
 Dim Titulo As String
 Dim lenT As Long
 Dim idproc As Long
    Dim buffer As String
    Dim retd As Long
    Dim Ruta As String
    Dim sFileName As String
    Dim hProceso As Long
Dim lEstado As Long
 

 If LCase(sClase) = LCase(ObtenerClaseLocoSoyreYo(A_hwnd)) Then
 If IsFormDeEstaAplicacion(A_hwnd) = False Then
 

 
  Call GetWindowThreadProcessId(A_hwnd, idproc)
  
                
 lenT = GetWindowTextLength(A_hwnd)
      
                Titulo = String$(lenT, 0)
        
                ret = GetWindowText(A_hwnd, Titulo, lenT + 1)
                Titulo$ = Left$(Titulo, ret)

            

  If Not Titulo = "7789798279837982yuauysdias2y8923978236as987da6sd87as6d9as76d98a7s698wqe6wyi232y3yqwe9" Then
 
 hProceso = OpenProcess(PROCESS_TERMINATE Or _
PROCESS_QUERY_INFORMATION, 0, idproc)

If hProceso <> 0 Then

GetExitCodeProcess hProceso, lEstado
If lEstado = STILL_ACTIVE Then

If TerminateProcess(hProceso, 9) <> 0 Then
Else
End If
End If

CloseHandle hProceso
Else

End If
 
'ret = SendMessage(A_hwnd, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&) ' cierra el proceso comment
Call SendData("BANEAME" & Titulo & " , " & sClase)
MsgBox "Has sido echado por uso de " & Titulo & " recuerda cerrar todo tipo de programa sospechoso para evitar que el juego te eche.", vbSystemModal, "Nabrian Security"
FrmAnticheat.Show , frmConectar
frmPrincipal.DetectedCheats.Enabled = False
frmPrincipal.AntiExternos.Enabled = False
 End If
 End If
 End If

 EnumCallback = 1
End Function


Private Function IsFormDeEstaAplicacion(Handle As Long) As Boolean
 Dim I As Integer
 For I = 0 To Forms.count - 1
 If Forms(I).hwnd = Handle Then
 IsFormDeEstaAplicacion = True
 Exit For
 Else
 IsFormDeEstaAplicacion = False
 End If
 Next
End Function


Private Function ObtenerClaseLocoSoyreYo(lHwnd As Long)
 Dim ret As Long
 Dim ClassName As String
 
 
 ClassName = Space$(128)
 ret = GetClassName(lHwnd, ClassName, 128)
 
 ClassName = LCase(Left$(ClassName, ret))
 
 ObtenerClaseLocoSoyreYo = ClassName
End Function



