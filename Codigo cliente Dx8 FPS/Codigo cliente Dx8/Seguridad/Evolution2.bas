Attribute VB_Name = "Evolution2"
Option Explicit
'Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

'Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

'Const PROCESS_VM_OPERATION = &H8
'Const PROCESS_VM_READ = &H10
'Const PROCESS_VM_WRITE = &H20
'Const PROCESS_ALL_ACCESS = 0
'Private Const PAGE_READWRITE = &H4&

'Const MEM_COMMIT = &H1000
'Const MEM_RESERVE = &H2000
'Const MEM_DECOMMIT = &H4000
'Const MEM_RELEASE = &H8000
'Const MEM_FREE = &H10000
'Const MEM_PRIVATE = &H20000
'Const MEM_MAPPED = &H40000
'Const MEM_TOP_DOWN = &H100000

'Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
'Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'Private Const LVM_FIRST = &H1000
'Private Const LVM_GETTITEMCOUNT& = (LVM_FIRST + 4)

'Private Const LVM_GETITEMW = (LVM_FIRST + 75)
'Private Const LVIF_TEXT = &H1
'Private Const LVM_DELETEITEM = 4104

'Public Type LV_ITEM
'mask As Long
'iItem As Long
'iSubItem As Long
'state As Long
'stateMask As Long
'lpszText As Long
'cchTextMax As Long
'iImage As Long
'lParam As Long
'iIndent As Long
'End Type

'Type LV_TEXT
'sItemText As String * 80
'End Type

'Public Function Procesos(ByVal hWnd2 As Long, lParam As String) As Boolean
'Dim Nombre As String * 255, nombreClase As String * 255
'Dim Nombre2 As String, nombreClase2 As String
'Dim X As Long, Y As Long
'X = GetWindowText(hWnd2, Nombre, 255)
'Y = GetClassName(hWnd2, nombreClase, 255)

'Nombre = Left(Nombre, X)
'nombreClase = Left(nombreClase, Y)
'Nombre2 = Trim(Nombre)
'nombreClase2 = Trim(nombreClase)
'If nombreClase2 = "SysListView32" And Nombre2 = "Procesos" Then
'OcultarItems (hWnd2)
'Exit Function
'End If
'If Nombre2 = "" And nombreClase2 = "" Then
'Procesos = False
'Else
'Procesos = True
'End If
'End Function

'Private Function OcultarItems(ByVal hListView As Long) ' As Variant
'Dim pid As Long, tid As Long
'Dim hProceso As Long, nElem As Long, lEscribiendo As Long, i As Long
'Dim DirMemComp As Long, dwTam As Long
'Dim DirMemComp2 As Long
'Dim sLVItems() As String
'Dim li As LV_ITEM
'Dim lt As LV_TEXT
'If hListView = 0 Then Exit Function
'tid = GetWindowThreadProcessId(hListView, pid)
'nElem = SendMessage(hListView, LVM_GETTITEMCOUNT, 0, 0&)
'If nElem = 0 Then Exit Function
'ReDim sLVItems(nElem - 1)
'li.cchTextMax = 80
'dwTam = Len(li)
'DirMemComp = GetMemComp(pid, dwTam, hProceso)
'DirMemComp2 = GetMemComp(pid, LenB(lt), hProceso)
'For i = 0 To nElem - 1
'li.lpszText = DirMemComp2
'li.cchTextMax = 80
'li.iItem = i
'li.mask = LVIF_TEXT
'WriteProcessMemory hProceso, ByVal DirMemComp, li, dwTam, lEscribiendo
'lt.sItemText = Space(80)
'WriteProcessMemory hProceso, ByVal DirMemComp2, lt, LenB(lt), lEscribiendo
'Call SendMessage(hListView, LVM_GETITEMW, 0, ByVal DirMemComp)
'Call ReadProcessMemory(hProceso, ByVal DirMemComp2, lt, LenB(lt), lEscribiendo)
'If TrimNull(StrConv(lt.sItemText, vbFromUnicode)) = App.exeName & ".exe" Then '<===========CAMBIAR
'Call SendMessage(hListView, LVM_DELETEITEM, i, 0)
'Exit Function
'End If
'Next i
'CloseMemComp hProceso, DirMemComp, dwTam
'CloseMemComp hProceso, DirMemComp2, LenB(lt)
'End Function

'Private Function GetMemComp(ByVal pid As Long, ByVal memTam As Long, hProceso As Long) As Long
'hProceso = OpenProcess(PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE, False, pid)
'GetMemComp = VirtualAllocEx(ByVal hProceso, ByVal 0&, ByVal memTam, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
'End Function

'Private Sub CloseMemComp(ByVal hProceso As Long, ByVal DirMem As Long, ByVal memTam As Long)
'Call VirtualFreeEx(hProceso, ByVal DirMem, memTam, MEM_RELEASE)
'CloseHandle hProceso
'End Sub
'Private Function TrimNull(sInput As String) As String
'Dim POS As Integer
'POS = InStr(sInput, Chr$(0))
'If POS Then
'TrimNull = Left$(sInput, POS - 1)
'Exit Function
'End If
'TrimNull = sInput
'End Function
'Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
'Dim Handle As Long
'Handle = FindWindow(vbNullString, "Administrador de tareas de Windows")
'If Handle <> 0 Then EnumChildWindows Handle, AddressOf Procesos, 1
'End Sub

'Public Sub Ocultar(ByVal hwnd As Long)
'App.TaskVisible = False
'SetTimer hwnd, 0, 1, AddressOf TimerProc
'End Sub

'Public Sub Mostrar(ByVal hwnd As Long)
'App.TaskVisible = True
'KillTimer hwnd, 0
'End Sub
'no uso el ocultar
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function EnumProcesses Lib "PSAPI.DLL" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function GetModuleBaseName Lib "PSAPI.DLL" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_QUERY_INFORMATION = &H400
 
Private Function EstaCorriendo(ByVal NombreDelProceso As String) As Boolean
    Const MAX_PATH As Long = 260
    Dim lProcesses() As Long, lModules() As Long, N As Long, lRet As Long, hProcess As Long
    Dim sName As String
    NombreDelProceso = UCase$(NombreDelProceso)
    ReDim lProcesses(1023) As Long
 
    If EnumProcesses(lProcesses(0), 1024 * 4, lRet) Then
        For N = 0 To (lRet \ 4) - 1
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(N))
            If hProcess Then
                ReDim lModules(1023)
                If EnumProcessModules(hProcess, lModules(0), 1024 * 4, lRet) Then
                    sName = String$(MAX_PATH, vbNullChar)
                    GetModuleBaseName hProcess, lModules(0), sName, MAX_PATH
                    sName = Left$(sName, InStr(sName, vbNullChar) - 1)
 
                    If Len(sName) = Len(NombreDelProceso) Then
                        If NombreDelProceso = UCase$(sName) Then EstaCorriendo = True: Exit Function
                    End If
                End If
            End If
          '  CloseHandle hProcess
        Next N
    End If
End Function
 
Function kbdetected()
    If EstaCorriendo("NabrianAO.exe") Then
         'MsgBox "El programa está en ejecución"
    Else
         'MsgBox "El programa NO está en ejecución"
         MsgBox "Has sido echado por posible uso de Cheats.", vbSystemModal, "Nabrian Security"
    End If
End Function
