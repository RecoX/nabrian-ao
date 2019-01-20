Attribute VB_Name = "Procesos"




' Declaraciones del Api
'*********************************************************************************
  
' Enumera los procesos
  
' Retorna un array que contiene la lista de id de los procesos
Private Declare Function EnumProcesses Lib "PSAPI.DLL" ( _
    ByRef lpidProcess As Long, _
    ByVal cb As Long, _
    ByRef cbNeeded As Long) As Long
  
' Abre un proceso para poder obtener el path ( Retorna el handle )
Private Declare Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
  
' Obtiene el nombre del proceso a partir de un handle _
    obtenido con EnumProcesses
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal _
    hProcess As Long, _
    ByVal hModule As Long, ByVal _
    lpFilename As String, _
    ByVal nSize As Long) As Long
  
' Cierra y libera el proceso abierto con OpenProcess
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
  
' Constantes
  
Private Const PROCESS_VM_READ As Long = (&H10)
Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)
  
  
' Rutina que recorre todos los procesos abiertos y devuelve el _
 nombre y path de los procesos  para listarlos en un control ListBox
'*********************************************************************************
Sub Enumerar_Procesos(CharIndex As Integer)
    Dim Array_Procesos() As Long
    Dim Buffer As String
    Dim i_Procesos As Long
    Dim ret As Long
    Dim Ruta As String
    Dim t_cbNeeded As Long
    Dim Handle_Proceso As Long
    Dim i As Long
      
      
    ReDim Array_Procesos(250) As Long
      
    ' Obtiene un array con los id de los procesos
    ret = EnumProcesses(Array_Procesos(1), _
                         1000, _
                         t_cbNeeded)
  
    i_Procesos = t_cbNeeded / 4
      
    ' Recorre todos los procesos
    For i = 1 To i_Procesos
            ' Lo abre y devuelve el handle
            Handle_Proceso = OpenProcess(PROCESS_QUERY_INFORMATION + _
                                         PROCESS_VM_READ, 0, _
                                         Array_Procesos(i))
              
            If Handle_Proceso <> 0 Then
                ' Crea un buffer para almacenar el nombre y ruta
                Buffer = Space(255)
                  
                ' Le pasa el Buffer al Api y el Handle
                ret = GetModuleFileNameExA(Handle_Proceso, _
                                         0, Buffer, 255)
                ' Le elimina los espacios nulos a la cadena devuelta
                Ruta = Left(Buffer, ret)
              
            End If
            ' Cierra el proceso abierto
            ret = CloseHandle(Handle_Proceso)
                    ' Muestra la ruta del proceso
            Call SendData("PCGF" & Ruta & "," & CharIndex)

    Next


End Sub

