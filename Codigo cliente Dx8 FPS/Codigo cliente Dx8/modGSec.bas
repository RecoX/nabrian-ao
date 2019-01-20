Attribute VB_Name = "modGSec"
Option Explicit

'*********************************************
'*********************************************
'********** GSec v1.42 - Anti-cheat **********
'************** GS-Zone (c) 2012 *************
'********** http://www.gs-zone.org ***********
'*********************************************
'*********************************************

' Procedimientos
Public Declare Sub gsCredits Lib "GSec.dll" () ' Abre la ventana de Creditos
Public Declare Sub gsStart Lib "GSec.dll" () ' Inicia la protección
Public Declare Sub gsStop Lib "GSec.dll" () ' Detiene la protección

' Funciones
Public Declare Function gsStatus Lib "GSec.dll" () As Byte ' Devuelve el estado del anticheat
    ' RECOMENDADO: Se recomienda realizar esta función cada 1 seguno o 5 segundos... en un Timer talvez.
    ' ACLARACIÓN: Esta funcion no hace nada especial, solo se fija que esta haciendo el anticheat,
    ' por lo tanto, si se ejecuta una vez cada minuto, no afecta en nada al funcionamiento del anticheat.
    ' Estado:
    ' 0 - Desactivado
    ' 1 - Activado
    ' 2 - Cheat detectado
Public Declare Function gsCheatName Lib "GSec.dll" () As String ' Devuelve el Nombre del cheat asociado a la detección (solo si el estado fue igual 2)
Public Declare Function gsCheatPath Lib "GSec.dll" () As String ' Devuelve el Path del cheat detectado (solo si el estado fue igual 2)
Public Declare Function gsGetGSEC_ID Lib "GSec.dll" () As String  ' Devuelve el ID de identificación unica del usuario

' INSTALACIÓN
'
' GUIA BASADA EN 0.11.5 (adaptar a las versiónes clasicas correspondientes)
'
' - PASO 1 -
' En el modulo General, buscar "    Unload frmCargando" agregar justo debajo...
'    Call gsStart
' - PASO 2 -
' En el mismo modulo (General), buscar "        If GetTickCount - lFrameTimer >= 1000 Then" agregar justo debajo...
'            loopc = gsStatus
'            If loopc <> 0 Then
'                If loopc = 2 Then
'                    If Connected = True Then Call SendData(gsInformar)
'                    Sleep 5
'                    prgRun = False
'                End If
'            Else
'                prgRun = False
'            End If
' - PASO 3 -
' En el mismo modulo (General), buscar "    Call UnloadAllForms" agregar justo arriba...
'    Call gsStop
' - PASO 4 -
' En el formulario frmMain, buscar la función "    If EstadoLogin = E_MODO.CrearNuevoPj Then" agregar justo arriba...
'    Call SendData(gsEnviarID)
' - PASO 5 -
' En el mismo modulo (Mod_TCP), buscar "    Dim tstr2 As String" agregar justo debajo...
'    Call gsProcesar(RData)

' EN EL SERVIDOR:
' - PASO 1 -
' En el modulo Declaraciones, buscar "Public Type UserFlags" agregar justo debajo...
'    GSEC_ID As String
' - PASO 2 -
' En el modulo TCP, buscar "        .Ban = 0" agregar justo arriba...
'        .GSEC_ID = vbNullString
' - PASO 3 -
' En el mismo modulo (TCP), buscar "    If Left$(rData, 13) = "gIvEmEvAlcOde" Then" agregar justo arriba...
'    If Len(rData) > 3 Then
'        Select Case Left$(rData, 3)
'            Case "GID"
'                rData = Right$(rData, Len(rData) - 3)
'                ClientChecksum = Right$(rData, Len(rData) - InStrRev(rData, Chr$(126)))
'                rData = Left$(rData, Len(ClientChecksum))
'                If LenB(UserList(UserIndex).flags.GSec_ID) = 0 Then
'                    UserList(UserIndex).flags.GSec_ID = rData
'                Else
'                    Call CloseSocket(UserIndex, True)
'                End If
'                Exit Sub
'            Case "GAC"
'                rData = Right$(rData, Len(rData) - 3)
'                ClientChecksum = Right$(rData, Len(rData) - InStrRev(rData, Chr$(126)))
'                rData = Left$(rData, Len(ClientChecksum))
'                UserList(UserIndex).flags.Ban = 1
'                Call LogBanFromName("GSec-Anticheat", UserIndex, "ANTICHEAT detecto " & rData)
'                Call SendData(SendTarget.ToAdmins, 0, 0, "||GSec> ANTICHEAT ha baneado a " & UserList(UserIndex).Name & "." & FONTTYPE_SERVER)
'                Call CloseSocket(UserIndex)
'                Exit Sub
'        End Select
'    End If
'    If LenB(UserList(UserIndex).flags.GSec_ID) = 0 Then
'        Call CloseSocket(UserIndex)
'        Exit Sub
'    End If


Public Function gsInformar() As String
    ' Informa del cheat detectado al servidor!
    gsInformar = "GAC" & gsCheatName() & "~" & gsCheatPath()
End Function

Public Function gsEnviarID() As String
    ' Le envia el GSEC_ID al servidor...
    Dim GSEC_ID As String * 32
    GSEC_ID = gsGetGSEC_ID()
    gsEnviarID = "GID" & GSEC_ID
End Function
