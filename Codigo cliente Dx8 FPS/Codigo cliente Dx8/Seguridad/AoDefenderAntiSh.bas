Attribute VB_Name = "AoDefenderAntiSh"
 
'director del proyecto: #Esteban(Neliam)

'servidor basado en fénixao 1.0
'medios de contacto:
'Skype: dc.esteban
'E-mail: nabrianao@gmail.com
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function IsDebuggerPresent Lib "kernel32" () As Long

Public AoDefTime As Long
Public AoDefCount As Integer

Public AoDefOriginalClientName As String
Public AoDefClientName As String

Public Sub AoDefClientOn()
MsgBox "Se ha detectado cambio de nombre en el ejecutable. No es posible ejecutar el cliente! ponle 'NabrianAO.exe'.", vbCritical, "Nabrian Security"
End Sub



Public Sub AoDefAntiShInitialize()
AoDefTime = GetTickCount()
End Sub
Public Function AoDefAntiSh(ByVal FramesPerSec) As Boolean
If GetTickCount - AoDefTime > 335 Or GetTickCount - AoDefTime < 235 Then
        AoDefCount = AoDefCount + 1
    Else
        AoDefCount = 0
    End If
    
    If FramesPerSec < 5 Then
    AoDefCount = AoDefCount + 1
    End If
    
    If AoDefCount > 30 Then
       AoDefAntiSh = True
       Exit Function
    End If

AoDefTime = GetTickCount()
AoDefAntiSh = False
End Function
Public Sub AoDefAntiShOn()
MsgBox "Se ha detectado uso de SpeedHack, el cliente será cerrado!.", vbCritical, "Nabrian Security"
End Sub


Public Function AoDefChangeName() As Boolean
If AoDefOriginalClientName <> AoDefClientName Then
AoDefChangeName = True
Exit Function
End If
AoDefChangeName = False
End Function

Public Function AoDefDebugger() As Boolean
If IsDebuggerPresent Then
AoDefDebugger = True
Exit Function
End If
AoDefDebugger = False
End Function
Public Sub AoDefAntiDebugger()
MsgBox "Se ha detectado un intento de Debuggear el cliente, su cliente será cerrado.!", vbCritical, "Nabrian Security"
End Sub

