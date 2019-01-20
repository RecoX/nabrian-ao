Attribute VB_Name = "Mod_DX"
'NabrianAO (www.nabrianao.net)
'director del proyecto: #Esteban(Neliam)

'servidor basado en fénixao 1.0
'medios de contacto:
'Skype: dc.esteban
'E-mail: nabrianao@gmail.com
Option Explicit

Public Const NumSoundBuffers = 20

Public DirectX As New DirectX7
Public DirectDraw As DirectDraw7
Public DirectSound As DirectSound

Public PrimarySurface As DirectDrawSurface7
Public PrimaryClipper As DirectDrawClipper
Public SecundaryClipper As DirectDrawClipper
Public BackBufferSurface As DirectDrawSurface7

Public SurfaceDB As New CBmpMan
Public Perf As DirectMusicPerformance
Public Seg As DirectMusicSegment
Public SegState As DirectMusicSegmentState
Public Loader As DirectMusicLoader

Public oldResHeight As Long, oldResWidth As Long
Public bNoResChange As Boolean

Public LastSoundBufferUsed As Integer
Public DSBuffers(1 To NumSoundBuffers) As DirectSoundBuffer

Public ddsd2 As DDSURFACEDESC2
Public ddsd4 As DDSURFACEDESC2
Public ddsd5 As DDSURFACEDESC2
Public ddsAlphaPicture As DirectDrawSurface7
Public ddsSpotLight As DirectDrawSurface7


Private Sub IniciarDirectSound()
Err.Clear
On Error GoTo fin
    Set DirectSound = DirectX.DirectSoundCreate("")
    If Err Then
        MsgBox "Error iniciando DirectSound"
        End
    End If
    
    LastSoundBufferUsed = 1
    
    Set Perf = DirectX.DirectMusicPerformanceCreate()
    Call Perf.Init(Nothing, 0)
    Perf.SetPort -1, 80
    Call Perf.SetMasterAutoDownload(True)
    
    Exit Sub
fin:

LogError "Error al iniciar IniciarDirectSound, asegurese de tener bien configurada la placa de sonido."

Musica = 1
FX = 1

End Sub

Private Sub LiberarDirectSound()
Dim cloop As Integer
For cloop = 1 To NumSoundBuffers
    Set DSBuffers(cloop) = Nothing
Next cloop
Set DirectSound = Nothing
End Sub

Private Sub IniciarDXobject(DX As DirectX7)

Err.Clear

On Error Resume Next

Set DX = New DirectX7

If Err Then
    MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
    LogError "Error producido por Set DX = New DirectX7"
    End
End If

End Sub

Private Sub IniciarDDobject(DD As DirectDraw7)
Err.Clear
On Error Resume Next
Set DD = DirectX.DirectDrawCreate("")
If Err Then
    MsgBox "No se puede iniciar DirectDraw. Por favor asegurese de tener la ultima version correctamente instalada."
    LogError "Error producido en Private Sub IniciarDDobject(DD As DirectDraw7)"
    End
End If
End Sub

Public Sub IniciarObjetosDirectX()

On Error Resume Next

    
            
If MsgBox("¿Desea Reproducir el Juego en Pantalla Completa?", vbQuestion + vbYesNo, "Resolución") = vbYes Then
        NoRes = 0
        Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "ModoVentana", 0)
    Else
        NoRes = 1
        Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "ModoVentana", 1)
End If

Call AddtoRichTextBox(frmCargando.Status, "Iniciando DirectX....", 255, 255, 255, 0, 0, True)
Call IniciarDXobject(DirectX)
Call AddtoRichTextBox(frmCargando.Status, "¡OK!", 11, 213, 105, 1, 0, False)

Call AddtoRichTextBox(frmCargando.Status, "Iniciando DirectDraw....", 255, 255, 255, 0, 0, True)
Call IniciarDDobject(DirectDraw)
Call AddtoRichTextBox(frmCargando.Status, "¡OK!", 11, 213, 105, 1, 0, False)

If Musica = 0 Or FX = 0 Then
    Call AddtoRichTextBox(frmCargando.Status, "Iniciando DirectSound....", 255, 255, 255, 0, 0, True)
    Call IniciarDirectSound
    Call AddtoRichTextBox(frmCargando.Status, "¡OK!", 11, 213, 105, 1, 0, False)
End If

Call AddtoRichTextBox(frmCargando.Status, "Analizando y preparando la placa de video....", 255, 255, 255, 0, 0, True)

  
    
Dim lRes As Long
Dim MidevM As typDevMODE
lRes = EnumDisplaySettings(0, 0, MidevM)
    
Dim intWidth As Integer
Dim intHeight As Integer


oldResWidth = Screen.Width \ Screen.TwipsPerPixelX
oldResHeight = Screen.Height \ Screen.TwipsPerPixelY

Dim CambiarResolucion As Boolean

If NoRes Then
    CambiarResolucion = (oldResWidth < 800 Or oldResHeight < 600)
Else
    CambiarResolucion = (oldResWidth <> 800 Or oldResHeight <> 600)
End If

If CambiarResolucion Then
      With MidevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = 800
            .dmPelsHeight = 600
            .dmBitsPerPel = 16
      End With
      lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
Else
      bNoResChange = True
End If

Call AddtoRichTextBox(frmCargando.Status, "¡TODO OK!", 11, 213, 105, 1, 0, False)

Exit Sub

End Sub

Public Sub LiberarObjetosDX()
Err.Clear
On Error GoTo fin:
Dim loopc As Integer

Set PrimarySurface = Nothing
Set PrimaryClipper = Nothing
Set BackBufferSurface = Nothing

LiberarDirectSound

Call SurfaceDB.BorrarTodo

Set DirectDraw = Nothing

For loopc = 1 To NumSoundBuffers
    Set DSBuffers(loopc) = Nothing
Next loopc


Set Loader = Nothing
Set Perf = Nothing
Set Seg = Nothing
Set DirectSound = Nothing

Set DirectX = Nothing
Exit Sub
fin: LogError "Error producido en Public Sub LiberarObjetosDX()"
End Sub

