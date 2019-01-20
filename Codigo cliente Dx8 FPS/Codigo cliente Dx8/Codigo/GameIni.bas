Attribute VB_Name = "GameIni"
 
'director del proyecto: #Esteban(Neliam)

'servidor basado en fénixao 1.0
'medios de contacto:
'Skype: dc.esteban
'E-mail: nabrianao@gmail.com
Option Explicit

Public Type tCabecera
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tGameIni
    Puerto As Long
    Musica As Byte
    FX As Byte
    tip As Byte
    Password As String
    name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer
End Type


    Public Type tRenderMods
        sName      As String * 7
        bUseVideo  As Long
        bNoCostas  As Long
        bUseMMX    As Long
        bNoAlpha   As Long
        bNoTScan   As Long
        bNoMusic   As Long
        bNoSound   As Long
        iImageSize As Long
    End Type
    
    Public Type AutoUpdate
        version As Long
        Fase As Byte
    End Type
    
    Public RenderMod As tRenderMods


Public MiCabecera As tCabecera
Public Config_Inicio As tGameIni
Public AUpdate As AutoUpdate
Public Sub EscribirUpdate(ByRef Update As AutoUpdate)
Dim N As Integer
N = FreeFile
Open App.Path & "\init\AutoUpdate.con" For Binary As #N

Put #N, , Update
Close #N
End Sub
Public Function LeerAutoUpdate() As AutoUpdate
Dim N As Integer
Dim Up As AutoUpdate
N = FreeFile
Open App.Path & "\init\AutoUpdate.con" For Binary As #N

Get #N, , Up

Close #N
LeerAutoUpdate = Up
End Function
Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
Cabecera.Desc = "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
Cabecera.CRC = Rnd * 100
Cabecera.MagicWord = Rnd * 10
End Sub

Public Function LeerGameIni() As tGameIni
Dim N As Integer
Dim GameIni As tGameIni
N = FreeFile
Open App.Path & "\init\Inicio.con" For Binary As #N
Get #N, , MiCabecera

Get #N, , GameIni

Close #N
LeerGameIni = GameIni
End Function

Public Sub EscribirGameIni(ByRef GameIniConfiguration As tGameIni)
Dim N As Integer
N = FreeFile
Open App.Path & "\init\Inicio.con" For Binary As #N
Put #N, , MiCabecera
Put #N, , GameIniConfiguration
Close #N
End Sub

