Attribute VB_Name = "MoD_MIDI"
'NabrianAO (www.nabrianao.net)
'director del proyecto: #Esteban(Neliam)

'servidor basado en fénixao 1.0
'medios de contacto:
'Skype: dc.esteban
'E-mail: nabrianao@gmail.com
Option Explicit

Public Const MIdi_Inicio = 6

Public CurMidi As String
Public LoopMidi As Byte
Public IsPlayingCheck As Boolean

Public GetStartTime As Long
Public Offset As Long
Public mtTime As Long
Public mtLength As Double
Public dTempo As Double


Dim timesig As DMUS_TIMESIGNATURE
Dim portcaps As DMUS_PORTCAPS

Dim msg As String
Dim time As Double
Dim Offset2 As Long
Dim ElapsedTime2 As Double
Dim fIsPaused As Boolean








