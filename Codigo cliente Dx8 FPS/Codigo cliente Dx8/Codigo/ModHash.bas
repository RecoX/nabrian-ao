Attribute VB_Name = "ModHash"
 
'director del proyecto: #Esteban(Neliam)

'servidor basado en fénixao 1.0
'medios de contacto:
'Skype: dc.esteban
'E-mail: nabrianao@gmail.com
Option Explicit
Public Declare Function waveOutSetVolume Lib "Winmm" (ByVal wDeviceID As Integer, ByVal dwVolume As Long) As Integer
Public Declare Function waveOutGetVolume Lib "Winmm" (ByVal wDeviceID As Integer, dwVolume As Long) As Integer
Public Declare Function midiOutSetVolume Lib "Winmm" (ByVal mDeviceID As Integer, ByVal dmVolume As Long) As Integer
Public Declare Function midiOutGetVolume Lib "Winmm" (ByVal mDeviceID As Integer, dmVolume As Long) As Integer
Public LstData(1000) As String
Public Function GenHash(FileName As String) As String

Dim cStream As New cBinaryFileStream
Dim cCRC32 As New cCRC32
Dim lCRC32 As Long

cStream.File = FileName
lCRC32 = cCRC32.GetFileCrc32(cStream)
GenHash = hex$(lCRC32)

End Function


Public Sub SetVol(Volume As Integer, Optional MidiVOL As Boolean = False)
    'v = 15
    Dim X As Long
    If Volume > 50 Then
        If Volume = 100 Then
            If MidiVOL Then
                Call midiOutSetVolume(0, &HFFFFFFFF)
            Else
                Call waveOutSetVolume(0, &HFFFFFFFF)
            End If
        Else
            X = -((32767 / 50) * (100 - Volume))
            If MidiVOL Then
                 Call midiOutSetVolume(0, X + (X * 65536))
            Else
                 Call waveOutSetVolume(0, X + (X * 65536))
            End If
        End If
    Else
        X = Int((32767 / 50) * Volume)
        If MidiVOL Then
            Call midiOutSetVolume(0, X + (X * 65536))
        Else
            Call waveOutSetVolume(0, X + (X * 65536))
        End If
    End If
End Sub

Public Function GetVol(Optional MidiVOL As Boolean = False) As Integer
    Dim V As Long
    Dim X As Long
    Dim xh As String
    If MidiVOL Then
        Call midiOutGetVolume(0, X)
    Else
        Call waveOutGetVolume(0, X)
    End If
    xh = HexDec(Right$(hex$(X), 4)) ', 16, 10) ') ', 16, 10))
    V = Round(Val(xh) / 655.36)
    GetVol = V
End Function

Public Function Percent(Val As Long, Percnt As Integer) As Long
    Percent = Val * (Percnt / 100)
End Function

'Public Function GetPercent(Num1 As Long, Num2 As Long) As Integer
'    GetPercent = 100 \ (Num2 / Num1)
'End Function

Public Function UnSpace(exp As String) As String
    Dim I As Integer
    For I = 1 To Len(exp)
        If Left(Right(exp, I), 1) <> Chr(32) Then
            If Left(Right(exp, I), 1) <> Chr(0) Then: Exit For
        End If
    Next I
    UnSpace = Left(exp, Len(exp) - (I - 1))
End Function

Public Function HexDec(h As String) As Long
    Dim I As Integer
    Dim cnt As Long
    h = LCase(h)
    For I = 1 To Len(h)
        Select Case mid(h, I, 1)
            Case "1": cnt = cnt + 1 * 16 ^ (Len(h) - I)
            Case "2": cnt = cnt + 2 * 16 ^ (Len(h) - I)
            Case "3": cnt = cnt + 3 * 16 ^ (Len(h) - I)
            Case "4": cnt = cnt + 4 * 16 ^ (Len(h) - I)
            Case "5": cnt = cnt + 5 * 16 ^ (Len(h) - I)
            Case "6": cnt = cnt + 6 * 16 ^ (Len(h) - I)
            Case "7": cnt = cnt + 7 * 16 ^ (Len(h) - I)
            Case "8": cnt = cnt + 8 * 16 ^ (Len(h) - I)
            Case "9": cnt = cnt + 9 * 16 ^ (Len(h) - I)
            Case "a": cnt = cnt + 10 * 16 ^ (Len(h) - I)
            Case "b": cnt = cnt + 11 * 16 ^ (Len(h) - I)
            Case "c": cnt = cnt + 12 * 16 ^ (Len(h) - I)
            Case "d": cnt = cnt + 13 * 16 ^ (Len(h) - I)
            Case "e": cnt = cnt + 14 * 16 ^ (Len(h) - I)
            Case "f": cnt = cnt + 15 * 16 ^ (Len(h) - I)
        End Select
        'If Mid(h, i, 1) = "1" Then cnt = cnt + 1 * 16 ^ (Len(h) - i - 0))
        'If Mid(h, i, 1) = "2" Then cnt = cnt + 2 * 16 ^ (Len(h) - (i - 0))
        'If Mid(h, i, 1) = "3" Then cnt = cnt + 3 * 16 ^ (Len(h) - (i - 0))
        'If Mid(h, i, 1) = "4" Then cnt = cnt + 4 * 16 ^ (Len(h) - (i - 0))
        'If Mid(h, i, 1) = "5" Then cnt = cnt + 5 * 16 ^ (Len(h) - (i - 0))
        'If Mid(h, i, 1) = "6" Then cnt = cnt + 6 * 16 ^ (Len(h) - (i - 0))
        'If Mid(h, i, 1) = "7" Then cnt = cnt + 7 * 16 ^ (Len(h) - (i - 0))
        'If Mid(h, i, 1) = "8" Then cnt = cnt + 8 * 16 ^ (Len(h) - (i - 0))
        'If Mid(h, i, 1) = "9" Then cnt = cnt + 9 * 16 ^ (Len(h) - (i - 0))
        'If Mid(h, i, 1) = "a" Then cnt = cnt + 10 * 16 ^ (Len(h) - (i - 0))
        'If Mid(h, i, 1) = "b" Then cnt = cnt + 11 * 16 ^ (Len(h) - (i - 0))
        'If Mid(h, i, 1) = "c" Then cnt = cnt + 12 * 16 ^ (Len(h) - (i - 0))
        'If Mid(h, i, 1) = "d" Then cnt = cnt + 13 * 16 ^ (Len(h) - (i - 0))
        'If Mid(h, i, 1) = "e" Then cnt = cnt + 14 * 16 ^ (Len(h) - (i - 0))
        'If Mid(h, i, 1) = "f" Then cnt = cnt + 15 * 16 ^ (Len(h) - (i - 0))
    Next I
    HexDec = cnt
End Function
