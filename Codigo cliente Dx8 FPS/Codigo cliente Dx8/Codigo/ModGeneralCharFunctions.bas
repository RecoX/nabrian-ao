Attribute VB_Name = "ModGeneralCharFunctions"
'FenixAO DirectX8
'Engine By ·Parra
'Arreglado By Thusing
'Algunas cosas tomadas del cliente de DarkTester


'Parra: Este modulo contiene funciones generales relacionadas con el Char o con el movimiento del mismo _
                que antes estaban en el TileEngine y en General pero que prefiero pasarlas aqui para tener el TileEngine más limpio
                
Option Explicit

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal arma As Integer, ByVal escudo As Integer, ByVal casco As Integer)
On Error Resume Next


If CharIndex > LastChar Then LastChar = CharIndex

If arma = 0 Then arma = 2
If escudo = 0 Then escudo = 2
If casco = 0 Then casco = 2

CharList(CharIndex).Head = HeadData(Head)

CharList(CharIndex).Body = BodyData(Body)

If Body > 83 And Body < 88 Then
    CharList(CharIndex).Navegando = 1
Else: CharList(CharIndex).Navegando = 0
End If

CharList(CharIndex).arma = WeaponAnimData(arma)
    
CharList(CharIndex).escudo = ShieldAnimData(escudo)
CharList(CharIndex).casco = CascoAnimData(casco)

CharList(CharIndex).Heading = Heading


CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.X = 0
CharList(CharIndex).MoveOffset.Y = 0


CharList(CharIndex).POS.X = X
CharList(CharIndex).POS.Y = Y


CharList(CharIndex).active = 1


MapData(X, Y).CharIndex = CharIndex

End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

CharList(CharIndex).active = 0
CharList(CharIndex).Criminal = 0
CharList(CharIndex).Privilegios = 0
CharList(CharIndex).FX = 0
CharList(CharIndex).FxLoopTimes = 0
CharList(CharIndex).invisible = False
CharList(CharIndex).Moving = 0
CharList(CharIndex).muerto = False
CharList(CharIndex).Nombre = ""
CharList(CharIndex).NombreNPC = ""
CharList(CharIndex).pie = False
CharList(CharIndex).POS.X = 0
CharList(CharIndex).POS.Y = 0
CharList(CharIndex).UsandoArma = False

End Sub
Function NextOpenChar()
Dim loopc As Integer

loopc = 1

Do While CharList(loopc).active
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function
Sub EraseChar(ByVal CharIndex As Integer)
On Error Resume Next

CharList(CharIndex).active = 0


If CharIndex = LastChar Then
    Do Until CharList(LastChar).active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If


MapData(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y).CharIndex = 0

Call ResetCharInfo(CharIndex)

End Sub
Sub MoveCharByHead(CharIndex As Integer, nheading As Byte)

Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = CharList(CharIndex).POS.X
Y = CharList(CharIndex).POS.Y


Select Case nheading

    Case NORTH
        addY = -1

    Case EAST
        addX = 1

    Case SOUTH
        addY = 1
    
    Case WEST
        addX = -1
        
End Select

nX = X + addX
nY = Y + addY

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).POS.X = nX
CharList(CharIndex).POS.Y = nY
MapData(X, Y).CharIndex = 0

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nheading

If ActivadoFps = 0 Then
CharList(CharIndex).scrollDirectionX = addX
CharList(CharIndex).scrollDirectionY = addY
End If
If UserEstado <> 1 Then Call DoPasosFx(CharIndex)


End Sub
Public Sub DoFogataFx()
If FX = 0 Then
    If bFogata Then
        bFogata = HayFogata()
        If Not bFogata Then Audio.StopWave
    Else
        bFogata = HayFogata()
        If bFogata Then Audio.PlayWave "fuego.wav", 0, 0, Enabled
    End If
End If
End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean

Dim X As Integer, Y As Integer

For Y = UserPos.Y - MinYBorder + 1 To UserPos.Y + MinYBorder - 1
  For X = UserPos.X - MinXBorder + 1 To UserPos.X + MinXBorder - 1
            
            If MapData(X, Y).CharIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
  Next X
Next Y

EstaPCarea = False

End Function
Public Function TickON(Cual As Integer, Cont As Integer) As Boolean
Static TickCount(200) As Integer
If Cont = 999 Then Exit Function
TickCount(Cual) = TickCount(Cual) + 1
If TickCount(Cual) < Cont Then
    TickON = False
Else
    TickCount(Cual) = 0
    TickON = True
End If
End Function
Sub PasosEfecto(ByVal CharIndex As Integer, TipoPaso As Integer)
If FX = 0 Then
If TipoPaso = 1 Then

If CharList(CharIndex).pie Then 'arena
                        Call Audio.PlayWave(0, SND_PASOS5)
Else
                        Call Audio.PlayWave(0, SND_PASOS6)
End If

ElseIf TipoPaso = 2 Then

If CharList(CharIndex).pie Then 'pasto
                        Call Audio.PlayWave(0, SND_PASOS3)
Else
                        Call Audio.PlayWave(0, SND_PASOS4)
End If

ElseIf TipoPaso = 3 Then

If CharList(CharIndex).pie Then
                        Call Audio.PlayWave(0, SND_PASOS1)
Else
                        Call Audio.PlayWave(0, SND_PASOS2)
End If

End If
End If
End Sub
Sub DoPasosFx(ByVal CharIndex As Integer)
Static pie As Boolean

If FX = 0 Then
If CharList(CharIndex).Navegando = 0 Then
    If UserMontando And EstaPCarea(CharIndex) And CharIndex = UserCharIndex Then
        If TickON(0, 4) Then Call Audio.PlayWave(0, SND_MONTANDO)
    Else
        If CharList(CharIndex).Criminal = 1 Then Exit Sub
        If Not CharList(CharIndex).muerto And EstaPCarea(CharIndex) And CharList(CharIndex).Privilegios = 0 Then
            CharList(CharIndex).pie = Not CharList(CharIndex).pie
                
                If MapData(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y).Graphic(1).GrhIndex >= 6000 And MapData(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y).Graphic(1).GrhIndex <= 6559 Then
                Call PasosEfecto(CharIndex, 2)
                ElseIf MapData(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y).Graphic(2).GrhIndex >= 7283 And MapData(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y).Graphic(2).GrhIndex <= 7327 Then
                Call PasosEfecto(CharIndex, 2)
                ElseIf MapData(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y).Graphic(1).GrhIndex >= 7704 And MapData(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y).Graphic(1).GrhIndex <= 8000 Then
                Call PasosEfecto(CharIndex, 1)
                ElseIf MapData(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y).Graphic(2).GrhIndex >= 7331 And MapData(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y).Graphic(2).GrhIndex <= 7375 Then
                Call PasosEfecto(CharIndex, 1)
                Else
                Call PasosEfecto(CharIndex, 3)
            End If
        End If
    End If
Else: Call Audio.PlayWave(0, SND_NAVEGANDO)
End If
End If

End Sub
Sub MoveCharByPosAndHead(CharIndex As Integer, nX As Integer, nY As Integer, nheading As Byte)

On Error Resume Next

Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer



X = CharList(CharIndex).POS.X
Y = CharList(CharIndex).POS.Y

MapData(X, Y).CharIndex = 0

addX = nX - X
addY = nY - Y




MapData(nX, nY).CharIndex = CharIndex


CharList(CharIndex).POS.X = nX
CharList(CharIndex).POS.Y = nY

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

If ActivadoFps = 0 Then
CharList(CharIndex).scrollDirectionX = Sgn(addX)
CharList(CharIndex).scrollDirectionY = Sgn(addY)
End If

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nheading


End Sub
Sub MoveCharByPos(CharIndex As Integer, nX As Integer, nY As Integer)
On Error Resume Next

Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nheading As Byte

X = CharList(CharIndex).POS.X
Y = CharList(CharIndex).POS.Y

MapData(X, Y).CharIndex = 0

addX = nX - X
addY = nY - Y


If Sgn(addX) = 1 Then nheading = EAST
If Sgn(addX) = -1 Then nheading = WEST
If Sgn(addY) = -1 Then nheading = NORTH
If Sgn(addY) = 1 Then nheading = SOUTH

MapData(nX, nY).CharIndex = CharIndex

CharList(CharIndex).POS.X = nX
CharList(CharIndex).POS.Y = nY

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)
If ActivadoFps = 0 Then
CharList(CharIndex).scrollDirectionX = Sgn(addX)
CharList(CharIndex).scrollDirectionY = Sgn(addY)
End If
CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nheading

End Sub
Sub MoveCharByPosConHeading(CharIndex As Integer, nX As Integer, nY As Integer, nheading As Byte)
On Error Resume Next

If InMapBounds(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y) Then MapData(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y).CharIndex = 0

MapData(nX, nY).CharIndex = CharIndex

CharList(CharIndex).POS.X = nX
CharList(CharIndex).POS.Y = nY

CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.X = 0
CharList(CharIndex).MoveOffset.Y = 0

CharList(CharIndex).Heading = nheading

End Sub
Sub MoveMe(Direction As Byte)

If CONGELADO Then Exit Sub

If Cartel Then Cartel = False

If ProxLegalPos(Direction) And Not UserMeditar And Not UserParalizado Then
    If TiempoTranscurrido(LastPaso) >= IntervaloPaso Then
        Call SendData("M" & Direction)
        Call DibujarMiniMapa
        LastPaso = Timer
        If Not UserDescansar Then
            Call EliminarChars(Direction)
            Call MoveCharByHead(UserCharIndex, Direction)
            Call MoveScreen(Direction)
            Call DoFogataFx
        End If
    End If
ElseIf CharList(UserCharIndex).Heading <> Direction Then Call SendData("CHEA" & Direction)
End If

frmPrincipal.mapa.Caption = NombreDelMapaActual & " [" & UserMap & " - " & UserPos.X & " - " & UserPos.Y & "]"

End Sub
Function ProxLegalPos(Direction As Byte) As Boolean

Select Case Direction
    Case NORTH
        ProxLegalPos = LegalPos(UserPos.X, UserPos.Y - 1)
    Case SOUTH
        ProxLegalPos = LegalPos(UserPos.X, UserPos.Y + 1)
    Case WEST
        ProxLegalPos = LegalPos(UserPos.X - 1, UserPos.Y)
    Case EAST
        ProxLegalPos = LegalPos(UserPos.X + 1, UserPos.Y)
End Select

End Function

Sub MoveScreen(Heading As Byte)

Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer

Select Case Heading

    Case NORTH
        Y = -1

    Case EAST
        X = 1

    Case SOUTH
        Y = 1
    
    Case WEST
        X = -1
        
End Select


tX = UserPos.X + X
tY = UserPos.Y + Y


If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
    Exit Sub
Else
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
End If

End Sub
Function HayFogata() As Boolean
Dim j As Integer, k As Integer
For j = UserPos.X - 8 To UserPos.X + 8
    For k = UserPos.Y - 6 To UserPos.Y + 6
        If InMapBounds(j, k) Then
            If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    HayFogata = True
                    Exit Function
            End If
        End If
    Next k
Next j
End Function
Sub RefreshAllChars()
Dim loopc As Integer

For loopc = 1 To LastChar
    If CharList(loopc).active = 1 Then
        'MapData(CharList(loopc).POS.X, CharList(loopc).POS.Y).CharIndex = loopc
    End If
Next loopc

End Sub
Function LegalPos(X As Integer, Y As Integer) As Boolean

    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        LegalPos = False
        Exit Function
    End If

    If MapData(X, Y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
 '¿Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        LegalPos = False
        Exit Function
    End If
   
    If Not UserNavegando Then
        If HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    End If
    
LegalPos = True

End Function
Function InMapBounds(X As Integer, Y As Integer) As Boolean
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        InMapBounds = False
        Exit Function
    End If

    InMapBounds = True
End Function
Function HayAgua(X As Integer, Y As Integer) As Boolean

If MapData(X, Y).Graphic(1).GrhIndex >= 1505 And _
   MapData(X, Y).Graphic(1).GrhIndex <= 1520 And _
   MapData(X, Y).Graphic(2).GrhIndex = 0 Then
            HayAgua = True
Else
            HayAgua = False
End If

End Function
Sub EliminarChars(Direction As Byte)
Dim X(2) As Integer
Dim Y(2) As Integer

Select Case Direction
    Case NORTH, SOUTH
        X(1) = UserPos.X - MinXBorder - 2
        X(2) = UserPos.X + MinXBorder + 2
    Case EAST, WEST
        Y(1) = UserPos.Y - MinYBorder - 2
        Y(2) = UserPos.Y + MinYBorder + 2
End Select

Select Case Direction
    Case NORTH
        Y(1) = UserPos.Y - MinYBorder - 3
        If Y(1) < 1 Then Y(1) = 1
        Y(2) = Y(1)
    Case EAST
        X(1) = UserPos.X + MinXBorder + 3
        If X(1) > 99 Then X(1) = 99
        X(2) = X(1)
    Case SOUTH
        Y(1) = UserPos.Y + MinYBorder + 3
        If Y(1) > 99 Then Y(1) = 99
        Y(2) = Y(1)
    Case WEST
        X(1) = UserPos.X - MinXBorder - 3
        If X(1) < 1 Then X(1) = 1
        X(2) = X(1)
End Select

For Y(0) = Y(1) To Y(2)
    For X(0) = X(1) To X(2)
        If X(0) > 6 And X(0) < 95 And Y(0) > 6 And Y(0) < 95 Then
            If MapData(X(0), Y(0)).CharIndex > 0 Then
                CharList(MapData(X(0), Y(0)).CharIndex).POS.X = 0
                CharList(MapData(X(0), Y(0)).CharIndex).POS.Y = 0
                MapData(X(0), Y(0)).CharIndex = 0
            End If
        End If
    Next
Next

End Sub



