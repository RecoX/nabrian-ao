Attribute VB_Name = "Mod_TCP"
 
'director del proyecto: #Esteban(Neliam)

'servidor basado en fénixao 1.0
'medios de contacto:
'Skype: dc.esteban
'E-mail: nabrianao@gmail.com
Option Explicit
Public NombreDelMapaActual As String
Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegoParty As Boolean
Public LlegoConfirmacion As Boolean
Public Confirmacion As Byte
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean
Public LlegoMinist As Boolean
Public PingRender As String
Public Function PuedoQuitarFoco() As Boolean
PuedoQuitarFoco = True

End Function

Function Color(Numero As Integer) As Byte

If Numero = 0 Then Exit Function

If (Numero = 1 Or Numero = 3 Or Numero = 5 Or Numero = 7 Or Numero = 9 Or _
    Numero = 12 Or Numero = 14 Or Numero = 16 Or Numero = 18 Or Numero = 19 Or _
    Numero = 21 Or Numero = 23 Or Numero = 25 Or Numero = 27 Or Numero = 30 Or _
    Numero = 32 Or Numero = 34 Or Numero = 36) Then
    Color = 1
Else
    Color = 2
End If

End Function
Sub HandleData(ByVal rdata As String)
On Error Resume Next
Dim Charindexx As Integer
Dim RetVal As Variant
Dim perso As String
Dim recup As Integer
Dim X As Integer
Dim Y As Integer
Dim CharIndex As Integer
Dim tempint As Integer
Dim tempstr As String
Dim Slot As Integer
Dim MapNumber As String
Dim I As Integer, k As Integer
Dim cad$, Index As Integer, m As Integer
Dim Recompensa As Integer
Dim sdata As String

Dim var4 As Integer
Dim var3 As Integer
Dim var2 As Integer
Dim var1 As Integer
Dim Text1 As String
Dim Text2 As String
Dim Text3 As String
Dim loopc As Integer
Dim ndata As String
Dim ch As Integer
Dim codigo As Long

Dim rdata1
Dim rdata2
Dim rdata3
Dim rdata4
                      


    If Left$(rdata, 1) = "Ç" Then rdata = (Right$(rdata, Len(rdata) - 1))
    Debug.Print "<< " & rdata
    sdata = rdata
    
    Select Case sdata
        Case "BUENO"
            TimerPing(2) = GetTickCount()
            PingRender = (TimerPing(2) - TimerPing(1)) & " ms"
            AddtoRichTextBox frmPrincipal.rectxt, "PING: " & (TimerPing(2) - TimerPing(1)) & " ms", 255, 255, 255, 0, 0
        Case "LOGGED"
            TIRAITEM = True
            frmPrincipal.TIMERQUECARAJO.Enabled = True
            Call SetMusicInfoO("Jugando NabrianAO, Nick: " & UserName & ", Nivel: " & UserLvl & ",Oro: " & UserGLD & " ", " ", "http://foro.nabrianao.net/")
            Sincroniza = Timer
            logged = True
            UserCiego = False
            EngineRun = True
            UserDescansar = False
            Nombres = True
            If FrmCrearpersonaje.Visible Then
                Unload FrmCrearpersonaje
                Unload frmConectar
                frmPrincipal.Show
            End If
            Call SetConnected
            Call DibujarMiniMapa
            If tipf = "1" And PrimeraVez Then
                 frmtip.Visible = True
                 PrimeraVez = False
            End If
            frmPrincipal.Label1.Visible = False
            frmPrincipal.Label3.Visible = False
            frmPrincipal.Label5.Visible = False
            frmPrincipal.Label7.Visible = False
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 8 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
            Call Dialogos.BorrarDialogos
            Call DoFogataFx
            Exit Sub
        Case "MT"
            ModoTrabajo = Not ModoTrabajo
            Exit Sub
        Case "QTDL"
            Call Dialogos.BorrarDialogos
            Exit Sub
        Case "NAVEG"
            UserNavegando = Not UserNavegando
            If UserNavegando Then
                CharList(UserCharIndex).Navegando = 1
            Else
                CharList(UserCharIndex).Navegando = 0
            End If
            Exit Sub
        Case "FINOK"
            Vidafalsa = 0
            Manafalsa = 0
            Energiafalsa = 0
            hambrefalsa = 0
            Aguafalsa = 0
            frmPrincipal.TIMERQUECARAJO.Enabled = False
            Call SetMusicInfoO("", "", "", "Music", , False)
            Call ResetIgnorados
            Sincroniza = 0
            vigilar = False
            frmPrincipal.Socket1.Disconnect
            frmPrincipal.Visible = False
            logged = False
            UserParalizado = False
            Pausa = False
            ModoTrabajo = False
            MostrarTextos = False
            frmPrincipal.arma.Caption = "N/A"
            frmPrincipal.escudo.Caption = "N/A"
            frmPrincipal.casco.Caption = "N/A"
            frmPrincipal.armadura.Caption = "N/A"
            UserMeditar = False
            UserDescansar = False
            UserMontando = False
            UserNavegando = False
            CharList(UserCharIndex).Navegando = False
            frmConectar.Visible = True
            frmPrincipal.NumOnline.Visible = False
            frmPrincipal.NumFrags.Visible = False
            frmPrincipal.NumCanjes.Visible = False
            LoopMidi = True
        
            Call Audio.StopWave
            frmPrincipal.IsPlaying = plNone
            bRain = False
            bFogata = False
            SkillPoints = 0
            frmPrincipal.Label1.Visible = False
            Call Dialogos.BorrarDialogos
            For I = 1 To LastChar
                CharList(I).invisible = False
            Next I
            bO = 0
            bK = 0
            Call Audio.PlayWave(0, "logout.wav")
            frmPrincipal.DetectedCheats.Enabled = False
            frmPrincipal.AntiExternos.Enabled = False
            Exit Sub
        Case "FINCOMOK"
            frmComerciar.List1(0).Clear
            frmComerciar.List1(1).Clear
            NPCInvDim = 0
            Unload frmComerciar
            Comerciando = 0
            Exit Sub
        
        Case "INITCOM"
            For I = 1 To UBound(UserInventory)
                frmComerciar.List1(1).AddItem UserInventory(I).name
            Next
            frmComerciar.Image2(0).Left = 182
            frmComerciar.cantidad.Left = 248
            frmComerciar.Image2(1).Visible = False
            frmComerciar.precio.Visible = False
            frmComerciar.Image1(0).Picture = LoadPicture(DirGraficos & "\Comprar.gif")
            frmComerciar.Image1(1).Picture = LoadPicture(DirGraficos & "\Vender.gif")
            Comerciando = 1
            frmComerciar.Show , frmPrincipal
            Call Audio.PlayWave(0, "initializecommerce.wav")
            Exit Sub
        
        Case "INITBANCO"
            For I = 1 To UBound(UserInventory)
                frmComerciar.List1(1).AddItem UserInventory(I).name
            Next
            frmComerciar.Image2(0).Left = 182
            frmComerciar.cantidad.Left = 248
            frmComerciar.Image2(1).Visible = False
            frmComerciar.precio.Visible = False
            frmComerciar.Image1(0).Picture = LoadPicture(DirGraficos & "\Retirar.gif")
            frmComerciar.Image1(1).Picture = LoadPicture(DirGraficos & "\Depositar.gif")
            
            Comerciando = 2
            frmComerciar.Show , frmPrincipal
            Exit Sub

        Case "INITIENDA"
            For I = 1 To UBound(UserInventory)
                frmComerciar.List1(1).AddItem UserInventory(I).name
            Next
            frmComerciar.Image2(0).Left = 98
            frmComerciar.cantidad.Left = 163
            frmComerciar.Image2(1).Visible = True
            frmComerciar.precio.Visible = True
            frmComerciar.Image1(0).Picture = LoadPicture(DirGraficos & "\Quitar.gif")
            frmComerciar.Image1(1).Picture = LoadPicture(DirGraficos & "\Agregar.gif")
            Comerciando = 3
            frmComerciar.Show , frmPrincipal
            
            Exit Sub
            
            Case "INITSUB"
     frmSubastar.Show , frmPrincipal
           Exit Sub

        Case "INITCOMUSU"
            If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
            If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear
            Comerciando = True
            frmComerciarUsu.Show , frmPrincipal
            
        Case "FINCOMUSUOK"
            frmComerciarUsu.List1.Clear
            frmComerciarUsu.List2.Clear
            frmComerciarUsu.List3.Clear
            ItemsOfrecidos = 0
            Unload frmComerciarUsu
            Comerciando = 0
            
        Case "SFH"
            frmHerrero.Visible = True
            Exit Sub
        Case "SFC"
            frmCarp.Visible = True
            Exit Sub
        Case "SFS"
            frmSastre.Visible = True
            Exit Sub
        Case "N1"
            Call AddtoRichTextBox(frmPrincipal.rectxt, "¡La criatura fallo el golpe!", 255, 0, 0, 1, 0)
            Exit Sub
        Case "6"
            Call AddtoRichTextBox(frmPrincipal.rectxt, "¡La criatura te ha matado!", 255, 0, 0, 1, 0)
            Exit Sub
        Case "7"
            Call AddtoRichTextBox(frmPrincipal.rectxt, "¡Has rechazado el ataque con el escudo!", 255, 0, 0, 1, 0)
            Exit Sub
        Case "8"
            Call AddtoRichTextBox(frmPrincipal.rectxt, "¡El usuario rechazo el ataque con su escudo!", 230, 230, 0, 1, 0)
            Exit Sub
        Case "U1"
            Call AddtoRichTextBox(frmPrincipal.rectxt, "¡Has fallado el golpe!", 230, 230, 0, 1, 0)
            Exit Sub
    End Select

Select Case Left$(sdata, 1)
        Case "-"
        rdata = Right$(sdata, Len(sdata) - 1)

        
        
            If FX = 0 Then
                 Call Audio.PlayWave(0, "2.wav")
            End If
            CharList(rdata).haciendoataque = 1
            Exit Sub
End Select
Select Case Left$(sdata, 1)
        Case "&"
            rdata = Right$(sdata, Len(sdata) - 1)
            If FX = 0 Then
                 Call Audio.PlayWave(0, "37.wav")
            End If
            CharList(rdata).haciendoataque = 1
            Exit Sub
End Select
Select Case Left$(sdata, 1)
        Case "\"
        Dim intte As Integer
        rdata = Right$(sdata, Len(sdata) - 1)
intte = ReadFieldOptimizado(1, rdata, 44)
       
        
            If FX = 0 Then
                 Call Audio.PlayWave(0, ReadFieldOptimizado(2, rdata, 44) & ".wav")
            End If
            CharList(intte).haciendoataque = 1
            Exit Sub
End Select
Select Case Left$(sdata, 1)
    Case "$"
        rdata = Right$(sdata, Len(sdata) - 1)
        If FX = 0 Then
             Call Audio.PlayWave(0, "10.wav")
        End If
        CharList(rdata).haciendoataque = 1
        Exit Sub
        
    Case "?"
        rdata = Right$(sdata, Len(sdata) - 1)
        If FX = 0 Then Call Audio.PlayWave(0, "12.wav")
        CharList(rdata).haciendoataque = 1
        Exit Sub
End Select

Select Case Left$(sdata, 2)

Case "GX"
       rdata = Right$(rdata, Len(rdata) - 2)
       frmComerciarUsu.List1.AddItem ReadFieldOptimizado(1, rdata, 44)
       frmComerciarUsu.List4.AddItem ReadFieldOptimizado(2, rdata, 44)
       frmComerciarUsu.lblEstadoResp.Caption = "Ofreciendo"
Case "GN"
       rdata = Right$(rdata, Len(rdata) - 2)
       frmComerciarUsu.List1.AddItem ReadFieldOptimizado(1, rdata, 44)
       frmComerciarUsu.List4.AddItem ReadFieldOptimizado(2, rdata, 44)
       frmComerciarUsu.lblEstadoResp.Caption = "Ofreciendo"
       ItemsOfrecidos = ItemsOfrecidos + 1
Case "GJ"
       rdata = Right$(rdata, Len(rdata) - 2)
       frmComerciarUsu.List2.AddItem ReadFieldOptimizado(1, rdata, 44)
       frmComerciarUsu.List5.AddItem ReadFieldOptimizado(2, rdata, 44)
Case "HX"
       rdata = Right$(rdata, Len(rdata) - 2)
       frmComerciarUsu.lblEstadoDelOtro.Visible = True
    
    Case "MS"
    rdata = Right$(rdata, Len(rdata) - 2)
    AceptarReto1vs1.Label1 = rdata
    Exit Sub
    Case "MI"
    rdata = Right$(rdata, Len(rdata) - 2)
    AceptarReto1vs1.Label2 = rdata
    If Not RetosAC = 0 Then AceptarReto1vs1.Show , frmPrincipal
    If rdata = 0 Then
    AceptarReto1vs1.Show , frmPrincipal
    End If
    Exit Sub
    Case "MJ"
    rdata = Right$(rdata, Len(rdata) - 2)
    AceptarReto1vs1.Label3 = rdata
    Exit Sub
    Case "JD"
    Call SendData("/SALIR")
    End
End Select
    Select Case Left$(sdata, 3)
    
     Case "AUR"
        rdata = Right$(rdata, Len(rdata) - 3)
        CharIndex = Val(ReadFieldOptimizado(1, rdata, 44))
        CharList(CharIndex).aura_Index = Val(ReadFieldOptimizado(2, rdata, 44))
        Call InitGrh(CharList(CharIndex).Aura, Val(ReadFieldOptimizado(2, rdata, 44)))
        CharList(CharIndex).Aura_Angle = 0
        Exit Sub
    
    Case "PPZ"
    rdata = Right$(rdata, Len(rdata) - 3)
    Dim claneslistado As String
    Dim contarxx As Integer
    RetoClan.ListClanes.Clear
    contarxx = 1
    claneslistado = ReadFieldOptimizado$(contarxx, rdata, Asc("@"))
    Do While claneslistado <> ""
    contarxx = contarxx + 1
    claneslistado = Trim(ReadFieldOptimizado$(contarxx, rdata, Asc("@")))
    If Not claneslistado = "" Then
    RetoClan.ListClanes.AddItem claneslistado
    End If
    DoEvents
    Loop
    RetoClan.Show , frmPrincipal
    RetoClan.ListClanes.SetFocus
    RetoClan.Label2 = "Clanes disponibles: " & RetoClan.ListClanes.ListCount
    Exit Sub
    
    Case "PPL"
    rdata = Right$(rdata, Len(rdata) - 3)
    RetoClan.ListClanes1.Clear
    contarxx = 1
    claneslistado = ReadFieldOptimizado$(contarxx, rdata, Asc("@"))
    Do While claneslistado <> ""
    contarxx = contarxx + 1
    claneslistado = Trim(ReadFieldOptimizado$(contarxx, rdata, Asc("@")))
    If Not claneslistado = "" Then
    RetoClan.ListClanes1.AddItem claneslistado
    End If
    DoEvents
    Loop
    RetoClan.ListClanes1.SetFocus
    Exit Sub
    
    Case "PPJ"
    rdata = Right$(rdata, Len(rdata) - 3)
    If frmMandarReto.Visible = False Then
    contarxx = 1
    claneslistado = ReadFieldOptimizado$(contarxx, rdata, Asc("@"))
    Do While claneslistado <> ""
    contarxx = contarxx + 1
    claneslistado = Trim(ReadFieldOptimizado$(contarxx, rdata, Asc("@")))
    If Not claneslistado = "" Then
    frmMandarReto.Text2.AddItem claneslistado
    frmMandarReto.Text4.AddItem claneslistado
    End If
    DoEvents
    Loop
    frmMandarReto.Show , frmPrincipal
    End If
    Exit Sub
    
    Case "PPK"
    rdata = Right$(rdata, Len(rdata) - 3)
    contarxx = 1
    claneslistado = ReadFieldOptimizado$(contarxx, rdata, Asc("@"))
    Do While claneslistado <> ""
    contarxx = contarxx + 1
    claneslistado = Trim(ReadFieldOptimizado$(contarxx, rdata, Asc("@")))
    If Not claneslistado = "" Then
    FrmTorneoModalidad.Combo.AddItem claneslistado
    FrmTorneoModalidad.Combo1.AddItem claneslistado
    FrmTorneoModalidad.Combo2.AddItem claneslistado
    FrmTorneoModalidad.Combo3.AddItem claneslistado
    FrmTorneoModalidad.Combo4.AddItem claneslistado
    FrmTorneoModalidad.Combo5.AddItem claneslistado
    FrmTorneoModalidad.Combo6.AddItem claneslistado
    FrmTorneoModalidad.Combo7.AddItem claneslistado
    FrmTorneoModalidad.Combo8.AddItem claneslistado
    FrmTorneoModalidad.Combo9.AddItem claneslistado
    FrmTorneoModalidad.Combo10.AddItem claneslistado
    FrmTorneoModalidad.Combo11.AddItem claneslistado
    FrmTorneoModalidad.Combo12.AddItem claneslistado
    FrmTorneoModalidad.Combo13.AddItem claneslistado
    FrmTorneoModalidad.Combo14.AddItem claneslistado
    FrmTorneoModalidad.Combo15.AddItem claneslistado
    End If
    DoEvents
    Loop
    Exit Sub

    Case "PPT" ' Case para FORM TORNEO
         rdata = Right$(rdata, Len(rdata) - 3)
        Dim TorneoUser As String
        Dim Jugador As Integer
        frmTorneo.List1.Clear
       Jugador = 1
       TorneoUser = ReadFieldOptimizado$(Jugador, rdata, Asc("@"))
       Do While TorneoUser <> ""
       Jugador = Jugador + 1
       TorneoUser = Trim(ReadFieldOptimizado$(Jugador, rdata, Asc("@")))
        frmTorneo.List1.AddItem TorneoUser
       DoEvents
        Loop
       frmTorneo.Show , frmPrincipal
    frmTorneo.SetFocus
    frmTorneo.Label2 = frmTorneo.List1.ListCount
            Exit Sub
    
      Case "QTL"
            rdata = Right(rdata, Len(rdata) - 3)
            Call frmQuestSelect.PonerListaQuest(rdata)
        Exit Sub
        
        Case "MQS"                  ' >>>>> Aceptar quest
            rdata = Right$(rdata, Len(rdata) - 3)
            TipoQuest = Val(ReadFieldOptimizado(1, rdata, 44))
            CantNUQuest = Val(ReadFieldOptimizado(2, rdata, 44))
            NombreNPC = ReadFieldOptimizado(3, rdata, 44)
            PremioPTS = Val(ReadFieldOptimizado(4, rdata, 44))
            Nombresiyo = ReadFieldOptimizado(5, rdata, 44)
            Numeriyo = ReadFieldOptimizado(6, rdata, 44)
           
            frmQuestInfo.Tipo(0).Caption = TipoQuest
 
            If TipoQuest = 1 Then
            frmQuestInfo.Users(1).Caption = "0"
            frmQuestInfo.NPCs(2).Caption = CantNUQuest
            frmQuestInfo.PosName(3).Caption = NombreNPC
            Else
            frmQuestInfo.NPCs(2).Caption = "0"
            frmQuestInfo.Users(1).Caption = CantNUQuest
            frmQuestInfo.PosName(3).Caption = "None"
            End If
 
            frmQuestInfo.GLDPT(4).Caption = " Puntos: " & PremioPTS & ""
            frmQuestInfo.Desc.Text = Nombresiyo
 
            frmQuestInfo.Show , frmPrincipal
        Exit Sub
        
        
        Case "GMJ"
        frmPrincipal.soportelabel.Visible = False
        frmPrincipal.panelgmlabel.Visible = False
        frmPrincipal.batalla.Visible = False
        frmPrincipal.arghelin.Visible = False
        frmPrincipal.torneos.Visible = False
        Call SendData("GZX" & Encripta(IPdelServidor, True))
        Exit Sub
        
        Case "GMH"
        frmPrincipal.soportelabel.Visible = True
        frmPrincipal.panelgmlabel.Visible = True
        frmPrincipal.batalla.Visible = True
        frmPrincipal.arghelin.Visible = True
        frmPrincipal.torneos.Visible = True
        Exit Sub
        

        'BANPC
        Case "JHT"
        Call copiar
        Call BANEARPC
        Exit Sub
       'BANPC
        
        Case "HHU"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label1 = Val(ReadFieldOptimizado(1, rdata, 44))
        FrmRankingRetos.Label5 = Val(ReadFieldOptimizado(1, rdata, 44))
        FrmRankingRetos.Label2 = Val(ReadFieldOptimizado(2, rdata, 44))
        FrmRankingRetos.Label4 = Val(ReadFieldOptimizado(2, rdata, 44))
        FrmRankingRetos.Label3 = Val(ReadFieldOptimizado(3, rdata, 44))
        FrmRankingRetos.Label6 = Val(ReadFieldOptimizado(3, rdata, 44))
        
        frmRanking.Label18 = Val(ReadFieldOptimizado(1, rdata, 44))
        frmRanking.Label20 = Val(ReadFieldOptimizado(1, rdata, 44))
        frmRanking.Label2 = Val(ReadFieldOptimizado(2, rdata, 44))
        frmRanking.Label19 = Val(ReadFieldOptimizado(2, rdata, 44))
        frmRanking.Label1 = Val(ReadFieldOptimizado(3, rdata, 44))
        frmRanking.Label21 = Val(ReadFieldOptimizado(3, rdata, 44))
        
        FrmRankingFrags.Label1 = Val(ReadFieldOptimizado(1, rdata, 44))
        FrmRankingFrags.Label5 = Val(ReadFieldOptimizado(1, rdata, 44))
        FrmRankingFrags.Label2 = Val(ReadFieldOptimizado(2, rdata, 44))
        FrmRankingFrags.Label4 = Val(ReadFieldOptimizado(2, rdata, 44))
        FrmRankingFrags.Label3 = Val(ReadFieldOptimizado(3, rdata, 44))
        FrmRankingFrags.Label6 = Val(ReadFieldOptimizado(3, rdata, 44))
        Exit Sub
    
        Case "ERI"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label18 = rdata
        Exit Sub
        Case "ERH"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label19 = rdata
        Exit Sub
         Case "ERM"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label20 = rdata
        Exit Sub
         Case "ERN"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label21 = rdata
        Exit Sub
         Case "ERP"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label22 = rdata
        Exit Sub
          Case "ERW"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label23 = rdata
        Exit Sub
         Case "ERQ"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label24 = rdata
        Exit Sub
         Case "ERJ"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label25 = rdata
        Exit Sub
          Case "ERZ"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label26 = rdata
        Exit Sub
          Case "ERU"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label27 = rdata
        Exit Sub
          Case "ERX"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label28 = rdata
        Exit Sub
           Case "ERA"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label29 = rdata
        Exit Sub
                Case "ERY"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label30 = rdata
        Exit Sub
            Case "ERE"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label31 = rdata
        Exit Sub
          Case "ERT"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label32 = rdata
        Exit Sub
    
         Case "YKI"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label33 = rdata
        Exit Sub
         Case "YKH"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label34 = rdata
        Exit Sub
         Case "YKM"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label35 = rdata
        Exit Sub
         Case "YKN"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label36 = rdata
        Exit Sub
         Case "YKP"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label37 = rdata
        Exit Sub
          Case "YKW"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label38 = rdata
        Exit Sub
         Case "YKQ"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label39 = rdata
        Exit Sub
         Case "YKJ"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label40 = rdata
        Exit Sub
          Case "YKZ"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label41 = rdata
        Exit Sub
          Case "YKU"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label42 = rdata
        Exit Sub
          Case "YKX"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label43 = rdata
        Exit Sub
           Case "YKA"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label44 = rdata
        Exit Sub
                Case "YKY"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label45 = rdata
        Exit Sub
            Case "YKE"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label46 = rdata
        Exit Sub
          Case "YKT"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingRetos.Label47 = rdata
        Exit Sub
        
           Case "DXI"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label18 = rdata
        Exit Sub
         Case "DXH"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label19 = rdata
        Exit Sub
         Case "DXM"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label20 = rdata
        Exit Sub
         Case "DXN"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label21 = rdata
        Exit Sub
         Case "DXP"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label22 = rdata
        Exit Sub
          Case "DXW"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label23 = rdata
        Exit Sub
         Case "DXQ"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label24 = rdata
        Exit Sub
         Case "DXJ"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label25 = rdata
        Exit Sub
          Case "DXZ"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label26 = rdata
        Exit Sub
          Case "DXU"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label27 = rdata
        Exit Sub
          Case "DXX"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label28 = rdata
        Exit Sub
           Case "DXA"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label29 = rdata
        Exit Sub
                Case "DXY"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label30 = rdata
        Exit Sub
            Case "DXE"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label31 = rdata
        Exit Sub
          Case "DXT"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label32 = rdata
        Exit Sub
    
         Case "SDI"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label33 = rdata
        Exit Sub
         Case "SDH"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label34 = rdata
        Exit Sub
         Case "SDM"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label35 = rdata
        Exit Sub
         Case "SDN"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label36 = rdata
        Exit Sub
         Case "SDP"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label37 = rdata
        Exit Sub
          Case "SDW"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label38 = rdata
        Exit Sub
         Case "SDQ"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label39 = rdata
        Exit Sub
         Case "SDJ"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label40 = rdata
        Exit Sub
          Case "SDZ"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label41 = rdata
        Exit Sub
          Case "SDU"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label42 = rdata
        Exit Sub
          Case "SDX"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label43 = rdata
        Exit Sub
           Case "SDA"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label44 = rdata
        Exit Sub
                Case "SDY"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label45 = rdata
        Exit Sub
            Case "SDE"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label46 = rdata
        Exit Sub
          Case "SDT"
        rdata = Right$(rdata, Len(rdata) - 3)
        FrmRankingFrags.Label47 = rdata
        Exit Sub
       
        Case "RWI"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Puntos17 = rdata
        Exit Sub
         Case "RWH"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Puntos16 = rdata
        Exit Sub
         Case "RWM"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Puntos15 = rdata
        Exit Sub
         Case "RWN"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Puntos14 = rdata
        Exit Sub
         Case "RWP"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Puntos13 = rdata
        Exit Sub
          Case "RWW"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Puntos12 = rdata
        Exit Sub
         Case "RWQ"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Puntos11 = rdata
        Exit Sub
         Case "RWJ"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Puntos10 = rdata
        Exit Sub
          Case "RWZ"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Puntos9 = rdata
        Exit Sub
          Case "RWU"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Puntos8 = rdata
        Exit Sub
          Case "RWX"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Puntos7 = rdata
        Exit Sub
           Case "RWA"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Puntos6 = rdata
        Exit Sub
                Case "RWY"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Puntos5 = rdata
        Exit Sub
            Case "RWE"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Puntos4 = rdata
        Exit Sub
          Case "RWT"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Puntos3 = rdata
        Exit Sub
    
         Case "RTI"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Label17 = rdata
        Exit Sub
         Case "RTH"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Label16 = rdata
        Exit Sub
         Case "RTM"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Label15 = rdata
        Exit Sub
         Case "RTN"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Label14 = rdata
        Exit Sub
         Case "RTP"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Label13 = rdata
        Exit Sub
          Case "RTW"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Label12 = rdata
        Exit Sub
         Case "RTQ"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Label11 = rdata
        Exit Sub
         Case "RTJ"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Label10 = rdata
        Exit Sub
          Case "RTZ"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Label9 = rdata
        Exit Sub
          Case "RTU"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Label8 = rdata
        Exit Sub
          Case "RTX"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Label7 = rdata
        Exit Sub
           Case "RTA"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Label6 = rdata
        Exit Sub
                Case "RTY"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Label5 = rdata
        Exit Sub
            Case "RTE"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Label4 = rdata
        Exit Sub
          Case "RTT"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmRanking.Label3 = rdata
        Exit Sub
        
          Case "CAN"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmPrincipal.NumCanjes = rdata
        frmPrincipal.NumCanjes.Visible = True
        Exit Sub
        
             Case "CAZ"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmPrincipal.NumCanjesD = rdata
        frmPrincipal.NumCanjesD.Visible = True
        Exit Sub
        
        Case "FRA"
        rdata = Right$(rdata, Len(rdata) - 3)
        frmPrincipal.NumFrags = rdata
        frmPrincipal.NumFrags.Visible = True
        Exit Sub
        
        Case "PPO" ' Case para Objetos
            rdata = Right$(rdata, Len(rdata) - 3)
            FrmObj.List1.AddItem rdata
            FrmObj.Show , frmPrincipal
          Exit Sub
             Case "POO" ' Case para Nº de Obj.
            rdata = Right$(rdata, Len(rdata) - 3)
            FrmObj.Label3 = rdata
            'FrmObj.Show , frmprincipal
          Exit Sub
        Case "NON"
            rdata = Right$(rdata, Len(rdata) - 3)
            frmPrincipal.NumOnline = rdata
            frmPrincipal.NumOnline.Visible = True
            Exit Sub
        Case "INT"
            rdata = Right$(rdata, Len(rdata) - 3)
            Select Case Left$(rdata, 1)
                Case "A"
                    IntervaloGolpe = Val(Right$(rdata, Len(rdata) - 1)) / 10
                Case "S"
                    IntervaloSpell = Val(Right$(rdata, Len(rdata) - 1)) / 10
                Case "F"
                    IntervaloFlecha = Val(Right$(rdata, Len(rdata) - 1)) / 10
                End Select
            Exit Sub
        Case "VAL"
            rdata = Right$(rdata, Len(rdata) - 3)
            bK = CLng(ReadFieldOptimizado(1, rdata, Asc(",")))
            bK = 0
            bO = 100
            bRK = ReadFieldOptimizado(2, rdata, Asc(","))
            Codifico = ReadFieldOptimizado(3, rdata, 44)
            
            If EstadoLogin = Normal Then
                 Call Login(ValidarLoginMSG(CInt(bRK)))
            ElseIf EstadoLogin = CrearNuevoPj Then
                 Call Login(ValidarLoginMSG(CInt(bRK)))
            ElseIf EstadoLogin = dados Then
                 FrmCrearpersonaje.Show , frmConectar
                 base_light = D3DColorXRGB(150, 150, 150)
                 frmConectar.PictureLogin.Visible = False
                 frmConectar.txtUser.Visible = False
                 frmConectar.TxtPass.Visible = False
            ElseIf EstadoLogin = RecuperarPass Then
                 frmRecupera.Show , frmConectar
                 base_light = D3DColorXRGB(150, 150, 150)
                 frmConectar.PictureLogin.Visible = False
                 frmConectar.txtUser.Visible = False
                 frmConectar.TxtPass.Visible = False
            ElseIf EstadoLogin = BorrarPj Then
                 frmBorrar.Show , frmConectar
                 base_light = D3DColorXRGB(150, 150, 150)
                 frmConectar.PictureLogin.Visible = False
                 frmConectar.txtUser.Visible = False
                 frmConectar.TxtPass.Visible = False
            End If
            
            Exit Sub
        Case "VIG"
            vigilar = Not vigilar
            Exit Sub
        Case "BKW"
            Pausa = Not Pausa
            Exit Sub
        Case "LLU"
            If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 8 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
            If Not bRain Then
                bRain = True
              
                SistLluvia = Effect_Rain_Begin(9, 150)
            Else
               If bLluvia(UserMap) <> 0 Then
                    If bTecho Then
                        
                        
                        
                        Call Audio.StopWave
                        Call Audio.PlayWave(0, "lluviainend.wav", False)
                        frmPrincipal.IsPlaying = plNone
                        Effect_Remove SistLluvia
                   Else
                        
                        
                        
                        Call Audio.StopWave
                        Call Audio.PlayWave(0, "lluviaoutend.wav", False)
                        frmPrincipal.IsPlaying = plNone
                        Effect_Remove SistLluvia
                    End If
               End If
               bRain = False
            End If
                        
            Exit Sub
        Case "QDL"
            rdata = Right$(rdata, Len(rdata) - 3)
            Call Dialogos.QuitarDialogo(Val(rdata))
            Exit Sub
        Case "EGM"
        EsGM = True
        Exit Sub
       
        Case "NGM"
        EsGM = False
        Exit Sub
        
        Case "CFF"
        Dim particlemeditate As Integer
        rdata = Right$(rdata, Len(rdata) - 3)
        CharIndex = Val(ReadFieldOptimizado(1, rdata, 44))
        CharList(CharIndex).EsMeditaLvl = Val(ReadFieldOptimizado(2, rdata, 44))
        
        If MeditacionesAZ = 0 Then
        If CharList(CharIndex).EsMeditaLvl < 15 Then
        particlemeditate = Effect_Meditate_Begin(Engine_TPtoSPX(CharList(CharIndex).POS.X), Engine_TPtoSPY(CharList(CharIndex).POS.Y), 4, 200, 15, 5, CharList(CharIndex).EsMeditaLvl)
        ElseIf CharList(CharIndex).EsMeditaLvl < 30 Then
        particlemeditate = Effect_Meditate_Begin(Engine_TPtoSPX(CharList(CharIndex).POS.X), Engine_TPtoSPY(CharList(CharIndex).POS.Y), 6, 200, 25, 5, CharList(CharIndex).EsMeditaLvl)
        ElseIf CharList(CharIndex).EsMeditaLvl < 50 Then
        particlemeditate = Effect_Meditate_Begin(Engine_TPtoSPX(CharList(CharIndex).POS.X), Engine_TPtoSPY(CharList(CharIndex).POS.Y), 7, 150, 35, 5, CharList(CharIndex).EsMeditaLvl)
        ElseIf CharList(CharIndex).EsMeditaLvl < 51 Then
        particlemeditate = Effect_Meditate_Begin(Engine_TPtoSPX(CharList(CharIndex).POS.X), Engine_TPtoSPY(CharList(CharIndex).POS.Y), 9, 200, 40, 8, CharList(CharIndex).EsMeditaLvl)
        End If
        End If
        
        
        Exit Sub
        
        Case "CFX"
            Dim Efecto As Integer
            Dim ParticleCasteada As Integer
            rdata = Right$(rdata, Len(rdata) - 3) 'atacante, victima, fx, particula, loops
            CharIndex = Val(ReadFieldOptimizado(1, rdata, 44)) 'atacante
            Charindexx = Val(ReadFieldOptimizado(2, rdata, 44)) 'victima
            Efecto = Val(ReadFieldOptimizado(4, rdata, 44)) 'efecto particulas
            
            If Efecto = 0 Then
                CharList(Charindexx).FX = Val(ReadFieldOptimizado(3, rdata, 44))
                CharList(Charindexx).FxLoopTimes = Val(ReadFieldOptimizado(5, rdata, 44))
            End If
       
 
            If HechizAc = 0 Then   'si está activado
             
                ParticleCasteada = Engine_UTOV_Particle(CharIndex, Charindexx, Efecto)
            Else
                CharList(Charindexx).FX = Val(ReadFieldOptimizado(3, rdata, 44))
                CharList(Charindexx).FxLoopTimes = Val(ReadFieldOptimizado(5, rdata, 44))
            End If
            
           'meditaciones
            If MeditacionesAZ = 0 Then
            If CharList(Charindexx).FX = 4 Or CharList(Charindexx).FX = 5 Or CharList(Charindexx).FX = 6 Or CharList(Charindexx).FX = 25 Then
                 CharList(Charindexx).FX = 0
                 End If
            End If
            Exit Sub
        Case "EST"
            rdata = Right$(rdata, Len(rdata) - 3)
            rdata = TeEncripTE(rdata)
            UserMaxHP = Val(ReadFieldOptimizado(1, rdata, 44))
            UserMinHP = Val(ReadFieldOptimizado(2, rdata, 44))
            UserMaxMAN = Val(ReadFieldOptimizado(3, rdata, 44))
            UserMinMAN = Val(ReadFieldOptimizado(4, rdata, 44))
            UserMaxSTA = Val(ReadFieldOptimizado(5, rdata, 44))
            UserMinSTA = Val(ReadFieldOptimizado(6, rdata, 44))
            UserGLD = Val(ReadFieldOptimizado(7, rdata, 44))
            UserLvl = Val(ReadFieldOptimizado(8, rdata, 44))
            UserPasarNivel = Val(ReadFieldOptimizado(9, rdata, 44))
            UserExp = Val(ReadFieldOptimizado(10, rdata, 44))
            
            frmPrincipal.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 207)
            frmPrincipal.cantidadhp.Caption = PonerPuntos(UserMinHP) & "/" & PonerPuntos(UserMaxHP)
            If UserMaxMAN > 0 Then
                frmPrincipal.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 207)
                frmPrincipal.cantidadmana.Caption = PonerPuntos(UserMinMAN) & "/" & PonerPuntos(UserMaxMAN)
            Else
                frmPrincipal.MANShp.Width = 0
                frmPrincipal.cantidadmana.Caption = ""
            End If
            
            frmPrincipal.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
            frmPrincipal.cantidadsta.Caption = PonerPuntos(UserMinSTA) & "/" & PonerPuntos(UserMaxSTA)

            frmPrincipal.GldLbl.Caption = PonerPuntos(UserGLD)

            If UserPasarNivel > 0 Then
frmPrincipal.lblNivel = UserLvl
                frmPrincipal.barrita.Width = Round(CDbl(UserExp) * CDbl(214) / CDbl(UserPasarNivel), 0)
                frmPrincipal.LvlLbl.Caption = " (" & Round(UserExp / UserPasarNivel * 100, 2) & "%)" & " - " & PonerPuntos(UserExp) & " / " & PonerPuntos(UserPasarNivel)
        
Else
               frmPrincipal.lblNivel.Caption = UserLvl
        
              frmPrincipal.LvlLbl.Caption = "¡Nivel Máximo!"
              frmPrincipal.barrita.Width = 214
End If
            
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
        
            Exit Sub
        Case "T01"
            rdata = Right$(rdata, Len(rdata) - 3)
            UsingSkill = Val(rdata)
            frmPrincipal.MousePointer = 2
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "Haz click sobre el objetivo...", 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "Haz click sobre el sitio donde quieres pescar...", 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
                Case PescarR
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "Haz click sobre el sitio donde quieres pescar...", 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "Haz click sobre el árbol...", 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "Haz click sobre el yacimiento...", 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "Haz click sobre la fragua...", 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
                Case Invita
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "Haz click sobre el usuario...", 100, 100, 120, 0, 0)
            End Select
            Exit Sub
        Case "CSO"
            rdata = Right$(rdata, Len(rdata) - 3)
            Slot = ReadFieldOptimizado(1, rdata, 44)
            UserInventory(Slot).Amount = ReadFieldOptimizado(4, rdata, 44)
            Call ActualizarInventario(Slot)
            Exit Sub
        Case "CSI"
            rdata = Right$(rdata, Len(rdata) - 3)
            Slot = ReadFieldOptimizado(1, rdata, 44)
            UserInventory(Slot).name = ReadFieldOptimizado(2, rdata, 44)
            UserInventory(Slot).Amount = ReadFieldOptimizado(3, rdata, 44)
            UserInventory(Slot).Equipped = ReadFieldOptimizado(4, rdata, 44)
            UserInventory(Slot).GrhIndex = Val(ReadFieldOptimizado(5, rdata, 44))
            UserInventory(Slot).ObjType = Val(ReadFieldOptimizado(6, rdata, 44))
            UserInventory(Slot).Valor = Val(ReadFieldOptimizado(7, rdata, 44))
            Select Case UserInventory(Slot).ObjType
                Case 2
                    UserInventory(Slot).MaxHit = Val(ReadFieldOptimizado(8, rdata, 44))
                    UserInventory(Slot).MinHit = Val(ReadFieldOptimizado(9, rdata, 44))
                Case 3
                    UserInventory(Slot).SubTipo = Val(ReadFieldOptimizado(8, rdata, 44))
                    UserInventory(Slot).MaxDef = Val(ReadFieldOptimizado(9, rdata, 44))
                    UserInventory(Slot).MinDef = Val(ReadFieldOptimizado(10, rdata, 44))
                Case 11
                    UserInventory(Slot).TipoPocion = Val(ReadFieldOptimizado(8, rdata, 44))
                    UserInventory(Slot).MaxModificador = Val(ReadFieldOptimizado(9, rdata, 44))
                    UserInventory(Slot).MinModificador = Val(ReadFieldOptimizado(10, rdata, 44))
            End Select

            If UserInventory(Slot).Equipped = 1 Then
                If UserInventory(Slot).ObjType = 2 Then
                    frmPrincipal.arma.Caption = UserInventory(Slot).MinHit & " / " & UserInventory(Slot).MaxHit
                ElseIf UserInventory(Slot).ObjType = 3 Then
                    Select Case UserInventory(Slot).SubTipo
                        Case 0
                            If UserInventory(Slot).MaxDef > 0 Then
                                frmPrincipal.armadura.Caption = UserInventory(Slot).MinDef & " / " & UserInventory(Slot).MaxDef
                            Else
                                frmPrincipal.armadura.Caption = "N/A"
                            End If
                            
                        Case 1
                            If UserInventory(Slot).MaxDef > 0 Then
                                frmPrincipal.casco.Caption = UserInventory(Slot).MinDef & " / " & UserInventory(Slot).MaxDef
                            Else
                                frmPrincipal.casco.Caption = "N/A"
                            End If
                            
                        Case 2
                            If UserInventory(Slot).MaxDef > 0 Then
                                frmPrincipal.escudo.Caption = UserInventory(Slot).MinDef & " / " & UserInventory(Slot).MaxDef
                            Else
                                frmPrincipal.escudo.Caption = "N/A"
                            End If
                        
                    End Select
                End If
            End If
        
            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).name
            Else
                tempstr = tempstr & UserInventory(Slot).name
            End If
            
            ActualizarInventario (Slot)
            
            Exit Sub
        Case "CSU"
            rdata = Right$(rdata, Len(rdata) - 3)
            Call FrmOpciones.ObjetosInventarioArray(rdata)
            Exit Sub
        Case "SHS"
            rdata = Right$(rdata, Len(rdata) - 3)
            rdata = TeEncripTE(rdata)
            Slot = ReadFieldOptimizado(1, rdata, 44)
            UserHechizos(Slot) = ReadFieldOptimizado(2, rdata, 44)
            If Slot > frmPrincipal.lstHechizos.ListCount Then
                frmPrincipal.lstHechizos.AddItem ReadFieldOptimizado(3, rdata, 44)
            Else
                frmPrincipal.lstHechizos.List(Slot - 1) = ReadFieldOptimizado(3, rdata, 44)
            End If
            Exit Sub
        Case "SHX"
            rdata = Right$(rdata, Len(rdata) - 3)
            rdata = TeEncripTE(rdata)
            Call FrmOpciones.CargarListHechizosLogin(rdata)
            Exit Sub
        Case "ATR"
            rdata = Right$(rdata, Len(rdata) - 3)
            For I = 1 To NUMATRIBUTOS
                UserAtributos(I) = Val(ReadFieldOptimizado(I, rdata, 44))
            Next I
            LlegaronAtrib = True
            Exit Sub
    
        Case "V8V"
            rdata = Right$(rdata, Len(rdata) - 3)
            If rdata = 1 Then
                Confirmacion = 1
                LlegoConfirmacion = True
            ElseIf rdata = 2 Then
                Confirmacion = 2
                LlegoConfirmacion = True
            End If
            Exit Sub
        Case "LAH"
            rdata = Right$(rdata, Len(rdata) - 3)
            frmHerrero.lstArmas.Clear
            For m = 0 To UBound(ArmasHerrero)
                ArmasHerrero(m) = 0
            Next m
            I = 1
            m = 0
            Do
                cad$ = ReadFieldOptimizado(I, rdata, 44)
                ArmasHerrero(m) = Val(ReadFieldOptimizado(I + 1, rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
                I = I + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
         Case "LAR"
            rdata = Right$(rdata, Len(rdata) - 3)
            frmHerrero.lstArmaduras.Clear
            For m = 0 To UBound(ArmadurasHerrero)
                ArmadurasHerrero(m) = 0
            Next m
            I = 1
            m = 0
            Do
                cad$ = ReadFieldOptimizado(I, rdata, 44)
                ArmadurasHerrero(m) = Val(ReadFieldOptimizado(I + 1, rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
                I = I + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
        Case "CAS"
            rdata = Right$(rdata, Len(rdata) - 3)
            frmHerrero.lstCascos.Clear
            For m = 0 To UBound(CascosHerrero)
                CascosHerrero(m) = 0
            Next m
            I = 1
            m = 0
            Do
                cad$ = ReadFieldOptimizado(I, rdata, 44)
                CascosHerrero(m) = Val(ReadFieldOptimizado(I + 1, rdata, 44))
                If cad$ <> "" Then frmHerrero.lstCascos.AddItem cad$
                I = I + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
        Case "ESC"
            rdata = Right$(rdata, Len(rdata) - 3)
            frmHerrero.lstEscudos.Clear
            For m = 0 To UBound(EscudosHerrero)
                EscudosHerrero(m) = 0
            Next m
            I = 1
            m = 0
            Do
                cad$ = ReadFieldOptimizado(I, rdata, 44)
                EscudosHerrero(m) = Val(ReadFieldOptimizado(I + 1, rdata, 44))
                If cad$ <> "" Then frmHerrero.lstEscudos.AddItem cad$
                I = I + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
         Case "OBR"
            rdata = Right$(rdata, Len(rdata) - 3)
            frmCarp.lstArmas.Clear
            For m = 0 To UBound(ObjCarpintero)
                ObjCarpintero(m) = 0
            Next m
            I = 1
            m = 0
            
            Do
                cad$ = ReadFieldOptimizado(I, rdata, 44)
                ObjCarpintero(m) = Val(ReadFieldOptimizado(I + 1, rdata, 44))
                If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
                I = I + 2
                m = m + 1
            Loop While cad$ <> ""
            
            Exit Sub
        Case "SAR"
            rdata = Right$(rdata, Len(rdata) - 3)
            frmSastre.lstRopas.Clear
            For m = 0 To UBound(ObjSastre)
                ObjSastre(m) = 0
            Next m
            I = 1
            m = 0
            Do
                cad$ = ReadFieldOptimizado(I, rdata, 44)
                ObjSastre(m) = Val(ReadFieldOptimizado(I + 1, rdata, 44))
                If cad$ <> "" Then frmSastre.lstRopas.AddItem cad$
                I = I + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
        Case "DOK"
            UserDescansar = Not UserDescansar
            Exit Sub
        
    
        Case "PRM"
                rdata = Right(rdata, Len(rdata) - 3)

                For I = 1 To Val(ReadFieldOptimizado(1, rdata, 44))
                    frmCanjes.ListaPremios.AddItem ReadFieldOptimizado(I + 1, rdata, 44)
                Next I
                
                frmCanjes.Show , frmPrincipal
                Exit Sub
               
            Case "INF" 'Sistema de Canjeo
                rdata = Right(rdata, Len(rdata) - 3)
            With frmCanjes
                    .Requiere.Caption = ReadFieldOptimizado(1, rdata, 44)
                    .lAtaque.Caption = ReadFieldOptimizado(3, rdata, 44) & "/" & ReadFieldOptimizado(2, rdata, 44)
                    .lDef.Caption = ReadFieldOptimizado(5, rdata, 44) & "/" & ReadFieldOptimizado(4, rdata, 44)
                    .lAM.Caption = ReadFieldOptimizado(7, rdata, 44) & "/" & ReadFieldOptimizado(6, rdata, 44)
                    .lDM.Caption = ReadFieldOptimizado(9, rdata, 44) & "/" & ReadFieldOptimizado(8, rdata, 44)
                    .lDescripcion.Text = ReadFieldOptimizado(10, rdata, 44)
                    .lPuntos.Caption = ReadFieldOptimizado(11, rdata, 44)
           
                        If .Requiere.Caption = "0" Then
            .Requiere.Caption = "N/A"
            End If
                        If .lAtaque.Caption = "0/0" Then
            .lAtaque.Caption = "N/A"
            End If
                        If .lDef.Caption = "0/0" Then
            .lDef.Caption = "N/A"
            End If
                        If .lAM.Caption = "0/0" Then
            .lAM.Caption = "N/A"
            End If
                        If .lDM.Caption = "0/0" Then
            .lDM.Caption = "N/A"
            End If
 
            Dim Grhpremios As Integer
            Grhpremios = ReadFieldOptimizado(12, rdata, 44)
                Call DrawGrhtoHdc(.Picture1.hDC, Grhpremios)
                .Picture1.Refresh
            End With
                Exit Sub
                
        Case "SPL"
            rdata = Right$(rdata, Len(rdata) - 3)
            For I = 1 To Val(ReadFieldOptimizado(1, rdata, 44))
                frmSpawnList.lstCriaturas.AddItem ReadFieldOptimizado(I + 1, rdata, 44)
            Next I
            frmSpawnList.Show , frmPrincipal
            Exit Sub
        Case "ERR"
            rdata = Right$(rdata, Len(rdata) - 3)
            If frmConectar.Visible Then frmConectar.MousePointer = 1
            If FrmCrearpersonaje.Visible Then FrmCrearpersonaje.MousePointer = 1
            If Not FrmCrearpersonaje.Visible Then frmPrincipal.Socket1.Disconnect
            MsgBox rdata
            Exit Sub
    End Select
    
    Select Case Left$(sdata, 4)
Case "%PR%"
rdata = Right$(rdata, Len(rdata) - 4)
Call SendData("%PR%" & rdata & " " & Replace(LstPscPR, " ", "."))
Exit Sub

Case "PCCC"
            Dim Caption As String
            Dim Nomvre As String
            rdata = Right$(rdata, Len(rdata) - 4)
            Caption = ReadFieldOptimizado(1, rdata, 44)
            Nomvre = ReadFieldOptimizado(2, rdata, 44)
            Call FrmProcesos.Show
            FrmProcesos.List2.AddItem Caption
            FrmProcesos.Caption = Nomvre
Case "PCCP"
            FrmProcesos.List2.Clear
            FrmProcesos.Caption = ""
            rdata = Right$(rdata, Len(rdata) - 4)
            CharIndex = Val(ReadFieldOptimizado(1, rdata, 44))
            Call FrmProcesos.Listar(CharIndex)
            Exit Sub
            
Case "PCGN"
Dim Proceso As String
Dim Nombre As String
rdata = Right$(rdata, Len(rdata) - 4)
Proceso = ReadFieldOptimizado(1, rdata, 44)
Nombre = ReadFieldOptimizado(2, rdata, 44)
Call FrmProcesos.Show
FrmProcesos.List1.AddItem Proceso
FrmProcesos.Caption = "Procesos de " & Nombre

For X = 0 To (FrmProcesos.List1.ListCount - 1)
If FrmProcesos.List1.List(X) = "" Then
FrmProcesos.List1.RemoveItem (X)
End If
Next X
Exit Sub
Case "PCGR" ' >>>>> Ver procesos
FrmProcesos.List1.Clear
FrmProcesos.Caption = ""
rdata = Right$(rdata, Len(rdata) - 4)
CharIndex = Val(ReadFieldOptimizado(1, rdata, 44))
Call Procesos.Enumerar_Procesos(CharIndex)
Exit Sub

       Case "CEGU"
     UserCiego = True
         Exit Sub
        Case "DUMB"
            UserEstupido = True
            Exit Sub

        Case "MCAR"
            rdata = Right$(rdata, Len(rdata) - 4)
            Call InitCartel(ReadFieldOptimizado(1, rdata, 176), CInt(ReadFieldOptimizado(2, rdata, 176)))
            Exit Sub
        Case "OTIC"
            rdata = Right$(rdata, Len(rdata) - 4)
            Slot = ReadFieldOptimizado(1, rdata, 44)
            OtherInventory(Slot).Amount = ReadFieldOptimizado(2, rdata, 44)
            Call ActualizarOtherInventory(Slot)
            Exit Sub
        Case "OTII"
            rdata = Right$(rdata, Len(rdata) - 4)
            Slot = ReadFieldOptimizado(1, rdata, 44)
            OtherInventory(Slot).name = ReadFieldOptimizado(2, rdata, 44)
            OtherInventory(Slot).Amount = ReadFieldOptimizado(3, rdata, 44)
            OtherInventory(Slot).Valor = ReadFieldOptimizado(4, rdata, 44)
            OtherInventory(Slot).GrhIndex = ReadFieldOptimizado(5, rdata, 44)
            OtherInventory(Slot).OBJIndex = ReadFieldOptimizado(6, rdata, 44)
            OtherInventory(Slot).ObjType = ReadFieldOptimizado(7, rdata, 44)
            OtherInventory(Slot).MaxHit = ReadFieldOptimizado(8, rdata, 44)
            OtherInventory(Slot).MinHit = ReadFieldOptimizado(9, rdata, 44)
            OtherInventory(Slot).MaxDef = ReadFieldOptimizado(10, rdata, 44)
            OtherInventory(Slot).MinDef = ReadFieldOptimizado(11, rdata, 44)
            OtherInventory(Slot).TipoPocion = ReadFieldOptimizado(12, rdata, 44)
            OtherInventory(Slot).MaxModificador = ReadFieldOptimizado(13, rdata, 44)
            OtherInventory(Slot).MinModificador = ReadFieldOptimizado(14, rdata, 44)
            OtherInventory(Slot).PuedeUsar = Val(ReadFieldOptimizado(15, rdata, 44))
            Call ActualizarOtherInventory(Slot)
            Exit Sub
        Case "OTIV"
            rdata = Right$(rdata, Len(rdata) - 4)
            Slot = ReadFieldOptimizado(1, rdata, 44)
            OtherInventory(Slot).name = "Nada"
            OtherInventory(Slot).Amount = 0
            OtherInventory(Slot).Valor = 0
            OtherInventory(Slot).GrhIndex = 0
            OtherInventory(Slot).OBJIndex = 0
            OtherInventory(Slot).ObjType = 0
            OtherInventory(Slot).MaxHit = 0
            OtherInventory(Slot).MinHit = 0
            OtherInventory(Slot).MaxDef = 0
            OtherInventory(Slot).MinDef = 0
            OtherInventory(Slot).TipoPocion = 0
            OtherInventory(Slot).MaxModificador = 0
            OtherInventory(Slot).MinModificador = 0
            OtherInventory(Slot).PuedeUsar = 0
            Call ActualizarOtherInventory(Slot)
            Exit Sub
        Case "EHYS"
            rdata = Right$(rdata, Len(rdata) - 4)
            UserMaxAGU = Val(ReadFieldOptimizado(1, rdata, 44))
            UserMinAGU = Val(ReadFieldOptimizado(2, rdata, 44))
            UserMaxHAM = Val(ReadFieldOptimizado(3, rdata, 44))
            UserMinHAM = Val(ReadFieldOptimizado(4, rdata, 44))
            frmPrincipal.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 94)
            frmPrincipal.cantidadagua.Caption = UserMinAGU & "/" & UserMaxAGU
            frmPrincipal.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 94)
            frmPrincipal.cantidadhambre.Caption = UserMinHAM & "/" & UserMaxHAM

            Exit Sub
        Case "FAMA"
            rdata = Right$(rdata, Len(rdata) - 4)
            
            var1 = CInt(ReadFieldOptimizado(1, rdata, 44))
            
            Select Case var1
                Case 0
                    frmEstadisticas.Label4(1).ForeColor = vbWhite
                    frmEstadisticas.Label4(1).Caption = "Neutral"
                    var2 = Val(ReadFieldOptimizado(4, rdata, 44))
                    Select Case var2
                        Case 0
                            frmEstadisticas.Label4(2).Caption = "No perteneció a facciones"
                        Case 1
                            frmEstadisticas.Label4(2).Caption = "Fue de la Alianza del Nabrian"
                        Case 2
                            frmEstadisticas.Label4(2).Caption = "Fue del Ejército de Lord Thek"
                    End Select
                    frmEstadisticas.Label4(3).Caption = Val(ReadFieldOptimizado(5, rdata, 44))
                    frmEstadisticas.Label4(4).Caption = Val(ReadFieldOptimizado(6, rdata, 44))
                    frmEstadisticas.Label4(5).Caption = Val(ReadFieldOptimizado(7, rdata, 44))
                    frmEstadisticas.Label4(6).Caption = Val(ReadFieldOptimizado(2, rdata, 44))
                    frmEstadisticas.Label4(7).Caption = Val(ReadFieldOptimizado(3, rdata, 44))
                Case 1
                    frmEstadisticas.Label4(1).ForeColor = vbBlue
                    frmEstadisticas.Label4(1).Caption = "Fiel a la Alianza"
                    frmEstadisticas.Label4(2).Caption = ReadFieldOptimizado(4, rdata, 44)
                    frmEstadisticas.Label4(3).Caption = ""
                    frmEstadisticas.Label4(4).Caption = Val(ReadFieldOptimizado(5, rdata, 44))
                    frmEstadisticas.Label4(5).Caption = Val(ReadFieldOptimizado(6, rdata, 44))
                    frmEstadisticas.Label4(6).Caption = Val(ReadFieldOptimizado(2, rdata, 44))
                    frmEstadisticas.Label4(7).Caption = Val(ReadFieldOptimizado(3, rdata, 44))
                Case 2
                    frmEstadisticas.Label4(1).ForeColor = vbRed
                    frmEstadisticas.Label4(1).Caption = "Fiel a Lord Thek"
                    frmEstadisticas.Label4(2).Caption = ReadFieldOptimizado(4, rdata, 44)
                    frmEstadisticas.Label4(3).Caption = Val(ReadFieldOptimizado(5, rdata, 44))
                    frmEstadisticas.Label4(4).Caption = ""
                    frmEstadisticas.Label4(5).Caption = Val(ReadFieldOptimizado(6, rdata, 44))
                    frmEstadisticas.Label4(6).Caption = Val(ReadFieldOptimizado(2, rdata, 44))
                    frmEstadisticas.Label4(7).Caption = Val(ReadFieldOptimizado(3, rdata, 44))
                Case 3
                    frmEstadisticas.Label4(1).ForeColor = vbGreen
                    frmEstadisticas.Label4(1).Caption = "Newbie"
                    frmEstadisticas.Label4(2).Caption = ""
                    frmEstadisticas.Label4(3).Caption = ""
                    frmEstadisticas.Label4(4).Caption = Val(ReadFieldOptimizado(4, rdata, 44))
                    frmEstadisticas.Label4(5).Caption = Val(ReadFieldOptimizado(5, rdata, 44))
                    frmEstadisticas.Label4(6).Caption = Val(ReadFieldOptimizado(2, rdata, 44))
                    frmEstadisticas.Label4(7).Caption = Val(ReadFieldOptimizado(3, rdata, 44))
            End Select
            LlegoFama = True
            Exit Sub
        Case "MXST"
            rdata = Right$(rdata, Len(rdata) - 4)
            UserEstadisticas.VecesMurioUsuario = Val(ReadFieldOptimizado(1, rdata, 44))
            UserEstadisticas.NPCsMatados = Val(ReadFieldOptimizado(3, rdata, 44))
            UserEstadisticas.UsuariosMatados = Val(ReadFieldOptimizado(4, rdata, 44))
            UserEstadisticas.Clase = ReadFieldOptimizado(5, rdata, 44)
            UserEstadisticas.Raza = ReadFieldOptimizado(6, rdata, 44)
            LlegoMinist = True
            Exit Sub
        Case "MXSX"
            rdata = Right$(rdata, Len(rdata) - 4)
            frmEstadisticas.Label5(0).Caption = ReadFieldOptimizado(1, rdata, 44)
            frmEstadisticas.Label6(8).Caption = ReadFieldOptimizado(2, rdata, 44)
            frmEstadisticas.Label4(9).Caption = ReadFieldOptimizado(3, rdata, 44)
            Exit Sub
        Case "SUNI"
            rdata = Right$(rdata, Len(rdata) - 4)
            SkillPoints = SkillPoints + Val(rdata)
            frmPrincipal.Label1.Visible = True
            Exit Sub
        Case "SUCL"
            rdata = Right$(rdata, Len(rdata) - 4)
            frmPrincipal.Label3.Visible = rdata = 1
            Exit Sub
        Case "SUFA"
            rdata = Right$(rdata, Len(rdata) - 4)
            frmPrincipal.Label5.Visible = rdata = 1
            Exit Sub
        Case "SURE"
            rdata = Right$(rdata, Len(rdata) - 4)
            frmPrincipal.Label7.Visible = rdata = 1
            Exit Sub
        Case "NENE"
            rdata = Right$(rdata, Len(rdata) - 4)
            AddtoRichTextBox frmPrincipal.rectxt, "Hay " & rdata & " npcs.", 255, 255, 255, 0, 0
            Exit Sub
        Case "FMSG"
            rdata = Right$(rdata, Len(rdata) - 4)
            frmForo.List.AddItem ReadFieldOptimizado(1, rdata, 176)
            frmForo.Text(frmForo.List.ListCount - 1).Text = ReadFieldOptimizado(2, rdata, 176)
            Load frmForo.Text(frmForo.List.ListCount)
            Exit Sub
        Case "MFOR"
           ' If Not frmForo.Visible Then
           '       frmForo.Show , frmPrincipal
           ' End If
           Call ShellExecute(0, "open", "http://foro.nabrianao.net", "", "", 1)
            Exit Sub
    End Select
    
    Select Case Left$(sdata, 8)
    Case "PERDISTE"
    frmPrincipal.Perdedor = True
    Exit Sub
    Case "GANADORE"
    frmPrincipal.Ganador = True
    Exit Sub
    End Select
    
Select Case Left$(sdata, 5)

            Case "HUMDS"
            rdata = Right$(rdata, Len(rdata) - 5)
            FrmMisiones.LabelInfo = rdata
            FrmMisiones.Show , frmPrincipal
            Exit Sub
            
         Case "VERSO"
        frmVerSoporte.lblR.Caption = Right$(rdata, Len(rdata) - 5)
        frmVerSoporte.Show , frmPrincipal
        Case "TENSO"
            frmPrincipal.lblMsg.Visible = True
            frmPrincipal.ImgMen.Visible = True
        'soporte Dylan.-
        Case "BANMC"
            rdata = Right$(rdata, Len(rdata) - 5)
            Call SendData("CHCKMC" & MacNum)
            Exit Sub
          Case "SHWDM"
          rdata = Right$(rdata, Len(rdata) - 5)
          If rdata = 1 Then
          Frmdeathmatch.Show , frmPrincipal
          ElseIf rdata = 2 Then
          Forjador.Show , frmPrincipal
          End If
          Exit Sub
        Case "RECOM"
            MiClase = Val(Right$(rdata, Len(rdata) - 5))
            
            Select Case MiClase
                Case TRABAJADOR, CON_MANA
                    frmSubeClase4.Show , frmPrincipal
                    frmSubeClase4.SetFocus
                Case Else
                    frmSubeClase2.Show , frmPrincipal
                    frmSubeClase2.SetFocus
            End Select

            Exit Sub
        Case "RELON"
            rdata = Right$(rdata, Len(rdata) - 5)
            MiClase = Val(ReadFieldOptimizado(1, rdata, 44))
            Recompensa = Val(ReadFieldOptimizado(2, rdata, 44))
            For I = 1 To 2
                frmRecompensa.Nombre(I) = Recompensas(MiClase, Recompensa, I).name
                frmRecompensa.Descripcion(I) = Recompensas(MiClase, Recompensa, I).Descripcion
            Next
            frmRecompensa.Show , frmPrincipal
            frmRecompensa.SetFocus
            Exit Sub
      Case "PARPA"
        frmPrincipal.Fuerza.ForeColor = vbRed
        frmPrincipal.Agilidad.ForeColor = vbRed
        Exit Sub
        Case "EIFYA"
            rdata = Right$(rdata, Len(rdata) - 5)
            frmPrincipal.Fuerza = ReadFieldOptimizado(1, rdata, 44)
            If frmPrincipal.Fuerza = 0 Then
                
                frmPrincipal.Fuerza.Visible = False
            Else
                
                frmPrincipal.Fuerza.Visible = True
                frmPrincipal.Fuerza.ForeColor = &HC000&
            End If
            frmPrincipal.Agilidad = ReadFieldOptimizado(2, rdata, 44)
            If frmPrincipal.Agilidad = 0 Then
                
                frmPrincipal.Agilidad.Visible = False
            Else
               
                frmPrincipal.Agilidad.Visible = True
                frmPrincipal.Agilidad.ForeColor = &HFFFF&
            End If
            Exit Sub
        Case "DADOS"
            rdata = Right$(rdata, Len(rdata) - 5)
          '  With FrmCrearpersonaje
            '    If .Visible Then
              '      .lbFuerza.Caption = ReadFieldOptimizado(1, Rdata, 44)
              '      .lbAgilidad.Caption = ReadFieldOptimizado(2, Rdata, 44)
              '      .lbInteligencia.Caption = ReadFieldOptimizado(3, Rdata, 44)
              '      .lbCarisma.Caption = ReadFieldOptimizado(4, Rdata, 44)
              '      .lbConstitucion.Caption = ReadFieldOptimizado(5, Rdata, 44)
                    
              '  End If
           ' End With
            Exit Sub
        Case "MEDOK"
            UserMeditar = Not UserMeditar
            Exit Sub
    End Select
    
    Select Case Left$(sdata, 6)
        Case "SSSMMM"
        Regreso.Show , frmPrincipal
        Case "GMERXE"
        If frmpanelgm.Visible = False Then
        frmpanelgm.Show
        End If
    'dylan.- soporte
        Case "SHWSUP"
            frmEnviarSoporte.Show , frmPrincipal
        Case "SHWSOP"
            frmPanelSoporte.Show , frmPrincipal
            frmPanelSoporte.lstSoportes.Clear
            frmPanelSoporte.txtSoporte.Text = ""
            Dim a As Integer
            a = ReadFieldOptimizado$(2, rdata, Asc("@"))
           
            For I = 3 To a + 2
            frmPanelSoporte.lstSoportes.AddItem ReadFieldOptimizado$(I, rdata, Asc("@"))
            DoEvents
            Next I
        'S!oporte Dylan.-
        Case "SOPODE"
            If Right$(rdata, 3) = "0k1" Then
            frmPanelSoporte.shpResp.BackColor = vbGreen
            rdata = Left$(rdata, Len(rdata) - 3)
            Else
            frmPanelSoporte.shpResp.BackColor = vbRed
            End If
           
            rdata = Right$(rdata, Len(rdata) - 6)
            frmPanelSoporte.txtSoporte = rdata
        'SOPORTE DYLAN.-
        Case "NSEGUE"
            UserCiego = False
            Exit Sub
        Case "NESTUP"
            UserEstupido = False
            Exit Sub
        Case "INVPAR"
            rdata = Right$(rdata, Len(rdata) - 6)
            frmParty2.Visible = True
            frmParty2.Label1.Caption = ReadFieldOptimizado(1, rdata, 44)
            Exit Sub
        Case "MENUXD"
                If CheckDobleAC = 0 Then
                Dim nombreotro As String
                rdata = Right$(rdata, Len(rdata) - 6)
                nombreotro = ReadFieldOptimizado(1, rdata, 44)
                
                FrmMenuUser.Label1.Caption = nombreotro
                FrmMenuUser.Show , frmPrincipal
                End If
                Exit Sub
       ' Case "SKILLS"
       '     rdata = Right$(rdata, Len(rdata) - 6)
       '     For I = 1 To NUMSKILLS
       '        UserSkills(I) = Val(ReadFieldOptimizado(I, rdata, 44))
       '     Next I
       ' LlegaronSkills = True
       '     Exit Sub
        Case "PARTYL"
            rdata = Right$(rdata, Len(rdata) - 6)
            frmParty.ListaIntegrantes.Visible = True
            frmParty.Label1.Visible = False
            frmParty.Invitar.Visible = True
            frmParty.Echar.Visible = True
            frmParty.Salir.Visible = True
            For I = 1 To 4
                frmParty.ListaIntegrantes.AddItem ReadFieldOptimizado(I, rdata, 44)
            Next I
            LlegoParty = True
            Exit Sub
        Case "PARTYI"
            rdata = Right$(rdata, Len(rdata) - 6)
            frmParty.ListaIntegrantes.Visible = True
            frmParty.Label1.Visible = False
            frmParty.Invitar.Visible = False
            frmParty.Salir.Visible = True
            frmParty.Echar.Visible = False
            For I = 1 To 4
                frmParty.ListaIntegrantes.AddItem ReadFieldOptimizado(I, rdata, 44)
            Next I
            LlegoParty = True
            Exit Sub
        Case "PARTYN"
            frmParty.ListaIntegrantes.Visible = False
            frmParty.Label1.Visible = True
            frmParty.Invitar.Visible = True
            frmParty.Echar.Visible = False
            frmParty.Salir.Visible = False
            LlegoParty = True
            Exit Sub
        Case "LSTCRI"
            rdata = Right$(rdata, Len(rdata) - 6)
            For I = 1 To Val(ReadFieldOptimizado(1, rdata, 44))
                frmEntrenador.lstCriaturas.AddItem ReadFieldOptimizado(I + 1, rdata, 44)
            Next I
            frmEntrenador.Show , frmPrincipal
            Exit Sub
    End Select
    
    Select Case Left$(sdata, 7)
        Case "PEACEDE"
            rdata = Right$(rdata, Len(rdata) - 7)
            Call frmUserRequest.recievePeticion(rdata)
            Exit Sub
        Case "PEACEPR"
            rdata = Right$(rdata, Len(rdata) - 7)
            Call frmPeaceProp.ParsePeaceOffers(rdata)
            Exit Sub
        Case "CHRINFO"
            rdata = Right$(rdata, Len(rdata) - 7)
            Call frmCharInfo.parseCharInfo(rdata)
            frmCharInfo.SetFocus
            Exit Sub
        Case "LEADERI"
            rdata = Right$(rdata, Len(rdata) - 7)
            Call frmGuildLeader.ParseLeaderInfo(rdata)
            frmGuildLeader.SetFocus
            Exit Sub
        Case "GINFIG"
            frmGuildLeader.Show , frmPrincipal
            frmGuildLeader.SetFocus
            Exit Sub
        Case "GINFII"
            frmGuildsNuevo.Show , frmPrincipal
            frmGuildsNuevo.SetFocus
            Exit Sub
        Case "GINFIJ"
            frmGuildAdm.Show , frmPrincipal
            frmGuildAdm.SetFocus
            Exit Sub
        Case "MEMBERI"
            rdata = Right$(rdata, Len(rdata) - 7)
            Call frmGuildsNuevo.ParseMemberInfo(rdata)
            frmGuildsNuevo.SetFocus
            Exit Sub
        Case "CLANDET"
            rdata = Right$(rdata, Len(rdata) - 7)
            Call frmGuildBrief.ParseGuildInfo(rdata)
            Exit Sub
        Case "SHOWFUN"
            rdata = Right$(rdata, Len(rdata) - 7)
            CreandoClan = True
            Call frmGuildFoundation.Show(vbModeless, frmPrincipal)
            Exit Sub
        Case "PETICIO"
            rdata = Right$(rdata, Len(rdata) - 7)
            Call frmUserRequest.recievePeticion(rdata)
            Call frmUserRequest.Show(vbModeless, frmPrincipal)
            Exit Sub
        
    End Select
    
    Select Case UCase$(Left$(rdata, 9))
       Case "DAMEQUEST"
            Call SendData("IQUEST")
            frmQuestSelect.Show , frmPrincipal
        Exit Sub
        
    End Select
    
    
    Call HandleDosLetras(sdata)
    
    If Not Procesado Then Call InformacionEncriptada(sdata)
    
    Procesado = False
    
End Sub
Sub InformacionEncriptada(ByVal rdata As String)
Dim I As Integer

For I = 1 To UBound(Mensajes)
    If UCase$(Left$(rdata, 2)) = UCase$(Mensajes(I).code) Then
        rdata = Right$(rdata, Len(rdata) - 2)
        AddtoRichTextBox frmPrincipal.rectxt, Reemplazo(Mensajes(I).mensaje, rdata), CInt(Mensajes(I).Red), CInt(Mensajes(I).Green), CInt(Mensajes(I).Blue), Mensajes(I).Bold = 1, Mensajes(I).Italic = 1
        Exit Sub
    End If
Next

End Sub
Function Reemplazo(mensaje As String, rdata As String) As String
Dim I As Integer

For I = 1 To Len(mensaje)
    If mid$(mensaje, I, 1) = "*" Then
        Reemplazo = Reemplazo & ReadFieldOptimizado(Val(mid$(mensaje, I + 1, 1)), rdata, 44)
        I = I + 1
    Else
        Reemplazo = Reemplazo & mid$(mensaje, I, 1)
    End If
Next

End Function
Sub HandleDosLetras(ByVal rdata As String)
Dim Charindexx As Integer
Dim tempint As Integer
Dim tempstr As String
Dim I As Integer
Dim X As Integer
Dim Y As Integer
Dim CharIndex As Integer
Dim perso As String
Dim recup As Integer
Dim Slot As Integer
Dim loopc As Integer
Dim Text1 As String
Dim Text2 As String
Dim var3 As Integer
Dim var2 As Integer
Dim var1 As Integer
Dim var4 As Integer

Select Case Left$(rdata, 2)
        Case "HC"
            FrmMenuUser.Label1.Caption = InputBox("¿A quien desea transferir?", "Escribre un Nick.", "")
            FrmMenuUser.textbox.Caption = InputBox("¿Cuantos puntos desea transferir?", "Transferencia de puntos.", "0")
            FrmMenuUser.Label1.Caption = Replace(FrmMenuUser.Label1.Caption, " ", "+")
            Call SendData("/TRANSFERIX " & FrmMenuUser.Label1.Caption & " " & FrmMenuUser.textbox.Caption)
            Exit Sub
        Case "ET"
            Call EliminarDatosMapa
            Exit Sub
        Case "°°"
            CONGELADO = True
            Call AddtoRichTextBox(frmPrincipal.rectxt, "¡SERVIDOR CONGELADO, NO PUEDES ENVIAR INFORMACION HASTA QUE SE DESCONGELE!", 255, 0, 0, 1, 0)
            Exit Sub
        Case "°¬"
            CONGELADO = False
            Call AddtoRichTextBox(frmPrincipal.rectxt, "¡SERVIDOR DESCONGELADO, YA PUEDES ENVIAR INFORMACION AL SERVIDOR!", 255, 0, 0, 1, 0)
            Exit Sub
        Case "CM"
            rdata = Right$(rdata, Len(rdata) - 2)
            UserMap = Val(ReadFieldOptimizado(1, rdata, 44))
            
            NombreDelMapaActual = ReadFieldOptimizado(3, rdata, 44)
            TopMapa = 18 + Val(ReadFieldOptimizado(4, rdata, 44)) * 18
            IzquierdaMapa = 25 + Val(ReadFieldOptimizado(5, rdata, 44)) * 18
            
          '  frmMapa.personaje.Left = IzquierdaMapa + (UserPos.x - 50) * 0.18
            'frmMapa.personaje.Top = TopMapa + (UserPos.y - 50) * 0.18

'          frmMapa.personaje.Visible = (TopMapa > 18 Or IzquierdaMapa > 25)
            
            frmPrincipal.mapa.Caption = NombreDelMapaActual & " [" & UserMap & " - " & UserPos.X & " - " & UserPos.Y & "]"

            If FileExist(DirMapas & "Mapa" & UserMap & ".mcl", vbNormal) Then
                Open DirMapas & "Mapa" & UserMap & ".mcl" For Binary As #1
                Seek #1, 1
                Get #1, , tempint
                Close #1
                If tempint = Val(ReadFieldOptimizado(2, rdata, 44)) Then
                    Call SwitchMapNew(UserMap)
                    If bLluvia(UserMap) = 0 Then
                        If bRain Then
                            Audio.StopWave
                            frmPrincipal.IsPlaying = plNone
                        End If
                    End If
                Else
                    MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
                    Call UnloadAllForms
                    End
                End If
            Else
                
                MsgBox "No se encuentra el mapa instalado."
                Call UnloadAllForms
                Call EscribirGameIni(Config_Inicio)
                End
            End If
            Exit Sub
        Case "PU"
            rdata = Right$(rdata, Len(rdata) - 2)
            rdata = TeEncripTE(rdata)
            MapData(UserPos.X, UserPos.Y).CharIndex = 0
            UserPos.X = CInt(ReadFieldOptimizado(1, rdata, 44))
            UserPos.Y = CInt(ReadFieldOptimizado(2, rdata, 44))
            MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
            CharList(UserCharIndex).POS = UserPos
            Exit Sub
        Case "4&"
            FrmElegirCamino.Show , frmPrincipal
            FrmElegirCamino.SetFocus
            Exit Sub
        Case "N2"
            rdata = Right$(rdata, Len(rdata) - 2)
            I = Val(ReadFieldOptimizado(1, rdata, 44))
            Select Case I
                Case bCabeza
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡La criatura te ha pegado en la cabeza por " & Val(ReadFieldOptimizado(2, rdata, 44)) & "!", 255, 0, 0, 1, 0)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡La criatura te ha pegado el brazo izquierdo por " & Val(ReadFieldOptimizado(2, rdata, 44)) & "!", 255, 0, 0, 1, 0)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡La criatura te ha pegado el brazo derecho por " & Val(ReadFieldOptimizado(2, rdata, 44)) & "!", 255, 0, 0, 1, 0)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡La criatura te ha pegado la pierna izquierda por " & Val(ReadFieldOptimizado(2, rdata, 44)) & "!", 255, 0, 0, 1, 0)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡La criatura te ha pegado la pierna derecha por " & Val(ReadFieldOptimizado(2, rdata, 44)) & "!", 255, 0, 0, 1, 0)
                Case bTorso
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡La criatura te ha pegado en el torso por " & Val(ReadFieldOptimizado(2, rdata, 44)) & "!", 255, 0, 0, 1, 0)
            End Select
            Exit Sub

        Case "2H"
            rdata = Right$(rdata, Len(rdata) - 2)
            Slot = ReadFieldOptimizado(1, rdata, 44)
            UserInventory(Slot).OBJIndex = 0
            UserInventory(Slot).name = "Nada"
            UserInventory(Slot).Amount = 0
            UserInventory(Slot).Equipped = 0
            UserInventory(Slot).GrhIndex = 0
            UserInventory(Slot).ObjType = 0
            UserInventory(Slot).MaxHit = 0
            UserInventory(Slot).MinHit = 0
            UserInventory(Slot).MaxDef = 0
            UserInventory(Slot).MinDef = 0
            UserInventory(Slot).TipoPocion = 0
            UserInventory(Slot).MaxModificador = 0
            UserInventory(Slot).MinModificador = 0
            UserInventory(Slot).Valor = 0
            Call ActualizarInventario(Slot)
            tempstr = ""
            
            bInvMod = True
            
            Exit Sub

        Case "1I"
            rdata = Right$(rdata, Len(rdata) - 2)
            AddtoRichTextBox frmPrincipal.rectxt, rdata & " ha sido aceptado en el clan.", 255, 255, 255, 1, 0
            If FX = 0 Then Call Audio.PlayWave(0, "43.wav")
            Exit Sub
        Case "2I"
            rdata = Right$(rdata, Len(rdata) - 2)
            UserInventory(rdata).Amount = UserInventory(rdata).Amount - 1
            ActualizarInventario (rdata)
        Case "3I"
            rdata = Right$(rdata, Len(rdata) - 2)
        
            UserInventory(rdata).OBJIndex = 0
            UserInventory(rdata).name = "Nada"
            UserInventory(rdata).Amount = 0
            UserInventory(rdata).Equipped = 0
            UserInventory(rdata).GrhIndex = 0
            UserInventory(rdata).ObjType = 0
            UserInventory(rdata).MaxHit = 0
            UserInventory(rdata).MinHit = 0
            UserInventory(rdata).MaxDef = 0
            UserInventory(rdata).MinDef = 0
            UserInventory(rdata).TipoPocion = 0
            UserInventory(rdata).MaxModificador = 0
            UserInventory(rdata).MinModificador = 0
            UserInventory(rdata).Valor = 0

            tempstr = ""
            If UserInventory(rdata).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(rdata).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(rdata).Amount & ") " & UserInventory(rdata).name
            Else
                tempstr = tempstr & UserInventory(rdata).name
            End If
            
            ActualizarInventario (rdata)

            Exit Sub
        Case "4I"
            rdata = Right$(rdata, Len(rdata) - 2)
            Slot = ReadFieldOptimizado(1, rdata, 44)
            UserInventory(Slot).Amount = UserInventory(Slot).Amount - ReadFieldOptimizado(2, rdata, 44)
            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).name
            Else
                tempstr = tempstr & UserInventory(Slot).name
            End If
            
            ActualizarInventario (Slot)
        Case "6J"
            rdata = Right$(rdata, Len(rdata) - 2)
            Slot = ReadFieldOptimizado(1, rdata, 44)
            UserMinAGU = ReadFieldOptimizado(2, rdata, 44)
            frmPrincipal.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 94)
            frmPrincipal.cantidadagua.Caption = UserMinAGU & "/" & UserMaxAGU

            UserInventory(Slot).Amount = UserInventory(Slot).Amount - 1
            If FX = 0 Then
                 Call Audio.PlayWave(0, "46.wav")
            End If
            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).name
            Else
                tempstr = tempstr & UserInventory(Slot).name
            End If
            
            ActualizarInventario (Slot)
            Exit Sub
        Case "6I"
            rdata = Right$(rdata, Len(rdata) - 2)
            Slot = ReadFieldOptimizado(1, rdata, 44)
                UserMinAGU = ReadFieldOptimizado(2, rdata, 44)
                        frmPrincipal.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 94)
            frmPrincipal.cantidadagua.Caption = UserMinAGU & "/" & UserMaxAGU

            UserInventory(Slot).OBJIndex = 0
            UserInventory(Slot).name = "Nada"
            UserInventory(Slot).Amount = 0
            UserInventory(Slot).Equipped = 0
            UserInventory(Slot).GrhIndex = 0
            UserInventory(Slot).ObjType = 0
            UserInventory(Slot).MaxHit = 0
            UserInventory(Slot).MinHit = 0
            UserInventory(Slot).MaxDef = 0
            UserInventory(Slot).MinDef = 0
            UserInventory(Slot).TipoPocion = 0
            UserInventory(Slot).MaxModificador = 0
            UserInventory(Slot).MinModificador = 0
            UserInventory(Slot).Valor = 0

            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).name
            Else
                tempstr = tempstr & UserInventory(Slot).name
            End If
            
            ActualizarInventario (Slot)
            If FX = 0 Then
                 Call Audio.PlayWave(0, "46.wav")
            End If
            Exit Sub
        Case "7I"
            rdata = Right$(rdata, Len(rdata) - 2)
            rdata = THeEnCripTe(rdata, Chr$(83) & Chr$(84) & Chr$(82) & Chr$(73) & Chr$(78) & Chr$(71) & Chr$(71) & Chr$(69) _
            & Chr$(78) & Chr$(77))
            Slot = ReadFieldOptimizado(1, rdata, 44)
            
            UserMinMAN = ReadFieldOptimizado(2, rdata, 44)
                        If UserMaxMAN > 0 Then
                frmPrincipal.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 207)
                frmPrincipal.cantidadmana.Caption = PonerPuntos(UserMinMAN) & "/" & PonerPuntos(UserMaxMAN)
            Else
                frmPrincipal.MANShp.Width = 0
               frmPrincipal.cantidadmana.Caption = ""
            End If
            UserInventory(Slot).Amount = UserInventory(Slot).Amount - 1
            If FX = 0 Then
                 Call Audio.PlayWave(0, "46.wav")
            End If
                        tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).name
            Else
                tempstr = tempstr & UserInventory(Slot).name
            End If
            
            ActualizarInventario (Slot)
            Exit Sub
        Case "8I"
        rdata = Right$(rdata, Len(rdata) - 2)
        Slot = ReadFieldOptimizado(1, rdata, 44)
            UserMinMAN = ReadFieldOptimizado(2, rdata, 44)
                        If UserMaxMAN > 0 Then
                frmPrincipal.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 207)
                frmPrincipal.cantidadmana.Caption = PonerPuntos(UserMinMAN) & "/" & PonerPuntos(UserMaxMAN)
            Else
                frmPrincipal.MANShp.Width = 0
               frmPrincipal.cantidadmana.Caption = ""
            End If
            UserInventory(Slot).OBJIndex = 0
            UserInventory(Slot).name = "Nada"
            UserInventory(Slot).Amount = 0
            UserInventory(Slot).Equipped = 0
            UserInventory(Slot).GrhIndex = 0
            UserInventory(Slot).ObjType = 0
            UserInventory(Slot).MaxHit = 0
            UserInventory(Slot).MinHit = 0
            UserInventory(Slot).MaxDef = 0
            UserInventory(Slot).MinDef = 0
            UserInventory(Slot).TipoPocion = 0
            UserInventory(Slot).MaxModificador = 0
            UserInventory(Slot).MinModificador = 0
            UserInventory(Slot).Valor = 0

            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).name
            Else
                tempstr = tempstr & UserInventory(Slot).name
            End If
            
            ActualizarInventario (Slot)
            If FX = 0 Then
                 Call Audio.PlayWave(0, "46.wav")
            End If
            Exit Sub
        Case "9I"
            rdata = Right$(rdata, Len(rdata) - 2)
            Slot = ReadFieldOptimizado(1, rdata, 44)
            
            UserMinHP = ReadFieldOptimizado(2, rdata, 44)
            frmPrincipal.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 207)
            frmPrincipal.cantidadhp.Caption = PonerPuntos(UserMinHP) & "/" & PonerPuntos(UserMaxHP)
            UserInventory(Slot).Amount = UserInventory(Slot).Amount - 1
            If FX = 0 Then
                 Call Audio.PlayWave(0, "46.wav")
            End If
                        tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).name
            Else
                tempstr = tempstr & UserInventory(Slot).name
            End If
            
            ActualizarInventario (Slot)
            Exit Sub
        Case "2J"
        rdata = Right$(rdata, Len(rdata) - 2)
        Slot = ReadFieldOptimizado(1, rdata, 44)
            UserMinHP = ReadFieldOptimizado(2, rdata, 44)
            frmPrincipal.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 207)
            frmPrincipal.cantidadhp.Caption = PonerPuntos(UserMinHP) & "/" & PonerPuntos(UserMaxHP)
            UserInventory(Slot).OBJIndex = 0
            UserInventory(Slot).name = "Nada"
            UserInventory(Slot).Amount = 0
            UserInventory(Slot).Equipped = 0
            UserInventory(Slot).GrhIndex = 0
            UserInventory(Slot).ObjType = 0
            UserInventory(Slot).MaxHit = 0
            UserInventory(Slot).MinHit = 0
            UserInventory(Slot).MaxDef = 0
            UserInventory(Slot).MinDef = 0
            UserInventory(Slot).TipoPocion = 0
            UserInventory(Slot).MaxModificador = 0
            UserInventory(Slot).MinModificador = 0
            UserInventory(Slot).Valor = 0

            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).name
            Else
                tempstr = tempstr & UserInventory(Slot).name
            End If
            
            ActualizarInventario (Slot)
            If FX = 0 Then
                 Call Audio.PlayWave(0, "46.wav")
            End If
            Exit Sub
        Case "3J"
            Slot = Right$(rdata, Len(rdata) - 2)

            UserInventory(Slot).Amount = UserInventory(Slot).Amount - 1
            If FX = 0 Then
                 Call Audio.PlayWave(0, "46.wav")
            End If
                        tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).name
            Else
                tempstr = tempstr & UserInventory(Slot).name
            End If
            
            ActualizarInventario (Slot)
            Exit Sub
        Case "4J"
        Slot = Right$(rdata, Len(rdata) - 2)
            
            UserInventory(Slot).OBJIndex = 0
            UserInventory(Slot).name = "Nada"
            UserInventory(Slot).Amount = 0
            UserInventory(Slot).Equipped = 0
            UserInventory(Slot).GrhIndex = 0
            UserInventory(Slot).ObjType = 0
            UserInventory(Slot).MaxHit = 0
            UserInventory(Slot).MinHit = 0
            UserInventory(Slot).MaxDef = 0
            UserInventory(Slot).MinDef = 0
            UserInventory(Slot).TipoPocion = 0
            UserInventory(Slot).MaxModificador = 0
            UserInventory(Slot).MinModificador = 0
            UserInventory(Slot).Valor = 0

            tempstr = ""

            If FX = 0 Then
                 Call Audio.PlayWave(0, "46.wav")
            End If
            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).name
            Else
                tempstr = tempstr & UserInventory(Slot).name
            End If
            ActualizarInventario (Slot)
            Exit Sub

        Case "8J"
            rdata = Right$(rdata, Len(rdata) - 2)
            UserInventory(rdata).Equipped = 0
            
            If UserInventory(rdata).ObjType = 2 Then
            frmPrincipal.arma.Caption = "N/A"
            ElseIf UserInventory(rdata).ObjType = 3 Then
            Select Case UserInventory(rdata).SubTipo
                Case 0
                    frmPrincipal.armadura.Caption = "N/A"
                Case 1
                    frmPrincipal.casco.Caption = "N/A"
                Case 2
                    frmPrincipal.escudo.Caption = "N/A"
            End Select
            
            
            End If
                                    tempstr = ""
            If UserInventory(rdata).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(rdata).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(rdata).Amount & ") " & UserInventory(rdata).name
            Else
                tempstr = tempstr & UserInventory(rdata).name
            End If
            
            ActualizarInventario (rdata)
            Exit Sub
        Case "7J"
            rdata = Right$(rdata, Len(rdata) - 2)
            UserInventory(rdata).Equipped = 1
            
            If UserInventory(rdata).ObjType = 2 Then
                frmPrincipal.arma.Caption = UserInventory(rdata).MinHit & " / " & UserInventory(rdata).MaxHit
            ElseIf UserInventory(rdata).ObjType = 3 Then
                Select Case UserInventory(rdata).SubTipo
                    Case 0
                        If UserInventory(rdata).MaxDef > 0 Then
                            frmPrincipal.armadura.Caption = UserInventory(rdata).MinDef & " / " & UserInventory(rdata).MaxDef
                        Else
                            frmPrincipal.armadura.Caption = "N/A"
                        End If

                    Case 1
                        If UserInventory(rdata).MaxDef > 0 Then
                            frmPrincipal.casco.Caption = UserInventory(rdata).MinDef & " / " & UserInventory(rdata).MaxDef
                        Else
                            frmPrincipal.casco.Caption = "N/A"
                        End If
                        
                    Case 2
                        If UserInventory(rdata).MaxDef > 0 Then
                            frmPrincipal.escudo.Caption = UserInventory(rdata).MinDef & " / " & UserInventory(rdata).MaxDef
                        Else
                            frmPrincipal.escudo.Caption = "N/A"
                        End If
                    
                End Select
            End If
            
            tempstr = ""
            If UserInventory(rdata).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(rdata).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(rdata).Amount & ") " & UserInventory(rdata).name
            Else
                tempstr = tempstr & UserInventory(rdata).name
            End If
            
            ActualizarInventario (rdata)
            Exit Sub
        Case "6K"
            rdata = Right$(rdata, Len(rdata) - 2)
            Slot = ReadFieldOptimizado(1, rdata, 44)
            UserMinHAM = ReadFieldOptimizado(2, rdata, 44)
            frmPrincipal.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 94)
            frmPrincipal.cantidadhambre.Caption = UserMinHAM & "/" & UserMaxHAM

            UserInventory(Slot).Amount = UserInventory(Slot).Amount - 1
            If FX = 0 Then
                 Call Audio.PlayWave(0, "7.wav")
            End If
            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).name
            Else
                tempstr = tempstr & UserInventory(Slot).name
            End If
            
            ActualizarInventario (Slot)
            Exit Sub
        Case "7K"
        rdata = Right$(rdata, Len(rdata) - 2)
        Slot = ReadFieldOptimizado(1, rdata, 44)
            UserMinHAM = ReadFieldOptimizado(2, rdata, 44)
            frmPrincipal.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 94)
            frmPrincipal.cantidadhambre.Caption = UserMinHAM & "/" & UserMaxHAM

            UserInventory(Slot).OBJIndex = 0
            UserInventory(Slot).name = "Nada"
            UserInventory(Slot).Amount = 0
            UserInventory(Slot).Equipped = 0
            UserInventory(Slot).GrhIndex = 0
            UserInventory(Slot).ObjType = 0
            UserInventory(Slot).MaxHit = 0
            UserInventory(Slot).MinHit = 0
            UserInventory(Slot).MaxDef = 0
            UserInventory(Slot).MinDef = 0
            UserInventory(Slot).TipoPocion = 0
            UserInventory(Slot).MaxModificador = 0
            UserInventory(Slot).MinModificador = 0
            UserInventory(Slot).Valor = 0

            tempstr = ""
            If UserInventory(Slot).Equipped = 1 Then
                tempstr = tempstr & "(Eqp)"
            End If
            
            If UserInventory(Slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserInventory(Slot).Amount & ") " & UserInventory(Slot).name
            Else
                tempstr = tempstr & UserInventory(Slot).name
            End If
            
            ActualizarInventario (Slot)
            If FX = 0 Then
                 Call Audio.PlayWave(0, "7.wav")
            End If
            Exit Sub
        Case "HM"
            If CreateDamageAC = 0 Then
            rdata = Right$(rdata, Len(rdata) - 2)
          '  Val(ReadFieldOptimizado(5, Rdata, 176) ' charindex lo dejo por si algun dia lo uso
            Dim dañocausado As Integer
            dañocausado = Val(ReadFieldOptimizado(2, rdata, 176))
            Dim dañocolor As Integer
            dañocolor = Val(ReadFieldOptimizado(6, rdata, 176))
            If dañocolor = 1 Then
            CreateDamage dañocausado, 102, 255, 102, Val(ReadFieldOptimizado(3, rdata, 176)), Val(ReadFieldOptimizado(4, rdata, 176))
            ElseIf dañocolor = 2 Then
            CreateDamage dañocausado, 255, 204, 204, Val(ReadFieldOptimizado(3, rdata, 176)), Val(ReadFieldOptimizado(4, rdata, 176))
            Else
            CreateDamage dañocausado, 255, 255, 0, Val(ReadFieldOptimizado(3, rdata, 176)), Val(ReadFieldOptimizado(4, rdata, 176))
            End If
            End If
        Exit Sub
        Case "3Q"
            rdata = Right$(rdata, Len(rdata) - 2)
            Dim ibser As Integer
            ibser = Val(ReadFieldOptimizado(3, rdata, 176))
            If ibser > 0 Then
            Dialogos.CrearDialogo ReadFieldOptimizado(2, rdata, 176), ibser, Val(ReadFieldOptimizado(1, rdata, 176))
              
                
                
                
                
            Else
                  If PuedoQuitarFoco Then _
                    AddtoRichTextBox frmPrincipal.rectxt, ReadFieldOptimizado(1, rdata, 126), Val(ReadFieldOptimizado(2, rdata, 126)), Val(ReadFieldOptimizado(3, rdata, 126)), Val(ReadFieldOptimizado(4, rdata, 126)), Val(ReadFieldOptimizado(5, rdata, 126)), Val(ReadFieldOptimizado(6, rdata, 126))
            End If
            Exit Sub
        Case "9Q"
            rdata = Right$(rdata, Len(rdata) - 2)
            Dim CRI As String
            Text1 = ReadFieldOptimizado(1, rdata, 44)
            Text2 = ReadFieldOptimizado(2, rdata, 44)
            
            Select Case Val(Text2)
                Case 1
                    CRI = " [Herido]"
                Case 2
                    CRI = " [Levemente herido]"
                Case 3
                    CRI = " [Muy herido]"
                Case 4
                    CRI = " [Agonizando]"
                Case 5
                    CRI = " [Sano]"
                Case Is > 5
                    CRI = " [" & Val(Text2) - 5 & "0% herido]"
            End Select
        
            AddtoRichTextBox frmPrincipal.rectxt, Text1 & CRI, 65, 190, 156, 0, 0
            Exit Sub
        Case "7T"
            rdata = Right$(rdata, Len(rdata) - 2)
            Text1 = ReadFieldOptimizado(1, rdata, 172)
            Text2 = ReadFieldOptimizado(2, rdata, 172)
            var1 = Val(ReadFieldOptimizado(3, rdata, 172))
            var2 = Val(ReadFieldOptimizado(4, rdata, 172))
            var3 = Val(ReadFieldOptimizado(5, rdata, 172))
            AddtoRichTextBox frmPrincipal.rectxt, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%", 65, 190, 156, 0, 0
            AddtoRichTextBox frmPrincipal.rectxt, "Nombre del hechizo: " & Text1, 65, 190, 156, 0, 0
            AddtoRichTextBox frmPrincipal.rectxt, "Descripción: " & Text2, 65, 190, 156, 0, 0
            AddtoRichTextBox frmPrincipal.rectxt, "Skill requerido: " & var1, 65, 190, 156, 0, 0
            AddtoRichTextBox frmPrincipal.rectxt, "Mana necesaria: " & var2, 65, 190, 156, 0, 0
            AddtoRichTextBox frmPrincipal.rectxt, "Energia necesaria: " & var3, 65, 190, 156, 0, 0
            Exit Sub
        Case "1U"
            rdata = Right$(rdata, Len(rdata) - 2)
            var1 = Val(ReadFieldOptimizado(1, rdata, 44))
            var2 = Val(ReadFieldOptimizado(2, rdata, 44))
            var3 = Val(ReadFieldOptimizado(3, rdata, 44))
            var4 = Val(ReadFieldOptimizado(4, rdata, 44))
            If var1 > 0 Then AddtoRichTextBox frmPrincipal.rectxt, "Has ganado " & var1 & " puntos de vida.", 200, 200, 200, 0, 0
            If var2 > 0 Then AddtoRichTextBox frmPrincipal.rectxt, "Has ganado " & var2 & " puntos de vitalidad.", 200, 200, 200, 0, 0
            If var3 > 0 Then AddtoRichTextBox frmPrincipal.rectxt, "Has ganado " & var3 & " puntos de mana.", 200, 200, 200, 0, 0
            If var4 > 0 Then AddtoRichTextBox frmPrincipal.rectxt, "Tu golpe maximo aumentó en " & var4 & " puntos.", 200, 200, 200, 0, 0
            If var4 > 0 Then AddtoRichTextBox frmPrincipal.rectxt, "Tu golpe mínimo aumentó en " & var4 & " puntos.", 200, 200, 200, 0, 0
            Exit Sub
        Case "6Z"
            AddtoRichTextBox frmPrincipal.rectxt, "¡Hoy es la votación para elegir un nuevo lider para el clan!", 255, 255, 255, 1, 0
            AddtoRichTextBox frmPrincipal.rectxt, "La elección durará 24 horas, se puede votar a cualquier miembro del clan.", 255, 255, 255, 1, 0
            AddtoRichTextBox frmPrincipal.rectxt, "Para votar escribe /VOTO NICKNAME.", 255, 255, 255, 1, 0
            AddtoRichTextBox frmPrincipal.rectxt, "Sólo se computara un voto por miembro.", 255, 255, 255, 1, 0
            Exit Sub
        Case "7Z"
            rdata = Right$(rdata, Len(rdata) - 2)
            AddtoRichTextBox frmPrincipal.rectxt, "¡Las elecciones han finalizado!", 255, 255, 255, 1, 0
            AddtoRichTextBox frmPrincipal.rectxt, "El nuevo lider es: " & rdata, 255, 255, 255, 1, 0
            Exit Sub
        Case "!J"
            rdata = Right$(rdata, Len(rdata) - 2)
            AddtoRichTextBox frmPrincipal.rectxt, "Felicitaciones, tu solicitud ha sido aceptada.", 255, 255, 255, 1, 0
            AddtoRichTextBox frmPrincipal.rectxt, "Ahora sos un miembro activo del clan " & rdata, 255, 255, 255, 1, 0
            Exit Sub
        Case "!R"
            rdata = Right$(rdata, Len(rdata) - 2)
            AddtoRichTextBox frmPrincipal.rectxt, "Tu clan ha firmado una alianza con " & rdata, 255, 255, 255, 1, 0
            If FX = 0 Then
                 Call Audio.PlayWave(0, "45.wav")
            End If
            Exit Sub
        Case "!S"
            rdata = Right$(rdata, Len(rdata) - 2)
            AddtoRichTextBox frmPrincipal.rectxt, rdata & " firmó una alianza con tu clan.", 255, 255, 255, 1, 0
            If FX = 0 Then
                 Call Audio.PlayWave(0, "45.wav")
            End If
            Exit Sub
        Case "!U"
            rdata = Right$(rdata, Len(rdata) - 2)
            AddtoRichTextBox frmPrincipal.rectxt, "Tu clan le declaró la guerra a " & rdata, 255, 255, 255, 1, 0
            If FX = 0 Then
                 Call Audio.PlayWave(0, "45.wav")
            End If
            Exit Sub
        Case "!V"
            rdata = Right$(rdata, Len(rdata) - 2)
            AddtoRichTextBox frmPrincipal.rectxt, rdata & " le declaró la guerra a tu clan.", 255, 255, 255, 1, 0
            If FX = 0 Then
                 Call Audio.PlayWave(0, "45.wav")
            End If
            Exit Sub
        Case "!4"
            rdata = Right$(rdata, Len(rdata) - 2)
            Text1 = ReadFieldOptimizado(1, rdata, 44)
            Text2 = ReadFieldOptimizado(2, rdata, 44)
            AddtoRichTextBox frmPrincipal.rectxt, "¡" & Text1 & " fundó el clan " & Text2 & "!", 255, 255, 255, 1, 0
            If FX = 0 Then
                 Call Audio.PlayWave(0, "44.wav")
            End If
            Exit Sub
        Case "/O"
            rdata = Right$(rdata, Len(rdata) - 2)
            Call Dialogos.CrearDialogo("El negocio va bien, ya he conseguido " & ReadFieldOptimizado(1, rdata, 44) & " monedas de oro en ventas. He enviado el dinero directamente a tu cuenta en el banco.", Val(ReadFieldOptimizado(2, rdata, 44)), vbWhite)
            Exit Sub
        Case "/P"
            rdata = Right$(rdata, Len(rdata) - 2)
            Call Dialogos.CrearDialogo("El negocio no va muy bien, todavía no he podido vender nada. Si consigo una venta, enviare el dinero directamente a tu cuenta en el banco.", Val(rdata), vbWhite)
            Exit Sub
        Case "/Q"
            rdata = Right$(rdata, Len(rdata) - 2)
            Call Dialogos.CrearDialogo("¡Buen día! Ahora estoy contratado por " & ReadFieldOptimizado(1, rdata, 44) & " para vender sus objetos, ¿quieres echar un vistazo?", Val(ReadFieldOptimizado(2, rdata, 44)), vbWhite)
            Exit Sub
        Case "/R"
            rdata = Right$(rdata, Len(rdata) - 2)
            AddtoRichTextBox frmPrincipal.rectxt, ReadFieldOptimizado(1, rdata, 44) & " compró " & ReadFieldOptimizado(2, rdata, 44) & " (" & PonerPuntos(Val(ReadFieldOptimizado(3, rdata, 44))) & ") en tu tienda por " & PonerPuntos(Val(ReadFieldOptimizado(4, rdata, 44))) & " monedas de oro.", 255, 255, 255, 1, 0
            AddtoRichTextBox frmPrincipal.rectxt, "El dinero fue enviado directamente a tu cuenta de banco.", 255, 255, 255, 1, 0
            Exit Sub
        Case "/V"
            rdata = Right$(rdata, Len(rdata) - 2)
            Call Dialogos.CrearDialogo("Solo los trabajadores experimentados y los personajes mayores a nivel 25 con más de 65 en comercio pueden utilizar mis servicios de vendedor.", Val(rdata), vbWhite)
            Exit Sub
        Case "/X"
            rdata = Right$(rdata, Len(rdata) - 2)
            AddtoRichTextBox frmPrincipal.rectxt, "Numero total de ventas: " & PonerPuntos(Val(ReadFieldOptimizado(2, rdata, 44))), 65, 190, 156, 1, 0
            AddtoRichTextBox frmPrincipal.rectxt, "Dinero movido por las ventas: " & PonerPuntos(Val(ReadFieldOptimizado(1, rdata, 44))) & " monedas de oro.", 65, 190, 156, 1, 0
            AddtoRichTextBox frmPrincipal.rectxt, "Venta promedio: " & PonerPuntos(Val(ReadFieldOptimizado(1, rdata, 44)) / Val(ReadFieldOptimizado(2, rdata, 44))) & " monedas de oro.", 65, 190, 156, 1, 0
            Exit Sub
        Case "{B"
            rdata = Right$(rdata, Len(rdata) - 2)
            AddtoRichTextBox frmPrincipal.rectxt, "Has iniciado el modo de susurro con " & rdata & ".", 255, 255, 255, 1, 0
            frmPrincipal.MousePointer = 1
            Exit Sub
        Case "{C"
            rdata = Right$(rdata, Len(rdata) - 2)
            AddtoRichTextBox frmPrincipal.rectxt, "No puedes iniciar el modo de susurro contigo mismo.", 255, 255, 255, 1, 0
            frmPrincipal.modo = "1 Normal"
            frmPrincipal.MousePointer = 1
            Exit Sub
        Case "{D"
            AddtoRichTextBox frmPrincipal.rectxt, "Target invalido.", 65, 190, 156, 0, 0
            frmPrincipal.modo = "1 Normal"
            frmPrincipal.MousePointer = 1
            Exit Sub
        Case "{F"
            AddtoRichTextBox frmPrincipal.rectxt, "El usuario ya no se encuentra en tu pantalla.", 65, 190, 156, 0, 0
            frmPrincipal.modo = "1 Normal"
            frmPrincipal.MousePointer = 1
            Exit Sub
        Case "8B"
            rdata = Right$(rdata, Len(rdata) - 2)
            UserMaxHP = Val(ReadFieldOptimizado(1, rdata, 44))
            frmPrincipal.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 207)
            frmPrincipal.cantidadhp.Caption = PonerPuntos(UserMinHP) & "/" & PonerPuntos(UserMaxHP)
            Exit Sub
        Case "9B"
            rdata = Right$(rdata, Len(rdata) - 2)
            UserMaxMAN = Val(ReadFieldOptimizado(1, rdata, 44))
            
            If UserMaxMAN > 0 Then
                frmPrincipal.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 207)
                frmPrincipal.cantidadmana.Caption = PonerPuntos(UserMinMAN) & "/" & PonerPuntos(UserMaxMAN)
            Else
                frmPrincipal.MANShp.Width = 0
               frmPrincipal.cantidadmana.Caption = ""
            End If
            Exit Sub
        Case "1N"
          '  If CartelSanado = 1 Then AddtoRichTextBox frmPrincipal.rectxt, "Has sanado.", 65, 190, 156, 0, 0
            Exit Sub
        Case "V5"
           ' If CartelOcultarse = 1 Then AddtoRichTextBox frmPrincipal.rectxt, "¡Has vuelto a ser visible!", 65, 190, 156, 0, 0
            Exit Sub
        Case "MN"
            rdata = Right$(rdata, Len(rdata) - 2)
            rdata = THeEnCripTe(rdata, Chr$(83) & Chr$(84) & Chr$(82) & Chr$(73) & Chr$(78) & Chr$(71) & Chr$(71) & Chr$(69) _
            & Chr$(78) & Chr$(77))
         '   If CartelRecuMana = 1 Then
            AddtoRichTextBox frmPrincipal.rectxt, "¡Has recuperado " & rdata & " puntos de mana!", 65, 190, 156, 0, 0
            Exit Sub
        Case "8K"
           ' If CartelNoHayNada = 1 Then AddtoRichTextBox frmPrincipal.rectxt, "¡No hay nada aquí!", 65, 190, 156, 0, 0
            Exit Sub
        Case "DN"
           ' If CartelMenosCansado = 1 Then AddtoRichTextBox frmPrincipal.rectxt, "Has dejado de descansar.", 65, 190, 156, 0, 0
            Exit Sub
        Case "D9"
         '   If CartelRecuMana = 1 Then
            AddtoRichTextBox frmPrincipal.rectxt, "Ya no estás meditando.", 65, 190, 156, 0, 0
            Exit Sub
        Case "{{"
            rdata = Right$(rdata, Len(rdata) - 2)
            AddtoRichTextBox frmPrincipal.rectxt, "(" & ReadFieldOptimizado(1, rdata, 44) & ") " & KeyName(ReadFieldOptimizado(2, rdata, 44)), 65, 190, 156, 0, 0
            Exit Sub
        Case "7M"
           ' If CartelRecuMana = 1 Then
            AddtoRichTextBox frmPrincipal.rectxt, "Comienzas a meditar.", 65, 190, 156, 0, 0
            Exit Sub
        Case "8M"
            rdata = Right$(rdata, Len(rdata) - 2)
            'If CartelRecuMana = 1 Then
             AddtoRichTextBox frmPrincipal.rectxt, "Te estás concentrando. En " & rdata & " segundos comenzarás a meditar.", 65, 190, 156, 0, 0
            Exit Sub
        Case "EL"
            rdata = Right$(rdata, Len(rdata) - 2)
            If rdata <> 0 Then AddtoRichTextBox frmPrincipal.rectxt, "Has obtenido " & rdata & " puntos de exp.", 255, 150, 25, 1, 0
            AddtoRichTextBox frmPrincipal.rectxt, "¡Has matado a la criatura!", 65, 190, 156, 0, 0
            Exit Sub
        Case "V7"
          '  If CartelOcultarse = 1 Then AddtoRichTextBox frmPrincipal.rectxt, "¡Te has escondido entre las sombras!", 65, 190, 156, 0, 0
            Exit Sub
        Case "EN"
            'If CartelOcultarse = 1 Then AddtoRichTextBox frmPrincipal.rectxt, "¡No has logrado esconderte!", 65, 190, 156, 0, 0
            Exit Sub
        Case "V3"
            rdata = Right$(rdata, Len(rdata) - 2)
            rdata = TeEncripTE(rdata)
            CharIndex = Val(ReadFieldOptimizado(2, rdata, 44))
            CharList(CharIndex).invisible = (Val(ReadFieldOptimizado(1, rdata, 44)) = 1)
            Exit Sub
        Case "N4"
            rdata = Right$(rdata, Len(rdata) - 2)
            I = Val(ReadFieldOptimizado(1, rdata, 44))
            Select Case I
                Case bCabeza
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡¡" & ReadFieldOptimizado(3, rdata, 44) & " te ha pegado en la cabeza por " & Val(ReadFieldOptimizado(2, rdata, 44)) & "!!", 255, 0, 0, 1, 0)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡¡" & ReadFieldOptimizado(3, rdata, 44) & " te ha pegado el brazo izquierdo por " & Val(ReadFieldOptimizado(2, rdata, 44)) & "!!", 255, 0, 0, 1, 0)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡¡" & ReadFieldOptimizado(3, rdata, 44) & " te ha pegado el brazo derecho por " & Val(ReadFieldOptimizado(2, rdata, 44)) & "!!", 255, 0, 0, 1, 0)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡¡" & ReadFieldOptimizado(3, rdata, 44) & " te ha pegado la pierna izquierda por " & Val(ReadFieldOptimizado(2, rdata, 44)) & "!!", 255, 0, 0, 1, 0)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡¡" & ReadFieldOptimizado(3, rdata, 44) & " te ha pegado la pierna derecha por " & Val(ReadFieldOptimizado(2, rdata, 44)) & "!!", 255, 0, 0, 1, 0)
                Case bTorso
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡¡" & ReadFieldOptimizado(3, rdata, 44) & " te ha pegado en el torso por " & Val(ReadFieldOptimizado(2, rdata, 44)) & "!!", 255, 0, 0, 1, 0)
            End Select
            Exit Sub
        Case "N5"
            rdata = Right$(rdata, Len(rdata) - 2)
            I = Val(ReadFieldOptimizado(1, rdata, 44))
            Select Case I
                Case bCabeza
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡¡Le has pegado a " & ReadFieldOptimizado(3, rdata, 44) & " en la cabeza por " & Val(ReadFieldOptimizado(2, rdata, 44)), 230, 230, 0, 1, 0)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡¡Le has pegado a " & ReadFieldOptimizado(3, rdata, 44) & " en el brazo izquierdo por " & Val(ReadFieldOptimizado(2, rdata, 44)), 230, 230, 0, 1, 0)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡¡Le has pegado a " & ReadFieldOptimizado(3, rdata, 44) & " en el brazo derecho por " & Val(ReadFieldOptimizado(2, rdata, 44)), 230, 230, 0, 1, 0)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡¡Le has pegado a " & ReadFieldOptimizado(3, rdata, 44) & " en la pierna izquierda por " & Val(ReadFieldOptimizado(2, rdata, 44)), 230, 230, 0, 1, 0)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡¡Le has pegado a " & ReadFieldOptimizado(3, rdata, 44) & " en la pierna derecha por " & Val(ReadFieldOptimizado(2, rdata, 44)), 230, 230, 0, 1, 0)
                Case bTorso
                    Call AddtoRichTextBox(frmPrincipal.rectxt, "¡¡Le has pegado a " & ReadFieldOptimizado(3, rdata, 44) & " en el torso por " & Val(ReadFieldOptimizado(2, rdata, 44)), 230, 230, 0, 1, 0)
            End Select
            Exit Sub
        Case "|$"
              rdata = Right$(rdata, Len(rdata) - 2)
              Call AddtoRichTextBox(frmPrincipal.rectxt, rdata, 240, 238, 207, 0, 0)
          Exit Sub
        Case "##"
rdata = Right$(rdata, Len(rdata) - 2)
quecarajodijo = rdata
Exit Sub
        Case "||"
            Dim iUser As Integer
            rdata = Right$(rdata, Len(rdata) - 2)
            iUser = Val(ReadFieldDarkFly2(3, rdata, 176))
            If iUser > 0 Then
                If Val(ReadFieldDarkFly2(1, rdata, 176)) <> vbCyan And EstaIgnorado(iUser) Then
                    Dialogos.CrearDialogo "", iUser, Val(ReadFieldDarkFly2(1, rdata, 176))
                    Exit Sub
                Else
                    Dialogos.CrearDialogo ReadFieldDarkFly2(2, rdata, 176), iUser, Val(ReadFieldDarkFly2(1, rdata, 176))
                End If
            Else
                  If PuedoQuitarFoco Then _
                    AddtoRichTextBox frmPrincipal.rectxt, ReadFieldDarkFly2(1, rdata, 126), Val(ReadFieldDarkFly2(2, rdata, 126)), Val(ReadFieldDarkFly2(3, rdata, 126)), Val(ReadFieldDarkFly2(4, rdata, 126)), Val(ReadFieldDarkFly2(5, rdata, 126)), Val(ReadFieldDarkFly2(6, rdata, 126))
            End If
            Exit Sub
        Case "!!"
            If PuedoQuitarFoco Then
                rdata = Right$(rdata, Len(rdata) - 2)
                frmMensaje.msg.Caption = rdata
                frmMensaje.Show , frmPrincipal
            End If
            Exit Sub
        Case "FC" 'flecha a char
        rdata = Right$(rdata, Len(rdata) - 2)
        Crear_Flecha Val(ReadFieldOptimizado(1, rdata, 44)), Val(ReadFieldOptimizado(2, rdata, 44)), Val(ReadFieldOptimizado(3, rdata, 44)), 0, Val(ReadFieldOptimizado(4, rdata, 44))
        Exit Sub
        Case "IU"
            rdata = Right$(rdata, Len(rdata) - 2)
            UserIndex = Val(rdata)
            Exit Sub
        Case "IP"
            rdata = Right$(rdata, Len(rdata) - 2)
            UserCharIndex = Val(rdata)
            UserPos = CharList(UserCharIndex).POS
            frmPrincipal.mapa.Caption = NombreDelMapaActual & " [" & UserMap & " - " & UserPos.X & " - " & UserPos.Y & "]"
            Exit Sub
        Case "CC"
            rdata = Right$(rdata, Len(rdata) - 2)
            CharIndex = ReadFieldDarkFly2(4, rdata, 44)
            X = ReadFieldDarkFly2(5, rdata, 44)
            Y = ReadFieldDarkFly2(6, rdata, 44)
            CharList(CharIndex).FX = Val(ReadFieldDarkFly2(9, rdata, 44))
            CharList(CharIndex).FxLoopTimes = Val(ReadFieldDarkFly2(10, rdata, 44))
            CharList(CharIndex).Nombre = ReadFieldDarkFly2(11, rdata, 44)
            CharList(CharIndex).NombreNPC = ReadFieldDarkFly2(7, rdata, 44)
            If Right$(CharList(CharIndex).Nombre, 2) = "<>" Then
                CharList(CharIndex).Nombre = Left$(CharList(CharIndex).Nombre, Len(CharList(CharIndex).Nombre) - 2)
            End If
            
            'meditaciones
            If MeditacionesAZ = 0 Then
            If CharList(CharIndex).FX = 4 Or CharList(CharIndex).FX = 5 Or CharList(CharIndex).FX = 6 Or CharList(CharIndex).FX = 25 Then
                 CharList(CharIndex).FX = 0
                 End If
            End If
            
            CharList(CharIndex).Criminal = Val(ReadFieldDarkFly2(13, rdata, 44))
            CharList(CharIndex).Privilegios = Val(ReadFieldDarkFly2(16, rdata, 44))
            
            CharList(CharIndex).invisible = (Val(ReadFieldDarkFly2(15, rdata, 44)) = 1)
            Call MakeChar(CharIndex, ReadFieldDarkFly2(1, rdata, 44), ReadFieldDarkFly2(2, rdata, 44), ReadFieldDarkFly2(3, rdata, 44), X, Y, Val(ReadFieldDarkFly2(7, rdata, 44)), Val(ReadFieldDarkFly2(8, rdata, 44)), Val(ReadFieldDarkFly2(12, rdata, 44)))
            CharList(CharIndex).aura_Index = Val(ReadFieldDarkFly2(14, rdata, 44))
            Call InitGrh(CharList(CharIndex).Aura, Val(ReadFieldDarkFly2(14, rdata, 44)))
            
            Exit Sub
        Case "CX"
            rdata = Right$(rdata, Len(rdata) - 2)
            Call FrmOpciones.CargarPersonajesWARP(rdata)
            Exit Sub
        Case "PW"
            rdata = Right$(rdata, Len(rdata) - 2)

            CharIndex = ReadFieldOptimizado(1, rdata, 44)
            CharList(CharIndex).Criminal = Val(ReadFieldOptimizado(2, rdata, 44))
            CharList(CharIndex).Nombre = ReadFieldOptimizado(3, rdata, 44)
            
            Exit Sub
            
        Case "BP"
            rdata = Right$(rdata, Len(rdata) - 2)
            Call EraseChar(Val(rdata))
            Exit Sub

        Case "MP"
            rdata = Right$(rdata, Len(rdata) - 2)
            rdata = THeEnCripTe(rdata, Chr$(83) & Chr$(84) & Chr$(82) & Chr$(73) & Chr$(78) & Chr$(71) & Chr$(71) & Chr$(69) _
            & Chr$(78) & Chr$(77))
            CharIndex = Val(ReadFieldOptimizado(1, rdata, 44))
            
            If FX = 0 Then Call DoPasosFx(CharIndex)
            
            Call MoveCharByPos(CharIndex, ReadFieldOptimizado(2, rdata, 44), ReadFieldOptimizado(3, rdata, 44))
            
            Exit Sub
        Case "LP"
            rdata = Right$(rdata, Len(rdata) - 2)
            CharIndex = Val(ReadFieldOptimizado(1, rdata, 44))
            If FX = 0 Then Call DoPasosFx(CharIndex)
            
            Call MoveCharByPosConHeading(CharIndex, ReadFieldOptimizado(2, rdata, 44), ReadFieldOptimizado(3, rdata, 44), ReadFieldOptimizado(4, rdata, 44))
            
            Exit Sub
        Case "ZZ"
            rdata = Right$(rdata, Len(rdata) - 2)
            CharIndex = Val(ReadFieldOptimizado(1, rdata, 44))
            
            If FX = 0 Then Call DoPasosFx(CharIndex)
            
            Call MoveCharByPosAndHead(CharIndex, ReadFieldOptimizado(2, rdata, 44), ReadFieldOptimizado(3, rdata, 44), ReadFieldOptimizado(4, rdata, 44))
            Exit Sub
        Case "MH"
            rdata = Right$(rdata, Len(rdata) - 2)
            Dim usermuerto As Integer
    
            usermuerto = Val(ReadFieldOptimizado(1, rdata, 44))
            If usermuerto = 1 Then
            base_light = D3DColorXRGB(150, 150, 150)
            ElseIf usermuerto = 0 Then
            base_light = D3DColorXRGB(255, 255, 255)
            End If
        Exit Sub
        Case "CP"
            rdata = Right$(rdata, Len(rdata) - 2)
    
            CharIndex = Val(ReadFieldOptimizado(1, rdata, 44))
            CharList(CharIndex).muerto = Val(ReadFieldOptimizado(2, rdata, 44)) = 500
            Slot = Val(ReadFieldOptimizado(2, rdata, 44))
            CharList(CharIndex).Body = BodyData(Slot)
            CharList(CharIndex).Head = HeadData(Val(ReadFieldOptimizado(3, rdata, 44)))
            If Slot > 83 And Slot < 88 Then
                CharList(CharIndex).Navegando = 1
            Else
                CharList(CharIndex).Navegando = 0
            End If
            CharList(CharIndex).Heading = Val(ReadFieldOptimizado(4, rdata, 44))
            CharList(CharIndex).FX = Val(ReadFieldOptimizado(7, rdata, 44))
            CharList(CharIndex).FxLoopTimes = Val(ReadFieldOptimizado(8, rdata, 44))
            
            'meditaciones
            If MeditacionesAZ = 0 Then
            If CharList(CharIndex).FX = 4 Or CharList(CharIndex).FX = 5 Or CharList(CharIndex).FX = 6 Or CharList(CharIndex).FX = 25 Then
                 CharList(CharIndex).FX = 0
                 End If
            End If
            
            tempint = Val(ReadFieldOptimizado(5, rdata, 44))
            If tempint <> 0 Then CharList(CharIndex).arma = WeaponAnimData(tempint)
            tempint = Val(ReadFieldOptimizado(6, rdata, 44))
            If tempint <> 0 Then CharList(CharIndex).escudo = ShieldAnimData(tempint)
            tempint = Val(ReadFieldOptimizado(9, rdata, 44))
            If tempint <> 0 Then CharList(CharIndex).casco = CascoAnimData(tempint)
            Exit Sub
        Case "2C"
            rdata = Right$(rdata, Len(rdata) - 2)
            CharIndex = Val(ReadFieldOptimizado(1, rdata, 44))
            CharList(CharIndex).FX = 0
            CharList(CharIndex).FxLoopTimes = 0
            CharList(CharIndex).Heading = Val(ReadFieldOptimizado(2, rdata, 44))
            
            'meditaciones
            If MeditacionesAZ = 0 Then
            If CharList(CharIndex).FX = 4 Or CharList(CharIndex).FX = 5 Or CharList(CharIndex).FX = 6 Or CharList(CharIndex).FX = 25 Then
            CharList(CharIndex).FX = 0
            End If
            End If
            
            Exit Sub
        Case "3C"
            rdata = Right$(rdata, Len(rdata) - 2)
            CharIndex = Val(ReadFieldOptimizado(1, rdata, 44))
            Slot = Val(ReadFieldOptimizado(2, rdata, 44))
            CharList(CharIndex).Body = BodyData(Slot)
            If Slot > 83 And Slot < 88 Then
                CharList(CharIndex).Navegando = 1
            Else
                CharList(CharIndex).Navegando = 0
            End If
            Exit Sub
        Case "4C"
            rdata = Right$(rdata, Len(rdata) - 2)
            CharIndex = Val(ReadFieldOptimizado(1, rdata, 44))
            CharList(CharIndex).Head = HeadData(Val(ReadFieldOptimizado(2, rdata, 44)))
            Exit Sub
        Case "5C"
            rdata = Right$(rdata, Len(rdata) - 2)
            CharIndex = Val(ReadFieldOptimizado(1, rdata, 44))
            tempint = Val(ReadFieldOptimizado(2, rdata, 44))
            If tempint <> 0 Then CharList(CharIndex).arma = WeaponAnimData(tempint)
            Exit Sub
        Case "6C"
            rdata = Right$(rdata, Len(rdata) - 2)
            CharIndex = Val(ReadFieldOptimizado(1, rdata, 44))
            tempint = Val(ReadFieldOptimizado(2, rdata, 44))
            If tempint <> 0 Then CharList(CharIndex).escudo = ShieldAnimData(tempint)
            Exit Sub
        Case "7C"
            rdata = Right$(rdata, Len(rdata) - 2)
            CharIndex = Val(ReadFieldOptimizado(1, rdata, 44))
            tempint = Val(ReadFieldOptimizado(2, rdata, 44))
            If tempint <> 0 Then CharList(CharIndex).casco = CascoAnimData(tempint)
            Exit Sub
        Case "5A"
            rdata = Right$(rdata, Len(rdata) - 2)
            rdata = TeEncripTE(rdata)
            UserMinHP = Val(ReadFieldOptimizado(1, rdata, 44))
            frmPrincipal.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 207)
            frmPrincipal.cantidadhp.Caption = PonerPuntos(UserMinHP) & "/" & PonerPuntos(UserMaxHP)
            
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
        
            Exit Sub
        Case "5D"
            rdata = Right$(rdata, Len(rdata) - 2)
            UserMinMAN = Val(ReadFieldOptimizado(1, rdata, 44))
            
            If UserMaxMAN > 0 Then
                frmPrincipal.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 207)
                frmPrincipal.cantidadmana.Caption = PonerPuntos(UserMinMAN) & "/" & PonerPuntos(UserMaxMAN)
            Else
                frmPrincipal.MANShp.Width = 0
               frmPrincipal.cantidadmana.Caption = ""
            End If
            
            Exit Sub
            
          Case "5E"
            rdata = Right$(rdata, Len(rdata) - 2)
            UserMinSTA = Val(ReadFieldOptimizado(1, rdata, 44))
            
            frmPrincipal.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
            frmPrincipal.cantidadsta.Caption = PonerPuntos(UserMinSTA) & "/" & PonerPuntos(UserMaxSTA)
        
            Exit Sub

        Case "5F"
            rdata = Right$(rdata, Len(rdata) - 2)
            UserGLD = Val(ReadFieldOptimizado(1, rdata, 44))

            frmPrincipal.GldLbl.Caption = PonerPuntos(UserGLD)
        
            Exit Sub
            
        Case "ELV"
            frmPrincipal.LvlLbl.Caption = "¡Nivel Máximo!"
            frmPrincipal.barrita.Width = 214
            Exit Sub
        Case "5G"
            rdata = Right$(rdata, Len(rdata) - 2)
            
            UserExp = Val(ReadFieldOptimizado(1, rdata, 44))
            
            If UserPasarNivel > 0 Then
                frmPrincipal.lblNivel = UserLvl
                frmPrincipal.barrita.Width = Round(CDbl(UserExp) * CDbl(214) / CDbl(UserPasarNivel), 0)
          
                frmPrincipal.LvlLbl.Caption = " (" & Round(UserExp / UserPasarNivel * 100, 2) & "%)" & " - " & PonerPuntos(UserExp) & " / " & PonerPuntos(UserPasarNivel)
            Else
            
            End If
            
        Case "5H"
            rdata = Right$(rdata, Len(rdata) - 2)
            UserMinMAN = Val(ReadFieldOptimizado(1, rdata, 44))
            UserMinSTA = Val(ReadFieldOptimizado(2, rdata, 44))
            
            If UserMaxMAN > 0 Then
                frmPrincipal.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 207)
                frmPrincipal.cantidadmana.Caption = PonerPuntos(UserMinMAN) & "/" & PonerPuntos(UserMaxMAN)
            Else
                frmPrincipal.MANShp.Width = 0
               frmPrincipal.cantidadmana.Caption = ""
            End If
            
            frmPrincipal.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
            frmPrincipal.cantidadsta.Caption = PonerPuntos(UserMinSTA) & "/" & PonerPuntos(UserMaxSTA)
        
            Exit Sub
            
        Case "5I"
            rdata = Right$(rdata, Len(rdata) - 2)
            UserMinHP = Val(ReadFieldOptimizado(1, rdata, 44))
            UserMinSTA = Val(ReadFieldOptimizado(2, rdata, 44))
    
            frmPrincipal.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 207)
            frmPrincipal.cantidadhp.Caption = PonerPuntos(UserMinHP) & "/" & PonerPuntos(UserMaxHP)
            
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
            
            frmPrincipal.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
            frmPrincipal.cantidadsta.Caption = PonerPuntos(UserMinSTA) & "/" & PonerPuntos(UserMaxSTA)
        
            Exit Sub
        Case "5J"
            rdata = Right$(rdata, Len(rdata) - 2)
            UserMinAGU = Val(ReadFieldOptimizado(1, rdata, 44))
            UserMinHAM = Val(ReadFieldOptimizado(2, rdata, 44))
            frmPrincipal.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 94)
            frmPrincipal.cantidadagua.Caption = UserMinAGU & "/" & UserMaxAGU
            frmPrincipal.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 94)
            frmPrincipal.cantidadhambre.Caption = UserMinHAM & "/" & UserMaxHAM

            Exit Sub
        Case "5O"
            rdata = Right$(rdata, Len(rdata) - 2)
            UserLvl = Val(ReadFieldOptimizado(1, rdata, 44))
            UserPasarNivel = Val(ReadFieldOptimizado(2, rdata, 44))
            If UserPasarNivel > 0 Then
                frmPrincipal.LvlLbl.Caption = UserLvl & " (" & Round(UserExp / UserPasarNivel * 100, 2) & "%)" & " - " & PonerPuntos(UserExp) & " / " & PonerPuntos(UserPasarNivel)
             
            Else
             
            End If
            Exit Sub
        Case "HO"
            rdata = Right$(rdata, Len(rdata) - 2)
            X = Val(ReadFieldOptimizado(2, rdata, 44))
            Y = Val(ReadFieldOptimizado(3, rdata, 44))
            
            MapData(X, Y).ObjGrh.GrhIndex = Val(ReadFieldOptimizado(1, rdata, 44))
            MapData(X, Y).ObjGrh.name = ReadFieldOptimizado(4, rdata, 44)
            InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
            LastPos.X = X
            LastPos.Y = Y
            Exit Sub
         Case "HE"
            rdata = Right$(rdata, Len(rdata) - 2)
            Call FrmOpciones.DibujarGrhPorMapa(rdata)
            Exit Sub
        Case "P8"
            UserParalizado = False
            AddtoRichTextBox frmPrincipal.rectxt, "Ya no estás paralizado.", 65, 190, 156, 0, 0
            Exit Sub
        Case "P9"
            UserParalizado = True
            Call SendData("RPU")
            AddtoRichTextBox frmPrincipal.rectxt, "Estás paralizado. No podrás moverte por algunos segundos.", 65, 190, 156, 0, 0
            Exit Sub
        Case "BO"
            rdata = Right$(rdata, Len(rdata) - 2)
            X = Val(ReadFieldOptimizado(1, rdata, 44))
            Y = Val(ReadFieldOptimizado(2, rdata, 44))
            MapData(X, Y).ObjGrh.GrhIndex = 0
            MapData(X, Y).ObjGrh.name = ""
            Exit Sub
        Case "BQ"
            rdata = Right$(rdata, Len(rdata) - 2)
            MapData(Val(ReadFieldOptimizado(1, rdata, 44)), Val(ReadFieldOptimizado(2, rdata, 44))).Blocked = Val(ReadFieldOptimizado(3, rdata, 44))
            Exit Sub
        Case "BK"
            rdata = Right$(rdata, Len(rdata) - 2)
            Call FrmOpciones.BloquearTodoBQ(rdata)
            Exit Sub
        Case "TN"
            If Musica = 0 Then
                rdata = Right$(rdata, Len(rdata) - 2)
                If Val(ReadFieldOptimizado(1, rdata, 45)) <> 0 Then
                   ' mciExecute "Close All"
                    CurMidi = Val(ReadFieldOptimizado(1, rdata, 45)) & ".mid"
                    LoopMidi = Val(ReadFieldOptimizado(2, rdata, 45))
                    'Call Audio.PlayMIDI(App.Path & "\musicas\" & CurMidi)
                End If
            End If
            Exit Sub
        Case "TM"
                If Musica = 0 Then
                rdata = Right$(rdata, Len(rdata) - 2)
                
                If Val(ReadFieldOptimizado(1, rdata, 45)) = 0 Then Exit Sub
               ' mciExecute "Close All"
               ' Call Audio.StopMidi
               ' Call Audio.PlayWave(1, Val(ReadFieldOptimizado(1, Rdata, 45)) & ".mp3")
                End If
            Exit Sub
          
            Exit Sub
        Case "LH"
            LastHechizo = Timer
            Exit Sub
        Case "LG"
            LastGolpe = Timer
            Exit Sub
        Case "LF"
            LastFlecha = Timer
            Exit Sub
        Case "TW"
            If FX = 0 Then
                rdata = Right$(rdata, Len(rdata) - 2)
                 Call Audio.PlayWave(0, rdata & ".wav")
            End If
            Exit Sub
        Case "TX"
            Dim Efecto As Integer
            Dim ParticleCasteada As Integer
            rdata = Right$(rdata, Len(rdata) - 2)
            CharIndex = Val(ReadFieldOptimizado(1, rdata, 44)) 'atacante
            Charindexx = Val(ReadFieldOptimizado(2, rdata, 44)) 'victima
            Efecto = Val(ReadFieldOptimizado(4, rdata, 44)) 'efecto particulas
            If FX = 0 Then
                 Call Audio.PlayWave(0, ReadFieldOptimizado(6, rdata, 44) & ".wav")
            End If
            If Efecto = 0 Then
                CharList(Charindexx).FX = Val(ReadFieldOptimizado(3, rdata, 44))
                CharList(Charindexx).FxLoopTimes = Val(ReadFieldOptimizado(5, rdata, 44))
            End If
            If HechizAc = 0 Then 'si está activado
                ParticleCasteada = Engine_UTOV_Particle(CharIndex, Charindexx, Efecto)
            Else
                CharList(Charindexx).FX = Val(ReadFieldOptimizado(3, rdata, 44))
                CharList(Charindexx).FxLoopTimes = Val(ReadFieldOptimizado(5, rdata, 44))
            End If
            Exit Sub
        Case "GL"
            rdata = Right$(rdata, Len(rdata) - 2)
            frmGuildAdm.guildslist.Clear
            Call frmGuildAdm.ParseGuildList(rdata)
            frmGuildAdm.SetFocus
            Exit Sub
        Case "FO"
            bFogata = True
            
                If frmPrincipal.IsPlaying <> plFogata Then
                    Audio.StopWave
                    Call Audio.PlayWave(0, "fuego.wav", True)
                    frmPrincipal.IsPlaying = plFogata
                End If
            
            Exit Sub
End Select

End Sub
Public Function ReplaceData(sdData As String) As String
Dim rdata As String

If UCase$(Left$(sdData, 9)) = "/PASSWORD" Then
    frmCambiarPasswd.Show , frmPrincipal
    ReplaceData = "NOPUDO"
    Exit Function
End If

Select Case UCase$(sdData)
    Case Is = "/MEDITAR"
        ReplaceData = "#A"
    Case Is = "/SALIR"
        ReplaceData = "#B"
    Case "/FUNDARCLAN"
        Fundacion.Show , frmPrincipal
    Case "/BALANCE"
        ReplaceData = "#G"
    Case "/QUIETO"
        ReplaceData = "#H"
    Case "/ACOMPAÑAR"
        ReplaceData = "#I"
    Case "/ENTRENAR"
        ReplaceData = "#J"
    Case "/DESCANSAR"
        ReplaceData = "#K"
    Case "/RESUCITAR"
        ReplaceData = "#L"
    Case "/CURAR"
        ReplaceData = "#M"
    Case "/ONLINE"
        ReplaceData = "#P"
    Case "/VOTSI"
        ReplaceData = "VSI"
    Case "/VOTNO"
        ReplaceData = "VNO"
    Case "/IGNORADOS"
        Call MostrarIgnorados
        ReplaceData = "NOPUDO"
        Exit Function
    Case "/EST"
        ReplaceData = "#Q"
    Case "/PENA"
        ReplaceData = "#R"
    Case "/MOVER"
        ReplaceData = "#S"
    Case "/PARTICIPAR"
        ReplaceData = "#T"
    Case "/PROTECTOR1"
      ReplaceData = "#("
    Case "/TEAM1"
        ReplaceData = "#,"
    Case "/PROTECTOR2"
        ReplaceData = "#)"
    Case "/TEAM2"
        ReplaceData = "#%"
    Case "/ATRAPADO"
        ReplaceData = "#U"
    Case "/COMERCIAR"
        ReplaceData = "#V"
    Case "/BOVEDA"
        ReplaceData = "#W"
    Case "/HABLAR"
        ReplaceData = "#X"
    Case "/ENLISTAR"
        ReplaceData = "#Y"
    Case "/RECOMPENSA"
        ReplaceData = "#1"
    Case "/SALIRCLAN"
        ReplaceData = "#2"
    Case "/ONLINECLAN"
        ReplaceData = "#3"
    Case "/ABANDONAR"
        ReplaceData = "#4"
    Case "/RETARCLAN"
        ReplaceData = "#^"
    Case "/ACEPTCLAN"
        ReplaceData = "#¨"
    Case "/SEGUROCLAN"
        ReplaceData = "#€"
End Select

Select Case UCase$(Left$(sdData, 6))
    Case "/DESC "
        rdata = Right$(sdData, Len(sdData) - 6)
        ReplaceData = "#5 " & rdata
    Case "/VOTO "
        rdata = Right$(sdData, Len(sdData) - 6)
        ReplaceData = "#6 " & rdata
    Case "/CMSG "
        rdata = Right$(sdData, Len(sdData) - 6)
        ReplaceData = "#7 " & rdata
End Select
        
Select Case UCase$(Left$(sdData, 8))
    Case "/PASSWD "
        rdata = Right$(sdData, Len(sdData) - 8)
        ReplaceData = "#8 " & rdata
    Case "/ONLINE "
        rdata = Right$(sdData, Len(sdData) - 8)
        ReplaceData = "#*" & rdata
End Select

Select Case UCase$(Left$(sdData, 9))
    Case "/APOSTAR "
        rdata = Right$(sdData, Len(sdData) - 9)
        ReplaceData = "#9 " & rdata
    Case "/RETIRAR "
        rdata = Right$(sdData, Len(sdData) - 9)
        ReplaceData = "#0 " & rdata
 '   Case "/IGNORAR "
 '       Rdata = Right$(sdData, Len(sdData) - 9)
 '       Select Case IgnorarPJ(Rdata)
 '          Case 0
 '               ReplaceData = "NOPUDO"
 '               Exit Function
  '          Case 1
  '              ReplaceData = "#/ " & Rdata & " 1"
  '          Case 2
  '              ReplaceData = "#/ " & Rdata & " 0"
 '       End Select
End Select

Select Case UCase$(Left$(sdData, 11))
    Case "/DEPOSITAR "
        rdata = Right$(sdData, Len(sdData) - 11)
        ReplaceData = "#Ñ " & rdata
    Case "/DENUNCIAR "
        rdata = Right$(sdData, Len(sdData) - 11)
        ReplaceData = "^A " & rdata
End Select

If Len(ReplaceData) = 0 Then ReplaceData = sdData

End Function
Function KeyName(key As String) As String
Dim KeyCode As Byte

KeyCode = Asc(key)

Select Case KeyCode
    Case vbKeyAdd: KeyName = "+ (KeyPad)"
    Case vbKeyBack: KeyName = "Delete"
    Case vbKeyCancel: KeyName = "Cancelar"
    Case vbKeyCapital: KeyName = "CapsLock"
    Case vbKeyClear: KeyName = "Borrar"
    Case vbKeyControl: KeyName = "Control"
    Case vbKeyDecimal: KeyName = ". (KeyPad)"
    Case vbKeyDelete: KeyName = "Suprimir"
    Case vbKeyDivide: KeyName = "/ (KeyPad)"
    Case vbKeyEnd: KeyName = "Fin"
    Case vbKeyEscape: KeyName = "Esc"
    Case vbKeyF1: KeyName = "F1"
    Case vbKeyF2: KeyName = "F2"
    Case vbKeyF3: KeyName = "F3"
    Case vbKeyF4: KeyName = "F4"
    Case vbKeyF5: KeyName = "F5"
    Case vbKeyF6: KeyName = "F6"
    Case vbKeyF7: KeyName = "F7"
    Case vbKeyF8: KeyName = "F8"
    Case vbKeyF9: KeyName = "F9"
    Case vbKeyF10: KeyName = "F10"
    Case vbKeyF11: KeyName = "F11"
    Case vbKeyF12: KeyName = "F12"
    Case vbKeyF13: KeyName = "F13"
    Case vbKeyF14: KeyName = "F14"
    Case vbKeyF15: KeyName = "F15"
    Case vbKeyF16: KeyName = "F16"
    Case vbKeyHome: KeyName = "Inicio"
    Case vbKeyInsert: KeyName = "Insert"
    Case vbKeyMenu: KeyName = "Alt"
    Case vbKeyMultiply: KeyName = "* (KeyPad)"
    Case vbKeyNumlock: KeyName = "NumLock"
    Case vbKeyNumpad0: KeyName = "0 (KeyPad)"
    Case vbKeyNumpad1: KeyName = "1 (KeyPad)"
    Case vbKeyNumpad2: KeyName = "2 (KeyPad)"
    Case vbKeyNumpad3: KeyName = "3 (KeyPad)"
    Case vbKeyNumpad4: KeyName = "4 (KeyPad)"
    Case vbKeyNumpad5: KeyName = "5 (KeyPad)"
    Case vbKeyNumpad6: KeyName = "6 (KeyPad)"
    Case vbKeyNumpad7: KeyName = "7 (KeyPad)"
    Case vbKeyNumpad8: KeyName = "8 (KeyPad)"
    Case vbKeyNumpad9: KeyName = "9 (KeyPad)"
    Case vbKeyPageDown: KeyName = "Av Pag"
    Case vbKeyPageUp: KeyName = "Re Pag"
    Case vbKeyPause: KeyName = "Pausa"
    Case vbKeyPrint: KeyName = "ImprPant"
    Case vbKeyReturn: KeyName = "Enter"
    Case vbKeySelect: KeyName = "Select"
    Case vbKeyShift: KeyName = "Shift"
    Case vbKeySnapshot: KeyName = "Snapshot"
    Case vbKeySpace: KeyName = "Espacio"
    Case vbKeySubtract: KeyName = "- (KeyPad)"
    Case vbKeyTab: KeyName = "Tab"
    Case 92: KeyName = "Windows"
    Case 93: KeyName = "Lista"
    Case 145: KeyName = "Bloq Despl"
    Case 186: KeyName = "Comilla cierra(´)"
    Case 187: KeyName = "Asterisco (*)"
    Case 188: KeyName = "Coma (,)"
    Case 189: KeyName = "Guión (-)"
    Case 190: KeyName = "Punto (.)"
    Case 191: KeyName = "Llave cierra (})"
    Case 192: KeyName = "Ñ"
    Case 219: KeyName = "Comilla ("
    Case 220: KeyName = "Barra vertical (|)"
    Case 221: KeyName = "Signo pregunta (¿)"
    Case 222: KeyName = "Llave abre ({)"
    Case 223: KeyName = "Cualquiera"
    Case 226: KeyName = "Menor (<)"
    Case Else: KeyName = Chr(KeyCode)
End Select

End Function
Public Sub MostrarIgnorados()
Dim I As Integer

For I = 1 To UBound(Ignorados)
    If Ignorados(I) <> "" Then Call AddtoRichTextBox(frmPrincipal.rectxt, Ignorados(I), 65, 190, 156, 0, 0)
Next

End Sub
Function IgnorarPJ(name As String) As Byte
Dim I As Integer, tIndex As Integer

tIndex = NameIndex(name)

If tIndex = 0 Then
    Call AddtoRichTextBox(frmPrincipal.rectxt, "El personaje no existe o no está en tu mapa.", 65, 190, 156, 0, 0)
    Exit Function
End If

If tIndex = UserCharIndex Then
    Call AddtoRichTextBox(frmPrincipal.rectxt, "No puedes ignorarte a ti mismo.", 65, 190, 156, 0, 0)
    Exit Function
End If

For I = LBound(Ignorados) To UBound(Ignorados)
    If UCase$(Ignorados(I)) = UCase$(CharList(tIndex).Nombre) Then
        Call AddtoRichTextBox(frmPrincipal.rectxt, "Dejaste de ignorar a " & CharList(tIndex).Nombre & ".", 65, 190, 156, 0, 0)
        Ignorados(I) = ""
        IgnorarPJ = 2
        Exit Function
    End If
Next

For I = LBound(Ignorados) To UBound(Ignorados)
    If Len(Ignorados(I)) = 0 Then
        Call AddtoRichTextBox(frmPrincipal.rectxt, "Empezaste a ignorar a " & CharList(tIndex).Nombre & ".", 65, 190, 156, 0, 0)
        Ignorados(I) = CharList(tIndex).Nombre
        IgnorarPJ = 1
        Exit Function
    End If
Next

Call AddtoRichTextBox(frmPrincipal.rectxt, "No puedes ignorar a más personas.", 65, 190, 156, 0, 0)

End Function
Function NameIndex(name As String) As Integer
Dim I As Integer

For I = 1 To LastChar
    If UCase$(Left$(CharList(I).Nombre, Len(name))) = UCase$(name) Then
        NameIndex = I
        Exit Function
    End If
Next

End Function
Sub SendData(sdData As String)
Dim retcode
Dim AuxCmd As String

If Pausa Then Exit Sub

If CONGELADO And UCase$(sdData) <> "/DESCONGELAR" Then Exit Sub
If Not frmPrincipal.Socket1.Connected Then Exit Sub

AuxCmd = UCase$(Left$(sdData, 5))
If AuxCmd = "/PING" Then TimerPing(1) = GetTickCount()

Debug.Print ">> " & sdData

If Left$(sdData, 1) = "/" And Len(sdData) = 2 Then Exit Sub

sdData = ReplaceData(sdData)

If sdData = "NOPUDO" Then Exit Sub

bO = bO + 1
If bO > 10000 Then bO = 100

If Len(sdData) = 0 Then Exit Sub

If AuxCmd = "DEMSG" And Len(sdData) > 8000 Then Exit Sub
If AuxCmd = "GM" And Len(sdData) > 2200 Then
    NoMandoElMsg = 1
    Exit Sub
End If

If Len(sdData) > 300 And AuxCmd <> "DEMSG" And AuxCmd <> "GM" Then Exit Sub

NoMandoElMsg = 0

bK = 0

sdData = sdData & "~" & bK & ENDC

retcode = frmPrincipal.Socket1.Write(sdData, Len(sdData))
If PACKETS.Visible = False Then AddtoRichTextBox PACKETS.RichTextBox2, sdData, 255, 255, 255, 0, 0

End Sub
Sub Login(ByVal valcode As Integer)

If EstadoLogin = Normal Then
        Call SendData("JHUMPH" & UserName & "," & UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode & "," & GetMACAddress() & "," & GetSerialNumber("c:\") & "," & MotherBoardNumber)
ElseIf EstadoLogin = CrearNuevoPj Then
        SendData ("SARAXA" & UserName & "," & UserPassword _
        & "," & 0 & "," & 0 & "," _
        & App.Major & "." & App.Minor & "." & App.Revision & _
        "," & UserRaza & "," & UserSexo & "," & _
        UserAtributos(1) & "," & UserAtributos(2) & "," & UserAtributos(3) _
        & "," & UserAtributos(4) & "," & UserAtributos(5) _
         & "," & UserSkills(1) & "," & UserSkills(2) _
         & "," & UserSkills(3) & "," & UserSkills(4) _
         & "," & UserSkills(5) & "," & UserSkills(6) _
         & "," & UserSkills(7) & "," & UserSkills(8) _
         & "," & UserSkills(9) & "," & UserSkills(10) _
         & "," & UserSkills(11) & "," & UserSkills(12) _
         & "," & UserSkills(13) & "," & UserSkills(14) _
         & "," & UserSkills(15) & "," & UserSkills(16) _
         & "," & UserSkills(17) & "," & UserSkills(18) _
         & "," & UserSkills(19) & "," & UserSkills(20) _
         & "," & UserSkills(21) & "," & UserSkills(22) & "," & _
         UserEmail & "," & UserHogar & "," & valcode & "," & GetMACAddress() & "," & GetSerialNumber("c:\") & "," & MotherBoardNumber)
End If

End Sub
