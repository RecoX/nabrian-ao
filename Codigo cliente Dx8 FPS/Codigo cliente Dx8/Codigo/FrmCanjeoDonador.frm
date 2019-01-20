VERSION 5.00
Begin VB.Form FrmCanjeoDonador 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   5550
      Left            =   255
      TabIndex        =   1
      Top             =   900
      Width           =   2730
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   585
      Left            =   4245
      ScaleHeight     =   585
      ScaleWidth      =   645
      TabIndex        =   0
      Top             =   1050
      Width           =   645
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ninguno de estos item de donador se cae al morir."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label lblPrecio 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   4440
      Width           =   2640
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label lblPermisos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   3120
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "FrmCanjeoDonador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Me.PICTURE = LoadPicture(App.Path & "\graficos\canjeodonador.gif")
List1.AddItem "Túnica dorada (ENANO)"
List1.AddItem "Túnica dorada (HUMANO)"
List1.AddItem "Túnica Suprema (ENANO)"
List1.AddItem "Túnica Suprema (HUMANO)"
List1.AddItem "Armadura Durlock (H)"
List1.AddItem "Armadura Infernal (E)"
List1.AddItem "Sombrero Infernal"
List1.AddItem "Sombrero Blanco"
List1.AddItem "Espada Ardiente"
List1.AddItem "Espada de Hielo"
List1.AddItem "Escudo Legendario"
List1.AddItem "Escudo Rondor"
List1.AddItem "Daga Acta"
List1.AddItem "Baculo Ancestral"
List1.AddItem "Arco Infernal"
List1.AddItem "Túnica alada (H/E)"
List1.AddItem "Montura de dragon amarillo"
List1.AddItem "Montura de dragon rojo"
List1.AddItem "Montura de corsario blanco"
List1.AddItem "Montura de corsario negro"
List1.AddItem "Gema de los Dioses"
List1.AddItem "Anillo de los Dioses Templarios"
List1.AddItem "150 Puntos de Canjeo"
List1.AddItem "350 Puntos de Canjeo"
List1.AddItem "800 Puntos de Canjeo"
End Sub



Private Sub Image2_Click()

If List1.Text = "Túnica dorada (ENANO)" Then Call SendData("/DEEW T1")
If List1.Text = "Túnica dorada (HUMANO)" Then Call SendData("/DEEW T2")
If List1.Text = "Túnica Suprema (ENANO)" Then Call SendData("/DEEW T3")
If List1.Text = "Túnica Suprema (HUMANO)" Then Call SendData("/DEEW T4")
If List1.Text = "Armadura Durlock (H)" Then Call SendData("/DEEW T5")
If List1.Text = "Armadura Infernal (E)" Then Call SendData("/DEEW T6")
If List1.Text = "Sombrero Infernal" Then Call SendData("/DEEW T7")
If List1.Text = "Sombrero Blanco" Then Call SendData("/DEEW T8")
If List1.Text = "Espada Ardiente" Then Call SendData("/DEEW T9")
If List1.Text = "Espada de Hielo" Then Call SendData("/DEEW T10")
If List1.Text = "Escudo Legendario" Then Call SendData("/DEEW T11")
If List1.Text = "Escudo Rondor" Then Call SendData("/DEEW T12")
If List1.Text = "Daga Acta" Then Call SendData("/DEEW T13")
If List1.Text = "Baculo Ancestral" Then Call SendData("/DEEW T14")
If List1.Text = "Arco Infernal" Then Call SendData("/DEEW T15")
If List1.Text = "Túnica alada (H/E)" Then Call SendData("/DEEW T16")
If List1.Text = "Montura de dragon amarillo" Then Call SendData("/DEEW T17")
If List1.Text = "Montura de dragon rojo" Then Call SendData("/DEEW T18")
If List1.Text = "Montura de corsario blanco" Then Call SendData("/DEEW T19")
If List1.Text = "Montura de corsario negro" Then Call SendData("/DEEW T20")
If List1.Text = "Gema de los Dioses" Then Call SendData("/DEEW T21")
If List1.Text = "Anillo de los Dioses Templarios" Then Call SendData("/DEEW T22")
If List1.Text = "150 Puntos de Canjeo" Then Call SendData("/DEEW T23")
If List1.Text = "350 Puntos de Canjeo" Then Call SendData("/DEEW T24")
If List1.Text = "800 Puntos de Canjeo" Then Call SendData("/DEEW T25")
Unload Me
End Sub

Private Sub Label1_Click()
Unload FrmCanjeoDonador
End Sub




Private Sub List1_Click()
Dim PICTURE As String
If List1.Text = "Túnica dorada (ENANO)" Then
    PICTURE = "804.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "150 Puntos de donador."
    lblStat.Caption = "Min: 45 / Max: 50"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Túnica dorada (HUMANO)" Then
    PICTURE = "804.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "150 Puntos de donador."
    lblStat.Caption = "Min: 45 / Max: 50"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Túnica Suprema (ENANO)" Then
    PICTURE = "662.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "100 Puntos de donador."
    lblStat.Caption = "Min: 43 / Max: 48"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Túnica Suprema (HUMANO)" Then
    PICTURE = "662.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "100 Puntos de donador."
    lblStat.Caption = "Min: 43 / Max: 48"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Armadura Durlock (H)" Then
    PICTURE = "16181.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "250 Puntos de donador."
    lblStat.Caption = "Min: 50 / Max: 55"
    lblPermisos.Caption = "Paladin - Guerrero - Clerigo - Arquero"
    End If
If List1.Text = "Armadura Infernal (E)" Then
    PICTURE = "10113.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "250 Puntos de donador."
    lblStat.Caption = "Min: 50 / Max: 55"
    lblPermisos.Caption = "Paladin - Guerrero - Clerigo - Arquero"
    End If
If List1.Text = "Sombrero Infernal" Then
    PICTURE = "16032.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "300 Puntos de donador."
    lblStat.Caption = "Min: 17 / Max: 19"
    lblPermisos.Caption = "Mago"
    End If
If List1.Text = "Sombrero Blanco" Then
    PICTURE = "16102.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "100 Puntos de donador."
    lblStat.Caption = "Min: 8 / Max: 8"
    lblPermisos.Caption = "Bardo - Clerigo - Asesino - Mago"
    End If
If List1.Text = "Espada Ardiente" Then
    PICTURE = "9629.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "300 Puntos de donador."
    lblStat.Caption = "Min: 21 / Max: 22"
    lblPermisos.Caption = "Paladin - Guerrero"
    End If
If List1.Text = "Espada de Hielo" Then
    PICTURE = "16072.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "300 Puntos de donador."
    lblStat.Caption = "Min: 18 / Max: 20"
    lblPermisos.Caption = "Clerigo - Paladin - Guerrero"
    End If
If List1.Text = "Escudo Legendario" Then
    PICTURE = "9574.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "150 Puntos de donador."
    lblStat.Caption = "Min: 9 / Max: 14"
    lblPermisos.Caption = "Clerigo - Paladin - Guerrero"
    End If
If List1.Text = "Escudo Rondor" Then
    PICTURE = "16122.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "150 Puntos de donador."
    lblStat.Caption = "Min: 8 / Max: 11"
    lblPermisos.Caption = "Bardo - Asesino"
    End If
If List1.Text = "Daga Acta" Then
    PICTURE = "3537.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "300 Puntos de donador."
    lblStat.Caption = "Min: 11 / Max: 12"
    lblPermisos.Caption = "Bardo - Asesino"
    End If
If List1.Text = "Baculo Ancestral" Then
    PICTURE = "16030.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "300 Puntos de donador."
    lblStat.Caption = "Min: 8 / Max: 11"
    lblPermisos.Caption = "Mago"
    End If
If List1.Text = "Arco Infernal" Then
    PICTURE = "16116.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "300 Puntos de donador."
    lblStat.Caption = "Min: 11 / Max: 16"
    lblPermisos.Caption = "Arquero"
    End If
If List1.Text = "Túnica alada (H/E)" Then
    PICTURE = "16465.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "500 Puntos de donador."
    lblStat.Caption = "Min: 52 / Max / 58"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Montura de dragon amarillo" Then
    PICTURE = "16460.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "750 Puntos de donador."
    lblStat.Caption = "Min: 55 / Max / 60"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Montura de dragon rojo" Then
    PICTURE = "14502.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "750 Puntos de donador."
    lblStat.Caption = "Min: 55 / Max / 60"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Montura de corsario blanco" Then
    PICTURE = "13169.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "400 Puntos de donador."
    lblStat.Caption = "Min: 50 / Max / 52"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Montura de corsario negro" Then
    PICTURE = "13167.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "400 Puntos de donador."
    lblStat.Caption = "Min: 50 / Max / 52"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Gema de los Dioses" Then
    PICTURE = "705.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "250 Puntos de donador."
    lblStat.Caption = "Fundacion de Clan Sin Requerimientos"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Anillo de los Dioses Templarios" Then
    PICTURE = "16239.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "250 Puntos de donador."
    lblStat.Caption = "Templario sin realizar las Misiones de THOR."
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "150 Puntos de Canjeo" Then
    PICTURE = "3.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "50 Puntos de donador."
    lblStat.Caption = "Te otorga 100 puntos de canjeo."
    lblPermisos.Caption = ""
    End If
If List1.Text = "350 Puntos de Canjeo" Then
    PICTURE = "3.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "100 Puntos de donador."
    lblStat.Caption = "Te otorga 350 puntos de canjeo."
    lblPermisos.Caption = ""
    End If
If List1.Text = "800 Puntos de Canjeo" Then
    PICTURE = "3.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "200 Puntos de donador."
    lblStat.Caption = "Te otorga 800 puntos de canjeo."
    lblPermisos.Caption = ""
    End If
    
    If EncriptGraficosActiva = True Then
    If Extract_File(Graphics, App.Path & "\GRAFICOS\", PICTURE, App.Path & "\GRAFICOS\") Then
        Picture1.PICTURE = LoadPicture(DirGraficos & PICTURE)
        Call Kill(App.Path & "\Graficos\*.bmp")
    End If
    Else
    Picture1.PICTURE = LoadPicture(DirGraficos & PICTURE)
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If bmoving = False And Button = vbLeftButton Then
      Dx3 = X
      dy = Y
      bmoving = True
   End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If bmoving And ((X <> Dx3) Or (Y <> dy)) Then
      Move Left + (X - Dx3), Top + (Y - dy)
   End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      bmoving = False
   End If
End Sub

