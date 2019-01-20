VERSION 5.00
Begin VB.Form frmCanjea 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6840
   ClientLeft      =   420
   ClientTop       =   315
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4250
      ScaleHeight     =   585
      ScaleWidth      =   645
      TabIndex        =   1
      Top             =   1050
      Width           =   645
   End
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
      TabIndex        =   0
      Top             =   905
      Width           =   2730
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
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   3120
      Top             =   5880
      Width           =   2655
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
      TabIndex        =   5
      Top             =   3720
      Width           =   2535
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
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   3000
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
      Left            =   3120
      TabIndex        =   3
      Top             =   4440
      Width           =   2760
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
      TabIndex        =   2
      Top             =   2280
      Width           =   2775
   End
End
Attribute VB_Name = "frmCanjea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Me.PICTURE = LoadPicture(App.Path & "\graficos\canjeo.gif")
List1.AddItem "Tunica de Rey (Altos)"
List1.AddItem "Tunica de Rey (Bajos)"
List1.AddItem "Ropa penitenciaria (altos)"
List1.AddItem "Ropa penitenciaria (Bajos)"
List1.AddItem "Sombrero de las sombras"
List1.AddItem "Baculo Sagrado"
List1.AddItem "Arco de la luz"
List1.AddItem "Espada de plata +1"
List1.AddItem "Espada fantasmal"
List1.AddItem "Escudo plateado"
List1.AddItem "Corona comun"
List1.AddItem "Corona de rey"
List1.AddItem "Armadura daedrica"
List1.AddItem "Armadura de asesino"
List1.AddItem "Armadura de asesino (E)"
List1.AddItem "Armadura Thek"
List1.AddItem "Daga templaria"
List1.AddItem "Tunica drumond oscura"
'List1.AddItem "Chupin celeste"
'List1.AddItem "Chupin rojo"
'List1.AddItem "Chupin Amarillo"
List1.AddItem "Tunica Champions negra"
List1.AddItem "Daga de hielo"
List1.AddItem "Escudo de aura"
List1.AddItem "Sacri"
List1.AddItem "Amuleto"
End Sub



Private Sub Image2_Click()

If List1.Text = "Tunica de Rey (Altos)" Then Call SendData("/CANJEO T1")
If List1.Text = "Tunica de Rey (Bajos)" Then Call SendData("/CANJEO T2")
If List1.Text = "Ropa penitenciaria (altos)" Then Call SendData("/CANJEO T3")
If List1.Text = "Ropa penitenciaria (Bajos)" Then Call SendData("/CANJEO T4")
If List1.Text = "Sombrero de las sombras" Then Call SendData("/CANJEO T5")
If List1.Text = "Baculo Sagrado" Then Call SendData("/CANJEO T6")
If List1.Text = "Arco de la luz" Then Call SendData("/CANJEO T7")
If List1.Text = "Espada de plata +1" Then Call SendData("/CANJEO T8")
If List1.Text = "Espada fantasmal" Then Call SendData("/CANJEO T9")
If List1.Text = "Escudo plateado" Then Call SendData("/CANJEO T10")
If List1.Text = "Corona comun" Then Call SendData("/CANJEO T11")
If List1.Text = "Corona de rey" Then Call SendData("/CANJEO T12")
If List1.Text = "Armadura daedrica" Then Call SendData("/CANJEO T13")
If List1.Text = "Armadura de asesino" Then Call SendData("/CANJEO T14")
If List1.Text = "Armadura de asesino (E)" Then Call SendData("/CANJEOPS T4")
If List1.Text = "Armadura Thek" Then Call SendData("/CANJEO T15")
If List1.Text = "Daga templaria" Then Call SendData("/CANJEO T16")
If List1.Text = "Tunica drumond oscura" Then Call SendData("/CANJEOP T2")
'If List1.Text = "Chupin celeste" Then Call SendData("/CANJEOP T3")
'If List1.Text = "Chupin rojo" Then Call SendData("/CANJEOPS T1")
'If List1.Text = "Chupin Amarillo" Then Call SendData("/CANJEOPS T2")
If List1.Text = "Tunica Champions negra" Then Call SendData("/CANJEOP T1")
If List1.Text = "Daga de hielo" Then Call SendData("/CANJEOP T5")
If List1.Text = "Escudo de aura" Then Call SendData("/CANJEOP T6")
If List1.Text = "Sacri" Then Call SendData("/CANJEOPS T3")
If List1.Text = "Amuleto" Then Call SendData("/CANJEOPS T5")

Unload Me
End Sub




Private Sub Label1_Click()
Unload frmCanjea
End Sub






Private Sub List1_Click()
Dim PICTURE As String
If List1.Text = "Tunica de Rey (Altos)" Then
    PICTURE = "685.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "400 Puntos de Canjeo"
    lblStat.Caption = "Min: 35 / Max: 40"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Tunica de Rey (Bajos)" Then
    PICTURE = "12044.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "400 Puntos de Canjeo"
    lblStat.Caption = "Min: 35 / Max: 40"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Ropa penitenciaria (altos)" Then
    PICTURE = "1571.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "150 Puntos de Canjeo"
    lblStat.Caption = "Min: 10 / Max: 20"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Ropa penitenciaria (Bajos)" Then
    PICTURE = "1572.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "150 Puntos de Canjeo"
    lblStat.Caption = "Min: 10 / Max: 20"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Sombrero de las sombras" Then
    PICTURE = "16124.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "450 Puntos de Canjeo"
    lblStat.Caption = "Min: 15 / Max: 18"
    lblPermisos.Caption = "Mago"
    End If
If List1.Text = "Baculo Sagrado" Then
    PICTURE = "10063.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "500 Puntos de Canjeo"
    lblStat.Caption = "Min: 10 / Max: 10"
    lblPermisos.Caption = "Mago"
    End If
If List1.Text = "Arco de la luz" Then
    PICTURE = "16114.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "500 Puntos de Canjeo"
    lblStat.Caption = "Min: 10 / Max: 16"
    lblPermisos.Caption = "Arquero"
    End If
If List1.Text = "Espada de plata +1" Then
    PICTURE = "9627.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "500 Puntos de Canjeo"
    lblStat.Caption = "Min: 17 / Max: 20"
    lblPermisos.Caption = "Clerigo, paladín, Guerrero"
    End If
If List1.Text = "Espada fantasmal" Then
    PICTURE = "9630.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "500 Puntos de Canjeo"
    lblStat.Caption = "Min: 20 / Max: 22"
    lblPermisos.Caption = "Paladín y guerrero."
    End If
If List1.Text = "Escudo plateado" Then
    PICTURE = "16068.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "475 Puntos de Canjeo"
    lblStat.Caption = "Min: 8 / Max: 14"
    lblPermisos.Caption = "Clerigo y paladín"
    End If
If List1.Text = "Corona comun" Then
    PICTURE = "2023.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "800 Puntos de Canjeo"
    lblStat.Caption = "Min: 14 / Max: 16 'RM'"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Corona de rey" Then
    PICTURE = "16100.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "1200 Puntos de Canjeo"
    lblStat.Caption = "Min: 16 / Max: 19 'RM'"
    lblPermisos.Caption = "Todas menos guerrero."
    End If
If List1.Text = "Armadura daedrica" Then
    PICTURE = "16197.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "450 Puntos de Canjeo"
    lblStat.Caption = "Min: 50 / Max: 50"
    lblPermisos.Caption = "Paladín y Guerrero"
    End If
If List1.Text = "Armadura de asesino" Then
    PICTURE = "4591.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "425 Puntos de Canjeo"
    lblStat.Caption = "Min: 30 / Max: 40"
    lblPermisos.Caption = "Asesino"
    End If
If List1.Text = "Armadura de asesino (E)" Then
    PICTURE = "16076.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "425 Puntos de Canjeo"
    lblStat.Caption = "Min: 30 / Max: 40"
    lblPermisos.Caption = "Asesino"
    End If
If List1.Text = "Armadura Thek" Then
    PICTURE = "16048.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "550 Puntos de Canjeo"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Clerigo"
    End If
If List1.Text = "Daga templaria" Then
    PICTURE = "1222.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "500 Puntos de Canjeo"
    lblStat.Caption = "Min: 10 / Max: 12"
    lblPermisos.Caption = "Bardo"
    End If
If List1.Text = "Tunica drumond oscura" Then
    PICTURE = "16034.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "375 Puntos de Canjeo"
    lblStat.Caption = "Min: 30 / Max: 35"
    lblPermisos.Caption = "Mago y bardo"
    End If
'If List1.Text = "Chupin celeste" Then
'    PICTURE = "16106.bmp"
'    lblNombre.Caption = List1.Text
'    lblPrecio.Caption = "350 Puntos de Canjeo"
'    lblStat.Caption = "Min: 25 / Max: 30"
'    lblPermisos.Caption = "Todas las clases"
'    End If
'If List1.Text = "Chupin rojo" Then
'    PICTURE = "16104.bmp"
'    lblNombre.Caption = List1.Text
'    lblPrecio.Caption = "350 Puntos de Canjeo"
 '   lblStat.Caption = "Min: 25 / Max: 30"
 '   lblPermisos.Caption = "Todas las clases"
 '   End If
'If List1.Text = "Chupin Amarillo" Then
'    PICTURE = "16108.bmp"
'    lblNombre.Caption = List1.Text
'    lblPrecio.Caption = "350 Puntos de Canjeo"
'    lblStat.Caption = "Min: 25 / Max: 30"
'    lblPermisos.Caption = "Todas las clases"
'    End If
If List1.Text = "Tunica Champions negra" Then
    PICTURE = "16052.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "425 Puntos de Canjeo"
    lblStat.Caption = "Min: 40 / Max: 40"
    lblPermisos.Caption = "Bardo, Mago, Clerigo"
    End If
If List1.Text = "Daga de hielo" Then
    PICTURE = "16118.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "475 Puntos de Canjeo"
    lblStat.Caption = "Min: 10 / Max: 12"
    lblPermisos.Caption = "Asesino"
    End If
If List1.Text = "Escudo de aura" Then
    PICTURE = "16064.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "440 Puntos de Canjeo"
    lblStat.Caption = "Min: 7 / Max: 11"
    lblPermisos.Caption = "Bardo, Asesino"
    End If
If List1.Text = "Sacri" Then
    PICTURE = "16511.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "5 Puntos de Canjeo"
    lblStat.Caption = "(Tus item del inventario no caeran)"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Amuleto" Then
    PICTURE = "16477.bmp"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "20 Puntos de Canjeo"
    lblStat.Caption = "(Entrar a Ultratumba)"
    lblPermisos.Caption = "Todas las clases"
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
