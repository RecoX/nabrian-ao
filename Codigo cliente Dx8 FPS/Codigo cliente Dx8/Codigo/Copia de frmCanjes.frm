VERSION 5.00
Begin VB.Form frmCanjes 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Sistema de canje"
   ClientHeight    =   6585
   ClientLeft      =   420
   ClientTop       =   315
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Copia de frmCanjes.frx":0000
   ScaleHeight     =   6585
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000001&
      Height          =   735
      Left            =   240
      ScaleHeight     =   735
      ScaleWidth      =   840
      TabIndex        =   1
      Top             =   600
      Width           =   840
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000004&
      Height          =   4320
      Left            =   240
      TabIndex        =   0
      Top             =   1580
      Width           =   2600
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   4800
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   3120
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2280
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblPermisos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2445
      Width           =   1695
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
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
      TabIndex        =   4
      Top             =   3180
      Width           =   1695
   End
   Begin VB.Label lblPrecio 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   4000
      Width           =   1695
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Tunica de rey"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
List1.AddItem "Ropa de las sombras"
List1.AddItem "Ropa de la alianza"
List1.AddItem "Escudo Dinal +1"
List1.AddItem "Gorrito de Navidad"
End Sub

Private Sub Image1_Click()
Me.Visible = False
End Sub

Private Sub Image2_Click()

If List1.Text = "Ropa de las sombras" Then Call SendData("/CANJEOT T1")
If List1.Text = "Ropa de la alianza" Then Call SendData("/CANJEOT T2")
If List1.Text = "Escudo Dinal +1" Then Call SendData("/CANJEOT T3")
If List1.Text = "Gorrito de Navidad" Then Call SendData("/CANJEOT T4")

End Sub

Private Sub Image3_Click()
frmCanjes.Visible = False
frmCanjea.Show
End Sub

Private Sub Image4_Click()
Me.Visible = False
End Sub

Private Sub list1_Click()
If List1.Text = "Ropa de las sombras" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16036.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "35 Torneos "
    lblStat.Caption = "50/50"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Ropa de la alianza" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16038.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "35 Torneos"
    lblStat.Caption = "50/50"
    lblPermisos.Caption = "Todas las clases"
    End If
If List1.Text = "Escudo Dinal +1" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16064.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "20 Torneos"
    lblStat.Caption = "10/12"
    lblPermisos.Caption = "Bardo"
    End If
If List1.Text = "Gorrito de Navidad" Then
    Picture1.Picture = LoadPicture(DirGraficos & "2363.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "4 Torneos"
    lblStat.Caption = "20/25"
    lblPermisos.Caption = "Todas las clases"
    End If
End Sub

