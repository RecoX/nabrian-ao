VERSION 5.00
Begin VB.Form Forjador 
   BorderStyle     =   0  'None
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   Picture         =   "Forjador.frx":0000
   ScaleHeight     =   5565
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
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
      Height          =   1335
      Left            =   4200
      ScaleHeight     =   1335
      ScaleWidth      =   3015
      TabIndex        =   5
      Top             =   2520
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   4020
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2715
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
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
      Height          =   735
      Left            =   3000
      ScaleHeight     =   735
      ScaleWidth      =   840
      TabIndex        =   0
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ninguno de estos item se cae en las afueras del mundo."
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
      Height          =   1095
      Left            =   3000
      TabIndex        =   7
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label LabelcLases 
      BackStyle       =   0  'Transparent
      Caption         =   "Clases: -------------------------------------"
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
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   7080
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ataque: N/A"
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
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Defensa: N/A"
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
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre: Selecciona un item."
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
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   4200
      Top             =   4080
      Width           =   3015
   End
End
Attribute VB_Name = "Forjador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Sub Image1_Click()
Unload Me
End Sub
Private Sub Form_Load()
List1.AddItem "Espada Infernal Reforzada"
List1.AddItem "Daga Infernal Reforzada"
List1.AddItem "Espada Celestial"
List1.AddItem "Báculo Infernal"
End Sub

Private Sub Image2_Click()
If List1.Text = "Espada Infernal Reforzada" Then Call SendData("/FORJE 1")
If List1.Text = "Daga Infernal Reforzada" Then Call SendData("/FORJE 2")
If List1.Text = "Espada Celestial" Then Call SendData("/FORJE 3")
If List1.Text = "Báculo Infernal" Then Call SendData("/FORJE 4")
Unload Me
End Sub





Private Sub List1_Click()
Picture2.Cls
If List1.Text = "Espada Infernal Reforzada" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16575.bmp")
    lblNombre.Caption = "Nombre:" & List1.Text
    Label1.Caption = "Defensa: N/A"
    Label2.Caption = "Ataque 21/23"
    LabelcLases.Caption = "Clases: Paladin - Guerrero."
    Picture2.Print "Espada Fantasmal"
    Picture2.Print "Tunica de Almas (Hades)"
    Picture2.Print "Armadura Infernal (Hades)"
    Picture2.Print "Gema Encantadora (Atenea)"
    End If
If List1.Text = "Daga Infernal Reforzada" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16579.bmp")
    lblNombre.Caption = "Nombre:" & List1.Text
    Label1.Caption = "Defensa: N/A"
    Label2.Caption = "Ataque 11/12"
    LabelcLases.Caption = "Clases: Bardo - Asesino."
    Picture2.Print "Daga Templaria"
    Picture2.Print "Armadura Daedrica (Hades)"
    Picture2.Print "Tunica Lagred (Atenea)"
    Picture2.Print "Gema Encantadora (Atenea)"
    End If
If List1.Text = "Espada Celestial" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16577.bmp")
    lblNombre.Caption = "Nombre:" & List1.Text
    Label1.Caption = "Defensa: N/A"
    Label2.Caption = "Ataque 18/21"
    LabelcLases.Caption = "Clases: Clerigo - Paladin - Guerrero."
    Picture2.Print "Espada de Plata + 1"
    Picture2.Print "Daga Helada (Poseidon)"
    Picture2.Print "Corona Celestial (Poseidon)"
    Picture2.Print "Gema Encantadora (Atenea)"
    End If
If List1.Text = "Báculo Infernal" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16569.bmp")
    lblNombre.Caption = "Nombre:" & List1.Text
    Label1.Caption = "Defensa: N/A"
    Label2.Caption = "Ataque N/A"
    LabelcLases.Caption = "Clases: Mago."
    Picture2.Print "Báculo Sagrado"
    Picture2.Print "Corona de Atenas (Atenea)"
    Picture2.Print "Corona Celestial (Poseidon)"
    Picture2.Print "Gema Encantadora (Atenea)"
    End If
End Sub


