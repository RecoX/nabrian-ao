VERSION 5.00
Begin VB.Form frmInterfases 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Gestor de Interfases"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Elegir Interfase"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Interfases Disponibles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.OptionButton Option3 
         Caption         =   "Evacuation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Doran, The Siege Tower"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Augur of the Skulls"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   2880
      Top             =   1200
      Width           =   3135
   End
End
Attribute VB_Name = "frmInterfases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1 = True Then
frmMain.Picture = LoadPicture(DirGraficos & "AugurPrincipal.jpg")
frmMain.imgFondoInvent.Picture = LoadPicture(DirGraficos & "AugurInventario.jpg")
frmMain.imgFondoHechizos.Picture = LoadPicture(DirGraficos & "AugurHechizos.jpg")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 1)
End If

If Option2 = True Then
frmMain.Picture = LoadPicture(DirGraficos & "DoranPrincipal.jpg")
frmMain.imgFondoInvent.Picture = LoadPicture(DirGraficos & "DoranInventario.jpg")
frmMain.imgFondoHechizos.Picture = LoadPicture(DirGraficos & "DoranHechizos.jpg")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 2)
End If

If Option3 = True Then
frmMain.Picture = LoadPicture(DirGraficos & "EvacuationPrincipal.jpg")
frmMain.imgFondoInvent.Picture = LoadPicture(DirGraficos & "EvacuationInventario.jpg")
frmMain.imgFondoHechizos.Picture = LoadPicture(DirGraficos & "EvacuationHechizos.jpg")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 3)
End If

If Option4 = True Then
frmMain.Picture = LoadPicture(DirGraficos & "HandPrincipal.jpg")
frmMain.imgFondoInvent.Picture = LoadPicture(DirGraficos & "HandInventario.jpg")
frmMain.imgFondoHechizos.Picture = LoadPicture(DirGraficos & "HandHechizos.jpg")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 4)
End If

If Option5 = True Then
frmMain.Picture = LoadPicture(DirGraficos & "IslandPrincipal.jpg")
frmMain.imgFondoInvent.Picture = LoadPicture(DirGraficos & "IslandInventario.jpg")
frmMain.imgFondoHechizos.Picture = LoadPicture(DirGraficos & "IslandHechizos.jpg")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 5)
End If

If Option6 = True Then
frmMain.Picture = LoadPicture(DirGraficos & "LlanowarPrincipal.jpg")
frmMain.imgFondoInvent.Picture = LoadPicture(DirGraficos & "LlanowarInventario.jpg")
frmMain.imgFondoHechizos.Picture = LoadPicture(DirGraficos & "LlanowarHechizos.jpg")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 6)
End If

If Option7 = True Then
frmMain.Picture = LoadPicture(DirGraficos & "LuminousPrincipal.jpg")
frmMain.imgFondoInvent.Picture = LoadPicture(DirGraficos & "LuminousInventario.jpg")
frmMain.imgFondoHechizos.Picture = LoadPicture(DirGraficos & "LuminousHechizos.jpg")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 7)
End If

If Option8 = True Then
frmMain.Picture = LoadPicture(DirGraficos & "WanderbrinePrincipal.jpg")
frmMain.imgFondoInvent.Picture = LoadPicture(DirGraficos & "WanderbrineInventario.jpg")
frmMain.imgFondoHechizos.Picture = LoadPicture(DirGraficos & "WanderbrineHechizos.jpg")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 8)
End If

If Option9 = True Then
frmMain.Picture = LoadPicture(DirGraficos & "WatchtowerPrincipal.jpg")
frmMain.imgFondoInvent.Picture = LoadPicture(DirGraficos & "WatchtowerInventario.jpg")
frmMain.imgFondoHechizos.Picture = LoadPicture(DirGraficos & "WatchtowerHechizos.jpg")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 9)
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Image1.Stretch = True
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Option1_Click()
Image1.Picture = LoadPicture(DirGraficos & "AugurPrincipal.jpg")
End Sub

Private Sub Option2_Click()
Image1.Picture = LoadPicture(DirGraficos & "DoranPrincipal.jpg")
End Sub

Private Sub Option3_Click()
Image1.Picture = LoadPicture(DirGraficos & "EvacuationPrincipal.jpg")
End Sub

Private Sub Option4_Click()
Image1.Picture = LoadPicture(DirGraficos & "HandPrincipal.jpg")
End Sub

Private Sub Option5_Click()
Image1.Picture = LoadPicture(DirGraficos & "IslandPrincipal.jpg")
End Sub

Private Sub Option6_Click()
Image1.Picture = LoadPicture(DirGraficos & "LlanowarPrincipal.jpg")
End Sub

Private Sub Option7_Click()
Image1.Picture = LoadPicture(DirGraficos & "LuminousPrincipal.jpg")
End Sub

Private Sub Option8_Click()
Image1.Picture = LoadPicture(DirGraficos & "WanderbrinePrincipal.jpg")
End Sub

Private Sub Option9_Click()
Image1.Picture = LoadPicture(DirGraficos & "WatchtowerPrincipal.jpg")
End Sub
