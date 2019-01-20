VERSION 5.00
Begin VB.Form Menu 
   BorderStyle     =   0  'None
   Caption         =   "Menu"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "       Presiona esc para ocultar el menú"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Image Image8 
      Height          =   615
      Left            =   2880
      Top             =   240
      Width           =   2175
   End
   Begin VB.Image Image7 
      Height          =   375
      Left            =   360
      Top             =   4800
      Width           =   4695
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   360
      Top             =   240
      Width           =   2175
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   360
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   360
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   360
      Top             =   2760
      Width           =   4695
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   360
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   360
      Top             =   3240
      Width           =   4695
   End
   Begin VB.Image Ehh 
      Height          =   375
      Left            =   360
      Top             =   4320
      Width           =   4695
   End
   Begin VB.Image Image1R 
      Height          =   375
      Left            =   360
      Top             =   2160
      Width           =   4695
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\graficos\menu.jpg")
End Sub

Private Sub Image1_Click()
Call Audio.PlayWave(0, SND_CLICK)

        LlegaronAtrib = False
        LlegaronSkills = False
        LlegoFama = False
        LlegoMinist = False
        SendData "ATRI"
        SendData "ESKI"
        SendData "FAMA"
        Do While Not LlegaronAtrib Or Not LlegoFama Or Not LlegoMinist
            DoEvents
        Loop
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Show , frmPrincipal
        LlegaronAtrib = False
        LlegaronSkills = False
        LlegoFama = False
        LlegoMinist = False
        Unload Menu

End Sub



Private Sub Ehh_Click()
Call Audio.PlayWave(0, SND_CLICK)
frmParty.ListaIntegrantes.Clear
LlegoParty = False
Call SendData("PARINF")
Do While Not LlegoParty
    DoEvents
Loop
frmParty.Visible = True
frmParty.SetFocus
LlegoParty = False
Unload Menu
End Sub


Private Sub Image1R_Click()
Call Audio.PlayWave(0, SND_CLICK)
Call SendData("/RANKING")
frmRanking.Show , frmPrincipal
Unload Menu
End Sub

Private Sub Image2_Click()
Call Audio.PlayWave(0, SND_CLICK)
       If frmGuildLeader.Visible Then frmGuildLeader.Visible = False
        If frmGuildsNuevo.Visible Then frmGuildsNuevo.Visible = False
        If frmGuildAdm.Visible Then frmGuildAdm.Visible = False
        Call SendData("GLINFO")
        Unload Menu
End Sub

Private Sub Image3_Click()
Call Audio.PlayWave(0, SND_CLICK)
        Call FrmOpciones.Show(vbModeless, frmPrincipal)
        Unload Menu
End Sub

Private Sub Image4_Click()
Call Audio.PlayWave(0, SND_CLICK)
ModoCanjes = 0
frmCanjes.Label1.Visible = True
Call SendData("CCANJE")
ModoCanjes = 1
Unload Menu
End Sub



Private Sub Image5_Click()
Call Audio.PlayWave(0, SND_CLICK)
Call salirmsgbox
End Sub

Private Sub Image6_Click()
Call Audio.PlayWave(0, SND_CLICK)
ModoCanjes = 0
frmCanjes.Label1.Visible = False
Call SendData("CDONAR")
ModoCanjes = 0
Unload Menu
End Sub

Private Sub Image7_Click()
Call Audio.PlayWave(0, SND_CLICK)
Call SendData("/SOPORTE")
Unload Menu
End Sub

Private Sub Image8_Click()
frmPrincipal.WindowState = vbMinimized
Unload Menu
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyEscape Then
Unload Menu
End If

End Sub

Private Sub Label1_Click()
Unload Me
End Sub
