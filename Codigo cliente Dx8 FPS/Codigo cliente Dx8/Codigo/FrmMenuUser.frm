VERSION 5.00
Begin VB.Form FrmMenuUser 
   BorderStyle     =   0  'None
   Caption         =   "                          Menu usuario"
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.Label textbox 
      Caption         =   "textbox"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   600
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   600
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   600
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   600
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   600
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "FrmMenuUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\graficos\paneluser.gif")
End Sub

Private Sub Image1_Click()
frmMandarReto.Text2.Text = Label1.Caption
Call SendData("/RETAR")
Unload Me
End Sub

Private Sub Image2_Click()
SendData ("/PAREJA " & Label1.Caption)
Unload Me
End Sub

Private Sub Image3_Click()
SendData ("/TPAREJA " & Label1.Caption)
Unload Me
End Sub

Private Sub Image4_Click()
SendData ("/COMERCIAR")
Unload Me
End Sub

Private Sub Image5_Click()
textbox.Caption = InputBox("¿Cuantos puntos desea transferir?", "Transferencia de puntos.", "0")
Label1.Caption = Replace(Label1.Caption, " ", "+")
Call SendData("/TRANSFERIX " & Label1.Caption & " " & textbox.Caption)
End Sub

Private Sub Label2_Click()
Unload Me
End Sub
