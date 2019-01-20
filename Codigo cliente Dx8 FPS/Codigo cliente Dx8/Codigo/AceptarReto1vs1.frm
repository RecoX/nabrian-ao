VERSION 5.00
Begin VB.Form AceptarReto1vs1 
   BorderStyle     =   0  'None
   Caption         =   "Aceptar reto"
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Modo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2040
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   215
      TabIndex        =   1
      Top             =   1410
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NickName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
End
Attribute VB_Name = "AceptarReto1vs1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\graficos\enviareto.gif")
End Sub



Private Sub Image1_Click()
Call SendData("/ACEPTO " & Label1.Caption)
Unload AceptarReto1vs1
End Sub

Private Sub Image2_Click()
Call SendData("/RECHAZO")
Unload AceptarReto1vs1
End Sub

