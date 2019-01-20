VERSION 5.00
Begin VB.Form EventoZombie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EventoZombies"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Text            =   "0"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CANCELAR"
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "COMENZAR"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ABRIR CUPOS PARA:"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "EventoZombie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call SendData("/EVENTOZOMBIE " & Text1.Text)
End Sub

Private Sub Command2_Click()
Call SendData("/SOMBIEC")
End Sub

Private Sub Command3_Click()
Call SendData("/ZCANCELAR")
End Sub
