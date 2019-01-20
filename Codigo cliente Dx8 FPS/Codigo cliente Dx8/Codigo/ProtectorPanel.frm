VERSION 5.00
Begin VB.Form ProtectorPanel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel de evento protector."
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3030
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Text            =   "0"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Abrir cupos"
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "CANCELAR"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Con items de canjeo"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sin items de canjeo"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
   End
End
Attribute VB_Name = "ProtectorPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
Call SendData("/PROTECTORTHE " & Text1.Text)
End Sub


Private Sub Command2_Click()
Call SendData("/DENOCHE 4")
End Sub

Private Sub Command3_Click()
Call SendData("/ZKGJX")
End Sub

Private Sub Command4_Click()
Call SendData("/DENOCHE 3")
End Sub

