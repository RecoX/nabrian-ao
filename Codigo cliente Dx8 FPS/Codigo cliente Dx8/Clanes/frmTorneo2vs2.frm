VERSION 5.00
Begin VB.Form Torneo2 
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command17 
      Caption         =   "Mandar ulla"
      Height          =   375
      Left            =   4680
      TabIndex        =   23
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Mandar ulla"
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Anunciar priemro"
      Height          =   735
      Left            =   4560
      TabIndex        =   21
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Anunciar 2"
      Height          =   615
      Left            =   240
      TabIndex        =   20
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      Caption         =   "2 - 1 A Favor"
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton Command12 
      Caption         =   "2 - 1 A Favor"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Pierden"
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton Command10 
      Caption         =   "2 - 0 A Favor"
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton Command9 
      Caption         =   "1 - 1"
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton Command8 
      Caption         =   "1 - 0 A Favor"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Ganan Torneo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3120
      TabIndex        =   11
      Text            =   "Nick Personaje"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Text            =   "Nick Personaje"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Pierden"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "2 - 0 A Favor"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "1 - 1"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "1 - 0 A Favor"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DARLE SUM 2"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DARLE SUM"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4440
      TabIndex        =   3
      Text            =   "Nick Personaje"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Text            =   "Nick Personaje"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "Nick Personaje"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Nick Personaje"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Torneo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
Call SendData("/TELEP" & " " & Text1.Text & " " & "16 67 13")
Call SendData("/TELEP" & " " & Text3.Text & " " & "16 87 24")
Call SendData("/TELEP" & " " & Text2.Text & " " & "16 68 13")
Call SendData("/TELEP" & " " & Text4.Text & " " & "16 86 24")
End Sub

Private Sub Command10_Click()
Call SendData("/RMSG 2 A 0 A favor de" & " " & Text3.Text & "-" & Text4.Text)
Call SendData("/REVIVIR" & " " & Text1.Text)
Call SendData("/REVIVIR" & " " & Text2.Text)
End Sub

Private Sub Command11_Click()
Call SendData("/RMSG Pierden:" & " " & Text3.Text & "-" & Text4.Text & " y quedan descalificados del Torneo.")
End Sub

Private Sub Command12_Click()
Call SendData("/RMSG 2 A 1 A favor de" & " " & Text1.Text & "-" & Text2.Text)
Call SendData("/REVIVIR" & " " & Text3.Text)
Call SendData("/REVIVIR" & " " & Text4.Text)
End Sub

Private Sub Command13_Click()
Call SendData("/RMSG 2 A 1 A favor de" & " " & Text3.Text & "-" & Text4.Text)
Call SendData("/REVIVIR" & " " & Text1.Text)
Call SendData("/REVIVIR" & " " & Text2.Text)
End Sub

Private Sub Command14_Click()
Call SendData("/RMSG Se enfrentan nuevamente:" & " " & Text1.Text & "-" & Text2.Text & " vs " & Text3.Text & "-" & Text4.Text)
Call SendData("/RMSG Esquinas sale en:")
Call SendData("/CUENTA 5")
End Sub

Private Sub Command15_Click()
Call SendData("/RMSG Se enfrentan " & " " & Text1.Text & "-" & Text2.Text & " vs " & Text3.Text & "-" & Text4.Text)
Call SendData("/RMSG Esquinas sale en:")
Call SendData("/CUENTA 5")
End Sub

Private Sub Command16_Click()
Call SendData("/TELEP" & " " & Text1.Text & " " & "1 50 50")
Call SendData("/TELEP" & " " & Text2.Text & " " & "1 50 50")
End Sub

Private Sub Command17_Click()
Call SendData("/TELEP" & " " & Text3.Text & " " & "1 50 50")
Call SendData("/TELEP" & " " & Text4.Text & " " & "1 50 50")
End Sub

Private Sub Command2_Click()
Call SendData("/TELEP" & " " & Text1.Text & " " & "16 67 13")
Call SendData("/TELEP" & " " & Text3.Text & " " & "16 87 24")
Call SendData("/TELEP" & " " & Text2.Text & " " & "16 68 13")
Call SendData("/TELEP" & " " & Text4.Text & " " & "16 86 24")
End Sub

Private Sub Command3_Click()
Call SendData("/RMSG 1 A 0 A favor de" & " " & Text1.Text & "-" & Text2.Text)
Call SendData("/REVIVIR" & " " & Text3.Text)
Call SendData("/REVIVIR" & " " & Text4.Text)
End Sub
Private Sub Command4_Click()
Call SendData("/RMSG Lo empatan" & " " & Text1.Text & "-" & Text2.Text)
Call SendData("/REVIVIR" & " " & Text3.Text)
Call SendData("/REVIVIR" & " " & Text4.Text)
End Sub

Private Sub command5_Click()
Call SendData("/RMSG 2 A 0 A favor de" & " " & Text1.Text & "-" & Text2.Text)
Call SendData("/REVIVIR" & " " & Text3.Text)
Call SendData("/REVIVIR" & " " & Text4.Text)
End Sub

Private Sub Command6_Click()
Call SendData("/RMSG Pierden:" & " " & Text1.Text & "-" & Text2.Text & " y quedan descalificados del Torneo.")
End Sub

Private Sub Command7_Click()
Call SendData("/RMSG Los Ganadores del Torneo son " & " " & Text5.Text & "-" & Text6.Text)
Call SendData("/RMSG Gracias por Participar..")
End Sub

Private Sub Command8_Click()
Call SendData("/RMSG 1 A 0 A favor de" & " " & Text3.Text & "-" & Text4.Text)
Call SendData("/REVIVIR" & " " & Text1.Text)
Call SendData("/REVIVIR" & " " & Text2.Text)
End Sub

Private Sub Command9_Click()
Call SendData("/RMSG Lo empatan" & " " & Text3.Text & "-" & Text4.Text)
Call SendData("/REVIVIR" & " " & Text1.Text)
Call SendData("/REVIVIR" & " " & Text2.Text)
End Sub
