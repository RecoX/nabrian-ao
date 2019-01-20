VERSION 5.00
Begin VB.Form frmQuestInfo 
   BorderStyle     =   0  'None
   Caption         =   "Quests Info"
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Desc 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   910
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label Tipo 
      BackStyle       =   0  'Transparent
      Caption         =   "NPC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   5
      Top             =   300
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   540
      Left            =   2640
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   240
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label GLDPT 
      BackStyle       =   0  'Transparent
      Caption         =   "10 Puntos de canjeos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   4
      Left            =   2670
      TabIndex        =   4
      Top             =   1480
      Width           =   2295
   End
   Begin VB.Label PosName 
      BackStyle       =   0  'Transparent
      Caption         =   "Momia - Dungeon abismal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   3
      Top             =   1185
      Width           =   2895
   End
   Begin VB.Label NPCs 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   880
      Width           =   1575
   End
   Begin VB.Label Users 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   590
      Width           =   2775
   End
End
Attribute VB_Name = "frmQuestInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "Quest.gif")
End Sub

Private Sub Image1_Click()
Call SendData("ACQT" & Numeriyo)
Unload Me
Unload frmQuestSelect
End Sub
Private Sub Image2_Click()
Unload Me
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

