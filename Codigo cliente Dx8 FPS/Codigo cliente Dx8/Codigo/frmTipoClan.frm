VERSION 5.00
Begin VB.Form frmTipoClan 
   Caption         =   "Facción del clan"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   2760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Continuar"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Caos"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Neutral"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Real"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Elige a que facción deseas que pretenezca el clan que vas a crear:"
      Height          =   555
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2625
   End
End
Attribute VB_Name = "frmTipoClan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
Dim i As Integer

For i = 0 To 2
    If Option1(i).value Then
        ClanType = i
        Exit For
    End If
Next

frmGuildDetails.Show
Unload Me

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving = False And Button = vbLeftButton Then
    DX = X
    dy = Y
    bmoving = True
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving And ((X <> DX) Or (Y <> dy)) Then Move Left + (X - DX), Top + (Y - dy)

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub
