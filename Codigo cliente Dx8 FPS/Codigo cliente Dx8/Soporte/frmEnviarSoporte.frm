VERSION 5.00
Begin VB.Form frmEnviarSoporte 
   BorderStyle     =   0  'None
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSoporte 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      MaxLength       =   300
      MultiLine       =   -1  'True
      TabIndex        =   0
      Tag             =   "Escriba el mensaje."
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   120
      MouseIcon       =   "frmEnviarSoporte.frx":0000
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   3360
      MouseIcon       =   "frmEnviarSoporte.frx":1982
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   2295
   End
End
Attribute VB_Name = "frmEnviarSoporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.PICTURE = LoadPicture(DirGraficos & "SGM.gif")
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0

If Len(txtSoporte) Then
    Call SendData("/ZOPORTE " & txtSoporte.Text)
End If
txtSoporte.Text = ""
Me.Hide
Case 1
txtSoporte.Text = ""
Me.Hide
End Select
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
