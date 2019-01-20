VERSION 5.00
Begin VB.Form Frmdeathmatch 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image2 
      Height          =   375
      Left            =   240
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2160
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Frmdeathmatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Form_Load()
Me.PICTURE = LoadPicture(App.Path & "\graficos\deathmatch.gif")
End Sub

Private Sub Image1_Click()
Call SendData("/ABANDONARDM")
Unload Frmdeathmatch
End Sub

Private Sub Image2_Click()
Call SendData("XDM")
Unload Frmdeathmatch
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
