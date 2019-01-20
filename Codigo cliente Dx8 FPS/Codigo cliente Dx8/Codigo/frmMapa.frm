VERSION 5.00
Begin VB.Form frmMapa 
   BorderStyle     =   0  'None
   ClientHeight    =   6675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   453
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   375
      Left            =   6360
      MouseIcon       =   "frmMapa.frx":0000
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public BotonMapa As Byte
Public MouseX As Long
Public MouseY As Long

Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "MapaDeJuego.gif")
End Sub
Private Sub Form_Click()

If BotonMapa = 2 Then Call TelepPorMapa(MouseX, MouseY)

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'personaje.Left = IzquierdaMapa + (UserPos.x - 50) * 0.18
'personaje.Top = TopMapa + (UserPos.y - 50) * 0.18

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

'personaje.Left = IzquierdaMapa + ((UserPos.x - 50) * 0.18)
'personaje.Top = TopMapa + ((UserPos.y - 50) * 0.18)

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'personaje.Left = IzquierdaMapa + (UserPos.x - 50) * 0.18
'personaje.Top = TopMapa + (UserPos.y - 50) * 0.18

End Sub

Private Sub Form_LostFocus()
Me.Visible = False
End Sub

Private Sub Form_GotFocus()
'personaje.Left = IzquierdaMapa + (UserPos.x - 50) * 0.18
'personaje.Top = TopMapa + (UserPos.y - 50) * 0.18
End Sub


Private Sub Image1_Click()
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
