VERSION 5.00
Begin VB.Form frmMSG 
   BorderStyle     =   0  'None
   Caption         =   "GM Messenger"
   ClientHeight    =   7230
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   6450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox mensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   1575
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4680
      Width           =   5295
   End
   Begin VB.TextBox GM 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2040
      TabIndex        =   2
      Text            =   "Cualquier GM disponible"
      Top             =   2470
      Width           =   2775
   End
   Begin VB.ComboBox categoria 
      Appearance      =   0  'Flat
      BackColor       =   &H00111720&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmMSG.frx":0000
      Left            =   2880
      List            =   "frmMSG.frx":0013
      TabIndex        =   1
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3000
      MouseIcon       =   "frmMSG.frx":0065
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   2895
   End
   Begin VB.Image command1 
      Height          =   375
      Left            =   600
      MouseIcon       =   "frmMSG.frx":036F
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   2175
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FénixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar
Private Sub command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "SGM.gif")

End Sub
Private Sub Image1_Click()
Dim GMs As String

If categoria.ListIndex = -1 Then
    MsgBox "El motivo del mensaje no es válido"
    Exit Sub
End If

If Len(mensaje.Text) > 250 Then
    MsgBox "La longitud del mensaje debe tener menos de 250 carácteres."
    Exit Sub
End If

If Len(GM.Text) = 0 Or GM.Text = "Cualquier GM disponible" Then
    GMs = "Ninguno"
Else: GMs = GM.Text
End If

If Len(mensaje.Text) = 0 Then
    MsgBox "Debes ingresar un mensaje."
    Exit Sub
End If

Call SendData("GM" & GMs & "¬" & categoria.List(categoria.ListIndex) & "¬" & mensaje.Text)

If NoMandoElMsg = 0 Then
    mensaje.Text = ""
    GM.Text = "Cualquier GM disponible"
    categoria.List(categoria.ListIndex) = ""
    AddtoRichTextBox frmPrincipal.rectxt, "El mensaje fue enviado. Dentro de algunas horas recibirás la respuesta en el mail registrado. Rogamos tengas paciencia y no escribas más de un mensaje sobre el mismo tema.", 252, 151, 53, 1, 0
    Unload Me
Else
    Call MsgBox("El mensaje es demasiado largo, por favor resumilo.")
End If

End Sub



Private Sub Label7_Click()

End Sub

Private Sub mensaje_Change()
mensaje.Text = LTrim(mensaje.Text)
End Sub


Private Sub mensaje_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 209) And (KeyAscii <> 241) And (KeyAscii <> 8) And (KeyAscii <> 32) And (KeyAscii <> 164) And (KeyAscii <> 165) Then
    If (Index <> 6) And ((KeyAscii < 40 Or KeyAscii > 122) Or (KeyAscii > 90 And KeyAscii < 96)) Then
        KeyAscii = 0
    End If
End If

 KeyAscii = Asc((Chr(KeyAscii)))
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If bmoving = False And Button = vbLeftButton Then
      Dx3 = x
      dy = y
      bmoving = True
   End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If bmoving And ((x <> Dx3) Or (y <> dy)) Then
      Move Left + (x - Dx3), Top + (y - dy)
   End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      bmoving = False
   End If
End Sub

