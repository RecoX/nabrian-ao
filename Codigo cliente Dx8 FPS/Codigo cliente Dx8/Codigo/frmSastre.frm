VERSION 5.00
Begin VB.Form frmSastre 
   BorderStyle     =   0  'None
   Caption         =   "Sastre"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   -90
   ClientWidth     =   5250
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "1"
      Top             =   3560
      Width           =   1695
   End
   Begin VB.ListBox lstRopas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2190
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   4065
   End
   Begin VB.Image command4 
      Height          =   375
      Left            =   0
      MouseIcon       =   "frmSastre.frx":0000
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   855
   End
   Begin VB.Image command3 
      Height          =   375
      Left            =   3000
      MouseIcon       =   "frmSastre.frx":030A
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "frmSastre"
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

Private Sub Command3_Click()
On Error Resume Next
Dim stxtCantBuffer As String
stxtCantBuffer = txtCantidad.Text

Call SendData("SCR" & ObjSastre(lstRopas.ListIndex) & " " & stxtCantBuffer)

Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()

Me.SetFocus
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "Sastreria.gif")

End Sub

Private Sub txtCantidad_Change()
If Val(txtCantidad.Text) < 0 Then
    txtCantidad.Text = 1
End If

If Val(txtCantidad.Text) > MAX_INVENTORY_OBJS Then
    txtCantidad.Text = 1
End If
If Not IsNumeric(txtCantidad.Text) Then txtCantidad.Text = "1"

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
