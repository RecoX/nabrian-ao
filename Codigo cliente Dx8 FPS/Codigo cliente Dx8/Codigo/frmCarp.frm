VERSION 5.00
Begin VB.Form frmCarp 
   BorderStyle     =   0  'None
   Caption         =   "Carpintero"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
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
   Begin VB.ListBox lstArmas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
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
      Height          =   2190
      ItemData        =   "frmCarp.frx":0000
      Left            =   600
      List            =   "frmCarp.frx":0002
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   870
      Width           =   4080
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      MouseIcon       =   "frmCarp.frx":0004
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3000
      MouseIcon       =   "frmCarp.frx":030E
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   1575
   End
End
Attribute VB_Name = "frmCarp"
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

Call SendData("CNC" & ObjCarpintero(lstArmas.ListIndex))

Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub



Private Sub Form_Load()
Me.PICTURE = LoadPicture(DirGraficos & "fragua-carpinteria.gif")
End Sub

Private Sub Image1_Click()
On Error Resume Next
Dim stxtCantBuffer As String
stxtCantBuffer = txtCantidad.Text

Call SendData("CNC" & ObjCarpintero(lstArmas.ListIndex) & " " & stxtCantBuffer)
Unload Me

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
Private Sub Image2_Click()

Unload Me

End Sub
