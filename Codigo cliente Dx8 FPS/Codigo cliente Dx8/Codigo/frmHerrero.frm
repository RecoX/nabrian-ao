VERSION 5.00
Begin VB.Form frmHerrero 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Herrero"
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   360
   ClientWidth     =   5250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Text            =   "1"
      Top             =   3545
      Width           =   1695
   End
   Begin VB.ListBox lstCascos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2340
      ItemData        =   "frmHerrero.frx":0000
      Left            =   600
      List            =   "frmHerrero.frx":0002
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   4080
   End
   Begin VB.ListBox lstEscudos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFFFF&
      Height          =   2370
      ItemData        =   "frmHerrero.frx":0004
      Left            =   600
      List            =   "frmHerrero.frx":0006
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   4080
   End
   Begin VB.ListBox lstArmas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFFFF&
      Height          =   2370
      ItemData        =   "frmHerrero.frx":0008
      Left            =   600
      List            =   "frmHerrero.frx":000A
      TabIndex        =   1
      Top             =   960
      Width           =   4080
   End
   Begin VB.ListBox lstArmaduras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFFFF&
      Height          =   2370
      ItemData        =   "frmHerrero.frx":000C
      Left            =   590
      List            =   "frmHerrero.frx":000E
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   4080
   End
   Begin VB.Image command7 
      Height          =   615
      Left            =   3840
      MouseIcon       =   "frmHerrero.frx":0010
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   615
   End
   Begin VB.Image command6 
      Height          =   615
      Left            =   1800
      MouseIcon       =   "frmHerrero.frx":031A
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   615
   End
   Begin VB.Image command4 
      Height          =   375
      Left            =   0
      MouseIcon       =   "frmHerrero.frx":0624
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   855
   End
   Begin VB.Image command3 
      Height          =   375
      Left            =   3000
      MouseIcon       =   "frmHerrero.frx":092E
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Image command2 
      Height          =   615
      Left            =   720
      MouseIcon       =   "frmHerrero.frx":0C38
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   615
   End
   Begin VB.Image command1 
      Height          =   615
      Left            =   2880
      MouseIcon       =   "frmHerrero.frx":0F42
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   615
   End
End
Attribute VB_Name = "frmHerrero"
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

lstArmaduras.Visible = False
lstArmas.Visible = True
lstEscudos.Visible = False
lstCascos.Visible = False
End Sub

Private Sub Command2_Click()

lstArmaduras.Visible = True
lstArmas.Visible = False
lstEscudos.Visible = False
lstCascos.Visible = False
End Sub

Private Sub Command3_Click()

On Error Resume Next
Dim stxtCantBuffer As String
stxtCantBuffer = txtCantidad.Text

If lstArmas.Visible Then
 Call SendData("CNS" & ArmasHerrero(lstArmas.ListIndex) & " " & stxtCantBuffer)
ElseIf lstArmaduras.Visible Then
 Call SendData("CNS" & ArmadurasHerrero(lstArmaduras.ListIndex) & " " & stxtCantBuffer)
 ElseIf lstEscudos.Visible Then
 Call SendData("CNS" & EscudosHerrero(lstEscudos.ListIndex) & " " & stxtCantBuffer)
 ElseIf lstCascos.Visible Then
 Call SendData("CNS" & CascosHerrero(lstCascos.ListIndex) & " " & stxtCantBuffer)
End If

Unload Me

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command6_Click()
lstArmaduras.Visible = False
lstArmas.Visible = False
lstEscudos.Visible = True
lstCascos.Visible = False
End Sub

Private Sub Command7_Click()
lstArmaduras.Visible = False
lstArmas.Visible = False
lstEscudos.Visible = False
lstCascos.Visible = True
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub
Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "Herreria.gif")

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

Private Sub txtCantidad_Change()

If Not IsNumeric(txtCantidad.Text) Then txtCantidad.Text = "1"

End Sub
