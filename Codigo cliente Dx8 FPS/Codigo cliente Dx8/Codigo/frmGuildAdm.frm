VERSION 5.00
Begin VB.Form frmGuildAdm 
   BorderStyle     =   0  'None
   Caption         =   "Lista de Clanes Registrados"
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4155
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
   ScaleHeight     =   3900
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox GuildsList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1950
      ItemData        =   "frmGuildAdm.frx":0000
      Left            =   600
      List            =   "frmGuildAdm.frx":0002
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
   Begin VB.Image command3 
      Height          =   375
      Left            =   3360
      MouseIcon       =   "frmGuildAdm.frx":0004
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   855
   End
   Begin VB.Image command1 
      Height          =   375
      Left            =   1320
      MouseIcon       =   "frmGuildAdm.frx":030E
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   1575
   End
End
Attribute VB_Name = "frmGuildAdm"
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
Dim GuildName As String


GuildName = guildslist.List(guildslist.ListIndex)
If Right$(GuildName, 1) = ")" Then GuildName = Left$(GuildName, Len(GuildName) - 4)

Call SendData("CLANDETAILS" & GuildName)
frmGuildAdm.Visible = False

End Sub
Private Sub Command3_Click()

Me.Visible = False
frmPrincipal.SetFocus

End Sub
Public Sub ParseGuildList(ByVal Rdata As String)
Dim j As Integer, k As Integer

k = CInt(ReadFieldOptimizado(1, Rdata, 44))

For j = 1 To k
    guildslist.AddItem ReadFieldOptimizado(1 + j, Rdata, 44)
Next j

Me.Show

End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "listadeclanes.gif")
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
