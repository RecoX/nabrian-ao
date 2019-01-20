VERSION 5.00
Begin VB.Form frmGuildNews 
   BorderStyle     =   0  'None
   Caption         =   "Portada del Clan"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox aliados 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000009&
      Height          =   1005
      ItemData        =   "frmGuildNews.frx":0000
      Left            =   610
      List            =   "frmGuildNews.frx":0002
      TabIndex        =   2
      Top             =   5040
      Width           =   4370
   End
   Begin VB.ListBox guerra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000009&
      Height          =   1005
      ItemData        =   "frmGuildNews.frx":0004
      Left            =   610
      List            =   "frmGuildNews.frx":0006
      TabIndex        =   1
      Top             =   3440
      Width           =   4370
   End
   Begin VB.TextBox news 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000009&
      Height          =   2055
      Left            =   610
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   4370
   End
   Begin VB.Image command1 
      Height          =   255
      Left            =   0
      MouseIcon       =   "frmGuildNews.frx":0008
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   855
   End
End
Attribute VB_Name = "frmGuildNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Private Sub Command1_Click()
On Error Resume Next
Unload Me
frmMain.SetFocus
End Sub

Public Sub ParseGuildNews(ByVal s As String)

news = Replace(ReadField(1, s, Asc("¬")), "º", vbCrLf)

Dim h%, j%

h% = Val(ReadField(2, s, Asc("¬")))

For j% = 1 To h%
    
    guerra.AddItem ReadField(j% + 2, s, Asc("¬"))
    
Next j%

j% = j% + 2

h% = Val(ReadField(j%, s, Asc("¬")))

For j% = j% + 1 To j% + h%
    
    aliados.AddItem ReadField(j%, s, Asc("¬"))
    
Next j%

Me.Show vbModeless, frmMain

End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "GuildNews.jpg")
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving = False And Button = vbLeftButton Then

      DX = X

      dy = Y

      bmoving = True

   End If

   

End Sub

 

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving And ((X <> DX) Or (Y <> dy)) Then

      Move Left + (X - DX), Top + (Y - dy)

   End If

   

End Sub

 

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = vbLeftButton Then

      bmoving = False

   End If

   

End Sub

