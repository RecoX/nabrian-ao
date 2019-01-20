VERSION 5.00
Begin VB.Form frmParty 
   BorderStyle     =   0  'None
   Caption         =   "Party"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox ListaIntegrantes 
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
      Height          =   750
      ItemData        =   "Party.frx":0000
      Left            =   600
      List            =   "Party.frx":0007
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No te encuentras en ningún party. Para formar uno, clickea en el botón ""Invitar al party""."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Top             =   4080
      Width           =   735
   End
   Begin VB.Image Salir 
      Height          =   360
      Left            =   810
      MouseIcon       =   "Party.frx":001D
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   2370
   End
   Begin VB.Image Echar 
      Height          =   360
      Left            =   810
      MouseIcon       =   "Party.frx":0327
      MousePointer    =   99  'Custom
      Top             =   2760
      Width           =   2370
   End
   Begin VB.Image Invitar 
      Height          =   360
      Left            =   810
      MouseIcon       =   "Party.frx":0631
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   2355
   End
End
Attribute VB_Name = "frmParty"
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

Private Sub Form_Load()


Me.Picture = LoadPicture(DirGraficos & "sistemaparty.gif")
Invitar.Picture = LoadPicture(DirGraficos & "invitar.gif")
Echar.Picture = LoadPicture(DirGraficos & "echarparty.gif")
Salir.Picture = LoadPicture(DirGraficos & "salirparty.gif")

End Sub

Private Sub Salir_Click()
Call SendData("PARSAL")
Unload frmParty
End Sub
Private Sub Echar_Click()
Call SendData("PARECH" & frmParty.ListaIntegrantes.Text)
Unload frmParty
End Sub
Private Sub Invitar_Click()
Call SendData("UK" & Invita)
Unload frmParty
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
