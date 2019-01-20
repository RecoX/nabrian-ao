VERSION 5.00
Begin VB.Form frmPeaceProp 
   BorderStyle     =   0  'None
   Caption         =   "Ofertas de paz"
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4725
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
   ScaleHeight     =   3885
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lista 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
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
      Height          =   1950
      ItemData        =   "frmPeaceProp.frx":0000
      Left            =   580
      List            =   "frmPeaceProp.frx":0002
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   0
      MouseIcon       =   "frmPeaceProp.frx":0004
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   600
      MouseIcon       =   "frmPeaceProp.frx":030E
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3120
      MouseIcon       =   "frmPeaceProp.frx":0618
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "frmPeaceProp"
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
Public Sub ParsePeaceOffers(ByVal s As String)
Dim T%, R%

T% = Val(ReadFieldOptimizado(1, s, 44))

For R% = 1 To T%
    Call lista.AddItem(ReadFieldOptimizado(R% + 1, s, 44))
Next R%

Me.Show vbModeless, frmPrincipal

End Sub
Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "OfertaDePazParaGuildMaster.gif")

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
Private Sub Image1_Click()
Call SendData("ACEPPEAT" & lista.List(lista.ListIndex))
Unload Me
End Sub

Private Sub Image2_Click()
Call SendData("PEACEDET" & lista.List(lista.ListIndex))
End Sub

Private Sub Image3_Click()
Unload Me
End Sub

