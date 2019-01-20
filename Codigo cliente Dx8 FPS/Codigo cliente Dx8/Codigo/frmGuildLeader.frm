VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   0  'None
   Caption         =   "Administración del Clan"
   ClientHeight    =   7950
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6825
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
   ScaleHeight     =   7950
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtguildnews 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3600
      Width           =   5535
   End
   Begin VB.ListBox members 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H00FFFFFF&
      Height          =   1590
      ItemData        =   "frmGuildLeader.frx":0000
      Left            =   3600
      List            =   "frmGuildLeader.frx":0002
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.ListBox guildslist 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H00FFFFFF&
      Height          =   1590
      ItemData        =   "frmGuildLeader.frx":0004
      Left            =   600
      List            =   "frmGuildLeader.frx":0006
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
   Begin VB.ListBox solicitudes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      ItemData        =   "frmGuildLeader.frx":0008
      Left            =   600
      List            =   "frmGuildLeader.frx":000A
      TabIndex        =   3
      Top             =   5520
      Width           =   2655
   End
   Begin VB.Label Miembros 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "El clan cuenta con x miembros"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   6550
      Width           =   2535
   End
   Begin VB.Image command4 
      Height          =   375
      Left            =   1200
      MouseIcon       =   "frmGuildLeader.frx":000C
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Image command3 
      Height          =   375
      Left            =   2640
      MouseIcon       =   "frmGuildLeader.frx":0316
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Image command2 
      Height          =   375
      Left            =   4080
      MouseIcon       =   "frmGuildLeader.frx":0620
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Image command1 
      Height          =   375
      Left            =   1200
      MouseIcon       =   "frmGuildLeader.frx":092A
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Image command8 
      Height          =   255
      Left            =   0
      MouseIcon       =   "frmGuildLeader.frx":0C34
      MousePointer    =   99  'Custom
      Top             =   7680
      Width           =   735
   End
   Begin VB.Image command7 
      Height          =   375
      Left            =   3360
      MouseIcon       =   "frmGuildLeader.frx":0F3E
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   2895
   End
   Begin VB.Image command6 
      Height          =   375
      Left            =   3480
      MouseIcon       =   "frmGuildLeader.frx":1248
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   2775
   End
   Begin VB.Image command5 
      Height          =   375
      Left            =   3480
      MouseIcon       =   "frmGuildLeader.frx":1552
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   2775
   End
End
Attribute VB_Name = "frmGuildLeader"
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

frmCharInfo.frmmiembros = 0
frmCharInfo.frmsolicitudes = True
Call SendData("1HRINFO<" & Solicitudes.List(Solicitudes.ListIndex))



End Sub

Private Sub Command2_Click()

frmCharInfo.frmmiembros = 15
Call SendData("1HRINFO<" & members.List(members.ListIndex))
Me.Visible = False


End Sub

Private Sub Command3_Click()

Dim k$

k$ = Replace(txtguildnews, vbCrLf, "º")

Call SendData("ACTGNEWS" & k$)

End Sub

Private Sub Command4_Click()
Dim GuildName As String


GuildName = guildslist.List(guildslist.ListIndex)
If Right$(GuildName, 1) = ")" Then GuildName = Left$(GuildName, Len(GuildName) - 4)

frmGuildBrief.EsLeader = True
Call SendData("CLANDETAILS" & GuildName)

End Sub

Private Sub Command5_Click()

Call frmGuildDetails.Show(vbModal, frmGuildLeader)



End Sub

Private Sub Command6_Click()
Call frmGuildURL.Show(vbModeless, frmGuildLeader)

End Sub

Private Sub Command7_Click()
Call SendData("ENVPROPP")
End Sub

Private Sub Command8_Click()
Me.Visible = False
frmPrincipal.SetFocus

End Sub
Private Function ListaDeClanes(ByVal Data As String) As Integer
Dim a As Integer
Dim i As Integer

a = Val(ReadFieldOptimizado(1, Data, Asc("¬")))
ReDim oClan(1 To a) As Clan

For i = 1 To a
    oClan(i).name = Left$(ReadFieldOptimizado(i + 1, Data, Asc("¬")), Len(ReadFieldOptimizado(i + 1, Data, Asc("¬"))) - 2)
    oClan(i).Relation = Right$(ReadFieldOptimizado(1 + i, Data, Asc("¬")), 1)
Next

For i = 1 To a
    If oClan(i).Relation = 4 Then
        Call guildslist.AddItem(oClan(i).name)
    End If
Next

For i = 1 To a
    If oClan(i).Relation = 1 Then
        Call guildslist.AddItem(oClan(i).name & " (A)")
    End If
Next

For i = 1 To a
    If oClan(i).Relation = 2 Then
        Call guildslist.AddItem(oClan(i).name & " (E)")
    End If
Next

For i = 1 To a
    If oClan(i).Relation = 0 Then
        Call guildslist.AddItem(oClan(i).name)
    End If
Next

ListaDeClanes = a + 2

End Function
Public Sub ParseLeaderInfo(ByVal Data As String)

guildslist.Clear
members.Clear
Solicitudes.Clear
txtguildnews = ""

If Me.Visible Then Exit Sub

Dim a As Integer
Dim b As Integer
Dim i As Integer

b = ListaDeClanes(Data)

a = Val(ReadFieldOptimizado(b, Data, Asc("¬")))

For i = 1 To a
    Call members.AddItem(ReadFieldOptimizado(b + i, Data, Asc("¬")))
Next

b = b + a + 1

Miembros.Caption = "El clan cuenta con " & a & " miembros."
txtguildnews = Replace(ReadFieldOptimizado(b, Data, Asc("¬")), "º", vbCrLf)

b = b + 1

a = Val(ReadFieldOptimizado(b, Data, Asc("¬")))

For i = 1 To a
    Solicitudes.AddItem ReadFieldOptimizado(b + i, Data, Asc("¬"))
Next

Call Me.Show(vbModeless, frmPrincipal)
Call Me.SetFocus

End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "GuildMaster.gif")

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
