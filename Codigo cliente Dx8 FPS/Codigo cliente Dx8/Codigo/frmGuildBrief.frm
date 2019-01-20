VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7620
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
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
   ScaleHeight     =   7650
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Desc 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   6240
      Width           =   6495
   End
   Begin VB.Image aliado 
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildBrief.frx":0000
      MousePointer    =   99  'Custom
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image command3 
      Height          =   375
      Left            =   5040
      MouseIcon       =   "frmGuildBrief.frx":030A
      MousePointer    =   99  'Custom
      Top             =   7080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image guerra 
      Height          =   495
      Left            =   1320
      MouseIcon       =   "frmGuildBrief.frx":0614
      MousePointer    =   99  'Custom
      Top             =   7560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image command1 
      Height          =   375
      Left            =   0
      MouseIcon       =   "frmGuildBrief.frx":091E
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   855
   End
   Begin VB.Image command2 
      Height          =   375
      Left            =   5040
      MouseIcon       =   "frmGuildBrief.frx":0C28
      MousePointer    =   99  'Custom
      Top             =   7080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   880
      Width           =   5535
   End
   Begin VB.Label fundador 
      BackStyle       =   0  'Transparent
      Caption         =   "Fundador:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label creacion 
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/0000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   1250
      Width           =   4335
   End
   Begin VB.Label lider 
      BackStyle       =   0  'Transparent
      Caption         =   "Nadie"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   1605
      Width           =   5745
   End
   Begin VB.Label web 
      BackStyle       =   0  'Transparent
      Caption         =   "Web site:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   2330
      Width           =   5835
   End
   Begin VB.Label eleccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Dias para proxima eleccion de lider:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Enemigos 
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes Enemigos:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Aliados 
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes Aliados:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Label Codex 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   7
      Top             =   3600
      Width           =   6495
   End
   Begin VB.Label Codex 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   6
      Top             =   3840
      Width           =   6495
   End
   Begin VB.Label Codex 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   5
      Top             =   4080
      Width           =   6495
   End
   Begin VB.Label Codex 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   4
      Top             =   4320
      Width           =   6495
   End
   Begin VB.Label Codex 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   3
      Top             =   4560
      Width           =   6495
   End
   Begin VB.Label Codex 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   2
      Top             =   4800
      Width           =   6495
   End
   Begin VB.Label Codex 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   1
      Top             =   5040
      Width           =   6495
   End
   Begin VB.Label Codex 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   0
      Top             =   5280
      Width           =   6495
   End
End
Attribute VB_Name = "frmGuildBrief"
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
Public EsLeader As Boolean
Public Sub ParseGuildInfo(ByVal buffer As String)
Dim BandoClan As Byte, BandoMio As Byte

BandoClan = Val(ReadFieldOptimizado(8, buffer, Asc("¬")))
BandoMio = Val(ReadFieldOptimizado(11, buffer, Asc("¬")))

If Not EsLeader Then
    Me.Picture = LoadPicture(DirGraficos & "DetallesDeClan.gif")
    guerra.Visible = False
    aliado.Visible = False
    Command3.Visible = False
    Command2.Visible = (BandoMio = BandoClan)
Else
    Me.Picture = LoadPicture(DirGraficos & "DetallesDeClanGuildMaster.gif")
    aliado.Visible = True
    guerra.Visible = True
    Command3.Visible = True
    Command2.Visible = False
End If

Select Case BandoClan
    Case 1
        Nombre.ForeColor = &HFF0000
    Case 2
        Nombre.ForeColor = &HFF&
    Case Else
        Nombre.ForeColor = &HE0E0E0
End Select

Nombre.Caption = ReadFieldOptimizado(1, buffer, Asc("¬"))
fundador.Caption = ReadFieldOptimizado(2, buffer, Asc("¬"))
creacion.Caption = ReadFieldOptimizado(3, buffer, Asc("¬"))
lider.Caption = ReadFieldOptimizado(4, buffer, Asc("¬"))
web.Caption = ReadFieldOptimizado(5, buffer, Asc("¬"))
Miembros.Caption = ReadFieldOptimizado(6, buffer, Asc("¬"))
Eleccion.Caption = ReadFieldOptimizado(7, buffer, Asc("¬"))
Enemigos.Caption = ReadFieldOptimizado(9, buffer, Asc("¬"))
Aliados.Caption = ReadFieldOptimizado(10, buffer, Asc("¬"))

Dim T%, k%
k% = Val(ReadFieldOptimizado(12, buffer, Asc("¬")))

For T% = 1 To k%
    Codex(T% - 1).Caption = ReadFieldOptimizado(12 + T%, buffer, Asc("¬"))
Next T%
Dim des$


des$ = ReadFieldOptimizado(12 + T%, buffer, Asc("¬"))

Desc = Replace(des$, "º", vbCrLf)

Me.Show vbModeless, frmPrincipal

End Sub

Private Sub aliado_Click()
Call SendData("DECALIAD" & Right$(Nombre, Len(Nombre.Caption)))
Unload Me
End Sub



Private Sub command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

Call frmGuildSol.RecieveSolicitud(Right$(Nombre, Len(Nombre.Caption)))
Call frmGuildSol.Show(vbModeless, frmGuildBrief)


End Sub

Private Sub Command3_Click()
frmCommet.Nombre = Right$(Nombre.Caption, Len(Nombre.Caption))
Call frmCommet.Show(vbModeless, frmGuildBrief)

End Sub


Private Sub guerra_Click()
Call SendData("DECGUERR" & Right$(Nombre.Caption, Len(Nombre.Caption)))
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
