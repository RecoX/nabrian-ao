VERSION 5.00
Begin VB.Form frmCharInfo 
   BorderStyle     =   0  'None
   Caption         =   "Info"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6000
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
   ScaleHeight     =   6375
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label neutrales 
      BackStyle       =   0  'Transparent
      Caption         =   "Neutrales asesinados:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   5280
      Width           =   4695
   End
   Begin VB.Image aceptar 
      Height          =   375
      Left            =   4440
      MouseIcon       =   "frmCharInfo.frx":0000
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   855
   End
   Begin VB.Image rechazar 
      Height          =   375
      Left            =   3120
      MouseIcon       =   "frmCharInfo.frx":030A
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Image desc 
      Height          =   375
      Left            =   1920
      MouseIcon       =   "frmCharInfo.frx":0614
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   975
   End
   Begin VB.Image echar 
      Height          =   375
      Left            =   600
      MouseIcon       =   "frmCharInfo.frx":091E
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      MouseIcon       =   "frmCharInfo.frx":0C28
      MousePointer    =   99  'Custom
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Ciudadanos 
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudadanos asesinados:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   4800
      Width           =   4695
   End
   Begin VB.Label criminales 
      BackStyle       =   0  'Transparent
      Caption         =   "Criminales asesinados:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   5040
      Width           =   4695
   End
   Begin VB.Label Solicitudes 
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitudes para ingresar a clanes:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   3000
      Width           =   4695
   End
   Begin VB.Label solicitudesRechazadas 
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitudes rechazadas:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   3240
      Width           =   4695
   End
   Begin VB.Label fundo 
      BackStyle       =   0  'Transparent
      Caption         =   "Fundo el clan:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   3480
      Width           =   4695
   End
   Begin VB.Label lider 
      BackStyle       =   0  'Transparent
      Caption         =   "Veces fue lider de clan:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Label integro 
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes que integro:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   3960
      Width           =   4695
   End
   Begin VB.Label Nombre 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Nivel 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Label Clase 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Clase:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Raza 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Raza:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label Genero 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Genero:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label Oro 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   4695
   End
   Begin VB.Label Banco 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Banco:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   4695
   End
   Begin VB.Label faccion 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   2400
      Width           =   4695
   End
End
Attribute VB_Name = "frmCharInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FenixAO DirectX8
'Engine By ·Parra
'Arreglado By Thusing
'Algunas cosas tomadas del cliente de DarkTester

Public frmmiembros As Byte
Public frmsolicitudes As Boolean
Private Sub Aceptar_Click()

frmmiembros = False
frmsolicitudes = False
Call SendData("ACEPTARI" & Right$(Nombre, Len(Nombre) - 8))
frmGuildLeader.Visible = False
Call SendData("GLINFO")

Unload Me

End Sub
Private Sub command1_Click()

Unload Me

End Sub
Public Sub parseCharInfo(ByVal Rdata As String)

Select Case frmmiembros
    Case 0
        Rechazar.Visible = True
        Aceptar.Visible = True
        Echar.Visible = False
        desc.Visible = True
    Case 1
        Rechazar.Visible = False
        Aceptar.Visible = False
        Echar.Visible = True
        desc.Visible = False
    Case 2
        Rechazar.Visible = False
        Aceptar.Visible = False
        Echar.Visible = False
        desc.Visible = False
End Select

Nombre.Caption = "Nombre: " & ReadField(1, Rdata, 44)
Raza.Caption = "Raza: " & ReadField(2, Rdata, 44)
Clase.Caption = "Clase: " & ReadField(3, Rdata, 44)
Genero.Caption = "Genero: " & ReadField(4, Rdata, 44)
Nivel.Caption = "Nivel: " & ReadField(5, Rdata, 44)
Oro.Caption = "Oro: " & ReadField(6, Rdata, 44)
Banco.Caption = "Banco: " & ReadField(7, Rdata, 44)

If Val(ReadField(8, Rdata, 44)) = 1 Then
    fundo.Caption = "Fundo el clan: " & ReadField(9, Rdata, 44)
Else
    fundo.Caption = "Fundo el clan: Ninguno"
End If

Solicitudes.Caption = "Solicitudes para ingresar a clanes: " & ReadField(10, Rdata, 44)
solicitudesRechazadas.Caption = "Solicitudes rechazadas: " & ReadField(11, Rdata, 44)
lider.Caption = "Veces fue lider de clan: " & ReadField(12, Rdata, 44)
integro.Caption = "Clanes que integro: " & ReadField(13, Rdata, 44)

Select Case Val(ReadField(14, Rdata, 44))
    Case 0
        faccion.ForeColor = vbWhite
        faccion.Caption = "Faccion: Neutral"
    Case 1
        faccion.ForeColor = vbBlue
        faccion.Caption = "Faccion: Alianza del Fenix"
    Case 2
        faccion.ForeColor = vbRed
        faccion.Caption = "Faccion: Ejército de Lord Thek"
End Select

neutrales.Caption = "Neutrales asesinados: " & ReadField(15, Rdata, 44)
Ciudadanos.Caption = "Ciudadanos asesinados: " & ReadField(16, Rdata, 44)
criminales.Caption = "Criminales asesinados: " & ReadField(17, Rdata, 44)

Me.Show vbModeless, frmMain

End Sub
Private Sub desc_Click()

Call SendData("ENVCOMEN" & Right$(Nombre, Len(Nombre) - 8))

End Sub
Private Sub Echar_Click()

Call SendData("ECHARCLA" & Right$(Nombre, Len(Nombre) - 8))
frmmiembros = 0
frmsolicitudes = False
frmGuildLeader.Visible = False
Call SendData("GLINFO")
Unload Me

End Sub
Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "CharInfo.gif")
Echar.Picture = LoadPicture(DirGraficos & "echar.gif")
desc.Picture = LoadPicture(DirGraficos & "desc.gif")
Aceptar.Picture = LoadPicture(DirGraficos & "aceptar.gif")
Rechazar.Picture = LoadPicture(DirGraficos & "rechazar.gif")

End Sub
Private Sub Image1_Click()

Unload Me

End Sub
Private Sub Rechazar_Click()

Call SendData("RECHAZAR" & Right$(Nombre, Len(Nombre) - 8))
frmmiembros = 0
frmsolicitudes = False
frmGuildLeader.Visible = False
Call SendData("GLINFO")
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

If bmoving And ((X <> Dx3) Or (Y <> dy)) Then Move Left + (X - Dx3), Top + (Y - dy)

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub
Private Sub Text1_Change()

If Val(Text1.Text) < 0 Then
    Text1.Text = MAX_INVENTORY_OBJS
End If

If Val(Text1.Text) > MAX_INVENTORY_OBJS And ItemElegido <> FLAGORO Then
    Text1.Text = 1
End If

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) Then
    If (Index <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If

End Sub
