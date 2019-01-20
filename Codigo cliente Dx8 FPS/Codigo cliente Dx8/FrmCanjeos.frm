VERSION 5.00
Begin VB.Form frmCanjes 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Sistema de Canjeo"
   ClientHeight    =   7500
   ClientLeft      =   420
   ClientTop       =   315
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   500
      Left            =   5400
      Picture         =   "FrmCanjeos.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   450
      TabIndex        =   9
      Top             =   1320
      Width           =   450
   End
   Begin VB.TextBox lDescripcion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1320
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3120
      Width           =   3015
   End
   Begin VB.ListBox ListaPremios 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   6030
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   3645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ATENCIÓN: Los item de canjeo se caen!! Usar sacri mas info, presionar ""F1"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   7200
      Width           =   7215
   End
   Begin VB.Label LabelNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Selecciona un item"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label lPuntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   6360
      TabIndex        =   7
      Top             =   7080
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lCantidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   7200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lAtaque 
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6100
      TabIndex        =   5
      Top             =   2000
      Width           =   1335
   End
   Begin VB.Label lDef 
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6100
      TabIndex        =   4
      Top             =   2330
      Width           =   1335
   End
   Begin VB.Label lAM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
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
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lDM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Requiere 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6100
      TabIndex        =   1
      Top             =   2630
      Width           =   1245
   End
   Begin VB.Image bSalir 
      Height          =   450
      Left            =   7080
      Top             =   0
      Width           =   480
   End
   Begin VB.Image bAceptar 
      Height          =   1395
      Left            =   4080
      Top             =   5400
      Width           =   3195
   End
End
Attribute VB_Name = "frmCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bAceptar_Click()
If ListaPremios.Text = "" Then 'selecciona el item culeao
Unload Me
Exit Sub
End If


If ModoCanjes = 1 Then
Call SendData("SPX" & ListaPremios.ListIndex + 1)
Else
Call SendData("SPM" & ListaPremios.ListIndex + 1)
End If
Unload Me
End Sub
     



Private Sub ListaPremios_Click()
If ModoCanjes = 1 Then
Call SendData("IPX" & ListaPremios.ListIndex + 1)
Else
Call SendData("IPM" & ListaPremios.ListIndex + 1)
End If
LabelNombre.Caption = ListaPremios.Text

End Sub

Private Sub bSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\graficos\canjeo.jpg")
If ModoCanjes = 1 Then
Call SendData("IPX" & ListaPremios.ListIndex + 1)
Else
Call SendData("IPM" & ListaPremios.ListIndex + 1)
End If
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

