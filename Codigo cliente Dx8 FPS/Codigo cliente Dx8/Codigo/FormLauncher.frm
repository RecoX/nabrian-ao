VERSION 5.00
Begin VB.Form FormLauncher 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   7395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormLauncher.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "FormLauncher.frx":000C
   ScaleHeight     =   7395
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   630
      Left            =   730
      Top             =   6270
      Width           =   3945
   End
   Begin VB.Image Image4 
      Height          =   630
      Left            =   730
      Top             =   5400
      Width           =   3945
   End
   Begin VB.Image foro 
      Height          =   630
      Left            =   730
      Top             =   4530
      Width           =   3945
   End
   Begin VB.Image Errores 
      Height          =   630
      Left            =   730
      Top             =   3650
      Width           =   3945
   End
   Begin VB.Image Image1 
      Height          =   630
      Left            =   730
      Top             =   1800
      Width           =   3945
   End
   Begin VB.Label LabelVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V 1.0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "FormLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Errores_Click()
'Call Audio.PlayWave(0,  SND_CLICK)
Shell "ErroresFIX.exe"
End Sub

Private Sub foro_Click()
ShellExecute Me.hwnd, "open", "http://foro.nabrianao.net/", "", "", 1
End Sub

Private Sub Image1_Click()
On Error Resume Next
'Call Audio.PlayWave(0,  SND_CLICK)
Call MainShell

End Sub

Private Sub Form_Load()
On Error Resume Next
'Me.Picture = LoadPicture(App.Path & "\graficos\Launcher.jpg")
VersionDelJuego = "v" & App.Major & "." & App.Minor & "." & App.Revision
LabelVersion = VersionDelJuego
'Call RunAsAdmin
SeguridadActiva = True 'ESTO ACTIVA Y DESACTIVA TODA LA SEGURIDAD DEL CLIENTE (CHEATS BASICO, CLASES, SH, 2Clients, ExeName, hash, y algunas encriptaciones (No todas algunas son apartes)).
EncriptGraficosActiva = False

Call GetSerialNumber2

'Errores.Picture = LoadPicture("Graficos\botonerror.jpg")
'Image1.Picture = LoadPicture("Graficos\botonjugar.jpg")
'Image2.Picture = LoadPicture("Graficos\botonsalir.jpg")
'foro.Picture = LoadPicture("Graficos\botonforo.jpg")
'Image4.Picture = LoadPicture("Graficos\botonweb.jpg")
End Sub


Private Sub Image2_Click()
Unload Me
End
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



Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image4_Click()
ShellExecute Me.hwnd, "open", "http://nabrianao.net/", "", "", 1
End Sub
