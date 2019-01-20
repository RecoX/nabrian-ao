VERSION 5.00
Begin VB.Form frmRecupera 
   BorderStyle     =   0  'None
   Caption         =   "                    Recuperar personaje"
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   4455
   Icon            =   "frmRecu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recuperar personaje"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   2805
   End
   Begin VB.TextBox txtMail 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   2625
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2625
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   4560
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3360
   End
   Begin VB.Line Line2 
      X1              =   4440
      X2              =   4440
      Y1              =   -480
      Y2              =   2880
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4560
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmRecupera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function Numeros(Tecla As Integer) As Integer
Dim strValido As String

strValido = "0123456789qwertyuiopasdfghjklñmnbvcxzQWERTYUIOPASDFGHJKLÑZXCVBNM"
If Tecla > 26 Then
If InStr(strValido, Chr(Tecla)) = 0 Then
Tecla = 0
End If
End If
Numeros = Tecla
End Function

Private Sub command1_Click()
 
Call SendData("RECUPE" & txtName & "," & txtMail)

Unload Me

End Sub


Private Sub Command2_Click()
frmPrincipal.Socket1.Disconnect

Unload Me
End Sub

Private Sub txtMail_KeyPress(KeyAscii As Integer)

KeyAscii = Numeros(KeyAscii)

End Sub

Private Sub Form_Unload(Cancel As Integer)

frmConectar.PictureLogin.Visible = True
frmConectar.txtUser.Visible = True
frmConectar.TxtPass.Visible = True
base_light = D3DColorXRGB(255, 255, 255)
End Sub
