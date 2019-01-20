VERSION 5.00
Begin VB.Form frmBorrar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2880
   ClientLeft      =   15
   ClientTop       =   -30
   ClientWidth     =   4530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodigo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      MouseIcon       =   "frmBorrar.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2400
      Width           =   1965
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2400
      MouseIcon       =   "frmBorrar.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2400
      Width           =   1995
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4350
   End
   Begin VB.Label Label2 
      Caption         =   "Una ves borrado el personaje        no podrá ser renovado."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   525
      Left            =   720
      TabIndex        =   7
      Top             =   480
      Width           =   2985
   End
   Begin VB.Label Label1 
      Caption         =   "Atención"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label4 
      Caption         =   "CODIGO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   825
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre del personaje:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2145
   End
End
Attribute VB_Name = "frmBorrar"
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

Private Sub cmdBorrar_Click()

Call SendData("BORRAR" & txtName & "," & txtCodigo)

Unload Me

End Sub

Private Sub Command2_Click()
frmPrincipal.Socket1.Disconnect
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmConectar.PictureLogin.Visible = True
frmConectar.txtUser.Visible = True
frmConectar.TxtPass.Visible = True
base_light = D3DColorXRGB(255, 255, 255)
End Sub


Private Sub txtCodigo_KeyPress(KeyAscii As Integer)

KeyAscii = Numeros(KeyAscii)

End Sub
