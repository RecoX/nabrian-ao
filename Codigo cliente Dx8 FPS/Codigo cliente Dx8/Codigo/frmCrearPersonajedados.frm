VERSION 5.00
Begin VB.Form FrmCrearpersonaje 
   BorderStyle     =   0  'None
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCorreo 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   3240
      Width           =   2715
   End
   Begin VB.TextBox txtPasswdCheck 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2760
      Width           =   1995
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2280
      Width           =   2250
   End
   Begin VB.ComboBox lstGenero 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      ItemData        =   "frmCrearPersonajedados.frx":0000
      Left            =   1200
      List            =   "frmCrearPersonajedados.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1700
      Width           =   2760
   End
   Begin VB.ComboBox lstRaza 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      ItemData        =   "frmCrearPersonajedados.frx":001D
      Left            =   1200
      List            =   "frmCrearPersonajedados.frx":0030
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   2760
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   0
      Top             =   720
      Width           =   2730
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Presione aquí si desea cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label lblPass2OK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   540
      Left            =   4200
      TabIndex        =   4
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label lblPassOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   540
      Left            =   4200
      TabIndex        =   3
      Top             =   2640
      Width           =   240
   End
   Begin VB.Image boton 
      Appearance      =   0  'Flat
      Height          =   645
      Index           =   0
      Left            =   360
      MouseIcon       =   "frmCrearPersonajedados.frx":005D
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   3840
   End
End
Attribute VB_Name = "FrmCrearpersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SkillPoints As Byte
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
Function CheckData() As Boolean

If UserRaza = 0 Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserHogar = 0 Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

If UserSexo = -1 Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

Dim I As Integer
For I = 1 To NUMATRIBUTOS
    If UserAtributos(I) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next I

CheckData = True

End Function
Private Sub boton_Click(Index As Integer)
Dim I As Integer
Dim k As Object


Select Case Index
    Case 0
        LlegoConfirmacion = False
        Confirmacion = 0


        
        UserName = txtNombre.Text
        
        If Right$(UserName, 1) = " " Then
            UserName = Trim(UserName)
            MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
        
        UserRaza = lstRaza.ListIndex + 1
        UserSexo = lstGenero.ListIndex
        UserHogar = 4
        
        UserAtributos(1) = 1
        UserAtributos(2) = 1
        UserAtributos(3) = 1
        UserAtributos(4) = 1
        UserAtributos(5) = 1
        
        If CheckData() Then
            UserPassword = MD5String(txtPasswd.Text)
            UserEmail = TxtCorreo.Text
            
            Dim PW As Integer
            Dim NUM As Boolean
            If IsNumeric(TxtCorreo) Then
            MsgBox "El codigo debe tener Numeros y Letras."
            Exit Sub
            End If
            PW = 1
            NUM = False
            For PW = 1 To Len(TxtCorreo)
            If mid(TxtCorreo, PW, 1) Like "#" Then
            PW = Len(TxtCorreo)
            NUM = True
            End If
            Next
            If NUM = False Then
            MsgBox "El codigo debe tener Numeros y Letras."
            Exit Sub
            End If
            
            If Len(Trim(TxtCorreo)) < 8 Then
                MsgBox "El codigo debe tener al menos 8 caracteres.", vbExclamation, "Nabrian AO"
                TxtCorreo = ""
                TxtCorreo = ""
                TxtCorreo.SetFocus
                Exit Sub
            End If
            
            If Len(Trim(txtPasswd)) = 0 Then
                MsgBox "Tenés que ingresar una contraseña.", vbExclamation, "NabrianAO"
                txtPasswd.SetFocus
                Exit Sub
            End If
            
            If Len(Trim(txtNombre)) < 3 Then
                MsgBox "El Nick del personaje debe tener al menos 3 caracteres.", vbExclamation, "NabrianAO"
                txtPasswd = ""
                txtPasswdCheck = ""
                txtPasswd.SetFocus
                Exit Sub
            End If
            
            
            
            If Trim(txtPasswd) <> Trim(txtPasswdCheck) Then
                MsgBox "Las contraseñas no coinciden.", vbInformation, "NabrianAO"
                txtPasswd = ""
                txtPasswdCheck = ""
                txtPasswd.SetFocus
                Exit Sub
            End If
    
            frmPrincipal.Socket1.HostName = IPdelServidor
            frmPrincipal.Socket1.RemotePort = PuertoDelServidor
    
            Call Audio.PlayWave(0, SND_CLICK)
            Call SendData("TIRDAD")
          '  Me.MousePointer = 11
            EstadoLogin = CrearNuevoPj
    
            If Not frmPrincipal.Socket1.Connected Then
                Call MsgBox("Error: Se ha perdido la conexión con el server.")
                Unload Me
            Else
                Call Login(ValidarLoginMSG(CInt(bRK)))
            End If
            
    
        
          
        End If

End Select

End Sub







Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "crear.jpg")
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmConectar.PictureLogin.Visible = True
frmConectar.txtUser.Visible = True
frmConectar.TxtPass.Visible = True
base_light = D3DColorXRGB(255, 255, 255)
End Sub

Private Sub Label1_Click()
frmPrincipal.Socket1.Disconnect
Call Audio.PlayWave(0, SND_CLICK)
Unload Me
End Sub

'Private Sub txtCorreo_GotFocus()

'MsgBox "Recuerda poner un código de seguridad que te acuerdes, ya que el mismo servira para recuperar o borrar tu personaje en caso de que lo pierdas.", 64, "Código de seguridad"

'End Sub

Private Sub txtPasswd_Change()

If Len(Trim(txtPasswd)) < 6 Then
    lblPass2OK = "s"
    lblPass2OK.ForeColor = &HC0&
    lblPassOK = "s"
    lblPassOK.ForeColor = &HC0&
    Exit Sub
End If

lblPass2OK = "s"
lblPass2OK.ForeColor = &H80FF&

If (txtPasswdCheck = txtPasswd) Then
    lblPassOK = "s"
    lblPassOK.ForeColor = &H80FF&
Else
    lblPassOK = "s"
    lblPassOK.ForeColor = &HC0&
End If

End Sub
Private Sub txtPasswdCheck_Change()

If Len(Trim(txtPasswd)) < 6 Then
    lblPass2OK = "s"
    lblPass2OK.ForeColor = &HC0&
    lblPassOK = "s"
    lblPassOK.ForeColor = &HC0&
    Exit Sub
End If

lblPass2OK = "s"
lblPass2OK.ForeColor = &H80FF&

If (txtPasswdCheck = txtPasswd) Then
    lblPassOK = "s"
    lblPassOK.ForeColor = &H80FF&
Else
    lblPassOK = "s"
    lblPassOK.ForeColor = &HC0&
End If

End Sub
Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub


Private Sub txtNombre_KeyPress(KeyAscii As Integer)
 'KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

