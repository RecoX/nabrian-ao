VERSION 5.00
Begin VB.Form frmConectar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerConexion 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   6240
   End
   Begin VB.TextBox TxtPass 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4320
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3840
      Width           =   3330
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4320
      MaxLength       =   20
      TabIndex        =   0
      Top             =   3045
      Width           =   3330
   End
   Begin VB.PictureBox renderconnect 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   9000
      Left            =   0
      Picture         =   "frmConnect.frx":000C
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   2
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox PictureLogin 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5145
         Left            =   3720
         ScaleHeight     =   5145
         ScaleWidth      =   4500
         TabIndex        =   5
         Top             =   1560
         Width           =   4500
         Begin VB.CheckBox Check1 
            Height          =   225
            Left            =   1200
            TabIndex        =   6
            Top             =   2880
            Width           =   210
         End
         Begin VB.Timer Timer1 
            Interval        =   100
            Left            =   600
            Top             =   2760
         End
         Begin VB.Label labelconex 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   4800
            Width           =   4575
         End
         Begin VB.Image Image3 
            Height          =   375
            Left            =   240
            Top             =   0
            Width           =   4215
         End
         Begin VB.Image Image1 
            Height          =   540
            Index           =   1
            Left            =   360
            MouseIcon       =   "frmConnect.frx":46F2
            MousePointer    =   99  'Custom
            Top             =   3360
            Width           =   3855
         End
         Begin VB.Image Image1 
            Height          =   555
            Index           =   0
            Left            =   360
            MouseIcon       =   "frmConnect.frx":49FC
            MousePointer    =   99  'Custom
            Top             =   4080
            Width           =   3825
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   555
         ScaleHeight     =   750
         ScaleWidth      =   10905
         TabIndex        =   4
         Top             =   7920
         Width           =   10905
         Begin VB.Image Image5 
            Height          =   510
            Left            =   8040
            Top             =   120
            Width           =   2580
         End
         Begin VB.Image Image4 
            Height          =   510
            Left            =   5400
            Top             =   120
            Width           =   2460
         End
         Begin VB.Image Image1 
            Height          =   510
            Index           =   3
            Left            =   2520
            MouseIcon       =   "frmConnect.frx":4D06
            MousePointer    =   99  'Custom
            Top             =   120
            Width           =   2580
         End
         Begin VB.Image Image1 
            Height          =   510
            Index           =   2
            Left            =   240
            MouseIcon       =   "frmConnect.frx":5010
            MousePointer    =   99  'Custom
            Top             =   120
            Width           =   2100
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   11640
         ScaleHeight     =   375
         ScaleWidth      =   345
         TabIndex        =   3
         Top             =   0
         Width           =   345
         Begin VB.Image Image2 
            Height          =   360
            Left            =   0
            Picture         =   "frmConnect.frx":531A
            Top             =   0
            Width           =   360
         End
      End
   End
End
Attribute VB_Name = "frmConectar"
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
Option Explicit
Dim yasepusolapass As Integer

 Const ENCRYPT = 1
 Const DECRYPT = 2
Private Sub Form_KeyPress(KeyAscii As Integer)
If IntervaloConexionLogin > 3 Then Exit Sub
If KeyAscii = vbKeyReturn Then
    Call Audio.PlayWave(0, SND_CLICK)
    
    If Check1.value = 1 Then
        Call WriteVar(App.Path & "\INIT\Recordar.dat", txtUser.Text, "Nombre", txtUser.Text)
        Call WriteVar(App.Path & "\INIT\Recordar.dat", txtUser.Text, "Password", EncryptString("smmmmmm", TxtPass.Text, ENCRYPT))
    End If
       
    If frmPrincipal.Socket1.Connected Then frmPrincipal.Socket1.Disconnect
    
    If frmConectar.MousePointer = 11 Then
    frmConectar.MousePointer = 1
        Exit Sub
    End If
    
    
    UserName = txtUser.Text
    Dim aux As String
    aux = TxtPass.Text
    UserPassword = MD5String(aux)
    If CheckUserData(False) = True Then
        frmPrincipal.Socket1.HostName = IPdelServidor
        frmPrincipal.Socket1.RemotePort = PuertoDelServidor
        
        EstadoLogin = Normal
        'Me.MousePointer = 11
        frmPrincipal.Socket1.Connect
    End If
End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then

    If MsgBox("¿Seguro que deseas salir?", vbYesNo, "Resolución") = vbYes Then
    frmCargando.Show
    frmCargando.Refresh
    AddtoRichTextBox frmCargando.Status, "Cerrando NabrianAO.", 150, 150, 150, 0, 0
    
    Call SaveGameini
    frmConectar.MousePointer = 1
    frmPrincipal.MousePointer = 1
    prgRun = False
    
    AddtoRichTextBox frmCargando.Status, "Liberando recursos..."
    frmCargando.Refresh
    AddtoRichTextBox frmCargando.Status, "OK!", 11, 213, 105, 1, 0
    AddtoRichTextBox frmCargando.Status, "¡¡Gracias por jugar NabrianAO vuelve pronto!!", 255, 255, 255, 0, 0
    frmCargando.Refresh
    Call UnloadAllForms
        End
        Else
        End If

End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
    

    
    
    


    
    
    KeyCode = 0
    Exit Sub
End If

End Sub


Private Sub Check1_Click()
    If Check1.value = 0 Then
    Call WriteVar(App.Path & "\INIT\Recordar.dat", txtUser.Text, "Nombre", "")
    Call WriteVar(App.Path & "\INIT\Recordar.dat", txtUser.Text, "Password", "")
    txtUser.Text = ""
    TxtPass.Text = ""
    yasepusolapass = 0
    End If
End Sub

Private Sub Form_Load()
yasepusolapass = 0
PictureLogin.Picture = LoadPicture(DirGraficos & "conectar.jpg")
Picture1.Picture = LoadPicture(DirGraficos & "menuconectar.jpg")
frmConectar.Icon = frmPrincipal.Icon
frmConectar.Caption = "NabrianAO - " & RandomNumber(2000, 3000)
base_light = D3DColorXRGB(255, 255, 255) 'volvemos por si esta muerto el culeao.
Call SwitchMapNew(RandomNumber(78, 79)) 'renderizamo conectar
UserMinHP = 1 'conectar renderizado pa que se vea a color
EngineRun = False

'mciExecute "Close All"
'Call Audio.StopMidi
'Call Audio.PlayWave(1, "intro.mp3")
    
 Dim j
 For Each j In Image1()
    j.Tag = "0"
 Next

 IntervaloPaso = 0.19
 IntervaloUsar = 0.14
 'Picture = LoadPicture(DirGraficos & "fondoconectar.jpg")


End Sub

Private Sub Image1_Click(Index As Integer)
If IntervaloConexionLogin > 3 Then Exit Sub
CurServer = 0


        
Select Case Index
    Case 0
        frmPrincipal.Socket1.HostName = IPdelServidor
        frmPrincipal.Socket1.RemotePort = PuertoDelServidor
        'Me.MousePointer = 11
        EstadoLogin = dados
        frmPrincipal.Socket1.Connect
        
    Case 1
    
       If Check1.value = 1 Then
        Call WriteVar(App.Path & "\INIT\Recordar.dat", txtUser.Text, "Nombre", txtUser.Text)
        Call WriteVar(App.Path & "\INIT\Recordar.dat", txtUser.Text, "Password", EncryptString("smmmmmm", TxtPass.Text, ENCRYPT))
       End If
    
        If frmPrincipal.Socket1.Connected Then frmPrincipal.Socket1.Disconnect
        
        If frmConectar.MousePointer = 11 Then
        frmConectar.MousePointer = 1
            Exit Sub
        End If
        
        
        
        UserName = txtUser.Text
        Dim aux As String
        aux = TxtPass.Text
        UserPassword = MD5String(aux)
        If CheckUserData(False) = True Then
            frmPrincipal.Socket1.HostName = IPdelServidor
            frmPrincipal.Socket1.RemotePort = PuertoDelServidor
            
            EstadoLogin = Normal
           ' Me.MousePointer = 11
            frmPrincipal.Socket1.Connect
        End If
        
Case 2
If frmPrincipal.Socket1.Connected Then frmPrincipal.Socket1.Disconnect
     
If frmConectar.MousePointer = 11 Then
frmConectar.MousePointer = 1
Exit Sub
End If
     
frmPrincipal.Socket1.HostName = IPdelServidor
frmPrincipal.Socket1.RemotePort = PuertoDelServidor
EstadoLogin = BorrarPj
'Me.MousePointer = 11
frmPrincipal.Socket1.Connect
Case 3
If frmPrincipal.Socket1.Connected Then frmPrincipal.Socket1.Disconnect
     
If frmConectar.MousePointer = 11 Then
frmConectar.MousePointer = 1
Exit Sub
End If
     
frmPrincipal.Socket1.HostName = IPdelServidor
frmPrincipal.Socket1.RemotePort = PuertoDelServidor
EstadoLogin = RecuperarPass
'Me.MousePointer = 11
frmPrincipal.Socket1.Connect

End Select

End Sub
Private Sub Image2_Click()
        If MsgBox("¿Seguro que deseas salir?", vbYesNo, "Cambio de resolución.") = vbYes Then
        End
        Else
        End If
End Sub
Private Sub imgGetPass_Click()
 
If frmPrincipal.Socket1.Connected Then frmPrincipal.Socket1.Disconnect
 
If frmConectar.MousePointer = 11 Then
    frmConectar.MousePointer = 1
    Exit Sub
End If
 
EstadoLogin = RecuperarPass
'Me.MousePointer = 11
frmPrincipal.Socket1.Connect
 
End Sub


Private Sub Image3_Click()
Call ShellExecute(Me.hwnd, "open", "http://foro.nabrianao.net", "", "", 1)
End Sub


Private Sub Image4_Click()
Call ShellExecute(Me.hwnd, "open", "http://www.nabrianao.net/manual/manual.html", "", "", 1)
End Sub

Private Sub Image5_Click()
Call ShellExecute(Me.hwnd, "open", "http://foro.nabrianao.net/", "", "", 1)
End Sub


Private Sub TimerConexion_Timer()
IntervaloConexionLogin = IntervaloConexionLogin - 1

If IntervaloConexionLogin = 6 Then
labelconex.Caption = "Debes esperar 3 segundos para volver a conectar."
ElseIf IntervaloConexionLogin = 5 Then
labelconex.Caption = "Debes esperar 2 segundos para volver a conectar."
ElseIf IntervaloConexionLogin = 4 Then
labelconex.Caption = "Debes esperar 1 segundos para volver a conectar."
ElseIf IntervaloConexionLogin = 3 Then
labelconex.Caption = ""
IntervaloConexionLogin = 0
frmConectar.TimerConexion = False
End If
End Sub

Private Sub txtUser_Change()
yasepusolapass = 0
End Sub

Private Sub Timer1_Timer()
    If yasepusolapass = 1 Then Exit Sub
    Dim usuariorecordado As String
    
    usuariorecordado = GetVar(App.Path & "\INIT\Recordar.dat", txtUser.Text, "Nombre")
    If UCase$(usuariorecordado) = UCase$(txtUser.Text) Then
    If txtUser.Text = "" Then Exit Sub
    TxtPass.Text = EncryptString("smmmmmm", GetVar(App.Path & "\INIT\Recordar.dat", txtUser.Text, "Password"), DECRYPT)
    Check1.value = 1
    yasepusolapass = 1
    End If
End Sub


Public Function EncryptString( _
    UserKey As String, Text As String, Action As Single _
    ) As String
    Dim UserKeyX As String
    Dim Temp     As Integer
    Dim Times    As Integer
    Dim I        As Integer
    Dim j        As Integer
    Dim N        As Integer
    Dim rtn      As String
      
    '//Get UserKey characters
    N = Len(UserKey)
    ReDim UserKeyASCIIS(1 To N)
    For I = 1 To N
        UserKeyASCIIS(I) = Asc(mid$(UserKey, I, 1))
    Next
          
    '//Get Text characters
    ReDim TextASCIIS(Len(Text)) As Integer
    For I = 1 To Len(Text)
        TextASCIIS(I) = Asc(mid$(Text, I, 1))
    Next
      
    '//Encryption/Decryption
    If Action = ENCRYPT Then
       For I = 1 To Len(Text)
           j = IIf(j + 1 >= N, 1, j + 1)
           Temp = TextASCIIS(I) + UserKeyASCIIS(j)
           If Temp > 255 Then
              Temp = Temp - 255
           End If
           rtn = rtn + Chr$(Temp)
       Next
    ElseIf Action = DECRYPT Then
       For I = 1 To Len(Text)
           j = IIf(j + 1 >= N, 1, j + 1)
           Temp = TextASCIIS(I) - UserKeyASCIIS(j)
           If Temp < 0 Then
              Temp = Temp + 255
           End If
           rtn = rtn + Chr$(Temp)
       Next
    End If
      
    '//Return
    EncryptString = rtn
End Function


