VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ClientFrm 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ChatBox - estado: desconectado"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox textnick 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "WST_Czec"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   840
      TabIndex        =   9
      Text            =   "nickname"
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "WST_Czec"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "WST_Czec"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   3735
   End
   Begin VB.CommandButton bntSend 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4680
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton bntConnect 
      BackColor       =   &H000000FF&
      Caption         =   "Conectar"
      Height          =   375
      Left            =   3480
      MaskColor       =   &H0000FFFF&
      TabIndex        =   4
      Tag             =   "Connect"
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "WST_Czec"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Text            =   "123"
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "WST_Czec"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Nombre:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Puerto"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "IP remota"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "ClientFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bntConnect_Click()
On Error GoTo ErrSub

    With Winsock1
        .Close
        .RemoteHost = txtIP
        .RemotePort = txtPort
        .Connect
    End With
Exit Sub
ErrSub:
MsgBox "Error : " & Err.Description, vbCritical
End Sub


Private Sub bntSend_Click()
On Error GoTo ErrSub


    Winsock1.SendData txtSend

    txtLog = txtLog & textnick.Text & ": " & txtSend & vbCrLf
    txtSend = ""

Exit Sub
ErrSub:
MsgBox "Error : " & Err.Description
Winsock1_Close ' cierra la conexión
End Sub



Private Sub txtLog_Change()

End Sub

Private Sub Winsock1_Close()

    Winsock1.Close  'Cierra la conexión
    txtLog = txtLog & "*** Desconectado" & vbCrLf

End Sub

Private Sub Winsock1_Connect()

txtLog = "Conectado a " & Winsock1.RemoteHostIP & vbCrLf

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim dat As String
    
    Winsock1.GetData dat, vbString
    txtLog = txtLog & "Servidor: " & dat & vbCrLf
    ClientFrm.Caption = "ChatBox - estado: conectado"

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, _
                           Description As String, _
                           ByVal Scode As Long, _
                           ByVal Source As String, _
                           ByVal HelpFile As String, _
                           ByVal HelpContext As Long, _
                           CancelDisplay As Boolean)

    txtLog = txtLog & "*** Error : " & Description & vbCrLf

    Winsock1_Close
End Sub
