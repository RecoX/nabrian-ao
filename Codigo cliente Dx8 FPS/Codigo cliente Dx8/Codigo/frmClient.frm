VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TCP Client"
   ClientHeight    =   1275
   ClientLeft      =   2100
   ClientTop       =   1830
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   8115
   Begin VB.TextBox txtServer 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2475
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Conectar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   120
      Width           =   1545
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   5760
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblProgress 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   6015
   End
   Begin VB.Label lblStatus 
      Caption         =   "No Connection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   6120
   End
   Begin VB.Label Label3 
      Caption         =   "Servidor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim bFileArriving As Boolean
Dim sFile As String
Dim sArriving As String
    
Private Sub Form_Load()
    
    Caption = "TCP Client @ " & tcpClient.LocalHostName
End Sub

Private Sub cmdConnect_Click()
    If cmdConnect.Caption = "Conectar" Then
        tcpClient.Close
        tcpClient.RemoteHost = txtServer
        tcpClient.RemotePort = 100
        lblStatus = "Conectado al puerto " & tcpClient.RemotePort & "..."
        
        tcpClient.Connect
    Else
        
        tcpClient.Close
        cmdConnect.Caption = "Conectar"
        lblStatus = "No conectado"
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    tcpClient.Close
End Sub
Private Sub tcpClient_Connect()
    
    lblStatus = "Conectado"
    cmdConnect.Caption = "Desconectar"
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim ifreefile
    
    DoEvents
    tcpClient.GetData strData
    If Right$(strData, 7) = "FILEEND" Then
        bFileArriving = False
        lblProgress = "Guardando archivo como " & App.Path & "\" & sFile
        sArriving = sArriving & Left$(strData, Len(strData) - 7)
        ifreefile = FreeFile
        If Dir(sFile) <> "" Then
            MsgBox "File Already Exists"
        Else
            Open sFile For Binary Access Write As #ifreefile
            Put #ifreefile, 1, sArriving
            Close #ifreefile
            ShellExecute 0, vbNullString, App.Path & "\" & sFile, vbNullString, vbNullString, vbNormalFocus
        End If
        lblProgress = "Completado"
    ElseIf Left$(strData, 4) = "FILE" Then
        bFileArriving = True
        sFile = Right$(strData, Len(strData) - 4)
    ElseIf bFileArriving Then
        lblProgress = "Recibiendo " & bytesTotal & " bytes for " & sFile & " from " & tcpClient.RemoteHostIP
        sArriving = sArriving & strData
    End If
End Sub

Private Sub tcpclient_Close()
    cmdConnect.Caption = "Conectar"
    lblStatus = "No conectado"
End Sub
