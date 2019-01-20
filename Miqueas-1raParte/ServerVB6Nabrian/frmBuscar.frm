VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmBuscar 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscador + User on data (WEB)"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Buscar por MOTHER"
      Height          =   375
      Left            =   4920
      TabIndex        =   18
      Top             =   5640
      Width           =   2055
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3975
      Left            =   7200
      TabIndex        =   16
      Top             =   240
      Width           =   4575
      ExtentX         =   8070
      ExtentY         =   7011
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer datosweb 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   5760
      Top             =   6720
   End
   Begin VB.Frame frameChars 
      BackColor       =   &H00000000&
      Caption         =   "Charfiles"
      ForeColor       =   &H00C0C0FF&
      Height          =   4935
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2775
      Begin VB.Timer Timer1 
         Interval        =   30000
         Left            =   2280
         Top             =   120
      End
      Begin VB.CommandButton cmbUpdateChars 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4080
         Width           =   2295
      End
      Begin VB.ListBox lstChars 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3660
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame frameResultados 
      BackColor       =   &H00000000&
      Caption         =   "Resultados"
      ForeColor       =   &H00C0C0FF&
      Height          =   4935
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtProcesado 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   4455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame frameAcciones 
      BackColor       =   &H00000000&
      Caption         =   "Acciones"
      ForeColor       =   &H00C0C0FF&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   6975
      Begin VB.CommandButton Command5 
         Caption         =   "MANUAL"
         Height          =   255
         Left            =   5760
         TabIndex        =   19
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Buscar usuarios por HD"
         Height          =   375
         Left            =   3000
         TabIndex        =   15
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox textweb 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Text            =   "http://nabrianao.000webhostapp.com/"
         Top             =   720
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Prender"
         Height          =   255
         Left            =   3240
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Apagar"
         Height          =   255
         Left            =   4440
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtDatos 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmbMail 
         Caption         =   "Buscar usuario por Código"
         Height          =   375
         Left            =   4800
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmbIP 
         Caption         =   "Busca usuarios por IP"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Estado: APAGADO"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   3015
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   3975
      Left            =   7200
      TabIndex        =   17
      Top             =   4320
      Width           =   3255
      ExtentX         =   5741
      ExtentY         =   7011
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Detector server.exe: AutoActivado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   6600
      Width           =   6975
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4680
      Width           =   2775
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim Nicks(90000) As String ' max 10M de pjs
Dim MaxChar As Long ' ultimo numero de la array nicks
Dim Analizado As Long

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String
    Dim sSpaces As String, szReturn As String
    szReturn = ""
    sSpaces = Space$(EmptySpaces)
    GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, file
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Sub BuscarCodigo(Email As String)
Dim Path As String
Dim Archivo As String
Dim ContarChar As Long
Dim Nick As String
Dim IpPreArray() As String
Dim IpArray(4) As String
Path = App.Path & "\Charfile\"

ContarChar = 1

txtProcesado.Text = ""

lblStatus = "ANALIZANDO CHARFILES - ESPERE"
DoEvents

Do While ContarChar <= MaxChar
    Archivo = GetVar(Path & Nicks(ContarChar - 1), "CONTACTO", "Email")
    If Archivo = Email Then
        txtProcesado.Text = txtProcesado.Text & Nicks(ContarChar - 1) & vbCrLf
    End If
    ContarChar = ContarChar + 1
Loop

lblStatus = ""


End Sub

Sub BuscarMother(Email As String)
Dim Path As String
Dim Archivo As String
Dim ContarChar As Long
Dim Nick As String
Dim IpPreArray() As String
Dim IpArray(4) As String
Path = App.Path & "\Charfile\"

ContarChar = 1

txtProcesado.Text = ""

lblStatus = "ANALIZANDO CHARFILES - ESPERE"
DoEvents

Do While ContarChar <= MaxChar
    Archivo = GetVar(Path & Nicks(ContarChar - 1), "INIT", "Mother")
    If Archivo = Email Then
        txtProcesado.Text = txtProcesado.Text & Nicks(ContarChar - 1) & vbCrLf
    End If
    ContarChar = ContarChar + 1
Loop

lblStatus = ""


End Sub
Sub BuscarIP(IP As String)
Dim Path As String
Dim Archivo As String
Dim ContarChar As Long
Dim Nick As String
Dim IpPreArray() As String
Dim IpArray(4) As String
Path = App.Path & "\Charfile\"

ContarChar = 1

txtProcesado.Text = ""

lblStatus = "ANALIZANDO CHARFILES - ESPERE"
DoEvents

Do While ContarChar <= MaxChar
    Archivo = GetVar(Path & Nicks(ContarChar - 1), "INIT", "LastIP")
    If Archivo = IP Then
        txtProcesado.Text = txtProcesado.Text & Nicks(ContarChar - 1) & vbCrLf
    End If
    ContarChar = ContarChar + 1
Loop

lblStatus = ""


End Sub

Sub BuscarHD(HD As String)
Dim Path As String
Dim Archivo As String
Dim ContarChar As Long
Dim Nick As String
Dim IpPreArray() As String
Dim IpArray(4) As String
Path = App.Path & "\Charfile\"

ContarChar = 1

txtProcesado.Text = ""

lblStatus = "ANALIZANDO CHARFILES - ESPERE"
DoEvents

Do While ContarChar <= MaxChar
    Archivo = GetVar(Path & Nicks(ContarChar - 1), "INIT", "LastHD")
    If Archivo = HD Then
        txtProcesado.Text = txtProcesado.Text & Nicks(ContarChar - 1) & vbCrLf
    End If
    ContarChar = ContarChar + 1
Loop

lblStatus = ""


End Sub

Private Sub cmbIP_Click()
If txtDatos.Text = "" Then
    lblStatus = "FALTAN DATOS"
    DoEvents
    Sleep (500)
    lblStatus = ""
    Exit Sub
End If

BuscarIP (txtDatos.Text)
End Sub

Sub BuscarChars()
Dim FileName As String
Dim count As Long
count = 1

FileName = Dir(App.Path & "\Charfile\*.chr", vbArchive)
lstChars.Clear
Do While FileName <> ""
    lstChars.AddItem FileName
    FileName = Dir
Loop

Do While count <= lstChars.ListCount
Nicks(count - 1) = lstChars.List(count - 1)
''MsgBox Nicks(count - 1)
count = count + 1
Loop

MaxChar = lstChars.ListCount
frameChars.Caption = "Charfiles: " & MaxChar

End Sub

Private Sub cmbMail_Click()

If txtDatos.Text = "" Then
    lblStatus = "FALTAN DATOS"
    DoEvents
    Sleep (500)
    lblStatus = ""
    Exit Sub
End If

BuscarCodigo (txtDatos.Text)
End Sub

Private Sub cmbUpdateChars_Click()
BuscarChars
DoEvents
End Sub

Private Sub Command1_Click()
datosweb.Enabled = True
Label2.Caption = "Estado: Prendido"
End Sub

Private Sub Command2_Click()
datosweb.Enabled = False
Label2.Caption = "Estado: APAGADO"
End Sub

Private Sub Command3_Click()
If txtDatos.Text = "" Then
    lblStatus = "FALTAN DATOS"
    DoEvents
    Sleep (500)
    lblStatus = ""
    Exit Sub
End If

BuscarHD (txtDatos.Text)
End Sub

Private Sub Command4_Click()
If txtDatos.Text = "" Then
    lblStatus = "FALTAN DATOS"
    DoEvents
    Sleep (500)
    lblStatus = ""
    Exit Sub
End If

BuscarMother (txtDatos.Text)
End Sub

Private Sub Command5_Click()
 On Error Resume Next

WebBrowser1.Navigate2 (textweb.Text & "useron.php?keyp=4568&min=" & GetVar(App.Path & "\Dat\UsersOn.siam", "USERSON", "UsuariosOnline") & "&max=" & GetVar(App.Path & "\Dat\UsersOn.siam", "USERSON", "PersonajesCreados"))


End Sub

Private Sub datosweb_Timer()
 On Error Resume Next
WebBrowser1.Navigate2 (textweb.Text & "useron.php?keyp=4568&min=" & GetVar(App.Path & "\Dat\UsersOn.siam", "USERSON", "UsuariosOnline") & "&max=" & GetVar(App.Path & "\Dat\UsersOn.siam", "USERSON", "PersonajesCreados"))


End Sub

Private Sub Form_Load()
BuscarChars
End Sub

Private Sub Label3_Click()
If Timer1.Enabled = False Then
Label3.Caption = "Detector server.exe: ACTIVADO"
Timer1.Enabled = True
ElseIf Timer1.Enabled = True Then
Label3.Caption = "Detector server.exe: DESACTIVADO"
Timer1.Enabled = False
End If
End Sub

Private Sub Timer1_Timer()
Shell "server.exe"
End Sub

