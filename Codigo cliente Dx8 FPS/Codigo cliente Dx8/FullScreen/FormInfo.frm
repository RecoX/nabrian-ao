VERSION 5.00
Begin VB.Form FormInfo 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   ClientHeight    =   1800
   ClientLeft      =   7935
   ClientTop       =   510
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   Picture         =   "FormInfo.frx":0000
   ScaleHeight     =   1800
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   Begin VB.Label ExpTotal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2400
      TabIndex        =   7
      Top             =   780
      Width           =   255
   End
   Begin VB.Label cantidadmana 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2400
      TabIndex        =   6
      Top             =   1350
      Width           =   255
   End
   Begin VB.Label cantidadhambre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1920
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label cantidadsta 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2520
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label cantidadagua 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1560
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label cantidadhp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label LblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Tu nick"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   230
      Width           =   2535
   End
   Begin VB.Label LvlLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   530
      Width           =   105
   End
   Begin VB.Shape ShapeExp 
      BackColor       =   &H009DCAE7&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080C0FF&
      Height          =   180
      Left            =   1583
      Top             =   800
      Width           =   1995
   End
   Begin VB.Shape Hpshp 
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   1583
      Top             =   1114
      Width           =   1995
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   180
      Left            =   1320
      Top             =   4200
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   180
      Left            =   1583
      Top             =   1380
      Width           =   1995
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   180
      Left            =   1200
      Top             =   4200
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   180
      Left            =   840
      Top             =   4200
      Visible         =   0   'False
      Width           =   1995
   End
End
Attribute VB_Name = "FormInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub moverForm()
    Dim res As Long
    '
    ReleaseCapture
    res = SendMessage(Me.hWnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
End Sub

Private Sub cantidadagua_Click()

End Sub

Private Sub cantidadmana_Click()

End Sub

Private Sub Form_Click()
frmMain.SetFocus
End Sub

Private Sub Form_DblClick()
frmMain.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button Then moverForm
End Sub

Private Sub LvlLbl_Click()

End Sub
