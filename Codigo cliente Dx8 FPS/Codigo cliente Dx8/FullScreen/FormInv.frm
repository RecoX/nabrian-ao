VERSION 5.00
Begin VB.Form FormInv 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   ClientHeight    =   3360
   ClientLeft      =   8955
   ClientTop       =   3570
   ClientWidth     =   2865
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   2865
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox inventariofd 
      BackColor       =   &H80000008&
      Height          =   2535
      Left            =   170
      ScaleHeight     =   2475
      ScaleWidth      =   2475
      TabIndex        =   2
      Top             =   175
      Width           =   2535
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   25
         Left            =   2265
         TabIndex        =   52
         Top             =   2235
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   25
         Left            =   1920
         TabIndex        =   51
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   25
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   1935
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   24
         Left            =   1785
         TabIndex        =   50
         Top             =   2235
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   24
         Left            =   1440
         TabIndex        =   49
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   24
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   1935
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   23
         Left            =   1305
         TabIndex        =   48
         Top             =   2235
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   23
         Left            =   960
         TabIndex        =   47
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   23
         Left            =   960
         Stretch         =   -1  'True
         Top             =   1935
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   22
         Left            =   825
         TabIndex        =   46
         Top             =   2235
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   22
         Left            =   480
         TabIndex        =   45
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   22
         Left            =   480
         Stretch         =   -1  'True
         Top             =   1935
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   21
         Left            =   345
         TabIndex        =   44
         Top             =   2235
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   21
         Left            =   0
         TabIndex        =   43
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   21
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1935
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   20
         Left            =   2265
         TabIndex        =   42
         Top             =   1740
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   20
         Left            =   1920
         TabIndex        =   41
         Top             =   1425
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   20
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   19
         Left            =   1785
         TabIndex        =   40
         Top             =   1740
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   19
         Left            =   1440
         TabIndex        =   39
         Top             =   1425
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   19
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   18
         Left            =   1305
         TabIndex        =   38
         Top             =   1740
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   18
         Left            =   960
         TabIndex        =   37
         Top             =   1425
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   18
         Left            =   960
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   17
         Left            =   825
         TabIndex        =   36
         Top             =   1740
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   17
         Left            =   480
         TabIndex        =   35
         Top             =   1425
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   17
         Left            =   480
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   16
         Left            =   345
         TabIndex        =   34
         Top             =   1740
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   16
         Left            =   0
         TabIndex        =   33
         Top             =   1425
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   16
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   15
         Left            =   2265
         TabIndex        =   32
         Top             =   1275
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   15
         Left            =   1920
         TabIndex        =   31
         Top             =   960
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   15
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   975
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   14
         Left            =   1785
         TabIndex        =   30
         Top             =   1275
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   14
         Left            =   1440
         TabIndex        =   29
         Top             =   960
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   14
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   975
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   13
         Left            =   1305
         TabIndex        =   28
         Top             =   1275
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   13
         Left            =   960
         TabIndex        =   27
         Top             =   960
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   13
         Left            =   960
         Stretch         =   -1  'True
         Top             =   975
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   12
         Left            =   825
         TabIndex        =   26
         Top             =   1275
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   12
         Left            =   480
         TabIndex        =   25
         Top             =   960
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   12
         Left            =   480
         Stretch         =   -1  'True
         Top             =   975
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   11
         Left            =   345
         TabIndex        =   24
         Top             =   1275
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   11
         Left            =   0
         TabIndex        =   23
         Top             =   960
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   11
         Left            =   0
         Stretch         =   -1  'True
         Top             =   975
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   10
         Left            =   2265
         TabIndex        =   22
         Top             =   795
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   10
         Left            =   1920
         TabIndex        =   21
         Top             =   480
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   10
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   495
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   9
         Left            =   1785
         TabIndex        =   20
         Top             =   795
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   9
         Left            =   1440
         TabIndex        =   19
         Top             =   480
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   9
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   495
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   8
         Left            =   1305
         TabIndex        =   18
         Top             =   795
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   8
         Left            =   960
         TabIndex        =   17
         Top             =   480
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   8
         Left            =   960
         Stretch         =   -1  'True
         Top             =   495
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   7
         Left            =   825
         TabIndex        =   16
         Top             =   795
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   7
         Left            =   480
         TabIndex        =   15
         Top             =   480
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   7
         Left            =   480
         Stretch         =   -1  'True
         Top             =   495
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   6
         Left            =   345
         TabIndex        =   14
         Top             =   795
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   6
         Left            =   0
         TabIndex        =   13
         Top             =   480
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   6
         Left            =   0
         Stretch         =   -1  'True
         Top             =   495
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   5
         Left            =   2265
         TabIndex        =   12
         Top             =   315
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   5
         Left            =   1920
         TabIndex        =   11
         Top             =   0
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   5
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   15
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   4
         Left            =   1785
         TabIndex        =   10
         Top             =   315
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   4
         Left            =   1440
         TabIndex        =   9
         Top             =   0
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   4
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   15
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   3
         Left            =   1305
         TabIndex        =   8
         Top             =   315
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   960
         TabIndex        =   7
         Top             =   0
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   3
         Left            =   960
         Stretch         =   -1  'True
         Top             =   15
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   2
         Left            =   825
         TabIndex        =   6
         Top             =   315
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   480
         TabIndex        =   5
         Top             =   0
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   2
         Left            =   480
         Stretch         =   -1  'True
         Top             =   15
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   1
         Left            =   345
         TabIndex        =   4
         Top             =   315
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   15
         Width           =   480
      End
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2340
      Left            =   360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   150
      Left            =   1200
      TabIndex        =   0
      Top             =   2880
      Width           =   75
   End
   Begin VB.Image CmdLanzar 
      Height          =   285
      Left            =   240
      MouseIcon       =   "FormInv.frx":0000
      MousePointer    =   99  'Custom
      Top             =   2880
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image cmdInfo 
      Height          =   285
      Left            =   1440
      MouseIcon       =   "FormInv.frx":0152
      MousePointer    =   99  'Custom
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   135
      Index           =   0
      Left            =   2400
      MouseIcon       =   "FormInv.frx":02A4
      MousePointer    =   99  'Custom
      Top             =   360
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   135
      Index           =   1
      Left            =   2400
      MouseIcon       =   "FormInv.frx":03F6
      MousePointer    =   99  'Custom
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "FormInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub moverForm()

End Sub

Private Sub cmdInfo_Click()
    Call SendData("INFS" & hlst.ListIndex + 1)
frmMain.SetFocus
End Sub

Private Sub CmdLanzar_Click()
    If hlst.List(hlst.ListIndex) <> "(None)" And TiempoTranscurrido(LastHechizo) >= IntervaloSpell And TiempoTranscurrido(Hechi) >= IntervaloSpell / 4 Then
    Call SendData("LH" & hlst.ListIndex + 1)
    Call SendData("UK" & Magia)
    End If
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)
If hlst.ListIndex = -1 Then Exit Sub

Select Case Index
Case 0 'subir
    If hlst.ListIndex = 0 Then Exit Sub
Case 1 'bajar
    If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
End Select

Call SendData("DESPHE" & Index + 1 & "," & hlst.ListIndex + 1)

Select Case Index
Case 0 'subir
    hlst.ListIndex = hlst.ListIndex - 1
Case 1 'bajar
    hlst.ListIndex = hlst.ListIndex + 1
End Select

frmMain.SetFocus


End Sub

Private Sub Form_Activate()
If frmMain.Visible = True Then frmMain.SetFocus
End Sub


Private Sub Form_Click()
frmMain.SetFocus
End Sub

Private Sub Form_DblClick()
frmMain.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'moverForm
If Button Then

    Dim res As Long
    ReleaseCapture
    res = SendMessage(FormInv.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
End If
End Sub

Private Sub GldLbl_Click()
ItemElegido = FLAGORO
If UserGLD > 0 Then frmCantidad.Show
frmCantidad.SetFocus
End Sub

Private Sub hlst_Click()
frmMain.SetFocus
End Sub

Private Sub hlst_DblClick()
frmMain.SetFocus
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub


Private Sub imgObjeto_Click(Index As Integer)
If ItemElegido <> Index And UserInventory(Index).Name <> "Nada" Then
    Shape1.Visible = True
    hlst.Visible = False
    Shape1.Top = imgObjeto(Index).Top '+ Inventariofd.Top
    Shape1.Left = imgObjeto(Index).Left '- Inventariofd.Left
    ItemElegido = Index

End If
 frmMain.SetFocus
    
End Sub
Private Sub imgObjeto_DblClick(Index As Integer)

If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
If ItemElegido = Index Then Call SendData("USE" & ItemElegido)
frmMain.SetFocus
End Sub

