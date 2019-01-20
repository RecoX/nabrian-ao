VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form FormConsola 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1845
   ClientLeft      =   240
   ClientTop       =   585
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormConsola.frx":0000
   ScaleHeight     =   1845
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1480
      Left            =   145
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   165
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   2619
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"FormConsola.frx":286DC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FormConsola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Click()
frmMain.SetFocus
End Sub


Private Sub RecTxt_Click()
frmMain.SetFocus
End Sub

Private Sub RecTxt_DblClick()
frmMain.SetFocus
End Sub

Private Sub RecTxt_KeyPress(KeyAscii As Integer)
frmMain.SetFocus
End Sub

Private Sub RecTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button Then moverForm
End Sub
Public Sub moverForm()
    Dim res As Long
    '
    ReleaseCapture
    res = SendMessage(FormConsola.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
End Sub
