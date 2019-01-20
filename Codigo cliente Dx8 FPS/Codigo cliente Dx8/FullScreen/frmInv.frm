VERSION 5.00
Begin VB.Form frmInv 
   BorderStyle     =   0  'None
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox hlst 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2790
      Left            =   240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   240
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
      Begin VB.Image CmdLanzar 
         Height          =   405
         Index           =   0
         Left            =   0
         MouseIcon       =   "frmInv.frx":0000
         MousePointer    =   99  'Custom
         Top             =   0
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Image cmdInfo 
         Height          =   405
         Index           =   0
         Left            =   1860
         MouseIcon       =   "frmInv.frx":0152
         MousePointer    =   99  'Custom
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image cmd 
         Height          =   405
         Left            =   0
         MouseIcon       =   "frmInv.frx":02A4
         MousePointer    =   99  'Custom
         Top             =   0
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Image cmd2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1950
         Left            =   0
         MouseIcon       =   "frmInv.frx":03F6
         MousePointer    =   99  'Custom
         Picture         =   "frmInv.frx":0548
         Top             =   150
         Visible         =   0   'False
         Width           =   1950
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Move"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   4080
      Width           =   495
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   1
      Left            =   2760
      MouseIcon       =   "frmInv.frx":498A
      MousePointer    =   99  'Custom
      Top             =   1260
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   0
      Left            =   2760
      MouseIcon       =   "frmInv.frx":4ADC
      MousePointer    =   99  'Custom
      Top             =   840
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "frmInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub moverForm()
    Dim res As Long
    '
    ReleaseCapture
    res = SendMessage(FormInv.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
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

Private Sub TirarItem()

    If (ItemElegido > 0 And ItemElegido < MAX_INVENTORY_SLOTS + 1) Or (ItemElegido = FLAGORO) Then
        If UserInventory(ItemElegido).Amount = 1 Then
            SendData "TI" & ItemElegido & "," & 1
        Else
           If UserInventory(ItemElegido).Amount > 1 Then
            frmCantidad.Show
           End If
        End If
    End If

End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image2_Click()

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
moverForm
End Sub
