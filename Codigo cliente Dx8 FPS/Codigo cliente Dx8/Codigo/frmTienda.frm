VERSION 5.00
Begin VB.Form frmTienda 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   463
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004DC488&
      Height          =   3930
      Index           =   0
      ItemData        =   "frmTienda.frx":0000
      Left            =   720
      List            =   "frmTienda.frx":0002
      TabIndex        =   3
      Top             =   2040
      Width           =   2490
   End
   Begin VB.TextBox precio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004DC488&
      Height          =   285
      Left            =   4440
      TabIndex        =   10
      Text            =   "0"
      Top             =   6645
      Width           =   1080
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004DC488&
      Height          =   3930
      Index           =   1
      Left            =   3840
      TabIndex        =   2
      Top             =   2040
      Width           =   2490
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   720
      ScaleHeight     =   570
      ScaleWidth      =   525
      TabIndex        =   1
      Top             =   720
      Width           =   555
   End
   Begin VB.TextBox cantidad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004DC488&
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   "1"
      Top             =   6645
      Width           =   600
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   6090
      MouseIcon       =   "frmTienda.frx":0004
      MousePointer    =   99  'Custom
      Top             =   6930
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   9
      Top             =   1035
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   5
      Left            =   1320
      TabIndex        =   8
      Top             =   1650
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   4080
      TabIndex        =   7
      Top             =   1275
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   3
      Left            =   4080
      TabIndex        =   6
      Top             =   1500
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   0
      Left            =   840
      MouseIcon       =   "frmTienda.frx":030E
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6150
      Width           =   2460
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   1
      Left            =   3840
      MouseIcon       =   "frmTienda.frx":0618
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6165
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   1335
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   2
      Left            =   4080
      TabIndex        =   4
      Top             =   1050
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "frmTienda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LastIndex1 As Integer
Public LastIndex2 As Integer
Private Sub Image1_Click(Index As Integer)

Call PlayWaveDS(SND_CLICK)

If List1(Index).List(List1(Index).ListIndex) = "Nada" Or List1(Index).ListIndex < 0 Then Exit Sub

Select Case Index
    Case 0
        frmTienda.List1(0).SetFocus
        LastIndex1 = List1(0).ListIndex
        Lista = 0
        Call SendData("SAVE" & List1(0).ListIndex + 1 & "," & cantidad.Text)
        
   Case 1
        LastIndex2 = List1(1).ListIndex
        If UserInventory(List1(1).ListIndex + 1).Equipped = 0 Then
            Lista = 1
            Call SendData("POVE" & List1(1).ListIndex + 1 & "," & cantidad.Text & "," & precio.Text)
        Else
            Call AddtoRichTextBox(frmMain.rectxt, "No podes poner a la venta el item porque lo estás usando.", 2, 51, 223, 1, 1)
            Exit Sub
        End If
End Select

End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
    Case 0
        If Image1(0).Tag = 1 Then
            Image1(0).Tag = 0
            Image1(1).Tag = 1
        End If
        
    Case 1
        If Image1(1).Tag = 1 Then
            Image1(1).Tag = 0
            Image1(0).Tag = 1
        End If
        
End Select

End Sub
Private Sub Image2_Click()
SendData ("FINTIE")
End Sub
Private Sub cantidad_Change()

If Val(cantidad.Text) < 0 Then
    cantidad.Text = 1
End If

If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
    cantidad.Text = 1
End If

End Sub
Private Sub cantidad_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If

End Sub
Private Sub Form_Deactivate()

Me.SetFocus

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving = False And Button = vbLeftButton Then
   DX = X
   dy = Y
   bmoving = True
End If

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub
Private Sub Form_Load()

frmTienda.Picture = LoadPicture(App.Path & "\Graficos\Tienda.jpg")

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Image1(0).Tag = 0 Then
    Image1(0).Tag = 1
End If

If Image1(1).Tag = 0 Then
    Image1(1).Tag = 1
End If

If bmoving And ((X <> DX) Or (Y <> dy)) Then Move Left + (X - DX), Top + (Y - dy)

End Sub
Private Sub List1_Click(Index As Integer)

Lista = Index
Call ActualizarInformacionTienda(Index)

End Sub
Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyE:
        If List1(1).ListIndex > -1 And List1(1).ListIndex < MAX_INVENTORY_SLOTS - 1 Then
            Call SendData("EQUI" & List1(1).ListIndex + 1)
        End If
End Select

End Sub

Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Image1(0).Tag = 0 Then Image1(0).Tag = 1
If Image1(1).Tag = 0 Then Image1(1).Tag = 1

End Sub
