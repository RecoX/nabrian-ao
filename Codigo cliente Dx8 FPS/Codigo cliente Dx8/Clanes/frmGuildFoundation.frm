VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   0  'None
   Caption         =   "Creación de un Clan"
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4500
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00111720&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox txtClanName 
      Appearance      =   0  'Flat
      BackColor       =   &H00111720&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2090
      Width           =   3375
   End
   Begin VB.Image command2 
      Height          =   375
      Left            =   3720
      MouseIcon       =   "frmGuildFoundation.frx":0000
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   735
   End
   Begin VB.Image command1 
      Height          =   375
      Left            =   2280
      MouseIcon       =   "frmGuildFoundation.frx":030A
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   1575
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FenixAO DirectX8
'Engine By ·Parra
'Arreglado By Thusing
'Algunas cosas tomadas del cliente de DarkTester

Private Sub command1_Click()
ClanName = txtClanName
Site = Text2
frmGuildDetails.Show
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub


Private Sub Form_Load()

If Len(txtClanName.Text) <= 30 Then
    If Not AsciiValidos(txtClanName) Then
        MsgBox "Nombre invalido."
        Exit Sub
    End If
Else
        MsgBox "Nombre demasiado extenso."
        Exit Sub
End If

Me.Picture = LoadPicture(DirGraficos & "GuildCreation.gif")


End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving = False And Button = vbLeftButton Then
    Dx3 = X
    dy = Y
    bmoving = True
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving And ((X <> Dx3) Or (Y <> dy)) Then Move Left + (X - Dx3), Top + (Y - dy)

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub
