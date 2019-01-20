VERSION 5.00
Begin VB.Form frmDruida 
   BorderStyle     =   0  'None
   Caption         =   "Alquimia"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000004&
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "1"
      Top             =   3565
      Width           =   1695
   End
   Begin VB.ListBox lstPociones 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2370
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   4095
   End
   Begin VB.Image cmdSalir 
      Height          =   255
      Left            =   0
      MouseIcon       =   "frmDruida.frx":0000
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   735
   End
   Begin VB.Image cmdCrear 
      Height          =   375
      Left            =   2880
      Top             =   3480
      Width           =   1575
   End
End
Attribute VB_Name = "frmDruida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCrear_Click()
On Error Resume Next
Dim stxtCantBuffer As String
stxtCantBuffer = txtCantidad.Text

Call SendData("DCI" & ObjDruida(lstPociones.ListIndex) & " " & stxtCantBuffer)

Unload Me
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "Alquimia.gif")
End Sub

Private Sub txtCantidad_Change()
If Val(txtCantidad.Text) < 0 Then
    txtCantidad.Text = 1
End If

If Val(txtCantidad.Text) > MAX_INVENTORY_OBJS Then
    txtCantidad.Text = 1
End If

If Not IsNumeric(txtCantidad.Text) Then txtCantidad.Text = "1"

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving = False And Button = vbLeftButton Then
    DX = X
    dy = Y
    bmoving = True
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving And ((X <> DX) Or (Y <> dy)) Then Move Left + (X - DX), Top + (Y - dy)

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub
