VERSION 5.00
Begin VB.Form frmVerSoporte 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "soporte"
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   Picture         =   "frmVerSoporte.frx":0000
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblR 
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4035
   End
   Begin VB.Image imgCerrar 
      Height          =   495
      Left            =   1080
      MouseIcon       =   "frmVerSoporte.frx":6B39
      Top             =   2880
      Width           =   2175
   End
End
Attribute VB_Name = "frmVerSoporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub imgCerrar_Click()
lblR.Caption = ""
Me.Hide
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If bmoving = False And Button = vbLeftButton Then
      Dx3 = X
      dy = Y
      bmoving = True
   End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If bmoving And ((X <> Dx3) Or (Y <> dy)) Then
      Move Left + (X - Dx3), Top + (Y - dy)
   End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      bmoving = False
   End If
End Sub

