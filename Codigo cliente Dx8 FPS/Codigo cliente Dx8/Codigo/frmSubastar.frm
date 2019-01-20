VERSION 5.00
Begin VB.Form frmSubastar 
   BorderStyle     =   0  'None
   Caption         =   "Subasta"
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox StartBid 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1680
      TabIndex        =   2
      Text            =   "20"
      Top             =   5040
      Width           =   1410
   End
   Begin VB.TextBox cantsubasta 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2040
      TabIndex        =   1
      Text            =   "1"
      Top             =   4545
      Width           =   975
   End
   Begin VB.ListBox ItemList 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3630
      ItemData        =   "frmSubastar.frx":0000
      Left            =   240
      List            =   "frmSubastar.frx":0002
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   165
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   1560
      Top             =   5640
      Width           =   1560
   End
End
Attribute VB_Name = "frmSubastar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Image1_Click()
If Not IsNumeric(cantsubasta.Text) Then Exit Sub
If Not IsNumeric(StartBid.Text) Then Exit Sub
If ItemList.Text = "Nada" Then Exit Sub
 
Call SendData("/INISUB " & ItemList.ListIndex + 1 & " " & cantsubasta.Text & " " & StartBid.Text & "")
Unload Me
 
End Sub
Private Sub Image2_Click()
Unload Me
End Sub


Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "Subastar.jpg")
           For I = 1 To UBound(UserInventory)
           
                        frmSubastar.ItemList.AddItem UserInventory(I).name
                        Next
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
