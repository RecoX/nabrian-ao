VERSION 5.00
Begin VB.Form FormBarInv 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1170
   ClientLeft      =   7530
   ClientTop       =   7545
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MouseIcon       =   "barraInv.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "barraInv.frx":08CA
   ScaleHeight     =   1170
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Begin VB.Shape Shape 
      BorderColor     =   &H80000004&
      BorderStyle     =   0  'Transparent
      Height          =   585
      Index           =   0
      Left            =   300
      Shape           =   3  'Circle
      Top             =   300
      Width           =   615
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H80000004&
      BorderStyle     =   0  'Transparent
      Height          =   585
      Index           =   4
      Left            =   3300
      Shape           =   3  'Circle
      Top             =   300
      Width           =   735
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H80000004&
      BorderStyle     =   0  'Transparent
      Height          =   585
      Index           =   3
      Left            =   2530
      Shape           =   3  'Circle
      Top             =   300
      Width           =   735
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H80000004&
      BorderStyle     =   0  'Transparent
      Height          =   585
      Index           =   2
      Left            =   1775
      Shape           =   3  'Circle
      Top             =   300
      Width           =   735
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H80000004&
      BorderStyle     =   0  'Transparent
      Height          =   585
      Index           =   1
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "FormBarInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.SetFocus
If Not Y >= 300 Then Exit Sub
If Y >= 900 Then Exit Sub
For a = 0 To 4
xDxD = Shape(a).Left + Shape(a).Height
If Shape(a).Left + Shape(a).Height >= X Then
Select Case a
Case 0
If FormConsola.Visible = False Then
FormConsola.Visible = True
FormConsola.Show , frmMain
FormBarInv.Shape(a).BorderStyle = 1
FormBarInv.Shape(a).BorderColor = vbYellow
Else
FormConsola.Visible = False
FormConsola.Hide
FormBarInv.Shape(a).BorderStyle = 0
FormBarInv.Shape(a).BorderColor = vbWhite
End If

Case 1
bInvMod = True
If Not FormInv.Visible = True Or FormInv.CmdLanzar.Visible = True Then
FormInv.Visible = True
FormInv.Show , frmMain
FormBarInv.Shape(a).BorderStyle = 1
FormBarInv.Shape(a).BorderColor = vbYellow
With FormInv
   .Picture = FormImagenes.Picture1.Picture 'LoadPicture(App.Path & "\Graficos\Centronuevoinventario.jpg")

    'DespInv(0).Visible = True
    'DespInv(1).Visible = True
   '.picInv.Visible = True
    .inventariofd.Visible = True
   .hlst.Visible = False
   .cmdInfo.Visible = False
    .CmdLanzar.Visible = False
    
    .cmdMoverHechi(0).Visible = True
    .cmdMoverHechi(1).Visible = True
    .GldLbl.Visible = True
    End With
Else

FormInv.Visible = False
FormInv.Hide
FormBarInv.Shape(a).BorderStyle = 0
FormBarInv.Shape(a).BorderColor = vbWhite
FormBarInv.Shape(a + 1).BorderStyle = 0
FormBarInv.Shape(a + 1).BorderColor = vbWhite
End If
    bInvMod = True
Case 2
    bInvMod = True
If Not FormInv.Visible = True Or FormInv.CmdLanzar.Visible = False Then
FormInv.Visible = True
bInvMod = True
FormInv.Show , frmMain
FormBarInv.Shape(a).BorderStyle = 1
FormBarInv.Shape(a).BorderColor = vbYellow
'FormBarInv.Shape(A + 1).BorderStyle = 1
'FormBarInv.Shape(A + 1).BorderColor = vbYellow
    With FormInv
    .Picture = FormImagenes.Picture2.Picture 'LoadPicture(App.Path & "\Graficos\Centronuevohechizos.jpg")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    'DespInv(0).Visible = False
    'DespInv(1).Visible = False
    .inventariofd.Visible = False
    .hlst.Visible = True
    .cmdInfo.Visible = True
    .CmdLanzar.Visible = True
    
    .cmdMoverHechi(0).Visible = True
    .cmdMoverHechi(1).Visible = True
    .GldLbl.Visible = False
End With
Else
FormInv.Visible = False
FormInv.Hide
FormBarInv.Shape(a).BorderStyle = 0
FormBarInv.Shape(a).BorderColor = vbWhite
FormBarInv.Shape(a - 1).BorderStyle = 0
FormBarInv.Shape(a - 1).BorderColor = vbWhite
End If
    bInvMod = True
Case 3
If FormListOpciones.Visible = False Then
FormListOpciones.Visible = True
FormListOpciones.Show , frmMain
'FormListOpciones.SetFocus
FormBarInv.Shape(a).BorderStyle = 1
FormBarInv.Shape(a).BorderColor = vbYellow
Else
FormListOpciones.Visible = False
FormListOpciones.Hide
FormBarInv.Shape(a).BorderStyle = 0
FormBarInv.Shape(a).BorderColor = vbWhite
End If
Case 4
If FormInfo.Visible = False Then
FormInfo.Visible = True
FormInfo.Show , frmMain
FormBarInv.Shape(a).BorderStyle = 1
FormBarInv.Shape(a).BorderColor = vbYellow
Else
FormInfo.Visible = False
FormInfo.Hide
FormBarInv.Shape(a).BorderStyle = 0
FormBarInv.Shape(a).BorderColor = vbWhite
End If

End Select

Exit For
End If
Next a
'frmMain.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Y >= 300 Then GoTo nooh
If Y >= 900 Then GoTo nooh
For C = 0 To 4
If Shape(C).Left + Shape(C).Height >= X Then
Shape(C).BorderStyle = 1
Exit For
End If
Next C

nooh:
For X = 0 To 4
'If Not C = Empty Then
If Not X = C Then

Select Case X
    Case 0
            If FormConsola.Visible = True Then
            Shape(X).BorderStyle = 1
            Shape(X).BorderColor = vbYellow
            Else
            Shape(X).BorderStyle = 0
            Shape(X).BorderColor = vbWhite
            End If

    Case 1
    
            If FormInv.Visible = True And FormInv.cmdInfo.Visible = False Then
            Shape(X).BorderStyle = 1
            Shape(X).BorderColor = vbYellow
            Else
            Shape(X).BorderStyle = 0
            Shape(X).BorderColor = vbWhite
            End If
    Case 2
            If FormInv.Visible = True And FormInv.cmdInfo.Visible = True Then
            Shape(X).BorderStyle = 1
            Shape(X).BorderColor = vbYellow
            Else
            Shape(X).BorderStyle = 0
            Shape(X).BorderColor = vbWhite
            End If
    
    Case 3
    
            If FormListOpciones.Visible = True Then
            Shape(X).BorderStyle = 1
            Shape(X).BorderColor = vbYellow
            Else
            Shape(X).BorderStyle = 0
            Shape(X).BorderColor = vbWhite
            End If
    Case 4
            If FormInfo.Visible = True Then
            Shape(X).BorderStyle = 1
            Shape(X).BorderColor = vbYellow
            Else
            Shape(X).BorderStyle = 0
            Shape(X).BorderColor = vbWhite
            End If
    
    End Select
'End If
End If
Next X

frmMain.SetFocus
If Button Then moverForm
End Sub
Public Sub moverForm()
    Dim res As Long
    '
    ReleaseCapture
   res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
End Sub

Private Sub Timer1_Timer()

End Sub
