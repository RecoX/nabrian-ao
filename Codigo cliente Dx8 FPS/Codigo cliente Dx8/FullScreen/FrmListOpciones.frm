VERSION 5.00
Begin VB.Form FormListOpciones 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   ClientHeight    =   2250
   ClientLeft      =   3555
   ClientTop       =   3570
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmListOpciones.frx":0000
   ScaleHeight     =   2250
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   Begin VB.Label Party 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1695
      Width           =   735
   End
   Begin VB.Label Quest 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   1340
      Width           =   855
   End
   Begin VB.Label Clanes 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   1000
      Width           =   855
   End
   Begin VB.Label Estadisticas 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   675
      Width           =   1575
   End
   Begin VB.Label Opciones 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FormListOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Clanes_Click()
            If Not frmGuildLeader.Visible Then _
                Call SendData("GLINFO")
End Sub

Private Sub Estadisticas_Click()
          LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            SendData "ATRI"
            SendData "ESKI"
            SendData "FEST"
            SendData "FAMA"
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            frmEstadisticas.SetFocus
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False

End Sub
Public Sub moverForm()
    Dim res As Long
    '
    ReleaseCapture
    res = SendMessage(Me.hWnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
End Sub

Private Sub Form_Click()
frmMain.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button Then
moverForm
End If

End Sub
Private Sub Opciones_Click()
frmOpciones.Visible = True
frmOpciones.Show , frmMain
frmOpciones.SetFocus
End Sub

Private Sub Party_Click()
frmParty.ListaIntegrantes.Clear
LlegoParty = False
Call SendData("PARINF")
Do While Not LlegoParty
    DoEvents
Loop
frmParty.Visible = True
frmParty.SetFocus
LlegoParty = False
End Sub


