VERSION 5.00
Begin VB.Form frmPanelSoporte 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Soporte Actual:"
   ClientHeight    =   5370
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Buscar > En caso de bug"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H0080FFFF&
      Caption         =   "Actualizar Lista!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdResp 
      BackColor       =   &H000000FF&
      Caption         =   "Responder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar Soporte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtRespuesta 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2280
      MaxLength       =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3000
      Width           =   4215
   End
   Begin VB.TextBox txtSoporte 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   2280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin VB.ListBox lstSoportes 
      BackColor       =   &H00000000&
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
      Height          =   4140
      ItemData        =   "frmPanelSoporte.frx":0000
      Left            =   120
      List            =   "frmPanelSoporte.frx":0007
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Respondido:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   0
      Width           =   975
   End
   Begin VB.Shape shpResp 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Soporte a responder:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Soporte recibido:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de Usuarios:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2055
   End
   Begin VB.Menu MnuP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuSum 
         Caption         =   "Traer"
      End
      Begin VB.Menu mnuIr 
         Caption         =   "Ir"
      End
      Begin VB.Menu mnuCarcel 
         Caption         =   "Carcel 40 Hierro"
         Index           =   0
      End
      Begin VB.Menu mnuCarcel 
         Caption         =   "Carcel 30 Hierro"
         Index           =   1
      End
      Begin VB.Menu mnuCarcel 
         Caption         =   "Carcel 20 Hierro"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmPanelSoporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
Me.Hide
End Sub

Private Sub cmdEliminar_Click()
'Dim UserNick As String
'UserNick = InputBox("Ingrese en este cuadro el nick a borrar.", "!!!")
Call SendData("/BORSO " & UCase$(ReadFieldOptimizado$(2, Me.Caption, Asc(":"))))
End Sub

Private Sub cmdFind_Click()
MsgBox "Este botón se utiliza en caso de que alguien no pueda enviar SOS y el SOS no pueda ser respondido.", vbOKOnly
Dim UserNick As String
UserNick = InputBox("Ingrese en este cuadro el nick a buscar.", "!!!")
Call SendData("/SOSDE " & UCase$(UserNick))
Me.Caption = "Soporte Actual:" & UCase$(UserNick)
End Sub

Private Sub cmdResp_Click()
shpResp.BackColor = vbGreen
Call SendData("/RESOS " & Right$(frmPanelSoporte.Caption, Len(frmPanelSoporte.Caption) - 15) & ";" & txtRespuesta)
End Sub

Private Sub cmdUpdate_Click()
frmPanelSoporte.Hide
Call SendData("/DAMESOS")
End Sub


Private Sub lstSoportes_DblClick()
If lstSoportes.ListIndex = -1 Then Exit Sub
Call SendData("/SOSDE " & lstSoportes.List(lstSoportes.ListIndex))
Me.Caption = "Soporte Actual:" & lstSoportes.List(lstSoportes.ListIndex)
Me.txtRespuesta = ""
End Sub

Private Sub lstSoportes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu MnuP
End If
End Sub

Private Sub mnuCarcel_Click(Index As Integer)
If lstSoportes.ListIndex = -1 Then Exit Sub
Select Case Index
Case 0
If lstSoportes.ListIndex = -1 Then Exit Sub
Call SendData("/CARCEL SOPORTE INVALIDO" & lstSoportes.List(lstSoportes.ListIndex) & "@40")
Case 1
If lstSoportes.ListIndex = -1 Then Exit Sub
Call SendData("/CARCEL SOPORTE INVALIDO" & lstSoportes.List(lstSoportes.ListIndex) & "@30")
Case 2
If lstSoportes.ListIndex = -1 Then Exit Sub
Call SendData("/CARCEL SOPORTE INVALIDO@" & lstSoportes.List(lstSoportes.ListIndex) & "@20")
End Select
End Sub

Private Sub mnuIr_Click()
If lstSoportes.ListIndex = -1 Then Exit Sub
Call SendData("/IRA " & lstSoportes.List(lstSoportes.ListIndex))
End Sub

Private Sub mnuSum_Click()
If lstSoportes.ListIndex = -1 Then Exit Sub
Call SendData("/SUM " & lstSoportes.List(lstSoportes.ListIndex))
End Sub

Private Sub txtSoporte_Click()
txtRespuesta.SetFocus
End Sub

