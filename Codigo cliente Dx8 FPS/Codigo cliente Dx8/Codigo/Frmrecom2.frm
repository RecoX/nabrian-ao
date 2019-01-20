VERSION 5.00
Begin VB.Form Frmrecom2 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      MouseIcon       =   "Frmrecom2.frx":0000
      MousePointer    =   99  'Custom
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image command2 
      Height          =   375
      Left            =   4800
      MouseIcon       =   "Frmrecom2.frx":030A
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Image Command1 
      Height          =   375
      Left            =   1560
      MouseIcon       =   "Frmrecom2.frx":0614
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   3710
      TabIndex        =   4
      Top             =   2150
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   490
      TabIndex        =   3
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Frmrecom2.frx":091E
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   6225
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4070
      TabIndex        =   1
      Top             =   1755
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   860
      TabIndex        =   0
      Top             =   1755
      Width           =   2535
   End
End
Attribute VB_Name = "Frmrecom2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
Select Case (MiClase)
Case Is = 38
SendData "REL12" '5 puntos de vida
Case Is = 39
SendData "REL12" '5 puntos de vida
Case Is = 41
SendData "REL12" '5 puntos de vida
Case Is = 42
SendData "REL12" '5 puntos de vida
Case Is = 44
SendData "REL12" '5 puntos de vida
Case Is = 45
SendData "REL12" '5 puntos de vida
Case Is = 47
SendData "REL12" '5 puntos de vida
Case Is = 48
SendData "REL12" '5 puntos de vida
Case Is = 50
SendData "REL12" '5 puntos de vida
Case Is = 51
SendData "REL12" '5 puntos de vida
End Select

Me.Refresh
Unload Me
Me.Hide

End Sub

Private Sub Command2_Click()
Select Case (MiClase)
Case Is = 38
SendData "REL16" '+10 de defensa en la armadura/tunica faccionaria
Case Is = 39
SendData "REL16" '+10 de defensa en la armadura/tunica faccionaria
Case Is = 41
SendData "REL17" '+15 de defensa en la armadura/tunica faccionaria
Case Is = 42
SendData "REL16" '+10 de defensa en la armadura/tunica faccionaria
Case Is = 44
SendData "REL16" '+10 de defensa en la armadura/tunica faccionaria
Case Is = 45
SendData "REL17" '+15 de defensa en la armadura/tunica faccionaria
Case Is = 47
SendData "REL16" '+10 de defensa en la armadura/tunica faccionaria
Case Is = 48
SendData "REL17" '+15 de defensa en la armadura/tunica faccionaria
Case Is = 50
SendData "REL17"  '+15 de defensa en la armadura/tunica faccionaria
Case Is = 51
SendData "REL17" '+15 de defensa en la armadura/tunica faccionaria
End Select
Me.Refresh
Unload Frmrecom2
End Sub

Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "Suclases2op.gif")

Select Case (MiClase)
Case Is = 38
Label1.Caption = "Vida"
Label2.Caption = "Defensa"
Label4.Caption = "Se te otorgan 5 puntos de vida extras de forma permanente."
Label5.Caption = "+10 de defensa a la armadura o túnica faccionaria (no surtirá efecto si no tienes una)."

Case Is = 39
Label1.Caption = "Vida"
Label2.Caption = "Defensa"
Label4.Caption = "Se te otorgan 5 puntos de vida extras de forma permanente."
Label5.Caption = "+10 de defensa a la armadura o túnica faccionaria (no surtirá efecto si no tienes una)."

Case Is = 41
Label1.Caption = "Vida"
Label2.Caption = "Defensa"
Label4.Caption = "Se te otorgan 5 puntos de vida extras de forma permanente."
Label5.Caption = "+15 de defensa a la armadura o túnica faccionaria (no surtirá efecto si no tienes una)."

Case Is = 42
Label1.Caption = "Vida"
Label2.Caption = "Defensa"
Label4.Caption = "Se te otorgan 5 puntos de vida extras de forma permanente."
Label5.Caption = "+10 de defensa a la armadura o túnica faccionaria (no surtirá efecto si no tienes una)."



'nuevas

Case Is = 44
Label1.Caption = "Vida"
Label2.Caption = "Defensa"
Label4.Caption = "Se te otorgan 5 puntos de vida extras de forma permanente."
Label5.Caption = "+10 de defensa a la armadura o túnica faccionaria (no surtirá efecto si no tienes una)."

Case Is = 45
Label1.Caption = "Vida"
Label2.Caption = "Defensa"
Label4.Caption = "Se te otorgan 5 puntos de vida extras de forma permanente."
Label5.Caption = "+15 de defensa a la armadura o túnica faccionaria (no surtirá efecto si no tienes una)."

Case Is = 47
Label1.Caption = "Vida"
Label2.Caption = "Defensa"
Label4.Caption = "Se te otorgan 5 puntos de vida extras de forma permanente."
Label5.Caption = "+10 de defensa a la armadura o túnica faccionaria (no surtirá efecto si no tienes una)."

Case Is = 48
Label1.Caption = "Vida"
Label2.Caption = "Defensa"
Label4.Caption = "Se te otorgan 5 puntos de vida extras de forma permanente."
Label5.Caption = "+15 de defensa a la armadura o túnica faccionaria (no surtirá efecto si no tienes una)."

Case Is = 50
Label1.Caption = "Vida"
Label2.Caption = "Defensa"
Label4.Caption = "Se te otorgan 5 puntos de vida extras de forma permanente."
Label5.Caption = "+15 de defensa a la armadura o túnica faccionaria (no surtirá efecto si no tienes una)."

Case Is = 51
Label1.Caption = "Vida"
Label2.Caption = "Defensa"
Label4.Caption = "Se te otorgan 5 puntos de vida extras de forma permanente."
Label5.Caption = "+15 de defensa a la armadura o túnica faccionaria (no surtirá efecto si no tienes una)."

End Select
End Sub

Private Sub Image1_Click()
Unload Me
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

Private Sub Label5_Click()

End Sub
