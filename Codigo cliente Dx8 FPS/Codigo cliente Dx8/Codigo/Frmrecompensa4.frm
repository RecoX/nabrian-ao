VERSION 5.00
Begin VB.Form frmSubeClase4 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "Topico = 11"
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Cerrar"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Más información"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   4550
      MouseIcon       =   "Frmrecompensa4.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Más información"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   1320
      MouseIcon       =   "Frmrecompensa4.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Más información"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   4550
      MouseIcon       =   "Frmrecompensa4.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Más información"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   1320
      MouseIcon       =   "Frmrecompensa4.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Image command1 
      Height          =   375
      Index           =   3
      Left            =   4800
      MouseIcon       =   "Frmrecompensa4.frx":0C28
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   2
      Left            =   1560
      MouseIcon       =   "Frmrecompensa4.frx":0F32
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   1
      Left            =   4680
      MouseIcon       =   "Frmrecompensa4.frx":123C
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   0
      Left            =   1560
      MouseIcon       =   "Frmrecompensa4.frx":1546
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label 7"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3760
      TabIndex        =   8
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4040
      TabIndex        =   7
      Top             =   4500
      Width           =   2415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   540
      TabIndex        =   6
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   860
      TabIndex        =   5
      Top             =   4500
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   3760
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   540
      TabIndex        =   3
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Frmrecompensa4.frx":1850
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   5415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   1800
      Width           =   2295
   End
End
Attribute VB_Name = "frmSubeClase4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
'director del proyecto: #Esteban(Neliam)

'servidor basado en fénixao 1.0
'medios de contacto:
'Skype: dc.esteban
'E-mail: nabrianao@gmail.com
Private Sub command1_Click(Index As Integer)

Call SendData("RSB" & Index + 1)
Unload Me

End Sub
Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "clase4.jpg")

Select Case (MiClase)
    Case TRABAJADOR
        Label1.Caption = "Experto en minerales"
        Label2.Caption = "Experto en madera"
        Label6.Caption = "Pescador"
        Label9.Caption = "Sastre"
        
        Label4.Caption = "El experto en minerales dedicará completamente su vida a todo lo relacionado con metales tales como el oro, la plata o el hierro. Podrá extraerlos o bien dedicarse a trabajarlos."
        Label5.Caption = "El experto en uso de madera ama el trabajo del carpintero o leñador. Generalmente los conocimientos se transmiten de padre a hijo, de generación en generación. Podrá talar o bien dedicarse a trabajarla."
        Label7.Caption = "Campesinos sin profesión deciden dedicarse a la relajada vida de los pescadores. Pacientes, humildes, y usualmente muy charlatanes y generosos. Muchos de los grandes guerreros de este mundo tienen amigos pescadores que en algún momento los han ayudado."
        Label10.Caption = "El sastre confecciona todo aquel ropaje que este hecho en base a pieles y tela. Sus productos proveen una gran defensa de estar transformados mágicamente, y pueden ser usados por la gran mayoría de los habitantes de estas tierras."
        
        Label3.Caption = "Es hora de elegir el tipo de trabajador que deseas ser. Recuerda que al elegir uno de estos caminos, estarás desechando para siempre cualquier otro."
        
    Case CON_MANA
        Label1.Caption = "Hechicero"
        Label2.Caption = "Orden Sagrada"
        Label6.Caption = "Naturalista"
        Label9.Caption = "Sigiloso"
        
        Label4.Caption = "Los hechiceros usan casi exclusivamente la magia para realizar sus ataques. En niveles avanzados pueden causar severos daños o invocar poderosísimas criaturas, sin necesidad de grandes equipamientos."
        Label5.Caption = "Los que pertenecen a esta orden actúan siempre en nombre de un dios, ya se benigno o maligno. Los que creen en la bondad tratan por lo general de utilizar la palabra, mientras que quienes fueron corrompidos, no dudarán en matar a quien se les cruce. "
        Label7.Caption = "Amantes de la naturaleza, el arte, los árboles y los ríos. Encuentran el placer en cosas simples del mundo: un poema, el olor de una rosa, una dulce melodía. Han tomado de la naturaleza la habilidad única de usarla en defensa propia."
        Label10.Caption = "Escondidos, agachados, ocultos van tras su presa y la eliminan en el mayor de los silencios. Juegan a favor de la Alianza matando figuras claves, caudillos, revolucionarios o bien en contra de la misma, tratando de cometer regicidio. "
        
        Label3.Caption = "Un nuevo paso has avanzado, y ahora debes elegir qué nivel de magia quieres alcanzar."
    
End Select

End Sub

Private Sub Form_LostFocus()

Me.Visible = False

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



Private Sub Label11_Click()
Unload Me
End Sub

Private Sub Label8_Click(Index As Integer)
Ayuda = 0
FrmAyuda.Show , frmSubeClase4

End Sub


