VERSION 5.00
Begin VB.Form Frmrecompensa3 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Más información"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2880
      MouseIcon       =   "Frmrecompensa3.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Más información"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4560
      MouseIcon       =   "Frmrecompensa3.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Más información"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      MouseIcon       =   "Frmrecompensa3.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      MouseIcon       =   "Frmrecompensa3.frx":091E
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   735
   End
   Begin VB.Image command2 
      Height          =   255
      Left            =   4800
      MouseIcon       =   "Frmrecompensa3.frx":0C28
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Image command1 
      Height          =   255
      Left            =   1560
      MouseIcon       =   "Frmrecompensa3.frx":0F32
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Image command3 
      Height          =   375
      Left            =   3120
      MouseIcon       =   "Frmrecompensa3.frx":123C
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   930
      TabIndex        =   6
      Top             =   4610
      Width           =   5415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H8000000E&
      Height          =   1455
      Left            =   900
      TabIndex        =   5
      Top             =   4920
      Width           =   5445
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3775
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   615
      TabIndex        =   3
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Frmrecompensa3.frx":1546
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1080
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
      Left            =   4180
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
Attribute VB_Name = "Frmrecompensa3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()

Select Case (MiClase)
Case Is = 4
SendData "RSB5"
AddtoRichTextBox frmMain.rectxt, "¡Ahora eres un Minero Aeluin! Restás 50 de vida.", 255, 250, 55, 1, 0  'Info
Case Is = 14
SendData "RSB15"
AddtoRichTextBox frmMain.rectxt, "¡Ahora eres un Leñador Eldalie! Restás 50 de vida.", 255, 250, 55, 1, 0  'Info
Case Is = 23
SendData "RSB24"
AddtoRichTextBox frmMain.rectxt, "¡Ahora eres un Pescador Brethil!", 255, 250, 55, 1, 0  'Info
Case Is = 27
SendData "RSB28"
AddtoRichTextBox frmMain.rectxt, "¡Ahora eres un Sastre Troron!", 255, 250, 55, 1, 0  'Info
Case Is = 31
SendData "RSB32"
AddtoRichTextBox frmMain.rectxt, "¡Ahora eres un Alquimista Oilosse!", 255, 250, 55, 1, 0  'Info
Case Is = 52
SendData "RSB53"
AddtoRichTextBox frmMain.rectxt, "¡Finalmente, eres un Tesorero!", 255, 250, 55, 1, 0  'Info
End Select
Me.Refresh
Unload Frmrecompensa3
End Sub

Private Sub Command2_Click()
Select Case (MiClase)
Case Is = 4
SendData "RSB6"
AddtoRichTextBox frmMain.rectxt, "¡Ahora eres un Minero Ciryatan!", 255, 250, 55, 1, 0  'Info
Case Is = 14
SendData "RSB16"
AddtoRichTextBox frmMain.rectxt, "¡Ahora eres un Leñador Andor!", 255, 250, 55, 1, 0  'Info
Case Is = 23
SendData "RSB25"
AddtoRichTextBox frmMain.rectxt, "¡Ahora eres un Pescador Hathol!", 255, 250, 55, 1, 0  'Info
Case Is = 27
SendData "RSB29"
AddtoRichTextBox frmMain.rectxt, "¡Ahora eres un Sastre Othar!", 255, 250, 55, 1, 0  'Info
Case Is = 31
SendData "RSB33"
AddtoRichTextBox frmMain.rectxt, "¡Ahora eres un Alquimista Míriel!", 255, 250, 55, 1, 0  'Info
Case Is = 52
SendData "RSB54"
AddtoRichTextBox frmMain.rectxt, "¡Finalmente, eres un Comandante!", 255, 250, 55, 1, 0  'Info
End Select
Me.Refresh
Unload Frmrecompensa3
End Sub

Private Sub Command3_Click()
Select Case (MiClase)
Case Is = 4
SendData "RSB7"
AddtoRichTextBox frmMain.rectxt, "¡Ahora eres un Minero Loth Maeg! Ganas 100 de vida.", 255, 250, 55, 1, 0  'Info
Case Is = 14
SendData "RSB17"
AddtoRichTextBox frmMain.rectxt, "¡Ahora eres un Leñador Deron!", 255, 250, 55, 1, 0  'Info
Case Is = 23
SendData "RSB26"
AddtoRichTextBox frmMain.rectxt, "¡Ahora eres un Pescador Wethrin!", 255, 250, 55, 1, 0  'Info
Case Is = 27
SendData "RSB30"
AddtoRichTextBox frmMain.rectxt, "¡Ahora eres un Sastre Nauglamir!", 255, 250, 55, 1, 0  'Info
Case Is = 31
SendData "RSB34"
AddtoRichTextBox frmMain.rectxt, "¡Ahora eres un Alquimista Uldor!", 255, 250, 55, 1, 0  'Info
Case Is = 52
SendData "RSB55"
AddtoRichTextBox frmMain.rectxt, "¡Finalmente, eres un Pirata!", 255, 250, 55, 1, 0  'Info
End Select
Me.Refresh
Unload Frmrecompensa3
End Sub

Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "Suclases3op.gif")

Select Case (MiClase)

Case Is = 4
Label1.Caption = "Minero Aeluin"
Label2.Caption = "Minero Ciryatan"
Label6.Caption = "Minero Loth Maeg"

Label4.Caption = "Los mineros Aeluin, escaparon atemorizados de la gran guerra que se desato en las cercanías de Banderbille mucho tiempo atrás. Sus descendientes heredaron dolorosos recuerdos de humillación que hacen su vida mucho más dura y escasa pero a su vez heredaron un increíble talento."
Label5.Caption = "Estos mineros pasaron sin pena ni gloria una de las etapas más duras que el imperio sufrió. Estuvieron ocultos en cuevas por años hasta que ya no tuvieron más opción que pelear, en la ultima alianza."
Label7.Caption = "La facción menos numerosa pero por lejos más valiente entre los mineros. Desde el principio lucharon sin temor por una causa la cual creyeron justa. Muchos murieron, y los herederos perdieron el talento nato, aunque un gran orgullo y una vida mucho mayor."

Label3.Caption = "Es hora de elegir el tipo de minero que deseas ser. Cada uno de ellos tiene distintas bonificaciones, por lo que piensa bien la utilidad que piensas darle."

Topico = 12

Case Is = 14
Label1.Caption = "Leñador Eldalie"
Label2.Caption = "Leñador Andor"
Label6.Caption = "Leñador Deron"

Label4.Caption = "Pequeños pero fuertes y agilidosos, los leñadores de la raza Eldalie trabajan con la misma dureza cada día para conseguir el pan. Algunos pocos inteligentes llegan a cosechar pequeñas fortunas, aunque la mayoría viven una vida de pobreza y tienen menor esperanza de vida."
Label5.Caption = "Los Andor son la clase de leñadores más tradicional. No existe Andor cuyo padre y abuelo no haya sido leñador anteriormente. De carácter amistoso y solidario, nunca se esfuerzan demasiado ya que su mutua ayuda les sirve para sobrevivir."
Label7.Caption = "Fuertes, robustos, entrenan muchísimo su físico antes de dedicarse a su tarea. Contrariamente a los Andor, son cerrados y fríos, muy poco sociables; lo cual ha traído como peor consecuencia su escasa técnica. Son de perfil bajo y pocos conocidos en estas tierras."

Label3.Caption = "Es hora de elegir el tipo de leñador que deseas ser. Cada uno de ellos tiene distintas bonificaciones, por lo que piensa bien la utilidad que piensas darle."

Topico = 13

Case Is = 23
Label1.Caption = "Pescador Brethil"
Label2.Caption = "Pescador Hathol"
Label6.Caption = "Pescador Wethrin"

Label4.Caption = "Conforman el sector más pobre de Las Tierras del Fénix. Pocos llegan a la vejez, ya que muchos mueren abandonados. Aprenden el oficio por imitación pero sorprendentemente son los que más peces pueden extraer en una jornada."
Label5.Caption = "Por repetición aprenden el arte pero no profundizan sus conocimientos. Utilizan la paciencia como técnica para extraer más de un pez a la vez. Son los más callados de todos los pescadores y no llegan a destacarse en las Tierras de Fénix Ao a pesar de que algunos logran pequeñas fortunas."
Label7.Caption = "Su falta de habilidad se compensa con su extraordinaria suerte. En la mayoría de sus intentos logran sacar deliciosos peses aunque pocas veces de a muchos."

Label3.Caption = "Es hora de elegir el tipo de pescador que deseas ser. Cada uno de ellos tiene distintas bonificaciones, por lo que piensa bien la utilidad que piensas darle."

Topico = 14

Case Is = 27
Label1.Caption = "Sastre Troron"
Label2.Caption = "Sastre Othar"
Label6.Caption = "Sastre Nauglamir"

Label4.Caption = "Los sastres Thoron son por lejos los más cultos de las Tierras del Fénix. Sus conocimientos son envidiados hasta por algunos clérigos y sus habilidades en sastrería son absolutamente asombrosas. En niveles avanzados pueden llegar a confeccionar hermosas vestimentas dotadas de importantes protecciones para las clases débiles."
Label5.Caption = "El origen de los Othar fue y continua siendo una verdadera leyenda. La historia se ha pasado de generación en generación hasta el día de hoy y dice que un noble, hace cientos de años, ordeno a un sastre hacer una vestimenta a base de menos pieles sin perder calidad. El sastre lo logró y se salvó de la horca, su técnica perduró con el tiempo."
Label7.Caption = "Los Nauglamir tienen una increíble facilidad para aprender las técnicas de sastrería, pero lamentablemente no han sabido aprovecharla. Aunque no esté comprobado, se dice que es hereditario; tanto esta facilidad como su extrema vagancia."

Label3.Caption = "Es hora de elegir el tipo de sastre que deseas ser. Cada uno de ellos tiene distintas bonificaciones, por lo que piensa bien la utilidad que piensas darle."
Topico = 15

Case Is = 31
Label1.Caption = "Alquimista Oilosse"
Label2.Caption = "Alquimista Míriel"
Label6.Caption = "Alquimista Uldor"

Label4.Caption = "Son excelentes botánicos: conocen cada hierba, planta, flor y árbol de las Tierras del Fénix y saben cada técnica de cultivo existente. No hay duda que nadie sabe más sobre plantaciones que ellos."
Label5.Caption = "Al contrario de los Oilosse, los Míriel han dejado un poco de lado la botánica y se han especializado en la alquimia propiamente dicha. No tienen tal capacidad para extraer raíces pero si una gran facilidad para ejercer la alquimia."
Label7.Caption = "Un balance entre los Muriel y los Oilosse; tienen una interesante capacidad para la alquimia sin dejar de lado el estudio de la botánica. Pueden hacer todo tipo de pociones y cultivar y extraer raíces a un velocidad media."

Label3.Caption = "Es hora de elegir el tipo de alquimista que deseas ser. Cada uno de ellos tiene distintas bonificaciones, por lo que piensa bien la utilidad que piensas darle."
Topico = 16

Case Is = 52
Label1.Caption = "Tesorero"
Label2.Caption = "Comandante"
Label6.Caption = "Pirata"

Label4.Caption = "Sacan mayor provecho del botín y lo guardan consigo en lugares imposibles de descubrir, pues saben que si son descubiertos tendrán una muerte segura."
Label5.Caption = "Los comandantes son dueños de sus propios barcos. Ellos mismos lo consiguen ahorrando durante décadas y ellos mismos reúnen una tripulación. Si el barco se hunde, ellos mueren con él."
Label7.Caption = "De consistencia fuerte, son llamados los guerreros del mar. Tienen características realmente similares a dicha clase, aunque en el agua son casi invencibles. Saben moverse en un barco como en su propia casa."

Label3.Caption = "Siendo navegante, puedes adquirir fuerza similar a la de un guerrero. Dedicarte a guardar tesoros. O simplemente, explorar el mundo en tu propio barco."
Topico = 17

End Select

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
Private Sub Image1_Click()
Unload Me

End Sub

Private Sub Label10_Click()
FrmAyuda.Show
End Sub

Private Sub Label8_Click()
FrmAyuda.Show
End Sub

Private Sub Label9_Click()
FrmAyuda.Show
End Sub
