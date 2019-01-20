VERSION 5.00
Begin VB.Form frmSubeClase2 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Cerrar"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Más información"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   4520
      MouseIcon       =   "Frmrecompensa2.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Más información"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   1290
      MouseIcon       =   "Frmrecompensa2.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Image Command1 
      Height          =   375
      Index           =   1
      Left            =   4800
      MouseIcon       =   "Frmrecompensa2.frx":0614
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Image Command1 
      Height          =   375
      Index           =   0
      Left            =   1440
      MouseIcon       =   "Frmrecompensa2.frx":091E
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   3720
      TabIndex        =   4
      Top             =   2145
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   495
      TabIndex        =   3
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Frmrecompensa2.frx":0C28
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
Attribute VB_Name = "frmSubeClase2"
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
                            
Me.PICTURE = LoadPicture(DirGraficos & "clase2.jpg")

Select Case MiClase
    Case CIUDADANO
        Label1.Caption = "Desabilitado"
        Label2.Caption = "Luchador"
        
        
        Label5.Caption = "El luchador usa sus habilidades en combate para ganar dinero. Prefiere una vida más arriesgada, llena de aventuras y emociones. Puede elegir diversas ramificaciones, tales como la magia, la espada o el arco."
        Label4.Caption = "Las clases trabajadoras fueron desabilitadas en un servidor de agite no hacen falta."
        Label3.Caption = "Ha llegado el momento de tomar la primera decisión importante de tu vida. A partir de esta elección se desarrollará todo de ahora en más, por lo que ten cuidado! Una vez tomada una decisión, ya no podrás volver marcha atrás."

    Case EXPERTO_MINERALES
        Label1.Caption = "Minero"
        Label2.Caption = "Herrero"
        
        Label4.Caption = "El minero, como bien dice su nombre, mina oro, plata y hierro en las tierras del Fénix. Podrán encontrar importantes minas en el continente aunque la más importante será alcanzable solo mediante navegación."
        Label5.Caption = "El herrero tiene una vida tan dura como la del minero. Forja poderosas armas, grandes escudos y fuertes armaduras para sobrevivir o, en ciertos casos, utilizarlas para beneficio personal. Lo hay en distintos tipos y con distintas características, ¡solo te resta elegir!"
        
        Label3.Caption = "Es momento de elegir qué rama de los minerales seguirás. Si quieres extraerlos, deberías pensar en ser minero. Si prefieres fabricar lingotes o armas, puedes elegir ser un herrero."
    
    Case EXPERTO_MADERA
        Label1.Caption = "Leñador"
        Label2.Caption = "Carpintero"
        
        Label4.Caption = "De procedencia humilde, trabajan para carpinteros o en ciertas ocasiones para poderosos terratenientes. Algunos llegan en su vida a talar suficiente madera para varias barcas."
        Label5.Caption = "Manejan el serrucho a la perfección y modelan la madera a gusto. Grandes diseñadores de barcas y pequeños productores de flechas. Constructores de hermosos y complejos arcos y simples constructores de amoblamiento para hogares."
        
        Label3.Caption = "Ahora debes tomar una decisión importante en tu vida. Si quieres dedicarte a la tala de árboles, elige ser Leñador. Si por el contrario quieres construir cosas a partir de madera, sé un buen Carpintero."
        
    Case LUCHADOR
        Label1.Caption = "Con uso de Mana"
        Label2.Caption = "Sin uso de Mana"
        
        Label4.Caption = "Esta tipo de luchadores utilizan en mayor o menor medida la magia, pudiendo combinarla con la espada o el arco. Pueden desatar poderosos conjuros y causar diversos efectos sobre su oponente y sobre si mismos."
        Label5.Caption = "Poco o nada les interesan las artes mágicas a quienes eligen este camino. Si sigues esta senda te basarás mucho más en tu poderío físico que en memorizar largos conjuros y complicados hechizos. Si tu fuerte no es la inteligencia, este es tu camino."
        
        Label3.Caption = "Ahora debes tomar una decisión importante en tu vida. Debes elegir entre aprender habilidades mágicas en mayor o menos medida o dedicarte a la fuerza bruta únicamente, dejando completamente de lado el uso de magia."

    Case HECHICERO
        Label1.Caption = "Mago"
        Label2.Caption = "Nigromante"
        
        Label4.Caption = "El mago puede usar el mejor hechizo de ataque en los niveles más avanzados. Su poder puede llegar a ser absolutamente devastador si aprende a combinarlos con eficacia y sabiduría."
        Label5.Caption = "El nigromante puede llegar a invocar una temible criatura tal como lo es el fuego fatuo. El fuego fatuo puede eliminar fácilmente rivales de poca envergadura y, siendo combinado con fuertes hechizos de ataque directo, puede eliminar a los más poderosos guerreros."
        
        Label3.Caption = "Eres alguien totalmente dedicado a la magia. Es momento de decidir si quieres ser poderoso por el daño de tus hechizos, o por la fuerza de los que invocas."

    Case ORDEN
        Label1.Caption = "Paladín"
        Label2.Caption = "Clérigo"
        
        Label4.Caption = "Prefieren predicar la palabra de Dios mediante la espada. Aman a sus dioses y dedican su entera vida a ellos. Hay paladines realmente adinerados y otros mucho más humildes. Por lo general, llevan su rol a extremos, pudiendo ser muy benévolos o realmente malvados."
        Label5.Caption = "Pasa gran parte de su vida dentro del templo, orando por las almas de las personas vivas y muertas del mundo. Así como los paladines, pueden ser buenos o malos dependiendo de la deidad a la que sigan. Son considerados las personas más cultas de las Tierras del Fénix."

    Case NATURALISTA
        Label1.Caption = "Bardo"
        Label2.Caption = "Druida"
        
        Label4.Caption = "El bardo es un verdadero experto en las artes musicales. Conoce cada nota y el efecto que estas producen al ser combinadas en hermosas melodías. Asombrosos y sorprendentes los bardos son."
        Label5.Caption = "Nacen y se crían en medio de la naturaleza. Tienen un nato rechazo a la ciudad y la civilización. Siempre que un druida tenga que entrar en combate, contará con el entero apoyo y ayuda de la naturaleza."

    Case SIGILOSO
        Label1.Caption = "Asesino"
        Label2.Caption = "Cazador"
        
        Label4.Caption = "Una fuerte apuñalada es suficiente para que su enemigo caiga derrotado sin siquiera saber quien fue. De poco físico y a su vez de poca piedad los asesinos son. No pueden llevar grandes armaduras ni importantes escudos, pero un certero golpe es más que suficiente."
        Label5.Caption = "Tras las sombras, con arco en mano y flecha preparada. La cuerda tiesa, el proyectil apuntando a la cabeza; sabe que suelta la cuerda es una muerte segura. Lo hace y no falla, su recompensa será grande y el lo sabe más que bien."
        
        Label3.Caption = "Ahora debes tomar una decisión importante en tu vida. Si quieres dedicarte a la tala de árboles, elige ser Leñador. Si por el contrario quieres construir cosas a partir de madera, sé un buen Carpintero."

    Case SIN_MANA
        Label1.Caption = "Desabilitado"
        Label2.Caption = "Caballero"
        
        Label4.Caption = "La clase pirata y ladron han sido desabilitadas en un servidor de agite no hacen falta."
        Label5.Caption = "Los caballeros deciden dedicar su vida al bien y luchan valientemente en las líneas del frente contra el enemigo. Tienen un control total de las armas ya sea al pelear cuerpo a cuerpo o al derrotar al enemigo a distancia con una poderosa flecha."
    
    Case BANDIDO
        Label1.Caption = "Pirata"
        
        
        Label4.Caption = "De consistencia fuerte, son llamados los guerreros del mar. Tienen características realmente similares a dicha clase, aunque en el agua son casi invencibles. Saben moverse en un barco como en su propia casa."
        
        Label3.Caption = "Puedes dedicarte al hurto o preferir navegar los mares de las Tierras de Fénix como un pirata."
        
    Case CABALLERO
        Label1.Caption = "Guerrero"
        Label2.Caption = "Arquero"
        
        Label4.Caption = "Dan golpes muy fuertes con sus espadas o puños. Tienen un impactante aspecto físico que los hace temidos por muchos, a pesar de que la mayoría sea de buen corazón. Suelen portar impresionantes espadas y grandes armaduras."
        Label5.Caption = "Se especializa en combate con arcos, aunque puede usar algunas pocas armas. Los arqueros son seres de una gran agilidad, velocidad y puntería. Los que llegan a niveles avanzados pueden partir una nuez a varios metros de distancia con los ojos vendados."
        
        Label3.Caption = "Puedes dedicarte al uso de la espada, o preferir manejar con precisión el arco. También podrías ser un gran navegante, o un delincuente."

End Select

End Sub

Private Sub Form_LostFocus()

Me.Visible = False

End Sub

Private Sub Image1_Click()

Unload Me

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


Private Sub Label6_Click(Index As Integer)

Ayuda = 0
FrmAyuda.Show , frmSubeClase2

End Sub
