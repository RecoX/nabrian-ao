VERSION 5.00
Begin VB.Form frmguiajuego 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8835
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmguiajuego.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Duda 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      ItemData        =   "frmguiajuego.frx":000C
      Left            =   120
      List            =   "frmguiajuego.frx":0031
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   8520
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000012&
      X1              =   1200
      X2              =   9000
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line6 
      X1              =   7800
      X2              =   8880
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   8280
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000012&
      X1              =   0
      X2              =   7800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecciona tu duda:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000B&
      X1              =   0
      X2              =   5040
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000012&
      X1              =   0
      X2              =   7800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label InfoLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecciona tu duda para filtrar Información."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7575
   End
   Begin VB.Line Line8 
      Index           =   1
      X1              =   0
      X2              =   9480
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9480
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmguiajuego"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Duda_click()

Select Case (Duda.List(Duda.ListIndex))
    Case Is = "¿Como ser Templario? y que Privilegio Obtengo."
    InfoLabel.Caption = "Para ser templario deberán Realizar las Misiones del NPC THOR, este mismo se encuentra en la ciudad de Ullathorpe y Arghelin.                                    ¿Que privilegio obtengo siendo templario?                                               -Viajes a dungeones (desde el vendedor de pasajes).                                -15% De devolución de los puntos al canjear items de Canjeos.                   -15% De devolución de los puntos al canjear items de Donador.                 - 5% más de probabilidad de Drop.                                                        - Cambio de meditación.                                                                                                                                                                            ¿De que otra forma puedo hacerme templario sin conseguir los items de las misiones de Thor?  En los item de donadores se encuentra el 'Anillo de los Dioses Templarios' el cual anulara todas las misiones y te hará templario."

    
    Case Is = "¿Como pasar de nivel?"
        InfoLabel.Caption = "Como notaran pasar de nivel es totalmente fácil, no lleva mas que 10 minutos en la sala de entrenamiento que esta al crear su usuario.                                                                                                                                                                                                                     Luego de llegar a nivel 45 tendrán que pasar 5 niveles, los cuales ya son mas complicados hasta el nivel 50 subirá las estadísticas de vida, mana, etc luego tendrán premios y recompensas."
      
     Case Is = "¿Como retar 1vs1?"
        InfoLabel.Caption = "En la ciudad de arghelin podran retar apretando la tecla 'F2', Con la posibilidad de elegir si el duelo sera de plante o 1vs1, podran apostar los canjeos que deseen, el reto sera al mejor de 3."
   
     Case Is = "¿Como retar 2vs2?"
        InfoLabel.Caption = "Simple sistema de 2vs2 puede ser con el panel 'f2' o poniendo /PAREJA NICK, para jugar ambos personajes deben estar en ISLA PIRATA."
     
     Case Is = "¿Como retar 3vs3?"
        InfoLabel.Caption = "En la ciudad de IP (ISLA PIRATA), deberan estar 3 usuarios en una party el creador del party debera poner el comando /TRIO cuando hayan 2 trios se los llevara a cabo el reto."
    
     Case Is = "¿Como apostar mi personaje?"
        InfoLabel.Caption = "En la ciudad de arghelin, con la tecla f8 podran elegir a su contricante el cual aceptara con la tecla f8 poniendo su codigo de seguridad. El personaje que pierda, se le cambiara el codigo por el personaje ganador. El reto es al mejor de 3."
   
     Case Is = "¿Como dominar castillo de clanes o cuartel faccionario?"
        InfoLabel.Caption = "Consiste en ir al mapa de castillo, deberan tener clan. En el castillo se encontrara el rey de clanes el cual miembro de un clan lo mate sera el clan que lo domine. Deberan defenderlo el tiempo que mas puedan. Al pasar 1 hora de domino el rey acumulara 2 puntos de canjeos al clan que lo domine. Al dominar otro clan o matar el rey por si mismo otorgara los canjeos acumulados a todos los usuarios de el clan que domine. De la misma manera funciona el Cuartel faccionario, para ver los puntos acumulados deberan tipear /CASTILLOS."
    
     Case Is = "¿Como conseguir monturas?"
        InfoLabel.Caption = "En los bosques del mundo se encuentran los NPCs que dropean monturas las cual pueden montar y equiparse los usuarios. Entre ellas existe la montura del tigre (tiene 15% de probabilidad de dropeo), montura de unicornio (25% de probabilidad de dropeo), montura de minotauro (15% probabilidad de dropeo) la montura de Dragon amarillo, rojo y los caballos los cuales se consiguen solamente donando o comprandosela a otro usuario."
     
     Case Is = "Clanes información"
        InfoLabel.Caption = "Para fundar clan deberan tener los siguientes requisitos: 500 puntos de canjeo (se consiguen donando o ganando torneos/quest dentro del juego), 8 puntos de torneos, 6 puntos de quest, Gema Celeste (se consigue matando un NPC), Gema Roja (se consigue matando un NPC), ser nivel 50. Al tener los requisitos podran fundar con el comando /FUNDARCLAN.                                                                                                                                                                Guerra de clanes:  El líder de un clan envía el comando /RETARCLAN 'Nick del líder o nombre del clan a retar', de esta manera se le abrirá un question donde preguntara si quiere retar a modalidad Cupos vs Cupos."
    
     Case Is = "¿Como hacer una misión?"
        InfoLabel.Caption = "En el mapa 2 (CIUDAD DE ARGHELIN), se encuentra el NPC de quest. El cual al clickearlo se les abrira un formulario en el cual pueden elegir su quest. Las quest son por puntos de canjeos, un ejemplo de quest es: MATAR 3 Rey Dragones."
 
     Case Is = "¿Como subastar un ITEM?"
        InfoLabel.Caption = "Para subastar un item tendran que poner /SUBASTAR y elegir el item poner el item que quieran subastar.. y poner el precio inicial.. los usuarios iran ofreciendo con el comando /OFRECER X puntos, la subasta tardara 4 minutos el usuario que ofresca mas canjes será el que se lleve el item subastado y el subastador se lleve los canjes."
  
End Select

End Sub



Private Sub Form_Load()
frmguiajuego.Caption = "Manual Iniciativo " & VersionDelJuego
End Sub

