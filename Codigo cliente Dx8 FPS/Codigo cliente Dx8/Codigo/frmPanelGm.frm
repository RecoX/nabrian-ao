VERSION 5.00
Begin VB.Form frmpanelgm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Torneo Panel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   53
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton Command9 
      Caption         =   "/VERTORNEO"
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
      Left            =   5640
      TabIndex        =   51
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Puede atacar GMS'"
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
      Left            =   5640
      TabIndex        =   50
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton Command42 
      Caption         =   "Pausar automáticos"
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
      Left            =   5640
      TabIndex        =   49
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton Command41 
      BackColor       =   &H000000FF&
      Caption         =   "Cancelar DeathMach"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton Command40 
      BackColor       =   &H0000FF00&
      Caption         =   "Abrir cupos DeathMach"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton Command39 
      BackColor       =   &H000000FF&
      Caption         =   "Cancelar Portales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton Command38 
      BackColor       =   &H000000FF&
      Caption         =   "Cancelar 2vs2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton Command37 
      BackColor       =   &H000000FF&
      Caption         =   "Cancelar Plantes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton Command36 
      BackColor       =   &H000000FF&
      Caption         =   "Cancelar 1vs1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton Command35 
      BackColor       =   &H0000FF00&
      Caption         =   "Abrir cupos Portales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton Command34 
      BackColor       =   &H0000FF00&
      Caption         =   "Abrir cupos 2vs2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H0000FF00&
      Caption         =   "Abrir cupos Plantes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H0000FF00&
      Caption         =   "Abrir cupos 1vs1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H000000FF&
      Caption         =   "BOTON LIBRE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H000080FF&
      Caption         =   "Bloquea zona"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H0000FF00&
      Caption         =   "Comenzar Carrera"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3000
      TabIndex        =   33
      Text            =   "0"
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Abrir cupos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Reglas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Anunciar que sale evento."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox textCarreras 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   29
      Text            =   "Hora Ejemplo (10:00:00)"
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Enviar Hora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   28
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SOPORTE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H00E0E0E0&
      Caption         =   "/SUBASTAR"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command26 
      BackColor       =   &H0000FF00&
      Caption         =   "Ver pregunta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H0000FF00&
      Caption         =   "Envar pista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   17
      Text            =   "Juega en argentina"
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H0000FF00&
      Caption         =   "Enviar pregunta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Text            =   "Messi"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Text            =   "Quien es el mejor jugador del mundo"
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Advertir usuario"
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
      TabIndex        =   12
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Ver procesos"
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
      TabIndex        =   11
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Banear usuario por tiempo"
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
      TabIndex        =   10
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Banear usuario"
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
      TabIndex        =   9
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Mandar carcel"
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
      TabIndex        =   8
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Buscar item"
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
      TabIndex        =   7
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H000000FF&
      Caption         =   "Desactivar/Activar Retos 1vs1,2vs2,3vs3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FF00&
      Caption         =   "Panel de Soportes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Panel The team"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Panel The Protector"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CFG AUTOMATICOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H00E0E0E0&
      Caption         =   "GRANPODER"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H008080FF&
      Caption         =   "BOTON LIBRE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   7920
      TabIndex        =   52
      Top             =   0
      Width           =   495
   End
   Begin VB.Line Line14 
      X1              =   5520
      X2              =   5520
      Y1              =   6600
      Y2              =   5040
   End
   Begin VB.Line Line11 
      X1              =   5520
      X2              =   11040
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label6 
      Caption         =   "Torneos/eventos automáticos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   5760
      TabIndex        =   38
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Evento carreras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   3000
      TabIndex        =   36
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label LabelText 
      Caption         =   "LabelText"
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
      Index           =   2
      Left            =   4680
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabelText 
      Caption         =   "LabelText"
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
      Index           =   1
      Left            =   4680
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabelText 
      Caption         =   "LabelText"
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
      Index           =   0
      Left            =   4800
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Pista:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Line Line13 
      X1              =   3480
      X2              =   3480
      Y1              =   4200
      Y2              =   6000
   End
   Begin VB.Label Label4 
      Caption         =   "Evento <PREGUNTAS>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Line Line12 
      X1              =   0
      X2              =   5520
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line10 
      X1              =   5520
      X2              =   11040
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line9 
      X1              =   0
      X2              =   5520
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line8 
      X1              =   3240
      X2              =   8760
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Retos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Paneles de eventos y sistemas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   5520
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   0
      X2              =   0
      Y1              =   4680
      Y2              =   1440
   End
   Begin VB.Line Line5 
      X1              =   5520
      X2              =   5520
      Y1              =   5880
      Y2              =   4320
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   0
      Y1              =   7680
      Y2              =   4440
   End
   Begin VB.Line Line3 
      X1              =   5520
      X2              =   5520
      Y1              =   5400
      Y2              =   2280
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   3120
      Y2              =   -120
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   5520
      X2              =   5520
      Y1              =   3240
      Y2              =   0
   End
End
Attribute VB_Name = "frmpanelgm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
FrmCfgAutomaticos.Show
Unload Me
End Sub

Private Sub Command11_Click()
Call SendData("/RETOACTIVADO")
Call SendData("/PAREJASACTIVADA")
Call SendData("/TRIEACTIVADO")
End Sub

Private Sub Command12_Click()
Call SendData("/BLOQUEAR")
End Sub

Private Sub Command13_Click()
Call SendData("/CARRERA")
End Sub

Private Sub Command14_Click()
LabelText(0).Caption = InputBox("Escribe el nombre del item que desea buscar.", "Buscar Item", "")
Call SendData("/BUSCAR " & LabelText(0).Caption)
End Sub

Private Sub Command15_Click()
LabelText(0).Caption = InputBox("Escribe el motivo.", "Motivo de carcel", "")
LabelText(1).Caption = InputBox("Escribe el Nick del usuario.", "Nick de usuario", "")
LabelText(2).Caption = InputBox("Escribe el Tiempo.", "Tiempo de carcel", "")
Call SendData("/CARCELTIEMPO " & LabelText(0).Caption & "@" & LabelText(1).Caption & "@" & LabelText(2).Caption)
End Sub

Private Sub Command16_Click()
LabelText(0).Caption = InputBox("Escribe el motivo.", "Motivo de baneo", "")
LabelText(1).Caption = InputBox("Escribe el Nick del usuario.", "Nick de usuario", "")
Call SendData("/BAN " & LabelText(0).Caption & "@" & LabelText(1).Caption)
End Sub

Private Sub Command17_Click()
LabelText(0).Caption = InputBox("Escribe el motivo.", "Motivo de ban", "")
LabelText(1).Caption = InputBox("Escribe el Nick del usuario.", "Nick de usuario", "")
LabelText(2).Caption = InputBox("Escribe el Tiempo.", "Tiempo de Ban", "")
Call SendData("/BANT " & LabelText(0).Caption & "@" & LabelText(1).Caption & "@" & LabelText(2).Caption)
End Sub

Private Sub Command18_Click()
LabelText(0).Caption = InputBox("Escribe el Nick del usuario.", "Ver procesos", "")
Call SendData("/VERPROCESOS " & LabelText(0).Caption)
End Sub

Private Sub Command19_Click()
LabelText(0).Caption = InputBox("Escribe el Nick del usuario.", "Advertir Usuario", "")
Call SendData("/ADVERTIR " & LabelText(0).Caption)
End Sub

Private Sub Command2_Click()
ProtectorPanel.Show
Unload frmpanelgm
End Sub

Private Sub Command20_Click()
Call SendData("/RMSG Reglas: Solo vale los hechizos inmovilizar y remover paralizis, no se puede entrar con clases sin mana.")
End Sub


Private Sub Command22_Click()
Call SendData("/RMSG En instantes realizaré una quest (Carreras), comando /PARTICIPAR, " & textCarreras.Text & " Hora del servidor abro.")
End Sub

Private Sub Command23_Click()
Call SendData("/HORA")
End Sub

Private Sub Command24_Click()
Call SendData("/PREG " & Text8.Text & "@" & Text9.Text)
End Sub

Private Sub Command25_Click()
Call SendData("/PISTAGM " & Text10.Text)
End Sub

Private Sub Command26_Click()
Call SendData("/RESP")
End Sub

Private Sub Command27_Click()
Call SendData("/ACTIVS")
End Sub

Private Sub Command28_Click()
Call SendData("/PODERACTIVADO")
End Sub

Private Sub Command29_Click()
Call SendData("/SOOPORTEACTIVADO")
End Sub

Private Sub Command3_Click()
Theteampanel.Show
Unload frmpanelgm
End Sub



Private Sub Command32_Click()
LabelText(0).Caption = InputBox("¿Cuantos cupos desea abrir? 1 = 2 ^^ 2 = 4 ^^ 3 = 8 ^^ 4 = 16 ^^ 5 = 32. ", "Abrir cupos", "")
Call SendData("/AUTO " & LabelText(0).Caption)
End Sub

Private Sub Command33_Click()
LabelText(0).Caption = InputBox("¿Cuantos cupos desea abrir? 1 = 2 ^^ 2 = 4 ^^ 3 = 8 ^^ 4 = 16 ^^ 5 = 32. ", "Abrir cupos", "")
Call SendData("/SATUROSP " & LabelText(0).Caption)
End Sub

Private Sub Command34_Click()
LabelText(0).Caption = InputBox("¿Cuantos cupos desea abrir? 1 = 2 ^^ 2 = 4 ^^ 3 = 8 ^^ 4 = 16 ^^ 5 = 32. ", "Abrir cupos", "")
Call SendData("/AUTO2V2 " & LabelText(0).Caption)
End Sub

Private Sub Command35_Click()
LabelText(0).Caption = InputBox("¿Cuantos cupos desea abrir?", "Abrir cupos", "")
Call SendData("/EVENTOPORTAL " & LabelText(0).Caption)
End Sub

Private Sub Command36_Click()
Call SendData("/CANCELAR")
End Sub

Private Sub Command37_Click()
Call SendData("/PCANCELAR")
End Sub

Private Sub Command38_Click()
Call SendData("/CANCELA2V2")
End Sub

Private Sub Command39_Click()
Call SendData("/PORTALCANCELAR")
End Sub

Private Sub Command4_Click()
Call SendData("/TORNEO " & Text1.Text)
End Sub

Private Sub Command40_Click()
LabelText(0).Caption = InputBox("¿Cuantos cupos desea abrir?", "Abrir cupos", "")
Call SendData("/DEATMAC " & LabelText(0).Caption)
End Sub

Private Sub Command41_Click()
Call SendData("/DCANCELAR")
End Sub

Private Sub Command42_Click()
Call SendData("/PAUSAR")
End Sub

Private Sub Command5_Click()
Call SendData("/DAMESOS")
Unload frmpanelgm
End Sub

Private Sub Command6_Click()
FrmTorneoModalidad.Show , frmPrincipal
Unload Me
End Sub

Private Sub Command8_Click()
Call SendData("/PUEDE")
End Sub

Private Sub Command9_Click()
Call SendData("/VERTORNEO")
End Sub

Private Sub Label7_Click()
Command31.Enabled = True
Command29.Enabled = True
Command27.Enabled = True
Command28.Enabled = True
Command21.Enabled = True
Command11.Enabled = True
End Sub
