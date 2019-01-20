VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   12075.47
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCorreo2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      Left            =   345
      TabIndex        =   33
      Top             =   3240
      Width           =   4080
   End
   Begin VB.TextBox txtPasswdCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   345
      PasswordChar    =   "*"
      TabIndex        =   35
      Top             =   4470
      Width           =   4080
   End
   Begin VB.TextBox txtPasswd 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   345
      PasswordChar    =   "*"
      TabIndex        =   34
      Top             =   3840
      Width           =   4080
   End
   Begin VB.TextBox txtCorreo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      Left            =   345
      TabIndex        =   32
      Top             =   2625
      Width           =   4080
   End
   Begin VB.ComboBox lstGenero 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonaje.frx":0000
      Left            =   5160
      List            =   "frmCrearPersonaje.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   1950
      Width           =   2753
   End
   Begin VB.ComboBox lstRaza 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonaje.frx":001D
      Left            =   5160
      List            =   "frmCrearPersonaje.frx":0030
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   1350
      Width           =   2753
   End
   Begin VB.ComboBox lstHogar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonaje.frx":005D
      Left            =   5160
      List            =   "frmCrearPersonaje.frx":006D
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   2580
      Width           =   2753
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   300
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   0
      Top             =   780
      Width           =   3615
   End
   Begin VB.Label modCarisma 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6718
      TabIndex        =   49
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label modAgilidad 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6720
      TabIndex        =   48
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label modConstitucion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6718
      TabIndex        =   47
      Top             =   5850
      Width           =   375
   End
   Begin VB.Label modInteligencia 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6720
      TabIndex        =   46
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label modfuerza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6718
      TabIndex        =   45
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblPass2OK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   4560
      TabIndex        =   44
      Top             =   3750
      Width           =   345
   End
   Begin VB.Label lbSabiduria 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "+3"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   180
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblMailOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   4560
      TabIndex        =   40
      Top             =   2520
      Width           =   330
   End
   Begin VB.Label lblMail2OK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   4560
      TabIndex        =   38
      Top             =   3150
      Width           =   345
   End
   Begin VB.Label lblPassOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   4560
      TabIndex        =   36
      Top             =   4350
      Width           =   345
   End
   Begin VB.Image Picture6 
      Height          =   375
      Left            =   5160
      MouseIcon       =   "frmCrearPersonaje.frx":0096
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image Picture5 
      Height          =   375
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":03A0
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image Picture2 
      Height          =   375
      Left            =   5520
      MouseIcon       =   "frmCrearPersonaje.frx":06AA
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image Picture1 
      Height          =   375
      Left            =   7080
      MouseIcon       =   "frmCrearPersonaje.frx":09B4
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   375
   End
   Begin VB.Image Picture8 
      Height          =   255
      Left            =   5280
      MouseIcon       =   "frmCrearPersonaje.frx":0CBE
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   255
   End
   Begin VB.Image Picture7 
      Height          =   255
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":0FC8
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   255
   End
   Begin VB.Image Picture4 
      Height          =   255
      Left            =   5520
      MouseIcon       =   "frmCrearPersonaje.frx":12D2
      MousePointer    =   99  'Custom
      Top             =   4440
      Width           =   255
   End
   Begin VB.Image Picture3 
      Height          =   255
      Left            =   7200
      MouseIcon       =   "frmCrearPersonaje.frx":15DC
      MousePointer    =   99  'Custom
      Top             =   4440
      Width           =   255
   End
   Begin VB.Image Picture10 
      Height          =   255
      Left            =   5640
      MouseIcon       =   "frmCrearPersonaje.frx":18E6
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   255
   End
   Begin VB.Image Picture9 
      Height          =   375
      Left            =   7080
      MouseIcon       =   "frmCrearPersonaje.frx":1BF0
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   23
      Left            =   11031
      TabIndex        =   31
      Top             =   7800
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   22
      Left            =   11031
      TabIndex        =   30
      Top             =   7500
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   21
      Left            =   11031
      TabIndex        =   29
      Top             =   7200
      Width           =   398
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   46
      Left            =   11415
      MouseIcon       =   "frmCrearPersonaje.frx":1EFA
      MousePointer    =   99  'Custom
      Top             =   7905
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   44
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":204C
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   42
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":219E
      MousePointer    =   99  'Custom
      Top             =   7290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   47
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":22F0
      MousePointer    =   99  'Custom
      Top             =   7920
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   45
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":2442
      MousePointer    =   99  'Custom
      Top             =   7605
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   43
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":2594
      MousePointer    =   99  'Custom
      Top             =   7305
      Width           =   195
   End
   Begin VB.Label puntosquedan 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "28"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6300
      TabIndex        =   28
      Top             =   2955
      Width           =   255
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   8865
      TabIndex        =   27
      Top             =   525
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   3
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":26E6
      MousePointer    =   99  'Custom
      Top             =   1185
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   5
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":2838
      MousePointer    =   99  'Custom
      Top             =   1440
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   7
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":298A
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   9
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":2ADC
      MousePointer    =   99  'Custom
      Top             =   2070
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   11
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":2C2E
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   13
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":2D80
      MousePointer    =   99  'Custom
      Top             =   2700
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   15
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":2ED2
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   17
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":3024
      MousePointer    =   99  'Custom
      Top             =   3270
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   19
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":3176
      MousePointer    =   99  'Custom
      Top             =   3615
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   21
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":32C8
      MousePointer    =   99  'Custom
      Top             =   3945
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   23
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":341A
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   25
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":356C
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   27
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":36BE
      MousePointer    =   99  'Custom
      Top             =   4815
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   1
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":3810
      MousePointer    =   99  'Custom
      Top             =   840
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":3962
      MousePointer    =   99  'Custom
      Top             =   870
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   2
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":3AB4
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":3C06
      MousePointer    =   99  'Custom
      Top             =   1500
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   6
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":3D58
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   8
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":3EAA
      MousePointer    =   99  'Custom
      Top             =   2085
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":3FFC
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":414E
      MousePointer    =   99  'Custom
      Top             =   2730
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   14
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":42A0
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   16
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":43F2
      MousePointer    =   99  'Custom
      Top             =   3360
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   18
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":4544
      MousePointer    =   99  'Custom
      Top             =   3630
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":4696
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   22
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":47E8
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":493A
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   26
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":4A8C
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   28
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":4BDE
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   29
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":4D30
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":4E82
      MousePointer    =   99  'Custom
      Top             =   5490
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   31
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":4FD4
      MousePointer    =   99  'Custom
      Top             =   5430
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":5126
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":5278
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":53CA
      MousePointer    =   99  'Custom
      Top             =   6105
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   35
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":551C
      MousePointer    =   99  'Custom
      Top             =   6090
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   36
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":566E
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   37
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":57C0
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   38
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":5912
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   39
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":5A64
      MousePointer    =   99  'Custom
      Top             =   6705
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   11400
      MouseIcon       =   "frmCrearPersonaje.frx":5BB6
      MousePointer    =   99  'Custom
      Top             =   6990
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   10875
      MouseIcon       =   "frmCrearPersonaje.frx":5D08
      MousePointer    =   99  'Custom
      Top             =   6990
      Width           =   135
   End
   Begin VB.Image boton 
      Height          =   255
      Index           =   1
      Left            =   120
      MouseIcon       =   "frmCrearPersonaje.frx":5E5A
      MousePointer    =   99  'Custom
      Top             =   8640
      Width           =   1125
   End
   Begin VB.Image boton 
      Appearance      =   0  'Flat
      Height          =   450
      Index           =   0
      Left            =   360
      MouseIcon       =   "frmCrearPersonaje.frx":5FAC
      MousePointer    =   99  'Custom
      Top             =   7800
      Width           =   4560
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   20
      Left            =   11031
      TabIndex        =   26
      Top             =   6900
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   19
      Left            =   11031
      TabIndex        =   25
      Top             =   6600
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   18
      Left            =   11031
      TabIndex        =   24
      Top             =   6285
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   17
      Left            =   11031
      TabIndex        =   23
      Top             =   5970
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   16
      Left            =   11031
      TabIndex        =   22
      Top             =   5685
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   15
      Left            =   11031
      TabIndex        =   21
      Top             =   5385
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   14
      Left            =   11031
      TabIndex        =   20
      Top             =   5070
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   13
      Left            =   11031
      TabIndex        =   19
      Top             =   4770
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   12
      Left            =   11031
      TabIndex        =   18
      Top             =   4470
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   11
      Left            =   11031
      TabIndex        =   17
      Top             =   4155
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   10
      Left            =   11031
      TabIndex        =   16
      Top             =   3840
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   9
      Left            =   11031
      TabIndex        =   15
      Top             =   3540
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   8
      Left            =   11031
      TabIndex        =   14
      Top             =   3225
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   7
      Left            =   11031
      TabIndex        =   13
      Top             =   2925
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   6
      Left            =   11031
      TabIndex        =   12
      Top             =   2610
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   5
      Left            =   11031
      TabIndex        =   11
      Top             =   2310
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   4
      Left            =   11031
      TabIndex        =   10
      Top             =   2010
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   3
      Left            =   11031
      TabIndex        =   9
      Top             =   1710
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   2
      Left            =   11031
      TabIndex        =   8
      Top             =   1395
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   0
      Left            =   11031
      TabIndex        =   7
      Top             =   780
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   1
      Left            =   11031
      TabIndex        =   6
      Top             =   1080
      Width           =   398
   End
   Begin VB.Label lbCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   6251
      TabIndex        =   5
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   6255
      TabIndex        =   4
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   6251
      TabIndex        =   3
      Top             =   5730
      Width           =   495
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   6255
      TabIndex        =   2
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label lbFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   6255
      TabIndex        =   1
      Top             =   3840
      Width           =   495
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public SkillPoints As Byte
Function CheckData() As Boolean

If UserRaza = 0 Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserHogar = 0 Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

If Val(puntosquedan.Caption) > 0 Then
    MsgBox "Asigne los atributos del personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    MsgBox "Asigne los skillpoints del personaje."
    Exit Function
End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    If UserAtributos(i) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

CheckData = True


End Function

Private Sub boton_Click(Index As Integer)

Call PlayWaveDS(SND_CLICK)

Select Case Index
    Case 0
        LlegoConfirmacion = False
        Confirmacion = 0
        Dim i As Integer
        Dim k As Object
        i = 1
        For Each k In Skill
            UserSkills(i) = k.Caption
            i = i + 1
        Next
        
        UserName = txtNombre.Text
        
        If Right$(UserName, 1) = " " Then
            UserName = RTrim(UserName)
            MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
        
        UserRaza = lstRaza.ListIndex + 1
        UserSexo = lstGenero.ListIndex
        UserHogar = lstHogar.ListIndex + 1
        
        UserAtributos(1) = Val(lbFuerza.Caption)
        UserAtributos(2) = Val(lbAgilidad.Caption)
        UserAtributos(3) = Val(lbInteligencia.Caption)
        UserAtributos(4) = Val(lbCarisma.Caption)
        UserAtributos(5) = Val(lbConstitucion.Caption)
        
        If UserAtributos(1) + UserAtributos(3) + UserAtributos(4) + UserAtributos(2) + UserAtributos(5) > 63 Then Exit Sub
        
     If CheckData() Then
     
     If CheckDatos() Then
    UserPassword = MD5String(txtPasswd.Text)
    UserEmail = txtCorreo.Text
    
    If Not CheckMailString(UserEmail) Then
            MsgBox "Direccion de mail invalida.", vbExclamation, "Fenix AO"
            Exit Sub
    End If
    
    If Trim(txtPasswd) = "" Or Trim(txtPasswdCheck) = "" Then
        MsgBox "Tenés que ingresar una contraseña, no se aceptan espacios en blanco.", vbInformation, "Fenix AO"
        txtPasswd = ""
        txtPasswdCheck = ""
        txtPasswd.SetFocus
        Exit Sub
    End If
        
    frmMain.Socket1.HostName = IPdelServidor
    frmMain.Socket1.RemotePort = PuertoDelServidor

    'SendNewChar = True
    Me.MousePointer = 11
    EstadoLogin = CrearNuevoPj
    
    If Not frmMain.Socket1.Connected Then
        MsgBox "Error: Se ha perdido la conexion con el server."
        Unload Me
        
    Else
        Call Login(ValidarLoginMSG(CInt(bRK)))
    End If
            If Musica = 0 Then
            CurMidi = DirMidi & "2.mid"
            LoopMidi = 1
            Call CargarMIDI(CurMidi)
            Call Play_Midi
        End If
        
    frmConnect.Picture = LoadPicture(App.Path & "\Graficos\conectar.jpg")
    Me.Visible = False
   '         Do While Not LlegoConfirmacion
   '             DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
   '         Loop
   '         If Confirmacion = 1 Then
   '    MsgBox "Un mail ha sido enviado a la casilla a la cual registraste tu personaje. Utiliza el código que en él aparece para así poder activar el personaje en el formulario correspondiente de Activar/desactivar personaje. Si no recibes el email, chequea la seguridad de tu cuenta y revisa la sección de correo no deseado para comprobar que no haya llegado allí."
   '
   '     frmConnect.FONDO.Picture = LoadPicture(App.Path & "\Graficos\conectar.jpg")
   '
   '     frmMain.Socket1.Disconnect
   '     frmConnect.MousePointer = 1
   '   Unload Me
   '   ElseIf Confirmacion = 2 Then
   '           EstadoLogin = dados
   '           Me.MousePointer = 1
        Else
        MsgBox "Error al crear el personaje, intentelo otra vez"
     End If
      
      
End If



    Case 1
        If Musica = 0 Then
            CurMidi = DirMidi & "2.mid"
            LoopMidi = 1
            Call CargarMIDI(CurMidi)
            Call Play_Midi
        End If
        
        frmConnect.Picture = LoadPicture(App.Path & "\Graficos\conectar.jpg")
        
        frmMain.Socket1.Disconnect
        frmConnect.MousePointer = 1
      Unload Me
      
End Select


End Sub
Private Sub command1_Click(Index As Integer)
Call PlayWaveDS(SND_CLICK)

Dim indice
If Index Mod 2 = 0 Then
    If SkillPoints > 0 Then
        indice = Index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = Index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

Puntos.Caption = SkillPoints
End Sub

Function CheckDatos() As Boolean

If txtPasswd.Text <> txtPasswdCheck.Text Then
    MsgBox "Las contraseñas que ingresaste no coinciden." & vbCrLf & vbCrLf & "Por favor volvé a ingresarlos.", vbInformation, "Fenix AO"
    txtPasswdCheck = ""
    txtPasswdCheck.SetFocus
    Exit Function
End If

If txtCorreo.Text <> txtCorreo2.Text Then
    MsgBox "Las direcciones de correo electrónico ingresadas no coinciden." & vbCrLf & vbCrLf & "Por favor volvé a ingresarlos.", vbInformation, "Fenix AO"
    txtCorreo2.SetFocus
    Exit Function
End If

CheckDatos = True

End Function
Private Sub Form_Load()

SkillPoints = 10
Puntos.Caption = SkillPoints
Me.Picture = LoadPicture(App.Path & "\graficos\CP-Interface.gif")

Select Case (lstRaza.List(lstRaza.ListIndex))
    Case "Humano"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = "+ 2"
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = "+ 1"
        modCarisma.Caption = ""
        
    Case "Elfo"
        modfuerza.Caption = "- 1"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+ 3"
        modInteligencia.Caption = "+ 2"
        modCarisma.Caption = "+ 2"
        
    Case "Elfo Oscuro"
        modfuerza.Caption = "- 1"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+ 2"
        modInteligencia.Caption = "+ 2"
        modCarisma.Caption = "- 1"
        
    Case "Enano"
        modfuerza.Caption = "+ 3"
        modConstitucion.Caption = "+ 3"
        modAgilidad.Caption = ""
        modInteligencia.Caption = "- 4"
        modCarisma.Caption = "- 1"
        
    Case "Gnomo"
        modfuerza.Caption = "- 4"
        modConstitucion.Caption = ""
        modAgilidad.Caption = ""
        modInteligencia.Caption = "+ 3"
        modCarisma.Caption = "+ 1"
        
End Select

End Sub

Private Sub Pîcture4_Click()

End Sub

Private Sub lstRaza_click()

Select Case (lstRaza.List(lstRaza.ListIndex))
    Case "Humano"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = "+ 2"
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = ""
        modCarisma.Caption = ""
        
    Case "Elfo"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = "+ 1"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "+ 2"
        modCarisma.Caption = "+ 2"
        
    Case "Elfo Oscuro"
        modfuerza.Caption = "+ 2"
        modConstitucion.Caption = "+ 1"
        modAgilidad.Caption = "+ 2"
        modInteligencia.Caption = "+ 1"
        modCarisma.Caption = "- 3"
        
    Case "Enano"
        modfuerza.Caption = "+ 3"
        modConstitucion.Caption = "+ 3"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "- 6"
        modCarisma.Caption = "- 2"
        
    Case "Gnomo"
        modfuerza.Caption = "- 4"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+ 3"
        modInteligencia.Caption = "+ 3"
        modCarisma.Caption = "+ 2"
        
    End Select

End Sub

Private Sub txtCorreo_Change()

If Not CheckMailString(txtCorreo) Then
    lblMailOK = "O"
    lblMailOK.ForeColor = &HC0&
Else
    lblMailOK = "P"
    lblMailOK.ForeColor = &H80FF&
    Beep
End If

If Not (UCase$(txtCorreo.Text) = UCase$(txtCorreo2.Text)) Or (txtCorreo2.Text) = "" Or Not CheckMailString(txtCorreo2) Then
    lblMail2OK = "O"
    lblMail2OK.ForeColor = &HC0&
Else
    lblMail2OK = "P"
    lblMail2OK.ForeColor = &H80FF&
    Beep
End If

End Sub
Private Sub txtCorreo_GotFocus()

Call MsgBox("La dirección de correo electrónico DEBE SER real.")

End Sub
Private Sub txtCorreo2_Change()

If Not CheckMailString(txtCorreo) Then
    lblMailOK = "O"
    lblMailOK.ForeColor = &HC0&
Else
    lblMailOK = "P"
    lblMailOK.ForeColor = &H80FF&
    Beep
End If

If Not (UCase$(txtCorreo.Text) = UCase$(txtCorreo2.Text)) Or (txtCorreo2.Text) = "" Or Not CheckMailString(txtCorreo2) Then
    lblMail2OK = "O"
    lblMail2OK.ForeColor = &HC0&
Else
    lblMail2OK = "P"
    lblMail2OK.ForeColor = &H80FF&
    Beep
End If

End Sub
Private Sub txtPasswd_Change()

If Trim(txtPasswd) = "" Then
    lblPass2OK = "O"
    lblPass2OK.ForeColor = &HC0&
Else
    lblPass2OK = "P"
    lblPass2OK.ForeColor = &H80FF&
    Beep
End If

If (txtPasswdCheck = txtPasswd) And txtPasswd <> "" Then
    lblPassOK = "P"
    lblPassOK.ForeColor = &H80FF&
    Beep
Else
    lblPassOK = "O"
    lblPassOK.ForeColor = &HC0&
End If

End Sub
Private Sub txtPasswdCheck_Change()

If Trim(txtPasswdCheck) = "" And Trim(txtPasswdCheck) = "" Then
    lblPassOK = "O"
    lblPassOK.ForeColor = &HC0&
Else
    If (txtPasswdCheck = txtPasswd) And txtPasswd <> "" Then
        lblPassOK = "P"
        lblPassOK.ForeColor = &H80FF&
        Beep
    Else
        lblPassOK = "O"
        lblPassOK.ForeColor = &HC0&
    End If
End If

End Sub
Private Sub Picture1_Click()

Call PlayWaveDS(SND_CLICK)

If Val(puntosquedan.Caption) > 0 Then
    If Val(lbCarisma.Caption) < 18 Then
    lbCarisma.Caption = Val(lbCarisma.Caption) + 1
    puntosquedan.Caption = Val(puntosquedan.Caption) - 1
    End If
End If

End Sub
Private Sub Picture2_Click()

Call PlayWaveDS(SND_CLICK)

If Val(lbCarisma.Caption) > 6 Then
    lbCarisma.Caption = Val(lbCarisma.Caption) - 1
    puntosquedan.Caption = Val(puntosquedan.Caption) + 1
End If

End Sub
Private Sub Picture3_Click()

Call PlayWaveDS(SND_CLICK)

If Val(puntosquedan.Caption) > 0 Then
    If Val(lbInteligencia.Caption) < 18 Then
    lbInteligencia.Caption = Val(lbInteligencia.Caption) + 1
    puntosquedan.Caption = Val(puntosquedan.Caption) - 1
    End If
End If

End Sub
Private Sub Picture4_Click()

Call PlayWaveDS(SND_CLICK)

If Val(lbInteligencia.Caption) > 6 Then
    lbInteligencia.Caption = Val(lbInteligencia.Caption) - 1
    puntosquedan.Caption = Val(puntosquedan.Caption) + 1
End If

End Sub
Private Sub Picture5_Click()

Call PlayWaveDS(SND_CLICK)

If Val(puntosquedan.Caption) > 0 Then
    If Val(lbConstitucion.Caption) < 18 Then
    lbConstitucion.Caption = Val(lbConstitucion.Caption) + 1
    puntosquedan.Caption = Val(puntosquedan.Caption) - 1
    End If
End If

End Sub
Private Sub Picture6_Click()

Call PlayWaveDS(SND_CLICK)

If Val(lbConstitucion.Caption) > 6 Then
    lbConstitucion.Caption = Val(lbConstitucion.Caption) - 1
    puntosquedan.Caption = Val(puntosquedan.Caption) + 1
End If

End Sub
Private Sub Picture7_Click()

Call PlayWaveDS(SND_CLICK)

If Val(puntosquedan.Caption) > 0 Then
    If Val(lbAgilidad.Caption) < 18 Then
    lbAgilidad.Caption = Val(lbAgilidad.Caption) + 1
    puntosquedan.Caption = Val(puntosquedan.Caption) - 1
    End If
End If

End Sub
Private Sub Picture8_Click()

Call PlayWaveDS(SND_CLICK)

If Val(lbAgilidad.Caption) > 6 Then
    lbAgilidad.Caption = Val(lbAgilidad.Caption) - 1
    puntosquedan.Caption = Val(puntosquedan.Caption) + 1
End If

End Sub
Private Sub Picture10_Click()

Call PlayWaveDS(SND_CLICK)

If Val(lbFuerza.Caption) > 6 Then
    lbFuerza.Caption = Val(lbFuerza.Caption) - 1
    puntosquedan.Caption = Val(puntosquedan.Caption) + 1
End If

End Sub
Private Sub Picture9_Click()

Call PlayWaveDS(SND_CLICK)

If Val(puntosquedan.Caption) > 0 Then
    If Val(lbFuerza.Caption) < 18 Then
    lbFuerza.Caption = Val(lbFuerza.Caption) + 1
    puntosquedan.Caption = Val(puntosquedan.Caption) - 1
    End If
End If

End Sub
Private Sub txtNombre_Change()

txtNombre.Text = LTrim(txtNombre.Text)

End Sub
Private Sub txtNombre_GotFocus()

Call MsgBox("Sea cuidadoso al seleccionar el nombre de su personaje, Argentum es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotación politica los administradores borrarán su personaje y no habrá ninguna posibilidad de recuperarlo.")

End Sub
Private Sub txtNombre_KeyPress(KeyAscii As Integer)

KeyAscii = Asc(UCase$(Chr(KeyAscii)))
 
End Sub
