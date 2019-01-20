VERSION 5.00
Begin VB.Form FrmTorneoModalidad 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Torneo hasta 4vs4"
   ClientHeight    =   8220
   ClientLeft      =   3345
   ClientTop       =   615
   ClientWidth     =   11820
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   3  'Dash-Dot
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8220
   ScaleMode       =   0  'User
   ScaleWidth      =   11820
   Begin VB.CommandButton Command68 
      Caption         =   "ACC"
      Height          =   195
      Left            =   10320
      TabIndex        =   125
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command63 
      BackColor       =   &H00000000&
      Caption         =   "Enviar Reglas"
      Height          =   255
      Left            =   10320
      MaskColor       =   &H00808080&
      TabIndex        =   94
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Caption         =   "Jugadores 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   83
      Top             =   3000
      Width           =   5415
      Begin VB.CommandButton Command33 
         Caption         =   "2-1 Ganan"
         Height          =   255
         Left            =   2400
         TabIndex        =   93
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command34 
         Caption         =   "3-1 Ganan"
         Height          =   255
         Left            =   3480
         TabIndex        =   92
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command32 
         Caption         =   "2-0 Ganan"
         Height          =   255
         Left            =   2400
         TabIndex        =   91
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command57 
         Caption         =   "3-2 Ganan"
         Height          =   255
         Left            =   3480
         TabIndex        =   90
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command36 
         Caption         =   "Empa 2-2"
         Height          =   255
         Left            =   1320
         TabIndex        =   89
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Anunciar Pelea"
         Height          =   255
         Left            =   0
         TabIndex        =   88
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command35 
         Caption         =   "Empa 1-1"
         Height          =   255
         Left            =   1320
         TabIndex        =   87
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command67 
         Caption         =   "3.0 Ganan"
         Height          =   255
         Left            =   3480
         TabIndex        =   86
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Llamar Equipos"
         Height          =   255
         Left            =   0
         TabIndex        =   85
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command31 
         Caption         =   "1-0 Ganan"
         Height          =   255
         Left            =   2400
         TabIndex        =   84
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command66 
      Caption         =   "HORA"
      Height          =   255
      Left            =   10320
      TabIndex        =   74
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   10320
      TabIndex        =   73
      Text            =   "NUMERO DE CUANTAS PAREJAS"
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   10320
      TabIndex        =   72
      Text            =   "HORA DEL EVENTO Ej: (19:50:00)"
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command65 
      BackColor       =   &H00000000&
      Caption         =   "Repetir lo de arriva"
      Height          =   255
      Left            =   10320
      MaskColor       =   &H00808080&
      TabIndex        =   71
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command64 
      BackColor       =   &H00000000&
      Caption         =   "Anunciar que vas a hacer el evento."
      Height          =   375
      Left            =   10320
      MaskColor       =   &H00808080&
      TabIndex        =   70
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10320
      TabIndex        =   66
      Text            =   "0"
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command45 
      Caption         =   "Abrir cupos"
      Height          =   255
      Left            =   10320
      TabIndex        =   65
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      Caption         =   "Jugadores 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   62
      Top             =   7080
      Width           =   5415
      Begin VB.CommandButton Command44 
         Caption         =   "2-1 Ganan"
         Height          =   255
         Left            =   2520
         TabIndex        =   101
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command55 
         Caption         =   "2-0 Ganan"
         Height          =   255
         Left            =   2520
         TabIndex        =   100
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command58 
         Caption         =   "3-2 Ganan"
         Height          =   255
         Left            =   3480
         TabIndex        =   99
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command43 
         Caption         =   "3-1 Ganan"
         Height          =   255
         Left            =   3480
         TabIndex        =   98
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Anunciar Pelea"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command53 
         Caption         =   "Empa 2-2"
         Height          =   255
         Left            =   1440
         TabIndex        =   96
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Llamar Equipos"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command62 
         Caption         =   "3-0 Ganan"
         Height          =   255
         Left            =   3480
         TabIndex        =   69
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command56 
         Caption         =   "1-0 Ganan"
         Height          =   255
         Left            =   2520
         TabIndex        =   64
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command54 
         Caption         =   "Empa 1-1"
         Height          =   255
         Left            =   1440
         TabIndex        =   63
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00000000&
      Caption         =   "Jugadores 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1095
      Left            =   5640
      TabIndex        =   60
      Top             =   7080
      Width           =   4575
      Begin VB.CommandButton Command50 
         Caption         =   "2-1 Ganan"
         Height          =   255
         Left            =   1200
         TabIndex        =   107
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command51 
         Caption         =   "2-0 Ganan"
         Height          =   255
         Left            =   1200
         TabIndex        =   106
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command52 
         Caption         =   "1-0 Ganan"
         Height          =   255
         Left            =   1200
         TabIndex        =   105
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command59 
         Caption         =   "3-2 Ganan"
         Height          =   255
         Left            =   2280
         TabIndex        =   104
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command49 
         Caption         =   "3-1 Ganan"
         Height          =   255
         Left            =   2280
         TabIndex        =   103
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command47 
         Caption         =   "Empatan 2-2"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command61 
         Caption         =   "3-0 Ganan"
         Height          =   255
         Left            =   2280
         TabIndex        =   68
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command48 
         Caption         =   "Empatan 1-1"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Primer Enfrentamiento"
      Height          =   495
      Left            =   10320
      TabIndex        =   59
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command29 
      Caption         =   "CERRAR"
      Height          =   195
      Left            =   10320
      TabIndex        =   53
      Top             =   7980
      Width           =   1455
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Cuenta"
      Height          =   375
      Left            =   11040
      TabIndex        =   50
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   10320
      TabIndex        =   49
      Text            =   "5"
      Top             =   3240
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Jugadores Ronda Nº2"
      ForeColor       =   &H000080FF&
      Height          =   2775
      Left            =   120
      TabIndex        =   25
      Top             =   4200
      Width           =   10095
      Begin VB.ComboBox Combo7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7080
         TabIndex        =   116
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox Combo6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7080
         TabIndex        =   115
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5520
         TabIndex        =   114
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5520
         TabIndex        =   113
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         TabIndex        =   112
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         TabIndex        =   111
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   110
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox Combo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   109
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1395
         Left            =   5400
         TabIndex        =   108
         Top             =   480
         Width           =   4575
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Pasan "
         Height          =   255
         Left            =   6720
         TabIndex        =   45
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Pierden"
         Height          =   255
         Left            =   8280
         TabIndex        =   44
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   8760
         TabIndex        =   32
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Añadir"
         Height          =   255
         Left            =   8760
         TabIndex        =   31
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   3480
         TabIndex        =   30
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Añadir"
         Height          =   255
         Left            =   3480
         TabIndex        =   29
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1395
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   4935
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Pasan "
         Height          =   255
         Left            =   1800
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Pierden"
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label26 
         BackColor       =   &H00000000&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5040
         TabIndex        =   52
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label25 
         BackColor       =   &H00000000&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5280
         TabIndex        =   51
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6360
         TabIndex        =   43
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   42
         Top             =   270
         Width           =   255
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         Caption         =   "Total de equipos:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5040
         TabIndex        =   41
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         Caption         =   "Total de equipos:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5100
         TabIndex        =   39
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5100
         TabIndex        =   38
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label19 
         BackColor       =   &H00000000&
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5100
         TabIndex        =   36
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label20 
         BackColor       =   &H00000000&
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5100
         TabIndex        =   35
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label21 
         BackColor       =   &H00000000&
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5100
         TabIndex        =   34
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label22 
         BackColor       =   &H00000000&
         Caption         =   "Vs.  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5100
         TabIndex        =   33
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Pasan"
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pasan"
      Height          =   255
      Left            =   6840
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Jugadores Ronda Nº1"
      ForeColor       =   &H000080FF&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.ComboBox Combo15 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6960
         TabIndex        =   124
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox Combo14 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6960
         TabIndex        =   123
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox Combo13 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5400
         TabIndex        =   122
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox Combo12 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5400
         TabIndex        =   121
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox Combo11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         TabIndex        =   120
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox Combo10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         TabIndex        =   119
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ComboBox Combo9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   118
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox Combo8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   117
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Pierden"
         Height          =   255
         Left            =   8400
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pierden"
         Height          =   255
         Left            =   3360
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Pasan "
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   8640
         TabIndex        =   14
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Añadir"
         Height          =   255
         Left            =   8640
         TabIndex        =   13
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Añadir"
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1395
         Left            =   5280
         TabIndex        =   2
         Top             =   480
         Width           =   4695
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1395
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label24 
         BackColor       =   &H00000000&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5160
         TabIndex        =   22
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label23 
         BackColor       =   &H00000000&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4920
         TabIndex        =   21
         Top             =   2400
         Width           =   135
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6480
         TabIndex        =   18
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   270
         Width           =   255
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Total de equipos:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5160
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Total de equipos:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4920
         TabIndex        =   10
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4920
         TabIndex        =   9
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4920
         TabIndex        =   7
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4920
         TabIndex        =   6
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Vs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4920
         TabIndex        =   5
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Vs.  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4920
         TabIndex        =   4
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Zona A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1455
      Left            =   10320
      TabIndex        =   46
      Top             =   240
      Width           =   1455
      Begin VB.CommandButton Command22 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   120
         TabIndex        =   77
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Cargar"
         Height          =   375
         Left            =   120
         TabIndex        =   76
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Borrar"
         Height          =   375
         Left            =   120
         TabIndex        =   75
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Zona B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1455
      Left            =   10320
      TabIndex        =   47
      Top             =   3960
      Width           =   1455
      Begin VB.CommandButton Command24 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   120
         TabIndex        =   80
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Cargar"
         Height          =   375
         Left            =   120
         TabIndex        =   79
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Borrar"
         Height          =   375
         Left            =   120
         TabIndex        =   78
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Jugadores 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1095
      Left            =   5640
      TabIndex        =   48
      Top             =   3000
      Width           =   4575
      Begin VB.CommandButton Command46 
         Caption         =   "3-2 Ganan"
         Height          =   255
         Left            =   2280
         TabIndex        =   82
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command40 
         Caption         =   "3-1 Ganan"
         Height          =   255
         Left            =   2280
         TabIndex        =   81
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command60 
         Caption         =   "3-0 Ganan"
         Height          =   255
         Left            =   2280
         TabIndex        =   67
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command42 
         Caption         =   "Empatan 2-2"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Empatan 1-1"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command39 
         Caption         =   "2-1 Ganan"
         Height          =   255
         Left            =   1200
         TabIndex        =   56
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command38 
         Caption         =   "2-0 Ganan"
         Height          =   255
         Left            =   1200
         TabIndex        =   55
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command37 
         Caption         =   "1-0 Ganan"
         Height          =   255
         Left            =   1200
         TabIndex        =   54
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Menu msg 
      Caption         =   "&Mensajes"
      Begin VB.Menu msg1 
         Caption         =   "¡Una extraordinaria pelea, ninguno de los equipos logra sacarse ventaja!"
         Shortcut        =   ^A
      End
      Begin VB.Menu msg2 
         Caption         =   "¡Esta pelea deja de que hablar, ambos equipos están dando un gran espectáculo!"
         Shortcut        =   ^B
      End
      Begin VB.Menu msg3 
         Caption         =   "Una pelea digna de ver, unos de los mejores enfrentamientos de este evento…"
         Shortcut        =   ^C
      End
      Begin VB.Menu msg4 
         Caption         =   "Remos, inmos, y mas inmos… ¡Que pelea muchachos!"
         Shortcut        =   ^D
      End
      Begin VB.Menu msg5 
         Caption         =   "¡Esquinas! ¡Mucha Suerte! Comienza en...    "
         Shortcut        =   ^E
      End
      Begin VB.Menu msg6 
         Caption         =   "Inmos, remos, apocas, una gran batalla, aun no se ha visto lo mejor!"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnunpc 
      Caption         =   "&NPC's"
      Begin VB.Menu bove 
         Caption         =   "Bóveda"
      End
      Begin VB.Menu sacer 
         Caption         =   "Sacerdote"
      End
      Begin VB.Menu potas 
         Caption         =   "Pociones"
      End
   End
End
Attribute VB_Name = "FrmTorneoModalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub bove_Click()
Call SendData("/ACC 57")
End Sub



Private Sub command1_Click()
If List3.ListIndex <> -1 Then
List3.RemoveItem List3.ListIndex
Label12 = List3.ListCount
End If
End Sub

Private Sub Command10_Click()
If List2.ListIndex <> -1 Then
List2.RemoveItem List2.ListIndex
Label11 = List2.ListCount

End If
End Sub

Private Sub Command11_Click()
If List4.ListIndex <> -1 Then
List4.RemoveItem List4.ListIndex
Label13 = List4.ListCount
End If
End Sub

Private Sub Command12_Click()
List4.AddItem Combo.Text & "-" & Combo1.Text & "-" & Combo2.Text & "-" & Combo3.Text
Label13 = List4.ListCount
End Sub

Private Sub Command13_Click()
Call SendData("/SUM " & ReadFieldOptimizado(1, List4, Asc("-")))
Call SendData("/SUM " & ReadFieldOptimizado(2, List4, Asc("-")))
Call SendData("/SUM " & ReadFieldOptimizado(3, List4, Asc("-")))
Call SendData("/SUM " & ReadFieldOptimizado(4, List4, Asc("-")))
Call SendData("/SUM " & ReadFieldOptimizado(1, List3, Asc("-")))
Call SendData("/SUM " & ReadFieldOptimizado(2, List3, Asc("-")))
Call SendData("/SUM " & ReadFieldOptimizado(3, List3, Asc("-")))
Call SendData("/SUM " & ReadFieldOptimizado(4, List3, Asc("-")))
End Sub

Private Sub Command14_Click()
Call SendData("/RMSG " & "Juegan la siguiente pelea: " & List4 & " Vs. " & List3)
End Sub

Private Sub Command16_Click()
Call SendData("/EXPLOTA " & ReadFieldOptimizado(1, List2, Asc("-")))
Call SendData("/EXPLOTA " & ReadFieldOptimizado(2, List2, Asc("-")))
Call SendData("/EXPLOTA " & ReadFieldOptimizado(3, List2, Asc("-")))
Call SendData("/EXPLOTA " & ReadFieldOptimizado(4, List2, Asc("-")))
Call SendData("/RMSG " & "Quedan eliminados del torneo: " & List2)
List2.RemoveItem List2.ListIndex
Label11 = List2.ListCount
End Sub

Private Sub Command17_Click()
If List4.ListIndex <> -1 Then
List1.AddItem List4
Call SendData("/PASAN " & ReadFieldOptimizado(1, List4, Asc("-")))
Call SendData("/PASAN " & ReadFieldOptimizado(2, List4, Asc("-")))
Call SendData("/PASAN " & ReadFieldOptimizado(3, List4, Asc("-")))
Call SendData("/PASAN " & ReadFieldOptimizado(4, List4, Asc("-")))
Call SendData("/RMSG " & "Pasan a la siguiente instancia: " & List4)
List4.RemoveItem List4.ListIndex
Label10 = List1.ListCount
Label13 = List4.ListCount
End If

End Sub

Private Sub Command18_Click()
If List1.ListIndex <> -1 Then
List4.AddItem List1
Call SendData("/PASAN " & ReadFieldOptimizado(1, List1, Asc("-")))
Call SendData("/PASAN " & ReadFieldOptimizado(2, List1, Asc("-")))
Call SendData("/PASAN " & ReadFieldOptimizado(3, List1, Asc("-")))
Call SendData("/PASAN " & ReadFieldOptimizado(4, List1, Asc("-")))
Call SendData("/RMSG " & "Pasan a la siguiente instancia: " & List1)
List1.RemoveItem List1.ListIndex
Label10 = List1.ListCount
Label13 = List4.ListCount
End If

End Sub

Private Sub Command19_Click()
Call SendData("/EXPLOTA " & ReadFieldOptimizado(1, List4, Asc("-")))
Call SendData("/EXPLOTA " & ReadFieldOptimizado(2, List4, Asc("-")))
Call SendData("/EXPLOTA " & ReadFieldOptimizado(3, List4, Asc("-")))
Call SendData("/EXPLOTA " & ReadFieldOptimizado(4, List4, Asc("-")))
Call SendData("/RMSG " & "Quedan eliminados del torneo: " & List4)
List4.RemoveItem List4.ListIndex
Label13 = List4.ListCount

End Sub

Private Sub Command2_Click()
List3.AddItem Combo4.Text & "-" & Combo5.Text & "-" & Combo6.Text & "-" & Combo7.Text
Label12 = List3.ListCount
End Sub

Private Sub Command20_Click()
Call SendData("/EXPLOTA " & ReadFieldOptimizado(1, List3, Asc("-")))
Call SendData("/EXPLOTA " & ReadFieldOptimizado(2, List3, Asc("-")))
Call SendData("/EXPLOTA " & ReadFieldOptimizado(3, List3, Asc("-")))
Call SendData("/EXPLOTA " & ReadFieldOptimizado(4, List3, Asc("-")))
Call SendData("/RMSG " & "Quedan eliminados del torneo: " & List3)
List3.RemoveItem List3.ListIndex
Label12 = List3.ListCount

End Sub

Private Sub Command21_Click()
If List3.ListIndex <> -1 Then
List2.AddItem List3
Call SendData("/PASAN " & ReadFieldOptimizado(1, List3, Asc("-")))
Call SendData("/PASAN " & ReadFieldOptimizado(2, List3, Asc("-")))
Call SendData("/PASAN " & ReadFieldOptimizado(3, List3, Asc("-")))
Call SendData("/PASAN " & ReadFieldOptimizado(4, List3, Asc("-")))
Call SendData("/RMSG " & "Pasan a la siguiente instancia: " & List3)
List3.RemoveItem List3.ListIndex
Label11 = List2.ListCount
Label12 = List3.ListCount
End If

End Sub

Private Sub Command22_Click()
If MsgBox("¿Está seguro que desea Guardar la Zona Nº1?", vbYesNo) = vbYes Then
Call GuardarLista(List1, "C:/list1.txt")
Call GuardarLista(List2, "C:/list2.txt")
End If
End Sub

Private Sub Command23_Click()
If MsgBox("¿Está seguro que desea cargar la Zona Nº1?", vbYesNo) = vbYes Then
Call LeerLista(List1, "C:/list1.txt")
Call LeerLista(List2, "C:/list2.txt")
End If
End Sub

Private Sub Command24_Click()
If MsgBox("¿Está seguro que desea guardar la Zona Nº2?", vbYesNo) = vbYes Then
Call GuardarLista(List4, "C:/list4.txt")
Call GuardarLista(List3, "C:/list3.txt")
End If
End Sub

Private Sub Command25_Click()
If MsgBox("¿Está seguro que desea cargar la Zona Nº2?", vbYesNo) = vbYes Then
Call LeerLista(List4, "C:/list4.txt")
Call LeerLista(List3, "C:/list3.txt")
End If
End Sub

Private Sub Command26_Click()
Call SendData("/RMSG " & "¡Esquinas! ¡Mucha Suerte! Comienza en...")
Call SendData("/cuenta " & Text5.Text)
End Sub

Private Sub Command27_Click()
List3.Clear
List4.Clear
Label13 = List3.ListCount
Label12 = List4.ListCount

End Sub

Private Sub Command28_Click()
List1.Clear
List2.Clear
Label10 = List1.ListCount
Label11 = List2.ListCount

End Sub

Private Sub Command29_Click()
Call GuardarLista(List1, "C:/list1.txt")
Call GuardarLista(List2, "C:/list2.txt")
Call GuardarLista(List3, "C:/list3.txt")
Call GuardarLista(List4, "C:/list4.txt")
Me.Hide
End Sub

Private Sub Command3_Click()
Call SendData("/RMSG " & "Juegan la siguiente pelea: " & List1 & " Vs. " & List2)

End Sub

Private Sub Command30_Click()
Call SendData("/RMSG " & "Primer Enfrentamiento: " & List1 & " Vs. " & List2)
End Sub


Private Sub Command31_Click()
Call SendData("/RMSG " & "1-0 Gana el team: " & List1)
End Sub

Private Sub Command32_Click()
Call SendData("/RMSG " & "2-0 Gana el team: " & List1)
End Sub

Private Sub Command33_Click()
Call SendData("/RMSG " & "2-1 Gana el team: " & List1)
End Sub

Private Sub Command34_Click()
Call SendData("/RMSG " & "3-1 Gana el team: " & List1)
End Sub

Private Sub Command35_Click()
Call SendData("/RMSG " & "1-1 Empatan team: " & List1)
End Sub

Private Sub Command36_Click()
Call SendData("/RMSG " & "2-2 Empatan team: " & List1)
End Sub



Private Sub Command37_Click()
Call SendData("/RMSG " & "1-0 Ganan el team: " & List2)
End Sub

Private Sub Command38_Click()
Call SendData("/RMSG " & "2-0 Ganan el team: " & List2)
End Sub

Private Sub Command39_Click()
Call SendData("/RMSG " & "2-1 Ganan el team: " & List2)
End Sub

Private Sub Command4_Click()
Call SendData("/SUM " & ReadFieldOptimizado(1, List1, Asc("-")))
Call SendData("/SUM " & ReadFieldOptimizado(2, List1, Asc("-")))
Call SendData("/SUM " & ReadFieldOptimizado(3, List1, Asc("-")))
Call SendData("/SUM " & ReadFieldOptimizado(4, List1, Asc("-")))
Call SendData("/SUM " & ReadFieldOptimizado(1, List2, Asc("-")))
Call SendData("/SUM " & ReadFieldOptimizado(2, List2, Asc("-")))
Call SendData("/SUM " & ReadFieldOptimizado(3, List2, Asc("-")))
Call SendData("/SUM " & ReadFieldOptimizado(4, List2, Asc("-")))
End Sub

Private Sub Command40_Click()
Call SendData("/RMSG " & "3-1 Ganan el team: " & List2)
End Sub

Private Sub Command41_Click()
Call SendData("/RMSG " & "1-1 Empatan team: " & List2)
End Sub

Private Sub Command42_Click()
Call SendData("/RMSG " & "2-2 Empatan team: " & List2)
End Sub

Private Sub Command43_Click()
Call SendData("/RMSG " & "3-1 Ganan el team: " & List4)
End Sub

Private Sub Command44_Click()
Call SendData("/RMSG " & "2-1 Ganan el team: " & List4)
End Sub



Private Sub Command45_Click()
Call SendData("/TORNEO " & Text6.Text)
End Sub

Private Sub Command46_Click()
Call SendData("/RMSG " & "3-2 Ganan el team: " & List2)
End Sub

Private Sub Command47_Click()
Call SendData("/RMSG " & "2-2 Empatan team: " & List3)
End Sub

Private Sub Command48_Click()
Call SendData("/RMSG " & "1-1 Empatan team: " & List3)
End Sub

Private Sub Command49_Click()
Call SendData("/RMSG " & "3-1 Ganan el team: " & List3)
End Sub

Private Sub Command5_Click()
If List2.ListIndex <> -1 Then
List3.AddItem List2
Call SendData("/PASAN " & ReadFieldOptimizado(1, List2, Asc("-")))
Call SendData("/PASAN " & ReadFieldOptimizado(2, List2, Asc("-")))
Call SendData("/PASAN " & ReadFieldOptimizado(3, List2, Asc("-")))
Call SendData("/PASAN " & ReadFieldOptimizado(4, List2, Asc("-")))
Call SendData("/RMSG " & "Pasan a la siguiente instancia: " & List2)
List2.RemoveItem List2.ListIndex
Label11 = List2.ListCount
Label12 = List3.ListCount
End If

End Sub

Private Sub Command50_Click()
Call SendData("/RMSG " & "2-1 Ganan el team: " & List3)
End Sub

Private Sub Command51_Click()
Call SendData("/RMSG " & "2-0 Ganan el team: " & List3)
End Sub

Private Sub Command52_Click()
Call SendData("/RMSG " & "1-0 Ganan el team: " & List3)
End Sub

Private Sub Command53_Click()
Call SendData("/RMSG " & "2-2 Empatan team: " & List4)
End Sub

Private Sub Command54_Click()
Call SendData("/RMSG " & "1-1 Empatan team: " & List4)
End Sub

Private Sub Command55_Click()
Call SendData("/RMSG " & "2-0 Ganan el team: " & List4)
End Sub

Private Sub Command56_Click()
Call SendData("/RMSG " & "1-0 Ganan el team: " & List4)
End Sub

Private Sub Command57_Click()
Call SendData("/RMSG " & "3-2 Gana el team: " & List1)
End Sub

Private Sub Command58_Click()
Call SendData("/RMSG " & "3-2 Ganan el team: " & List4)
End Sub

Private Sub Command59_Click()
Call SendData("/RMSG " & "3-2 Ganan el team: " & List3)
End Sub

Private Sub Command6_Click()
Call SendData("/EXPLOTA " & ReadFieldOptimizado(1, List1, Asc("-")))
Call SendData("/EXPLOTA " & ReadFieldOptimizado(2, List1, Asc("-")))
Call SendData("/EXPLOTA " & ReadFieldOptimizado(3, List1, Asc("-")))
Call SendData("/EXPLOTA " & ReadFieldOptimizado(4, List1, Asc("-")))
Call SendData("/RMSG " & "Quedan eliminados del torneo: " & List1)
List1.RemoveItem List1.ListIndex
Label10 = List1.ListCount

End Sub

Private Sub Command60_Click()
Call SendData("/RMSG " & "3-0 Ganan el team: " & List2)
End Sub

Private Sub Command61_Click()
Call SendData("/RMSG " & "3-0 Ganan el team: " & List3)
End Sub

Private Sub Command62_Click()
Call SendData("/RMSG " & "3-0 Ganan el team: " & List4)
End Sub

Private Sub Command63_Click()
Call SendData("/RMSG Reglas: No valdra items de canjeos, no vale repetir clase, no valen Hechizos como (invisibiidad, Resucitar), envia uno solo por pareja.")
End Sub

Private Sub Command64_Click()
Call SendData("/RMSG En instantes realizaré un torneo 2vs2, para " & Text8.Text & " parejas, comando /PARTICIPAR, " & Text7.Text & " Hora del servidor abro.")
End Sub

Private Sub Command65_Click()
Call SendData("/RMSG Repito: En instantes realizaré un torneo 2vs2, para " & Text8.Text & " parejas, comando /PARTICIPAR, " & Text7.Text & " Hora del servidor abro.")
End Sub

Private Sub Command66_Click()
Call SendData("/HORA")
End Sub

Private Sub Command67_Click()
Call SendData("/RMSG " & "3-0 Gana el team: " & List1)
End Sub

Private Sub Command68_Click()
Combo.Clear
Combo1.Clear
Combo2.Clear
Combo3.Clear
Combo4.Clear
Combo5.Clear
Combo6.Clear
Combo7.Clear
Combo8.Clear
Combo9.Clear
Combo10.Clear
Combo11.Clear
Combo12.Clear
Combo13.Clear
Combo14.Clear
Combo15.Clear
Call SendData("/PANELGM")
End Sub

Private Sub Command7_Click()
List1.AddItem Combo8.Text & "-" & Combo9.Text & "-" & Combo10.Text & "-" & Combo11.Text
Label10 = List1.ListCount
End Sub

Private Sub Command8_Click()
If List1.ListIndex <> -1 Then
List1.RemoveItem List1.ListIndex
Label10 = List1.ListCount

End If
End Sub

Private Sub Command9_Click()
List2.AddItem Combo12.Text & "-" & Combo13.Text & "-" & Combo14.Text & "-" & Combo15.Text
Label11 = List2.ListCount
End Sub

Private Sub cuatrovscuatro_Click()

End Sub

Private Sub Form_Load()
'List1.DragIcon = LoadPicture(App.Path & "\Graficos\Drag.ico")
'List2.DragIcon = LoadPicture(App.Path & "\Graficos\Drag.ico")
'List3.DragIcon = LoadPicture(App.Path & "\Graficos\Drag.ico")
'List4.DragIcon = LoadPicture(App.Path & "\Graficos\Drag.ico")
Call LeerLista(List1, "C:/list1.txt")
Call LeerLista(List2, "C:/list2.txt")
Call LeerLista(List3, "C:/list3.txt")
Call LeerLista(List4, "C:/list4.txt")
End Sub


Private Sub Label23_Click()
If List1.ListIndex <> -1 Then
List2.AddItem List1
List1.RemoveItem List1.ListIndex
Label10 = List1.ListCount
Label11 = List2.ListCount
End If
End Sub

Private Sub Label24_Click()
If List2.ListIndex <> -1 Then
   'Eliminamos el elemento que se encuentra seleccionado
   List1.AddItem List2
List2.RemoveItem List2.ListIndex
Label11 = List2.ListCount
Label10 = List1.ListCount

End If
End Sub

Private Sub Label25_Click()
If List3.ListIndex <> -1 Then
List4.AddItem List3
List3.RemoveItem List3.ListIndex
Label13 = List4.ListCount
Label12 = List3.ListCount
End If
End Sub

Private Sub Label26_Click()
If List4.ListIndex <> -1 Then
List3.AddItem List4
List4.RemoveItem List4.ListIndex
Label13 = List4.ListCount
Label12 = List3.ListCount
End If
End Sub

Private Sub List2_Click()
    SincListBox List2, List1
 '  Label11 = List2.ListCount
End Sub

Private Sub List1_Scroll()
    'Sincronizar también el primer item mostrado en la lista
   List2.TopIndex = List1.TopIndex
End Sub
Private Sub List2_Scroll()
    'Sincronizar también el primer item mostrado en la lista
    List1.TopIndex = List2.TopIndex
End Sub

Private Sub List1_Click()
    SincListBox List1, List2
   ' Label10 = List1.ListCount
End Sub
Private Sub QuitarListSelected(unList As Control)
    'Quitar los elementos seleccionados del listbox indicado
    'Parámetros:
    '   unList      el List a controlar
    '
    Dim i&
    
    With unList
        'Sólo hacer el bucle si permite multiselección
        If .MultiSelect Then
            For i = 0 To .ListCount - 1
                .Selected(i) = False
            Next
        End If
    End With
End Sub

Private Sub ListSelected(elListOrig As Control, elListDest As Control)
    'Marca en el ListDest los elementos seleccionados del ListOrig
    '
    'Los dos listbox deben tener el mismo número de elementos
    '
    Dim i&
    
    'Por si no tienen los mismos elementos
    On Local Error Resume Next
    
    With elListOrig
        For i = 0 To .ListCount - 1
            'Si el origen está seleccionado...
            If .Selected(i) Then
                elListDest.Selected(i) = .Selected(i)
            Else
                'sino, quitar la posible selección
                elListDest.Selected(i) = False
            End If
        Next
    End With
        
    Err = 0
End Sub

Private Sub PonerListSelected(elListOrig As Control, elListDest As Control)
    'Marca en el ListDest los elementos seleccionados del ListOrig
    '
    'Los dos listbox deben tener el mismo número de elementos
    '
    Dim i&
    
    'Por si no tienen los mismos elementos
    On Local Error Resume Next
    
    With elListOrig
        For i = 0 To .ListCount - 1
            elListDest.Selected(i) = .Selected(i)
        Next
    End With
        
    Err = 0
End Sub

Private Sub SincListBox(elListOrig As Control, elListDest As Control)
    Static EnListBox As Boolean
        
    'Sincronizar el elListDest con el elListOrig
    If Not EnListBox Then
    
        EnListBox = True
        
'        'Desmarcar los elementos seleccionados
'        QuitarListSelected elListDest
'
'        'Marcar en el 1º ListBox los seleccionados del 2º
'        PonerListSelected elListOrig, elListDest
        
        'Poner en el ListDest los mismos que en ListOrig
        ListSelected elListOrig, elListDest
        
        'Posicionar el elemento superior
     '   elListDest.TopIndex = elListOrig.TopIndex
        
        EnListBox = False
    End If
End Sub

Private Sub List3_Click()
SincListBox List3, List4
'Label12 = List3.ListCount
End Sub

Private Sub List4_Click()
SincListBox List4, List3
'Label13 = List4.ListCount
End Sub

Sub GuardarLista(listax As ListBox, Donde As String)
Dim fnum As Integer
On Error GoTo Ninguno
    fnum = FreeFile
    Open Donde For Output As fnum
    
    Dim i As Integer
    For i = 0 To listax.ListCount
            Print #fnum, listax.List(i)
        DoEvents
    Next i
    
    Close fnum
   ' MsgBox "Torneo Guardado."
Ninguno:
End Sub

Sub LeerLista(listax As ListBox, Donde As String)
Dim fnum As Integer
Dim txt As String
On Error GoTo Ninguno

fnum = FreeFile
    Open Donde For Input As fnum
    Do While Not EOF(fnum)
        Line Input #fnum, txt
        listax.AddItem txt
        'Texto.Text = Texto.Text & vbCrLf & txt
    Loop
    Close fnum
    MsgBox "Torneo Cargado."
    Label10 = List1.ListCount
    Label11 = List2.ListCount
    Label13 = List4.ListCount
    Label12 = List3.ListCount
Ninguno:
End Sub

Private Sub List1_DragDrop(Source As Control, X As Single, Y As Single)
    
    ' Si el control es el List2 entonces OK ..
    If Source Is List4 Then
       If List4.ListIndex <> -1 Then
            List1.AddItem List4.List(List4.ListIndex)
            List4.RemoveItem List4.ListIndex
            Label10 = List1.ListCount
            Label13 = List4.ListCount
       End If
    End If
End Sub


' Inicia la operación de arrastre, es decir el drag para List1
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, _
                            X As Single, Y As Single)
    List1.Drag vbBeginDrag
End Sub

' Al soltar el item en List2, se agrega al mismo y se elimina el del List1

Private Sub List4_DragDrop(Source As Control, _
                           X As Single, Y As Single)
    
    ' Si el control es el List1 entonces..
    If Source Is List1 Then
        If List1.ListIndex <> -1 Then
           List4.AddItem List1.List(List1.ListIndex)
           List1.RemoveItem List1.ListIndex
           Label10 = List1.ListCount
           Label13 = List4.ListCount
        End If
    End If
End Sub

' Comienza el Drag para el List2
Private Sub List4_MouseDown(Button As Integer, Shift As Integer, _
                                        X As Single, Y As Single)
    List4.Drag vbBeginDrag
End Sub

'#####################################################################
'####################### AHORA LIST2 y LIST3 #########################
'#####################################################################

Private Sub List2_DragDrop(Source As Control, X As Single, Y As Single)
    
    ' Si el control es el List2 entonces OK ..
    If Source Is List3 Then
       If List3.ListIndex <> -1 Then
            List2.AddItem List3.List(List3.ListIndex)
            List3.RemoveItem List3.ListIndex
            Label11 = List2.ListCount
            Label12 = List3.ListCount
       End If
    End If
End Sub


' Inicia la operación de arrastre, es decir el drag para List1
Private Sub List2_MouseDown(Button As Integer, Shift As Integer, _
                            X As Single, Y As Single)
    List2.Drag vbBeginDrag
End Sub

' Al soltar el item en List2, se agrega al mismo y se elimina el del List1

Private Sub List3_DragDrop(Source As Control, _
                           X As Single, Y As Single)
    
    ' Si el control es el List1 entonces..
    If Source Is List2 Then
        If List2.ListIndex <> -1 Then
           List3.AddItem List2.List(List2.ListIndex)
           List2.RemoveItem List2.ListIndex
           Label11 = List2.ListCount
           Label12 = List3.ListCount
        End If
    End If
End Sub

' Comienza el Drag para el List2
Private Sub List3_MouseDown(Button As Integer, Shift As Integer, _
                                        X As Single, Y As Single)
    List3.Drag vbBeginDrag
End Sub




Private Sub msg1_Click()
Call SendData("/RMSG " & "¡Una extraordinaria pelea, ninguno de los equipos logra sacarse ventaja!")
End Sub

Private Sub msg2_Click()
Call SendData("/RMSG " & "¡Esta pelea deja mucho que hablar, ambos equipos están dando un gran espectáculo!")
End Sub

Private Sub msg3_Click()
Call SendData("/RMSG " & "Una pelea digna de ver, unos de los mejores enfrentamientos de este evento…")
End Sub

Private Sub msg4_Click()
Call SendData("/RMSG " & "Remos, inmos, y mas inmos… ¡Que pelea muchachos!")
End Sub

Private Sub msg5_Click()
Call SendData("/RMSG " & "¡Esquinas! ¡Mucha Suerte! Comienza en...")
End Sub

Private Sub msg6_Click()
Call SendData("/RMSG " & "Inmos, remos, apocas, una gran batalla, aun no se ha visto lo mejor!")
End Sub

Private Sub potas_Click()
Call SendData("/ACC 14")
End Sub

Private Sub sacer_Click()
Call SendData("/ACC 5")
End Sub

Private Sub torneo_Click()

End Sub

Private Sub tresvstres_Click()

End Sub

Private Sub Text2_Change()

End Sub
