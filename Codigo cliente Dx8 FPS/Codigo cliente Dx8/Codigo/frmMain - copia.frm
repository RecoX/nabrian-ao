VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{B370EF78-425C-11D1-9A28-004033CA9316}#2.0#0"; "Captura.ocx"
Begin VB.Form frmPrincipal 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   720
      Top             =   3000
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   "FlamiusAO"
      HostName        =   "FlamiusAO"
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   10200
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   10200
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer CierraL 
      Interval        =   2500
      Left            =   2160
      Top             =   3000
   End
   Begin VB.Timer Perdedor 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6720
      Top             =   7200
   End
   Begin VB.Timer Ganador 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6240
      Top             =   7200
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   106
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1860
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7080
      TabIndex        =   105
      Text            =   "Text1"
      Top             =   10080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer DetectedCheats 
      Interval        =   300
      Left            =   1200
      Top             =   3000
   End
   Begin VB.PictureBox Minimap 
      AutoRedraw      =   -1  'True
      Height          =   1425
      Left            =   6840
      ScaleHeight     =   91
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   100
      Top             =   360
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Timer FPSTIMER 
      Interval        =   10000
      Left            =   2640
      Top             =   3000
   End
   Begin VB.Timer AntiExternos 
      Interval        =   6000
      Left            =   1680
      Top             =   3000
   End
   Begin VB.Frame frInvent 
      BorderStyle     =   0  'None
      Height          =   3885
      Left            =   8520
      TabIndex        =   36
      Top             =   2040
      Width           =   3270
      Begin VB.Image Shape2 
         Height          =   480
         Left            =   480
         Picture         =   "frmMain.frx":1642
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   480
         Left            =   3360
         Top             =   3000
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   10
         Left            =   2400
         TabIndex        =   87
         Top             =   1260
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   9
         Left            =   1920
         TabIndex        =   86
         Top             =   1305
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   25
         Left            =   2400
         TabIndex        =   85
         Top             =   2880
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   24
         Left            =   1920
         TabIndex        =   84
         Top             =   2865
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   23
         Left            =   1440
         TabIndex        =   83
         Top             =   2880
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   22
         Left            =   960
         TabIndex        =   82
         Top             =   2880
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   21
         Left            =   480
         TabIndex        =   81
         Top             =   2880
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   20
         Left            =   2400
         TabIndex        =   80
         Top             =   2340
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   19
         Left            =   1920
         TabIndex        =   79
         Top             =   2340
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   18
         Left            =   1440
         TabIndex        =   78
         Top             =   2340
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   17
         Left            =   960
         TabIndex        =   77
         Top             =   2340
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   16
         Left            =   480
         TabIndex        =   76
         Top             =   2340
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   15
         Left            =   2400
         TabIndex        =   75
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   14
         Left            =   1920
         TabIndex        =   74
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   13
         Left            =   1440
         TabIndex        =   73
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   12
         Left            =   960
         TabIndex        =   72
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   11
         Left            =   480
         TabIndex        =   71
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   7
         Left            =   960
         TabIndex        =   70
         Top             =   1260
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   6
         Left            =   480
         TabIndex        =   69
         Top             =   1260
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   5
         Left            =   2400
         TabIndex        =   68
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   4
         Left            =   1920
         TabIndex        =   67
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   960
         TabIndex        =   66
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   8
         Left            =   1440
         TabIndex        =   65
         Top             =   1260
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   64
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   1440
         TabIndex        =   63
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   62
         Top             =   1080
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   61
         Top             =   1080
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   4
         Left            =   2280
         TabIndex        =   60
         Top             =   1080
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   5
         Left            =   2760
         TabIndex        =   59
         Top             =   1080
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   6
         Left            =   840
         TabIndex        =   58
         Top             =   1620
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   7
         Left            =   1320
         TabIndex        =   57
         Top             =   1620
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   8
         Left            =   1800
         TabIndex        =   56
         Top             =   1620
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   9
         Left            =   2280
         TabIndex        =   55
         Top             =   1620
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   10
         Left            =   2760
         TabIndex        =   54
         Top             =   1620
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   11
         Left            =   840
         TabIndex        =   53
         Top             =   2160
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   12
         Left            =   1320
         TabIndex        =   52
         Top             =   2160
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   13
         Left            =   1800
         TabIndex        =   51
         Top             =   2160
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   15
         Left            =   2760
         TabIndex        =   50
         Top             =   2160
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   20
         Left            =   2760
         TabIndex        =   49
         Top             =   2700
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   19
         Left            =   2280
         TabIndex        =   48
         Top             =   2700
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   18
         Left            =   1800
         TabIndex        =   47
         Top             =   2700
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   17
         Left            =   1320
         TabIndex        =   46
         Top             =   2700
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   16
         Left            =   840
         TabIndex        =   45
         Top             =   2700
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   21
         Left            =   840
         TabIndex        =   44
         Top             =   3240
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   22
         Left            =   1320
         TabIndex        =   43
         Top             =   3240
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   23
         Left            =   1800
         TabIndex        =   42
         Top             =   3240
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   24
         Left            =   2280
         TabIndex        =   41
         Top             =   3240
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   25
         Left            =   2760
         TabIndex        =   40
         Top             =   3240
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   3
         Left            =   1800
         TabIndex        =   39
         Top             =   1080
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   14
         Left            =   2280
         TabIndex        =   38
         Top             =   2115
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   25
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   2910
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   24
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   2910
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   23
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   2910
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   22
         Left            =   960
         Stretch         =   -1  'True
         Top             =   2910
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   21
         Left            =   480
         Stretch         =   -1  'True
         Top             =   2910
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   20
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   2370
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   19
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   2370
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   18
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   2370
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   17
         Left            =   960
         Stretch         =   -1  'True
         Top             =   2370
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   16
         Left            =   480
         Stretch         =   -1  'True
         Top             =   2370
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   15
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   1830
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   14
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   1830
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   13
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   1830
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   12
         Left            =   960
         Stretch         =   -1  'True
         Top             =   1830
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   11
         Left            =   480
         Stretch         =   -1  'True
         Top             =   1830
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   10
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   1290
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   9
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   1290
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   8
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   1290
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   7
         Left            =   960
         Stretch         =   -1  'True
         Top             =   1290
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   6
         Left            =   480
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   5
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   750
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   4
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   750
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   2
         Left            =   960
         Stretch         =   -1  'True
         Top             =   750
         Width           =   480
      End
      Begin VB.Label lblHechizos 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         MouseIcon       =   "frmMain.frx":1786
         MousePointer    =   99  'Custom
         TabIndex        =   37
         Top             =   240
         Width           =   1200
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   1
         Left            =   480
         Stretch         =   -1  'True
         Top             =   750
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Height          =   480
         Index           =   3
         Left            =   1440
         Top             =   750
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   135
         Index           =   0
         Left            =   960
         MouseIcon       =   "frmMain.frx":1A90
         MousePointer    =   99  'Custom
         Top             =   3600
         Width           =   195
      End
      Begin VB.Image Image5 
         Height          =   135
         Index           =   1
         Left            =   1320
         MouseIcon       =   "frmMain.frx":1D9A
         MousePointer    =   99  'Custom
         Top             =   3600
         Width           =   195
      End
      Begin VB.Image Image5 
         Height          =   195
         Index           =   2
         Left            =   240
         MouseIcon       =   "frmMain.frx":20A4
         MousePointer    =   99  'Custom
         Top             =   3600
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   195
         Index           =   3
         Left            =   600
         MouseIcon       =   "frmMain.frx":23AE
         MousePointer    =   99  'Custom
         Top             =   3600
         Width           =   255
      End
      Begin VB.Image imgFondoInvent 
         Height          =   3930
         Left            =   0
         Top             =   0
         Width           =   3270
      End
   End
   Begin VB.Frame frHechizos 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   8520
      TabIndex        =   28
      Top             =   2040
      Width           =   3240
      Begin VB.ListBox lstHechizos 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2430
         Left            =   360
         TabIndex        =   29
         Top             =   840
         Width           =   2595
      End
      Begin VB.Label lblArriba 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         MouseIcon       =   "frmMain.frx":26B8
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   600
         Width           =   180
      End
      Begin VB.Label lblAbajo 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         MouseIcon       =   "frmMain.frx":29C2
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   600
         Width           =   180
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Height          =   480
         Left            =   2040
         MouseIcon       =   "frmMain.frx":2CCC
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   3360
         Width           =   1050
      End
      Begin VB.Label lblInvent 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         MouseIcon       =   "frmMain.frx":2FD6
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label lblCh 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   120
         TabIndex        =   31
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label lblLanzar 
         BackStyle       =   0  'Transparent
         Height          =   480
         Left            =   120
         MouseIcon       =   "frmMain.frx":32E0
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   3360
         Width           =   1545
      End
      Begin VB.Image imgFondoHechizos 
         Height          =   3930
         Left            =   0
         Top             =   0
         Width           =   3270
      End
   End
   Begin VB.Timer TIMERQUECARAJO 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   5760
      Top             =   7200
   End
   Begin Captura.wndCaptura Captura1 
      Left            =   1200
      Top             =   3480
      _ExtentX        =   688
      _ExtentY        =   688
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   8640
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   23
      Top             =   9840
      Visible         =   0   'False
      Width           =   975
   End
   Begin RichTextLib.RichTextBox rectxt 
      Height          =   1320
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   405
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   2328
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":35EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3240
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
   End
   Begin VB.PictureBox renderer 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   6240
      Left            =   150
      ScaleHeight     =   416
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   541
      TabIndex        =   108
      Top             =   2220
      Width           =   8115
   End
   Begin VB.Label Moverpantalla 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   109
      Top             =   0
      Width           =   6735
   End
   Begin VB.Image Image8 
      Height          =   375
      Left            =   10920
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label MB 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   8400
      TabIndex        =   107
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(100%)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9900
      TabIndex        =   104
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label NumCanjesD 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10440
      TabIndex        =   103
      Top             =   8355
      Width           =   615
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimap"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6840
      TabIndex        =   102
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimap"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6840
      TabIndex        =   101
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "C. Clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   99
      Top             =   1875
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "ULLA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   98
      Top             =   1875
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Arghelin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   97
      Top             =   1875
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Isla pirata"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   96
      Top             =   1875
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Zona espera "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   95
      Top             =   1875
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Sala GM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   94
      Top             =   1875
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Panel GM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6315
      TabIndex        =   93
      Top             =   1875
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "SOPORTES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   7290
      TabIndex        =   92
      Top             =   1875
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   10200
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label lblletra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   91
      Top             =   1800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label cantidadmana 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9720
      TabIndex        =   90
      Top             =   6585
      Width           =   1650
   End
   Begin VB.Label cantidadhp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9720
      TabIndex        =   89
      Top             =   6990
      Width           =   1650
   End
   Begin VB.Image Hpshp 
      Height          =   195
      Left            =   9210
      Picture         =   "frmMain.frx":3668
      Top             =   6990
      Width           =   2700
   End
   Begin VB.Image MANShp 
      Height          =   195
      Left            =   9210
      Picture         =   "frmMain.frx":37AC
      Top             =   6585
      Width           =   2700
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "FPS:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7680
      TabIndex        =   88
      Top             =   120
      Width           =   375
   End
   Begin VB.Image barrita 
      Height          =   300
      Left            =   8580
      Picture         =   "frmMain.frx":38F0
      Top             =   1305
      Width           =   3180
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   8640
      TabIndex        =   27
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label NumCanjes 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10440
      TabIndex        =   26
      Top             =   8130
      Width           =   615
   End
   Begin VB.Label NumFrags 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   7320
      TabIndex        =   25
      Top             =   8715
      Width           =   615
   End
   Begin VB.Image ImgMen 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   8640
      MouseIcon       =   "frmMain.frx":6324
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":6FEE
      Top             =   6075
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSoporte 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   8640
      Picture         =   "frmMain.frx":9173
      Top             =   6075
      Width           =   480
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6960
      TabIndex        =   24
      Top             =   9720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   22
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblNivel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "45"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   9240
      TabIndex        =   21
      Top             =   990
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   3
      Left            =   0
      MouseIcon       =   "frmMain.frx":B784
      MousePointer    =   99  'Custom
      Top             =   9000
      Width           =   45
   End
   Begin VB.Image Party 
      Height          =   135
      Left            =   9840
      MouseIcon       =   "frmMain.frx":BA8E
      MousePointer    =   99  'Custom
      Top             =   9840
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label NumOnline 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   210
      Left            =   5850
      TabIndex        =   20
      Top             =   8715
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   11160
      TabIndex        =   19
      Top             =   1020
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   11520
      TabIndex        =   18
      Top             =   1020
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6840
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10800
      TabIndex        =   16
      Top             =   1020
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label modo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "1 Normal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Agilidad 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   9075
      TabIndex        =   14
      Top             =   7620
      Width           =   225
   End
   Begin VB.Label Fuerza 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   9750
      TabIndex        =   13
      Top             =   7620
      Width           =   225
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   840
      Top             =   0
      Width           =   7455
   End
   Begin VB.Label casco 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3660
      TabIndex        =   0
      Top             =   8730
      Width           =   540
   End
   Begin VB.Label armadura 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   555
      TabIndex        =   11
      Top             =   8730
      Width           =   540
   End
   Begin VB.Label escudo 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   8730
      Width           =   540
   End
   Begin VB.Label arma 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2730
      TabIndex        =   9
      Top             =   8730
      Width           =   525
   End
   Begin VB.Label mapa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ullathorpe"
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
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   8700
      Width           =   3495
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   8160
      Top             =   9840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label cantidadagua 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   9960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label cantidadsta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12600
      TabIndex        =   7
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label cantidadhambre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   7920
      TabIndex        =   5
      Top             =   9840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   11280
      MouseIcon       =   "frmMain.frx":BD98
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00003E25&
      X1              =   16
      X2              =   551.467
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   11640
      MouseIcon       =   "frmMain.frx":C0A2
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   495
   End
   Begin VB.Label fpstext 
      BackStyle       =   0  'Transparent
      Caption         =   "84"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   8040
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Neliam"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8640
      TabIndex        =   3
      Top             =   570
      Width           =   3105
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H00008080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6600
      Top             =   9720
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   75
      Left            =   9480
      TabIndex        =   2
      Top             =   10080
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   120
      Left            =   9600
      Top             =   9960
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   240
      Left            =   8880
      Top             =   9960
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7920
      TabIndex        =   1
      Top             =   10080
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LabelVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V 1.0.0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   110
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NabrianAO (www.nabrianao.net)
'director del proyecto: #Esteban(Neliam)

'servidor basado en fénixao 1.0
'medios de contacto:
'Skype: dc.esteban
'E-mail: nabrianao@gmail.com
Option Explicit
Dim StartX, StartY
Dim h As Integer

Private Type POINTAPI
    X As Long
    Y As Long
End Type
 
Private Declare Function GetClassName Lib "user32" Alias _
 "GetClassNameA" ( _
 ByVal hwnd As Long, _
 ByVal lpGetClassNameA As String, _
 ByVal nMaxCount As Long) As Long
 
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As _
Long, ByVal yPoint As Long) As Long
Dim Mouse As POINTAPI

Private Type BLENDFUNCTION
BlendOp As Byte
BlendFlags As Byte
SourceConstantAlpha As Byte
AlphaFormat As Byte
End Type
Private Const AC_SRC_OVER = &H0
   
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal xOriginDest As Long, ByVal yOriginDest As Long, ByVal WidthDest As Long, ByVal HeightDest As Long, ByVal hDCsrc As Long, ByVal xOriginSrc As Long, ByVal yOriginSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
   
Dim Blend As BLENDFUNCTION
Dim blendlong As Long
Dim Contador As Integer

Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long

Dim POS(0) As DSBPOSITIONNOTIFY
Public IsPlaying As Byte
Public boton As Integer

Dim endEvent As Long




Private Sub AntiExternos_Timer()
If logged Then
If FindWindow(vbNullString, UCase$("Cheat Engine 5.1.1")) Then
    Call AoDefCheatDetect("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("AutoClick 2.2")) Then
    Call AoDefCheatDetect("AutoClick")
ElseIf FindWindow(vbNullString, UCase$("ART-MONEY")) Then
    Call AoDefCheatDetect("Art-Money")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 5.0")) Then
    Call AoDefCheatDetect("Cheat Engine 5.0")
ElseIf FindWindow(vbNullString, UCase$("CROWN MAKRO")) Then
    Call AoDefCheatDetect("Crown Makro")
ElseIf FindWindow(vbNullString, UCase$("A TRABAJAR...")) Then
    Call AoDefCheatDetect("Macro")
ElseIf FindWindow(vbNullString, UCase$("ews")) Then
    Call AoDefCheatDetect("Macro")
ElseIf FindWindow(vbNullString, UCase$("Pts")) Then
    Call AoDefCheatDetect("Auto Potas")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 5.2")) Then
    Call AoDefCheatDetect("Cheat Engine 5.2")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 5.6")) Then
    Call AoDefCheatDetect("Cheat Engine 5.6")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 5.7")) Then
    Call AoDefCheatDetect("Cheat Engine 5.7")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 5.9")) Then
    Call AoDefCheatDetect("Cheat Engine 5.9")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 6.0")) Then
    Call AoDefCheatDetect("Cheat Engine 6.0")
ElseIf FindWindow(vbNullString, UCase$("SOLOCOVO?")) Then
    Call AoDefCheatDetect("SOLOCOVO?")
ElseIf FindWindow(vbNullString, UCase$("MACROCRACK <GONZA_VI@HOTMAIL.COM>")) Then
    Call AoDefCheatDetect("MACROCRACK <GONZA_VI@HOTMAIL.COM>")
ElseIf FindWindow(vbNullString, UCase$("MACROCRACK <GONZA_VJ@HOTMAIL.COM>")) Then
    Call AoDefCheatDetect("MACROCRACK <GONZA_VJ@HOTMAIL.COM>")
ElseIf FindWindow(vbNullString, UCase$("MACRO CRACK <GONZA_VI@HOTMAIL.COM>")) Then
    Call AoDefCheatDetect("MACRO CRACK <GONZA_VI@HOTMAIL.COM>")
ElseIf FindWindow(vbNullString, UCase$("MACRO CRACK <GONZA_VJ@HOTMAIL.COM>")) Then
    Call AoDefCheatDetect("MACRO CRACK <GONZA_VJ@HOTMAIL.COM>")
ElseIf FindWindow(vbNullString, UCase$("CHITS")) Then
    Call AoDefCheatDetect("CHITS")
ElseIf FindWindow(vbNullString, UCase$("ORKAM")) Then
    Call AoDefCheatDetect("ORKAM")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V5.4")) Then
    Call AoDefCheatDetect("Cheat Engine V5.4")
ElseIf FindWindow(vbNullString, UCase$("Countach")) Then
    Call AoDefCheatDetect("Countach")
ElseIf FindWindow(vbNullString, UCase$("MacroRecorder")) Then
    Call AoDefCheatDetect("MacroRecorder")
ElseIf FindWindow(vbNullString, UCase$("Ultimatemacros")) Then
    Call AoDefCheatDetect("Ultimatemacros")
ElseIf FindWindow(vbNullString, UCase$("MacroLauncher")) Then
    Call AoDefCheatDetect("MacroLauncher")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 5.5")) Then
    Call AoDefCheatDetect("Cheat Engine 5.5")
ElseIf FindWindow(vbNullString, UCase$("Auto Remo- TheFrank^")) Then
    Call AoDefCheatDetect("Auto Remo- TheFrank^")
ElseIf FindWindow(vbNullString, UCase$("WPE PRO")) Then
    Call AoDefCheatDetect("WPE PRO")
ElseIf FindWindow(vbNullString, UCase$("WPE PRO - " & AoDefOriginalClientName & ".exe")) Then
    Call AoDefCheatDetect("WPE PRO")
ElseIf FindWindow(vbNullString, UCase$("WPE PRO - [WPEPRO2]")) Then
    Call AoDefCheatDetect("WPE PRO")
ElseIf FindWindow(vbNullString, UCase$("WPE PRO [WPEPRO2]")) Then
    Call AoDefCheatDetect("WPE PRO")
ElseIf FindWindow(vbNullString, UCase$("WPE PRO - " & AoDefOriginalClientName & ".exe" & " - [WPEPRO2]")) Then
    Call AoDefCheatDetect("WPE PRO")
ElseIf FindWindow(vbNullString, UCase$("rPE - rEdoX Packet Editor")) Then
    Call AoDefCheatDetect("rPE - rEdoX Packet Editor")
ElseIf FindWindow(vbNullString, UCase$("MACRO FOWL")) Then
    Call AoDefCheatDetect("MACRO FOWL")
ElseIf FindWindow(vbNullString, UCase$("MINI MACRO BY FOWL WWW.XTREME-ZONE.NET")) Then
    Call AoDefCheatDetect("MINI MACRO BY FOWL WWW.XTREME-ZONE.NET")
ElseIf FindWindow(vbNullString, UCase$("MACROSARAZA")) Then
    Call AoDefCheatDetect("MACROSARAZA")
ElseIf FindWindow(vbNullString, UCase$("Macroncmurd")) Then
    Call AoDefCheatDetect("Macroncmurd")
ElseIf FindWindow(vbNullString, UCase$("AUTOTRAINING")) Then
    Call AoDefCheatDetect("AUTOTRAINING")
ElseIf FindWindow(vbNullString, UCase$("0RK4M Version 1.5")) Then
    Call AoDefCheatDetect("0RK4M Version 1.5")
ElseIf FindWindow(vbNullString, UCase$("cmd")) Then
    Call AoDefCheatDetect("cmd")
ElseIf FindWindow(vbNullString, UCase$("X-Z MULTIMACRO VERSION II BY THEGABYX WWW.XTREME-ZONE.NET")) Then
    Call AoDefCheatDetect("X-Z MULTIMACRO VERSION II BY THEGABYX WWW.XTREME-ZONE.NET")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 6.0")) Then
    Call AoDefCheatDetect("Cheat Engine 6.0")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 6.1")) Then
    Call AoDefCheatDetect("Cheat Engine 6.1")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 5.4")) Then
    Call AoDefCheatDetect("Cheat Engine 5.4")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 5.5")) Then
    Call AoDefCheatDetect("Cheat Engine 5.5")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 5.8")) Then
    Call AoDefCheatDetect("Cheat Engine 5.8")
ElseIf FindWindow(vbNullString, UCase$("SoLocoVo?")) Then
    Call AoDefCheatDetect("SOLOCOVO?")
ElseIf FindWindow(vbNullString, UCase$("-=[ANUBYS RADAR]=-")) Then
    Call AoDefCheatDetect("-=[ANUBYS RADAR]=-")
ElseIf FindWindow(vbNullString, UCase$("CRAZY SPEEDER 1.05")) Then
    Call AoDefCheatDetect("CRAZY SPEEDER 1.05")
ElseIf FindWindow(vbNullString, UCase$("SET !XSPEED.NET")) Then
    Call AoDefCheatDetect("SET !XSPEED.NET")
ElseIf FindWindow(vbNullString, UCase$("SPEEDERXP V1.80 - UNREGISTERED")) Then
    Call AoDefCheatDetect("SPEEDERXP V1.80 - UNREGISTERED")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 5.3")) Then
    Call AoDefCheatDetect("Cheat Engine 5.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 5.1")) Then
    Call AoDefCheatDetect("Cheat Engine 5.1")
ElseIf FindWindow(vbNullString, UCase$("A SPEEDER")) Then
    Call AoDefCheatDetect("A SPEEDER")
ElseIf FindWindow(vbNullString, UCase$("MEMO :P")) Then
    Call AoDefCheatDetect("MEMO :P")
ElseIf FindWindow(vbNullString, UCase$("ORK4M VERSION 1.5")) Then
    Call AoDefCheatDetect("ORK4M VERSION 1.5")
ElseIf FindWindow(vbNullString, UCase$("BY FEDEX")) Then
    Call AoDefCheatDetect("By Fedex")
ElseIf FindWindow(vbNullString, UCase$("!XSPEED.NET +4.59")) Then
    Call AoDefCheatDetect("!Xspeeder")
ElseIf FindWindow(vbNullString, UCase$("CAMBIA TITULOS DE CHEATS BY FEDEX")) Then
    Call AoDefCheatDetect("Cambia titulos")
ElseIf FindWindow(vbNullString, UCase$("NEWENG OCULTO")) Then
    Call AoDefCheatDetect("NEWENG OCULTO")
ElseIf FindWindow(vbNullString, UCase$("SERBIO ENGINE")) Then
    Call AoDefCheatDetect("SERBIO ENGINE")
ElseIf FindWindow(vbNullString, UCase$("REYMIX ENGINE 5.3 PUBLIC")) Then
    Call AoDefCheatDetect("REYMIX ENGINE 5.3 PUBLIC")
ElseIf FindWindow(vbNullString, UCase$("REY ENGINE 5.2")) Then
    Call AoDefCheatDetect("REY ENGINE 5.2")
ElseIf FindWindow(vbNullString, UCase$("AUTOCLICK - BY NIO_SHOOTER")) Then
    Call AoDefCheatDetect("AUTOCLICK - BY NIO_SHOOTER")
ElseIf FindWindow(vbNullString, UCase$("TONNER MINER! :D [REG][SKLOV] 2.0")) Then
    Call AoDefCheatDetect("TONNER MINER! :D [REG][SKLOV] 2.0")
ElseIf FindWindow(vbNullString, UCase$("Buffy The vamp Slayer")) Then
    Call AoDefCheatDetect("Buffy The vamp Slayer")
ElseIf FindWindow(vbNullString, UCase$("Blorb Slayer 1.12.552 (BETA)")) Then
    Call AoDefCheatDetect("Blorb Slayer 1.12.552 (BETA)")
ElseIf FindWindow(vbNullString, UCase$("PumaEngine3.0")) Then
    Call AoDefCheatDetect("PumaEngine3.0")
ElseIf FindWindow(vbNullString, UCase$("Vicious Engine 5.0")) Then
    Call AoDefCheatDetect("Vicious Engine 5.0")
ElseIf FindWindow(vbNullString, UCase$("AkumaEngine33")) Then
    Call AoDefCheatDetect("AkumaEngine33")
ElseIf FindWindow(vbNullString, UCase$("Spuc3ngine")) Then
    Call AoDefCheatDetect("Spuc3ngine")
ElseIf FindWindow(vbNullString, UCase$("Ultra Engine")) Then
    Call AoDefCheatDetect("Ultra Engine")
ElseIf FindWindow(vbNullString, UCase$("Engine")) Then
    Call AoDefCheatDetect("Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V5.6")) Then
    Call AoDefCheatDetect("Cheat Engine V5.6")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V5.5")) Then
    Call AoDefCheatDetect("Cheat Engine V5.5")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4")) Then
    Call AoDefCheatDetect("Cheat Engine V4.4")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4 German Add-On")) Then
    Call AoDefCheatDetect("Cheat Engine V4.4 German Add-On")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.3")) Then
    Call AoDefCheatDetect("Cheat Engine V4.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.2")) Then
    Call AoDefCheatDetect("Cheat Engine V4.2")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.1.1")) Then
    Call AoDefCheatDetect("Cheat Engine V4.1.1")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.3")) Then
    Call AoDefCheatDetect("Cheat Engine V3.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.2")) Then
    Call AoDefCheatDetect("Cheat Engine V3.2")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.1")) Then
    Call AoDefCheatDetect("Cheat Engine V3.1")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine")) Then
    Call AoDefCheatDetect("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("danza engine 5.2.150")) Then
    Call AoDefCheatDetect("danza engine 5.2.150")
ElseIf FindWindow(vbNullString, UCase$("zenx engine")) Then
    Call AoDefCheatDetect("zenx engine")
ElseIf FindWindow(vbNullString, UCase$("MACROMAKER")) Then
    Call AoDefCheatDetect("Macro Maker")
ElseIf FindWindow(vbNullString, UCase$("MACREOMAKER - EDIT MACRO")) Then
    Call AoDefCheatDetect("MACREOMAKER - EDIT MACRO")
ElseIf FindWindow(vbNullString, UCase$("By Fedex")) Then
    Call AoDefCheatDetect("Macro Fedex")
ElseIf FindWindow(vbNullString, UCase$("Macro Mage 1.0")) Then
    Call AoDefCheatDetect("Macro Mage")
ElseIf FindWindow(vbNullString, UCase$("Auto* v0.4 (c) 2001 Pete Powa")) Then
    Call AoDefCheatDetect("Macro Fisher")
ElseIf FindWindow(vbNullString, UCase$("Kizsada")) Then
    Call AoDefCheatDetect("Macro K33")
ElseIf FindWindow(vbNullString, UCase$("Makro K33")) Then
    Call AoDefCheatDetect("Macro K33")
ElseIf FindWindow(vbNullString, UCase$("Super Saiyan")) Then
    Call AoDefCheatDetect("El Chit del Geri")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete")) Then
    Call AoDefCheatDetect("Piringulete")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete 2003")) Then
    Call AoDefCheatDetect("Piringulete 2003")
ElseIf FindWindow(vbNullString, UCase$("TUKY2005")) Then
    Call AoDefCheatDetect("Makro Tuky")
ElseIf FindWindow(vbNullString, UCase$("Volks")) Then
    Call AoDefCheatDetect("TURBINAS")
ElseIf FindWindow(vbNullString, UCase$("Turbinas")) Then
    Call AoDefCheatDetect("TURBINAS")
ElseIf FindWindow(vbNullString, UCase$("msn")) Then
    Call AoDefCheatDetect("msn")
ElseIf FindWindow(vbNullString, UCase$("Volks")) Then
    Call AoDefCheatDetect("TURBINAS")
ElseIf FindWindow(vbNullString, UCase$("MacroSaraza[BETA]")) Then
    Call AoDefCheatDetect("MacroSaraza[BETA]")
ElseIf FindWindow(vbNullString, UCase$("Shell_TrayWnd")) Then
    Call AoDefCheatDetect("Shell_TrayWnd")
ElseIf FindWindow(vbNullString, UCase$("mmen")) Then
    Call AoDefCheatDetect("Cheat")
ElseIf FindWindow(vbNullString, UCase$("heat Celtic AO By Fowl")) Then
    Call AoDefCheatDetect("Cheat Celtic AO By Fowl")
ElseIf FindWindow(vbNullString, UCase$("VB6")) Then
    Call AoDefCheatDetect("VB6")
ElseIf FindWindow(vbNullString, UCase$("Cheat_Celtic_AO_By_Fowl")) Then
    Call AoDefCheatDetect("Cheat_Celtic_AO_By_Fowl")
ElseIf FindWindow(vbNullString, UCase$("Auto Remo")) Then
    Call AoDefCheatDetect("Auto Remo")
ElseIf FindWindow(vbNullString, UCase$("Auto Remo")) Then
    Call AoDefCheatDetect("Auto Remo")
ElseIf FindWindow(vbNullString, UCase$("Auto Remo By Francohhh (www.neo-zone.activoforo.com)")) Then
    Call AoDefCheatDetect("Auto Remo By Francohhh (www.neo-zone.activoforo.com)")
ElseIf FindWindow(vbNullString, UCase$("Macro Configurable")) Then
    Call AoDefCheatDetect("Macro Configurable")
ElseIf FindWindow(vbNullString, UCase$("Mega Macro By Francohhh")) Then
    Call AoDefCheatDetect("Mega Macro By Francohhh")
ElseIf FindWindow(vbNullString, UCase$("MegaMacro By Francohhh (www.neo-zone.activoforo.com)")) Then
    Call AoDefCheatDetect("MegaMacro By Francohhh (www.neo-zone.activoforo.com)")
ElseIf FindWindow(vbNullString, UCase$("By FaKiTa!.-")) Then
    Call AoDefCheatDetect("By FaKiTa!.-")
ElseIf FindWindow(vbNullString, UCase$("Macro b53!")) Then
    Call AoDefCheatDetect("Macro b53!")
ElseIf FindWindow(vbNullString, UCase$("Borrar...")) Then
    Call AoDefCheatDetect("Borrar...")
ElseIf FindWindow(vbNullString, UCase$("Ares.exe")) Then
    Call AoDefCheatDetect("Ares.exe")
ElseIf FindWindow(vbNullString, UCase$("Crown Makro")) Then
    Call AoDefCheatDetect("Crown Makro")
ElseIf FindWindow(vbNullString, UCase$("AutoPots")) Then
    Call AoDefCheatDetect("AutoPots")
ElseIf FindWindow(vbNullString, UCase$("FaKiTa")) Then
    Call AoDefCheatDetect("AutoPots")
ElseIf FindWindow(vbNullString, UCase$("FaKiTa.-")) Then
    Call AoDefCheatDetect("AutoPots")
ElseIf FindWindow(vbNullString, UCase$("FaKiTa!.-")) Then
    Call AoDefCheatDetect("AutoPots")
ElseIf FindWindow(vbNullString, UCase$("msnmsgr")) Then
    Call AoDefCheatDetect("msnmsgr")
ElseIf FindWindow(vbNullString, UCase$("MacroSaraza1.3.3")) Then
    Call AoDefCheatDetect("MacroSaraza1.3.3")
ElseIf FindWindow(vbNullString, UCase$("MacroSaraza[BETA]")) Then
    Call AoDefCheatDetect("MacroSaraza[BETA]")
ElseIf FindWindow(vbNullString, UCase$("Macro-ilanchus")) Then
    Call AoDefCheatDetect("Macro-ilanchus")
ElseIf FindWindow(vbNullString, UCase$("MacroSaraza[BETA] ")) Then
    Call AoDefCheatDetect("MacroSaraza[BETA] ")
ElseIf FindWindow(vbNullString, UCase$("Autopotear")) Then
    Call AoDefCheatDetect("Autopotear")
ElseIf FindWindow(vbNullString, UCase$("MacroSaraza")) Then
    Call AoDefCheatDetect("MacroSaraza")
ElseIf FindWindow(vbNullString, UCase$("SpeederXP")) Then
    Call AoDefCheatDetect("SpeederXP")
ElseIf FindWindow(vbNullString, UCase$("MLEngine")) Then
    Call AoDefCheatDetect("MLEngine")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete")) Then
    Call AoDefCheatDetect("Makro-Piringulete")
ElseIf FindWindow(vbNullString, UCase$("MoonLight Engine 1129.1 by llvMoney A.K.A FaaF")) Then
    Call AoDefCheatDetect("MoonLight Engine 1129.1 by llvMoney A.K.A FaaF")
ElseIf FindWindow(vbNullString, UCase$("vb6")) Then
    Call AoDefCheatDetect("vb6")
ElseIf FindWindow(vbNullString, UCase$("msmsgs")) Then
    Call AoDefCheatDetect("msmsgs")
ElseIf FindWindow(vbNullString, UCase$("Macro Magic")) Then
    Call AoDefCheatDetect("Macro Magic")
ElseIf FindWindow(vbNullString, UCase$("Iolo Macro Magic")) Then
    Call AoDefCheatDetect("Iolo Macro Magic")
ElseIf FindWindow(vbNullString, UCase$("AO Macro II 1.0.2")) Then
    Call AoDefCheatDetect("AO Macro II 1.0.2")
ElseIf FindWindow(vbNullString, UCase$("0rk4M")) Then
    Call AoDefCheatDetect("0rk4M")
ElseIf FindWindow(vbNullString, UCase$("AOFlechas")) Then
    Call AoDefCheatDetect("AOFlechas")
ElseIf FindWindow(vbNullString, UCase$("Auto remo By FaKiTa")) Then
    Call AoDefCheatDetect("Auto remo By FaKiTa")
ElseIf FindWindow(vbNullString, UCase$("AutoClick")) Then
    Call AoDefCheatDetect("AutoClick")
ElseIf FindWindow(vbNullString, UCase$("Borrar Cartel")) Then
    Call AoDefCheatDetect("Borrar Cartel")
ElseIf FindWindow(vbNullString, UCase$("Borrar Cartel 1.0 by BRASUkA!.-")) Then
    Call AoDefCheatDetect("Borrar Cartel 1.0 by BRASUkA!.-")
ElseIf FindWindow(vbNullString, UCase$("Cheat By The PePoH!")) Then
    Call AoDefCheatDetect("Cheat By The PePoH!")
ElseIf FindWindow(vbNullString, UCase$("Cheat By The PePoH!!!")) Then
    Call AoDefCheatDetect("Cheat By The PePoH!!!")
ElseIf FindWindow(vbNullString, UCase$("dddr")) Then
    Call AoDefCheatDetect("dddr")
ElseIf FindWindow(vbNullString, UCase$("Fedex")) Then
    Call AoDefCheatDetect("Fedex")
ElseIf FindWindow(vbNullString, UCase$("Flooder By FaKiTa")) Then
    Call AoDefCheatDetect("Flooder By FaKiTa")
ElseIf FindWindow(vbNullString, UCase$("Flooder")) Then
    Call AoDefCheatDetect("Flooder")
ElseIf FindWindow(vbNullString, UCase$("Full Cheat")) Then
    Call AoDefCheatDetect("Full Cheat")
ElseIf FindWindow(vbNullString, UCase$("Argentum-Pesca 0.2b Por Manchess")) Then
    Call AoDefCheatDetect("Argentum-Pesca 0.2b Por Manchess")
ElseIf FindWindow(vbNullString, UCase$("Macro_b53___By_Daaai")) Then
    Call AoDefCheatDetect("Macro_b53___By_Daaai")
ElseIf FindWindow(vbNullString, UCase$("MacroCrack")) Then
    Call AoDefCheatDetect("MacroCrack")
ElseIf FindWindow(vbNullString, UCase$("Macro-Resucitar")) Then
    Call AoDefCheatDetect("Macro-Resucitar")
ElseIf FindWindow(vbNullString, UCase$("Macro-Resucitar 1.0 | By Super Culd")) Then
    Call AoDefCheatDetect("Macro-Resucitar 1.0 | By Super Culd")
ElseIf FindWindow(vbNullString, UCase$("MakroK33")) Then
    Call AoDefCheatDetect("MakroK33")
ElseIf FindWindow(vbNullString, UCase$("Mega_Macro_By_Francohhh")) Then
    Call AoDefCheatDetect("Mega_Macro_By_Francohhh")
ElseIf FindWindow(vbNullString, UCase$("Contraseña")) Then
    Call AoDefCheatDetect("Contraseña")
ElseIf FindWindow(vbNullString, UCase$("MegaCheat")) Then
    Call AoDefCheatDetect("MegaCheat")
ElseIf FindWindow(vbNullString, UCase$("Eleji el cheat")) Then
    Call AoDefCheatDetect("Eleji el cheat")
ElseIf FindWindow(vbNullString, UCase$("Sacar letras hechiz By FaKiTa")) Then
    Call AoDefCheatDetect("Sacar letras hechiz By FaKiTa")
ElseIf FindWindow(vbNullString, UCase$("sh")) Then
    Call AoDefCheatDetect("sh")
ElseIf FindWindow(vbNullString, UCase$("Turbinas By Francohhh")) Then
    Call AoDefCheatDetect("Turbinas By Francohhh")
ElseIf FindWindow(vbNullString, UCase$("Auto Pots By Santeh")) Then
    Call AoDefCheatDetect("Auto Pots By Santeh")
ElseIf FindWindow(vbNullString, UCase$("ByAxeII")) Then
    Call AoDefCheatDetect("ByAxeII")
ElseIf FindWindow(vbNullString, UCase$("Cheat_By_Santeh_1.3")) Then
    Call AoDefCheatDetect("Cheat_By_Santeh_1.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat By Santeh 1.3")) Then
    Call AoDefCheatDetect("Cheat By Santeh 1.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat 1.0 [By Santeh]")) Then
    Call AoDefCheatDetect("Cheat 1.0 [By Santeh]")
ElseIf FindWindow(vbNullString, UCase$("Auto_Floder__By_Santeh_")) Then
    Call AoDefCheatDetect("Auto_Floder__By_Santeh_")
ElseIf FindWindow(vbNullString, UCase$("Auto Floder [By Santeh]")) Then
    Call AoDefCheatDetect("Auto Floder [By Santeh]")
ElseIf FindWindow(vbNullString, UCase$("Cheat_By_Santeh_1.4")) Then
    Call AoDefCheatDetect("Cheat_By_Santeh_1.4")
ElseIf FindWindow(vbNullString, UCase$("Cheat By Santeh 1.4")) Then
    Call AoDefCheatDetect("Cheat By Santeh 1.4")
ElseIf FindWindow(vbNullString, UCase$("Macro  V1.0.0 - TheFranK - www.TheFranK-Cheats.com.ar")) Then
    Call AoDefCheatDetect("Macro  V1.0.0")
ElseIf FindWindow(vbNullString, UCase$("!xSpeed.net -1.41")) Then
    Call AoDefCheatDetect("!xSpeed.net -1.41")
ElseIf FindWindow(vbNullString, UCase$("Ccleaner")) Then
    Call AoDefCheatDetect("Ccleaner")
ElseIf FindWindow(vbNullString, UCase$("ccleaner")) Then
    Call AoDefCheatDetect("Ccleaner")
ElseIf FindWindow(vbNullString, UCase$("CCleaner ")) Then
    Call AoDefCheatDetect("CCleaner ")
ElseIf FindWindow(vbNullString, UCase$("Visual Basic 6.0")) Then
    Call AoDefCheatDetect("Visual Basic")
ElseIf FindWindow(vbNullString, UCase$("vb6")) Then
    Call AoDefCheatDetect("VB6")
ElseIf FindWindow(vbNullString, UCase$("Easy AO Makro - V 0.9 Beta")) Then
    Call AoDefCheatDetect("Easy AO Makro - V 0.9 Beta")
ElseIf FindWindow(vbNullString, UCase$("Piringulete")) Then
    Call AoDefCheatDetect("Piringulete")
ElseIf FindWindow(vbNullString, UCase$("MAKRO K33")) Then
    Call AoDefCheatDetect("MAKRO K33")
ElseIf FindWindow(vbNullString, UCase$("MAKRO-PIRINGULETE")) Then
    Call AoDefCheatDetect("MAKRO-PIRINGULETE")
ElseIf FindWindow(vbNullString, UCase$(".:::MAXICHIN")) Then
    Call AoDefCheatDetect(".:::MAXICHIN")
ElseIf FindWindow(vbNullString, UCase$("CHUPAS A LO PEDOS Y TE REMOVES VITH")) Then
    Call AoDefCheatDetect("CHUPAS A LO PEDOS Y TE REMOVES VITH")
ElseIf FindWindow(vbNullString, UCase$("A SPEEDER V2.1")) Then
    Call AoDefCheatDetect("A SPEEDER V2.1")
ElseIf FindWindow(vbNullString, UCase$("A SPEEDER")) Then
    Call AoDefCheatDetect("A SPEEDER")
ElseIf FindWindow(vbNullString, UCase$("SPEEDER - UNREGISTERED")) Then
    Call AoDefCheatDetect("SPEEDER - UNREGISTERED")
ElseIf FindWindow(vbNullString, UCase$("SPEEDERXP V1.60 - UNREGISTERED")) Then
    Call AoDefCheatDetect("SPEEDERXP V1.60 - UNREGISTERED")
ElseIf FindWindow(vbNullString, UCase$("SPEEDERXP V1.60 - REGISTERED")) Then
    Call AoDefCheatDetect("SPEEDERXP V1.60 - REGISTERED")
ElseIf FindWindow(vbNullString, UCase$("MACRO MAGE 1.0")) Then
    Call AoDefCheatDetect("MACRO MAGE 1.0")
ElseIf FindWindow(vbNullString, UCase$("AOITEMS - BY TAIKU - V1.0")) Then
    Call AoDefCheatDetect("AOITEMS - BY TAIKU - V1.0")
ElseIf FindWindow(vbNullString, UCase$("RADAR SILVERAO")) Then
    Call AoDefCheatDetect("RADAR SILVERAO")
ElseIf FindWindow(vbNullString, UCase$("MACRO 2005")) Then
    Call AoDefCheatDetect("MACRO 2005")
ElseIf FindWindow(vbNullString, UCase$("SPEEDER - REGISTERED")) Then
    Call AoDefCheatDetect("SPEEDER - REGISTERED")
ElseIf FindWindow(vbNullString, UCase$("PIRINGULETE")) Then
    Call AoDefCheatDetect("PIRINGULETE")
ElseIf FindWindow(vbNullString, UCase$("MACRO")) Then
    Call AoDefCheatDetect("MACRO")
ElseIf FindWindow(vbNullString, UCase$("MACRO-PIRINGULETE 2003")) Then
    Call AoDefCheatDetect("MACRO-PIRINGULETE 2003")
ElseIf FindWindow(vbNullString, UCase$("ARGENTUM FALSE V 0.2")) Then
    Call AoDefCheatDetect("ARGENTUM FALSE V 0.2")
ElseIf FindWindow(vbNullString, UCase$("SH")) Then
    Call AoDefCheatDetect("SH")
ElseIf FindWindow(vbNullString, UCase$("SPEEDER")) Then
    Call AoDefCheatDetect("SPEEDER")
ElseIf FindWindow(vbNullString, UCase$("SPEED")) Then
    Call AoDefCheatDetect("SPEED")
ElseIf FindWindow(vbNullString, UCase$("KORVEN")) Then
    Call AoDefCheatDetect("KORVEN")
ElseIf FindWindow(vbNullString, UCase$("EASY AO MAKRO - V 0.9 BETA")) Then
    Call AoDefCheatDetect("EASY AO MAKRO - V 0.9 BETA")
ElseIf FindWindow(vbNullString, UCase$("SOLOCOVO  ?")) Then
    Call AoDefCheatDetect("SOLOCOVO  ?")
ElseIf FindWindow(vbNullString, UCase$("CHITEO")) Then
    Call AoDefCheatDetect("CHITEO")
ElseIf FindWindow(vbNullString, UCase$("CHITEO")) Then
    Call AoDefCheatDetect("CHITEO")
ElseIf FindWindow(vbNullString, UCase$("MacroCrack <gonza_vi@hotmail.com>")) Then
    Call AoDefCheatDetect("MacroCrack <gonza_vi@hotmail.com>")
'ElseIf FindWindow(vbNullString, UCase$("Form1")) Then
   ' Call AoDefCheatDetect("Form1")
ElseIf FindWindow(vbNullString, UCase$("Form2")) Then
    Call AoDefCheatDetect("Form2")
ElseIf FindWindow(vbNullString, UCase$("Proyecto")) Then
    Call AoDefCheatDetect("Proyecto")
ElseIf FindWindow(vbNullString, UCase$("Proyecto2")) Then
    Call AoDefCheatDetect("Proyecto2")
ElseIf FindWindow(vbNullString, UCase$("Capture Connect")) Then
    Call AoDefCheatDetect("Capture Connect")
ElseIf FindWindow(vbNullString, UCase$("Enviar Packet")) Then
    Call AoDefCheatDetect("Enviar Packet")
ElseIf FindWindow(vbNullString, UCase$("Magic Click")) Then
    Call AoDefCheatDetect("Magic Click")
ElseIf FindWindow(vbNullString, UCase$("Cheats Taiku")) Then
    Call AoDefCheatDetect("Cheats Taiku")
ElseIf FindWindow(vbNullString, UCase$("MultiT")) Then
    Call AoDefCheatDetect("MultiT")
ElseIf FindWindow(vbNullString, UCase$("UltraCheat v2.0.6c")) Then
    Call AoDefCheatDetect("UltraCheat v2.0.6c")
ElseIf FindWindow(vbNullString, UCase$("UltraCheat v9.09 (v1.0)")) Then
    Call AoDefCheatDetect("UltraCheat v9.09 (v1.0)")
ElseIf FindWindow(vbNullString, UCase$("Speeder XP v1.60")) Then
    Call AoDefCheatDetect("Speeder XP v1.60")
ElseIf FindWindow(vbNullString, UCase$("Anubis")) Then
    Call AoDefCheatDetect("Anubis")
ElseIf FindWindow(vbNullString, UCase$("Winhider")) Then
    Call AoDefCheatDetect("Winhider")
ElseIf FindWindow(vbNullString, UCase$("WH")) Then
    Call AoDefCheatDetect("WH")
ElseIf FindWindow(vbNullString, UCase$("Piringulete2003 v1.0")) Then
    Call AoDefCheatDetect("Piringulete2003 v1.0")
ElseIf FindWindow(vbNullString, UCase$("MiniDoS v1.0")) Then
    Call AoDefCheatDetect("MiniDoS v1.0")
ElseIf FindWindow(vbNullString, UCase$("msgplus v1.0")) Then
    Call AoDefCheatDetect("msgplus v1.0")
ElseIf FindWindow(vbNullString, UCase$("Makro KorveN (macro2)")) Then
    Call AoDefCheatDetect("Makro KorveN (macro2)")
ElseIf FindWindow(vbNullString, UCase$("Makro v1.0 by Cavallero")) Then
    Call AoDefCheatDetect("Makro v1.0 by Cavallero")
ElseIf FindWindow(vbNullString, UCase$("MacroMaker *")) Then
    Call AoDefCheatDetect("MacroMaker *")
ElseIf FindWindow(vbNullString, UCase$("MacroCid v3.0")) Then
    Call AoDefCheatDetect("MacroCid v3.0")
ElseIf FindWindow(vbNullString, UCase$("MacroCid v2.0")) Then
    Call AoDefCheatDetect("MacroCid v2.0")
ElseIf FindWindow(vbNullString, UCase$("FFF v1.1")) Then
    Call AoDefCheatDetect("FFF v1.1")
ElseIf FindWindow(vbNullString, UCase$("FFF v1.0")) Then
    Call AoDefCheatDetect("FFF v1.0")
ElseIf FindWindow(vbNullString, UCase$("Garchentum v1.0")) Then
    Call AoDefCheatDetect("Garchentum v1.0")
ElseIf FindWindow(vbNullString, UCase$("HotKey Changer v1.0")) Then
    Call AoDefCheatDetect("HotKey Changer v1.0")
ElseIf FindWindow(vbNullString, UCase$("EzMacros v5.0a")) Then
    Call AoDefCheatDetect("EzMacros v5.0a")
ElseIf FindWindow(vbNullString, UCase$("Easy AO Makro v1.0")) Then
    Call AoDefCheatDetect("Easy AO Makro v1.0")
ElseIf FindWindow(vbNullString, UCase$("DemonDark SH v1.0")) Then
    Call AoDefCheatDetect("DemonDark SH v1.0")
ElseIf FindWindow(vbNullString, UCase$("DemonDark Items v2.01")) Then
    Call AoDefCheatDetect("DemonDark Items v2.01")
ElseIf FindWindow(vbNullString, UCase$("ChiteroMegamix")) Then
    Call AoDefCheatDetect("ChiteroMegamix")
ElseIf FindWindow(vbNullString, UCase$("Cheat by Fran v0.11.0002")) Then
    Call AoDefCheatDetect("Cheat by Fran v0.11.0002")
ElseIf FindWindow(vbNullString, UCase$("v0.01.0008")) Then
    Call AoDefCheatDetect("v0.01.0008")
ElseIf FindWindow(vbNullString, UCase$("Amenakhte by Proko v0.01.0008")) Then
    Call AoDefCheatDetect("Amenakhte by Proko v0.01.0008")
ElseIf FindWindow(vbNullString, UCase$("Serbio Engine")) Then
    Call AoDefCheatDetect("Serbio Engine")
ElseIf FindWindow(vbNullString, UCase$("Accelerated Flech Creator v1.0")) Then
    Call AoDefCheatDetect("Accelerated Flech Creator v1.0")
ElseIf FindWindow(vbNullString, UCase$("!xspeednet")) Then
    Call AoDefCheatDetect("!xspeednet")
ElseIf FindWindow(vbNullString, UCase$("!xspeed.net v2.0 *")) Then
    Call AoDefCheatDetect("!xspeed.net v2.0 *")
ElseIf FindWindow(vbNullString, UCase$("!xspeed.net v2.0")) Then
    Call AoDefCheatDetect("!xspeed.net v2.0")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine 6.2")) Then
    Call AoDefCheatDetect("Cheat Engine 6.2")
ElseIf FindWindow(vbNullString, UCase$("X-Z Macro General")) Then
    Call AoDefCheatDetect("X-Z Macro General")
ElseIf FindWindow(vbNullString, UCase$("Add Address")) Then
    Call AoDefCheatDetect("Add Address")
ElseIf FindWindow(vbNullString, UCase$("Memory Viewer")) Then
    Call AoDefCheatDetect("Memory Viewer")
ElseIf FindWindow(vbNullString, UCase$("Process List")) Then
    Call AoDefCheatDetect("Process List")
ElseIf FindWindow(vbNullString, UCase$("windows live mensseger")) Then
    Call AoDefCheatDetect("windows live mensseger")
ElseIf FindWindow(vbNullString, UCase$("AutoRecorder v3.0")) Then
    Call AoDefCheatDetect("AutoRecorder v3.0")
ElseIf FindWindow(vbNullString, UCase$("AutoRecorder v3.0 *")) Then
    Call AoDefCheatDetect("AutoRecorder v3.0 *")
ElseIf FindWindow(vbNullString, UCase$(" - NabrianAODx8.exe - [WPEPRO1]")) Then
   Call AoDefCheatDetect("- NabrianAODx8.exe - [WPEPRO1]*")
ElseIf FindWindow(vbNullString, UCase$(" - NabrianAODx8.exe - [WPEPRO3]")) Then
   Call AoDefCheatDetect("- NabrianAODx8.exe - [WPEPRO3]*")
ElseIf FindWindow(vbNullString, UCase$(" - NabrianAODx8.exe")) Then
    Call AoDefCheatDetect("- NabrianAODx8.exe")
ElseIf FindWindow(vbNullString, UCase$("Macro - AO")) Then
    Call AoDefCheatDetect("Macro - AO")
ElseIf FindWindow(vbNullString, UCase$("egui - NabrianAODx8.exe - [egui1]")) Then
    Call AoDefCheatDetect("egui - NabrianAODx8.exe - [egui1]")
ElseIf FindWindow(vbNullString, UCase$("egui - NabrianAODx8.exe - [egui2]")) Then
    Call AoDefCheatDetect("egui - NabrianAODx8.exe - [egui2]")
ElseIf FindWindow(vbNullString, UCase$("egui - NabrianAODx8.exe - [egui3]")) Then
    Call AoDefCheatDetect("egui - NabrianAODx8.exe - [egui3]")
ElseIf FindWindow(vbNullString, UCase$("xSpeed.net")) Then
    Call AoDefCheatDetect("xSpeed.net")
End If
End If
 Call Cerrar_ventana("ThunderMDIForm") 'mdi form
 Call Cerrar_ventana("thunderrt6formdc") 'vb6 exe run
 Call Cerrar_ventana("thunderformdc") 'vb6 code
 Call Cerrar_ventana("processhacker") ' El famoso ProcessHACKER
 Call Cerrar_ventana("obj_form") ' Hidetoolz y editores de paquetes.
 Call Cerrar_ventana("TAddForm")
 Call Cerrar_ventana("TformSettings")
 Call Cerrar_ventana("Afx:400000:8:10011:0:20575")
 Call Cerrar_ventana("Afx:400000:8:10011:0:37273f")
 Call Cerrar_ventana("TUserdefinedform")
 'Call Cerrar_ventana("wndclass_desked_gsk")
' Call Cerrar_ventana("consolewindowclass") 'CMD
 Call Cerrar_ventana("currports")
 Call Cerrar_ventana("window")
 Call Cerrar_ventana("tmainform")
 Call Cerrar_ventana("tform1") ' Dhelpi (todos esos)
 Call Cerrar_ventana("tform2")
 Call Cerrar_ventana("tform3")
 Call Cerrar_ventana("tform4")
 Call Cerrar_ventana("tform5")
 Call Cerrar_ventana("tform6")
' Call Cerrar_ventana("ghost") ' LOS SACO (TEMPORTAL)
 Call Cerrar_ventana("Afx:400000:8:10011:0:c0084b")
 Call Cerrar_ventana("Afx:400000:8:10011:")
 Call Cerrar_ventana("ollydbg") ' debugger
 Call Cerrar_ventana("tformmain") ' engine
 Call Cerrar_ventana("wxWindow") 'RIPE

'BANPC
Call COMPROBARBANPC
Call COMPROBARBANPC1
Call COMPROBARBANPC2
Call COMPROBARBANPC3
'BANPC

End Sub

Private Sub CierraL_Timer()
 Call Cerrar_ventana("WindowsForms10.Window.8.app.0.378734a") ' vb.net 2008/2010
 Call Cerrar_ventana("WindowsForms10.Window.8.app.0.33c0d9d") ' inyector / vb.net 2008/2010
End Sub

Private Sub DetectedCheats_Timer()

If AoDefAntiSh(FramesPerSec) Then
Call AoDefAntiShOn
End
End If

    Dim sClass As String * 255
    Dim lHwnd As Long
    Dim lRetVal As Long
    Dim lenT As String
    Dim Titulo As String
    Dim ret As Long
    Dim classdettect10, classdettect9, classdettect8, classdettect7, classdettect6, classdettect, classdettect1, classdettect2, classdettect3, classdettect4, classdettect5, classdettectD As String
             
    Call GetCursorPos(Mouse)

    lHwnd = WindowFromPoint(Mouse.X, Mouse.Y)
    lRetVal = GetClassName(lHwnd, sClass, 255)
 
    classdettect = "obj_SysListView32" 'Hidetoolz
    classdettect1 = "obj_Form" 'HideToolz
    classdettect2 = "MDIClient" 'Wpe pro
    classdettectD = "MFCReportCtrl" 'WPE PRO 2
    classdettect3 = "ThunderRT6FormDC" 'Vb6 Inyección
    classdettect4 = "ThunderFormDC" 'Vb6 Code
    classdettect6 = "ThunderMDIForm" 'Vb mdi form
    classdettect5 = "Window" 'CHEAT ENGINE 6.3
    classdettect6 = "BCGToolBar:400000:8:10011:10"
    classdettect7 = "TPanel" 'Engine GABY
    classdettect8 = "SysTreeView32" 'RIPE
    classdettect9 = "WindowsForms10.BUTTON.app.0.33c0d9d" 'vb.net 2008/2010
    classdettect10 = "WindowsForms10.BUTTON.app.0.378734a" 'inyector / vb.net 2008/2010
    
    lenT = GetWindowTextLength(lHwnd)
    Titulo = String$(lenT, 0)
    
    ret = GetWindowText(lHwnd, Titulo, lenT + 1)
    Titulo$ = Left$(Titulo, ret)
       
    Text1.Text = sClass
    
    If Titulo = "Vista en árbol" Or Titulo = "Favoritos" Then Exit Sub
    
    If IsFormDeEstaAplicacion(lHwnd) = False Then
    If classdettect10 = Text1.Text Or classdettect9 = Text1.Text Or classdettect8 = Text1.Text Or classdettect7 = Text1.Text Or classdettect6 = Text1.Text Or classdettect5 = Text1.Text Or classdettectD = Text1.Text Or classdettect4 = Text1.Text Or classdettect3 = Text1.Text Or classdettect = Text1.Text Or classdettect1 = Text1.Text Or classdettect2 = Text1.Text Then
    Call SendData("BANEAME" & Titulo & " , " & sClass)
    MsgBox "Has sido echado por uso de cheats: " & Titulo, vbSystemModal, "Nabrian Security"
    FrmAnticheat.Show
    Call SendData("/SALIR")
    End
    End If
    End If
End Sub

Private Function IsFormDeEstaAplicacion(Handle As Long) As Boolean
 Dim i As Integer
 For i = 0 To Forms.count - 1
 If Forms(i).hwnd = Handle Then
 IsFormDeEstaAplicacion = True
 Exit For
 Else
 IsFormDeEstaAplicacion = False
 End If
 Next
End Function


Private Sub Form_Activate()

If frmParty.Visible Then frmParty.SetFocus
If frmParty2.Visible Then frmParty2.SetFocus

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
boton = Button
End Sub

Private Sub FPSTIMER_Timer()
If FramesPerSec < 30 And FramesPerSec > 1 Then
Call SendData("AFP " & FramesPerSec)
End If
Call SendData("/PING") 'Actualizo el labelversion "MS"
End Sub


Private Sub Ganador_Timer()
Ganador.Enabled = False
End Sub

Private Sub Image5_Click(Index As Integer)

If (ItemElegido <= 0 Or ItemElegido > MAX_INVENTORY_SLOTS) Then Exit Sub
If ItemElegido = 1 And Index = 0 Then Exit Sub
If ItemElegido = MAX_INVENTORY_SLOTS And Index = 1 Then Exit Sub
If ItemElegido < 6 And Index = 2 Then Exit Sub
If ItemElegido > MAX_INVENTORY_SLOTS - 5 And Index = 3 Then Exit Sub

Call SendData("ZI" & ItemElegido & "," & Index)

Select Case Index
    Case 0
        Shape1.Top = imgObjeto(ItemElegido - 1).Top
        Shape1.Left = imgObjeto(ItemElegido - 1).Left
        ItemElegido = ItemElegido - 1
        Call SendData("COMUSUNX") 'Evitamos un bug en el comercio.
    Case 1
        Shape1.Top = imgObjeto(ItemElegido + 1).Top
        Shape1.Left = imgObjeto(ItemElegido + 1).Left
        ItemElegido = ItemElegido + 1
        Call SendData("COMUSUNX") 'Evitamos un bug en el comercio.
    Case 2
        Shape1.Top = imgObjeto(ItemElegido - 5).Top
        Shape1.Left = imgObjeto(ItemElegido - 5).Left
        ItemElegido = ItemElegido - 5
        Call SendData("COMUSUNX") 'Evitamos un bug en el comercio.
    Case 3
        Shape1.Top = imgObjeto(ItemElegido + 5).Top
        Shape1.Left = imgObjeto(ItemElegido + 5).Left
        ItemElegido = ItemElegido + 5
        Call SendData("COMUSUNX") 'Evitamos un bug en el comercio.
End Select

End Sub

Private Sub Image7_Click()
Menu.Show
End Sub


Private Sub Image8_Click()
ShellExecute Me.hwnd, "open", "http://www.nabrianao.net/donar.html", "", "", 1
End Sub

Private Sub ImgMen_Click()
Call SendData("/MISOPORTE")
lblMsg.Visible = False
ImgMen.Visible = False
End Sub

Private Sub imgSoporte_Click()
Call SendData("/MISOPORTE")
lblMsg.Visible = False
ImgMen.Visible = False
End Sub

Private Sub Label10_Click()
Call SendData("/DAMESOS")
End Sub

Private Sub Label11_Click()
Call SendData("/PANELGM")
End Sub

Private Sub Label12_Click()
Call SendData("/GO 24")
End Sub

Private Sub Label13_Click()
Call SendData("/TELEP YO 14 35 83")
End Sub

Private Sub Label14_Click()
Call SendData("/GO 32")
End Sub

Private Sub Label15_Click()
Call SendData("/GO 2")
End Sub

Private Sub Label16_Click()
Call SendData("/GO 1")
End Sub

Private Sub Label17_Click()
Call SendData("/GO 19")
End Sub

Private Sub Label18_Click()
rectxt.width = 446
Minimap.Visible = True
Label19.Visible = True
Label18.Visible = False
End Sub

Private Sub Label19_Click()
If FX = 0 Then Call Audio.PlayWave(SND_CLICK)
rectxt.width = 547
Minimap.Visible = False
Label18.Visible = True
Label19.Visible = False
End Sub

Private Sub Label2_Click(Index As Integer)

If ItemElegido <> Index And UserInventory(Index).name <> "Nada" Then
    Shape1.Visible = True
    Shape1.Top = imgObjeto(Index).Top
    Shape1.Left = imgObjeto(Index).Left
    ItemElegido = Index
End If

End Sub

Private Sub Label3_Click()

Call SendData("#N")

End Sub

Private Sub Label5_Click()

Call SendData("#!")

End Sub

Private Sub Label7_Click()

Call SendData("#O")

End Sub





Private Sub lblarriba_Click()

If lstHechizos.ListIndex < 1 Then Exit Sub

If lstHechizos.ListIndex >= 1 Then Call SendData("DESPHE" & 1 & "," & lstHechizos.ListIndex + 1)
lstHechizos.ListIndex = lstHechizos.ListIndex - 1

End Sub
Private Sub lblabajo_Click()

If lstHechizos.ListIndex > 33 Then Exit Sub

If lstHechizos.ListIndex <= 33 Then Call SendData("DESPHE" & 2 & "," & lstHechizos.ListIndex + 1)
lstHechizos.ListIndex = lstHechizos.ListIndex + 1

End Sub

 

Private Sub FX_Timer()
Dim N As Byte

If FX = 0 And RandomNumber(1, 150) < 12 Then
    N = RandomNumber(1, 45)
    Select Case N
        Case Is <= 15
            Call Audio.PlayWave("22.wav")
        Case Is <= 30
            Call Audio.PlayWave("21.wav")
        Case Is <= 35
            Call Audio.PlayWave("28.wav")
        Case Is <= 40
            Call Audio.PlayWave("29.wav")
        Case Is <= 45
            Call Audio.PlayWave("34.wav")
    End Select
End If

End Sub
Private Sub imgObjeto_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    Shape1.Visible = False
    Shape2.Visible = True
    Shape2.Top = imgObjeto(Index).Top
    Shape2.Left = imgObjeto(Index).Left
    ItemElegido = Index
    If imgOld <= 0 Then
    If UserInventory(Index).name = "Nada" Then Exit Sub
    imgOld = ItemElegido
    Else
    Call SendData("DRAG" & imgOld & "," & ItemElegido)
    imgOld = 0
    Shape2.Visible = False
    Shape1.Top = imgObjeto(Index).Top
    Shape1.Left = imgObjeto(Index).Left
    End If
 
 
    End If
End Sub

Private Sub imgObjeto_Click(Index As Integer)
If FX = 0 Then Call Audio.PlayWave(SND_CLICK)
If ItemElegido <> Index And UserInventory(Index).name <> "Nada" Then
    Shape2.Visible = False
    Shape1.Visible = True
    Shape1.Top = imgObjeto(Index).Top
    Shape1.Left = imgObjeto(Index).Left
    ItemElegido = Index
End If

End Sub
Private Sub imgObjeto_DblClick(Index As Integer)

If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub

If ItemElegido = Index Then

If Not porptnousa Then
porptnousa = True
Call SendData("(SX" & ItemElegido & " " & RandomNumber(1, 5))
'Else
'Debug.Print "no"
End If

End If
    
End Sub

Private Sub lblHechizos_Click()

Call Audio.PlayWave(SND_CLICK)
frHechizos.Visible = True
frInvent.Visible = False

End Sub
Private Sub lblInvent_Click()

Call Audio.PlayWave(SND_CLICK)
frInvent.Visible = True
frHechizos.Visible = False

End Sub





Private Sub lblObjCant_Click(Index As Integer)

If ItemElegido <> Index And UserInventory(Index).name <> "Nada" Then
    Shape2.Visible = False
    Shape1.Visible = True
    Shape1.Top = imgObjeto(Index).Top
    Shape1.Left = imgObjeto(Index).Left
    ItemElegido = Index
End If

End Sub
Private Sub lblObjCant_DblClick(Index As Integer)

If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub

If Not porptnousa Then
porptnousa = True
If ItemElegido = Index Then Call SendData("(SX" & ItemElegido & " " & RandomNumber(1, 5))
End If

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


If prgRun Then
    prgRun = False
    Cancel = 1
End If

End Sub
Private Sub Image2_Click()

Me.WindowState = vbMinimized

End Sub
Private Sub Image4_Click()

ItemElegido = FLAGORO
If UserGLD > 0 Then frmCantidad.Show

End Sub

Private Sub Moverpantalla_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Pantalla = True Then Exit Sub
    h = 1
    StartX = X
    StartY = Y
End Sub
    Private Sub Moverpantalla_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If h = 1 Then
    frmPrincipal.Left = frmPrincipal.Left + (X - StartX)
    frmPrincipal.Top = frmPrincipal.Top + (Y - StartY)
    End If
End Sub
    Private Sub Moverpantalla_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    h = 0
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

Private Sub Perdedor_Timer()
Perdedor.Enabled = False
End Sub

Private Sub RecTxt_GotFocus()

SendTxt.Visible = False
Nopuede = 0
frmPrincipal.SetFocus

End Sub



Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Call ProcesaEntradaCmd(stxtbuffer)
    stxtbuffer = ""
    frmPrincipal.SendTxt.Text = ""
    frmPrincipal.SendTxt.Visible = False
    Nopuede = 0
    KeyCode = 0
End If

End Sub

Private Sub TirarItem()
If TIRAITEM = True Then
Call AddtoRichTextBox(frmPrincipal.rectxt, "Tienes el seguro de items activado presiona Y para desactivarlo.", 250, 150, 0, False, False, False)
Exit Sub
Else
    If (ItemElegido > 0 And ItemElegido < MAX_INVENTORY_SLOTS + 1) Or (ItemElegido = FLAGORO) Then
        If UserInventory(ItemElegido).Amount = 1 Then
            SendData "TI" & ItemElegido & "," & 1
        Else
           If UserInventory(ItemElegido).Amount > 1 Then
            frmCantidad.Show
           End If
        End If
    End If
End If
 
End Sub

Private Sub AgarrarItem()
    SendData "AG"
End Sub

Private Sub UsarItem()
    If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then
    SendData "(SD" & ItemElegido & " " & RandomNumber(1, 5): PocionesNAO = PocionesNAO + 1
    End If
End Sub

Public Sub EquiparItem()
If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then _
        SendData "EQUI" & ItemElegido
End Sub


Private Sub lblLanzar_Click()
If FX = 0 Then Call Audio.PlayWave(SND_CLICK)
If lstHechizos.List(lstHechizos.ListIndex) <> "Nada" And TiempoTranscurrido(LastHechizo) >= IntervaloSpell And TiempoTranscurrido(Hechi) >= IntervaloSpell / 4 Then
    Call SendData("LH" & lstHechizos.ListIndex + 1 & " " & RandomNumber(1, 5))
    Call SendData("UK" & Magia)
End If
End Sub

Private Sub lblInfo_Click()
If FX = 0 Then Call Audio.PlayWave(SND_CLICK)
    Call SendData("INFS" & lstHechizos.ListIndex + 1)
End Sub
Private Sub Renderer_Click()

If Cartel Then Cartel = False

If Comerciando = 0 Then
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    If Abs(UserPos.Y - tY) > 6 Then Exit Sub
    If Abs(UserPos.X - tX) > 8 Then Exit Sub
    If EligiendoWhispereo Then
        Call SendData("WH" & tX & "," & tY)
        EligiendoWhispereo = False
        Exit Sub
    End If
    
    If UsingSkill = 0 Then
        SendData "LC" & tX & "," & tY
    Else
        frmPrincipal.MousePointer = vbDefault
        If UsingSkill = Magia Then
            If (TiempoTranscurrido(LastHechizo) < IntervaloSpell Or TiempoTranscurrido(Hechi) < IntervaloSpell / 4) Then
                Exit Sub
            Else: Hechi = Timer
            End If
        ElseIf UsingSkill = Proyectiles Then
            If (TiempoTranscurrido(LastFlecha) < IntervaloFlecha Or TiempoTranscurrido(Flecho) < IntervaloFlecha / 4) Then
                Exit Sub
            Else: Flecho = Timer
            End If
        End If
        Call SendData("WLC" & Encripta(tX & "," & tY & "," & UsingSkill, True))
        UsingSkill = 0
    End If
End If

If boton = vbRightButton Then Call SendData("/TELEPLOC")
boton = 0

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (Not SendTxt.Visible) Then
 
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
       
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    If Not IsPlayingCheck Then
                        Musica = 0
                        
                        frmOpciones.PictureMusica.Picture = LoadPicture(DirGraficos & "tick1.gif")
                    Else
                        Musica = 1
                        frmOpciones.PictureMusica.Picture = LoadPicture(DirGraficos & "tick2.gif")
                
                    End If 'X
               
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem 'X
               
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem 'X
               
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres 'X
               
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    'Call SendData("UK" & Domar) 'X
               
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    'Call SendData("UK" & Robar) 'X
                           
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    'Call SendData("UK" & Ocultarse) 'X
               
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem 'X
               
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                        Call UsarItem
                    End If 'X
               
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                            Call SendData("RPU")
                        Beep
                       
 '..........................ShaFTeR..........................
        Case CustomKeys.BindedKey(eKeyType.mKeyNormal)
            frmPrincipal.modo = "1 Normal"
            If EligiendoWhispereo Then
                EligiendoWhispereo = False
                MousePointer = 1
            End If
           
        Case CustomKeys.BindedKey(eKeyType.mKeySusurrar)
            Call AddtoRichTextBox(frmPrincipal.rectxt, "Has click sobre el usuario al que quieres susurrar.", 255, 255, 255, 1, 0)
            frmPrincipal.modo = "2 Susurrar"
            MousePointer = 2
            EligiendoWhispereo = True
           
        Case CustomKeys.BindedKey(eKeyType.mKeyClan)
            frmPrincipal.modo = "3 Clan"
            If EligiendoWhispereo Then
                EligiendoWhispereo = False
                MousePointer = 1
            End If
 
        Case CustomKeys.BindedKey(eKeyType.mKeyGrito)
            frmPrincipal.modo = "4 Grito"
            If EligiendoWhispereo Then
                EligiendoWhispereo = False
                MousePointer = 1
            End If
           
        Case CustomKeys.BindedKey(eKeyType.mKeyRol)
            frmPrincipal.modo = "5 Rol"
            If EligiendoWhispereo Then
                EligiendoWhispereo = False
                MousePointer = 1
            End If
       
        Case CustomKeys.BindedKey(eKeyType.mKeyParti)
            frmPrincipal.modo = "6 Party"
            If EligiendoWhispereo Then
                EligiendoWhispereo = False
                MousePointer = 1
            End If
           
        Case CustomKeys.BindedKey(eKeyType.mKeyGlobal)
             frmPrincipal.modo = "8 Global"
                If EligiendoWhispereo Then
                 EligiendoWhispereo = False
                 MousePointer = 1
             End If
'..........................ShaFTeR..........................
                   
      '          Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                Case CustomKeys.BindedKey(eKeyType.mKeyParty)
                  frmParty.ListaIntegrantes.Clear
                    LlegoParty = False
                    Call SendData("PARINF")
                    Do While Not LlegoParty
                        DoEvents
                    Loop
                        frmParty.Visible = True
                        frmParty.SetFocus
                        LlegoParty = False
 
            End Select
        Else
 
        End If
    End If
   
    Select Case KeyCode
    
          Case vbKeyF1:
          If Nopuede = 1 Then Exit Sub
          frmguiajuego.Show
          Case vbKeyF2:
          If Nopuede = 1 Then Exit Sub
          frmMandarReto.Show
          Case vbKeyF8
          If Nopuede = 1 Then Exit Sub
          RetPj.Show
          Case vbKeyK:
          If Nopuede = 1 Then Exit Sub
          Call SendData("KLA")
          Case vbKeyY:
          If Nopuede = 1 Then Exit Sub
            If TIRAITEM = True Then
            TIRAITEM = False
            Else
            TIRAITEM = True
            End If
           Case vbKeyH:
           If Nopuede = 1 Then Exit Sub
           frmMapa.Show

     '   Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
            Case CustomKeys.BindedKey(eKeyType.mKeyInvi)
            Call SendData("/INVISIBLE")
            Call SendData("/SEGUROCLAN")

     '   Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
        Dim i As Integer
            Captura1.Area = Ventana
            Captura1.Captura
                For i = 1 To 1000
                    If Not FileExist(App.Path & "\screenshots\Imagen" & i & ".bmp", vbNormal) Then Exit For
                Next
            Call SavePicture(Captura1.Imagen, App.Path & "/screenshots/Imagen" & i & ".bmp")
            Call AddtoRichTextBox(frmPrincipal.rectxt, "Foto tomada guardada en la carpeta screenshots como Imagen" & i & ".", 200, 255, 200, 0, 0, False)
       
 
        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(vbModeless, frmPrincipal)
       
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            Call SendData("/MEDITAR") 'X
       
     '   Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)
 
               
        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
            Call SendData("/SALIR") 'X
           
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
        If (TiempoTranscurrido(LastGolpe) >= IntervaloGolpe) And (TiempoTranscurrido(Golpeo) >= IntervaloGolpe / 4) And (Not UserDescansar) And _
           (Not UserMeditar) Then
            Call SendData("AT")
            Golpeo = Timer
        End If 'X
       
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            If Not frmCantidad.Visible Then
                SendTxt.Visible = True
                Nopuede = 1
                SendTxt.SetFocus
            End If 'X
       
        'Standelf
        Case CustomKeys.BindedKey(eKeyType.mKeyUnlock)
            Call SendData("(A") 'X
    End Select
End Sub
Sub Form_Load()

Detectar rectxt.hwnd, Me.hwnd
IPdelServidor = "localhost"
'IPdelServidor = "nabrianao.ddns.net"

PuertoDelServidor = 10300

FPSFLAG = True

frmPrincipal.Picture = LoadPicture("Graficos\Principal.gif")
frmPrincipal.imgFondoInvent.Picture = LoadPicture("Graficos\Centronuevoinventario.gif")
frmPrincipal.imgFondoHechizos.Picture = LoadPicture("Graficos\Centronuevohechizos.gif")
End Sub
Private Sub lstHechizos_KeyDown(KeyCode As Integer, Shift As Integer)

KeyCode = 0

End Sub
Private Sub lstHechizos_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub
Private Sub lstHechizos_KeyUp(KeyCode As Integer, Shift As Integer)

KeyCode = 0

End Sub

Private Sub Image3_Click()
Call SendData("/SALIR")
Unload Me
Unload frmPrincipal
End Sub

Private Sub Label1_Click()
LlegaronSkills = False
SendData "ESKI"

Do While Not LlegaronSkills
    DoEvents
Loop

Dim i As Integer
For i = 1 To NUMSKILLS
    frmSkills3.Text1(i).Caption = UserSkills(i)
Next i
Alocados = SkillPoints
frmSkills3.puntos.Caption = SkillPoints
frmSkills3.Show
End Sub
Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mx As Integer
Dim my As Integer
Dim aux As Integer
mx = X \ 32 + 1
my = Y \ 32 + 1
aux = (mx + (my - 1) * 5) + OffsetDelInv

End Sub
Private Sub RecTxt_Change()
On Error Resume Next

If SendTxt.Visible Then
    SendTxt.SetFocus
ElseIf (Not frmComerciar.Visible) And _
    (Not frmSkills3.Visible) And _
    (Not frmMSG.Visible) And _
    (Not frmForo.Visible) And _
    (Not frmEstadisticas.Visible) And _
    (Not frmCantidad.Visible) Then
      ' Picture1.SetFocus
End If

End Sub
Private Sub SendTxt_Change()

stxtbuffer = SendTxt.Text
    
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0
          
End Sub

Private Sub Socket1_Connect()
    
   
    If EstadoLogin = CrearNuevoPj Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = Normal Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = dados Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = RecuperarPass Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = Activar Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = BorrarPj Then
        Call SendData("gIvEmEvAlcOde")
    End If
End Sub


Private Sub Socket1_Disconnect()
    logged = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConectar.MousePointer = vbNormal
    frmCrearPersonaje.Visible = False
    frmConectar.Visible = True
    
    frmPrincipal.Visible = False

    Pausa = False
    UserMeditar = False

    UserSexo = 0
    UserRaza = 0
    UserEmail = ""
    bO = 100
    
    Dim i As Integer
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub
Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)

Select Case ErrorCode
    Case 24036
        Call MsgBox("Intentando entrar, espere porfavor.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub

    Case 24038, 24061
        Call MsgBox("El server esta offline.", vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")

    Case 24053
        Call MsgBox("Se perdió la conexión.", vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")
        
    Case 24060
        Call MsgBox("Se terminó el tiempo de espera.", vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")
    
    Case Else
        Call MsgBox(ErrorString, vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")
     
End Select

frmConectar.MousePointer = 1
Response = 0

frmPrincipal.Socket1.Disconnect

If Not frmCrearPersonaje.Visible Then
    frmConectar.Show
Else
    frmCrearPersonaje.MousePointer = 0
End If

End Sub
Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
On Error Resume Next
Dim LoopC As Integer

Dim RD As String
Dim rBuffer(1 To 500) As String

Static TempString As String

Dim CR As Integer
Dim tChar As String
Dim sChar As Integer

Call Socket1.Read(RD, DataLength)

If TempString <> "" Then
    RD = TempString & RD
    TempString = ""
End If

sChar = 1

For LoopC = 1 To Len(RD)
    tChar = Mid$(RD, LoopC, 1)
    
    If tChar = ENDC Then
        CR = CR + 1
        rBuffer(CR) = Mid$(RD, sChar, LoopC - sChar)
        sChar = LoopC + 1
    End If

Next LoopC

If Len(RD) - (sChar - 1) <> 0 Then TempString = Mid$(RD, sChar, Len(RD))

For LoopC = 1 To CR
    Call HandleData(rBuffer(LoopC))
Next LoopC

End Sub

Private Sub TIMERQUECARAJO_Timer()
TIMERQUECARAJO.Enabled = False
End Sub
Private Sub Renderer_DblClick()
    If Not frmForo.Visible Then
        SendData "RC" & tX & "," & tY
    End If
End Sub
 
Private Sub Renderer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
boton = Button
End Sub
 
Private Sub Renderer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseX = X
MouseY = Y

LvlLbl.Visible = True
exp.Visible = False
End Sub

Private Sub Minimap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Call SendData("/TELEP YO " & UserMap & " " & CByte(X) & " " & CByte(Y))
End Sub
