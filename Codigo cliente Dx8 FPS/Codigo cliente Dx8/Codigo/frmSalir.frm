VERSION 5.00
Begin VB.Form FrmMisiones 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   2730
      Width           =   1455
   End
   Begin VB.Frame FF 
      BackColor       =   &H00000000&
      Caption         =   "Misiones (TEMPLARIO)"
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
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      Begin VB.Label LabelInfo 
         BackColor       =   &H00000000&
         Caption         =   "Cargando información."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   -80
      Width           =   5295
   End
End
Attribute VB_Name = "FrmMisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub command1_Click()
Call SendData("/MISION")
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


