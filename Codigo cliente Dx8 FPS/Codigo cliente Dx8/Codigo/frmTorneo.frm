VERSION 5.00
Begin VB.Form frmTorneo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Quieren participar"
   ClientHeight    =   3825
   ClientLeft      =   2760
   ClientTop       =   3960
   ClientWidth     =   3225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3030
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jugadores:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu Sumonear 
      Caption         =   "Sumonear"
      Index           =   0
   End
End
Attribute VB_Name = "frmTorneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()

Unload Me
End Sub



Private Sub Sumonear_Click(Index As Integer)
If List1.ListIndex = -1 Then Exit Sub
Call SendData("/SUM " & ReadFieldOptimizado(1, List1, Asc(":")))
List1.RemoveItem List1.ListIndex
Label2 = List1.ListCount
End Sub

Private Sub Form_Load()


'Label2 = List1.ListCount
End Sub
