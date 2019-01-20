VERSION 5.00
Begin VB.Form Regreso 
   BorderStyle     =   0  'None
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   495
      Left            =   3840
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   600
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "Regreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.PICTURE = LoadPicture(App.Path & "\graficos\Regreso.gif")
End Sub

Private Sub Image1_Click()
Call SendData("/REGRESAR")
Unload Regreso
End Sub

Private Sub Image2_Click()
Unload Regreso
End Sub
