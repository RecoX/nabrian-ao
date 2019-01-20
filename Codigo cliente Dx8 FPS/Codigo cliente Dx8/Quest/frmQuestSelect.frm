VERSION 5.00
Begin VB.Form frmQuestSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Elegir Quest"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmQuestSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub PonerListaQuest(ByVal Rdata As String)

Dim j As Integer, k As Integer
For j = 0 To List1.ListCount - 1
    Me.List1.RemoveItem 0
Next j
k = CInt(ReadFieldOptimizado(1, Rdata, 44))

For j = 1 To k
    List1.AddItem ReadFieldOptimizado(1 + j, Rdata, 44)
Next j

Me.Show , frmPrincipal

End Sub
Private Sub List1_Click()
Call SendData("INFD" & List1.Text)
End Sub
