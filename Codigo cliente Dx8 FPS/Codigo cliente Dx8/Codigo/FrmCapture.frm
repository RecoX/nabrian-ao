VERSION 5.00
Begin VB.Form frmCapture 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAPTURE THE FLAG"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Bloquear"
      Height          =   495
      Left            =   9840
      TabIndex        =   21
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   9840
      TabIndex        =   20
      Text            =   "5"
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cuenta"
      Height          =   495
      Left            =   8760
      TabIndex        =   19
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reglas"
      Height          =   495
      Left            =   8760
      TabIndex        =   18
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ACTIV DES-ACTIV "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9000
      TabIndex        =   17
      Top             =   360
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Jugadores"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.CommandButton Command6 
         Caption         =   "Sumonear"
         Height          =   375
         Left            =   5880
         TabIndex        =   23
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Sumonear"
         Height          =   375
         Left            =   1320
         TabIndex        =   22
         Top             =   4080
         Width           =   975
      End
      Begin VB.ListBox List1 
         ForeColor       =   &H000000C0&
         Height          =   2205
         Left            =   480
         TabIndex        =   11
         Top             =   1080
         Width           =   3375
      End
      Begin VB.ListBox List2 
         ForeColor       =   &H00C00000&
         Height          =   2205
         Left            =   4560
         TabIndex        =   10
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Text            =   "Jugador"
         Top             =   4680
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4560
         TabIndex        =   8
         Text            =   "Jugador"
         Top             =   4680
         Width           =   2415
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   7080
         TabIndex        =   6
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Quitar"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Quitar"
         Height          =   255
         Left            =   6840
         TabIndex        =   4
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Borrar Lista"
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Borrar Lista"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Act. Listas"
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Criminales"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Ciudadanos"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4560
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Vs."
         Height          =   375
         Left            =   3960
         TabIndex        =   14
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "1º 2º 3º 4º 5º 6º 7º 8º 9º 10º"
         Height          =   1935
         Left            =   4320
         TabIndex        =   13
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "1º 2º 3º 4º 5º 6º 7º 8º 9º 10º"
         Height          =   2055
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call SendData("/CAPTURETF " & text5)
End Sub


  


Private Sub Command10_Click()

If List2.ListIndex > -1 Then List2.RemoveItem List2.ListIndex

End Sub

Private Sub Command11_Click()
List2.Clear
End Sub

Private Sub Command12_Click()
List1.Clear
End Sub

Private Sub Command13_Click()


Dim n_Items As Long
Dim i As Integer
Dim Items As String
Dim El_Item As String
Dim length As Long
    n_Items = SendMessage(List2.hwnd, LB_GETCOUNT, 0, 0)
    For i = 0 To n_Items - 1
        length = SendMessage(List2.hwnd, LB_GETTEXTLEN, i, 0)
           
        El_Item = Space$(length + 1)

        length = SendMessage(List2.hwnd, LB_GETTEXT, i, ByVal El_Item)
         Call SendData("/SUM " & Replace(El_Item, Chr(0), vbNullString))
        'Items = Items & Replace(El_Item, Chr(0), vbNullString)
  
    Next i
           
   ' MsgBox Items, vbInformation
Call SendData("BLOQTEAMCIUDA")
End Sub

Private Sub Command14_Click()
Dim n_Items As Long
Dim i As Integer
Dim Items As String
Dim El_Item As String
Dim length As Long
    n_Items = SendMessage(List1.hwnd, LB_GETCOUNT, 0, 0)
    For i = 0 To n_Items - 1
        length = SendMessage(List1.hwnd, LB_GETTEXTLEN, i, 0)
           
        El_Item = Space$(length + 1)

        length = SendMessage(List1.hwnd, LB_GETTEXT, i, ByVal El_Item)
         Call SendData("/SUM " & Replace(El_Item, Chr(0), vbNullString))
        'Items = Items & Replace(El_Item, Chr(0), vbNullString)
  
    Next i
Call SendData("BLOQTEAMCAOS")
End Sub

Private Sub Command15_Click()
Call SendData("/VERCAPTURE")
End Sub

Private Sub Command2_Click()
Call SendData("/RVSG El objetivo del juego es llegar a la base enemiga y agarrar su bandera. Una vez que el jugador la tenga, debe volver a su base para ganar la ronda. Cuando éste muere, la bandera cae.~244~168~9~1~0")
End Sub

Private Sub Command3_Click()
Call SendData("/RVSG Por favor todos contra la pared de arriba, el que no este pegado contra la pared quedara descalificado y no jugara. Comenzamos a la cuenta de...~244~168~9~1~0")
Call SendData("/FCUENTA " & Text3)
End Sub

Private Sub Command4_Click()
Call SendData("/BLOQCAPTURE")
End Sub

Private Sub Command5_Click()
Dim n_Items As Long
Dim i As Integer
Dim Items As String
Dim El_Item As String
Dim length As Long
    n_Items = SendMessage(List1.hwnd, LB_GETCOUNT, 0, 0)
    For i = 0 To n_Items - 1
        length = SendMessage(List1.hwnd, LB_GETTEXTLEN, i, 0)
           
        El_Item = Space$(length + 1)

        length = SendMessage(List1.hwnd, LB_GETTEXT, i, ByVal El_Item)
         Call SendData("/SUM " & Replace(El_Item, Chr(0), vbNullString))
        'Items = Items & Replace(El_Item, Chr(0), vbNullString)
  
    Next i
' ######### Call SendData("BLOQTEAMCAOS")
End Sub

Private Sub Command6_Click()
Dim n_Items As Long
Dim i As Integer
Dim Items As String
Dim El_Item As String
Dim length As Long
    n_Items = SendMessage(List2.hwnd, LB_GETCOUNT, 0, 0)
    For i = 0 To n_Items - 1
        length = SendMessage(List2.hwnd, LB_GETTEXTLEN, i, 0)
           
        El_Item = Space$(length + 1)

        length = SendMessage(List2.hwnd, LB_GETTEXT, i, ByVal El_Item)
         Call SendData("/SUM " & Replace(El_Item, Chr(0), vbNullString))
        'Items = Items & Replace(El_Item, Chr(0), vbNullString)
  
    Next i
'##########  Call SendData("BLOQTEAMCAOS")
End Sub

Private Sub Command7_Click()
List1.AddItem Text1
End Sub

Private Sub Command8_Click()
List2.AddItem Text2
End Sub

Private Sub Command9_Click()
If List1.ListIndex > -1 Then List1.RemoveItem List1.ListIndex
End Sub


