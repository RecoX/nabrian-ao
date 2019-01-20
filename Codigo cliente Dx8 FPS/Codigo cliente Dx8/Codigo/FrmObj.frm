VERSION 5.00
Begin VB.Form FrmObj 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Objetos"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
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
      TabIndex        =   9
      Text            =   "Nombre del obj"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Buscar"
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
      Left            =   2040
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.CheckBox FraseCompleta 
      BackColor       =   &H00000000&
      Caption         =   "Frase Completa"
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
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Crear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "1"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3300
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Numeros de objetos:"
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
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Cantidad"
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
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   4560
      Width           =   735
   End
End
Attribute VB_Name = "FrmObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
' Declaración del api SendMessage
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As String) As Long
 
'Constante "Mensaje" para buscar una cadena en el ListBox
Private Const LB_FINDSTRING = &H18F
Private Const LB_FINDSTRINGEXACT As Long = &H1A2
 
'Recibe la cadena y el valor de tipo boolean para _
 determinar si busca o no la cadena completa
Private Sub Buscar_ListBox(Frase As String)
 
Dim indice As Long
       
    ' Tipo de búsqueda
    If FraseCompleta Then
       indice = SendMessage(List1.hwnd, LB_FINDSTRINGEXACT, -1, Frase)
    Else
       indice = SendMessage(List1.hwnd, LB_FINDSTRING, -1, Frase)
    End If
       
       
    If indice Then
        ' se encontró la frase entonces la selecciona
        List1.ListIndex = indice
    End If
End Sub
 
 
 
 
 
Private Sub command1_Click()
Call SendData("/ITEM " & ReadFieldOptimizado(2, List1, Asc("-")) & " " & Text1.Text)
End Sub
 
Private Sub Command2_Click()
Call SendData("/BUSCAR " & Text3.Text)
End Sub

Private Sub Text2_Change()
Call Buscar_ListBox(Text2)
End Sub

