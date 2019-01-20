VERSION 5.00
Begin VB.Form RetoClan 
   BorderStyle     =   0  'None
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox ListClanes 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1980
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   2590
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Text            =   "Apuesta"
      Top             =   840
      Width           =   1575
   End
   Begin VB.ListBox ListClanes1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1980
      Left            =   2950
      TabIndex        =   3
      Top             =   1440
      Width           =   1730
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   240
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de clanes con su lider online."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Apostar puntos de canjeo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4600
      TabIndex        =   2
      Top             =   95
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   2520
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "RetoClan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

  
  
Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\graficos\RetoClan.jpg")
End Sub

Private Sub Image1_Click()
If Text1.Text = "Apuesta" Then Text1.Text = 0
If Text1.Text = "" Then Text1.Text = 0
Call SendData("/RETARCLAN " & ListClanes1 & " " & Text1.Text)
Unload Me
End Sub

Private Sub Image2_Click()
 If MsgBox("¿Seguro que deseas aceptar el reto?", vbYesNo, "Reto de clanes") = vbYes Then
        Call SendData("/ACEPTCLAN")
        Unload Me
        Else
        Unload Me
  End If
        
End Sub

Private Sub Label1_Click()
Unload Me
End Sub



Private Sub Listclanes_Click()
    SincListBox ListClanes, ListClanes1
End Sub

Private Sub Listclanes1_Scroll()
    'Sincronizar también el primer item mostrado en la lista
   ListClanes.TopIndex = ListClanes1.TopIndex
End Sub
Private Sub Listclanes_Scroll()
    'Sincronizar también el primer item mostrado en la lista
    ListClanes1.TopIndex = ListClanes.TopIndex
End Sub

Private Sub Listclanes1_Click()
    SincListBox ListClanes1, ListClanes
   ' Label10 = Listclanes1.ListCount
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If bmoving = False And Button = vbLeftButton Then
      Dx3 = X
      dy = Y
      bmoving = True
   End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If bmoving And ((X <> Dx3) Or (Y <> dy)) Then
      Move Left + (X - Dx3), Top + (Y - dy)
   End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      bmoving = False
   End If
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

Private Sub Text1_Click()
If Text1.Text = "Apuesta" Then Text1.Text = ""
End Sub
