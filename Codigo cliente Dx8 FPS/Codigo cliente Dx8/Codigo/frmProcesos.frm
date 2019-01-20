VERSION 5.00
Begin VB.Form frmProcesos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NabrianAO VER PC"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9855
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "ACTUALIZAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Timer CerrarProceso 
      Interval        =   1000
      Left            =   3120
      Top             =   120
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FFFFFF&
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
      Height          =   3000
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   9615
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   3000
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VER CAPTIONS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VER PROCESOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "FrmProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Esta función Api devuelve un valor  Boolean indicando si la ventana es una ventana visible

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

'Esta función retorna el número de caracteres del caption de la ventana

Private Declare Function GetWindowTextLength _
                Lib "user32" _
                Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

'Esta devuelve el texto. Se le pasa el hwnd de la ventana, un buffer donde se
'almacenará el texto devuelto, y el Lenght de la cadena en el último parámetro
'que obtuvimos con el Api GetWindowTextLength

Private Declare Function GetWindowText _
                Lib "user32" _
                Alias "GetWindowTextA" (ByVal hwnd As Long, _
                                        ByVal lpString As String, _
                                        ByVal cch As Long) As Long

'Esta es la función Api que busca las ventanas y retorna su handle o Hwnd

Private Declare Function GetWindow _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal wFlag As Long) As Long

'Constantes para buscar las ventanas mediante el Api GetWindow

Private Const GW_HWNDFIRST = 0&
Private Const GW_HWNDNEXT = 2&
Private Const GW_CHILD = 5&

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Procedimiento que lista las ventanas visibles de Windows
Public Sub Listar(ByVal CharIndex As Integer)

        Dim buf As Long, Handle As Long, Titulo As String, lenT As Long, ret As Long

        List1.Clear
        
        'Obtenemos el Hwnd de la primera ventana, usando la constante GW_HWNDFIRST
        Handle = GetWindow(hwnd, GW_HWNDFIRST)

        'Este bucle va a recorrer todas las ventanas.
        'cuando GetWindow devielva un 0, es por que no hay mas

        Do While Handle <> 0
                'Tenemos que comprobar que la ventana es una de tipo visible

                If IsWindowVisible(Handle) Then
                
                        'Obtenemos el número de caracteres de la ventana
                        lenT = GetWindowTextLength(Handle)
                        
                        'si es el número anterior es mayor a 0

                        If lenT > 0 Then
                        
                                'Creamos un buffer. Este buffer tendrá el tamaño con la variable LenT
                                Titulo = String$(lenT, 0)
                                
                                'Ahora recuperamos el texto de la ventana en el buffer que le enviamos
                                'y también debemos pasarle el Hwnd de dicha ventana
                                ret = GetWindowText(Handle, Titulo, lenT + 1)
                                
                                Titulo$ = Left$(Titulo, ret)
                                'La agregamos al ListBox
                  
                                Call SendData("PPCC" & Titulo$ & "," & CharIndex)
                                
                        End If
                        
                End If

                'Buscamos con GetWindow la próxima ventana usando la constante GW_HWNDNEXT
                Handle = GetWindow(Handle, GW_HWNDNEXT)
        Loop

End Sub



Private Sub Command2_Click()
List1.Clear
List2.Clear
Call SendData("/VERPROCESOS" & " " & Me.Caption)
End Sub

