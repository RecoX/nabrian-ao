VERSION 5.00
Begin VB.Form frmMandarReto 
   BorderStyle     =   0  'None
   Caption         =   "Sistema de retos"
   ClientHeight    =   6030
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   3960
   Icon            =   "frmMandarReto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "De plante"
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
      Height          =   210
      Left            =   720
      TabIndex        =   6
      Top             =   4320
      Width           =   200
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Sin items de canje"
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
      Height          =   210
      Left            =   720
      TabIndex        =   5
      Top             =   4710
      Width           =   200
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "2vs2"
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
      Height          =   270
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   225
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000F0F0F&
      Caption         =   "1vs1"
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
      Height          =   270
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   2
      Text            =   "Apuesta"
      Top             =   1920
      Width           =   3255
   End
   Begin VB.ComboBox Text4 
      BackColor       =   &H80000011&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Width           =   3255
   End
   Begin VB.ComboBox text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   360
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3480
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmMandarReto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Declaración del Api SendMessageLong
Private Declare Function SendMessageLong _
    Lib "user32" _
    Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long
  
'Flag para la tecla BackSpace
Private KeyRetroceso As Boolean
  

  
'A este procedimento le enviamos como _
parámetro el Control Combo que queremos utilizar.
Public Function Autocompletar_Combo(Combo As ComboBox)
  
Dim i As Integer, posSelect As Integer
  
    Select Case (KeyRetroceso Or Len(Combo.Text) = 0)
        Case True
            KeyRetroceso = False
            Exit Function
    End Select
  
    With Combo
  
    'Recorremos todos los elementos del combo
    For i = 0 To .ListCount - 1
        'Si hay coincidencia
        If InStr(1, .List(i), .Text, vbTextCompare) = 1 Then
            posSelect = .SelStart
            'Mostramos el texto en el combo
            .Text = .List(i)
            'Indicamos el comienzo de la selección
            .SelStart = posSelect
            'Acá seleccionamos el texto
            .SelLength = Len(.Text) - posSelect
  
            Exit For
        End If
    Next i
  
    End With
End Function
  
'Este procedimiento es para ocultar o desplegar _
el combo cuando presionamos el enter

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\graficos\reto.jpg")
End Sub

Private Sub Image1_Click()
If Text1.Text = "Apuesta" Then Text1.Text = 0
If Text1.Text = "" Then Text1.Text = 0

text2.Text = Replace(text2.Text, " ", "+")
If Check2.value = Checked Then
Call SendData("/RETPLE " & text2.Text)
Else
Call SendData("/RETPLA " & text2.Text)
End If

If Option1.value = True Then
If Me.text2 = "" Then
MsgBox "Debes insertar el nick de un personaje."
Exit Sub
End If
If Not (IsNumeric(Text1.Text)) Then
MsgBox "los canjes que ingreses deben ser de caracter numerico"
Exit Sub
End If
Call SendData("/RETO " & text2.Text & " " & Text1.Text)
End If

If Check1.value = Checked Then
Call SendData("/RETCAN " & text2.Text)
Else
Call SendData("/RETCEN " & text2.Text)
End If

If Option2.value = True Then

If Me.Text4 = "" Then
MsgBox "debes escribir el nick de tu pareja."
Exit Sub
End If
If (IsNumeric(Text4.Text)) Then
MsgBox "El nombre que ingreses no puede contener numeros."
Exit Sub
End If
Call SendData("/PAREJA " & Text4.Text)

End If

Unload Me
End Sub

Private Sub Image2_Click()
Unload Me
End Sub


Private Sub Text1_Click()
If Text1.Text = "Apuesta" Then Text1.Text = ""
End Sub

Private Sub Text2_Change()
    'Le pasamos el ComboBox que queremos, en este caso un text2
    Autocompletar_Combo text2
End Sub
  
Private Sub text2_KeyDown(KeyCode As Integer, Shift As Integer)
  
    Select Case KeyCode
        'Si la tecla presionada es Backspace o la tecla Delete
        Case vbKeyBack, vbKeyDelete
            Select Case Len(text2.Text)
                Case Is <> 0
                    KeyRetroceso = True
  
            End Select
    End Select
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
  
Dim resp As Integer
  
    If KeyAscii = 13 Then
        'Si le pasamos a SendMessageLong el valor False lo cierra
        resp = SendMessageLong(text2.hwnd, &H14F, False, 0)
    Else
        'si le pasamos True a SendMessageLong lo adespliega, es decir cuando
        'presionamos una tecla diferente al Enter
        resp = SendMessageLong(text2.hwnd, &H14F, True, 0)
    End If
End Sub

'Este procedimiento es para ocultar o desplegar _
el combo cuando presionamos el enter
Private Sub text4_Change()
    'Le pasamos el ComboBox que queremos, en este caso un text4
    Autocompletar_Combo Text4
End Sub
  
Private Sub text4_KeyDown(KeyCode As Integer, Shift As Integer)
  
    Select Case KeyCode
        'Si la tecla presionada es Backspace o la tecla Delete
        Case vbKeyBack, vbKeyDelete
            Select Case Len(Text4.Text)
                Case Is <> 0
                    KeyRetroceso = True
  
            End Select
    End Select
End Sub
Private Sub text4_KeyPress(KeyAscii As Integer)
  
Dim resp As Integer
  
    If KeyAscii = 13 Then
        'Si le pasamos a SendMessageLong el valor False lo cierra
        resp = SendMessageLong(Text4.hwnd, &H14F, False, 0)
    Else
        'si le pasamos True a SendMessageLong lo adespliega, es decir cuando
        'presionamos una tecla diferente al Enter
        resp = SendMessageLong(Text4.hwnd, &H14F, True, 0)
    End If
End Sub







Private Sub Option1_Click()
Text1.Enabled = True
text2.Enabled = True
Text4.Enabled = False
Text4.BackColor = &H80000011
Text1.BackColor = &H80000005
text2.BackColor = &H80000005
Check1.Enabled = True
Check2.Enabled = True
End Sub

Private Sub Option2_Click()
Text1.Enabled = False
text2.Enabled = False
Text4.Enabled = True
Text1.BackColor = &H80000011
text2.BackColor = &H80000011
Text4.BackColor = &H80000005
Check1.value = False
Check1.Enabled = False
Check2.value = False
Check2.Enabled = False
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
