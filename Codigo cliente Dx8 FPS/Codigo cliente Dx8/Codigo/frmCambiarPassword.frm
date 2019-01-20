VERSION 5.00
Begin VB.Form frmCambiarPasswd 
   BorderStyle     =   0  'None
   Caption         =   "Cambiar password"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox ConfirPasswdNuevo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox PasswdNuevo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox PasswdViejo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3360
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Image Cancelar 
      Height          =   375
      Left            =   2880
      MouseIcon       =   "frmCambiarPassword.frx":0000
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Image Aceptar 
      Height          =   375
      Left            =   720
      MouseIcon       =   "frmCambiarPassword.frx":030A
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   1935
   End
End
Attribute VB_Name = "frmCambiarPasswd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FénixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar

Private Sub Aceptar_Click()

If Me.PasswdNuevo <> Me.ConfirPasswdNuevo Then
    Call MsgBox("El password nuevo no coincide con su confirmación.")
    Exit Sub
End If
    
If Len(Me.PasswdNuevo) < 6 Then
    Call AddtoRichTextBox(frmPrincipal.rectxt, "El password nuevo debe tener al menos 6 caracteres.", 65, 190, 156, 0, 0)
    Exit Sub
End If

Call SendData("PASS" & MD5String(Me.PasswdViejo) & "," & MD5String(Me.PasswdNuevo))

Unload Me

End Sub

Private Sub Cancelar_Click()
Unload Me
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Form_Load()

Me.PICTURE = LoadPicture(DirGraficos & "password.gif")

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


