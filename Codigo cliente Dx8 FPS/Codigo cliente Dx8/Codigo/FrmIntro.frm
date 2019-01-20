VERSION 5.00
Begin VB.Form FrmIntro 
   BorderStyle     =   0  'None
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   Icon            =   "FrmIntro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   600
      MouseIcon       =   "FrmIntro.frx":F172
      MousePointer    =   99  'Custom
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   600
      MouseIcon       =   "FrmIntro.frx":F47C
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   600
      MouseIcon       =   "FrmIntro.frx":F786
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   600
      MouseIcon       =   "FrmIntro.frx":FA90
      MousePointer    =   99  'Custom
      Top             =   4680
      Width           =   3135
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   600
      MouseIcon       =   "FrmIntro.frx":FD9A
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   3135
   End
End
Attribute VB_Name = "FrmIntro"
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

Private Sub Form_Load()

Me.Picture = LoadPicture(App.Path & "\Graficos\MenuRapido.jpg")

Dim corriendo As Integer
Dim i As Long
Dim proc As PROCESSENTRY32
Dim snap As Long
Dim pepe As String

Dim exeName As String
snap = CreateToolhelpSnapshot(TH32CS_SNAPALL, 0)
proc.dwSize = Len(proc)
theloop = ProcessFirst(snap, proc)
i = 0
While theloop <> 0
    exeName = proc.szexeFile
    Text1.Text = proc.szexeFile
    If Text1.Text = "NabrianAONoDinamicoDx8.exe" Or Text1.Text = "NabrianAODx8.exe" Then
        corriendo = corriendo + 1
        Text1.Text = ""
    End If
    i = i + 1
    theloop = ProcessNext(snap, proc)
Wend
CloseHandle snap

End Sub
Private Sub Image2_Click()
If FindWindow(vbNullString, UCase$("NabrianAO" & " V " & App.Major & "." & App.Minor & "")) Then
    MsgBox "No está permitido el uso de doble cliente", vbExclamation
    End
Else
Call Main
End If
End Sub

Private Sub Image3_Click()
ShellExecute Me.hwnd, "open", App.Path & "/aosetup.exe", "", "", 1
End Sub

Private Sub Image4_Click()
ShellExecute Me.hwnd, "open", "http://www.nabrianao.net", "", "", 1

End Sub

Private Sub Image5_Click()
ShellExecute Me.hwnd, "open", "http://www.nabrianao.net", "", "", 1

End Sub

Private Sub Image6_Click()
Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If bmoving = False And Button = vbLeftButton Then
      Dx3 = x
      dy = y
      bmoving = True
   End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If bmoving And ((x <> Dx3) Or (y <> dy)) Then
      Move Left + (x - Dx3), Top + (y - dy)
   End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      bmoving = False
   End If
End Sub

