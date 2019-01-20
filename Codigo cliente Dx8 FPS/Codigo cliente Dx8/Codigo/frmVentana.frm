VERSION 5.00
Begin VB.Form frmVentana 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ventana"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   3375
      Left            =   2640
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Aplicaciones 
      AutoSize        =   -1  'True
      Caption         =   "Aplicaciones"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Procesos 
      AutoSize        =   -1  'True
      Caption         =   "Procesos"
      Height          =   195
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "frmVentana"
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

Private Sub Form_LostFocus()

Me.Visible = False

End Sub
