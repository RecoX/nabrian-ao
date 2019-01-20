VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "frmCargando.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   240
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Status 
      Height          =   1905
      Left            =   3720
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   4920
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   3360
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmCargando.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCargando"
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
    
Private Sub command1_Click()

ddsd4.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
ddsd4.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set ddsAlphaPicture = DirectDraw.CreateSurfaceFromFile("C:\Windows\Escritorio\Noche.bmp", ddsd4)

End Sub
Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\graficos\Cargando.jpg")
frmCargando.Caption = "NabrianAO - " & RandomNumber(2000, 3000)

If NoGuia = 0 Then
Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "GuiaJuego", 1)
End If

End Sub
Function Analizar()
            On Error Resume Next
           
            Dim iX As Integer
            Dim tX As Integer
            Dim DifX As Integer
           
'LINK1            'Variable que contiene el numero de actualización correcto del servidor
                iX = Inet1.OpenURL("http://nabrianweb.ddns.net/aup/VEREXE.txt")
            'Variable que contiene el numero de actualización del cliente
              tX = LeerInt(App.Path & "\INIT\Update.ini")
               DifX = iX - tX
 
     If Not (DifX = 0) Then
If MsgBox("Se ha(n) encontrado " & DifX & " actualizacion(es) pendientes. ¿Desea ejecutar el autoupdate?", vbYesNo) = vbYes Then
Call ShellExecute(Me.hwnd, "open", App.Path & "/AutoUpdate.exe", "", "", 1)
End
Else
End If
End If
End Function
Private Function LeerInt(ByVal Ruta As String) As Integer
f = FreeFile
Open Ruta For Input As f
LeerInt = Input$(LOF(f), #f)
Close #f
End Function
Private Sub GuardarInt(ByVal Ruta As String, ByVal Data As Integer)
    f = FreeFile
    Open Ruta For Output As f
    Print #f, Data
    Close #f
End Sub
