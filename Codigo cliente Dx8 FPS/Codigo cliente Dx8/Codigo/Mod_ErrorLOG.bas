Attribute VB_Name = "Mod_ErrorLOG"
 
'director del proyecto: #Esteban(Neliam)

'servidor basado en fénixao 1.0
'medios de contacto:
'Skype: dc.esteban
'E-mail: nabrianao@gmail.com
Option Explicit

Public Sub LogError(Desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile
Open App.Path & "\errores.log" For Append As #nfile
Print #nfile, Desc
Close #nfile
End Sub

