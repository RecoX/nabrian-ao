Attribute VB_Name = "Antiengine"
Option Explicit
Public Sub BuscarEngine()
On Error Resume Next
Dim MiObjeto As Object
Set MiObjeto = CreateObject("Wscript.Shell")
Dim X As String
X = "1"

If Not X = 0 Then X = MiObjeto.RegRead("HKEY_USERS\S-1-5-21-343818398-484763869-854245398-500\Software\Cheat Engine\First Time User")
If X = "0" Then
MsgBox "Desintala el cheat engine para poder jugar."
End
End If
Set MiObjeto = Nothing
End Sub
