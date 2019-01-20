Attribute VB_Name = "Declares"
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal CrcKey As Long, ByVal CrcString As String) As Long

Function GetVar(file As String, Main As String, Var As String) As String
Dim sSpaces As String
  
sSpaces = Space$(5000)
  
getprivateprofilestring Main, Var, "", sSpaces, Len(sSpaces), file

GetVar = RTrim(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function
