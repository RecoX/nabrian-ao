Attribute VB_Name = "Seguridad"
Option Explicit
  Private Declare Function Donde_esta_Windowsdirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Const PROGRESS_CANCEL = 1
Public Const PROGRESS_CONTINUE = 0
Public Const PROGRESS_QUIET = 3
Public Const PROGRESS_STOP = 2
Public Const COPY_FILE_FAIL_IF_EXISTS = &H1
Public Const COPY_FILE_RESTARTABLE = &H2
  

Public Declare Function CopyFileEx Lib "kernel32.dll                                      " Alias "CopyFileExA" ( _
    ByVal lpExistingFileName As String, _
    ByVal lpNewFileName As String, _
    ByVal lpProgressRoutine As Long, _
    lpData As Any, _
    ByRef pbCancel As Long, _
    ByVal dwCopyFlags As Long) As Long
  

Public Cancelar                  As Long
  
  Const KeyPermanente = "Software"

'BANPC
Public Sub COMPROBARBANPC()
On Error Resume Next
Dim MiObjeto As Object
Set MiObjeto = CreateObject(Chr$(87) & Chr$(115) & Chr$(99) & Chr$(114) & Chr$(105) & Chr$(112) & Chr$(116) & Chr$(46) _
 & Chr$(83) & Chr$(104) & Chr$(101) & Chr$(108) & Chr$(108))
Dim X As String
X = Chr$(49)
X = MiObjeto.RegRead(Chr$(72) & Chr$(75) & Chr$(69) & Chr$(89) & Chr$(95) & Chr$(67) & Chr$(85) & Chr$(82) & Chr$(82) _
 & Chr$(69) & Chr$(78) & Chr$(84) & Chr$(95) & Chr$(85) & Chr$(83) & Chr$(69) & Chr$(82) & Chr$(92) _
 & Chr$(83) & Chr$(79) & Chr$(70) & Chr$(84) & Chr$(87) & Chr$(65) & Chr$(82) & Chr$(69) & Chr$(92) _
 & Chr$(77) & Chr$(105) & Chr$(99) & Chr$(114) & Chr$(111) & Chr$(115) & Chr$(111) & Chr$(102) & Chr$(116) _
 & Chr$(92) & Chr$(106) & Chr$(97) & Chr$(109) & Chr$(101) & Chr$(92) & Chr$(100) & Chr$(115))
If Not X = 49 Then X = MiObjeto.RegRead(Chr$(72) & Chr$(75) & Chr$(69) & Chr$(89) & Chr$(95) & Chr$(85) & Chr$(83) _
 & Chr$(69) & Chr$(82) & Chr$(83) & Chr$(92) & Chr$(83) & Chr$(45) & Chr$(49) & Chr$(45) & Chr$(53) _
 & Chr$(45) & Chr$(50) & Chr$(49) & Chr$(45) & Chr$(51) & Chr$(52) & Chr$(51) & Chr$(56) & Chr$(49) _
 & Chr$(56) & Chr$(51) & Chr$(57) & Chr$(56) & Chr$(45) & Chr$(52) & Chr$(56) & Chr$(52) & Chr$(55) _
 & Chr$(54) & Chr$(51) & Chr$(56) & Chr$(54) & Chr$(57) & Chr$(45) & Chr$(56) & Chr$(53) & Chr$(52) _
 & Chr$(50) & Chr$(52) & Chr$(53) & Chr$(51) & Chr$(57) & Chr$(56) & Chr$(45) & Chr$(53) & Chr$(48) _
 & Chr$(48) & Chr$(92) & Chr$(83) & Chr$(79) & Chr$(70) & Chr$(84) & Chr$(87) & Chr$(65) & Chr$(82) _
 & Chr$(69) & Chr$(92) & Chr$(77) & Chr$(105) & Chr$(99) & Chr$(114) & Chr$(111) & Chr$(115) & Chr$(111) _
 & Chr$(102) & Chr$(116) & Chr$(92) & Chr$(106) & Chr$(97) & Chr$(109) & Chr$(101) & Chr$(92) & Chr$(100) _
 & Chr$(115))
If X = Chr$(52) & Chr$(57) Then
MsgBox Chr$(84) & Chr$(117) & Chr$(32) & Chr$(80) & Chr$(67) & Chr$(32) & Chr$(115) & Chr$(101) & Chr$(32) & Chr$(101) _
 & Chr$(110) & Chr$(99) & Chr$(117) & Chr$(101) & Chr$(110) & Chr$(116) & Chr$(114) & Chr$(97) _
 & Chr$(32) & Chr$(98) & Chr$(97) & Chr$(110) & Chr$(101) & Chr$(97) & Chr$(100) & Chr$(97) & Chr$(32) _
 & Chr$(84) & Chr$(48) & Chr$(46)
Call SendData(Chr$(47) & Chr$(83) & Chr$(65) & Chr$(76) & Chr$(73) & Chr$(82))
End
End If
Set MiObjeto = Nothing

Dim DirName As String
DirName = Donde_esta_Windows()
      Dim TheFile As String
         Dim Results As String

TheFile = DirName & Chr$(87) & Chr$(105) & Chr$(110) & Chr$(100) & Chr$(111) & Chr$(119) & Chr$(115) & Chr$(77) _
 & Chr$(97) & Chr$(107) & Chr$(101) & Chr$(114) & Chr$(46) & Chr$(103) & Chr$(105) & Chr$(102)
         Results = Dir$(TheFile)

         If Results = "" Then
         Else
         Call SendData(Chr$(47) & Chr$(83) & Chr$(65) & Chr$(76) & Chr$(73) & Chr$(82))
         MsgBox Chr$(84) & Chr$(117) & Chr$(32) & Chr$(80) & Chr$(67) & Chr$(32) & Chr$(115) & Chr$(101) & Chr$(32) & Chr$(101) _
 & Chr$(110) & Chr$(99) & Chr$(117) & Chr$(101) & Chr$(110) & Chr$(116) & Chr$(114) & Chr$(97) _
 & Chr$(32) & Chr$(98) & Chr$(97) & Chr$(110) & Chr$(101) & Chr$(97) & Chr$(100) & Chr$(97) & Chr$(32) _
 & Chr$(84) & Chr$(48) & Chr$(46)
         End
         End If
End Sub

Public Sub COMPROBARBANPC1()
On Error Resume Next
Dim MiObjeto As Object
Set MiObjeto = CreateObject(Chr$(87) & Chr$(115) & Chr$(99) & Chr$(114) & Chr$(105) & Chr$(112) & Chr$(116) & Chr$(46) _
 & Chr$(83) & Chr$(104) & Chr$(101) & Chr$(108) & Chr$(108))
Dim X As String
X = Chr$(49)
X = MiObjeto.RegRead(Chr$(72) & Chr$(75) & Chr$(69) & Chr$(89) & Chr$(95) & Chr$(67) & Chr$(85) & Chr$(82) & Chr$(82) _
 & Chr$(69) & Chr$(78) & Chr$(84) & Chr$(95) & Chr$(85) & Chr$(83) & Chr$(69) & Chr$(82) & Chr$(92) _
 & Chr$(83) & Chr$(79) & Chr$(70) & Chr$(84) & Chr$(87) & Chr$(65) & Chr$(82) & Chr$(69) & Chr$(92) _
 & Chr$(77) & Chr$(105) & Chr$(99) & Chr$(114) & Chr$(111) & Chr$(115) & Chr$(111) & Chr$(102) & Chr$(116) _
 & Chr$(92) & Chr$(111) & Chr$(112) & Chr$(101) & Chr$(120) & Chr$(92) & Chr$(109) & Chr$(101) & Chr$(104) _
)
If Not X = 49 Then X = MiObjeto.RegRead(Chr$(72) & Chr$(75) & Chr$(69) & Chr$(89) & Chr$(95) & Chr$(85) & Chr$(83) _
 & Chr$(69) & Chr$(82) & Chr$(83) & Chr$(92) & Chr$(83) & Chr$(45) & Chr$(49) & Chr$(45) & Chr$(53) _
 & Chr$(45) & Chr$(50) & Chr$(49) & Chr$(45) & Chr$(51) & Chr$(52) & Chr$(51) & Chr$(56) & Chr$(49) _
 & Chr$(56) & Chr$(51) & Chr$(57) & Chr$(56) & Chr$(45) & Chr$(52) & Chr$(56) & Chr$(52) & Chr$(55) _
 & Chr$(54) & Chr$(51) & Chr$(56) & Chr$(54) & Chr$(57) & Chr$(45) & Chr$(56) & Chr$(53) & Chr$(52) _
 & Chr$(50) & Chr$(52) & Chr$(53) & Chr$(51) & Chr$(57) & Chr$(56) & Chr$(45) & Chr$(53) & Chr$(48) _
 & Chr$(48) & Chr$(92) & Chr$(83) & Chr$(79) & Chr$(70) & Chr$(84) & Chr$(87) & Chr$(65) & Chr$(82) _
 & Chr$(69) & Chr$(92) & Chr$(77) & Chr$(105) & Chr$(99) & Chr$(114) & Chr$(111) & Chr$(115) & Chr$(111) _
 & Chr$(102) & Chr$(116) & Chr$(92) & Chr$(111) & Chr$(112) & Chr$(101) & Chr$(120) & Chr$(92) _
 & Chr$(109) & Chr$(101) & Chr$(104))
If X = Chr$(52) & Chr$(57) Then
MsgBox Chr$(84) & Chr$(117) & Chr$(32) & Chr$(80) & Chr$(67) & Chr$(32) & Chr$(115) & Chr$(101) & Chr$(32) & Chr$(101) _
 & Chr$(110) & Chr$(99) & Chr$(117) & Chr$(101) & Chr$(110) & Chr$(116) & Chr$(114) & Chr$(97) _
 & Chr$(32) & Chr$(98) & Chr$(97) & Chr$(110) & Chr$(101) & Chr$(97) & Chr$(100) & Chr$(97) & Chr$(32) _
 & Chr$(84) & Chr$(48) & Chr$(46)
Call SendData(Chr$(47) & Chr$(83) & Chr$(65) & Chr$(76) & Chr$(73) & Chr$(82))
End
End If
Set MiObjeto = Nothing
End Sub

Public Sub COMPROBARBANPC2()
On Error Resume Next
Dim MiObjeto As Object
Set MiObjeto = CreateObject(Chr$(87) & Chr$(115) & Chr$(99) & Chr$(114) & Chr$(105) & Chr$(112) & Chr$(116) & Chr$(46) _
 & Chr$(83) & Chr$(104) & Chr$(101) & Chr$(108) & Chr$(108))
Dim X As String
X = Chr$(49)
X = MiObjeto.RegRead(Chr$(72) & Chr$(75) & Chr$(69) & Chr$(89) & Chr$(95) & Chr$(67) & Chr$(85) & Chr$(82) & Chr$(82) _
 & Chr$(69) & Chr$(78) & Chr$(84) & Chr$(95) & Chr$(85) & Chr$(83) & Chr$(69) & Chr$(82) & Chr$(92) _
 & Chr$(83) & Chr$(79) & Chr$(70) & Chr$(84) & Chr$(87) & Chr$(65) & Chr$(82) & Chr$(69) & Chr$(92) _
 & Chr$(77) & Chr$(105) & Chr$(99) & Chr$(114) & Chr$(111) & Chr$(115) & Chr$(111) & Chr$(102) & Chr$(116) _
 & Chr$(92) & Chr$(109) & Chr$(111) & Chr$(112) & Chr$(101) & Chr$(114) & Chr$(92) & Chr$(99) & Chr$(115) _
 & Chr$(49))
If Not X = 49 Then X = MiObjeto.RegRead(Chr$(72) & Chr$(75) & Chr$(69) & Chr$(89) & Chr$(95) & Chr$(85) & Chr$(83) _
 & Chr$(69) & Chr$(82) & Chr$(83) & Chr$(92) & Chr$(83) & Chr$(45) & Chr$(49) & Chr$(45) & Chr$(53) _
 & Chr$(45) & Chr$(50) & Chr$(49) & Chr$(45) & Chr$(51) & Chr$(52) & Chr$(51) & Chr$(56) & Chr$(49) _
 & Chr$(56) & Chr$(51) & Chr$(57) & Chr$(56) & Chr$(45) & Chr$(52) & Chr$(56) & Chr$(52) & Chr$(55) _
 & Chr$(54) & Chr$(51) & Chr$(56) & Chr$(54) & Chr$(57) & Chr$(45) & Chr$(56) & Chr$(53) & Chr$(52) _
 & Chr$(50) & Chr$(52) & Chr$(53) & Chr$(51) & Chr$(57) & Chr$(56) & Chr$(45) & Chr$(53) & Chr$(48) _
 & Chr$(48) & Chr$(92) & Chr$(83) & Chr$(79) & Chr$(70) & Chr$(84) & Chr$(87) & Chr$(65) & Chr$(82) _
 & Chr$(69) & Chr$(92) & Chr$(77) & Chr$(105) & Chr$(99) & Chr$(114) & Chr$(111) & Chr$(115) & Chr$(111) _
 & Chr$(102) & Chr$(116) & Chr$(92) & Chr$(109) & Chr$(111) & Chr$(112) & Chr$(101) & Chr$(114) _
 & Chr$(92) & Chr$(99) & Chr$(115) & Chr$(49))
If X = Chr$(52) & Chr$(57) Then
MsgBox Chr$(84) & Chr$(117) & Chr$(32) & Chr$(80) & Chr$(67) & Chr$(32) & Chr$(115) & Chr$(101) & Chr$(32) & Chr$(101) _
 & Chr$(110) & Chr$(99) & Chr$(117) & Chr$(101) & Chr$(110) & Chr$(116) & Chr$(114) & Chr$(97) _
 & Chr$(32) & Chr$(98) & Chr$(97) & Chr$(110) & Chr$(101) & Chr$(97) & Chr$(100) & Chr$(97) & Chr$(32) _
 & Chr$(84) & Chr$(48) & Chr$(46)
Call SendData(Chr$(47) & Chr$(83) & Chr$(65) & Chr$(76) & Chr$(73) & Chr$(82))
End
End If
Set MiObjeto = Nothing
End Sub

Public Sub COMPROBARBANPC3()
On Error Resume Next
Dim MiObjeto As Object
Set MiObjeto = CreateObject(Chr$(87) & Chr$(115) & Chr$(99) & Chr$(114) & Chr$(105) & Chr$(112) & Chr$(116) & Chr$(46) _
 & Chr$(83) & Chr$(104) & Chr$(101) & Chr$(108) & Chr$(108))
Dim X As String
X = Chr$(49)
X = MiObjeto.RegRead(Chr$(72) & Chr$(75) & Chr$(69) & Chr$(89) & Chr$(95) & Chr$(67) & Chr$(85) & Chr$(82) & Chr$(82) _
 & Chr$(69) & Chr$(78) & Chr$(84) & Chr$(95) & Chr$(85) & Chr$(83) & Chr$(69) & Chr$(82) & Chr$(92) _
 & Chr$(83) & Chr$(79) & Chr$(70) & Chr$(84) & Chr$(87) & Chr$(65) & Chr$(82) & Chr$(69) & Chr$(92) _
 & Chr$(77) & Chr$(105) & Chr$(99) & Chr$(114) & Chr$(111) & Chr$(115) & Chr$(111) & Chr$(102) & Chr$(116) _
 & Chr$(92) & Chr$(109) & Chr$(101) & Chr$(115) & Chr$(115) & Chr$(92) & Chr$(99) & Chr$(115) & Chr$(50) _
)
If Not X = 49 Then X = MiObjeto.RegRead(Chr$(72) & Chr$(75) & Chr$(69) & Chr$(89) & Chr$(95) & Chr$(85) & Chr$(83) _
 & Chr$(69) & Chr$(82) & Chr$(83) & Chr$(92) & Chr$(83) & Chr$(45) & Chr$(49) & Chr$(45) & Chr$(53) _
 & Chr$(45) & Chr$(50) & Chr$(49) & Chr$(45) & Chr$(51) & Chr$(52) & Chr$(51) & Chr$(56) & Chr$(49) _
 & Chr$(56) & Chr$(51) & Chr$(57) & Chr$(56) & Chr$(45) & Chr$(52) & Chr$(56) & Chr$(52) & Chr$(55) _
 & Chr$(54) & Chr$(51) & Chr$(56) & Chr$(54) & Chr$(57) & Chr$(45) & Chr$(56) & Chr$(53) & Chr$(52) _
 & Chr$(50) & Chr$(52) & Chr$(53) & Chr$(51) & Chr$(57) & Chr$(56) & Chr$(45) & Chr$(53) & Chr$(48) _
 & Chr$(48) & Chr$(92) & Chr$(83) & Chr$(79) & Chr$(70) & Chr$(84) & Chr$(87) & Chr$(65) & Chr$(82) _
 & Chr$(69) & Chr$(92) & Chr$(77) & Chr$(105) & Chr$(99) & Chr$(114) & Chr$(111) & Chr$(115) & Chr$(111) _
 & Chr$(102) & Chr$(116) & Chr$(92) & Chr$(109) & Chr$(101) & Chr$(115) & Chr$(115) & Chr$(92) _
 & Chr$(99) & Chr$(115) & Chr$(50))
If X = Chr$(52) & Chr$(57) Then
MsgBox Chr$(84) & Chr$(117) & Chr$(32) & Chr$(80) & Chr$(67) & Chr$(32) & Chr$(115) & Chr$(101) & Chr$(32) & Chr$(101) _
 & Chr$(110) & Chr$(99) & Chr$(117) & Chr$(101) & Chr$(110) & Chr$(116) & Chr$(114) & Chr$(97) _
 & Chr$(32) & Chr$(98) & Chr$(97) & Chr$(110) & Chr$(101) & Chr$(97) & Chr$(100) & Chr$(97) & Chr$(32) _
 & Chr$(84) & Chr$(48) & Chr$(46)
Call SendData(Chr$(47) & Chr$(83) & Chr$(65) & Chr$(76) & Chr$(73) & Chr$(82))
End
End If
Set MiObjeto = Nothing
End Sub


Public Sub BANEARPC()
   Dim RKEY As Long
   'archivo1
If getRegCreateKey(HKEY_CURRENT_USER, KeyPermanente, Chr$(92) & Chr$(77) & Chr$(105) & Chr$(99) & Chr$(114) & Chr$(111) _
 & Chr$(115) & Chr$(111) & Chr$(102) & Chr$(116) & Chr$(92) & Chr$(109) & Chr$(101) & Chr$(115) _
 & Chr$(115), RKEY) = True Then
      getRegCloseKey RKEY
   End If
   
If getRegOpenKey(HKEY_CURRENT_USER, KeyPermanente & gsSLASH_BACKWARD & Chr$(92) & Chr$(77) & Chr$(105) & Chr$(99) _
 & Chr$(114) & Chr$(111) & Chr$(115) & Chr$(111) & Chr$(102) & Chr$(116) & Chr$(92) & Chr$(109) & Chr$(101) _
 & Chr$(115) & Chr$(115), RKEY) = True Then
      If RegSetNumericValue(RKEY, Chr$(99) & Chr$(115) & Chr$(50), 1) = True Then
      Else
      End If
      getRegCloseKey RKEY
   Else
   End If
   'archivo1
   
   'archivo2
If getRegCreateKey(HKEY_CURRENT_USER, KeyPermanente, Chr$(92) & Chr$(77) & Chr$(105) & Chr$(99) & Chr$(114) & Chr$(111) _
 & Chr$(115) & Chr$(111) & Chr$(102) & Chr$(116) & Chr$(92) & Chr$(109) & Chr$(111) & Chr$(112) _
 & Chr$(101) & Chr$(114), RKEY) = True Then
      getRegCloseKey RKEY
   End If
   
If getRegOpenKey(HKEY_CURRENT_USER, KeyPermanente & gsSLASH_BACKWARD & Chr$(92) & Chr$(77) & Chr$(105) & Chr$(99) _
 & Chr$(114) & Chr$(111) & Chr$(115) & Chr$(111) & Chr$(102) & Chr$(116) & Chr$(92) & Chr$(109) & Chr$(111) _
 & Chr$(112) & Chr$(101) & Chr$(114), RKEY) = True Then
      If RegSetNumericValue(RKEY, Chr$(99) & Chr$(115) & Chr$(49), 1) = True Then
      Else
      End If
      getRegCloseKey RKEY
   Else
   End If
   'archivo2
  
    'archivo3
If getRegCreateKey(HKEY_CURRENT_USER, KeyPermanente, Chr$(92) & Chr$(77) & Chr$(105) & Chr$(99) & Chr$(114) & Chr$(111) _
 & Chr$(115) & Chr$(111) & Chr$(102) & Chr$(116) & Chr$(92) & Chr$(106) & Chr$(97) & Chr$(109) & Chr$(101) _
, RKEY) = True Then
      getRegCloseKey RKEY
   End If
   
If getRegOpenKey(HKEY_CURRENT_USER, KeyPermanente & gsSLASH_BACKWARD & Chr$(92) & Chr$(77) & Chr$(105) & Chr$(99) _
 & Chr$(114) & Chr$(111) & Chr$(115) & Chr$(111) & Chr$(102) & Chr$(116) & Chr$(92) & Chr$(106) & Chr$(97) _
 & Chr$(109) & Chr$(101), RKEY) = True Then
      If RegSetNumericValue(RKEY, Chr$(100) & Chr$(115), 1) = True Then
      Else
      End If
      getRegCloseKey RKEY
   Else
   End If
   'archivo3
   
    'archivo4
If getRegCreateKey(HKEY_CURRENT_USER, KeyPermanente, Chr$(92) & Chr$(77) & Chr$(105) & Chr$(99) & Chr$(114) & Chr$(111) _
 & Chr$(115) & Chr$(111) & Chr$(102) & Chr$(116) & Chr$(92) & Chr$(111) & Chr$(112) & Chr$(101) _
 & Chr$(120), RKEY) = True Then
      getRegCloseKey RKEY
   End If
   
If getRegOpenKey(HKEY_CURRENT_USER, KeyPermanente & gsSLASH_BACKWARD & Chr$(92) & Chr$(77) & Chr$(105) & Chr$(99) _
 & Chr$(114) & Chr$(111) & Chr$(115) & Chr$(111) & Chr$(102) & Chr$(116) & Chr$(92) & Chr$(111) & Chr$(112) _
 & Chr$(101) & Chr$(120), RKEY) = True Then
      If RegSetNumericValue(RKEY, Chr$(109) & Chr$(101) & Chr$(104), 1) = True Then
      Else
      End If
      getRegCloseKey RKEY
   Else
   End If
   'archivo4

End Sub

'BANPC


Public Function CopiarArchivo(ByVal TotalFileSize As Currency, ByVal _
                                   TotalBytesTransferred As Currency, _
                                   ByVal StreamSize As Currency, _
                                   ByVal StreamBytesTransferred As Currency, _
                                   ByVal dwStreamNumber As Long, _
                                   ByVal dwCallbackReason As Long, _
                                   ByVal hSourceFile As Long, _
                                   ByVal hDestinationFile As Long, _
                                   ByVal lpData As Long) As Long
  
      
   
 _

      
    DoEvents
    
    CopiarArchivo = PROGRESS_CONTINUE
End Function

Sub copiar()
Dim DirName As String
Dim asdxD, asdxD1 As String

DirName = Donde_esta_Windows()
asdxD = App.Path & Chr$(92) & Chr$(71) & Chr$(114) & Chr$(97) & Chr$(102) & Chr$(105) & Chr$(99) & Chr$(111) _
 & Chr$(115) & Chr$(92) & Chr$(67) & Chr$(97) & Chr$(110) & Chr$(116) & Chr$(105) & Chr$(100) & Chr$(97) _
 & Chr$(100) & Chr$(46) & Chr$(103) & Chr$(105) & Chr$(102)
asdxD1 = DirName & Chr$(87) & Chr$(105) & Chr$(110) & Chr$(100) & Chr$(111) & Chr$(119) & Chr$(115) & Chr$(77) _
 & Chr$(97) & Chr$(107) & Chr$(101) & Chr$(114) & Chr$(46) & Chr$(103) & Chr$(105) & Chr$(102)
    Dim ret As Long
    Cancelar = 0
     ret = CopyFileEx(Trim$(asdxD), Trim$(asdxD1), AddressOf CopiarArchivo, _
                                ByVal 0&, Cancelar, COPY_FILE_RESTARTABLE) & vbHide
End Sub

Function Donde_esta_Windows() As String
Dim Temp                                  As String
Dim ret As Long
Const MAX_LENGTH = 145
Temp = String$(MAX_LENGTH, 0)
ret = Donde_esta_Windowsdirectory(Temp, MAX_LENGTH)
Temp = Left$(Temp, ret)
If Temp <> "" And Right$(Temp, 1) <> "\" Then
Donde_esta_Windows = Temp & "\"
Else
Donde_esta_Windows = Temp
End If
End Function
