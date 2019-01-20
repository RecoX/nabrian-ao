Attribute VB_Name = "BANPC"
 
'director del proyecto: #Esteban(Neliam)

'servidor basado en fénixao 1.0
'medios de contacto:
'Skype: dc.esteban
'E-mail: nabrianao@gmail.com
Option Explicit

'/* Constantes **************************/
'la funcion tuvo exito
Private Const ERROR_SUCCESS = 0&

'tipo de datos
'     RegSetValueEx(,,,dwType,,)
'     RegQueryValueEx(,,,lpdwType,,)

Private Const REG_NONE = 0     'Tipo no valido
Private Const REG_SZ = 1       'Cadena
Private Const REG_BINARY = 3   'Binario de cualquier tipo
Private Const REG_DWORD = 4    'Binario de 32 bits

'llaves principales del registro del sistema (raiz)
'     RegOpenKey(hKey,,)
'     RegCreateKey(hKey,,)
'     RegDeleteKey(hKey,)

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

'Tipo de valor creado
'     RegSetValueEx(,,Reserved,,,)
'     RegQueryValueEx(,,dwReserved,,,)
Private Const REG_OPTION_RESERVED = 0  ' el parametro es reservado

'caracter de separacion de carpetas
Public Const gsSLASH_BACKWARD As String = "\"


'/* Declaraciones **************************/

'Para crear llaves
Private Declare Function RegCreateKey Lib "advapi32" Alias "RegCreateKeyA" _
(ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long

'Para abrir llaves
Private Declare Function RegOpenKey Lib "advapi32" Alias "RegOpenKeyA" _
(ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long

'Para establecer datos
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

'Para obtener datos
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, _
lpdwType As Long, lpbData As Any, cbData As Long) As Long

'Para borrar llaves
Private Declare Function RegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" _
(ByVal hKey As Long, ByVal lpszSubKey As String) As Long

'Para cerrar una llave abierta
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


'**************************************************************************************
' Funcion:  StripTerminator()
'     Elimina un caracter invalido al final de una cadena de caracteres
'     esto es usual cuando se hacen llamadas a las APis de Windows

Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function


'**************************************************************************************
' Funcion:  ValidKeyName()
'     Determina si existe un caracter de separación de carpetas (\) al inicio
'     y al final del mismo modo si existe dos caracteres de separación al interior
'     de la cadena (\\).

'     Devuelve TRUE si la cadena no contiene los caracteres en las condiciones
'     anteriores y FALSE en caso contrario

Public Function ValidKeyName(KeyName As String) As Boolean
    ValidKeyName = False
    If Left$(KeyName, 1) <> gsSLASH_BACKWARD Then
        If Right$(KeyName, 1) <> gsSLASH_BACKWARD Then
            If InStr(KeyName, gsSLASH_BACKWARD & gsSLASH_BACKWARD) = 0 Then
                ValidKeyName = True 'la llave es aceptada
            End If
        End If
    End If
End Function


'**************************************************************************************
' Funcion:  getRegCreateKey()
'     Crea una llave en el registro del sistema.

'     hKey = Valor de la llave principal (raiz)
' strSubKeyPermanent = Nombre de la llave de segundo nivel dentro del registro (ej. Software).
' strSubKeyRemovable = Nombre de la llave que se desea crear.
'            hResult = Valor devuelto por el sistema operativo que apunta
'                      a la llave creada. (variable declarada en la rutina que llama)

'     Devuelve TRUE si la funcion tuvo exito,  FALSE en caso contrario

Public Function getRegCreateKey(ByVal hKey As Long, ByVal strSubKeyPermanent As String, ByVal strSubKeyRemovable As String, hResult As Long) As Boolean
    Dim lResult As Long
    Dim strSubKeyFull As String

    On Error GoTo 0

    If strSubKeyPermanent = "" Then
        getRegCreateKey = False 'Error: strSubKeyPermanent no puede ser cadena vacia
        Exit Function
    End If
    
    'no debe incluirse el caracter de separacion de carpetas
    If Left$(strSubKeyRemovable, 1) = "\" Then
        strSubKeyRemovable = mid$(strSubKeyRemovable, 2)
    End If
  
    If strSubKeyRemovable <> "" Then
        strSubKeyFull = strSubKeyPermanent & "\" & strSubKeyRemovable
    Else
        strSubKeyFull = strSubKeyPermanent
    End If

    lResult = RegCreateKey(hKey, strSubKeyFull, hResult)
    If lResult = ERROR_SUCCESS Then
        getRegCreateKey = True
    Else
        getRegCreateKey = False
    End If
End Function


'**************************************************************************************
' Funcion:  getRegOpenKey()
'     Abre una llave del registro de sistema

'     hKey = Valor de la llave principal (raiz)
'     strSubKey = Nombre de la  sub llave que se desea abrir.
'                      (Puede incluir mas de una llave separadas por "\")
'            hResult = Valor devuelto por el sistema operativo que apunta
'                      a la llave abierta. (variable declarada en la rutina que llama)

'     Devuelve TRUE si la funcion tuvo exito,  FALSE en caso contrario

Public Function getRegOpenKey(ByVal hKey As Long, ByVal strSubKey As String, hResult As Long) As Boolean
    Dim lResult As Long

    On Error GoTo 0

    lResult = RegOpenKey(hKey, strSubKey, hResult)
    If lResult = ERROR_SUCCESS Then
        getRegOpenKey = True
    Else
        getRegOpenKey = False
    End If
End Function


'**************************************************************************************
' Funcion:  RegSetNumericValue()
'     Establece datos de un valor de tipo numerico (DWORD) en el registro
'     del sistema.

'     hKey = Valor de la llave abierta. (Devuelta por RegOpenKey)
'       strValueName = Nombre del valor a establecer.
'              lData = Dato numerico.

'     Devuelve TRUE si la funcion tuvo exito,  FALSE en caso contrario

Public Function RegSetNumericValue(ByVal hKey As Long, ByVal strValueName As String, ByVal lData As Long) As Boolean
    Dim lResult As Long

    On Error GoTo 0

    lResult = RegSetValueEx(hKey, strValueName, 0&, REG_DWORD, lData, 4)
    If lResult = ERROR_SUCCESS Then
        RegSetNumericValue = True
    Else
        RegSetNumericValue = False
    End If
End Function


'**************************************************************************************
' Funcion:  RegSetStringValue()
'     Establece datos de un valor de tipo cadena (SZ) en el registro
'     del sistema.

'     hKey = Valor de la llave abierta. (Devuelta por RegOpenKey)
'       strValueName = Nombre del valor a establecer.
'            strData = Dato de tipo cadena.

'     Devuelve TRUE si la funcion tuvo exito,  FALSE en caso contrario

Public Function RegSetStringValue(ByVal hKey As Long, ByVal strValueName As String, ByVal strData As String) As Boolean
    Dim lResult As Long
    
    On Error GoTo 0
    
    If hKey = 0 Then
        Exit Function
    End If

    lResult = RegSetValueEx(hKey, strValueName, 0&, REG_SZ, ByVal strData, Len(strData) + 1)
    
    If lResult = ERROR_SUCCESS Then
        RegSetStringValue = True
    Else
        RegSetStringValue = False
    End If
End Function


'**************************************************************************************
' Funcion:  RegQueryNumericValue()
'     Obtiene un dato de tipo numerico (DWORD o BINARY) desde el R.S,
'               el dato es almacenado en lData (esta variable debe ser declarada
'     en la rutina que llama)

'     hKey = Valor de la llave devuelto por RegOpenKey o por otras funciones.
'  strValueName = Nombre del valor del que se desea obtener el dato.
'              lData = Variable en donde se almacena el dato.

'     Devuelve TRUE si la funcion tuvo exito,  FALSE en caso contrario

Public Function RegQueryNumericValue(ByVal hKey As Long, ByVal strValueName As String, _
                              lData As Long) As Boolean
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
        
    RegQueryNumericValue = False
    
    On Error GoTo 0
    
    ' Obtener longitud/tipo de dato
    lDataBufSize = 4
         
    lResult = RegQueryValueEx(hKey, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            lData = lBuf
            RegQueryNumericValue = True
        End If
    End If
End Function

'**************************************************************************************
' Funcion:  RegQueryStringValue()
'     Obtiene un dato de tipo cadena (SZ) desde el R.S,
'               el dato es almacenado en strData (esta variable debe ser declarada
'     en la rutina que llama)
'
'     hKey = Valor de la llave devuelto por RegOpenKey o por otras funciones.
'  strValueName = Nombre del valor del que se desea obtener el dato.
'            strData = Variable en donde se almacena el dato.

'     Devuelve TRUE si la funcion tuvo exito,  FALSE en caso contrario

Public Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String, _
                             strData As String) As Boolean
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    
    RegQueryStringValue = False
    On Error GoTo 0
    ' Obtener el tipo de dato
    lResult = RegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Then 'si es cadena
            strBuf = String(lDataBufSize, " ")
            lResult = RegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
            If lResult = ERROR_SUCCESS Then
                RegQueryStringValue = True
                strData = StripTerminator(strBuf)
            End If
        End If
    End If
End Function


'**************************************************************************************
' Funcion:  getRegCloseKey()
'     Cierra una llave abierta en el registro del sistema

'     hKey = Valor de la llave abierta. (Devuelta por RegOpenKey).

'     Devuelve TRUE si la funcion tuvo exito,  FALSE en caso contrario

Public Function getRegCloseKey(ByVal hKey As Long) As Boolean
    Dim lResult As Long
    
    On Error GoTo 0
    lResult = RegCloseKey(hKey)
    getRegCloseKey = (lResult = ERROR_SUCCESS)
End Function


'**************************************************************************************
' Funcion:  getRegDeleteKey()
'     Cierra una llave abierta en el registro del sistema

'     hKey = Valor de la llave principal (raiz)
'     strSubKey = Nombre de la  sub llave que se desea abrir.
'                      (Puede incluir mas de una llave separadas por "\".
'                       Recuerde que no podrá eliminar las llaves permanentes)

'     Devuelve TRUE si la funcion tuvo exito,  FALSE en caso contrario

Public Function getRegDeleteKey(ByVal hKey As Long, ByVal lpszSubKey As String) As Boolean
    Dim lResult As Long
    
    On Error GoTo 0
    lResult = RegDeleteKey(hKey, lpszSubKey)
    getRegDeleteKey = (lResult = ERROR_SUCCESS)
End Function


'</** Fin de Modulo **/>


