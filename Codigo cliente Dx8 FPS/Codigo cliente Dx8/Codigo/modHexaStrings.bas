Attribute VB_Name = "modHexaStrings"
Option Explicit
Public Function hexMd52Asc(ByVal MD5 As String) As String
    Dim i As Integer, l As String
    
    MD5 = UCase$(MD5)
    If Len(MD5) Mod 2 = 1 Then MD5 = "0" & MD5
    
    For i = 1 To Len(MD5) \ 2
        l = mid$(MD5, (2 * i) - 1, 2)
        hexMd52Asc = hexMd52Asc & Chr(hexHex2Dec(l))
    Next i
End Function

Public Function hexHex2Dec(ByVal hex As String) As Long
    Dim i As Integer, l As String
    For i = 1 To Len(hex)
        l = mid$(hex, i, 1)
        Select Case l
            Case "A": l = 10
            Case "B": l = 11
            Case "C": l = 12
            Case "D": l = 13
            Case "E": l = 14
            Case "F": l = 15
        End Select
        
        hexHex2Dec = (l * 16 ^ ((Len(hex) - i))) + hexHex2Dec
    Next i
End Function

Public Function txtOffset(ByVal Text As String, ByVal off As Integer) As String
    Dim i As Integer, l As String
    For i = 1 To Len(Text)
        l = mid$(Text, i, 1)
        txtOffset = txtOffset & Chr((Asc(l) + off) Mod 256)
    Next i
End Function

Function Encripta(Text As String, Encriptar As Boolean) As String
 
On Error GoTo a:
 
Dim a() As Integer
Dim b() As Integer
Dim Contraseñas(9) As String
Dim i As Integer
Dim ii As Integer
Dim R As String
Dim CI As Byte
Dim ss As Integer

Contraseñas(0) = Chr$(55) & Chr$(56) & Chr$(50) & Chr$(51) & Chr$(49) & Chr$(55) & Chr$(51) & Chr$(55) & Chr$(49) _
 & Chr$(50) & Chr$(48) & Chr$(55) & Chr$(97) & Chr$(55) & Chr$(56) & Chr$(115) & Chr$(100) & Chr$(55) _
 & Chr$(56) & Chr$(97) & Chr$(115) & Chr$(48) & Chr$(55) & Chr$(56) & Chr$(100) & Chr$(55) & Chr$(56) _
 & Chr$(97) & Chr$(115) & Chr$(48) & Chr$(55) & Chr$(56) & Chr$(57) & Chr$(100) & Chr$(49) & Chr$(50) _
 & Chr$(51) & Chr$(49) & Chr$(50) & Chr$(51)
Contraseñas(1) = Chr$(104) & Chr$(102) & Chr$(104) & Chr$(102) & Chr$(100) & Chr$(98) & Chr$(118) & Chr$(110) _
 & Chr$(118) & Chr$(110) & Chr$(118) & Chr$(99) & Chr$(110) & Chr$(102) & Chr$(103) & Chr$(115) _
 & Chr$(100) & Chr$(102) & Chr$(101) & Chr$(119) & Chr$(114) & Chr$(113) & Chr$(119) & Chr$(101) _
 & Chr$(114) & Chr$(119) & Chr$(113) & Chr$(101) & Chr$(114)
Contraseñas(2) = Chr$(111) & Chr$(117) & Chr$(105) & Chr$(115) & Chr$(97) & Chr$(100) & Chr$(102) & Chr$(117) _
 & Chr$(105) & Chr$(115) & Chr$(97) & Chr$(100) & Chr$(117) & Chr$(105) & Chr$(112) & Chr$(102) _
 & Chr$(97) & Chr$(117) & Chr$(115) & Chr$(100) & Chr$(117) & Chr$(105) & Chr$(102) & Chr$(119) _
 & Chr$(101) & Chr$(114) & Chr$(119) & Chr$(101) & Chr$(114)
Contraseñas(3) = Chr$(99) & Chr$(120) & Chr$(118) & Chr$(98) & Chr$(99) & Chr$(120) & Chr$(118) & Chr$(98) & Chr$(105) _
 & Chr$(111) & Chr$(112) & Chr$(99) & Chr$(120) & Chr$(98) & Chr$(99) & Chr$(120) & Chr$(112) & Chr$(111) _
 & Chr$(105) & Chr$(118) & Chr$(98) & Chr$(100) & Chr$(102) & Chr$(103) & Chr$(114) & Chr$(101) _
 & Chr$(116)
Contraseñas(4) = Chr$(115) & Chr$(100) & Chr$(102) & Chr$(103) & Chr$(100) & Chr$(102) & Chr$(103) & Chr$(105) _
 & Chr$(114) & Chr$(112) & Chr$(111) & Chr$(101) & Chr$(116) & Chr$(101) & Chr$(114) & Chr$(112) _
 & Chr$(116) & Chr$(112) & Chr$(105) & Chr$(111) & Chr$(112) & Chr$(111) & Chr$(105) & Chr$(98) _
 & Chr$(99) & Chr$(118) & Chr$(98) & Chr$(99) & Chr$(118) & Chr$(98) & Chr$(99) & Chr$(118) & Chr$(98) _

Contraseñas(5) = Chr$(102) & Chr$(104) & Chr$(102) & Chr$(103) & Chr$(104) & Chr$(102) & Chr$(104) & Chr$(110) _
 & Chr$(98) & Chr$(110) & Chr$(99) & Chr$(118) & Chr$(110) & Chr$(118) & Chr$(99) & Chr$(110)
Contraseñas(6) = Chr$(115) & Chr$(97) & Chr$(100) & Chr$(97) & Chr$(115) & Chr$(100) & Chr$(113) & Chr$(119) _
 & Chr$(101) & Chr$(113) & Chr$(119) & Chr$(117) & Chr$(101) & Chr$(55) & Chr$(56) & Chr$(55) & Chr$(56) _
 & Chr$(57) & Chr$(56) & Chr$(49) & Chr$(50) & Chr$(51) & Chr$(49) & Chr$(50) & Chr$(51)
Contraseñas(7) = Chr$(103) & Chr$(104) & Chr$(102) & Chr$(103) & Chr$(104) & Chr$(100) & Chr$(104) & Chr$(114) _
 & Chr$(116) & Chr$(121) & Chr$(114) & Chr$(116) & Chr$(121) & Chr$(101) & Chr$(114) & Chr$(121) _
 & Chr$(101) & Chr$(114)
Contraseñas(8) = Chr$(116) & Chr$(101) & Chr$(121) & Chr$(114) & Chr$(116) & Chr$(121) & Chr$(101) & Chr$(114) _
 & Chr$(116) & Chr$(121) & Chr$(101) & Chr$(114) & Chr$(103) & Chr$(104) & Chr$(102) & Chr$(104) _
 & Chr$(100) & Chr$(102) & Chr$(103) & Chr$(104) & Chr$(100) & Chr$(102) & Chr$(104) & Chr$(100) _
 & Chr$(102) & Chr$(104)
Contraseñas(9) = Chr$(114) & Chr$(116) & Chr$(121) & Chr$(114) & Chr$(116) & Chr$(101) & Chr$(114) & Chr$(116) _
 & Chr$(102) & Chr$(98) & Chr$(99) & Chr$(118) & Chr$(98) & Chr$(99) & Chr$(118) & Chr$(98) & Chr$(99) _
 & Chr$(118) & Chr$(98) & Chr$(99) & Chr$(118) & Chr$(98) & Chr$(99) & Chr$(118) & Chr$(98)
 
 
'********* que contraseña hay q usar? *********
If Not Encriptar Then
    CI = Val(Asc(Left(Text, 1))) - 10
    Text = Right(Text, Len(Text) - 1)
End If
'**********************************************
 
'para no llamar a cada rato a la function
ss = Len(Text)
 
'Por las dudas
If ss <= 0 Then Exit Function
 
ReDim a(1 To ss) As Integer
 
    For i = 1 To ss
        a(i) = Asc(mid(Text, i, 1))
    Next i
 
 
    If Encriptar Then
 
        '****** Separamos la Contraseña ******
            CI = RandomNumber(0, 9)
            ReDim b(1 To Len(Contraseñas(CI))) As Integer
 
            For i = 1 To Len(Contraseñas(CI))
                b(i) = Asc(mid(Contraseñas(CI), i, 1))
            Next i
        '*************************************
 
        For i = 1 To ss
            If ii >= UBound(b) Then ii = 0
            ii = ii + 1
            a(i) = a(i) + b(ii)
            If a(i) > 255 Then a(i) = a(i) - 255
            R = R + Chr(a(i))
        Next i
 
        Encripta = Chr(CI + 10) & R
 
    Else
       
    '****** Separamos la Contraseña ******
        ReDim b(1 To Len(Contraseñas(CI))) As Integer
       
        For i = 1 To Len(Contraseñas(CI))
            b(i) = Asc(mid(Contraseñas(CI), i, 1))
        Next i
    '*************************************
       
        For i = 1 To ss
        If ii >= UBound(b) Then ii = 0
            ii = ii + 1
            a(i) = a(i) - b(ii)
            If a(i) < 0 Then
            a(i) = a(i) + 255
            End If
            R = R + Chr(a(i))
        Next i
       
        Encripta = R
   
    End If
 
a:
 
End Function


'ENCRIPTACION
' Text1.Text = THeEnCripTe(Text1.Text, "asdasd")
Function THeEnCripTe(ByVal s As String, ByVal P As String) As String
Dim i As Integer, R As String
Dim C1 As Integer, C2 As Integer
R = ""
If Len(P) > 0 Then
For i = 1 To Len(s)
C1 = Asc(mid(s, i, 1))
If i > Len(P) Then
C2 = Asc(mid(P, i Mod Len(P) + 1, 1))
Else
C2 = Asc(mid(P, i, 1))
End If
C1 = C1 - C2 - 64
If Sgn(C1) = -1 Then C1 = 256 + C1
R = R + Chr(C1)
Next i
Else
R = s
End If
THeEnCripTe = R
End Function
'ENCRIPTACION

'ENCRIPTT

Private Function MamasiTEEX(X As Integer) As String
    If X > 9 Then
        MamasiTEEX = Chr(X + 55)
    Else
        MamasiTEEX = CStr(X)
    End If
End Function
Private Function MoveEltoto(X As String) As Integer
      
    Dim X1 As String
    Dim X2 As String
    Dim Temp As Integer
      
    X1 = mid(X, 1, 1)
    X2 = mid(X, 2, 1)
      
    If IsNumeric(X1) Then
        Temp = 16 * Int(X1)
    Else
        Temp = (Asc(X1) - 55) * 16
    End If
      
    If IsNumeric(X2) Then
        Temp = Temp + Int(X2)
    Else
        Temp = Temp + (Asc(X2) - 55)
    End If
      
    ' retorno
    MoveEltoto = Temp
      
End Function

Function TeEncripTE(DataValue As Variant) As Variant
      
    Dim X As Long
    Dim Temp As String
    Dim HexByte As String
      
    For X = 1 To Len(DataValue) Step 2
          
        HexByte = mid(DataValue, X, 2)
        Temp = Temp & Chr(MoveEltoto(HexByte))
          
    Next X
    ' retorno
    TeEncripTE = Temp
      
End Function
'ENCRIPTT

Public Function Encriptar(sTexto As String) As String
Dim i As Integer
Dim CodeAscii As Integer 'Almacena el codigo Ascii de la letra
Dim sLetra As String 'Almacena una letra
    'Bucle que recorre cada letra del sTexto
    For i = 1 To Len(sTexto)
        sLetra = mid(sTexto, i, 1) 'Almacena la letra
            CodeAscii = ((Asc(sLetra) + 123) - 123) 'Obtiene el Ascii del sLetra
            If CodeAscii < 100 Then  'Si es menor que 100
                Encriptar = Encriptar & "0" & CodeAscii 'Imprime un 0 delante para que tenga 3 caracteres
            Else
                Encriptar = Encriptar & CodeAscii 'Lo deja talcual
            End If
    DoEvents 'Realiza cada evento
    Next i
End Function 'Fin de la funcion
 
Public Function DesEncriptar(ByVal sTexto As String) As String
On Error Resume Next   'En caso de error continua
Dim i, T As Integer
Dim sCodeAscii As String
Dim lnCodeAscii As Long
    T = 1
    'Bucle que recorre el sTexto y toma de a 3 caracteres
    For i = 1 To Len(sTexto) / 3
        sCodeAscii = mid(sTexto, T, 3) 'Toma 3 caracteres y los almacena en sCodeAscii
        lnCodeAscii = ((Val(sCodeAscii) - 123) + 123) 'Tranforma sCodeAscii en numero y lo almacena en lnCodeAscii
        T = T + 3 'Aumenta en 3 la variable t
            DesEncriptar = DesEncriptar & Chr(lnCodeAscii) 'Transforma el Ascii al caracter correspondient e
    DoEvents 'Realiza eventos
    Next i
End Function 'Fin de la funcion
