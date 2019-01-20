Attribute VB_Name = "Carteles"
 
'director del proyecto: #Esteban(Neliam)

'servidor basado en fénixao 1.0
'medios de contacto:
'Skype: dc.esteban
'E-mail: nabrianao@gmail.com
Option Explicit
Const XPosCartel = 360
Const YPosCartel = 335
Const MAXLONG = 40


Public Cartel As Boolean
Public Leyenda As String
Public LeyendaFormateada() As String
Public textura As Integer


Sub InitCartel(Ley As String, Grh As Integer)
If Not Cartel Then
    Leyenda = Ley
    textura = Grh
    Cartel = True
    ReDim LeyendaFormateada(0 To (Len(Ley) \ (MAXLONG \ 2)))
                
    Dim I As Integer, k As Integer, anti As Integer
    anti = 1
    k = 0
    I = 0
    Call DarFormato(Leyenda, I, k, anti)
    I = 0
    Do While LeyendaFormateada(I) <> "" And I < UBound(LeyendaFormateada)
        
       I = I + 1
    Loop
    ReDim Preserve LeyendaFormateada(0 To I)
Else
    Exit Sub
End If
End Sub


Private Function DarFormato(s As String, I As Integer, k As Integer, anti As Integer)
If anti + I <= Len(s) + 1 Then
    If ((I >= MAXLONG) And mid$(s, anti + I, 1) = " ") Or (anti + I = Len(s)) Then
        LeyendaFormateada(k) = mid$(s, anti, I + 1)
        k = k + 1
        anti = anti + I + 1
        I = 0
    Else
        I = I + 1
    End If
    Call DarFormato(s, I, k, anti)
End If
End Function


Sub DibujarCartel()
If Not Cartel Then Exit Sub
Dim X As Integer, Y As Integer
X = XPosCartel + 20
Y = YPosCartel + 60

Dim j As Integer, desp As Integer

For j = 0 To UBound(LeyendaFormateada)

  desp = desp + (frmPrincipal.font.Size) + 5
Next
End Sub

