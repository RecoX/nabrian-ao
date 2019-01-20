Attribute VB_Name = "DibujarInventario"
 
'director del proyecto: #Esteban(Neliam)

'servidor basado en fénixao 1.0
'medios de contacto:
'Skype: dc.esteban
'E-mail: nabrianao@gmail.com
Option Explicit
 
Public Const XCantItems = 5
 
Public OffsetDelInv As Integer
Public ItemElegido As Integer
Public imgOld As Integer

Public mx As Integer
Public my As Integer
 
Private bStaticInit  As Boolean
Private r1           As RECT, r2 As RECT, auxr As RECT
Private rBox         As RECT
Private rBoxFrame(2) As RECT
Private iFrameMod    As Integer
Sub ActualizarOtherInventory(Slot As Integer)
 
If OtherInventory(Slot).OBJIndex = 0 Then
    frmComerciar.List1(0).List(Slot - 1) = "Nada"
Else
    frmComerciar.List1(0).List(Slot - 1) = OtherInventory(Slot).name
End If
 
If frmComerciar.List1(0).ListIndex = Slot - 1 And lista = 0 Then Call ActualizarInformacionComercio(0)
 
End Sub
Sub ActualizarInventario(Slot As Integer)
Dim OBJIndex As Long
Dim NameSize As Byte
 
If UserInventory(Slot).Amount = 0 Then
    frmPrincipal.imgObjeto(Slot).ToolTipText = "Nada"
    frmPrincipal.lblObjCant(Slot).ToolTipText = "Nada"
    frmPrincipal.lblObjCant(Slot).Caption = ""
    If ItemElegido = Slot Then frmPrincipal.Shape1.Visible = False
Else
    frmPrincipal.imgObjeto(Slot).ToolTipText = UserInventory(Slot).name
    frmPrincipal.lblObjCant(Slot).ToolTipText = UserInventory(Slot).name
    frmPrincipal.lblObjCant(Slot).Caption = CStr(UserInventory(Slot).Amount)
    If ItemElegido = Slot Then frmPrincipal.Shape1.Visible = True
End If
 
If UserInventory(Slot).GrhIndex > 0 Then
    If EncriptGraficosActiva = True Then
        'ENCRIPT GRAFICOS
    If Extract_File(Graphics, App.Path & "\Graficos\", GrhData(UserInventory(Slot).GrhIndex).FileNum & ".bmp", App.Path & "\Graficos\") Then
        frmPrincipal.imgObjeto(Slot).PICTURE = LoadPicture(DirGraficos & GrhData(UserInventory(Slot).GrhIndex).FileNum & ".bmp")
        Call Kill(App.Path & "\Graficos\*.bmp")
    End If
   'ENCRIPT GRAFICOS
    Else
     frmPrincipal.imgObjeto(Slot).PICTURE = LoadPicture(DirGraficos & GrhData(UserInventory(Slot).GrhIndex).FileNum & ".bmp") 'lectura sin encript

   End If
Else
    frmPrincipal.imgObjeto(Slot).PICTURE = LoadPicture()
End If
 
If UserInventory(Slot).Equipped > 0 Then
    frmPrincipal.Label2(Slot).Visible = True
Else
    frmPrincipal.Label2(Slot).Visible = False
End If
 
If frmComerciar.Visible Then
    If UserInventory(Slot).Amount = 0 Then
        frmComerciar.List1(1).List(Slot - 1) = "Nada"
     Else
        frmComerciar.List1(1).List(Slot - 1) = UserInventory(Slot).name
    End If
    If frmComerciar.List1(1).ListIndex = Slot - 1 And lista = 1 Then Call ActualizarInformacionComercio(1)
End If
 
End Sub
