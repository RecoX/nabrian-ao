VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBmpMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Enum eModoBMPMan
    mododinamico = 0
    ModoCacheado = 1
End Enum

Private Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type BMP_DB_ENTRY
    FileName As Integer
    Height As Long
    Width As Long
    Size As Long
    Accesos As Long
    Cont As Long
End Type

Private Type BMP_INFO
    Height As Long
    Width As Long
    Size As Long
End Type

Private mModo As eModoBMPMan
Private mBMPDB() As BMP_DB_ENTRY
Private mDDraw As DirectDraw7
Private mSurfaceDB() As DirectDrawSurface7
Private mKeys() As Integer

Private mMaxEntries As Long
Private mContador As Long
Private mCantidadGraficos As Long

Private GrhPath As String

Public Property Let Path(ByVal inString As String)
GrhPath = inString
End Property
Public Function GetBMP(FileName As Integer) As DirectDrawSurface7
Dim RetVal As Integer

Select Case mModo
Case mododinamico
    If mKeys(FileName) <> 0 Then
        mBMPDB(mKeys(FileName)).Accesos = mBMPDB(mKeys(FileName)).Accesos + 1
        Set GetBMP = mSurfaceDB(mKeys(FileName))
    Else
        RetVal = BuscaVacio()
        If RetVal = -1 Then
            RetVal = 1
            RetVal = BorraMenosUsado()
        End If
        Call CrearGrafico(FileName, RetVal)
        mBMPDB(RetVal).Accesos = mBMPDB(RetVal).Accesos + 1
        mKeys(FileName) = RetVal
        Set GetBMP = mSurfaceDB(RetVal)
    End If
Case ModoCacheado
    Set GetBMP = mSurfaceDB(FileName)
End Select

End Function
Public Function Init(Optional modo As eModoBMPMan = ModoCacheado) As Long
On Error Resume Next
Dim loopc As Integer
mModo = modo

Select Case modo
    Case mododinamico
        If mMaxEntries < 1 Then
            Init = -1
            Exit Function
        End If
        
        ReDim mBMPDB(0 To mMaxEntries)
        ReDim mSurfaceDB(0 To mMaxEntries)
        ReDim mKeys(0 To 17000) As Integer
        
        mCantidadGraficos = 0
        mContador = 0
        Init = 1
    
    Case ModoCacheado
        ReDim mSurfaceDB(1 To 20000) 'Config_Inicio.NumeroDeBMPs
        For loopc = 1 To 20000 'Config_Inicio.NumeroDeBMPs
            Call CrearGrafico(loopc, loopc)
        Next
    
End Select

Exit Function

End Function

Public Function BorrarTodo() As Long

On Error Resume Next
Dim i As Integer

Select Case mModo
Case mododinamico
    For i = 1 To mMaxEntries
        Call ZeroMemory(mBMPDB(i), Len(mBMPDB(i)))
        Set mSurfaceDB(i) = Nothing
    Next
Case ModoCacheado
    For i = 1 To 20000 'Config_Inicio.NumeroDeBMPs
        Call ZeroMemory(mBMPDB(i), Len(mBMPDB(i)))
        Set mSurfaceDB(i) = Nothing
    Next
End Select

BorrarTodo = 1
mCantidadGraficos = 0
End Function

Public Function Borrar(FileName As Integer) As Integer

Borrar = -1
Call ZeroMemory(mBMPDB(mKeys(FileName)), Len(mBMPDB(mKeys(FileName))))
Set mSurfaceDB(mKeys(FileName)) = Nothing
mKeys(FileName) = 0
Borrar = 1
mCantidadGraficos = mCantidadGraficos - 1
End Function

Private Function CrearGrafico(Archivo As Integer, bIndex As Integer) As Integer

Dim ddsd As DDSURFACEDESC2, BMPInfo As BMP_INFO, RetVal As Long, ddck As DDCOLORKEY
ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH

Call GetBMPHeader(GrhPath & Archivo & ".bmp", BMPInfo)

With ddsd
    If RenderMod.bUseVideo Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    
    .lHeight = BMPInfo.Height
    .lWidth = BMPInfo.Width
End With

ddck.high = 0: ddck.low = 0

If mModo = mododinamico Then
    With mBMPDB(bIndex)
        .FileName = Archivo
        .Accesos = 0
        .Height = BMPInfo.Height
        .Width = BMPInfo.Width
        .Size = BMPInfo.Size
        .Cont = mContador
    End With
    mContador = mContador + 1

mCantidadGraficos = mCantidadGraficos + 1
End If
If Archivo <> 0 Then Set mSurfaceDB(bIndex) = mDDraw.CreateSurfaceFromFile(GrhPath & Archivo & ".bmp", ddsd)
mSurfaceDB(bIndex).SetColorKey DDCKEY_SRCBLT, ddck

CrearGrafico = 1

End Function

Private Sub Class_Initialize()
mMaxEntries = 0
End Sub

Private Function BuscaVacio() As Integer
Dim i As Integer
For i = 1 To mMaxEntries
    If (mSurfaceDB(i) Is Nothing) Then
        BuscaVacio = i
        Exit Function
    End If
Next

BuscaVacio = -1
End Function

Private Function BorraMenosUsado() As Integer

BorraMenosUsado = 1

Dim lValor As Long
Dim lCont As Long

lValor = mBMPDB(1).Accesos
lCont = mBMPDB(1).Cont

Dim i As Integer

For i = 1 To mMaxEntries
    If mBMPDB(i).Accesos < lValor And mBMPDB(i).Cont < lCont Then
        lValor = mBMPDB(i).Accesos
        lCont = mBMPDB(i).Cont
        BorraMenosUsado = i
    End If
Next
mCantidadGraficos = mCantidadGraficos - 1

mKeys(mBMPDB(BorraMenosUsado).FileName) = 0
Set mSurfaceDB(BorraMenosUsado) = Nothing
Call ZeroMemory(mBMPDB(BorraMenosUsado), Len(mBMPDB(BorraMenosUsado)))

End Function

Public Property Let lpDirectDraw7(ByRef objDirectDraw7 As DirectDraw7)
Set mDDraw = objDirectDraw7
End Property
Private Function GetBMPHeader(Archivo As String, ByRef lpBMPInfo As BMP_INFO) As Long
Dim nfile As Integer, H_BMPH As BITMAPFILEHEADER, H_BMPI As BITMAPINFOHEADER
nfile = FreeFile()

Open Archivo For Binary Access Read As #nfile
Get #nfile, , H_BMPH
Get #nfile, , H_BMPI
Close #nfile






lpBMPInfo.Height = H_BMPI.biHeight
lpBMPInfo.Width = H_BMPI.biWidth
lpBMPInfo.Size = H_BMPH.bfSize

GetBMPHeader = 1

End Function

Public Property Get MaxEntries() As Long
MaxEntries = mMaxEntries
End Property

Public Property Let MaxEntries(ByVal vNewValue As Long)
mMaxEntries = vNewValue
End Property

Public Property Get CantidadGraficos() As Variant
CantidadGraficos = mCantidadGraficos
End Property

