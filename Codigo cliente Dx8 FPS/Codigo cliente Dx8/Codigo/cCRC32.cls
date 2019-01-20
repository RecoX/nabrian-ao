VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCRC32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Option Explicit




Private crc32Table() As Long
Private Const BUFFER_SIZE As Long = 8192

Public Function GetByteArrayCrc32(ByRef buffer() As Byte) As Long
   
   Dim crc32Result As Long
   crc32Result = &HFFFFFFFF
      
   Dim i As Integer
   Dim iLookup As Integer
   
   For i = LBound(buffer) To UBound(buffer)
      iLookup = (crc32Result And &HFF) Xor buffer(i)
      crc32Result = ((crc32Result And &HFFFFFF00) \ &H100) And 16777215
      crc32Result = crc32Result Xor crc32Table(iLookup)
   Next i
   
   GetByteArrayCrc32 = Not (crc32Result)


End Function

Public Function GetFileCrc32(ByRef stream As cBinaryFileStream) As Long

   Dim crc32Result As Long
   crc32Result = &HFFFFFFFF

   Dim buffer(0 To BUFFER_SIZE - 1) As Byte
   Dim readSize As Long
   readSize = BUFFER_SIZE

   Dim count As Integer
   count = stream.Read(buffer, readSize)
   
   Dim i As Integer
   Dim iLookup As Integer
   Dim tot As Integer
   
   Do While (count > 0)
      For i = 0 To count - 1
         iLookup = (crc32Result And &HFF) Xor buffer(i)
         crc32Result = ((crc32Result And &HFFFFFF00) \ &H100) And 16777215
         crc32Result = crc32Result Xor crc32Table(iLookup)
      Next i
      count = stream.Read(buffer, readSize)
   Loop

   GetFileCrc32 = Not (crc32Result)

End Function

Private Sub Class_Initialize()

    
    
    Dim dwPolynomial As Long
    dwPolynomial = &HEDB88320
    Dim i As Integer, j As Integer

    ReDim crc32Table(256)
    Dim dwCrc As Long

    For i = 0 To 255
        dwCrc = i
        For j = 8 To 1 Step -1
            If (dwCrc And 1) Then
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                dwCrc = dwCrc Xor dwPolynomial
            Else
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If
        Next j
        crc32Table(i) = dwCrc
    Next i

End Sub
