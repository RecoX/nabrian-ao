VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBinaryFileStream"
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

Private m_sFile As String
Private m_iFile As Integer
Private m_iLen As Long
Private m_iOffset As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Property Get File() As String
   File = m_sFile
End Property
Public Property Let File(ByVal sFile As String)
   Dispose
   m_sFile = sFile
   Dim lErr As Long
   If (FileExists(m_sFile, lErr)) Then
      m_iFile = FreeFile
      Open m_sFile For Binary Access Read Lock Write As #m_iFile
      m_iLen = LOF(m_iFile)
   Else
   
   End If
End Property

Private Function FileExists(ByVal sFile As String, ByRef lErr As Long) As Boolean
   
   lErr = 0
   On Error Resume Next
   Dim sDir As String
   sDir = Dir$(sFile)
   lErr = Err.Number
   On Error GoTo 0
   
   If (lErr = 0) Then
      If (Len(sDir) > 0) Then
         FileExists = True
      Else
         lErr = 53
      End If
   End If
   
End Function

Public Property Get Length() As Long
   Length = m_iLen
End Property

Public Function Read( _
      buffer() As Byte, _
      ByVal readSize As Long _
   ) As Long
   
   Dim lReadSize As Long
   lReadSize = readSize
   If (m_iOffset + lReadSize >= m_iLen) Then
      readSize = m_iLen - m_iOffset
      If (readSize > 0) Then
         ReDim newBuffer(0 To readSize - 1) As Byte
         Get #m_iFile, , newBuffer
         CopyMemory buffer(0), newBuffer(0), readSize
      Else
         Dispose
      End If
      m_iOffset = m_iOffset + readSize
   Else
      
      Get #m_iFile, , buffer
      m_iOffset = m_iOffset + readSize
   End If
   Read = readSize
   
End Function

Public Sub Dispose()
   If (m_iFile) Then
      Close #m_iFile
      m_iFile = 0
   End If
End Sub

Private Sub Class_Terminate()
   Dispose
End Sub
