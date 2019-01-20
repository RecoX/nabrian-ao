Attribute VB_Name = "MD5"
Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal T As Long, ByVal R As String)

Public Function MD5String(P As String) As String

    Dim R As String * 32, T As Long
    R = Space(32)
    T = Len(P)
    MDStringFix P, T, R
    MD5String = R
End Function
