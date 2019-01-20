Attribute VB_Name = "mdlEstado"
Option Explicit
 
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Type COPYDATASTRUCT
dwData As Long
cbData As Long
lpData As Long
End Type
Private Const WM_COPYDATA = &H4A
Private Sub Form_Load()
Call SetMusicInfoO("Myself", "Debut", "The Song That Never Ends")
End Sub
Public Sub SetMusicInfoO(ByRef r_sArtist As String, ByRef r_sAlbum As String, ByRef r_sTitle As String, Optional ByRef r_sWMContentID As String = vbNullString, Optional ByRef r_sFormat As String = "{0} - {1}", Optional ByRef r_bShow As Boolean = True)
Dim udtData As COPYDATASTRUCT
Dim sBuffer As String
Dim hMSGRUI As Long
sBuffer = "\0Music\0" & Abs(r_bShow) & "\0" & r_sFormat & "\0" & r_sArtist & "\0" & r_sTitle & "\0" & r_sAlbum & "\0" & r_sWMContentID & "\0" & vbNullChar
 
udtData.dwData = &H547
udtData.lpData = StrPtr(sBuffer)
udtData.cbData = LenB(sBuffer)
 
Do
hMSGRUI = FindWindowEx(0&, hMSGRUI, "MsnMsgrUIManager", vbNullString)
 
If (hMSGRUI > 0) Then
Call SendMessage(hMSGRUI, WM_COPYDATA, 0, VarPtr(udtData))
End If
 
Loop Until (hMSGRUI = 0)
End Sub
 

