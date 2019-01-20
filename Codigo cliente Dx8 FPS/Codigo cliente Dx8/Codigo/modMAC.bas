Attribute VB_Name = "modMACYHD"
Option Explicit
 
Public Const NCBASTAT As Long = &H33
Public Const NCBNAMSZ As Long = 16
Public Const HEAP_ZERO_MEMORY As Long = &H8
Public Const HEAP_GENERATE_EXCEPTIONS As Long = &H4
Public Const NCBRESET As Long = &H32
 
Public Type NET_CONTROL_BLOCK 'NCB
ncb_command As Byte
ncb_retcode As Byte
ncb_lsn As Byte
ncb_num As Byte
ncb_buffer As Long
ncb_length As Integer
ncb_callname As String * NCBNAMSZ
ncb_name As String * NCBNAMSZ
ncb_rto As Byte
ncb_sto As Byte
ncb_post As Long
ncb_lana_num As Byte
ncb_cmd_cplt As Byte
ncb_reserve(9) As Byte ' Reserved, must be 0
ncb_event As Long
End Type
 
Public Type ADAPTER_STATUS
adapter_address(5) As Byte
rev_major As Byte
reserved0 As Byte
adapter_type As Byte
rev_minor As Byte
duration As Integer
frmr_recv As Integer
frmr_xmit As Integer
iframe_recv_err As Integer
xmit_aborts As Integer
xmit_success As Long
recv_success As Long
iframe_xmit_err As Integer
recv_buff_unavail As Integer
t1_timeouts As Integer
ti_timeouts As Integer
Reserved1 As Long
free_ncbs As Integer
max_cfg_ncbs As Integer
max_ncbs As Integer
xmit_buf_unavail As Integer
max_dgram_size As Integer
pending_sess As Integer
max_cfg_sess As Integer
max_sess As Integer
max_sess_pkt_size As Integer
name_count As Integer
End Type
 
Public Type NAME_BUFFER
name As String * NCBNAMSZ
name_num As Integer
name_flags As Integer
End Type
 
Public Type ASTAT
adapt As ADAPTER_STATUS
NameBuff(30) As NAME_BUFFER
End Type
 
Public Declare Function Netbios Lib "netapi32.dll" _
(pncb As NET_CONTROL_BLOCK) As Byte
 
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, ByVal _
hpvSource As Long, ByVal _
cbCopy As Long)
 
Public Declare Function GetProcessHeap Lib "Kernel32" () As Long
 
Public MotherBoardNumber As String

Public Declare Function HeapAlloc Lib "Kernel32" _
(ByVal hHeap As Long, ByVal dwFlags As Long, _
ByVal dwBytes As Long) As Long
 
Public Declare Function HeapFree Lib "Kernel32" _
(ByVal hHeap As Long, _
ByVal dwFlags As Long, _
lpMem As Any) As Long

Declare Function GetVolumeInformation Lib "kernel32.dll" Alias _
"GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal _
lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, _
lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal _
nFileSystemNameSize As Long) As Long
 
Public Function GetMACAddress() As String
'Devuelve la dirección MAC del dispositivo de red
'instalado, devuelve una cadena formateada
 
Dim Tmp As String, pASTAT As Long
Dim NCB As NET_CONTROL_BLOCK, AST As ASTAT
 
'The IBM NetBIOS 3.0 specifications defines four basic
'NetBIOS environments under the NCBRESET command. Win32
'follows the OS/2 Dynamic Link Routine (DLR) environment.
'This means that the first NCB issued by an application
'must be a NCBRESET, with the exception of NCBENUM.
'The Windows NT implementation differs from the IBM
'NetBIOS 3.0 specifications in the NCB_CALLNAME field.
NCB.ncb_command = NCBRESET
Call Netbios(NCB)
 
'To get the Media Access Control (MAC) address for an
'ethernet adapter programmatically, use the Netbios()
'NCBASTAT command and provide a "*" as the name in the
'NCB.ncb_CallName field (in a 16-chr string).
NCB.ncb_callname = "* "
NCB.ncb_command = NCBASTAT
 
'For machines with multiple network adapters you need to
'enumerate the LANA numbers and perform the NCBASTAT
'command on each. Even when you have a single network
'adapter, it is a good idea to enumerate valid LANA numbers
'first and perform the NCBASTAT on one of the valid LANA
'numbers. It is considered bad programming to hardcode the
'LANA number to 0 (see the comments section below).
NCB.ncb_lana_num = 0
NCB.ncb_length = Len(AST)
 
pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS _
Or HEAP_ZERO_MEMORY, NCB.ncb_length)
 
If pASTAT = 0 Then
Debug.Print "memory allocation failed!"
Exit Function
End If
 
NCB.ncb_buffer = pASTAT
Call Netbios(NCB)
 
CopyMemory AST, NCB.ncb_buffer, Len(AST)
 
Tmp = format$(hex(AST.adapt.adapter_address(0)), "00") & "-" & _
format$(hex(AST.adapt.adapter_address(1)), "00") & "-" & _
format$(hex(AST.adapt.adapter_address(2)), "00") & "-" & _
format$(hex(AST.adapt.adapter_address(3)), "00") & "-" & _
format$(hex(AST.adapt.adapter_address(4)), "00") & "-" & _
format$(hex(AST.adapt.adapter_address(5)), "00")
 
HeapFree GetProcessHeap(), 0, pASTAT
 
GetMACAddress = Tmp
End Function
Public Function GetSerialNumber( _
    ByVal sDrive As String) As Long

    If Len(sDrive) Then
        If InStr(sDrive, "\\") = 1 Then
            ' Make sure we end in backslash for UNC
            If Right$(sDrive, 1) <> "\" Then
                sDrive = sDrive & "\"
            End If
        Else
            ' If not UNC, take first letter as drive
            sDrive = Left$(sDrive, 1) & ":\"
        End If
    Else
        ' Else just use current drive
        sDrive = vbNullString
    End If

    ' Grab S/N -- Most params can be NULL
    Call GetVolumeInformation( _
        sDrive, vbNullString, 0, GetSerialNumber, _
        ByVal 0&, ByVal 0&, vbNullString, 0)
End Function

Function GetSerialNumber2()

    Dim wmi As Object
    Dim mos As Object
    Dim mo As Object
    
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set mos = wmi.ExecQuery("Select * from Win32_Baseboard")

    For Each mo In mos
        MotherBoardNumber = mo.SerialNumber
    Next
End Function
