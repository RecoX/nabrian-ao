Attribute VB_Name = "VolumeModificacion"
Option Explicit
 
' declaraciones api , constantes y udt
'''''''''''''''''''''''''''''''''''''''''''''''''
 
Private Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Private Const MMSYSERR_NOERROR = 0
Private Const MAXPNAMELEN = 32
Private Const MIXER_LONG_NAME_CHARS = 64
Private Const MIXER_SHORT_NAME_CHARS = 16
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Private Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Private Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
 
Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = _
               (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
               
Private Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = _
               (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
 
Private Const MIXERLINE_COMPONENTTYPE_SRC_LINE = _
               (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
 
Private Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
 
Private Const MIXERCONTROL_CONTROLTYPE_FADER = _
               (MIXERCONTROL_CT_CLASS_FADER Or _
               MIXERCONTROL_CT_UNITS_UNSIGNED)
 
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME = _
               (MIXERCONTROL_CONTROLTYPE_FADER + 1)
 
Private Declare Function mixerClose Lib "winmm.dll" _
               (ByVal hmx As Long) As Long
   
Private Declare Function mixerGetControlDetails Lib "winmm.dll" _
               Alias "mixerGetControlDetailsA" _
               (ByVal hmxobj As Long, _
               pmxcd As MIXERCONTROLDETAILS, _
               ByVal fdwDetails As Long) As Long
   
Private Declare Function mixerGetDevCaps Lib "winmm.dll" _
               Alias "mixerGetDevCapsA" _
               (ByVal uMxId As Long, _
               ByVal pmxcaps As MIXERCAPS, _
               ByVal cbmxcaps As Long) As Long
   
Private Declare Function mixerGetID Lib "winmm.dll" _
               (ByVal hmxobj As Long, _
               pumxID As Long, _
               ByVal fdwId As Long) As Long
               
Private Declare Function mixerGetLineControls Lib "winmm.dll" _
               Alias "mixerGetLineControlsA" _
               (ByVal hmxobj As Long, _
               pmxlc As MIXERLINECONTROLS, _
               ByVal fdwControls As Long) As Long
               
Private Declare Function mixerGetLineInfo Lib "winmm.dll" _
               Alias "mixerGetLineInfoA" _
               (ByVal hmxobj As Long, _
               pmxl As MIXERLINE, _
               ByVal fdwInfo As Long) As Long
               
Private Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
 
Private Declare Function mixerMessage Lib "winmm.dll" _
               (ByVal hmx As Long, _
               ByVal uMsg As Long, _
               ByVal dwParam1 As Long, _
               ByVal dwParam2 As Long) As Long
               
Private Declare Function mixerOpen Lib "winmm.dll" _
               (phmx As Long, _
               ByVal uMxId As Long, _
               ByVal dwCallback As Long, _
               ByVal dwInstance As Long, _
               ByVal fdwOpen As Long) As Long
               
Private Declare Function mixerSetControlDetails Lib "winmm.dll" _
               (ByVal hmxobj As Long, _
               pmxcd As MIXERCONTROLDETAILS, _
               ByVal fdwDetails As Long) As Long
               
Private Declare Sub CopyStructFromPtr Lib "kernel32" _
               Alias "RtlMoveMemory" _
               (struct As Any, _
               ByVal ptr As Long, ByVal cb As Long)
               
Private Declare Sub CopyPtrFromStruct Lib "kernel32" _
               Alias "RtlMoveMemory" _
               (ByVal ptr As Long, _
               struct As Any, _
               ByVal cb As Long)
               
Private Declare Function GlobalAlloc Lib "kernel32" _
               (ByVal wFlags As Long, _
               ByVal dwBytes As Long) As Long
               
Private Declare Function GlobalLock Lib "kernel32" _
               (ByVal hMem As Long) As Long
               
Private Declare Function GlobalFree Lib "kernel32" _
               (ByVal hMem As Long) As Long
 
Private Type MIXERCAPS
    wMid As Integer                   '  manufacturer id
    wPid As Integer                   '  product id
    vDriverVersion As Long            '  version of the driver
    szPname As String * MAXPNAMELEN   '  product name
    fdwSupport As Long                '  misc. support bits
    cDestinations As Long             '  count of destinations
End Type
 
Private Type MIXERCONTROL
    cbStruct As Long           '  size in Byte of MIXERCONTROL
    dwControlID As Long        '  unique control id for mixer device
    dwControlType As Long      '  MIXERCONTROL_CONTROLTYPE_xxx
    fdwControl As Long         '  MIXERCONTROL_CONTROLF_xxx
    cMultipleItems As Long     '  if MIXERCONTROL_CONTROLF_MULTIPLE set
    szShortName As String * MIXER_SHORT_NAME_CHARS  ' short name of control
    szName As String * MIXER_LONG_NAME_CHARS        ' long name of control
    lMinimum As Long           '  Minimum value
    lMaximum As Long           '  Maximum value
    reserved(10) As Long       '  reserved structure space
End Type
 
Private Type MIXERCONTROLDETAILS
    cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
    dwControlID As Long    '  control id to get/set details on
    cChannels As Long      '  number of channels in paDetails array
    item As Long           '  hwndOwner or cMultipleItems
    cbDetails As Long      '  size of _one_ details_XX struct
    paDetails As Long      '  pointer to array of details_XX structs
End Type
 
Private Type MIXERCONTROLDETAILS_UNSIGNED
    dwValue As Long        '  value of the control
End Type
 
Private Type MIXERLINE
    cbStruct As Long               '  size of MIXERLINE structure
    dwDestination As Long          '  zero based destination index
    dwSource As Long               '  zero based source index (if source)
    dwLineID As Long               '  unique line id for mixer device
    fdwLine As Long                '  state/information about line
    dwUser As Long                 '  driver specific information
    dwComponentType As Long        '  component type line connects to
    cChannels As Long              '  number of channels line supports
    cConnections As Long           '  number of connections (possible)
    cControls As Long              '  number of controls at this line
    szShortName As String * MIXER_SHORT_NAME_CHARS
    szName As String * MIXER_LONG_NAME_CHARS
    dwType As Long
    dwDeviceID As Long
    wMid  As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
End Type
 
Private Type MIXERLINECONTROLS
    cbStruct As Long       '  size in Byte of MIXERLINECONTROLS
    dwLineID As Long       '  line id (from MIXERLINE.dwLineID)
                           '  MIXER_GETLINECONTROLSF_ONEBYID or
    dwControl As Long      '  MIXER_GETLINECONTROLSF_ONEBYTYPE
    cControls As Long      '  count of controls pmxctrl points to
    cbmxctrl As Long       '  size in Byte of _one_ MIXERCONTROL
    pamxctrl As Long       '  pointer to first MIXERCONTROL array
End Type
             
 Private Declare Function waveOutGetNumDevs _
   Lib "winmm.dll" () As Long
 
Private hMixer As Long
Private volCtrl As MIXERCONTROL
Private rc As Long
Private ok As Boolean
Private VolActual As Long
 
 
Function Verificar_tarjeta() As Boolean
     Dim ret As Long
     ret = waveOutGetNumDevs()
     
     If ret >= 0 Then
        Verificar_tarjeta = True
     Else
        Verificar_tarjeta = False
     End If
End Function
 
Private Function GetVolumeControl(ByRef hMixer As Long, _
                        ByVal componentType As Long, _
                        ByVal ctrlType As Long, _
                        ByRef mxc As MIXERCONTROL) As Boolean
                       
    ' This function attempts to obtain a mixer control.
    ' Returns True if successful.
    Dim mxlc As MIXERLINECONTROLS
    Dim mxl As MIXERLINE
    Dim hMem As Long
   
    mxl.cbStruct = Len(mxl)
    mxl.dwComponentType = componentType
   
    ' Obtain a line corresponding to the component type
    rc = mixerGetLineInfo(hMixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   
    If (MMSYSERR_NOERROR = rc) Then
        mxlc.cbStruct = Len(mxlc)
        mxlc.dwLineID = mxl.dwLineID
        mxlc.dwControl = ctrlType
        mxlc.cControls = 1
        mxlc.cbmxctrl = Len(mxc)
       
        ' Allocate a buffer for the control
        hMem = GlobalAlloc(&H40, Len(mxc))
        mxlc.pamxctrl = GlobalLock(hMem)
        mxc.cbStruct = Len(mxc)
       
        ' Get the control
        rc = mixerGetLineControls(hMixer, _
                                  mxlc, _
                                  MIXER_GETLINECONTROLSF_ONEBYTYPE)
       
        If (MMSYSERR_NOERROR = rc) Then
            GetVolumeControl = True
           
            ' Copy the control into the destination structure
            CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
                         
        Else
            GetVolumeControl = False
        End If
       
        GlobalFree (hMem)
        Exit Function
    End If
   
    GetVolumeControl = False
End Function
 
Private Function GetVolumen(ByRef hMixer As Long, _
                                ByRef mxc As MIXERCONTROL) As Long
 
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Vol As MIXERCONTROLDETAILS_UNSIGNED
    Dim hMem2 As Long
   
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Vol)
   
    ' Allocate a buffer for the control value buffer
    hMem2 = GlobalAlloc(&H40, Len(Vol))
    mxcd.paDetails = GlobalLock(hMem2)
    mxcd.cChannels = 1
   
    ' Get the control value
    rc = mixerGetControlDetails(hMixer, _
                               mxcd, _
                               MIXER_GETCONTROLDETAILSF_VALUE)
   
    '
    ' Copy the data into the control value buffer
    CopyStructFromPtr Vol, mxcd.paDetails, Len(Vol)
    '
    GlobalFree (hMem2)
   
    If (rc = MMSYSERR_NOERROR) Then
        GetVolumen = Vol.dwValue
       
    Else
        GetVolumen = -1&
       
    End If
End Function
 
Private Function SetVolumeControl(ByVal hMixer As Long, _
                        mxc As MIXERCONTROL, _
                        ByVal volume As Long) As Boolean
    ' This function sets the value for a volume control.
    ' Returns True if successful
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Vol As MIXERCONTROLDETAILS_UNSIGNED
    Dim hMem As Long
    'Dim rc As Long
   
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Vol)
   
    ' Allocate a buffer for the control value buffer
    hMem = GlobalAlloc(&H40, Len(Vol))
    mxcd.paDetails = GlobalLock(hMem)
    mxcd.cChannels = 1
    Vol.dwValue = volume
   
    ' Copy the data into the control value buffer
    CopyPtrFromStruct mxcd.paDetails, Vol, Len(Vol)
   
    ' Set the control value
    rc = mixerSetControlDetails(hMixer, _
                               mxcd, _
                               MIXER_SETCONTROLDETAILSF_VALUE)
   
    GlobalFree (hMem)
    If (MMSYSERR_NOERROR = rc) Then
        SetVolumeControl = True
    Else
        SetVolumeControl = False
    End If
   
End Function
 
 Function OpenMixer() As Long
 
'    // Open the mixer. This opens the mixer with a deviceID of 0. If you
'    // have a single sound card/mixer, then this will open it. If you have
'    // multiple sound cards/mixers, the deviceIDs will be 0, 1, 2, and
'    // so on.
'    rc = mixerOpen(&hMixer, 0,0,0,0);
'    if (MMSYSERR_NOERROR == rc) {
'           // Couldn't open the mixer.
'    }
'
    ' Open the mixer with deviceID 0.
    rc = mixerOpen(hMixer, 0, 0, 0, 0)
    If ((MMSYSERR_NOERROR <> rc)) Then
        'MsgBox "Couldn't open the mixer."
        OpenMixer = -1
        Exit Function
    End If
       
    ' Get the waveout volume control
    ok = GetVolumeControl(hMixer, _
                         MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                         MIXERCONTROL_CONTROLTYPE_VOLUME, _
                         volCtrl)
       
    If (ok = True) Then
        OpenMixer = GetVolumen(hMixer, volCtrl)
    Else
        OpenMixer = -1
    End If
End Function
 
Sub CloseMixer()
    On Error Resume Next
    Call mixerClose(hMixer)
End Sub
 
Public Property Get Volumen() As Byte
    Dim v As Long
    v = GetVolumen(hMixer, volCtrl)
    Volumen = CByte((v / 65535) * 100)
End Property
 
Public Property Let Volumen(ByVal NewValue As Byte)
    Dim v As Long
   
    If (NewValue > 100) Then
       MsgBox "El Valor máximo no puede ser superior a 100", vbCritical
       NewValue = 100
    End If
   
 
    v = CLng(NewValue * 65535) / 100
    If Not (v > volCtrl.lMaximum Or v < volCtrl.lMinimum) Then
        Call SetVolumeControl(hMixer, volCtrl, v)
    End If
End Property

