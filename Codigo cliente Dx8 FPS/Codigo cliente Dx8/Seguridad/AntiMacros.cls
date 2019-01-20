VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AntiMacros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 
Private Const MINIMOCLICK = 15 'en ms
Private Const MINIMOKEY = 15 'en ms
Private Const MAXKEYCODES = 1023
 
Private m_CountsPerSecond As Currency
Private m_Start As Currency
Private m_Stop As Currency
Private k_Start() As Currency
Private k_Stop() As Currency
Private m_ApiOverhead As Currency
 
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
 
Private Sub Class_Initialize()
    'Does the system support a performance counter
    If QueryPerformanceFrequency(m_CountsPerSecond) Then
        Dim I As Long, TotalOverhead As Currency
       
        'Find out how long it takes the system to call the API function
        For I = 1 To 1000
            QueryPerformanceCounter m_Start
            QueryPerformanceCounter m_Stop
            TotalOverhead = TotalOverhead + m_Stop - m_Start
        Next I
        m_ApiOverhead = TotalOverhead / 1000
        Debug.Print m_ApiOverhead
    Else
        m_CountsPerSecond = 1
    End If
    m_Start = 0
    m_Stop = 0
    ReDim k_Start(1 To MAXKEYCODES)
    ReDim k_Stop(1 To MAXKEYCODES)
End Sub
 
Public Property Get Supported() As Boolean
    'Does the system support a performance counter
    Supported = QueryPerformanceCounter(0)
End Property
 
'Public Sub StartTimer()
Public Sub ClickRatonDown()
    'Get the start time
    QueryPerformanceCounter m_Start
    m_Stop = 0
End Sub
 
'Public Sub StopTimer()
Public Function ClickRatonUP() As Boolean
On Error Resume Next
    'Get the end time
    QueryPerformanceCounter m_Stop
    If m_Start And m_Stop Then
        ClickRatonUP = (m_Stop - m_Start - m_ApiOverhead) / m_CountsPerSecond * 1000 > MINIMOCLICK
    Else
        ClickRatonUP = False
    End If
End Function
Public Sub ClickKeyDown(key As Integer)
    QueryPerformanceCounter k_Start(key)
    m_Stop = 0
End Sub
 
'Public Sub StopTimer()
Public Function ClickKeyUP(key As Integer) As Boolean
    'Get the end time
    QueryPerformanceCounter k_Stop(key)
    If k_Start(key) And k_Stop(key) Then
        ClickKeyUP = (k_Stop(key) - k_Start(key) - m_ApiOverhead) / m_CountsPerSecond * 1000 > MINIMOKEY
        k_Start(key) = 0
        k_Stop(key) = 0
    Else
        ClickKeyUP = False
    End If
End Function
 

