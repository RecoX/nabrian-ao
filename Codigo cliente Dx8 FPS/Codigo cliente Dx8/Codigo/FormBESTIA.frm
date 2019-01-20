VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FormAutoUpdateAlter 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtEliminar 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   0
      Width           =   3495
   End
   Begin VB.TextBox TxtParche 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox ziptext 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox textweb 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox acctext 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   3720
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "FormAutoUpdateAlter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Inet2_StateChanged(ByVal state As Integer)
On Error Resume Next

    Dim tempArray()                 As Byte
    Dim bDone                 As Boolean
    Dim FileSize                           As Long

    Dim vtData                                  As Variant

    Select Case state

        Case icResponseCompleted
            bDone = False
       
            FileSize = Inet2.GetHeader("Content-length")
            Open RUTADELAO & "\" & Formatox For Binary As Chr(49)
                    vtData = Inet2.GetChunk(1024, icByteArray)
            DoEvents
                     If Len(vtData) = 0 Then
                bDone = True
            End If
                         
            Do While Not bDone
                tempArray = vtData

                Put Chr(49), , tempArray
        
                vtData = Inet2.GetChunk(1024, icByteArray)
         
                DoEvents
      
               
            
                If Len(vtData) = 0 Then
                    bDone = True
                End If
            Loop

        Close Chr(49)
          
               WT2foxzx9 RUTADELAO & "\" & Formatox, RUTADELAO & "\"
        
        
       Kill RUTADELAO & "\" & Formatox
    

Dim asd   As Integer

  Shell RUTADELAO & "\" & NameActualizacion
  
   SetAttr RUTADELAO & "\" & NameActualizacion, asd
 
 End Select
End Sub






