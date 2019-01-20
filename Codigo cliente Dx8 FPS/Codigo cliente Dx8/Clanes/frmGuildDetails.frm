VERSION 5.00
Begin VB.Form frmGuildDetails 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCodex1 
      ForeColor       =   &H00111720&
      Height          =   285
      Index           =   0
      Left            =   930
      TabIndex        =   8
      Top             =   3960
      Width           =   5775
   End
   Begin VB.TextBox txtCodex1 
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   930
      TabIndex        =   7
      Top             =   4320
      Width           =   5775
   End
   Begin VB.TextBox txtCodex1 
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   930
      TabIndex        =   6
      Top             =   4680
      Width           =   5775
   End
   Begin VB.TextBox txtCodex1 
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   930
      TabIndex        =   5
      Top             =   5040
      Width           =   5775
   End
   Begin VB.TextBox txtCodex1 
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   930
      TabIndex        =   4
      Top             =   5400
      Width           =   5775
   End
   Begin VB.TextBox txtCodex1 
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   930
      TabIndex        =   3
      Top             =   5760
      Width           =   5775
   End
   Begin VB.TextBox txtCodex1 
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   930
      TabIndex        =   2
      Top             =   6120
      Width           =   5775
   End
   Begin VB.TextBox txtCodex1 
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   7
      Left            =   930
      TabIndex        =   1
      Top             =   6480
      Width           =   5775
   End
   Begin VB.TextBox txtDesc 
      ForeColor       =   &H00111720&
      Height          =   1575
      Left            =   700
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   $"frmGuildDetails.frx":0000
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   720
      TabIndex        =   9
      Top             =   3000
      Width           =   6255
   End
   Begin VB.Image Command1 
      Height          =   375
      Index           =   0
      Left            =   0
      Top             =   7320
      Width           =   855
   End
   Begin VB.Image Command1 
      Height          =   495
      Index           =   1
      Left            =   5640
      Top             =   6840
      Width           =   1455
   End
End
Attribute VB_Name = "frmGuildDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FenixAO DirectX8
'Engine By ·Parra
'Arreglado By Thusing
'Algunas cosas tomadas del cliente de DarkTester

Option Explicit


Private Sub command1_Click(Index As Integer)
Select Case Index

Case 0
    Unload Me
Case 1
    Dim fdesc$
    fdesc$ = Replace(txtDesc, vbCrLf, "º", , , vbBinaryCompare)
    




    
    Dim k As Integer
    Dim Cont As Integer
    Cont = 0
    For k = 0 To txtCodex1.UBound




        If Len(txtCodex1(k).Text) > 0 Then Cont = Cont + 1
    Next
    
    If Cont < 4 Then
        MsgBox "Debes definir al menos cuatro mandamientos."
        Exit Sub
    End If
    
    Dim chunk As String
    
    If CreandoClan Then
        chunk = "CIG" & fdesc$ & "¬" & ClanName & "¬" & Site
    Else
        chunk = "DESCOD" & fdesc$
    End If
    
    chunk = chunk & "¬"
    
    For k = 0 To Cont - 1
        chunk = chunk & txtCodex1(k) & "|"
    Next
    
    Call SendData(Left$(chunk, Len(chunk) - 1))
    
    CreandoClan = False
    
    Unload Me
    
End Select



End Sub

Private Sub Form_Deactivate()

If Not frmGuildLeader.Visible Then
    Me.SetFocus
Else
    
End If


End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(DirGraficos & "GuildDetailsCodex.gif")

End Sub
