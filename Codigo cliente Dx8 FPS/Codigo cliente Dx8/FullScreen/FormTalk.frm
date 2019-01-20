VERSION 5.00
Begin VB.Form FormTalk 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   315
   ClientLeft      =   3150
   ClientTop       =   8355
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   315
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   0
      Visible         =   0   'False
      Width           =   6840
   End
End
Attribute VB_Name = "FormTalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SendTxt_Change()
stxtbuffer = SendTxt.Text
End Sub
Private Sub SendTxt_KeyPress(KeyAscii As Integer)
If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
        Call ProcesaEntradaCmd(stxtbuffer)
        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
        FormTalk.Visible = False
        FormTalk.Hide
End If

End Sub

