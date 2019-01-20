Attribute VB_Name = "SolucionErrores"
Option Explicit
Public Declare Function IsUserAnAdmin Lib "SHELL32" () As Long


Public Sub RunAsAdmin()
    On Error GoTo Err
        If IsUserAnAdmin = 0 Then
            Dim numberOfMe As Integer
            numberOfMe = getNumberOfProcess(App.exeName & ".exe")
            Dim objShell As Object
            Set objShell = CreateObject("Shell.Application")
            objShell.ShellExecute App.Path & "\" & App.exeName & ".exe", "", "", "runas", 0
            Set objShell = Nothing
            While getNumberOfProcess("consent.exe") > 0
                'No hacer nada
            Wend
            If Not getNumberOfProcess(App.exeName & ".exe") > numberOfMe Then
                Call RunAsAdmin
            Else
                End
            End If
        Else

        End If
        Exit Sub
Err:
End Sub


Private Function getNumberOfProcess(ByVal Process As String) As Integer
    Dim objWMIService, colProcesses
    Set objWMIService = GetObject("winmgmts:")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name='" & Process & "'")
    getNumberOfProcess = colProcesses.Count
End Function

