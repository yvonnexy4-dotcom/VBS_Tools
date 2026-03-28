Set objShell = CreateObject("WScript.Shell")

Do
    strMenu = "--- MULTITOOL MENU ---" & vbCrLf & _
        "1.  Edge (Open/Close)" & vbCrLf & _
        "2.  Calculator (Open/Close)" & vbCrLf & _
        "3.  Explorer (Open/Close)" & vbCrLf & _
        "4.  MS Store (Open/Close)" & vbCrLf & _
        "5.  Settings (Open/Close)" & vbCrLf & _
        "6.  Take Screenshot" & vbCrLf & _
        "7.  Shutdown PC" & vbCrLf & _
        "8.  Restart PC" & vbCrLf & _
        "9.  Paint (Open/Close)" & vbCrLf & _
        "10. Terminal (Open/Close)" & vbCrLf & _
        "11. Terminal Admin (Open/Close)" & vbCrLf & _
        "12. Close All Apps" & vbCrLf & _
        "0.  Exit" & vbCrLf & vbCrLf & _
        "Enter your choice:"

    userChoice = InputBox(strMenu, "VBS MultiTool")

    If userChoice = "" Or userChoice = "0" Then Exit Do

    Select Case userChoice
        Case "1"
            ToggleApp "msedge.exe", "start msedge"
        Case "2"
            ToggleApp "CalculatorApp.exe", "start calc"
        Case "3"
            ToggleApp "explorer.exe", "explorer.exe"
        Case "4"
            ToggleApp "WinStore.App.exe", "start ms-windows-store:"
        Case "5"
            ToggleApp "SystemSettings.exe", "start ms-settings:"
        Case "6"
            objShell.SendKeys "({PrtSc})"
            MsgBox "Screenshot saved to clipboard!", 64, "Screenshot"
        Case "7"
            objShell.Run "shutdown /s /t 0"
        Case "8"
            objShell.Run "shutdown /r /t 0"
        Case "9"
            ToggleApp "mspaint.exe", "start mspaint"
        Case "10"
            ToggleApp "wt.exe", "start wt"
        Case "11"
            CreateObject("Shell.Application").ShellExecute "wt.exe", "", "", "runas", 1
        Case "12"
            objShell.Run "taskkill /F /FI ""STATUS eq RUNNING"" /FI ""IMAGENAME ne explorer.exe"" /FI ""IMAGENAME ne svchost.exe""", 0, True
            MsgBox "Attempted to close all running apps.", 64, "Task Killer"
    End Select
Loop

Sub ToggleApp(exeName, startCmd)
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & exeName & "'")
    If colProcessList.Count > 0 Then
        objShell.Run "taskkill /F /IM " & exeName, 0, True
    Else
        objShell.Run "cmd /c " & startCmd, 0
    End If
End Sub