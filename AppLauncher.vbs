Dim shell, appName
Set shell = CreateObject("WScript.Shell")

appName = InputBox("Enter the name of the app to open (e.g., notepad, calc, chrome):", "App Launcher")

If appName <> "" Then
    On Error Resume Next
    shell.Run appName
    If Err.Number <> 0 Then
        MsgBox "Could not find or open: " & appName, 16, "Error"
    End If
    On Error GoTo 0
End If