Dim input, shell, fso, tempFile, htmlPath
input = InputBox("Enter the text or word you want to turn red:", "Red Text Tool")

If input <> "" Then
    Set shell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    htmlPath = shell.ExpandEnvironmentStrings("%TEMP%\redtext.html")
    
    Set tempFile = fso.CreateTextFile(htmlPath, True)
    tempFile.WriteLine "<html><body style='display:flex;justify-content:center;align-items:center;height:100vh;margin:0;font-family:sans-serif;'>"
    tempFile.WriteLine "<h1 style='color:red;font-size:48px;'>" & input & "</h1>"
    tempFile.WriteLine "</body></html>"
    tempFile.Close
    
    shell.Run "mshta.exe """ & htmlPath & """", 1, True
    fso.DeleteFile htmlPath
End If