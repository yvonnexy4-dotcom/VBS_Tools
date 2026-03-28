Dim fileName, fileContent, fso, outFile, shell

fileName = InputBox("Enter the name of your note:", "Note Name")

If fileName <> "" Then
    fileContent = InputBox("Enter the text for your note:", "Note Text")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set outFile = fso.CreateTextFile(fileName & ".txt", True)
    
    outFile.WriteLine(fileContent)
    outFile.Close
    
    Set shell = CreateObject("WScript.Shell")
    shell.Run fileName & ".txt"
End If