Attribute VB_Name = "modPythonBridge"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
#End If

Public Function TerrainQuoteJSON(ByVal value As String) As String
    Dim i As Long, code As Long, ch As String, result As String
    result = Chr$(34)
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1): code = AscW(ch)
        Select Case code
            Case 34: result = result & "\" & Chr$(34)
            Case 92: result = result & "\\"
            Case 0 To 31: result = result & "\u" & Right$("0000" & Hex$(code), 4)
            Case Else: result = result & ch
        End Select
    Next i
    TerrainQuoteJSON = result & Chr$(34)
End Function

Public Function TerrainBool(ByVal value As Boolean) As String
    If value Then TerrainBool = "true" Else TerrainBool = "false"
End Function

Public Sub TerrainWriteUTF8(ByVal path As String, ByVal text As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText text
    stream.SaveToFile path, 2
    stream.Close
End Sub

Public Function TerrainReadUI(ByVal path As String) As String
    Dim fs As Object, stream As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FileExists(path) Then Exit Function
    Set stream = fs.OpenTextFile(path, 1, False, -1)
    TerrainReadUI = stream.ReadAll
    stream.Close
End Function

Public Sub TerrainRuntimeDefaults(ByRef sourceRoot As String, ByRef pythonExe As String, ByRef environmentFile As String)
    Dim fs As Object, project As Object, candidate As String, matches As Long
    Dim config As String, stream As Object
    sourceRoot = GetSetting("USGSTerrainDownloader", "Runtime", "SourceRoot", "")
    pythonExe = GetSetting("USGSTerrainDownloader", "Runtime", "Python", "")
    environmentFile = GetSetting("USGSTerrainDownloader", "Runtime", "EnvironmentFile", "")
    On Error GoTo Unavailable
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Len(sourceRoot) = 0 Then
        For Each project In Application.VBE.VBProjects
            If project.Name = "USGSTerrainDownloader" Then
                matches = matches + 1
                candidate = fs.GetParentFolderName(project.FileName)
            End If
        Next project
        If matches = 1 Then
            If fs.FileExists(fs.BuildPath(candidate, "python\launch_worker.exe")) Then sourceRoot = candidate
        End If
    End If
    config = Environ$("LOCALAPPDATA") & "\USGSTerrainDownloaderConfig"
    If Len(environmentFile) = 0 And fs.FileExists(config & "\runtime-environment.json") Then environmentFile = config & "\runtime-environment.json"
    If Len(pythonExe) = 0 And fs.FileExists(config & "\runtime-python.txt") Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2: stream.Charset = "utf-8": stream.Open
        stream.LoadFromFile config & "\runtime-python.txt"
        candidate = Trim$(stream.ReadText): stream.Close
        If fs.FileExists(candidate) And LCase$(fs.GetExtensionName(candidate)) = "exe" Then pythonExe = candidate
    End If
Unavailable:
    ' Discovery is optional; unavailable paths remain editable, never guessed.
End Sub

Public Function TerrainNewJobFolder(ByVal destination As String) As String
    Dim fs As Object, path As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    TerrainRequire fs.FolderExists(destination), "Destination must be an existing writable folder. Browse or create it first."
    path = fs.BuildPath(destination, "run-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & fs.GetTempName)
    fs.CreateFolder path
    TerrainNewJobFolder = path
End Function

Private Function CommandQuote(ByVal value As String) As String
    TerrainRequire InStr(value, Chr$(34)) = 0 And InStr(value, vbCr) = 0 And InStr(value, vbLf) = 0, "Invalid command path."
    TerrainRequire Right$(value, 1) <> "\", "Command argument cannot end in a slash."
    CommandQuote = Chr$(34) & value & Chr$(34)
End Function

Public Function TerrainStartWorker(ByVal sourceRoot As String, ByVal requestPath As String) As Object
    Dim shell As Object, fs As Object, launcher As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    launcher = fs.BuildPath(sourceRoot, "python\launch_worker.exe")
    TerrainRequire fs.FileExists(launcher), "Source folder must contain python\launch_worker.exe beside main.py."
    Set shell = CreateObject("WScript.Shell")
    Set TerrainStartWorker = shell.Exec(CommandQuote(launcher) & " " & CommandQuote(requestPath))
End Function

Public Sub TerrainYield()
    DoEvents
    Sleep 100
End Sub

Public Sub TerrainCancelJob(ByVal folder As String)
    If Len(folder) > 0 Then TerrainWriteUTF8 folder & "\cancel.flag", "cancel"
End Sub

Public Function TerrainBrowseFolder(ByVal title As String) As String
    Dim shell As Object, folder As Object
    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, title, 0)
    If Not folder Is Nothing Then TerrainBrowseFolder = folder.Self.path
End Function

Public Sub TerrainOpenFolder(ByVal path As String)
    Dim shell As Object, fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    TerrainRequire fs.FolderExists(path), "Folder does not exist."
    Set shell = CreateObject("Shell.Application")
    shell.Explore path
End Sub

Public Function TerrainDefaultDestination() As String
    Dim fs As Object, folder As String, stream As Object, probe As String
    On Error GoTo LocalFallback
    Set fs = CreateObject("Scripting.FileSystemObject")
    TerrainRequire Len(ActiveDesignFile.path) > 0, "DGN has no project folder."
    folder = fs.BuildPath(ActiveDesignFile.path, "USGS_Terrain")
    If Not fs.FolderExists(folder) Then fs.CreateFolder folder
    probe = fs.BuildPath(folder, fs.GetTempName)
    Set stream = fs.CreateTextFile(probe, False): stream.Close
    fs.DeleteFile probe
    TerrainDefaultDestination = folder
    Exit Function
LocalFallback:
    On Error GoTo Failed
    Set fs = CreateObject("Scripting.FileSystemObject")
    folder = Environ$("LOCALAPPDATA") & "\USGSTerrainDownloader"
    If Not fs.FolderExists(folder) Then fs.CreateFolder folder
    TerrainDefaultDestination = folder
    Exit Function
Failed:
    TerrainDefaultDestination = Environ$("TEMP")
End Function
