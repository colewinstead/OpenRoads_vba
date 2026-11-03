Attribute VB_Name = "modTerrainFileTests"
Option Explicit

Public Sub RunTerrainFileIOSmoke()
    Dim fs As Object, stream As Object, path As String, report As String
    Dim handle As Integer, opened As Boolean, value As String, message As String
    Dim errorNumber As Long, errorText As String, started As Double, elapsed As Double
    On Error GoTo Failed
    Set fs = CreateObject("Scripting.FileSystemObject")
    path = fs.BuildPath(Environ$("TEMP"), fs.GetTempName)
    report = Environ$("TEMP") & "\USGSTerrainDownloader-FileTests.json"
    TerrainWriteUTF8 report, "{""status"":""running""}"
    TerrainRequire TerrainReadUI(path) = "", "Missing UI file did not read as empty."
    Set stream = fs.CreateTextFile(path, True, True): stream.Close
    TerrainRequire TerrainReadUI(path) = "", "Empty Unicode file did not read as empty."
    value = "ready" & vbCrLf & ChrW$(&H394) & " US survey feet"
    Set stream = fs.CreateTextFile(path, True, True)
    stream.Write value: stream.Close: Set stream = Nothing
    TerrainRequire TerrainReadUI(path) = value, "Unicode UI text changed."
    handle = FreeFile
    Open path For Binary Access Read Write Lock Read Write As #handle
    opened = True
    started = Timer
    On Error Resume Next
    value = TerrainReadUI(path)
    errorNumber = Err.Number: errorText = Err.Description
    On Error GoTo Failed
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400#
    Close #handle: opened = False
    TerrainRequire errorNumber = 70 Or errorNumber = 75, "Exclusive lock did not raise a file-access error."
    TerrainRequire InStr(errorText, path) > 0, "Persistent read error omitted the file path."
    TerrainRequire elapsed >= 0.7, "Read failed immediately instead of retrying."
    TerrainRequire InStr(TerrainReadUI(path), "ready") = 1, "Read did not recover after unlocking."
    fs.DeleteFile path
    TerrainWriteUTF8 report, "{""status"":""passed"",""unicode_and_empty"":true,""locked_read_error"":" & CStr(errorNumber) & _
        ",""retry_seconds"":" & TerrainNumber(elapsed) & ",""file_path_reported"":true,""read_after_unlock"":true,""dgn_writes"":false}"
    Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    If opened Then Close #handle
    If Not stream Is Nothing Then stream.Close
    If Not fs Is Nothing Then
        If fs.FileExists(path) Then fs.DeleteFile path
    End If
    TerrainWriteUTF8 report, "{""status"":""failed"",""error"":" & TerrainQuoteJSON(message) & "}"
End Sub