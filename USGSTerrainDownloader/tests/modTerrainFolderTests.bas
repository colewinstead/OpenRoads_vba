Attribute VB_Name = "modTerrainFolderTests"
Option Explicit

Public Sub RunTerrainFolderSmoke()
    Dim fs As Object, root As String, folder As String, name As String, report As String
    Dim paths As Object, i As Long, collisionSeen As Boolean, message As String
    Dim stream As Object, marker As String, errorNumber As Long
    On Error GoTo Failed
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set paths = CreateObject("Scripting.Dictionary")
    root = fs.BuildPath(Environ$("TEMP"), fs.GetTempName)
    fs.CreateFolder root
    report = Environ$("TEMP") & "\USGSTerrainDownloader-FolderTests.json"
    TerrainWriteUTF8 report, "{""status"":""running""}"
    For i = 1 To 40
        folder = TerrainNewJobFolder(root, "search")
        name = fs.GetFileName(folder)
        TerrainRequire name Like "search-####-##-##_##-##-##*", "Search folder name format failed."
        TerrainRequire Not paths.Exists(folder), "A folder was reused."
        paths.Add folder, True
        If Len(name) > Len("search-2026-11-03_10-38-21") Then collisionSeen = True
    Next i
    TerrainRequire collisionSeen, "Collision suffix was not exercised."
    folder = TerrainNewJobFolder(root, "process")
    TerrainRequire fs.GetFileName(folder) Like "terrain-####-##-##_##-##-##", "Terrain folder name format failed."
    folder = TerrainNewJobFolder(root)
    TerrainRequire fs.GetFileName(folder) Like "run-####-##-##_##-##-##", "Legacy one-argument caller failed."
    marker = fs.BuildPath(root, "keep.txt")
    Set stream = fs.CreateTextFile(marker): stream.Write "keep": stream.Close
    On Error Resume Next
    folder = TerrainNewJobFolder(root, "../unsafe")
    errorNumber = Err.Number
    On Error GoTo Failed
    TerrainRequire errorNumber <> 0, "Invalid folder purpose was accepted."
    TerrainRequire fs.FileExists(marker), "Existing files were removed."
    TerrainRequire Left$(root, Len(Environ$("TEMP")) + 1) = Environ$("TEMP") & "\", "Test cleanup escaped TEMP."
    fs.DeleteFolder root
    TerrainWriteUTF8 report, "{""status"":""passed"",""search_names"":true,""terrain_names"":true,""unique_collision_suffixes"":true,""legacy_caller"":true,""invalid_purpose_rejected"":true,""dgn_writes"":false}"
    Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    TerrainWriteUTF8 report, "{""status"":""failed"",""error"":" & TerrainQuoteJSON(message) & "}"
End Sub