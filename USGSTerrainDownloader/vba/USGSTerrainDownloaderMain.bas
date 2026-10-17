Attribute VB_Name = "USGSTerrainDownloaderMain"
Option Explicit

Private mModel As ModelReference
Private mPicker As clsTerrainBoundaryPicker
Private mBoundary As String
Private mOpen As Boolean
Private mUnitsContext As String
Private mGCSContext As String

Public Sub Main()
    On Error GoTo Failed
    If mOpen Then
        frmUSGSTerrainDownloader.Show vbModeless
        Exit Sub
    End If
    TerrainRequire ActiveModelReference.Type = msdModelTypeNormal, "Open an active design model, not a sheet or drawing model."
    Set mModel = ActiveModelReference
    mUnitsContext = TerrainModelDescription()
    mGCSContext = TerrainDetectedGCS()
    mBoundary = "": mOpen = True
    frmUSGSTerrainDownloader.Show vbModeless
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "USGS Terrain Downloader"
End Sub

Public Sub TerrainRequire(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2801, "USGS Terrain Downloader", message
End Sub

Public Sub TerrainCheckContext(Optional ByVal checkGCS As Boolean = False)
    TerrainRequire mOpen And Not mModel Is Nothing, "Start the terrain downloader again."
    TerrainRequire mModel.IsActive, "The active model changed. Close this dialog and start again in the intended model."
    TerrainRequire TerrainModelDescription() = mUnitsContext, "Model working units changed. Close the dialog and select the boundary again."
    If checkGCS Then TerrainRequire TerrainDetectedGCS() = mGCSContext, "Model GCS changed. Close the dialog and verify the new coordinate mapping."
End Sub

Public Function TerrainBoundaryJSON() As String
    TerrainCheckContext
    TerrainRequire Len(mBoundary) > 0, "Draw a rectangle or select a closed straight-sided shape first."
    TerrainBoundaryJSON = mBoundary
End Function

Public Sub TerrainPick(ByVal rectangle As Boolean)
    On Error GoTo Failed
    TerrainCheckContext
    Set mPicker = New clsTerrainBoundaryPicker
    frmUSGSTerrainDownloader.Hide
    If rectangle Then
        CommandState.StartPrimitive mPicker
    Else
        CommandState.StartLocate mPicker
    End If
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Boundary selection"
    frmUSGSTerrainDownloader.Show vbModeless
End Sub

Public Sub TerrainAcceptBoundary(vertices() As Point3d)
    TerrainCheckContext
    mBoundary = TerrainVerticesJSON(vertices)
    frmUSGSTerrainDownloader.BoundaryChanged "Boundary captured: " & CStr(UBound(vertices) - LBound(vertices) + 1) & " vertices in model XY master units. Search calculates geographic extent."
End Sub

Public Sub TerrainPickerEnded()
    Set mPicker = Nothing
    If mOpen Then frmUSGSTerrainDownloader.Show vbModeless
End Sub

Public Sub TerrainClose()
    mOpen = False
    If Not mPicker Is Nothing Then mPicker.StopPick
    Set mPicker = Nothing
    Set mModel = Nothing
    mBoundary = ""
End Sub
