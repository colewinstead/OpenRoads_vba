Attribute VB_Name = "modTerrainORDTests"
Option Explicit

Private EvidenceFile As String

Private Function ElementIDs() As String
    Dim scan As ElementEnumerator
    Set scan = ActiveModelReference.Scan
    Do While scan.MoveNext
        ElementIDs = ElementIDs & DLongToString(scan.Current.ID) & ","
    Loop
End Function

Private Function SettingsSnapshot() As String
    Dim name As Variant
    SettingsSnapshot = ActiveDesignFile.FullName & "|" & TerrainModelDescription() & "|" & _
        ActiveSettings.Level.Name & "|" & CStr(ActiveSettings.Color) & "|" & _
        CStr(ActiveSettings.LineWeight) & "|" & ActiveSettings.LineStyle.Name & "|" & TerrainDetectedGCS()
    For Each name In Array("MS_VBAAUTOLOADPROJECTS", "MS_VBASEARCHDIRECTORIES", "MS_VBACOPYOUT", "MS_VBANEWPROJECTDIRECTORY")
        SettingsSnapshot = SettingsSnapshot & "|" & ActiveWorkspace.ConfigurationVariableValue(CStr(name), True)
    Next name
End Function

Public Sub RunTerrainORDSmoke()
    Dim beforeIDs As String, beforeSettings As String, detected As String, boundary As String
    Dim points(0 To 3) As Point3d, recovered() As Point3d, shape As ShapeElement, i As Long
    Dim gcs As GeographicCoordinateSystem, message As String, testZ As Double, unit As MeasurementUnit
    Dim sourceRoot As String, pythonExe As String, environmentFile As String, expectedSource As String
    EvidenceFile = Environ$("TEMP") & "\USGSTerrainDownloader-ORDTests.json"
    On Error GoTo Failed
    TerrainWriteUTF8 EvidenceFile, "{""status"":""running""}"
    beforeIDs = ElementIDs(): beforeSettings = SettingsSnapshot()
    TerrainRequire TerrainNumber(0.5) = "0.5", "Positive JSON fraction failed."
    TerrainRequire TerrainNumber(-0.5) = "-0.5", "Negative JSON fraction failed."
    TerrainRequire TerrainEPSGReference(" EPSG:6576 ") = "EPSG:6576", "Plain EPSG source reference failed."
    TerrainRequire TerrainEPSGReference("epsg:6576") = "EPSG:6576", "EPSG prefix normalization failed."
    TerrainRequire TerrainEPSGReference("TN83/2011F") = "", "A GCS name was guessed as EPSG."
    TerrainRequire TerrainEPSGReference("EPSG:6576 modified") = "", "Ambiguous source text was accepted."
    TerrainRequire TerrainEPSGReference("EPSG:0") = "" And TerrainEPSGReference("EPSG:-6576") = "", "Invalid EPSG source text was accepted."
    unit.Base = msdMeasurementBaseMeter: unit.UnitsPerBaseDenominator = 1
    unit.Label = "ambiguous"
    unit.UnitsPerBaseNumerator = 1
    TerrainRequire TerrainMasterUnitName(unit) = "m", "Metre definition failed."
    unit.UnitsPerBaseNumerator = 1# / 0.3048
    TerrainRequire TerrainMasterUnitName(unit) = "ft", "International foot definition failed."
    unit.UnitsPerBaseNumerator = 3937# / 1200#
    TerrainRequire TerrainMasterUnitName(unit) = "us-ft", "US survey foot definition failed."
    unit.UnitsPerBaseNumerator = 2
    TerrainRequire TerrainMasterUnitName(unit) = "", "Unsupported custom unit was guessed."
    unit.Base = msdMeasurementBaseNone
    TerrainRequire TerrainMasterUnitName(unit) = "", "Unitless model was guessed."
    unit.Base = msdMeasurementBaseMeter: unit.UnitsPerBaseDenominator = 0
    TerrainRequire TerrainMasterUnitName(unit) = "", "Invalid unit ratio was accepted."
    unit = ActiveModelReference.MasterUnit
    TerrainRuntimeDefaults sourceRoot, pythonExe, environmentFile
    Set gcs = ActiveModelReference.GetGCS(True)
    If Not gcs Is Nothing Then expectedSource = TerrainEPSGReference(gcs.Source)
    detected = TerrainDetectedGCS()
    If Not gcs Is Nothing Then TerrainRequire InStr(1, detected, gcs.Name, vbBinaryCompare) > 0, "Late-bound GCS name differs from actual GCS."
    Load frmUSGSTerrainDownloader
    With frmUSGSTerrainDownloader
        TerrainRequire .Controls("txtCRS").Text = expectedSource, "CRS suggestion differs from assigned GCS source reference."
        If Len(TerrainMasterUnitName(unit)) = 0 Then
            TerrainRequire .Controls("cboMaster").ListIndex = -1, "Unsupported model unit was guessed."
        Else
            TerrainRequire .Controls("cboMaster").Value = TerrainMasterUnitName(unit), "Master unit differs from exact model definition."
        End If
        TerrainRequire .Controls("txtRoot").Text = sourceRoot, "Source folder discovery differs."
        TerrainRequire .Controls("txtPython").Text = pythonExe, "Python discovery differs."
        TerrainRequire .Controls("txtEnv").Text = environmentFile, "GIS environment discovery differs."
        TerrainRequire Not .Controls("chkConfirmed").Value, "Grid mapping must start unconfirmed."
        TerrainRequire .Controls("cboSourceZ").ListIndex = -1, "Source Z units must require a choice."
        TerrainRequire .Controls("txtVertical").Text = "", "Vertical CRS must start undeclared."
        TerrainRequire .Controls("lblGCS").Caption = detected, "Form GCS display differs from probe."
    End With
    USGSTerrainDownloaderMain.Main
    TerrainRequire frmUSGSTerrainDownloader.Visible, "Modeless form did not become visible."
    testZ = 0
    If ActiveModelReference.Is3D Then testZ = 7
    points(0) = Point3dFromXYZ(123, 456, testZ)
    points(1) = Point3dFromXYZ(223, 456, testZ)
    points(2) = Point3dFromXYZ(223, 556, testZ)
    points(3) = Point3dFromXYZ(123, 556, testZ)
    Set shape = CreateShapeElement1(Nothing, points, msdFillModeNotFilled)
    recovered = shape.AsVertexList.GetVertices
    TerrainRequire LBound(recovered) = 0 And UBound(recovered) >= 3, "Shape vertices missing."
    For i = 0 To 3
        TerrainRequire recovered(i).X = points(i).X And recovered(i).Y = points(i).Y And recovered(i).Z = points(i).Z, _
            "Shape vertex " & CStr(i) & " changed: dX=" & TerrainNumber(recovered(i).X - points(i).X) & _
            ", dY=" & TerrainNumber(recovered(i).Y - points(i).Y) & ", dZ=" & TerrainNumber(recovered(i).Z - points(i).Z)
    Next i
    TerrainAcceptBoundary recovered
    boundary = TerrainBoundaryJSON()
    TerrainRequire boundary = TerrainVerticesJSON(recovered), "Programmatic boundary acceptance changed vertices."
    TerrainClose
    Unload frmUSGSTerrainDownloader
    TerrainRequire ElementIDs() = beforeIDs, "DGN element IDs changed."
    TerrainRequire SettingsSnapshot() = beforeSettings, "Active settings, model context or VBA configuration changed."
    TerrainWriteUTF8 EvidenceFile, "{""status"":""passed"",""gcs_probe"":" & TerrainQuoteJSON(detected) & _
        ",""model"":" & TerrainQuoteJSON(TerrainModelDescription()) & ",""element_ids"":" & TerrainQuoteJSON(beforeIDs) & _
        ",""boundary"":" & boundary & ",""source_crs_suggestion"":" & TerrainQuoteJSON(expectedSource) & ",""master_unit"":" & TerrainQuoteJSON(TerrainMasterUnitName(unit)) & _
        ",""source_root"":" & TerrainQuoteJSON(sourceRoot) & ",""python"":" & TerrainQuoteJSON(pythonExe) & _
        ",""environment_file"":" & TerrainQuoteJSON(environmentFile) & _
        ",""settings_unchanged"":true,""graphical_picking_tested"":false,""model_is_3d"":" & TerrainBool(ActiveModelReference.Is3D) & "}"
    Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    TerrainClose
    Unload frmUSGSTerrainDownloader
    On Error GoTo ReportFailed
    TerrainWriteUTF8 EvidenceFile, "{""status"":""failed"",""error"":" & TerrainQuoteJSON(message) & "}"
    Exit Sub
ReportFailed:
    MsgBox message & vbCrLf & "Could not write evidence: " & Err.Description, vbExclamation, "ORD native self-check"
End Sub

Public Sub RunTerrainLauncherSmoke()
    Dim sourceRoot As String, pythonExe As String, environmentFile As String
    Dim folder As String, report As String, request As String, worker As Object
    Dim started As Date, message As String, code As Long, status As String
    RunTerrainORDSmoke
    report = Environ$("TEMP") & "\USGSTerrainDownloader-LauncherTests.json"
    On Error GoTo Failed
    TerrainWriteUTF8 report, "{""status"":""running""}"
    TerrainRuntimeDefaults sourceRoot, pythonExe, environmentFile
    TerrainRequire Len(sourceRoot) > 0 And Len(pythonExe) > 0, "Configure the application runtime before testing its launcher."
    folder = TerrainNewJobFolder(Environ$("TEMP"))
    request = "{""schema_version"":0,""action"":""search"",""python_executable"":" & TerrainQuoteJSON(pythonExe) & _
        ",""environment_file"":" & TerrainQuoteJSON(environmentFile) & "}"
    TerrainWriteUTF8 folder & "\request.json", request
    Set worker = TerrainStartWorker(sourceRoot, folder & "\request.json")
    started = Now
    Do While worker.status = 0
        TerrainYield
        TerrainRequire DateDiff("s", started, Now) < 45, "Launcher test timed out; see " & folder
    Loop
    code = worker.ExitCode: status = TerrainReadUI(folder & "\status.txt")
    TerrainRequire code = 2, "Expected Python's invalid-schema exit code 2; actual=" & CStr(code)
    TerrainRequire InStr(1, status, "Unsupported request schema", vbTextCompare) > 0, "Python did not reach schema validation: " & status
    TerrainWriteUTF8 report, "{""status"":""passed"",""expected_worker_exit"":2,""job_folder"":" & TerrainQuoteJSON(folder) & _
        ",""worker_status"":" & TerrainQuoteJSON(status) & ",""network_used"":false,""dgn_writes"":false}"
    Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    TerrainCancelJob folder
    TerrainWriteUTF8 report, "{""status"":""failed"",""error"":" & TerrainQuoteJSON(message) & "}"
    On Error GoTo 0
End Sub

Public Sub RunTerrainVerticalMetadataSmoke()
    Dim report As String, beforeIDs As String, beforeSettings As String, message As String
    Dim selected() As Boolean, i As Long, count As Long, evidence As String
    report = Environ$("TEMP") & "\USGSTerrainDownloader-VerticalTests.json"
    On Error GoTo Failed
    TerrainWriteUTF8 report, "{""status"":""running""}"
    beforeIDs = ElementIDs(): beforeSettings = SettingsSnapshot()
    With frmUSGSTerrainDownloader
        TerrainRequire .Visible, "Open the actual search results and select metadata-supported products first."
        TerrainRequire .Controls("lstProducts").ListCount > 0, "Search results are required."
        ' This diagnostic selects all returned records; use a search with supported metadata.
        For i = 0 To .Controls("lstProducts").ListCount - 1
            .Controls("lstProducts").Selected(i) = True
        Next i
        ReDim selected(0 To .Controls("lstProducts").ListCount - 1)
        For i = 0 To UBound(selected)
            selected(i) = .Controls("lstProducts").Selected(i)
            If selected(i) Then count = count + 1
        Next i
        TerrainRequire count > 0, "Select at least one metadata-supported product."
        TerrainRequire .Controls("cboSourceZ").Value = "m", "Source elevation unit was not detected as metres."
        TerrainRequire .Controls("txtVertical").Text = "EPSG:5703", "NAVD88 vertical CRS was not filled."
        evidence = .Controls("txtEvidence").Text
        TerrainRequire UBound(Split(evidence, "USGS TNM source ")) = count, "Missing selected-product citations."
        TerrainRequire InStr(evidence, "All bare earth elevation values are in meters") > 0, "Source declaration was not retained."
        TerrainRequire Not .Controls("cboSourceZ").Enabled And .Controls("txtVertical").Locked And .Controls("txtEvidence").Locked, "Detected metadata fields must be protected."
        For i = 0 To UBound(selected)
            .Controls("lstProducts").Selected(i) = False
        Next i
        TerrainRequire .Controls("cboSourceZ").ListIndex = -1 And .Controls("txtVertical").Text = "" And .Controls("txtEvidence").Text = "", "Deselection left stale vertical declarations."
        TerrainRequire .Controls("cboSourceZ").Enabled And Not .Controls("txtVertical").Locked And Not .Controls("txtEvidence").Locked, "Manual fields did not unlock."
        For i = 0 To UBound(selected)
            .Controls("lstProducts").Selected(i) = selected(i)
        Next i
        TerrainRequire .Controls("txtEvidence").Text = evidence, "Restoring selection changed its citations."
    End With
    TerrainRequire ElementIDs() = beforeIDs And SettingsSnapshot() = beforeSettings, "DGN elements or settings changed."
    TerrainWriteUTF8 report, "{""status"":""passed"",""selected_count"":" & CStr(count) & ",""source_unit"":""m"",""vertical_crs"":""EPSG:5703"",""citations"":" & TerrainQuoteJSON(evidence) & ",""selection_clear_restore"":true,""settings_unchanged"":true}"
    Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    TerrainWriteUTF8 report, "{""status"":""failed"",""error"":" & TerrainQuoteJSON(message) & "}"
    On Error GoTo 0
End Sub

Public Sub RunTerrainCellSizeSmoke()
    Dim report As String, folder As String, beforeIDs As String, beforeSettings As String, message As String
    Dim fs As Object, stream As Object, baseline As String
    report = Environ$("TEMP") & "\USGSTerrainDownloader-CellSizeTests.json"
    On Error GoTo Failed
    TerrainWriteUTF8 report, "{""status"":""running""}"
    RunTerrainORDSmoke
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set stream = fs.OpenTextFile(Environ$("TEMP") & "\USGSTerrainDownloader-ORDTests.json", 1)
    baseline = stream.ReadAll: stream.Close
    TerrainRequire InStr(baseline, """status"":""passed""") > 0, "Native baseline failed."
    beforeIDs = ElementIDs(): beforeSettings = SettingsSnapshot()
    folder = Environ$("TEMP") & "\USGSTerrainDownloader-CellSizeFixtures"
    USGSTerrainDownloaderMain.Main
    ' The temporary form method only sets mSearchFolder and calls actual LoadProducts.
    With frmUSGSTerrainDownloader
        CallByName frmUSGSTerrainDownloader, "TerrainTestLoadProducts", VbMethod, folder & "\us-ft"
        TerrainRequire Abs(CDbl(.Controls("txtResolution").Text) - 3937# / 1200#) < 0.0000000001, "Survey-foot default failed."
        .Controls("cboOutput").Value = "m"
        TerrainRequire Abs(CDbl(.Controls("txtResolution").Text) - 3937# / 1200#) < 0.0000000001, "Elevation unit changed horizontal cell size."
        CallByName frmUSGSTerrainDownloader, "TerrainTestLoadProducts", VbMethod, folder & "\m"
        TerrainRequire CDbl(.Controls("txtResolution").Text) = 1, "New horizontal default did not replace the prior default."
        .Controls("txtResolution").Text = "5"
        CallByName frmUSGSTerrainDownloader, "TerrainTestLoadProducts", VbMethod, folder & "\us-ft"
        TerrainRequire .Controls("txtResolution").Text = "5", "Manual cell size was overwritten."
        .Controls("txtResolution").Text = ""
        CallByName frmUSGSTerrainDownloader, "TerrainTestLoadProducts", VbMethod, folder & "\us-ft"
        TerrainRequire Abs(CDbl(.Controls("txtResolution").Text) - 3937# / 1200#) < 0.0000000001, "Blank field was not filled."
        CallByName frmUSGSTerrainDownloader, "TerrainTestLoadProducts", VbMethod, folder & "\unknown"
        TerrainRequire .Controls("txtResolution").Text = "", "Unsupported dataset retained a stale automatic default."
        TerrainRequire Not .Controls("chkConfirmed").Value, "Grid mapping was automatically confirmed."
    End With
    TerrainClose
    Unload frmUSGSTerrainDownloader
    TerrainRequire ElementIDs() = beforeIDs And SettingsSnapshot() = beforeSettings, "DGN elements or settings changed."
    TerrainWriteUTF8 report, "{""status"":""passed"",""us_ft_default"":3.28083333333,""metric_default"":1,""manual_override_preserved"":true,""unknown_default_cleared"":true,""elevation_unit_independent"":true,""settings_unchanged"":true}"
    Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    TerrainClose
    Unload frmUSGSTerrainDownloader
    TerrainWriteUTF8 report, "{""status"":""failed"",""error"":" & TerrainQuoteJSON(message) & "}"
    On Error GoTo 0
End Sub
