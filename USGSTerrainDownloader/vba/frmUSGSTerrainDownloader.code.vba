Option Explicit

' Paste into a blank UserForm named frmUSGSTerrainDownloader. No designer controls required.
Private WithEvents cmdRectangle As MSForms.CommandButton
Private WithEvents cmdShape As MSForms.CommandButton
Private WithEvents cmdBrowse As MSForms.CommandButton
Private WithEvents cmdSearch As MSForms.CommandButton
Private WithEvents cmdProcess As MSForms.CommandButton
Private WithEvents cmdCancel As MSForms.CommandButton
Private WithEvents cmdOpen As MSForms.CommandButton
Private WithEvents lstProducts As MSForms.ListBox
Private txtCRS As MSForms.TextBox, txtResolution As MSForms.TextBox
Private txtVertical As MSForms.TextBox, txtEvidence As MSForms.TextBox
Private txtRoot As MSForms.TextBox, txtPython As MSForms.TextBox
Private txtEnv As MSForms.TextBox, txtOutput As MSForms.TextBox, txtCap As MSForms.TextBox
Private txtDetails As MSForms.TextBox
Private cboMaster As MSForms.ComboBox, cboOutput As MSForms.ComboBox
Private cboSourceZ As MSForms.ComboBox, cboDataset As MSForms.ComboBox
Private chkConfirmed As MSForms.CheckBox, chkClip As MSForms.CheckBox
Private chkOverlap As MSForms.CheckBox, chkGaps As MSForms.CheckBox, chkUnknown As MSForms.CheckBox
Private lblExtent As MSForms.Label, lblStatus As MSForms.Label, lblTotals As MSForms.Label
Private mBusy As Boolean, mJob As String, mSearchFolder As String, mLastFolder As String
Private mRows() As String

Private Function Label(ByVal name As String, ByVal caption As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, Optional ByVal h As Single = 18) As MSForms.Label
    Set Label = Me.Controls.Add("Forms.Label.1", name, True)
    With Label
        .caption = caption: .Left = x: .Top = y: .Width = w: .Height = h
    End With
End Function

Private Function Box(ByVal name As String, ByVal caption As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, Optional ByVal value As String = "") As MSForms.TextBox
    Dim heading As MSForms.Label
    Set heading = Label("lbl" & name, caption, x, y + 3, 120)
    Set Box = Me.Controls.Add("Forms.TextBox.1", name, True)
    With Box
        .Left = x + 122: .Top = y: .Width = w: .Height = 21: .value = value
    End With
End Function

Private Function Button(ByVal name As String, ByVal caption As String, ByVal x As Single, ByVal y As Single, ByVal w As Single) As MSForms.CommandButton
    Set Button = Me.Controls.Add("Forms.CommandButton.1", name, True)
    With Button
        .caption = caption: .Left = x: .Top = y: .Width = w: .Height = 26
    End With
End Function

Private Function Check(ByVal name As String, ByVal caption As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, Optional ByVal checked As Boolean = False) As MSForms.CheckBox
    Set Check = Me.Controls.Add("Forms.CheckBox.1", name, True)
    With Check
        .caption = caption: .Left = x: .Top = y: .Width = w: .Height = 22: .value = checked
    End With
End Function

Private Function Choices(ByVal name As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal values As String, Optional ByVal index As Long = -1) As MSForms.ComboBox
    Dim item As Variant
    Set Choices = Me.Controls.Add("Forms.ComboBox.1", name, True)
    With Choices
        .Left = x: .Top = y: .Width = w: .Height = 21: .Style = fmStyleDropDownList
        For Each item In Split(values, "|")
            .AddItem CStr(item)
        Next item
        .ListIndex = index
    End With
End Function

Private Sub UserForm_Initialize()
    Dim heading As MSForms.Label, folder As String, fs As Object
    Me.caption = "USGS Terrain Downloader"
    Me.Width = 806: Me.Height = 690: Me.ScrollHeight = 781: Me.ScrollBars = fmScrollBarsVertical
    Me.StartUpPosition = 1: Me.Font.Name = "Segoe UI": Me.Font.Size = 9
    Set heading = Label("lblTitle", "USGS TERRAIN DOWNLOADER", 18, 12, 740, 25)
    heading.Font.Size = 14: heading.Font.Bold = True
    Set heading = Label("lblBoundary", "BOUNDARY / PROJECT COORDINATES", 18, 45, 365)
    heading.Font.Bold = True
    Set cmdRectangle = Button("cmdRectangle", "Draw Rectangle", 18, 67, 165)
    Set cmdShape = Button("cmdShape", "Select Existing Shape", 195, 67, 185)
    Set heading = Label("lblGCS", TerrainDetectedGCS(), 18, 100, 365, 38)
    Set txtCRS = Box("txtCRS", "Source EPSG / WKT", 18, 141, 240)
    txtCRS.ControlTipText = "Exact horizontal projected CRS for model grid coordinates, including datum realization and foot definition."
    Set heading = Label("lblMaster", "Model master unit", 18, 173, 120)
    Set cboMaster = Choices("cboMaster", 140, 170, 240, "m|ft|us-ft")
    Set chkConfirmed = Check("chkConfirmed", "I verified CRS, master units and grid XY; no local offsets, rotation, ground scale or GCS modifiers.", 18, 198, 365)
    chkConfirmed.Height = 38
    Set lblExtent = Label("lblExtent", TerrainModelDescription() & vbCrLf & "Select boundary; Search calculates WGS84 extent.", 18, 244, 365, 52)
    Set heading = Label("lblProcessing", "PROCESSING / VERTICAL EVIDENCE", 410, 45, 365)
    heading.Font.Bold = True
    Set chkClip = Check("chkClip", "Clip to original selected boundary", 410, 67, 365, True)
    Set heading = Label("lblSourceZ", "Source elevation unit", 410, 100, 120)
    Set cboSourceZ = Choices("cboSourceZ", 532, 97, 240, "m|ft|us-ft")
    Set txtVertical = Box("txtVertical", "Source vertical CRS", 410, 125, 240)
    txtVertical.ControlTipText = "For example EPSG:5703 ONLY when NAVD88 is verified in selected source metadata. No datum shift is performed."
    Set txtEvidence = Box("txtEvidence", "Evidence / metadata", 410, 153, 240)
    Set heading = Label("lblOutputZ", "Output elevation unit", 410, 184, 120)
    Set cboOutput = Choices("cboOutput", 532, 181, 240, "m|ft|us-ft", 0)
    Set txtResolution = Box("txtResolution", "Cell size (CRS units)", 410, 209, 240)
    Set chkOverlap = Check("chkOverlap", "Accept overlaps: last selected list row wins", 410, 237, 365)
    Set chkGaps = Check("chkGaps", "Allow/report NoData gaps inside boundary", 410, 263, 365)
    Set heading = Label("lblFixed", "Compatible tiles are mosaicked and reprojected to source DGN CRS. Originals are preserved. Not survey-grade. Fence and automatic ORD attachment are not enabled.", 18, 303, 754, 28)
    Set heading = Label("lblDataset", "USGS data source", 18, 341, 120)
    Set cboDataset = Choices("cboDataset", 140, 338, 320, "3DEP 1-meter DEM|3DEP 1/3-arc-second DEM|LiDAR LAS/LAZ (download only)", 0)
    Set txtCap = Box("txtCap", "Download cap (GiB)", 480, 338, 170, "5")
    Set heading = Label("lblColumns", "Select products after Search. Columns: product / resolution / publication date / estimated MiB", 18, 370, 754)
    Set lstProducts = Me.Controls.Add("Forms.ListBox.1", "lstProducts", True)
    With lstProducts
        .Left = 18: .Top = 391: .Width = 754: .Height = 96
        .ColumnCount = 4: .ColumnWidths = "440 pt;90 pt;100 pt;95 pt"
        .MultiSelect = fmMultiSelectMulti
    End With
    Set txtDetails = Me.Controls.Add("Forms.TextBox.1", "txtDetails", True)
    With txtDetails
        .Left = 18: .Top = 495: .Width = 754: .Height = 56
        .MultiLine = True: .Locked = True: .ScrollBars = fmScrollBarsBoth
        .value = "Product metadata will appear here. Publication date is not an accuracy ranking. Verify source vertical datum/unit before processing."
    End With
    Set lblTotals = Label("lblTotals", "No products selected.", 18, 559, 450)
    Set chkUnknown = Check("chkUnknown", "Accept unknown sizes within hard cap", 480, 555, 292)
    Set txtRoot = Box("txtRoot", "Source code folder", 18, 585, 632, GetSetting("USGSTerrainDownloader", "Runtime", "SourceRoot", ""))
    txtRoot.ControlTipText = "Full path to USGSTerrainDownloader folder containing python\launch_worker.ps1."
    Set txtPython = Box("txtPython", "Python executable", 18, 613, 632, GetSetting("USGSTerrainDownloader", "Runtime", "Python", ""))
    Set txtEnv = Box("txtEnv", "GIS environment JSON", 18, 641, 632, GetSetting("USGSTerrainDownloader", "Runtime", "EnvironmentFile", ""))
    txtEnv.ControlTipText = "Optional child environment JSON. See python\osgeo4w-environment.example.json."
    folder = TerrainDefaultDestination()
    Set txtOutput = Box("txtOutput", "Destination folder", 18, 669, 540, folder)
    Set cmdBrowse = Button("cmdBrowse", "Browse...", 688, 666, 84)
    Set lblStatus = Label("lblStatus", "Ready. Draw/select boundary, confirm coordinates, then Search USGS.", 18, 703, 754, 28)
    Set cmdSearch = Button("cmdSearch", "Search USGS", 18, 741, 140)
    Set cmdProcess = Button("cmdProcess", "Download / Process", 170, 741, 170)
    Set cmdOpen = Button("cmdOpen", "Open Output Folder", 352, 741, 170)
    Set cmdCancel = Button("cmdCancel", "Cancel / Close", 632, 741, 140)
    cmdCancel.Cancel = True
    cmdRectangle.Accelerator = "R": cmdShape.Accelerator = "E"
    cmdSearch.Accelerator = "S": cmdProcess.Accelerator = "D"
    cmdOpen.Accelerator = "O": cmdBrowse.Accelerator = "B"
End Sub

Public Sub BoundaryChanged(ByVal message As String)
    mSearchFolder = "": lstProducts.Clear
    lblExtent.caption = message: lblTotals.caption = "Search required for new boundary."
End Sub

Private Sub cmdRectangle_Click()
    If Not mBusy Then TerrainPick True
End Sub
Private Sub cmdShape_Click()
    If Not mBusy Then TerrainPick False
End Sub
Private Sub cmdBrowse_Click()
    Dim path As String
    path = TerrainBrowseFolder("Choose a writable project terrain-data folder")
    If Len(path) > 0 Then txtOutput.value = path
End Sub

Private Function DatasetKey() As String
    Select Case cboDataset.ListIndex
        Case 0: DatasetKey = "dem1m"
        Case 1: DatasetKey = "dem13"
        Case 2: DatasetKey = "lidar"
        Case Else: TerrainRequire False, "Select a USGS dataset."
    End Select
End Function

Private Function RequestJSON(ByVal action As String) As String
    Dim result As String, ids As String, i As Long, fields As Variant, cap As Double, resolution As Double
    TerrainCheckContext True
    TerrainRequire chkConfirmed.value And Len(Trim$(txtCRS.Text)) > 0 And cboMaster.ListIndex >= 0, "Explicitly verify source CRS, model master units and grid-coordinate mapping."
    result = "{""schema_version"":1,""action"":" & TerrainQuoteJSON(action) & _
        ",""boundary"":" & TerrainBoundaryJSON() & ",""source_crs"":" & TerrainQuoteJSON(Trim$(txtCRS.Text)) & _
        ",""master_unit"":" & TerrainQuoteJSON(cboMaster.value) & ",""crs_confirmed"":true,""coordinates_are_grid"":true" & _
        ",""dataset"":" & TerrainQuoteJSON(DatasetKey()) & ",""python_executable"":" & TerrainQuoteJSON(Trim$(txtPython.Text)) & _
        ",""environment_file"":" & TerrainQuoteJSON(Trim$(txtEnv.Text)) & ",""dgn_context"":" & TerrainQuoteJSON(TerrainModelDescription()) & _
        ",""dgn_gcs_probe"":" & TerrainQuoteJSON(TerrainDetectedGCS())
    If action = "process" Then
        TerrainRequire Len(mSearchFolder) > 0, "Search USGS first."
        For i = 0 To lstProducts.ListCount - 1
            If lstProducts.Selected(i) Then
                fields = Split(mRows(i), vbTab)
                If Len(ids) > 0 Then ids = ids & ","
                ids = ids & TerrainQuoteJSON(CStr(fields(0)))
            End If
        Next i
        TerrainRequire Len(ids) > 0, "Select at least one product."
        TerrainRequire IsNumeric(txtCap.Text), "Download cap must be numeric GiB."
        cap = CDbl(txtCap.Text)
        TerrainRequire cap > 0 And cap <= 100, "Download cap must be > 0 and <= 100 GiB."
        result = result & ",""selected_ids"": [" & ids & "],""search_manifest"":" & TerrainQuoteJSON(mSearchFolder & "\search.json") & _
            ",""max_download_bytes"":" & Format$(Fix(cap * 1073741824#), "0") & _
            ",""unknown_size_confirmed"":" & TerrainBool(chkUnknown.value)
        If DatasetKey() <> "lidar" Then
            TerrainRequire cboSourceZ.ListIndex >= 0 And Len(Trim$(txtVertical.Text)) > 0 And Len(Trim$(txtEvidence.Text)) >= 12, "Verify selected sources' vertical datum/unit and enter source metadata evidence. Mixed datums are unsupported."
            TerrainRequire IsNumeric(txtResolution.Text), "Cell size must be numeric in source CRS native horizontal units."
            resolution = CDbl(txtResolution.Text)
            TerrainRequire resolution > 0, "Cell size must be positive."
            result = result & ",""source_vertical"": {""unit"":" & TerrainQuoteJSON(cboSourceZ.value) & _
                ",""crs"":" & TerrainQuoteJSON(Trim$(txtVertical.Text)) & ",""evidence"":" & TerrainQuoteJSON(Trim$(txtEvidence.Text)) & "}" & _
                ",""resolution"":" & TerrainNumber(resolution) & ",""elevation_unit"":" & TerrainQuoteJSON(cboOutput.value) & _
                ",""clip"":" & TerrainBool(chkClip.value) & ",""overlap_order_confirmed"":" & TerrainBool(chkOverlap.value) & _
                ",""allow_gaps"":" & TerrainBool(chkGaps.value)
        End If
    End If
    RequestJSON = result & "}"
End Function

Private Sub SetBusy(ByVal busy As Boolean)
    Dim control As MSForms.Control
    mBusy = busy
    For Each control In Me.Controls
        If TypeName(control) <> "Label" Then control.Enabled = Not busy
    Next control
    cmdCancel.Enabled = True
End Sub

Private Sub RunWorker(ByVal action As String)
    Dim request As String, worker As Object, text As String, errorText As String
    On Error GoTo Failed
    If mBusy Then Exit Sub
    request = RequestJSON(action)
    SaveSetting "USGSTerrainDownloader", "Runtime", "SourceRoot", Trim$(txtRoot.Text)
    SaveSetting "USGSTerrainDownloader", "Runtime", "Python", Trim$(txtPython.Text)
    SaveSetting "USGSTerrainDownloader", "Runtime", "EnvironmentFile", Trim$(txtEnv.Text)
    mJob = TerrainNewJobFolder(Trim$(txtOutput.Text)): mLastFolder = mJob
    TerrainWriteUTF8 mJob & "\request.json", request
    If action = "search" Then mSearchFolder = "": lstProducts.Clear
    SetBusy True
    lblStatus.caption = "Starting " & action & "..."
    Set worker = TerrainStartWorker(Trim$(txtRoot.Text), mJob & "\request.json")
    Do While worker.status = 0
        text = TerrainReadUI(mJob & "\status.txt")
        If Len(text) > 0 Then lblStatus.caption = Replace(text, vbCrLf, ": ")
        TerrainYield
    Loop
    text = TerrainReadUI(mJob & "\status.txt")
    If Len(text) > 0 Then lblStatus.caption = Replace(text, vbCrLf, ": ")
    TerrainRequire worker.ExitCode = 0, "Worker failed/cancelled. " & text & vbCrLf & "See " & mJob & "\error.log. If Python did not start, verify environment paths and PowerShell policy."
    If action = "search" Then
        mSearchFolder = mJob
        LoadProducts
        lblExtent.caption = TerrainReadUI(mJob & "\extent.txt")
    End If
    SetBusy False
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    If Not worker Is Nothing Then
        If worker.status = 0 Then
            TerrainCancelJob mJob
            Do While worker.status = 0
                TerrainYield
            Loop
        End If
    End If
    On Error GoTo 0
    SetBusy False
    lblStatus.caption = errorText
    MsgBox errorText, vbExclamation, "USGS Terrain Downloader"
End Sub

Private Sub LoadProducts()
    Dim contents As String, i As Long, fields As Variant, size As String
    lstProducts.Clear
    contents = TerrainReadUI(mSearchFolder & "\products.tsv")
    If Len(contents) = 0 Then Exit Sub
    mRows = Split(contents, vbCrLf)
    For i = 0 To UBound(mRows)
        fields = Split(mRows(i), vbTab)
        TerrainRequire UBound(fields) = 8, "Invalid products.tsv row."
        size = "unknown"
        If IsNumeric(fields(2)) Then size = Format$(CDbl(fields(2)) / 1048576#, "0.0")
        lstProducts.AddItem fields(1)
        lstProducts.List(i, 1) = fields(4): lstProducts.List(i, 2) = fields(3): lstProducts.List(i, 3) = size
    Next i
    RefreshTotals
End Sub

Private Sub RefreshTotals()
    Dim i As Long, count As Long, total As Double, unknown As Long, fields As Variant
    For i = 0 To lstProducts.ListCount - 1
        If lstProducts.Selected(i) Then
            count = count + 1: fields = Split(mRows(i), vbTab)
            If IsNumeric(fields(2)) Then total = total + CDbl(fields(2)) Else unknown = unknown + 1
        End If
    Next i
    lblTotals.caption = CStr(count) & " / " & CStr(lstProducts.ListCount) & " selected; known estimate " & Format$(total / 1073741824#, "0.000") & " GiB; " & CStr(unknown) & " unknown sizes."
End Sub

Private Sub lstProducts_Change()
    Dim fields As Variant
    If lstProducts.ListIndex >= 0 Then
        fields = Split(mRows(lstProducts.ListIndex), vbTab)
        txtDetails.value = "Coverage: " & fields(5) & vbCrLf & "Metadata: " & fields(7) & vbCrLf & fields(8)
    End If
    RefreshTotals
End Sub
Private Sub cmdSearch_Click()
    RunWorker "search"
End Sub
Private Sub cmdProcess_Click()
    RunWorker "process"
End Sub
Private Sub cmdOpen_Click()
    On Error GoTo Failed
    If Len(mLastFolder) > 0 Then TerrainOpenFolder mLastFolder Else TerrainOpenFolder txtOutput.Text
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Output folder"
End Sub
Private Sub cmdCancel_Click()
    If mBusy Then
        TerrainCancelJob mJob
        lblStatus.caption = "Cancellation requested. Waiting for current network/GDAL operation to stop..."
    Else
        Unload Me
    End If
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mBusy Then
        Cancel = True
        TerrainCancelJob mJob
        lblStatus.caption = "Cancellation requested. Dialog will remain until worker stops."
    Else
        TerrainClose
    End If
End Sub
