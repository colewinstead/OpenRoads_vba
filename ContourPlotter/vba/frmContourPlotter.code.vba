Option Explicit
' SPDX-License-Identifier: MIT
' Paste into a NEW blank UserForm named frmContourPlotter.
Private mControls As Object, mSurfaceNames As Collection
Private WithEvents cmdBrowse As MSForms.CommandButton
Private WithEvents cmdEngine As MSForms.CommandButton
Private WithEvents cmdInspect As MSForms.CommandButton
Private WithEvents cmdPreview As MSForms.CommandButton
Private WithEvents cmdGenerate As MSForms.CommandButton
Private WithEvents cmdRegenerate As MSForms.CommandButton
Private WithEvents cmdCancel As MSForms.CommandButton
Private WithEvents cmdPrevious As MSForms.CommandButton
Private WithEvents chkByLevel As MSForms.CheckBox
Private mInspectedSource As String

Private Function AddControl(parent As Object, ByVal kind As String, ByVal name As String, _
        ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single) As Object
    Dim c As Object
    Set c = parent.Controls.Add("Forms." & kind & ".1", name, True)
    c.Left = x: c.Top = y: c.Width = width: c.Height = height
    mControls.Add name, c
    Set AddControl = c
End Function

Private Sub Label(parent As Object, ByVal name As String, ByVal value As String, _
        ByVal x As Single, ByVal y As Single, ByVal width As Single, Optional ByVal height As Single = 20)
    Dim c As Object
    Set c = AddControl(parent, "Label", name, x, y, width, height)
    c.Caption = value
End Sub

Private Function TextBox(parent As Object, ByVal name As String, ByVal caption As String, _
        ByVal y As Single, ByVal value As String, Optional ByVal width As Single = 450) As Object
    Label parent, "lbl" & name, caption, 18, y + 3, 190
    Set TextBox = AddControl(parent, "TextBox", name, 220, y, width, 22)
    TextBox.Value = value
End Function

Private Function Combo(parent As Object, ByVal name As String, ByVal caption As String, _
        ByVal y As Single, ByVal values As String) As MSForms.ComboBox
    Dim c As MSForms.ComboBox, item As Variant
    Label parent, "lbl" & name, caption, 18, y + 3, 190
    Set c = AddControl(parent, "ComboBox", name, 220, y, 450, 22)
    c.Style = fmStyleDropDownList
    If Len(values) > 0 Then
        For Each item In Split(values, "|")
            c.AddItem CStr(item)
        Next item
    End If
    Set Combo = c
End Function

Private Function Button(parent As Object, ByVal name As String, ByVal caption As String, _
        ByVal x As Single, ByVal y As Single, ByVal width As Single) As MSForms.CommandButton
    Set Button = AddControl(parent, "CommandButton", name, x, y, width, 27)
    Button.Caption = caption
End Function

Private Function Check(parent As Object, ByVal name As String, ByVal caption As String, _
        ByVal y As Single, ByVal initial As Boolean) As MSForms.CheckBox
    Set Check = AddControl(parent, "CheckBox", name, 18, y, 650, 24)
    Check.Caption = caption: Check.Value = initial
End Function

Private Sub UserForm_Initialize()
    Dim pages As MSForms.MultiPage, p As Object, c As Object
    Set mControls = CreateObject("Scripting.Dictionary")
    Set mSurfaceNames = New Collection
    Me.Caption = "Contour Plotter | OpenRoads Designer"
    Me.Width = 742: Me.Height = 644
    Me.StartUpPosition = 1
    Me.Font.Name = "Segoe UI": Me.Font.Size = 10
    Label Me, "heading", "CONTOUR PLOTTER", 18, 14, 688, 25
    mControls("heading").Font.Bold = True: mControls("heading").Font.Size = 14
    Label Me, "target", CPFileName & vbCrLf & "Model: " & CPModelName & " | master units: " & CPUnitName, 18, 44, 688, 38
    Set pages = AddControl(Me, "MultiPage", "pages", 12, 90, 708, 432)
    Do While pages.Pages.Count < 3
        pages.Pages.Add
    Loop
    pages.Pages(0).Caption = "1  Source and coordinates"
    pages.Pages(1).Caption = "2  Contours and labels"
    pages.Pages(2).Caption = "3  MDOT levels and output"
    Set p = pages.Pages(0)
    Set c = TextBox(p, "source", "Merged terrain LandXML", 18, "", 338)
    Set cmdBrowse = Button(p, "browse", "Browse...", 566, 16, 104)
    Set c = TextBox(p, "engine", "ContourPlotter.exe", 50, GetSetting(CP_APP, "Settings", "Engine", ""), 338)
    Set cmdEngine = Button(p, "engineBrowse", "Browse...", 566, 48, 104)
    Set c = Combo(p, "surface", "Source surface", 88, "")
    Set cmdInspect = Button(p, "inspect", "Inspect LandXML", 220, 121, 160)
    Label p, "sourceStats", "Export the already merged terrain to LandXML in ORD. Native terrain export is manual in this version.", 18, 157, 652, 38
    Set c = Combo(p, "axis", "Stored point axis order", 202, "E N Z (xy)|N E Z (ne)")
    Set c = Combo(p, "horizontal", "Source horizontal unit", 232, "m|ft|us-ft")
    Set c = Combo(p, "vertical", "Source vertical unit", 262, "m|ft|us-ft")
    Set c = TextBox(p, "crs", "Verified CRS / local grid", 292, "")
    Set c = TextBox(p, "datum", "Verified vertical datum", 322, "")
    Label p, "gcs", CPGcsName & vbCrLf & "ft = international foot; us-ft = US survey foot. No reprojection or datum conversion.", 18, 354, 652, 40

    Set p = pages.Pages(1)
    Set c = TextBox(p, "minor", "Minor interval (source Z unit)", 18, "1")
    Set c = TextBox(p, "major", "Major interval (source Z unit)", 48, "5")
    Set c = TextBox(p, "base", "Contour base elevation", 78, "0")
    Set c = Combo(p, "labels", "Automatic labels", 108, "Major only|Major and minor|None")
    c.ListIndex = 0
    Set c = TextBox(p, "spacing", "Label spacing (model units)", 138, "250")
    Set c = TextBox(p, "scale", "Model units per paper inch", 168, "50")
    Set c = TextBox(p, "paperMajor", "Major label height (paper in)", 198, "0.10", 160)
    Label p, "lblPaperMinor", "Minor (in)", 396, 201, 80
    Set c = AddControl(p, "TextBox", "paperMinor", 486, 198, 184, 22)
    c.Value = "0.10"
    Set c = TextBox(p, "clearance", "Paper gap clearance (inches)", 228, "0.02")
    Set c = Combo(p, "dimension", "Geometry dimension", 258, "2D (Z = 0)|3D (contour elevation)")
    c.ListIndex = 0
    Set c = TextBox(p, "maxLabels", "Maximum labels per path", 288, "100")
    Set c = Check(p, "gaps", "Break contour under elevation labels", 322, True)
    Label p, "scaleHelp", "Example: 0.10 in x 50 ft/in = 5 ft model text height. For a metric 1:500 plot, enter 12.7 m/in. Smoothing is disabled. Short fragments are reported, never removed.", 18, 354, 652, 40

    Set p = pages.Pages(2)
    Set c = Combo(p, "majorLevel", "Major contour level", 18, ""): CPFillLevels c, "majorLevel"
    Set c = Combo(p, "minorLevel", "Minor contour level", 48, ""): CPFillLevels c, "minorLevel"
    Set c = Combo(p, "majorLabel", "Major label level", 78, ""): CPFillLevels c, "majorLabel"
    Set c = Combo(p, "minorLabel", "Minor label level", 108, ""): CPFillLevels c, "minorLabel"
    Set c = Combo(p, "style", "DGN / library text style", 138, ""): CPFillStyles c
    Set chkByLevel = Check(p, "byLevel", "Use Existing Level Symbology", 173, True)
    Set c = TextBox(p, "color", "Override color (0-255)", 207, "0")
    Set c = TextBox(p, "weight", "Override weight (0-31)", 237, "0")
    Set c = TextBox(p, "lineStyle", "Override line style name", 267, "")
    Set c = TextBox(p, "previous", "Previous run journal (.cpj)", 305, _
        GetSetting(CP_APP, "LastRuns", CPFileName & "|" & CPModelName, ""), 338)
    Set cmdPrevious = Button(p, "previousBrowse", "Browse...", 566, 303, 104)
    Label p, "stagedHelp", "Regenerate stages a complete replacement, then asks before deleting recorded prior members, including manual edits. Keep both run folders until the DGN is saved and checked.", 18, 345, 652, 48
    EnableOverrides
    Label Me, "status", "Inspect the LandXML, verify coordinates, then choose existing levels and a text style.", 18, 532, 688, 36
    Set cmdPreview = Button(Me, "preview", "Analyze / Preview", 18, 575, 156)
    Set cmdGenerate = Button(Me, "generate", "Generate", 344, 575, 120)
    Set cmdRegenerate = Button(Me, "regenerate", "Regenerate", 474, 575, 120)
    Set cmdCancel = Button(Me, "cancel", "Close", 604, 575, 102)
    cmdCancel.Cancel = True
End Sub

Public Sub SetStatus(ByVal message As String)
    mControls("status").Caption = message
    Me.Repaint
End Sub

Private Sub EnableOverrides()
    Dim name As Variant
    For Each name In Array("color", "weight", "lineStyle")
        mControls(name).Enabled = Not chkByLevel.Value
    Next name
End Sub

Private Sub chkByLevel_Click()
    If Not mControls Is Nothing Then EnableOverrides
End Sub

Private Sub cmdEngine_Click()
    Dim value As String
    value = CPBrowse("Select packaged ContourPlotter.exe", CStr(mControls("engine").Value), "Contour engine (*.exe)|*.exe|All files (*.*)|*.*")
    If Len(value) > 0 Then
        mControls("engine").Value = value
        SaveSetting CP_APP, "Settings", "Engine", value
    End If
End Sub

Private Sub cmdBrowse_Click()
    Dim value As String
    value = CPBrowse("Select already merged terrain LandXML", CStr(mControls("source").Value), "LandXML files (*.xml)|*.xml|All files (*.*)|*.*")
    If Len(value) > 0 Then mControls("source").Value = value
End Sub

Private Sub cmdPrevious_Click()
    Dim value As String
    value = CPBrowse("Select prior elements.cpj from its run folder", CStr(mControls("previous").Value), "Contour journals (*.cpj)|*.cpj|All files (*.*)|*.*")
    If Len(value) > 0 Then mControls("previous").Value = value
End Sub

Private Sub Busy(ByVal value As Boolean)
    Dim name As Variant
    CPBusy = value
    For Each name In mControls.Keys
        If name <> "cancel" And name <> "status" Then mControls(name).Enabled = Not value
    Next name
    If value Then cmdCancel.Caption = "Cancel" Else cmdCancel.Caption = "Close"
    If Not value Then EnableOverrides
End Sub

Private Sub cmdInspect_Click()
    Dim runID As String, job As String, lines As Variant, row As Variant, i As Long, message As String
    On Error GoTo Failed
    If CPBusy Then Exit Sub
    CPCancelled = False: Busy True
    CPPreflight
    mControls("source").Value = CPPath(CStr(mControls("source").Value))
    runID = CPNewRunID(): job = CPNewJob(runID)
    CPStatus "Inspecting source surfaces and declared metadata..."
    CPRunEngine CStr(mControls("engine").Value), job, CPRequest("inspect", runID, job, CStr(mControls("source").Value), "{}", "{}")
    lines = Split(Replace(CPReadUnicode(job & "\surfaces.tsv"), vbCr, ""), vbLf)
    row = Split(CStr(lines(0)), vbTab)
    CPRequire UBound(row) = 3, "Invalid surface inventory."
    CPRequire row(0) = "INSPECT1", "Invalid inspection version."
    SetUnit "horizontal", CStr(row(1))
    SetUnit "vertical", CStr(row(2))
    mControls("crs").Value = row(3)
    mControls("surface").Clear
    Set mSurfaceNames = New Collection
    For i = 1 To UBound(lines)
        If Len(lines(i)) > 0 Then
            row = Split(CStr(lines(i)), vbTab)
            CPRequire UBound(row) = 3, "Invalid surface inventory row."
            mSurfaceNames.Add CStr(row(1))
            mControls("surface").AddItem row(1) & " | " & row(2) & " points, " & row(3) & " faces"
        End If
    Next i
    If mSurfaceNames.Count = 1 Then mControls("surface").ListIndex = 0
    mInspectedSource = CPPath(CStr(mControls("source").Value))
    SetStatus "Choose a surface and resolve axis order, units, CRS and vertical datum. Face counts include invisible source faces until analysis."
    Busy False
    Exit Sub
Failed:
    message = Err.Description
    Busy False
    SetStatus message
    MsgBox message, vbExclamation, "Contour Plotter"
End Sub

Private Sub SetUnit(ByVal name As String, ByVal value As String)
    Dim unit As String, i As Long
    Select Case value
        Case "meter", "metre": unit = "m"
        Case "foot", "internationalFoot": unit = "ft"
        Case "USSurveyFoot": unit = "us-ft"
    End Select
    mControls(name).ListIndex = -1
    For i = 0 To mControls(name).ListCount - 1
        If mControls(name).List(i) = unit Then mControls(name).ListIndex = i
    Next i
End Sub

Private Function Number(ByVal name As String) As Double
    Dim value As String
    value = Trim$(CStr(mControls(name).Value))
    CPRequire Len(value) > 0 And IsNumeric(value), "Enter a valid number for " & name & "."
    Number = CDbl(value)
    CPRequire Abs(Number) <= 1000000000#, "Input exceeds the safety range: " & name
End Function

Private Function Spatial() As String
    Dim index As Long, axis As String
    CPRequire CPPath(CStr(mControls("source").Value)) = mInspectedSource And Len(mInspectedSource) > 0, "Inspect this LandXML before proceeding."
    index = mControls("surface").ListIndex
    CPRequire index >= 0 And index < mSurfaceNames.Count, "Select the named LandXML surface."
    CPRequire mControls("axis").ListIndex >= 0, "Explicitly select stored axis order."
    If mControls("axis").ListIndex = 0 Then axis = "xy" Else axis = "ne"
    CPRequire mControls("horizontal").ListIndex >= 0 And mControls("vertical").ListIndex >= 0, "Resolve horizontal and vertical units."
    CPRequire Len(Trim$(CStr(mControls("crs").Value))) > 0, "Enter a verified CRS/local grid."
    CPRequire Len(Trim$(CStr(mControls("datum").Value))) > 0, "Enter a verified vertical datum."
    Spatial = "{""surface_index"":" & CStr(index) & ",""surface_name"":" & CPJson(CStr(mSurfaceNames(index + 1))) & _
        ",""axis_order"":" & CPJson(axis) & ",""horizontal_unit"":" & CPJson(CStr(mControls("horizontal").Value)) & _
        ",""vertical_unit"":" & CPJson(CStr(mControls("vertical").Value)) & ",""model_unit"":" & CPJson(CPUnitName) & _
        ",""crs"":" & CPJson(CStr(mControls("crs").Value)) & ",""vertical_datum"":" & CPJson(CStr(mControls("datum").Value)) & "}"
End Function

Private Function Settings(ByVal widths As String, ByVal boxHeight As Double, ByVal minorBoxHeight As Double) As String
    Dim mode As String, gaps As String, maxLabels As Double
    Select Case mControls("labels").ListIndex
        Case 0: mode = "major"
        Case 1: mode = "both"
        Case 2: mode = "none"
        Case Else: CPRequire False, "Select the label mode."
    End Select
    If mControls("gaps").Value Then gaps = "true" Else gaps = "false"
    maxLabels = Number("maxLabels")
    CPRequire maxLabels = Fix(maxLabels) And maxLabels >= 1 And maxLabels <= 10000, "Maximum labels must be an integer in 1..10000."
    Settings = "{""minor"":" & CPNumber(Number("minor")) & ",""major"":" & CPNumber(Number("major")) & _
        ",""base"":" & CPNumber(Number("base")) & ",""label_mode"":" & CPJson(mode) & _
        ",""label_spacing"":" & CPNumber(Number("spacing")) & ",""paper_height_major"":" & CPNumber(Number("paperMajor")) & _
        ",""paper_height_minor"":" & CPNumber(Number("paperMinor")) & _
        ",""scale"":" & CPNumber(Number("scale")) & ",""clearance_paper"":" & CPNumber(Number("clearance")) & _
        ",""dimension"":" & CStr(mControls("dimension").ListIndex + 2) & ",""break_lines"":" & gaps & _
        ",""max_labels"":" & CStr(CLng(maxLabels)) & ",""text_widths"":" & widths
    If boxHeight > 0 Then Settings = Settings & ",""text_box_height"":" & CPNumber(boxHeight)
    If minorBoxHeight > 0 Then Settings = Settings & ",""text_box_height_minor"":" & CPNumber(minorBoxHeight)
    Settings = Settings & "}"
End Function

Private Sub cmdPreview_Click()
    Workflow False, False
End Sub

Private Sub cmdGenerate_Click()
    Workflow True, False
End Sub

Private Sub cmdRegenerate_Click()
    Workflow True, True
End Sub

Private Sub Workflow(ByVal insert As Boolean, ByVal regenerate As Boolean)
    Dim runID As String, job As String, spatialJSON As String, settingsJSON As String, request As String
    Dim analysis As Variant, widths As String, boxHeight As Double, height As Double, minorHeight As Double, minorBoxHeight As Double
    Dim prior As String, message As String, color As Long, weight As Long, item As Variant
    Dim level As Level, style As TextStyle, h As Integer, header As String, fields As Variant
    On Error GoTo Failed
    If CPBusy Then Exit Sub
    CPCancelled = False
    CPPreflight
    spatialJSON = Spatial()
    settingsJSON = Settings("{}", 0#, 0#)
    If regenerate Then
        prior = CPPath(CStr(mControls("previous").Value))
        CPRequire CreateObject("Scripting.FileSystemObject").FileExists(prior), "Select a previous committed elements.cpj."
    End If
    Busy True
    runID = CPNewRunID(): job = CPNewJob(runID)
    CPStatus "Analyzing source triangles and contour elevations..."
    CPRunEngine CStr(mControls("engine").Value), job, CPRequest("analyze", runID, job, CStr(mControls("source").Value), spatialJSON, settingsJSON)
    analysis = Split(Replace(Replace(CPReadUnicode(job & "\analysis.tsv"), vbCr, ""), vbLf, ""), vbTab)
    CPRequire UBound(analysis) = 13, "Invalid analysis metadata."
    height = CPParseNumber(CStr(analysis(10))): minorHeight = CPParseNumber(CStr(analysis(11)))
    mControls("sourceStats").Caption = CStr(analysis(0)) & " | " & CStr(analysis(1)) & " points | " & _
        CStr(analysis(2)) & " visible faces | Z " & CStr(analysis(3)) & " to " & CStr(analysis(4)) & " " & CStr(analysis(12))
    widths = "{}": boxHeight = height: minorBoxHeight = minorHeight
    If mControls("labels").ListIndex <> 2 Then
        CPStatus "Measuring elevation strings with the selected DGN text style..."
        widths = CPMeasuredWidths(job, CStr(mControls("style").Value), height, minorHeight, boxHeight, minorBoxHeight)
    End If
    settingsJSON = Settings(widths, boxHeight, minorBoxHeight)
    request = CPRequest("generate", runID, job, CStr(mControls("source").Value), spatialJSON, settingsJSON)
    request = Left$(request, Len(request) - 1) & ",""expected_source_sha256"":" & CPJson(CStr(analysis(13))) & "}"
    CPStatus "Generating connected contours, labels and physical gaps..."
    CPRunEngine CStr(mControls("engine").Value), job, request
    CPPreflight
    h = FreeFile
    Open job & "\geometry.cp2" For Input Access Read Lock Write As #h
    Line Input #h, header
    Close #h: h = 0
    fields = Split(header, vbTab)
    CPRequire UBound(fields) = 13, "Invalid output header."
    If Not insert Then
        SetStatus "Preview: " & CStr(fields(4)) & " native line strings; " & CStr(fields(5)) & " labels. No DGN writes. " & job
        Busy False
        Exit Sub
    End If
    Set level = CPLevel(CStr(mControls("majorLevel").Value))
    Set level = CPLevel(CStr(mControls("minorLevel").Value))
    If mControls("labels").ListIndex <> 2 Then Set level = CPLevel(CStr(mControls("majorLabel").Value))
    If mControls("labels").ListIndex = 1 Then Set level = CPLevel(CStr(mControls("minorLabel").Value))
    If Not chkByLevel.Value Then
        CPRequire Number("color") = Fix(Number("color")) And Number("weight") = Fix(Number("weight")), "Overrides must be integers."
        color = CLng(Number("color")): weight = CLng(Number("weight"))
    End If
    message = CPImport(job & "\geometry.cp2", job, runID, CStr(mControls("majorLevel").Value), _
        CStr(mControls("minorLevel").Value), CStr(mControls("majorLabel").Value), CStr(mControls("minorLabel").Value), _
        CStr(mControls("style").Value), height, minorHeight, boxHeight, minorBoxHeight, CBool(chkByLevel.Value), color, weight, CStr(mControls("lineStyle").Value), _
        "Declared source CRS: " & CStr(analysis(5)) & vbCrLf & "Confirmed grid: " & CStr(mControls("crs").Value) & _
        vbCrLf & "Vertical datum: " & CStr(mControls("datum").Value) & " | Axis: " & CStr(mControls("axis").Value), CLng(mControls("dimension").ListIndex + 2))
    SaveSetting CP_APP, "Settings", "Engine", CStr(mControls("engine").Value)
    SaveSetting CP_APP, "Settings", "TextStyle", CStr(mControls("style").Value)
    For Each item In Array("majorLevel", "minorLevel", "majorLabel", "minorLabel")
        SaveSetting CP_APP, "Levels", CStr(item), CStr(mControls(item).Value)
    Next item
    SaveSetting CP_APP, "LastRuns", CPFileName & "|" & CPModelName, job & "\elements.cpj"
    mControls("previous").Value = job & "\elements.cpj"
    SetStatus message
    If regenerate Then CPRemovePrevious prior
    Busy False
    MsgBox message, vbInformation, "Contour Plotter"
    Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    If h <> 0 Then Close #h
    On Error GoTo 0
    Busy False
    SetStatus message
    MsgBox message, vbExclamation, "Contour Plotter"
End Sub

Private Sub cmdCancel_Click()
    If CPBusy Then
        CPCancelled = True
        SetStatus "Cancellation requested..."
    Else
        Unload Me
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CPBusy Then
        CPCancelled = True
        Cancel = True
    End If
End Sub
