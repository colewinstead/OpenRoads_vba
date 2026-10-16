Option Explicit

' Paste into a NEW blank UserForm named frmPerpendicularLines.
' Controls are created here; no FRX blob or VBA-project automation is needed.
Private WithEvents txtSpacing As MSForms.TextBox
Private WithEvents txtLength As MSForms.TextBox
Private WithEvents txtStart As MSForms.TextBox
Private WithEvents txtEnd As MSForms.TextBox
Private WithEvents cboPlacement As MSForms.ComboBox
Private WithEvents cboDirection As MSForms.ComboBox
Private WithEvents chkStart As MSForms.CheckBox
Private WithEvents chkEnd As MSForms.CheckBox
Private WithEvents cmdGenerate As MSForms.CommandButton
Private WithEvents cmdCancel As MSForms.CommandButton
Private WithEvents cmdPickStart As MSForms.CommandButton
Private WithEvents cmdPickEnd As MSForms.CommandButton
Private WithEvents cmdFullAlignment As MSForms.CommandButton
Private lblLevel As MSForms.Label
Private lblUnits As MSForms.Label
Private lblStatus As MSForms.Label
Private lblEstimate As MSForms.Label
Private mInitializing As Boolean
Private mRequestedPick As Long

Private Function AddLabel(ByVal controlName As String, ByVal caption As String, _
        ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single) As MSForms.Label
    Dim label As MSForms.Label
    Set label = Me.Controls.Add("Forms.Label.1", controlName, True)
    label.caption = Replace(caption, "&", "")
    If InStr(caption, "&") > 0 Then label.Accelerator = Mid$(caption, InStr(caption, "&") + 1, 1)
    label.Left = x: label.Top = y: label.width = width: label.height = height
    Set AddLabel = label
End Function

Private Function AddText(ByVal controlName As String, ByVal caption As String, _
        ByVal y As Single, ByVal value As String) As MSForms.TextBox
    Dim label As MSForms.Label, box As MSForms.TextBox
    Set label = AddLabel("lbl" & controlName, caption, 18, y + 3, 220, 18)
    Set box = Me.Controls.Add("Forms.TextBox.1", controlName, True)
    box.Left = 250: box.Top = y: box.width = 230: box.height = 22
    box.value = value
    Set AddText = box
End Function

Private Sub UserForm_Initialize()
    Dim label As MSForms.Label
    mInitializing = True
    Me.caption = "Perpendicular Line Generator"
    Me.width = 512: Me.height = 540
    Me.StartUpPosition = 1
    Me.Font.Name = "Segoe UI": Me.Font.Size = 10
    Set label = AddLabel("lblHeading", "PERPENDICULAR LINE GENERATOR", 18, 14, 462, 22)
    label.Font.Bold = True: label.Font.Size = 12
    Set label = AddLabel("lblDistances", "All distances use model MASTER units, measured along geometry from the chosen direction's beginning. These are NOT displayed ORD stations.", 18, 42, 462, 38)
    Set lblStatus = AddLabel("lblStatus", "Alignment: not selected", 18, 84, 462, 32)
    Set lblLevel = AddLabel("lblLevel", "Active level:", 18, 118, 462, 18)
    Set lblUnits = AddLabel("lblUnits", "Working units:", 18, 138, 462, 18)
    Set txtSpacing = AddText("txtSpacing", "Line &Spacing", 168, Format$(150#, "0.00"))
    Set txtLength = AddText("txtLength", "Total Line &Length", 198, Format$(200#, "0.00"))
    Set label = AddLabel("lblPlacement", "&Placement", 18, 231, 220, 18)
    Set cboPlacement = Me.Controls.Add("Forms.ComboBox.1", "cboPlacement", True)
    With cboPlacement
        .Left = 250: .Top = 228: .width = 230: .height = 22
        .Style = fmStyleDropDownList
        .AddItem "Centered on Alignment": .AddItem "Left of Alignment": .AddItem "Right of Alignment"
        .ListIndex = 0
    End With
    Set txtStart = AddText("txtStart", "&Start Distance Along Alignment", 258, Format$(0#, "0.00"))
    Set txtEnd = AddText("txtEnd", "&End Distance", 288, "End of Alignment")
    txtStart.width = 150: txtEnd.width = 150
    Set cmdPickStart = Me.Controls.Add("Forms.CommandButton.1", "cmdPickStart", True)
    With cmdPickStart
        .caption = "Pick Start": .Left = 410: .Top = 258: .width = 70: .height = 22
    End With
    Set cmdPickEnd = Me.Controls.Add("Forms.CommandButton.1", "cmdPickEnd", True)
    With cmdPickEnd
        .caption = "Pick End": .Left = 410: .Top = 288: .width = 70: .height = 22
    End With
    Set label = AddLabel("lblDirection", "&Direction", 18, 321, 220, 18)
    Set cboDirection = Me.Controls.Add("Forms.ComboBox.1", "cboDirection", True)
    With cboDirection
        .Left = 250: .Top = 318: .width = 230: .height = 22
        .Style = fmStyleDropDownList
        .AddItem "Forward": .AddItem "Reverse": .ListIndex = 0
    End With
    Set chkStart = Me.Controls.Add("Forms.CheckBox.1", "chkStart", True)
    With chkStart
        .caption = "Include Starting Line": .Left = 18: .Top = 356: .width = 220: .height = 22: .value = True
        .Accelerator = "I"
    End With
    Set chkEnd = Me.Controls.Add("Forms.CheckBox.1", "chkEnd", True)
    With chkEnd
        .caption = "Include Ending Line": .Left = 250: .Top = 356: .width = 230: .height = 22: .value = False
        .Accelerator = "N"
    End With
    Set lblEstimate = AddLabel("lblEstimate", "Estimated lines:", 18, 390, 462, 34)
    Set label = AddLabel("lblLimit", "VBA geometry only. Native civil spirals/station equations are not interpreted. Left/right follow the chosen direction. Limit: 10,000 lines.", 18, 428, 462, 36)
    Set cmdGenerate = Me.Controls.Add("Forms.CommandButton.1", "cmdGenerate", True)
    With cmdGenerate
        .caption = "Generate Lines": .Accelerator = "G"
        .Left = 250: .Top = 476: .width = 140: .height = 28: .Default = True
    End With
    Set cmdCancel = Me.Controls.Add("Forms.CommandButton.1", "cmdCancel", True)
    With cmdCancel
        .caption = "Cancel": .Left = 400: .Top = 476: .width = 80: .height = 28: .Cancel = True
    End With
    Set cmdFullAlignment = Me.Controls.Add("Forms.CommandButton.1", "cmdFullAlignment", True)
    With cmdFullAlignment
        .caption = "Full Alignment": .Accelerator = "F"
        .Left = 18: .Top = 476: .width = 130: .height = 28
    End With
    mInitializing = False
End Sub

Public Sub PrepareDialog(ByVal status As String)
    mRequestedPick = 0
    RefreshContext status
End Sub

Public Property Get RequestedPick() As Long
    RequestedPick = mRequestedPick
End Property

Public Sub ReadPickSettings(ByVal target As Long, ByRef reverse As Boolean, ByRef previewLength As Double, _
        ByRef placement As Long, ByRef firstDistance As Double)
    previewLength = ReadNumber(txtLength, "Total Line Length")
    PLGRequire previewLength > 4# * PLGDistanceTolerance(), "Total Line Length must be positive and exceed four distance tolerances."
    placement = cboPlacement.ListIndex
    PLGRequire placement >= 0 And cboDirection.ListIndex >= 0, "Choose placement and direction."
    reverse = (cboDirection.ListIndex = 1)
    If target = 2 Then
        firstDistance = ReadNumber(txtStart, "Start Distance")
        PLGRequire firstDistance >= 0# And firstDistance <= PLGPathLength(), "Start Distance is outside the geometry. Pick Start again."
    End If
End Sub

Public Sub SetPickedDistance(ByVal target As Long, ByVal distance As Double)
    If target = 1 Then
        txtStart.value = CStr(distance)
    ElseIf Abs(distance - PLGPathLength()) <= PLGDistanceTolerance() Then
        txtEnd.value = "End of Alignment"
    Else
        txtEnd.value = CStr(distance)
    End If
    RefreshEstimate
End Sub

Private Sub cmdPickStart_Click()
    mRequestedPick = 1
    Me.Hide
End Sub

Private Sub cmdPickEnd_Click()
    mRequestedPick = 2
    Me.Hide
End Sub

Private Sub cmdFullAlignment_Click()
    mInitializing = True
    txtStart.value = Format$(0#, "0.00")
    txtEnd.value = "End of Alignment"
    mInitializing = False
    RefreshEstimate
End Sub

Public Sub RefreshContext(ByVal status As String)
    lblStatus.caption = "Alignment: " & status
    lblLevel.caption = "Active level: " & ActiveSettings.Level.Name & " (captured again on Generate)"
    lblUnits.caption = "Working units - " & PLGUnitsDescription()
    RefreshEstimate
End Sub

Private Function ReadNumber(box As MSForms.TextBox, ByVal field As String) As Double
    PLGRequire Len(Trim$(box.Text)) > 0, field & " is required."
    PLGRequire IsNumeric(box.Text), field & " must be a number in master units (no station notation or unit suffix)."
    ReadNumber = CDbl(box.Text)
    PLGRequire Abs(ReadNumber) <= 1000000000#, field & " must not exceed 1 billion master units."
End Function

Private Function ReadSettings(first As Double, last As Double, spacing As Double, totalLength As Double, _
        distances() As Double) As Long
    spacing = ReadNumber(txtSpacing, "Line Spacing")
    totalLength = ReadNumber(txtLength, "Total Line Length")
    PLGRequire totalLength > 4# * PLGDistanceTolerance(), "Total Line Length must be positive and exceed four distance tolerances."
    first = ReadNumber(txtStart, "Start Distance")
    If Len(Trim$(txtEnd.Text)) = 0 Or LCase$(Trim$(txtEnd.Text)) = "end of alignment" Then
        last = PLGPathLength()
    Else
        last = ReadNumber(txtEnd, "End Distance")
    End If
    PLGRequire cboPlacement.ListIndex >= 0 And cboDirection.ListIndex >= 0, "Choose placement and direction."
    ReadSettings = PLGBuildDistances(first, last, spacing, CBool(chkStart.value), CBool(chkEnd.value), distances)
End Function

Private Sub RefreshEstimate()
    Dim first As Double, last As Double, spacing As Double, totalLength As Double, distances() As Double, n As Long
    If mInitializing Then Exit Sub
    On Error GoTo Invalid
    n = ReadSettings(first, last, spacing, totalLength, distances)
    lblEstimate.caption = "Estimated lines: " & CStr(n) & " | An optional ending line can shorten the final interval."
    cmdGenerate.Enabled = True
    Exit Sub
Invalid:
    lblEstimate.caption = Err.Description
    cmdGenerate.Enabled = False
End Sub

Private Sub cmdGenerate_Click()
    Dim first As Double, last As Double, spacing As Double, totalLength As Double, distances() As Double
    Dim n As Long, result As String
    On Error GoTo Failed
    n = ReadSettings(first, last, spacing, totalLength, distances)
    cmdGenerate.Enabled = False
    result = PLGGenerate(first, last, spacing, totalLength, cboPlacement.ListIndex, _
        cboDirection.ListIndex = 1, CBool(chkStart.value), CBool(chkEnd.value))
    MsgBox result, vbInformation, "Perpendicular Line Generator"
    cmdGenerate.Enabled = True
    Me.Hide
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Perpendicular Line Generator"
    RefreshEstimate
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub txtSpacing_Change()
    RefreshEstimate
End Sub
Private Sub txtLength_Change()
    RefreshEstimate
End Sub
Private Sub txtStart_Change()
    RefreshEstimate
End Sub
Private Sub txtEnd_Change()
    RefreshEstimate
End Sub
Private Sub cboPlacement_Change()
    RefreshEstimate
End Sub
Private Sub cboDirection_Change()
    RefreshEstimate
End Sub
Private Sub chkStart_Click()
    RefreshEstimate
End Sub
Private Sub chkEnd_Click()
    RefreshEstimate
End Sub
