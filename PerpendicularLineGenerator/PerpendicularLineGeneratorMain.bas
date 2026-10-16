Attribute VB_Name = "PerpendicularLineGeneratorMain"
Option Explicit

' Key-in: VBA RUN [PerpendicularLineGenerator]PerpendicularLineGeneratorMain.Main
Private mSource As Element
Private mModel As ModelReference
Private mSessionActive As Boolean
Private mStatus As String
Private mPicker As clsPerpendicularLocate

Public Sub Main()
    On Error GoTo Failed
    If mSessionActive Then Exit Sub
    PLGRequire ActiveModelReference.Type <> msdModelTypeSheet, "Open the alignment's design/drawing model before running this tool. Sheet models are excluded."
    CommandState.StartLocate New clsPerpendicularLocate
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Perpendicular Line Generator"
    CommandState.StartDefaultCommand
End Sub

Public Function PLGBeginAlignment(source As Element) As Boolean
    Dim reason As String
    On Error GoTo Failed
    If mSessionActive Then Exit Function
    PLGRequire ActiveModelReference.Type <> msdModelTypeSheet, "Open the alignment's design/drawing model before running this tool. Sheet models are excluded."
    PLGRequire source.ModelReference.IsActive And Not source.ModelReference.IsAttachment, _
        "Select geometry in the active model. Reference attachments are not supported."
    If Not PLGLoadPath(source, reason) Then
        MsgBox reason, vbExclamation, "Unsupported alignment"
        Exit Function
    End If
    Set mSource = source
    Set mModel = ActiveModelReference
    mSessionActive = True
    mStatus = reason
    frmPerpendicularLines.RefreshContext mStatus
    BeginRangePick 1, True
    PLGBeginAlignment = True
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Perpendicular Line Generator"
    PLGCloseSession
End Function

Public Sub PLGCloseSession()
    Set mPicker = Nothing
    mSessionActive = False
    Set mSource = Nothing
    Set mModel = Nothing
    PLGClearPath
End Sub

Public Function PLGPickContextValid() As Boolean
    On Error GoTo Done
    If Not mSessionActive Then Exit Function
    If mModel Is Nothing Then Exit Function
    PLGPickContextValid = mModel.IsActive
Done:
End Function

Private Sub ShowSettingsDialog()
    Dim target As Long
    On Error GoTo Failed
    PLGRequire PLGPickContextValid(), "The active model changed. Run the tool again."
    frmPerpendicularLines.PrepareDialog mStatus
    frmPerpendicularLines.Show vbModal
    target = frmPerpendicularLines.RequestedPick
    If target > 0 Then
        BeginRangePick target, False
    Else
        PLGCloseSession
    End If
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Perpendicular Line Generator"
    PLGCloseSession
End Sub

Private Sub BeginRangePick(ByVal target As Long, ByVal followWithEnd As Boolean)
    Dim reverse As Boolean, previewLength As Double, placement As Long, firstDistance As Double
    Dim errorText As String, stopError As Long
    On Error GoTo Failed
    PLGRequire PLGPickContextValid(), "The active model changed. Run the tool again."
    frmPerpendicularLines.ReadPickSettings target, reverse, previewLength, placement, firstDistance
    Set mPicker = New clsPerpendicularLocate
    mPicker.ConfigurePick target, reverse, previewLength, placement, firstDistance, followWithEnd
    CommandState.StartPrimitive mPicker
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    Err.Clear
    If Not mPicker Is Nothing Then mPicker.StopPick Else CommandState.StartDefaultCommand
    stopError = Err.Number
    Set mPicker = Nothing
    On Error GoTo 0
    If stopError <> 0 Then
        PLGCloseSession
        MsgBox errorText & vbCrLf & "The picker could not be stopped. Start another ORD command before retrying.", vbExclamation, "Perpendicular Line Generator"
        Exit Sub
    End If
    MsgBox errorText, vbExclamation, "Perpendicular Line Generator"
    If PLGPickContextValid() Then ShowSettingsDialog Else PLGCloseSession
End Sub

Public Sub PLGFinishRangePick(ByVal target As Long, ByVal accepted As Boolean, _
        ByVal distance As Double, ByVal followWithEnd As Boolean)
    On Error GoTo Failed
    Set mPicker = Nothing
    PLGRequire PLGPickContextValid(), "The active model changed. Run the tool again."
    If accepted Then frmPerpendicularLines.SetPickedDistance target, distance
    If accepted And target = 1 And followWithEnd Then
        BeginRangePick 2, False
    Else
        ShowSettingsDialog
    End If
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Perpendicular Line Generator"
    PLGCloseSession
End Sub

Public Function PLGUnitsDescription() As String
    Dim master As MeasurementUnit, subUnit As MeasurementUnit
    master = ActiveModelReference.MasterUnit
    subUnit = ActiveModelReference.SubUnit
    PLGUnitsDescription = "Master: " & master.Label & " | Sub: " & subUnit.Label
End Function

Public Function PLGGenerate(ByVal first As Double, ByVal last As Double, ByVal spacing As Double, _
        ByVal totalLength As Double, ByVal placement As Long, ByVal reverse As Boolean, _
        ByVal includeFirst As Boolean, ByVal includeLast As Boolean) As String
    Dim level As Level, style As LineStyle, color As Long, weight As Long
    Dim lines As New Collection, line As LineElement, currentSource As Element
    Dim distances() As Double, n As Long, i As Long, attempted As Long, remaining As Long
    Dim p As Point3d, t As Point3d, a As Point3d, b As Point3d
    Dim reason As String, originalError As String
    On Error GoTo Failed
    PLGRequire Not mSource Is Nothing, "Select an alignment first."
    PLGRequire mModel.IsActive, "The active model changed. Cancel and run the macro again."
    PLGRequire Not ActiveModelReference.IsReadOnly, "The active model/DGN is read-only."
    ' Capture settings on Generate; do not reuse the values shown when the dialog opened.
    Set level = ActiveSettings.Level
    Set style = ActiveSettings.LineStyle
    color = ActiveSettings.Color
    weight = ActiveSettings.LineWeight
    PLGRequire level.ElementAccess = msdLevelElementAccessAll, _
        "The active level must allow modification/deletion so failed writes can be rolled back. Choose an unlocked writable level."
    Set currentSource = mModel.GetElementByID(mSource.ID)
    If Not PLGLoadPath(currentSource, reason) Then Err.Raise vbObjectError + 2605, , reason
    n = PLGBuildDistances(first, last, spacing, includeFirst, includeLast, distances)
    For i = 1 To n
        PLGEvaluate distances(i), reverse, p, t
        PLGEndpoints p, t, totalLength, placement, a, b
        Set line = CreateLineElement2(Nothing, a, b)
        line.Color = color
        line.LineWeight = weight
        Set line.LineStyle = style
        PLGRequire Abs(line.Length - totalLength) <= PLGDistanceTolerance(), "A prepared line has the wrong length."
        lines.Add line
    Next i
    ' All geometry and all LineElements now exist in memory; only this loop writes.
    For i = 1 To lines.Count
        Set line = lines(i)
        attempted = i
        ActiveModelReference.AddElement line
        ' Bentley forbids assigning a Level object before AddElement.
        Set line.Level = level
        line.Rewrite
    Next i
    PLGGenerate = "Generated " & CStr(n) & " perpendicular lines at " & Format(spacing, "0.00") & _
        " master-unit spacing on level " & level.Name & "." & vbCrLf & PLGUnitsDescription()
    If includeLast Then PLGGenerate = PLGGenerate & vbCrLf & "The ending line may have a shorter final spacing."
    Exit Function
Failed:
    originalError = Err.Description
    ' Remove only this run's elements; never issue a broad Undo command.
    On Error Resume Next
    For i = attempted To 1 Step -1
        Err.Clear
        Set line = lines(i)
        If DLongToString(line.ID) <> "0" Then mModel.RemoveElement line
        If Err.Number <> 0 Then remaining = remaining + 1
    Next i
    On Error GoTo 0
    If remaining > 0 Then
        originalError = originalError & vbCrLf & "Rollback could not remove " & CStr(remaining) & _
            " element(s). Use ordinary Undo and inspect the model before retrying."
    ElseIf attempted > 0 Then
        originalError = originalError & vbCrLf & "This run's added elements were removed."
    Else
        originalError = originalError & vbCrLf & "No elements were added."
    End If
    Err.Raise vbObjectError + 2604, "Perpendicular Lines", originalError
End Function
