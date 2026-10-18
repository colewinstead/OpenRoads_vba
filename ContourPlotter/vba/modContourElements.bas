Attribute VB_Name = "modContourElements"
Option Explicit
' SPDX-License-Identifier: MIT
Public Const CP_XDATA As String = "OpenRoadsContourPlotter_v1"

Public Sub CPTag(element As Element, ByVal runID As String)
    Dim data(0 To 0) As XDatum
    data(0).Type = msdXDatumTypeString
    data(0).Value = "CP1|" & runID
    element.SetXData CP_XDATA, data
End Sub

Public Function CPTagValue(element As Element) As String
    Dim data() As XDatum
    On Error GoTo Missing
    If Not element.HasXData(CP_XDATA) Then Exit Function
    data = element.GetXData(CP_XDATA)
    If UBound(data) >= 0 Then
        If data(0).Type = msdXDatumTypeString Then CPTagValue = CStr(data(0).Value)
    End If
Missing:
End Function

Public Sub CPApply(element As Element, level As Level, ByVal useByLevel As Boolean, _
        ByVal color As Long, ByVal weight As Long, ByVal styleName As String)
    Dim style As LineStyle
    ' Bentley's Level object assignment follows AddElement, as in the existing tool.
    Set element.Level = level
    If useByLevel Then
        element.Color = ByLevelColor
        element.LineWeight = ByLevelLineWeight
        Set element.LineStyle = ByLevelLineStyle
    Else
        CPRequire color >= 0 And color <= 255 And weight >= 0 And weight <= 31, "Invalid element color/weight override."
        Set style = ActiveDesignFile.LineStyles.Item(styleName)
        CPRequire Not style Is Nothing, "Line style override unavailable."
        element.Color = color
        element.LineWeight = weight
        Set element.LineStyle = style
    End If
    element.Class = msdElementClassPrimary
    element.GraphicGroup = 0
    element.Rewrite
    CPRequire element.Level.Name = level.Name, "Level assignment did not survive Rewrite."
    If useByLevel Then
        CPRequire element.Color = ByLevelColor And element.LineWeight = ByLevelLineWeight, "ByLevel color/weight did not survive Rewrite."
        CPRequire element.LineStyle.Name = ByLevelLineStyle.Name, "ByLevel line style did not survive Rewrite."
    End If
End Sub

Public Sub CPStartJournal(ByVal filename As String, ByVal runID As String, ByRef journal As Object)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    CPRequire Not fso.FileExists(filename), "Insertion journal already exists."
    CPRequire InStr(CPFileName & CPModelName, vbTab) = 0, "DGN/model names containing tabs are unsupported."
    Set journal = fso.CreateTextFile(filename, False, True)
    journal.WriteLine "CPJ1" & vbTab & runID
    journal.WriteLine CPFileName
    journal.WriteLine CPModelName
    journal.WriteLine CPNumber(CDbl(ActiveDesignFile.DateCreated))
    journal.Close
    Set journal = fso.OpenTextFile(filename, 8, False, -1)
End Sub

Public Function CPRollback(ByVal added As Collection, ByVal runID As String) As Long
    Dim i As Long, element As Element
    On Error Resume Next
    For i = added.Count To 1 Step -1
        Err.Clear
        Set element = added(i)
        If DLongToString(element.ID) <> "0" Then
            If CPTagValue(element) = "CP1|" & runID Then
                CPModel.RemoveElement element
            Else
                Err.Raise vbObjectError + 2813, , "New element lost its run tag."
            End If
        End If
        If Err.Number <> 0 Then CPRollback = CPRollback + 1
    Next i
    On Error GoTo 0
End Function

Public Sub CPRemovePrevious(ByVal filename As String)
    Dim stream As Object, fso As Object, header As Variant, row As Variant, line As String
    Dim ids As New Collection, stamps As New Collection, runID As String, element As Element
    Dim i As Long, changed As Long, missing As Long, removed As Long, lastLine As String, message As String
    On Error GoTo Failed
    CPPreflight
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.OpenTextFile(filename, 1, False, -1)
    header = Split(stream.ReadLine, vbTab)
    CPRequire UBound(header) = 1, "Invalid prior journal."
    CPRequire header(0) = "CPJ1", "Invalid prior journal version."
    runID = header(1)
    CPRequire stream.ReadLine = CPFileName, "Prior journal belongs to another DGN path."
    CPRequire stream.ReadLine = CPModelName, "Prior journal belongs to another model."
    CPRequire stream.ReadLine = CPNumber(CDbl(ActiveDesignFile.DateCreated)), "Prior journal belongs to a different DGN file."
    Do While Not stream.AtEndOfStream
        line = stream.ReadLine
        row = Split(line, vbTab)
        If row(0) = "ID" Then
            CPRequire UBound(row) = 2, "Invalid journal ID record."
            ids.Add CStr(row(1))
            stamps.Add CStr(row(2))
        Else
            lastLine = line
            CPRequire stream.AtEndOfStream, "Trailing or unknown journal records."
        End If
    Loop
    stream.Close
    CPRequire lastLine = "COMMITTED" & vbTab & CStr(ids.Count), "Prior insertion was not committed; follow recovery instructions."
    ' Resolve and check every recorded ID/tag before the first removal.
    For i = 1 To ids.Count
        Set element = CPFindElement(CStr(ids(i)))
        If element Is Nothing Then
            missing = missing + 1
        Else
            CPRequire CPTagValue(element) = "CP1|" & runID, "Prior ID/tag mismatch; no previous elements were deleted."
            CPRequire Not element.IsLocked And element.Level.ElementAccess = msdLevelElementAccessAll, "Prior element is locked; replacement retained."
            If CPNumber(CDbl(element.DateLastModified)) <> CStr(stamps(i)) Then changed = changed + 1
        End If
    Next i
    message = "The replacement is complete. Remove " & CStr(ids.Count - missing) & " recorded elements from prior run " & runID & "?" & vbCrLf & _
        CStr(changed) & " element timestamps changed; " & CStr(missing) & " elements are already absent." & vbCrLf & _
        "This includes any manually edited members of that run. Copies with new element IDs are preserved." & vbCrLf & _
        "Choose No to keep both runs for review."
    If MsgBox(message, vbYesNo + vbExclamation + vbDefaultButton2, "Confirm staged replacement") <> vbYes Then Exit Sub
    CPPreflight
    For i = 1 To ids.Count
        Set element = CPFindElement(CStr(ids(i)))
        If Not element Is Nothing Then
            CPRequire CPTagValue(element) = "CP1|" & runID, "Prior ID/tag changed; removal stopped."
            CPModel.RemoveElement element
            removed = removed + 1
        End If
        If i Mod 128 = 0 Then
            CPStatus "Removing recorded prior output: " & CStr(i) & "/" & CStr(ids.Count)
            CPPreflight
        End If
    Next i
    MsgBox "Removed " & CStr(removed) & " recorded prior elements. Replacement retained.", vbInformation, "Contour Plotter"
    Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise vbObjectError + 2814, "Contour Plotter", message & vbCrLf & _
        "Replacement retained. Prior elements removed: " & CStr(removed) & ". Keep both run folders for recovery."
End Sub

Public Function CPFindElement(ByVal value As String) As Element
    Dim id As DLong
    id = DLongFromString(value)
    ' GetElementByID returns Nothing for an absent ID; unexpected COM errors propagate.
    Set CPFindElement = CPModel.GetElementByID(id)
End Function

Public Sub ListTaggedRuns()
    Dim ee As ElementEnumerator, element As Element, counts As Object, tag As String, key As Variant, message As String
    On Error GoTo Failed
    Set counts = CreateObject("Scripting.Dictionary")
    Set ee = ActiveModelReference.Scan
    Do While ee.MoveNext
        Set element = ee.Current
        tag = CPTagValue(element)
        If Len(tag) > 0 Then
            If Not counts.Exists(tag) Then counts.Add tag, 0
            counts(tag) = counts(tag) + 1
        End If
    Loop
    For Each key In counts.Keys
        message = message & CStr(key) & ": " & CStr(counts(key)) & vbCrLf
    Next key
    If Len(message) = 0 Then message = "No Contour Plotter tags found in the active model."
    MsgBox message, vbInformation, "Contour Plotter run inventory (read only)"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Contour Plotter"
End Sub
