Attribute VB_Name = "modContourImport"
Option Explicit
' SPDX-License-Identifier: MIT
Private mBounds(0 To 3) As Double, mDimension As Long, mExpectedZ As Double

Public Function CPParseNumber(ByVal text As String) As Double
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "^-?[0-9]+(\.[0-9]+)?([eE][+-]?[0-9]+)?$"
    End If
    CPRequire Len(text) <= 64 And re.Test(text), "Invalid invariant number in engine output: " & Left$(text, 64)
    CPParseNumber = Val(text)
    CPRequire Abs(CPParseNumber) <= 1000000000000#, "Engine number exceeds the coordinate safety range."
End Function

Private Function CPInteger(ByVal text As String, ByVal minimum As Long, ByVal maximum As Long) As Long
    Dim value As Double
    value = CPParseNumber(text)
    CPRequire value = Fix(value) And value >= minimum And value <= maximum, "Invalid integer/count in engine output."
    CPInteger = CLng(value)
End Function

Private Function CPFields(ByVal f As Integer) As Variant
    Dim line As String
    CPRequire Not EOF(f), "Truncated contour output."
    Line Input #f, line
    CPRequire Len(line) > 0 And Len(line) <= 4096, "Invalid contour record length."
    CPFields = Split(line, vbTab)
End Function

Private Function CPPoint(ByVal x As String, ByVal y As String, ByVal z As String) As Point3d
    Dim p As Point3d, tolerance As Double
    p.X = CPParseNumber(x): p.Y = CPParseNumber(y): p.Z = CPParseNumber(z)
    tolerance = 2# / CPUOR + (Abs(p.X) + Abs(p.Y)) * 0.00000000000001
    CPRequire p.X >= mBounds(0) - tolerance And p.X <= mBounds(2) + tolerance, "X outside confirmed source bounds."
    CPRequire p.Y >= mBounds(1) - tolerance And p.Y <= mBounds(3) + tolerance, "Y outside confirmed source bounds."
    CPRequire Abs(p.Z - mExpectedZ) <= 2# / CPUOR + Abs(mExpectedZ) * 0.00000000000001, "Z differs from contour elevation."
    If mDimension = 2 Then CPRequire p.Z = 0#, "Nonzero elevation in 2D output."
    CPPoint = p
End Function

Public Function CPImport(ByVal filename As String, ByVal job As String, ByVal runID As String, _
        ByVal majorLevel As String, ByVal minorLevel As String, ByVal majorLabel As String, ByVal minorLabel As String, _
        ByVal styleName As String, ByVal height As Double, ByVal minorHeight As Double, _
        ByVal boxHeight As Double, ByVal minorBoxHeight As Double, ByVal useByLevel As Boolean, _
        ByVal overrideColor As Long, ByVal overrideWeight As Long, ByVal overrideStyle As String, _
        ByVal sourceContext As String, ByVal expectedDimension As Long) As String
    Dim f As Integer, pass As Long, header As Variant, row As Variant, vertex As Variant
    Dim expected(0 To 2) As Long, counts(0 To 2) As Long, i As Long, n As Long, kind As String
    Dim p As Point3d, points() As Point3d, element As Element, text As TextElement
    Dim levels(0 To 3) As Level, target As Level, added As New Collection, journal As Object, fso As Object
    Dim w As Double, h As Double, angle As Double, reservedWidth As Double, labelHeight As Double
    Dim expectedHeight As Double, expectedBoxHeight As Double
    Dim reason As String, remaining As Long, journalName As String, started As Double, checked As Double
    Dim explicitStyle As LineStyle
    On Error GoTo Failed
    CPPreflight
    If Not useByLevel Then
        CPRequire overrideColor >= 0 And overrideColor <= 255 And overrideWeight >= 0 And overrideWeight <= 31, "Invalid symbology override."
        CPRequire Len(overrideStyle) > 0, "Enter an existing line style override."
        Set explicitStyle = ActiveDesignFile.LineStyles.Item(overrideStyle)
        CPRequire Not explicitStyle Is Nothing, "Line style override unavailable."
    End If
    Set levels(0) = CPLevel(majorLevel): Set levels(1) = CPLevel(minorLevel)
    If Len(majorLabel) > 0 Then Set levels(2) = CPLevel(majorLabel)
    If Len(minorLabel) > 0 Then Set levels(3) = CPLevel(minorLabel)
    f = FreeFile
    ' Hold a deny-write lock through BOTH passes so validation applies to inserted bytes.
    Open filename For Input Access Read Lock Write As #f
    started = Timer
    For pass = 1 To 2
        Seek #f, 1
        header = CPFields(f)
        CPRequire UBound(header) = 13, "Invalid CP2 header."
        CPRequire header(0) = "CP2" And header(1) = runID, "Unexpected contour output version/run."
        CPRequire header(2) = CPUnitName, "Output/model unit mismatch."
        CPRequire header(3) = "m" Or header(3) = "ft" Or header(3) = "us-ft", "Unresolved output vertical unit."
        For i = 0 To 2
            expected(i) = CPInteger(CStr(header(i + 4)), 0, 100000000)
            counts(i) = 0
        Next i
        CPRequire CPParseNumber(CStr(header(7))) = height, "Output text height changed."
        CPRequire CPParseNumber(CStr(header(8))) = minorHeight, "Minor output text height changed."
        For i = 0 To 3
            mBounds(i) = CPParseNumber(CStr(header(i + 9)))
        Next i
        CPRequire mBounds(2) >= mBounds(0) And mBounds(3) >= mBounds(1), "Invalid source bounds."
        mDimension = CPInteger(CStr(header(13)), 2, 3)
        CPRequire mDimension = expectedDimension, "Output dimension differs from the requested setting."
        If mDimension = 3 Then CPRequire CPModel.Is3D, "3D contours require a 3D design model."
        If pass = 2 Then
            CPPreflight
            CPRequire MsgBox("Insert into:" & vbCrLf & CPFileName & vbCrLf & CPModelName & _
                vbCrLf & CStr(expected(0)) & " line strings; " & CStr(expected(1)) & " ordinary text elements." & _
                vbCrLf & "Model units: " & CPUnitName & "; elevation labels: " & CStr(header(3)) & _
                vbCrLf & "Bounds: " & CStr(header(9)) & ", " & CStr(header(10)) & " to " & CStr(header(11)) & ", " & CStr(header(12)) & _
                vbCrLf & "Contour levels: " & majorLevel & " / " & minorLevel & vbCrLf & sourceContext & _
                vbCrLf & CPGcsName & vbCrLf & "Confirm this source grid, axis mapping, vertical datum and location match the active model.", _
                vbYesNo + vbExclamation + vbDefaultButton2, "Confirm native DGN placement") = vbYes, "Placement cancelled before insertion."
            journalName = job & "\elements.cpj"
            CPStartJournal journalName, runID, journal
        End If
        Do
            row = CPFields(f)
            kind = CStr(row(0))
            If kind = "E" Then
                CPRequire UBound(row) = 3, "Invalid CP2 footer."
                For i = 0 To 2
                    CPRequire CPInteger(CStr(row(i + 1)), 0, 100000000) = counts(i), "Footer count mismatch."
                    CPRequire counts(i) = expected(i), "Header count mismatch."
                Next i
                CPRequire EOF(f), "Unexpected data after CP2 footer."
                Exit Do
            End If
            CPRequire kind = "C" Or kind = "L", "Unknown contour record."
            If kind = "C" Then
                CPRequire UBound(row) = 4, "Invalid contour record."
            Else
                CPRequire UBound(row) = 10, "Invalid label record."
            End If
            Call CPInteger(CStr(row(1)), 1, 100000000)
            mExpectedZ = 0#
            If mDimension = 3 Then mExpectedZ = CPParseNumber(CStr(row(2))) * CPMetres(CStr(header(3))) / CPMetres(CPUnitName)
            Call CPParseNumber(CStr(row(2)))
            CPRequire row(3) = "M" Or row(3) = "m", "Invalid contour classification."
            If kind = "C" Then
                n = CPInteger(CStr(row(4)), 2, 1024)
                ReDim points(0 To n - 1)
                For i = 0 To n - 1
                    vertex = CPFields(f)
                    CPRequire UBound(vertex) = 3, "Invalid vertex record."
                    CPRequire vertex(0) = "V", "Expected a vertex."
                    points(i) = CPPoint(CStr(vertex(1)), CStr(vertex(2)), CStr(vertex(3)))
                    If i > 0 Then
                        CPRequire points(i).X <> points(i - 1).X Or points(i).Y <> points(i - 1).Y, "Duplicate consecutive contour vertex."
                    End If
                Next i
                Set element = CreateLineElement1(Nothing, points)
                If row(3) = "M" Then Set target = levels(0) Else Set target = levels(1)
                counts(0) = counts(0) + 1: counts(2) = counts(2) + n
            Else
                p = CPPoint(CStr(row(4)), CStr(row(5)), CStr(row(6)))
                angle = CPParseNumber(CStr(row(7)))
                labelHeight = CPParseNumber(CStr(row(8)))
                reservedWidth = CPParseNumber(CStr(row(9)))
                If row(3) = "M" Then
                    expectedHeight = height: expectedBoxHeight = boxHeight
                Else
                    expectedHeight = minorHeight: expectedBoxHeight = minorBoxHeight
                End If
                CPRequire Abs(angle) <= 1.570796326795 And labelHeight = expectedHeight And reservedWidth > 0, "Invalid text angle/footprint."
                CPRequire CPParseNumber(CStr(row(10))) = CPParseNumber(CStr(row(2))), "Label value differs from elevation."
                Set text = CPText(CStr(row(10)), p, angle, labelHeight, styleName)
                text.GetTotalTextSize w, h
                CPRequire w <= reservedWidth + 2# / CPUOR And h <= expectedBoxHeight + 2# / CPUOR, "Font footprint changed; analyze and measure again."
                Set element = text
                If row(3) = "M" Then Set target = levels(2) Else Set target = levels(3)
                CPRequire Not target Is Nothing, "Required label level missing."
                counts(1) = counts(1) + 1
            End If
            If pass = 2 Then
                CPTag element, runID
                added.Add element
                CPModel.AddElement element
                CPApply element, target, useByLevel, overrideColor, overrideWeight, overrideStyle
                CPRequire CPTagValue(element) = "CP1|" & runID, "Persistent run tag did not survive Rewrite."
                journal.WriteLine "ID" & vbTab & DLongToString(element.ID) & vbTab & CPNumber(CDbl(element.DateLastModified))
                If added.Count Mod 128 = 0 Then
                    journal.Close
                    Set journal = CreateObject("Scripting.FileSystemObject").OpenTextFile(journalName, 8, False, -1)
                End If
            End If
            If (counts(0) + counts(1)) Mod 128 = 0 Then
                CPStatus "Pass " & CStr(pass) & "/2: " & CStr(counts(0) + counts(1)) & " elements"
                CPPreflight
            End If
        Loop
        If pass = 1 Then checked = Timer
    Next pass
    CPPreflight
    Close #f: f = 0
    journal.WriteLine "COMMITTED" & vbTab & CStr(added.Count)
    journal.Close
    Set journal = Nothing
    CPImport = "Inserted " & CStr(counts(0)) & " line strings and " & CStr(counts(1)) & " text elements." & vbCrLf & _
        "Validation: " & Format$(CPElapsed(started, checked), "0.00") & " s; insertion: " & _
        Format$(CPElapsed(checked, Timer), "0.00") & " s." & vbCrLf & "Run folder: " & job
    Exit Function
Failed:
    reason = Err.Description
    On Error Resume Next
    If f <> 0 Then Close #f
    If Not journal Is Nothing Then journal.Close
    remaining = CPRollback(added, runID)
    On Error GoTo 0
    Err.Raise vbObjectError + 2815, "Contour Plotter", reason & vbCrLf & _
        "New elements still requiring inspection: " & CStr(remaining) & ". Previous successful output was retained." & vbCrLf & _
        "Recovery folder: " & job
End Function

Private Function CPElapsed(ByVal first As Double, ByVal last As Double) As Double
    CPElapsed = last - first
    If CPElapsed < 0# Then CPElapsed = CPElapsed + 86400#
End Function

Private Function CPMetres(ByVal unit As String) As Double
    Select Case unit
        Case "m": CPMetres = 1#
        Case "ft": CPMetres = 0.3048
        Case "us-ft": CPMetres = 1200# / 3937#
        Case Else: CPRequire False, "Unsupported CP2 linear unit."
    End Select
End Function
