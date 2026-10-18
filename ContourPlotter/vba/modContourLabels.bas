Attribute VB_Name = "modContourLabels"
Option Explicit
' SPDX-License-Identifier: MIT
Public Function CPText(ByVal value As String, origin As Point3d, ByVal angle As Double, _
        ByVal height As Double, ByVal styleName As String) As TextElement
    Dim text As TextElement, style As TextStyle, rotation As Matrix3d, ratio As Double
    Dim actualWidth As Double, actualHeight As Double
    Set style = CPStyle(styleName)
    ratio = style.Width / style.Height
    rotation = Matrix3dFromAxisAndRotationAngle(2, angle)
    Set text = CreateTextElement1(Nothing, value, origin, rotation)
    Set text.TextStyle = style
    ' Change the element's local style; never change the object fetched from the DGN table.
    text.TextStyle.Height = height
    text.TextStyle.Width = height * ratio
    text.TextStyle.Justification = msdTextJustificationCenterCenter
    text.TextStyle.UsesColor = False
    text.TextStyle.BorderAndBackgroundVisible = False
    Set style = text.TextStyle
    Set text.TextStyle = style
    text.GetTotalTextSize actualWidth, actualHeight
    CPRequire actualWidth > 0# And actualHeight > 0#, "Text style produced an invalid physical footprint."
    ' Normalize model-space geometry without changing active annotation scale.
    text.ScaleUniform origin, height / actualHeight
    text.Origin = origin
    text.GetTotalTextSize actualWidth, actualHeight
    CPRequire Abs(actualHeight - height) <= 2# / CPUOR, "Text physical height could not be calibrated."
    Set CPText = text
End Function

Public Function CPMeasuredWidths(ByVal job As String, ByVal styleName As String, _
        ByVal height As Double, ByVal minorHeight As Double, ByRef boxHeight As Double, ByRef minorBoxHeight As Double) As String
    Dim f As Integer, line As String, text As TextElement, p As Point3d
    Dim w As Double, h As Double, result As String, message As String, fields As Variant, textHeight As Double
    On Error GoTo Failed
    boxHeight = height: minorBoxHeight = minorHeight
    result = "{"
    f = FreeFile
    Open job & "\levels.tsv" For Input Access Read Lock Write As #f
    Do While Not EOF(f)
        Line Input #f, line
        fields = Split(line, vbTab)
        CPRequire UBound(fields) = 1, "Invalid classified elevation from engine."
        line = CStr(fields(0))
        CPRequire Len(line) > 0 And Len(line) <= 64, "Invalid elevation text from engine."
        CPRequire fields(1) = "M" Or fields(1) = "m", "Invalid elevation class."
        Call CPParseNumber(line)
        If fields(1) = "M" Then textHeight = height Else textHeight = minorHeight
        Set text = CPText(line, p, 0#, textHeight, styleName)
        text.GetTotalTextSize w, h
        CPRequire w > 0 And h > 0, "The DGN font could not be measured."
        If Len(result) > 1 Then result = result & ","
        result = result & CPJson(line) & ":" & CPNumber(w)
        If fields(1) = "M" Then
            If h > boxHeight Then boxHeight = h
        Else
            If h > minorBoxHeight Then minorBoxHeight = h
        End If
    Loop
    Close #f
    CPMeasuredWidths = result & "}"
    Exit Function
Failed:
    message = Err.Description
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
    Err.Raise vbObjectError + 2811, "Contour Plotter", message
End Function
