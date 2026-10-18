Attribute VB_Name = "ContourPlotterMain"
Option Explicit
' SPDX-License-Identifier: MIT
Public CPModel As ModelReference
Public CPFileName As String, CPModelName As String, CPUnitName As String, CPGcsName As String
Public CPUOR As Double, CPCancelled As Boolean, CPBusy As Boolean
Public Const CP_APP As String = "OpenRoadsContourPlotter"

Public Sub Main()
    On Error GoTo Failed
    CPRequire Not CPBusy, "Contour Plotter is already processing."
    CPCancelled = False
    Set CPModel = ActiveModelReference
    CPFileName = ActiveDesignFile.FullName
    CPModelName = CPModel.Name
    CPUnitName = CPWorkingUnit()
    CPGcsName = CPDetectedGCS()
    CPUOR = CPModel.UORsPerMasterUnit
    CPPreflight
    Load frmContourPlotter
    frmContourPlotter.Show vbModal
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Contour Plotter"
End Sub

Public Sub CPRequire(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2810, "Contour Plotter", message
End Sub

Public Function CPWorkingUnit() As String
    Dim unit As MeasurementUnit, metres As Double
    unit = ActiveModelReference.MasterUnit
    CPRequire unit.Base = msdMeasurementBaseMeter, "Model master units must be linear."
    CPRequire unit.UnitsPerBaseNumerator > 0 And unit.UnitsPerBaseDenominator > 0, "Invalid model unit ratio."
    metres = unit.UnitsPerBaseDenominator / unit.UnitsPerBaseNumerator
    If Abs(metres - 1#) < 0.000000000001 Then CPWorkingUnit = "m"
    If Abs(metres - 0.3048) < 0.000000000001 Then CPWorkingUnit = "ft"
    If Abs(metres - 1200# / 3937#) < 0.000000000001 Then CPWorkingUnit = "us-ft"
    CPRequire Len(CPWorkingUnit) > 0, "Supported master units: metre, international foot, US survey foot."
End Function

Public Function CPDetectedGCS() As String
    Dim gcs As GeographicCoordinateSystem
    On Error GoTo Unknown
    Set gcs = ActiveModelReference.GetGCS(True)
    If gcs Is Nothing Then GoTo Unknown
    CPDetectedGCS = gcs.Name & " | " & gcs.DatumName & " | " & gcs.VerticalDatumName & " | " & gcs.Units
    Exit Function
Unknown:
    CPDetectedGCS = "(No readable DGN GCS; explicitly confirm the local grid.)"
End Function

Public Sub CPPreflight()
    CPRequire Not CPModel Is Nothing, "Run ContourPlotterMain.Main first."
    CPRequire CPModel.IsActive, "The active model changed; restart Contour Plotter."
    CPRequire ActiveDesignFile.FullName = CPFileName And ActiveModelReference.Name = CPModelName, "DGN/model changed."
    CPRequire Not CPModel.IsAttachment, "Reference models are not output targets."
    CPRequire CPModel.Type = msdModelTypeNormal, "Use a design model; sheet/drawing models are rejected."
    CPRequire Not CPModel.IsReadOnly And Not CPModel.IsLocked, "The active model is read-only or locked."
    CPRequire CPWorkingUnit() = CPUnitName And CPModel.UORsPerMasterUnit = CPUOR, "Model working units/resolution changed."
    CPRequire CPDetectedGCS() = CPGcsName, "Model GCS changed while processing."
    CPRequire Not CPCancelled, "Cancelled."
End Sub

Public Function CPNumber(ByVal number As Double) As String
    CPNumber = Trim$(Str$(number))
    If Left$(CPNumber, 1) = "." Then CPNumber = "0" & CPNumber
    If Left$(CPNumber, 2) = "-." Then CPNumber = "-0" & Mid$(CPNumber, 2)
End Function

Public Function CPJson(ByVal value As String) As String
    Dim i As Long, c As Long, result As String
    result = Chr$(34)
    For i = 1 To Len(value)
        c = AscW(Mid$(value, i, 1)) And &HFFFF&
        Select Case c
            Case 34: result = result & "\" & Chr$(34)
            Case 92: result = result & "\\"
            Case 32 To 126: result = result & ChrW$(c)
            Case Else: result = result & "\u" & Right$("0000" & Hex$(c), 4)
        End Select
    Next i
    CPJson = result & Chr$(34)
End Function

Public Function CPReadUnicode(ByVal filename As String) As String
    Dim stream As Object
    Set stream = CreateObject("Scripting.FileSystemObject").OpenTextFile(filename, 1, False, -1)
    CPReadUnicode = stream.ReadAll
    stream.Close
End Function

Public Sub CPStatus(ByVal message As String)
    frmContourPlotter.SetStatus message
    DoEvents
End Sub
