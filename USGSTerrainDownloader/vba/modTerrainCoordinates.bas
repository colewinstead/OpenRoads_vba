Attribute VB_Name = "modTerrainCoordinates"
Option Explicit

Public Function TerrainNumber(ByVal value As Double) As String
    ' Str$ uses a period even on decimal-comma Windows installations.
    TerrainNumber = Trim$(Str$(value))
    If Left$(TerrainNumber, 1) = "." Then TerrainNumber = "0" & TerrainNumber
    If Left$(TerrainNumber, 2) = "-." Then TerrainNumber = "-0" & Mid$(TerrainNumber, 2)
End Function

Public Function TerrainVerticesJSON(vertices() As Point3d) As String
    Dim i As Long, result As String, tolerance As Double
    TerrainRequire UBound(vertices) - LBound(vertices) >= 2, "Boundary needs three corners."
    TerrainRequire UBound(vertices) - LBound(vertices) < 20000, "Too many boundary vertices."
    tolerance = 2# / ActiveModelReference.UORsPerMasterUnit
    result = "["
    For i = LBound(vertices) To UBound(vertices)
        TerrainRequire Abs(vertices(i).Z - vertices(LBound(vertices)).Z) <= tolerance, "Select a horizontal shape. Tilted/nonplanar boundaries are unsupported."
        If i > LBound(vertices) Then result = result & ","
        result = result & "[" & TerrainNumber(vertices(i).X) & "," & TerrainNumber(vertices(i).Y) & "]"
    Next i
    TerrainVerticesJSON = result & "]"
End Function

Public Function TerrainModelDescription() As String
    Dim master As MeasurementUnit
    master = ActiveModelReference.MasterUnit
    TerrainModelDescription = "Model: " & ActiveModelReference.Name & " | master label: " & master.Label & _
        " | UOR/master: " & TerrainNumber(ActiveModelReference.UORsPerMasterUnit)
End Function

Public Function TerrainMasterUnitName(unit As MeasurementUnit) As String
    Dim metres As Double
    If unit.Base <> msdMeasurementBaseMeter Then Exit Function
    If unit.UnitsPerBaseNumerator <= 0 Or unit.UnitsPerBaseDenominator <= 0 Then Exit Function
    metres = unit.UnitsPerBaseDenominator / unit.UnitsPerBaseNumerator
    ' Labels such as a prime do not distinguish international and US survey feet.
    If Abs(metres - 1#) < 0.000000000001 Then TerrainMasterUnitName = "m"
    If Abs(metres - 0.3048) < 0.000000000001 Then TerrainMasterUnitName = "ft"
    If Abs(metres - 1200# / 3937#) < 0.000000000001 Then TerrainMasterUnitName = "us-ft"
End Function

Public Function TerrainDetectedGCS() As String
    Dim model As Object, gcs As Object
    On Error GoTo Unavailable
    Set model = ActiveModelReference
    ' shortcut: probe GCS name only; add a verified full-definition adapter for automatic EPSG/modifier support.
    ' Bentley COM GetGCS(True) reads the model GCS, not a guessed EPSG mapping.
    Set gcs = CallByName(model, "GetGCS", VbMethod, True)
    If gcs Is Nothing Then GoTo Unavailable
    TerrainDetectedGCS = CStr(CallByName(gcs, "Name", VbGet))
    If Len(TerrainDetectedGCS) = 0 Then GoTo Unavailable
    TerrainDetectedGCS = "Bentley GCS name: " & TerrainDetectedGCS & ". Confirm equivalent EPSG/WKT and model units below."
    Exit Function
Unavailable:
    TerrainDetectedGCS = "DGN GCS unavailable through COM. Explicit, verified EPSG/WKT and master units are REQUIRED."
End Function
