Attribute VB_Name = "modPerpendicularTests"
Option Explicit

' Run with the generator/picker INACTIVE. Creates in-memory fixtures only.
' Key-in: VBA RUN [PerpendicularLineGenerator]modPerpendicularTests.RunSelfTests
Private Const PI As Double = 3.14159265358979

Private Sub LoadFixture(source As Element)
    Dim reason As String
    If Not PLGLoadPath(source, reason) Then Err.Raise vbObjectError + 2690, , reason
End Sub

Private Sub Near(ByVal actual As Double, ByVal expected As Double, ByVal message As String)
    Dim tolerance As Double
    tolerance = PLGDistanceTolerance() * 2#
    PLGRequire Abs(actual - expected) <= tolerance, message & ": " & CStr(actual) & " versus " & CStr(expected)
End Sub

Public Sub RunSelfTests()
    Dim source As Element, curve As New BsplineCurve, arc As ArcElement
    Dim parts(0 To 1) As ChainableElement, vertices(0 To 2) As Point3d, poles(0 To 3) As Point3d, knots() As Double
    Dim p As Point3d, t As Point3d, a As Point3d, b As Point3d, cursor As Point3d, anchor As Point3d
    Dim crossing(0 To 3) As Point3d, pickedDistance As Double
    Dim distances() As Double, n As Long, i As Long, reason As String
    If PLGPickContextValid() Then
        MsgBox "Finish or cancel the generator before running self-tests.", vbExclamation, "Perpendicular Line Generator tests"
        Exit Sub
    End If
    On Error GoTo Failed
    Set source = CreateLineElement2(Nothing, Point3dFromXYZ(0#, 0#, 0#), Point3dFromXYZ(1000#, 0#, 0#))
    LoadFixture source
    cursor = Point3dFromXYZ(375#, 40#, 999#)
    Near PLGNearestDistance(cursor, False, anchor), 375#, "Straight graphical pick distance (ignore cursor Z)"
    Near anchor.X, 375#, "Straight graphical pick X"
    Near anchor.Y, 0#, "Straight graphical pick Y"
    Near PLGNearestDistance(cursor, True, anchor), 625#, "Reverse graphical pick distance"
    cursor = Point3dFromXYZ(-50#, 40#, 0#)
    Near PLGNearestDistance(cursor, False, anchor), 0#, "Pick before path clamps to beginning"
    cursor = Point3dFromXYZ(1200#, 40#, 0#)
    Near PLGNearestDistance(cursor, False, anchor), 1000#, "Pick after path clamps to end"
    n = PLGBuildDistances(0#, 1000#, 150#, True, False, distances)
    PLGRequire n = 7, "1,000-unit fixture should have 7 lines."
    For i = 1 To n
        Near distances(i), (i - 1) * 150#, "Path spacing"
        PLGEvaluate distances(i), False, p, t
        Near p.X, distances(i), "Straight X"
        Near p.Y, 0#, "Straight Y"
        Near t.X, 1#, "Straight tangent"
    Next i
    n = PLGBuildDistances(0#, 1000#, 150#, True, True, distances)
    PLGRequire n = 8, "Optional end should add one line."
    Near distances(n), 1000#, "Exact optional end"
    n = PLGBuildDistances(0#, 900#, 150#, False, False, distances)
    PLGRequire n = 5, "Both endpoints excluded should leave 150 through 750."
    Near distances(1), 150#, "Excluded start"
    Near distances(n), 750#, "Excluded regular-grid end"
    n = PLGBuildDistances(25#, 925#, 150#, True, True, distances)
    PLGRequire n = 7, "Exact-grid optional end must not be duplicated."
    Near distances(n), 925#, "Offset range end"
    PLGEvaluate 150#, True, p, t
    Near p.X, 850#, "Reverse distance"
    Near t.X, -1#, "Reverse tangent"
    PLGEndpoints p, t, 200#, 1, a, b
    Near a.Y, -200#, "Reverse left side"
    Near b.Y, 0#, "Reverse left anchor"
    PLGEvaluate 150#, False, p, t
    For i = 0 To 2
        PLGEndpoints p, t, 200#, i, a, b
        Near PLGPointDistance(a, b), 200#, "Total length"
        Near (a.X - b.X) * t.X + (a.Y - b.Y) * t.Y, 0#, "Perpendicularity"
        If i = 0 Then
            Near a.Y, 100#, "Centered left"
            Near b.Y, -100#, "Centered right"
        ElseIf i = 1 Then
            Near a.Y, 200#, "Left-only extent"
            Near b.Y, 0#, "Left-only anchor"
        Else
            Near a.Y, 0#, "Right-only anchor"
            Near b.Y, -200#, "Right-only extent"
        End If
    Next i
    ExpectBadSpacing 0#
    ExpectBadSpacing -150#
    ExpectBadSpacing 0.001
    Set source = CreateLineElement2(Nothing, Point3dFromXYZ(1000#, 0#, 0#), Point3dFromXYZ(0#, 0#, 0#))
    LoadFixture source
    PLGEvaluate 150#, False, p, t
    Near p.X, 850#, "Reversed element distance"
    Near t.X, -1#, "Reversed element tangent"
    cursor = Point3dFromXYZ(375#, 40#, 0#)
    Near PLGNearestDistance(cursor, False, anchor), 625#, "Graphical pick on reversed element"

    Set arc = CreateArcElement2(Nothing, Point3dFromXYZ(0#, 0#, 0#), 500#, 500#, Matrix3dIdentity, 0#, PI / 2#)
    LoadFixture arc
    Near PLGPathLength(), 500# * PI / 2#, "Circular arc length"
    PLGEvaluate 150#, False, p, t
    Near p.X, 500# * Cos(0.3), "Circular X"
    Near p.Y, 500# * Sin(0.3), "Circular Y"
    Near t.X, -Sin(0.3), "Circular tangent X"
    Near t.Y, Cos(0.3), "Circular tangent Y"
    cursor = Point3dFromXYZ(600# * Cos(0.3), 600# * Sin(0.3), 999#)
    Near PLGNearestDistance(cursor, False, anchor), 150#, "Circular graphical pick uses arc length, not parameter fraction or chord"
    Near anchor.X, 500# * Cos(0.3), "Circular graphical pick X"
    Near anchor.Y, 500# * Sin(0.3), "Circular graphical pick Y"
    Near PLGNearestDistance(cursor, True, anchor), PLGPathLength() - 150#, "Circular reverse graphical pick"
    PLGEndpoints p, t, 200#, 0, a, b
    Near (a.X - b.X) * t.X + (a.Y - b.Y) * t.Y, 0#, "Circular perpendicularity"

    Set parts(0) = CreateLineElement2(Nothing, Point3dFromXYZ(-1000#, 0#, 0#), Point3dFromXYZ(0#, 0#, 0#))
    Set parts(1) = CreateArcElement2(Nothing, Point3dFromXYZ(0#, 500#, 0#), 500#, 500#, Matrix3dIdentity, -PI / 2#, PI / 2#)
    Set source = CreateComplexStringElement1(parts)
    LoadFixture source
    PLGEvaluate 1050#, False, p, t
    Near p.X, 500# * Sin(0.1), "Cumulative distance across tangent/arc join"
    Near p.Y, 500# * (1# - Cos(0.1)), "Join arc Y"
    cursor = Point3dFromXYZ(p.X - 20# * Sin(0.1), p.Y + 20# * Cos(0.1), 0#)
    Near PLGNearestDistance(cursor, False, anchor), 1050#, "Graphical pick keeps cumulative distance across join"
    cursor = Point3dFromXYZ(0#, -20#, 0#)
    Near PLGNearestDistance(cursor, False, anchor), 1000#, "Adjacent components share one unambiguous join distance"
    For i = -1 To 1
        PLGEvaluate 1000# + i * 0.001, False, p, t
        Near Sqr(t.X * t.X + t.Y * t.Y), 1#, "Join tangent normalization"
    Next i

    ' A cubic Bezier with a known midpoint and derivative; this is NOT a clothoid.
    poles(0) = Point3dFromXYZ(0#, 0#, 0#): poles(1) = Point3dFromXYZ(100#, 100#, 0#)
    poles(2) = Point3dFromXYZ(200#, 100#, 0#): poles(3) = Point3dFromXYZ(300#, 0#, 0#)
    curve.SetPoles poles
    curve.Order = 4
    curve.SetKnots knots
    Set source = CreateBsplineCurveElement1(Nothing, curve)
    LoadFixture source
    PLGEvaluate curve.ComputeCurveLength(0#, 0.5), False, p, t
    Near p.X, 150#, "B-spline X"
    Near p.Y, 75#, "B-spline Y"
    Near t.X, 1#, "B-spline derivative X"
    Near t.Y, 0#, "B-spline derivative Y"
    cursor = Point3dFromXYZ(150#, 200#, 0#)
    pickedDistance = PLGNearestDistance(cursor, False, anchor)
    Near pickedDistance, curve.ComputeCurveLength(0#, 0.5), "B-spline graphical pick path distance"
    Near anchor.X, 150#, "B-spline graphical pick X"
    Near anchor.Y, 75#, "B-spline graphical pick Y"

    vertices(0) = Point3dFromXYZ(0#, 0#, 0#): vertices(1) = Point3dFromXYZ(150#, 0#, 0#)
    vertices(2) = Point3dFromXYZ(150#, 200#, 0#)
    Set source = CreateLineElement1(Nothing, vertices)
    LoadFixture source
    PLGEvaluate 200#, False, p, t
    Near p.X, 150#, "Line string second component X"
    Near p.Y, 50#, "Line string cumulative distance"
    cursor = Point3dFromXYZ(180#, 50#, 0#)
    Near PLGNearestDistance(cursor, False, anchor), 200#, "Line string graphical pick cumulative distance"
    ExpectCorner
    crossing(0) = Point3dFromXYZ(-100#, -100#, 0#): crossing(1) = Point3dFromXYZ(100#, 100#, 0#)
    crossing(2) = Point3dFromXYZ(-100#, 100#, 0#): crossing(3) = Point3dFromXYZ(100#, -100#, 0#)
    Set source = CreateLineElement1(Nothing, crossing)
    LoadFixture source
    ExpectAmbiguousPick
    Set source = CreateLineElement2(Nothing, Point3dFromXYZ(0#, 0#, 0#), Point3dFromXYZ(100#, 0#, 10#))
    PLGRequire Not PLGLoadPath(source, reason), "Changing-Z geometry should be rejected."
    PLGClearPath
    MsgBox "All in-memory self-tests passed. No DGN elements were added." & vbCrLf & _
        "Native civil spirals, form behavior, symbology, write rollback, and Undo still need the manual test plan.", _
        vbInformation, "Perpendicular Line Generator tests"
    Exit Sub
Failed:
    reason = Err.Description
    PLGClearPath
    MsgBox "Self-test FAILED: " & reason, vbCritical, "Perpendicular Line Generator tests"
End Sub

Private Sub ExpectAmbiguousPick()
    Dim cursor As Point3d, anchor As Point3d, distance As Double, errorNumber As Long
    cursor = Point3dFromXYZ(0#, 0#, 0#)
    On Error Resume Next
    distance = PLGNearestDistance(cursor, False, anchor)
    errorNumber = Err.Number
    On Error GoTo 0
    PLGRequire errorNumber <> 0, "Different chain distances at a crossing should require an unambiguous pick."
End Sub

Private Sub ExpectBadSpacing(ByVal spacing As Double)
    Dim distances() As Double, n As Long, errorNumber As Long
    On Error Resume Next
    n = PLGBuildDistances(0#, 1000#, spacing, True, False, distances)
    errorNumber = Err.Number
    On Error GoTo 0
    PLGRequire errorNumber <> 0, "Invalid/excessive spacing was accepted."
End Sub

Private Sub ExpectCorner()
    Dim p As Point3d, t As Point3d, errorNumber As Long
    On Error Resume Next
    PLGEvaluate 150#, False, p, t
    errorNumber = Err.Number
    On Error GoTo 0
    PLGRequire errorNumber <> 0, "A placement at a non-tangent vertex should be rejected."
End Sub
