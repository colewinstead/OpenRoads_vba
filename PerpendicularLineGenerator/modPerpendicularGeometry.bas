Attribute VB_Name = "modPerpendicularGeometry"
Option Explicit

' Geometry distances and coordinates are in active-model master units.
' Bentley ORD 2026 MicroStationVBA.chm: BsplineCurve.FromElement,
' ComputeCurveLength, EvaluatePointAtDistance, EvaluatePointTangent.
' shortcut: evaluate exposed VBA geometry only; use a Civil SDK add-in when native spirals or station equations are required.
Public Const PLG_MAX_LINES As Long = 10000
Private Const MAX_PARTS As Long = 10000
Private Const ANGLE_TOL As Double = 0.000001
Private mCurve() As BsplineCurve
Private mLength() As Double
Private mBackwards() As Boolean
Private mCount As Long
Private mTotal As Double
Private mTolerance As Double
Private mPlaneZ As Double
Private mPlaneTolerance As Double

Public Sub PLGClearPath()
    Erase mCurve
    Erase mLength
    Erase mBackwards
    mCount = 0
    mTotal = 0#
End Sub

Public Function PLGPathLength() As Double
    PLGPathLength = mTotal
End Function

Public Function PLGDistanceTolerance() As Double
    PLGDistanceTolerance = mTolerance
End Function

Public Sub PLGRequire(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2601, "Perpendicular Lines", message
End Sub

Public Function PLGPointDistance(a As Point3d, b As Point3d) As Double
    PLGPointDistance = Sqr((a.X - b.X) ^ 2 + (a.Y - b.Y) ^ 2 + (a.Z - b.Z) ^ 2)
End Function

Private Function UnitXY(v As Point3d) As Point3d
    Dim magnitude As Double
    magnitude = Sqr(v.X * v.X + v.Y * v.Y)
    PLGRequire magnitude > 0#, "The geometry has a zero XY tangent (cusp or degenerate curve)."
    UnitXY = Point3dFromXYZ(v.X / magnitude, v.Y / magnitude, 0#)
End Function

Private Sub SameTangent(a As Point3d, b As Point3d)
    Dim u As Point3d, v As Point3d
    u = UnitXY(a)
    v = UnitXY(b)
    PLGRequire PLGPointDistance(u, v) <= ANGLE_TOL, _
        "A checked spline knot or requested component join has a corner, cusp, or discontinuous tangent. Use smooth geometry or avoid a non-tangent component join."
End Sub

Public Function PLGLoadPath(source As Element, ByRef reason As String) As Boolean
    Dim expected As Point3d, ending As Point3d, sourceLength As Double
    On Error GoTo Failed
    PLGClearPath
    PLGRequire source.IsTraversableElement, "This element does not expose traversable VBA geometry."
    sourceLength = GetElementPathLength(source)
    PLGRequire sourceLength > 0# And sourceLength <= 1000000000#, "Alignment length must be positive and at most 1 billion master units."
    mTolerance = 2# / ActiveModelReference.UORsPerMasterUnit
    If mTolerance < sourceLength * 0.0000000001 Then mTolerance = sourceLength * 0.0000000001
    ReDim mCurve(1 To MAX_PARTS)
    ReDim mLength(1 To MAX_PARTS)
    ReDim mBackwards(1 To MAX_PARTS)
    expected = source.AsTraversableElement.PointAtDistance(0#)
    ending = source.AsTraversableElement.PointAtDistance(sourceLength)
    mPlaneZ = expected.Z
    mPlaneTolerance = Abs(mPlaneZ) * 0.00000000000001
    If mPlaneTolerance < 0.000000000001 Then mPlaneTolerance = 0.000000000001
    PLGRequire PLGPointDistance(expected, ending) > mTolerance, "Closed alignments are not supported; use an open path."
    AppendElement source, expected, 0
    PLGRequire mCount > 0, "No supported geometry components were found."
    PLGRequire PLGPointDistance(expected, ending) <= mTolerance, "Component order does not match the selected element's direction."
    PLGRequire Abs(mTotal - sourceLength) <= mTolerance, "Component lengths disagree with the selected geometry; civil/approximated geometry is not supported."
    PLGRequire mTotal > mTolerance, "Alignment is shorter than the model distance tolerance."
    reason = "Supported VBA path: " & CStr(mCount) & " component(s), length " & Format(mTotal, "0.00")
    PLGLoadPath = True
    Exit Function
Failed:
    reason = Err.Description & vbCrLf & _
        "Native ORD spirals/custom civil elements require a Civil SDK evaluator. No lines were created."
    PLGClearPath
End Function

Private Function GetElementPathLength(source As Element) As Double
    Select Case source.Type
        Case msdElementTypeLine, msdElementTypeLineString
            GetElementPathLength = source.AsLineElement.Length
        Case msdElementTypeArc
            GetElementPathLength = source.AsArcElement.Length
        Case msdElementTypeBsplineCurve
            GetElementPathLength = source.AsBsplineCurveElement.Length
        Case msdElementTypeComplexString
            GetElementPathLength = source.AsComplexStringElement.Length
        Case Else
            Err.Raise vbObjectError + 2602, , "Unsupported element type " & CStr(source.Type) & ". Its native geometry cannot be evaluated by this tool."
    End Select
End Function

Private Sub AppendElement(source As Element, expected As Point3d, ByVal depth As Long)
    Dim ee As ElementEnumerator, child As Element, vertices() As Point3d
    Dim i As Long, first As Long, last As Long, stepIndex As Long
    Dim piece As LineElement, curve As BsplineCurve
    PLGRequire depth <= 16, "Complex geometry nesting exceeds 16 levels."
    Select Case source.Type
        Case msdElementTypeComplexString
            Set ee = source.AsComplexStringElement.GetSubElements
            Do While ee.MoveNext
                Set child = ee.Current
                AppendElement child, expected, depth + 1
            Loop
        Case msdElementTypeLineString
            vertices = source.AsVertexList.GetVertices
            first = LBound(vertices): last = UBound(vertices): stepIndex = 1
            PLGRequire last > first, "The line string has fewer than two vertices."
            If PLGPointDistance(expected, vertices(first)) > mTolerance Then
                first = UBound(vertices): last = LBound(vertices): stepIndex = -1
            End If
            For i = first To last - stepIndex Step stepIndex
                Set piece = CreateLineElement2(Nothing, vertices(i), vertices(i + stepIndex))
                Set curve = New BsplineCurve
                curve.FromElement piece
                AppendCurve curve, expected
            Next i
        Case msdElementTypeLine, msdElementTypeArc, msdElementTypeBsplineCurve
            Set curve = New BsplineCurve
            curve.FromElement source
            AppendCurve curve, expected
        Case Else
            Err.Raise vbObjectError + 2602, , "Unsupported element/component type " & CStr(source.Type) & _
                ". Acceptable types: line, line string, arc, B-spline, and complex strings made only from these."
    End Select
End Sub

Private Sub AppendCurve(curve As BsplineCurve, expected As Point3d)
    Dim a As Point3d, b As Point3d, pole As Point3d
    Dim i As Long, curveLength As Double, backwards As Boolean
    PLGRequire mCount < MAX_PARTS, "Alignment exceeds 10,000 geometry components."
    PLGRequire Not curve.Closed, "Closed curves are not supported; use an open alignment."
    a = curve.EvaluatePoint(0#): b = curve.EvaluatePoint(1#)
    PLGRequire PLGPointDistance(a, b) > mTolerance, "Closed or zero-length component is not supported."
    For i = 0 To curve.PolesCount - 1
        pole = curve.Pole(i)
        PLGRequire Abs(pole.Z - mPlaneZ) <= mPlaneTolerance, _
            "Only horizontal geometry is supported. A component has changing Z or a tilted plane."
    Next i
    If PLGPointDistance(expected, a) > mTolerance Then
        PLGRequire PLGPointDistance(expected, b) <= mTolerance, "Alignment components have a gap or unsupported ordering."
        backwards = True
    End If
    curveLength = curve.ComputeCurveLength
    PLGRequire curveLength > mTolerance, "A component is zero length or below model resolution."
    ' Reject internal spline discontinuities instead of choosing an arbitrary derivative.
    CheckSplineKnots curve
    mCount = mCount + 1
    Set mCurve(mCount) = curve
    mLength(mCount) = curveLength
    mBackwards(mCount) = backwards
    mTotal = mTotal + curveLength
    If backwards Then expected = a Else expected = b
End Sub

Private Sub CheckSplineKnots(curve As BsplineCurve)
    Dim i As Long, k As Double, h As Double, leftK As Double, rightK As Double
    Dim a As Point3d, b As Point3d, ta As Point3d, tb As Point3d
    For i = 0 To curve.KnotsCount - 1
        k = curve.Knot(i)
        If k > 0# And k < 1# Then
            leftK = 0#: rightK = 1#
            If i > 0 Then
                If curve.Knot(i - 1) < k Then leftK = curve.Knot(i - 1)
            End If
            If i < curve.KnotsCount - 1 Then
                If curve.Knot(i + 1) > k Then rightK = curve.Knot(i + 1)
            End If
            h = (k - leftK) * 0.00000001
            If h > (rightK - k) * 0.00000001 Then h = (rightK - k) * 0.00000001
            PLGRequire k - h < k And k + h > k, "Spline knot intervals are too small to evaluate reliably."
            a = curve.EvaluatePointTangent(ta, k - h)
            b = curve.EvaluatePointTangent(tb, k + h)
            SameTangent ta, tb
            ' Compare each side with its limit; do not mistake a gap for a curve.
            a = curve.EvaluatePoint(k)
            b = curve.EvaluatePoint(k - h * 0.01)
            PLGRequire PLGPointDistance(a, b) <= mTolerance, "Spline has a discontinuous internal knot."
            b = curve.EvaluatePoint(k + h * 0.01)
            PLGRequire PLGPointDistance(a, b) <= mTolerance, "Spline has a discontinuous internal knot."
        End If
    Next i
End Sub

Private Sub EvaluatePart(ByVal index As Long, ByVal distance As Double, p As Point3d, t As Point3d)
    Dim parameter As Double, checkPoint As Point3d, residual As Double
    If mBackwards(index) Then distance = mLength(index) - distance
    If distance <= 0# Then
        parameter = 0#: p = mCurve(index).EvaluatePoint(0#)
    ElseIf distance >= mLength(index) Then
        parameter = 1#: p = mCurve(index).EvaluatePoint(1#)
    Else
        p = mCurve(index).EvaluatePointAtDistance(parameter, distance, 0#, 0.000000000001)
    End If
    PLGRequire parameter >= 0# And parameter <= 1#, "Curve evaluator returned an invalid parameter."
    residual = Abs(mCurve(index).ComputeCurveLength(0#, parameter) - distance)
    PLGRequire residual <= mTolerance, "Curve distance inversion failed its independent length check."
    checkPoint = mCurve(index).EvaluatePointTangent(t, parameter)
    PLGRequire PLGPointDistance(p, checkPoint) <= mTolerance, "Curve point and tangent evaluators disagree."
    t = UnitXY(t)
    If mBackwards(index) Then t = Point3dFromXYZ(-t.X, -t.Y, 0#)
End Sub

Public Sub PLGEvaluate(ByVal distance As Double, ByVal reverse As Boolean, p As Point3d, t As Point3d)
    Dim i As Long, accumulated As Double, partDistance As Double
    Dim otherPoint As Point3d, otherTangent As Point3d
    PLGRequire mCount > 0, "No alignment geometry is loaded."
    PLGRequire distance >= 0# And distance <= mTotal, "Distance is outside the geometry."
    If reverse Then distance = mTotal - distance
    For i = 1 To mCount
        partDistance = distance - accumulated
        If partDistance <= mLength(i) Or i = mCount Then
            If partDistance < 0# Then partDistance = 0#
            If partDistance > mLength(i) Then partDistance = mLength(i)
            EvaluatePart i, partDistance, p, t
            If partDistance <= mTolerance And i > 1 Then
                EvaluatePart i - 1, mLength(i - 1), otherPoint, otherTangent
                SameTangent t, otherTangent
            End If
            If mLength(i) - partDistance <= mTolerance And i < mCount Then
                EvaluatePart i + 1, 0#, otherPoint, otherTangent
                SameTangent t, otherTangent
            End If
            If reverse Then t = Point3dFromXYZ(-t.X, -t.Y, 0#)
            Exit Sub
        End If
        accumulated = accumulated + mLength(i)
    Next i
End Sub

Public Function PLGNearestDistance(cursor As Point3d, ByVal reverse As Boolean, anchor As Point3d) As Double
    Dim i As Long, bestIndex As Long, parameter As Double, accumulated As Double, partDistance As Double
    Dim distances() As Double, pathDistances() As Double, points() As Point3d
    Dim checkPoint As Point3d, rotation As Matrix3d, measured As Double
    PLGRequire mCount > 0, "No alignment geometry is loaded."
    ReDim distances(1 To mCount)
    ReDim pathDistances(1 To mCount)
    ReDim points(1 To mCount)
    rotation = Matrix3dIdentity
    For i = 1 To mCount
        distances(i) = mCurve(i).ComputeMinimumDistance(points(i), parameter, cursor, rotation)
        PLGRequire parameter >= 0# And parameter <= 1#, "Closest-point evaluator returned a parameter outside the curve."
        checkPoint = mCurve(i).EvaluatePoint(parameter)
        PLGRequire PLGPointDistance(points(i), checkPoint) <= mTolerance, "Closest-point and curve evaluators disagree."
        measured = Sqr((cursor.X - points(i).X) ^ 2 + (cursor.Y - points(i).Y) ^ 2)
        PLGRequire Abs(measured - distances(i)) <= mTolerance, "Closest-point evaluator returned an inconsistent XY distance."
        partDistance = mCurve(i).ComputeCurveLength(0#, parameter)
        If mBackwards(i) Then partDistance = mLength(i) - partDistance
        pathDistances(i) = accumulated + partDistance
        If bestIndex = 0 Then
            bestIndex = i
        ElseIf distances(i) < distances(bestIndex) Then
            bestIndex = i
        End If
        accumulated = accumulated + mLength(i)
    Next i
    For i = 1 To mCount
        If distances(i) <= distances(bestIndex) + mTolerance Then
            PLGRequire Abs(pathDistances(i) - pathDistances(bestIndex)) <= mTolerance, _
                "This pick is equally close to different alignment locations. Move away from the crossing or overlap."
        End If
    Next i
    ' shortcut: Bentley returns one minimum per B-spline; use typed distances when that component has multiple equally close locations.
    PLGNearestDistance = pathDistances(bestIndex)
    If reverse Then PLGNearestDistance = mTotal - PLGNearestDistance
    If PLGNearestDistance <= mTolerance Then PLGNearestDistance = 0#
    If mTotal - PLGNearestDistance <= mTolerance Then PLGNearestDistance = mTotal
    anchor = points(bestIndex)
End Function

Public Function PLGBuildDistances(ByVal first As Double, ByVal last As Double, ByVal spacing As Double, _
        ByVal includeFirst As Boolean, ByVal includeLast As Boolean, distances() As Double) As Long
    Dim k As Long, maxK As Long, n As Long, d As Double, ratio As Double
    PLGRequire mCount > 0, "No alignment geometry is loaded."
    PLGRequire spacing > 4# * mTolerance And spacing <= 1000000000#, "Spacing must be positive, exceed four distance tolerances, and be at most 1 billion master units."
    PLGRequire first >= 0# And last <= mTotal And last - first > mTolerance, "Require 0 <= start < end <= geometry length."
    ratio = (last - first) / spacing
    PLGRequire ratio <= PLG_MAX_LINES + 1#, "More than 10,000 lines requested. Increase spacing or shorten the range."
    maxK = CLng(Fix(ratio))
    ReDim distances(1 To PLG_MAX_LINES + 2)
    For k = 0 To maxK
        d = first + CDbl(k) * spacing
        If d < last - mTolerance And (k > 0 Or includeFirst) Then
            n = n + 1: distances(n) = d
        End If
    Next k
    If includeLast Then n = n + 1: distances(n) = last
    PLGRequire n > 0, "The settings request no lines. Include an endpoint or reduce spacing."
    PLGRequire n <= PLG_MAX_LINES, "More than 10,000 lines requested. Increase spacing or shorten the range."
    ReDim Preserve distances(1 To n)
    PLGBuildDistances = n
End Function

Public Sub PLGEndpoints(p As Point3d, t As Point3d, ByVal totalLength As Double, _
        ByVal placement As Long, a As Point3d, b As Point3d)
    Dim leftDistance As Double, rightDistance As Double
    PLGRequire totalLength > 4# * mTolerance And totalLength <= 1000000000#, "Line length must be positive, exceed four distance tolerances, and be at most 1 billion master units."
    Select Case placement
        Case 0: leftDistance = totalLength / 2#: rightDistance = totalLength / 2#
        Case 1: leftDistance = totalLength
        Case 2: rightDistance = totalLength
        Case Else: Err.Raise vbObjectError + 2603, , "Choose a placement option."
    End Select
    a = Point3dFromXYZ(p.X - t.Y * leftDistance, p.Y + t.X * leftDistance, p.Z)
    b = Point3dFromXYZ(p.X + t.Y * rightDistance, p.Y - t.X * rightDistance, p.Z)
    PLGRequire Abs(PLGPointDistance(a, b) - totalLength) <= mTolerance, "Line length failed the coordinate precision check."
End Sub
