VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsPerpendicularLocate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ILocateCommandEvents
Implements IPrimitiveCommandEvents

Private mTarget As Long
Private mReverse As Boolean
Private mPreviewLength As Double
Private mPlacement As Long
Private mFirstDistance As Double
Private mFollowWithEnd As Boolean
Private mFinishing As Boolean

Public Sub ConfigurePick(ByVal target As Long, ByVal reverse As Boolean, ByVal previewLength As Double, _
        ByVal placement As Long, ByVal firstDistance As Double, ByVal followWithEnd As Boolean)
    mTarget = target
    mReverse = reverse
    mPreviewLength = previewLength
    mPlacement = placement
    mFirstDistance = firstDistance
    mFollowWithEnd = followWithEnd
End Sub

Private Sub ILocateCommandEvents_Start()
    Dim criteria As LocateCriteria
    Set criteria = CommandState.CreateLocateCriteria(False)
    CommandState.SetLocateCriteria criteria
    ShowCommand "Perpendicular Line Generator"
    ShowPrompt "Identify horizontal geometry, then accept. Reset cancels."
End Sub

Private Sub ILocateCommandEvents_LocateFilter(ByVal Element As Element, Point As Point3d, Accepted As Boolean)
    On Error GoTo Rejected
    Accepted = False
    If Element.ModelReference.IsAttachment Then Exit Sub
    If Not Element.ModelReference.IsActive Then Exit Sub
    Select Case Element.Type
        Case msdElementTypeLine, msdElementTypeLineString, msdElementTypeArc, _
                msdElementTypeBsplineCurve, msdElementTypeComplexString
            Accepted = True
        Case Else
            ShowPrompt "Unsupported type. Native civil spirals need a Civil SDK evaluator. Select a standard linear element."
    End Select
    Exit Sub
Rejected:
    Accepted = False
    ShowPrompt Err.Description
End Sub

Private Sub ILocateCommandEvents_Accept(ByVal Element As Element, Point As Point3d, ByVal View As View)
    On Error GoTo Failed
    If Not PLGBeginAlignment(Element) Then CommandState.StartLocate New clsPerpendicularLocate
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Perpendicular Line Generator"
    CommandState.StartDefaultCommand
End Sub

Private Sub ILocateCommandEvents_LocateReset()
    CommandState.StartDefaultCommand
End Sub

Private Sub ILocateCommandEvents_LocateFailed()
    ShowPrompt "No supported element found. Select a line, line string, arc, B-spline, or complex string. Reset cancels."
End Sub

Private Sub ILocateCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub ILocateCommandEvents_Cleanup()
End Sub

Private Sub IPrimitiveCommandEvents_Start()
    Dim direction As String
    If mReverse Then direction = "Reverse" Else direction = "Forward"
    ShowCommand "Perpendicular Line Generator - Range"
    If mTarget = 1 Then
        ShowPrompt "Pick START (" & direction & ") on the alignment. Move for preview, data point to accept. Reset opens settings."
    Else
        ShowPrompt "Pick END after the start in the chosen direction. Move for preview, data point to accept. Reset opens settings."
    End If
    CommandState.EnableAccuSnap
    CommandState.StartDynamics
End Sub

Private Sub DrawLimit(ByVal distance As Double, ByVal drawMode As MsdDrawingMode)
    Dim p As Point3d, t As Point3d, a As Point3d, b As Point3d, preview As LineElement
    PLGEvaluate distance, mReverse, p, t
    PLGEndpoints p, t, mPreviewLength, mPlacement, a, b
    Set preview = CreateLineElement2(Nothing, a, b)
    preview.Redraw drawMode
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
    Dim distance As Double, anchor As Point3d, prompt As String
    On Error GoTo Failed
    PLGRequire PLGPickContextValid(), "The active model changed. Reset or start the tool again."
    distance = PLGNearestDistance(Point, mReverse, anchor)
    If mTarget = 2 Then DrawLimit mFirstDistance, DrawMode
    DrawLimit distance, DrawMode
    If mTarget = 1 Then prompt = "START" Else prompt = "END"
    prompt = prompt & " geometry distance: " & Format(distance, "0.000") & " master units. Click to accept; Reset opens settings."
    If mTarget = 2 And distance <= mFirstDistance + PLGDistanceTolerance() Then
        prompt = "END must follow START in the chosen direction. Reset to change Direction or choose the full alignment."
    End If
    ShowPrompt prompt
    Exit Sub
Failed:
    ShowPrompt Err.Description
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim distance As Double, anchor As Point3d, p As Point3d, t As Point3d
    On Error GoTo Failed
    PLGRequire PLGPickContextValid(), "The active model changed. Reset or start the tool again."
    distance = PLGNearestDistance(Point, mReverse, anchor)
    PLGEvaluate distance, mReverse, p, t
    If mTarget = 2 Then
        PLGRequire distance > mFirstDistance + PLGDistanceTolerance(), _
            "END must follow START in the chosen direction. Pick farther along, or Reset to change Direction."
    End If
    FinishPick True, distance
    Exit Sub
Failed:
    ShowPrompt Err.Description
End Sub

Private Sub FinishPick(ByVal accepted As Boolean, ByVal distance As Double)
    StopPick
    PLGFinishRangePick mTarget, accepted, distance, mFollowWithEnd
End Sub

Public Sub StopPick()
    mFinishing = True
    CommandState.StartDefaultCommand
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    FinishPick False, 0#
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
    If Not mFinishing Then PLGCloseSession
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub
