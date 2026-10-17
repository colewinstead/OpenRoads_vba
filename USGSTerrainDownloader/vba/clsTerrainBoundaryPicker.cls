VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsTerrainBoundaryPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IPrimitiveCommandEvents
Implements ILocateCommandEvents

Private mFirst As Point3d
Private mHaveFirst As Boolean
Private mFinishing As Boolean

Private Sub RectangleVertices(last As Point3d, vertices() As Point3d)
    ReDim vertices(0 To 4)
    vertices(0) = mFirst
    vertices(1) = Point3dFromXYZ(last.X, mFirst.Y, mFirst.Z)
    vertices(2) = Point3dFromXYZ(last.X, last.Y, mFirst.Z)
    vertices(3) = Point3dFromXYZ(mFirst.X, last.Y, mFirst.Z)
    vertices(4) = mFirst
End Sub

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "USGS Terrain Downloader - Rectangle"
    ShowPrompt "Pick first corner. Rectangle follows model XY, even in a rotated view. Reset cancels."
    CommandState.EnableAccuSnap
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    Dim vertices() As Point3d, tolerance As Double
    On Error GoTo Failed
    TerrainCheckContext
    If Not mHaveFirst Then
        mFirst = Point: mHaveFirst = True
        ShowPrompt "Pick opposite corner. Preview is temporary; no DGN element will be added."
        CommandState.StartDynamics
    Else
        tolerance = 2# / ActiveModelReference.UORsPerMasterUnit
        TerrainRequire Abs(Point.X - mFirst.X) > tolerance And Abs(Point.Y - mFirst.Y) > tolerance, "Rectangle needs positive XY width and height."
        RectangleVertices Point, vertices
        TerrainAcceptBoundary vertices
        Finish
    End If
    Exit Sub
Failed:
    ShowPrompt Err.Description
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
    Dim vertices() As Point3d, preview As LineElement
    On Error GoTo Failed
    TerrainCheckContext
    If Not mHaveFirst Then Exit Sub
    RectangleVertices Point, vertices
    Set preview = CreateLineElement1(Nothing, vertices)
    preview.Redraw DrawMode
    Exit Sub
Failed:
    ShowPrompt Err.Description
End Sub

Public Sub StopPick()
    mFinishing = True
    CommandState.StartDefaultCommand
End Sub

Private Sub Finish()
    StopPick
    TerrainPickerEnded
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    Finish
End Sub
Private Sub IPrimitiveCommandEvents_Cleanup()
    If Not mFinishing Then TerrainPickerEnded
End Sub
Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub

Private Sub ILocateCommandEvents_Start()
    Dim criteria As LocateCriteria
    Set criteria = CommandState.CreateLocateCriteria(False)
    CommandState.SetLocateCriteria criteria
    ShowCommand "USGS Terrain Downloader - Closed Shape"
    ShowPrompt "Identify a straight-sided shape in the active model, then accept. Reset cancels."
End Sub

Private Sub ILocateCommandEvents_LocateFilter(ByVal Element As Element, Point As Point3d, Accepted As Boolean)
    On Error GoTo Failed
    Accepted = False
    If Element.ModelReference.IsAttachment Then Exit Sub
    If Not Element.ModelReference.IsActive Then Exit Sub
    Accepted = (Element.Type = msdElementTypeShape)
    If Not Accepted Then ShowPrompt "Select a simple closed shape. Complex/curved shapes, holes, cells and references are unsupported."
    Exit Sub
Failed:
    Accepted = False
    ShowPrompt Err.Description
End Sub

Private Sub ILocateCommandEvents_Accept(ByVal Element As Element, Point As Point3d, ByVal View As View)
    Dim vertices() As Point3d
    On Error GoTo Failed
    TerrainCheckContext
    TerrainRequire Element.ModelReference.IsActive And Not Element.ModelReference.IsAttachment, "Select in the active model."
    TerrainRequire Element.Type = msdElementTypeShape, "Only simple closed shapes are supported."
    vertices = Element.AsVertexList.GetVertices
    TerrainAcceptBoundary vertices
    Finish
    Exit Sub
Failed:
    ShowPrompt Err.Description
End Sub

Private Sub ILocateCommandEvents_LocateReset()
    Finish
End Sub
Private Sub ILocateCommandEvents_LocateFailed()
    ShowPrompt "No supported shape found. Reset cancels."
End Sub
Private Sub ILocateCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
End Sub
Private Sub ILocateCommandEvents_Cleanup()
    If Not mFinishing Then TerrainPickerEnded
End Sub
