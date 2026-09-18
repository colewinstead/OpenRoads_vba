Option Explicit

Public Sub BmrXS_Grid_trans_Fast()

    Dim startPoint As Point3d
    Dim point As Point3d

    ' The recording turned off roughly 1,300 levels with a redraw after each one.
    CadInputQueue.SendKeyin "SET LEVELS OFF ALL;SELVIEW 1"

    CadInputQueue.SendKeyin "ON=""XS_SHEET_MINOR_GRIDS"";SELVIEW 1"
    CadInputQueue.SendKeyin "ON=""XS_SHEET_MEDIUM_GRIDS"";SELVIEW 1"
    CadInputQueue.SendKeyin "ON=""XS_SHEET_MAJOR_GRIDS"";SELVIEW 1"

    CadInputQueue.SendCommand "PLACE FENCE ICON"

    startPoint.X = -54075.1839491899
    startPoint.Y = 52560.8125842561
    startPoint.Z = 0#

    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X + 126120.02492787
    point.Y = startPoint.Y - 100622.074311266
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "CHANGE ATTRIBUTES"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES USEACTIVE ON"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE LEVEL"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE COLOR"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE LINESTYLE"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE WEIGHT"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE TRANSPARENCY"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE PRIORITY"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE ELEMENTCLASS"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE TEMPLATE"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE FILLCOLOR"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES MAKECOPY OFF"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENTIREELEMENT OFF"

    SetCExpressionValue "tcb->msToolSettings.general.useFence", 1, "CHANGEATTRIBS"
    CadInputQueue.SendCommand "LOCK FENCE INSIDE"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET TRANSPARENCY 40"

    point.X = startPoint.X + 79985.2106144282
    point.Y = startPoint.Y - 27032.7960836237
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendKeyin "ON=""XS_SHEET_LABELS_STATION"";SELVIEW 1"
    CadInputQueue.SendKeyin "ON=""XS_SHEET_LABELS_OFFSET"";SELVIEW 1"
    CadInputQueue.SendKeyin "ON=""XS_SHEET_LABELS_ELEVATION"";SELVIEW 1"
    CadInputQueue.SendKeyin "ON=""XS_P_TX_FS_BS"";SELVIEW 1"
    CadInputQueue.SendKeyin "ON=""XS_P_FINISHED_GRADE_ELEV_TX"";SELVIEW 1"
    CadInputQueue.SendKeyin "ON=""SHEET_BORDER"";SELVIEW 1"
    CadInputQueue.SendKeyin "ON=""Draft_Named_Boundary"";SELVIEW 1"
    CadInputQueue.SendKeyin "ON=""SHEET_TITLE_BLOCK_TX"";SELVIEW 1"
    CadInputQueue.SendKeyin "ON=""SHEET_TITLE_BLOCK"";SELVIEW 1"
    CadInputQueue.SendKeyin "ON=""SHEET_PRELIM"";SELVIEW 1"
    CadInputQueue.SendKeyin "ON=""SHEET_PLOTTING_SHAPE"";SELVIEW 1"
    CadInputQueue.SendKeyin "ON=""SHEET_MDOT_LOGO"";SELVIEW 1"
    CadInputQueue.SendKeyin "ON=""SHEET_LOGO"";SELVIEW 1"
    CadInputQueue.SendKeyin "ON=""SHEET_CONSULTANT_STAMP"";SELVIEW 1"

    CadInputQueue.SendCommand "FIT VIEW EXTENDED 1"
    CommandState.StartDefaultCommand

End Sub
