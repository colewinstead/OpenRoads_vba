Option Explicit

Public Sub Wrapper_SelectedDGNs_AllSheets()

    Dim selectedFiles As Collection
    Dim filePath As String
    Dim originalDgn As String
    Dim answer As VbMsgBoxResult
    Dim item As Variant
    Dim filesProcessed As Long
    Dim filesFailed As Long
    Dim sheetsProcessed As Long
    Dim sheetCount As Long
    Dim failedFiles As String

    Set selectedFiles = New Collection
    originalDgn = ActiveDesignFile.FullName

    Do
        filePath = PickDgnFile()
        If Len(filePath) = 0 Then Exit Do

        If Not CollectionContains(selectedFiles, filePath) Then
            selectedFiles.Add filePath
        End If

        answer = MsgBox( _
            "Added:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
            "Do you want to add another DGN?", _
            vbYesNo + vbQuestion, _
            "Select DGN Files")

        If answer = vbNo Then Exit Do
    Loop

    If selectedFiles.Count = 0 Then
        MsgBox "No DGN files were selected.", vbExclamation
        Exit Sub
    End If

    answer = MsgBox( _
        "Run the recorded macro on every Sheet model in " & _
        selectedFiles.Count & " selected DGN file(s)?", _
        vbYesNo + vbQuestion, _
        "Confirm Batch Run")

    If answer <> vbYes Then Exit Sub

    For Each item In selectedFiles
        filePath = CStr(item)
        sheetCount = 0

        If ProcessDgnFile(filePath, sheetCount) Then
            filesProcessed = filesProcessed + 1
            sheetsProcessed = sheetsProcessed + sheetCount
        Else
            filesFailed = filesFailed + 1
            failedFiles = failedFiles & vbCrLf & filePath
        End If
    Next item

    RestoreOriginalDgn originalDgn

    MsgBox "Batch processing complete." & vbCrLf & vbCrLf & _
           "DGN files processed: " & filesProcessed & vbCrLf & _
           "Sheet models processed: " & sheetsProcessed & vbCrLf & _
           "DGN files failed: " & filesFailed & _
           IIf(filesFailed > 0, vbCrLf & vbCrLf & "Failed files:" & failedFiles, ""), _
           IIf(filesFailed > 0, vbExclamation, vbInformation), _
           "All Sheets Batch"

End Sub

Public Sub Wrapper_SelectedDGNs_AllXSDrawingModels()

    Dim selectedFiles As Collection
    Dim filePath As String
    Dim originalDgn As String
    Dim answer As VbMsgBoxResult
    Dim item As Variant
    Dim filesProcessed As Long
    Dim filesFailed As Long
    Dim modelsProcessed As Long
    Dim modelCount As Long
    Dim failedFiles As String

    Set selectedFiles = New Collection
    originalDgn = ActiveDesignFile.FullName

    Do
        filePath = PickDgnFile()
        If Len(filePath) = 0 Then Exit Do

        If Not CollectionContains(selectedFiles, filePath) Then
            selectedFiles.Add filePath
        End If

        answer = MsgBox( _
            "Added:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
            "Do you want to add another DGN?", _
            vbYesNo + vbQuestion, _
            "Select DGN Files")

        If answer = vbNo Then Exit Do
    Loop

    If selectedFiles.Count = 0 Then
        MsgBox "No DGN files were selected.", vbExclamation
        Exit Sub
    End If

    answer = MsgBox( _
        "Run the recorded macro on every non-sheet station model in " & _
        selectedFiles.Count & " selected DGN file(s)?", _
        vbYesNo + vbQuestion, _
        "Confirm XS Drawing Batch")

    If answer <> vbYes Then Exit Sub

    For Each item In selectedFiles
        filePath = CStr(item)
        modelCount = 0

        If ProcessDgnFileXSDrawings(filePath, modelCount) Then
            filesProcessed = filesProcessed + 1
            modelsProcessed = modelsProcessed + modelCount
        Else
            filesFailed = filesFailed + 1
            failedFiles = failedFiles & vbCrLf & filePath
        End If
    Next item

    RestoreOriginalDgn originalDgn

    MsgBox "XS drawing batch processing complete." & vbCrLf & vbCrLf & _
           "DGN files processed: " & filesProcessed & vbCrLf & _
           "XS drawing models processed: " & modelsProcessed & vbCrLf & _
           "DGN files failed: " & filesFailed & _
           IIf(filesFailed > 0, vbCrLf & vbCrLf & "Failed files:" & failedFiles, ""), _
           IIf(filesFailed > 0, vbExclamation, vbInformation), _
           "All XS Drawings Batch"

End Sub

Public Sub Wrapper_ProjectWise_AllSheets()

    RunProjectWiseBatch False

End Sub

Public Sub Wrapper_ProjectWise_AllXSDrawingModels()

    RunProjectWiseBatch True

End Sub

Public Sub Batch_CurrentDGN_AllSheets()

    Dim modelCount As Long

    ProcessActiveSheetModels modelCount

    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"
    CadInputQueue.SendCommand "FILEDESIGN"
    DoEvents

End Sub

Public Sub Batch_CurrentDGN_AllXSDrawingModels()

    Dim modelCount As Long

    ProcessActiveXSDrawingModels modelCount

    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"
    CadInputQueue.SendCommand "FILEDESIGN"
    DoEvents

End Sub

Public Sub Batch_ActiveModel_Sheet()

    If ActiveModelReference.Type <> msdModelTypeSheet Then Exit Sub

    BmrXS_Grid_trans_Fast

    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"
    CadInputQueue.SendCommand "FILEDESIGN"
    DoEvents

End Sub

Public Sub Batch_ActiveModel_XSDrawing()

    If Not IsXsDrawingModel(ActiveModelReference) Then Exit Sub

    BmrXS_Grid_trans_Fast

    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"
    CadInputQueue.SendCommand "FILEDESIGN"
    DoEvents

End Sub

Private Sub RunProjectWiseBatch(ByVal processXsDrawings As Boolean)

    Dim originalDgn As String
    Dim previousDgn As String
    Dim filesProcessed As Long
    Dim modelsProcessed As Long
    Dim currentModelCount As Long
    Dim answer As VbMsgBoxResult
    Dim batchLabel As String

    originalDgn = ActiveDesignFile.FullName

    If processXsDrawings Then
        batchLabel = "XS drawing models"
    Else
        batchLabel = "Sheet models"
    End If

    Do
        answer = MsgBox( _
            "Click OK to open the Bentley/ProjectWise file dialog." & vbCrLf & vbCrLf & _
            "Select the next DGN containing " & batchLabel & ".", _
            vbOKCancel + vbInformation, _
            "Select ProjectWise DGN")

        If answer <> vbOK Then Exit Do

        previousDgn = ActiveDesignFile.FullName

        If Not OpenDgnUsingBentleyDialog(previousDgn) Then
            answer = MsgBox( _
                "No different DGN was opened." & vbCrLf & vbCrLf & _
                "Do you want to try again?", _
                vbYesNo + vbExclamation, _
                "DGN Not Selected")

            If answer = vbYes Then
                GoTo SelectNextDgn
            Else
                Exit Do
            End If
        End If

        currentModelCount = 0

        If processXsDrawings Then
            ProcessActiveXSDrawingModels currentModelCount
        Else
            ProcessActiveSheetModels currentModelCount
        End If

        filesProcessed = filesProcessed + 1
        modelsProcessed = modelsProcessed + currentModelCount

        CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"
        CadInputQueue.SendCommand "FILEDESIGN"
        DoEvents

        answer = MsgBox( _
            "Finished:" & vbCrLf & ActiveDesignFile.FullName & vbCrLf & vbCrLf & _
            batchLabel & " processed: " & currentModelCount & vbCrLf & vbCrLf & _
            "Open another DGN from ProjectWise?", _
            vbYesNo + vbQuestion, _
            "Continue Batch")

        If answer <> vbYes Then Exit Do

SelectNextDgn:
    Loop

    RestoreOriginalDgn originalDgn

    MsgBox "ProjectWise batch processing complete." & vbCrLf & vbCrLf & _
           "DGN files processed: " & filesProcessed & vbCrLf & _
           batchLabel & " processed: " & modelsProcessed, _
           vbInformation, _
           "ProjectWise Batch"

End Sub

Private Function OpenDgnUsingBentleyDialog(ByVal previousDgn As String) As Boolean

    Dim i As Long

    On Error GoTo OpenFailed

    ' Bentley's native Open dialog invokes ProjectWise when integration is enabled.
    CadInputQueue.SendCommand "DIALOG OPENFILE"

    ' DoEvents allows the queued modal dialog and selected DGN to finish opening.
    For i = 1 To 20
        DoEvents
    Next i

    OpenDgnUsingBentleyDialog = _
        (StrComp(ActiveDesignFile.FullName, previousDgn, vbTextCompare) <> 0)

    Exit Function

OpenFailed:
    OpenDgnUsingBentleyDialog = False

End Function

Private Function ProcessDgnFile(ByVal filePath As String, ByRef sheetCount As Long) As Boolean

    Dim dgn As DesignFile
    Dim modelNames As Collection
    Dim oModel As ModelReference
    Dim modelName As Variant

    On Error GoTo ProcessFailed

    If Len(Dir$(filePath)) = 0 Then Exit Function

    Set dgn = Application.OpenDesignFile(filePath, False)
    Set modelNames = New Collection

    ' Build the list before activating models.
    For Each oModel In ActiveDesignFile.Models
        If oModel.Type = msdModelTypeSheet Then
            modelNames.Add oModel.Name
        End If
    Next oModel

    For Each modelName In modelNames
        Set oModel = ActiveDesignFile.Models.Item(CStr(modelName))
        oModel.Activate
        DoEvents

        RunRecordedMacroOnCurrentSheet
        sheetCount = sheetCount + 1
    Next modelName

    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"
    CadInputQueue.SendCommand "FILEDESIGN"
    DoEvents

    ProcessDgnFile = True
    Exit Function

ProcessFailed:
    ProcessDgnFile = False

End Function

Private Sub ProcessActiveSheetModels(ByRef modelCount As Long)

    Dim modelNames As Collection
    Dim oModel As ModelReference
    Dim modelName As Variant

    Set modelNames = New Collection

    For Each oModel In ActiveDesignFile.Models
        If oModel.Type = msdModelTypeSheet Then
            modelNames.Add oModel.Name
        End If
    Next oModel

    For Each modelName In modelNames
        Set oModel = ActiveDesignFile.Models.Item(CStr(modelName))
        oModel.Activate
        DoEvents

        RunRecordedMacroOnCurrentSheet
        modelCount = modelCount + 1
    Next modelName

End Sub

Private Sub ProcessActiveXSDrawingModels(ByRef modelCount As Long)

    Dim modelNames As Collection
    Dim oModel As ModelReference
    Dim modelName As Variant

    Set modelNames = New Collection

    For Each oModel In ActiveDesignFile.Models
        If IsXsDrawingModel(oModel) Then
            modelNames.Add oModel.Name
        End If
    Next oModel

    For Each modelName In modelNames
        Set oModel = ActiveDesignFile.Models.Item(CStr(modelName))
        oModel.Activate
        DoEvents

        RunRecordedMacroOnCurrentXSDrawingModel
        modelCount = modelCount + 1
    Next modelName

End Sub

Private Function ProcessDgnFileXSDrawings(ByVal filePath As String, ByRef modelCount As Long) As Boolean

    Dim dgn As DesignFile
    Dim modelNames As Collection
    Dim oModel As ModelReference
    Dim modelName As Variant

    On Error GoTo ProcessFailed

    If Len(Dir$(filePath)) = 0 Then Exit Function

    Set dgn = Application.OpenDesignFile(filePath, False)
    Set modelNames = New Collection

    ' XS drawing models are non-sheet models whose names contain a station token.
    For Each oModel In ActiveDesignFile.Models
        If IsXsDrawingModel(oModel) Then
            modelNames.Add oModel.Name
        End If
    Next oModel

    For Each modelName In modelNames
        Set oModel = ActiveDesignFile.Models.Item(CStr(modelName))
        oModel.Activate
        DoEvents

        RunRecordedMacroOnCurrentXSDrawingModel
        modelCount = modelCount + 1
    Next modelName

    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"
    CadInputQueue.SendCommand "FILEDESIGN"
    DoEvents

    ProcessDgnFileXSDrawings = True
    Exit Function

ProcessFailed:
    ProcessDgnFileXSDrawings = False

End Function

Private Function IsXsDrawingModel(ByVal oModel As ModelReference) As Boolean

    If oModel.Type = msdModelTypeSheet Then Exit Function
    If InStr(1, oModel.Name, "[Sheet]", vbTextCompare) > 0 Then Exit Function

    ' Project XS model names use station notation such as 434+50.0000.
    IsXsDrawingModel = (InStr(1, oModel.Name, "+", vbBinaryCompare) > 0)

End Function

Private Sub RunRecordedMacroOnCurrentSheet()

    BmrXS_Grid_trans_Fast

End Sub

Private Sub RunRecordedMacroOnCurrentXSDrawingModel()

    BmrXS_Grid_trans_Fast

End Sub

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

Private Function PickDgnFile() As String

    Dim dialog As Object
    Dim initialFolder As String

    On Error GoTo UsePathPrompt

    Set dialog = CreateObject("UserAccounts.CommonDialog")

    initialFolder = ParentFolder(ActiveDesignFile.FullName)
    If Len(initialFolder) > 0 Then dialog.InitialDir = initialFolder

    dialog.Filter = "DGN Files|*.dgn|All Files|*.*"
    dialog.FilterIndex = 1

    If dialog.ShowOpen Then
        PickDgnFile = dialog.FileName
    End If

    Exit Function

UsePathPrompt:
    Err.Clear
    PickDgnFile = Trim$(InputBox( _
        "The Windows file picker was unavailable." & vbCrLf & _
        "Enter or paste the full path to a DGN file.", _
        "Select DGN File"))

End Function

Private Function CollectionContains(ByVal values As Collection, ByVal searchValue As String) As Boolean

    Dim item As Variant

    For Each item In values
        If StrComp(CStr(item), searchValue, vbTextCompare) = 0 Then
            CollectionContains = True
            Exit Function
        End If
    Next item

End Function

Private Function ParentFolder(ByVal filePath As String) As String

    Dim separatorPosition As Long

    separatorPosition = InStrRev(filePath, "\")
    If separatorPosition > 0 Then
        ParentFolder = Left$(filePath, separatorPosition - 1)
    End If

End Function

Private Sub RestoreOriginalDgn(ByVal originalDgn As String)

    On Error Resume Next

    If Len(originalDgn) > 0 Then
        If StrComp(ActiveDesignFile.FullName, originalDgn, vbTextCompare) <> 0 Then
            Application.OpenDesignFile originalDgn, False
        End If
    End If

    On Error GoTo 0

End Sub
