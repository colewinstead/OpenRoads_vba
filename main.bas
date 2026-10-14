Option Explicit

Private selectedVbaProjectName As String
Private selectedVbaMacroName As String
Private selectedVbaProjectPath As String

' Runs the macro chosen from an .mvba file.
Private Sub RunConfiguredMacro()

    If Len(selectedVbaMacroName) > 0 Then
        CadInputQueue.SendKeyin _
            "VBA RUN [" & selectedVbaProjectName & "]" & selectedVbaMacroName
        DoEvents
    End If

End Sub

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

    If Not SelectBatchMacro() Then Exit Sub

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

    If Not SelectBatchMacro() Then Exit Sub

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

    If Not EnsureBatchMacroSelected() Then Exit Sub

    ProcessActiveSheetModels modelCount

    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"
    CadInputQueue.SendCommand "FILEDESIGN"
    DoEvents

End Sub

Public Sub Batch_CurrentDGN_AllXSDrawingModels()

    Dim modelCount As Long

    If Not EnsureBatchMacroSelected() Then Exit Sub

    ProcessActiveXSDrawingModels modelCount

    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"
    CadInputQueue.SendCommand "FILEDESIGN"
    DoEvents

End Sub

Public Sub Configure_ORD_Batch_Macro()

    selectedVbaProjectName = vbNullString
    selectedVbaMacroName = vbNullString
    selectedVbaProjectPath = vbNullString
    ClearSavedBatchMacroSelection

    If SelectBatchMacro() Then
        MsgBox "Macro selected:" & vbCrLf & vbCrLf & _
               "[" & selectedVbaProjectName & "]" & selectedVbaMacroName & _
               vbCrLf & vbCrLf & _
               "You can now start ORD Batch Process.", _
               vbInformation, _
               "ORD Batch Macro Ready"
    End If

End Sub

Public Sub Batch_ActiveModel_Sheet()

    If ActiveModelReference.Type <> msdModelTypeSheet Then Exit Sub
    If Not EnsureBatchMacroSelected() Then Exit Sub

    RunConfiguredMacro

    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"
    CadInputQueue.SendCommand "FILEDESIGN"
    DoEvents

End Sub

Public Sub Batch_ActiveModel_XSDrawing()

    If Not IsXsDrawingModel(ActiveModelReference) Then Exit Sub
    If Not EnsureBatchMacroSelected() Then Exit Sub

    RunConfiguredMacro

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

    If Not SelectBatchMacro() Then Exit Sub

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

    RunConfiguredMacro

End Sub

Private Sub RunRecordedMacroOnCurrentXSDrawingModel()

    RunConfiguredMacro

End Sub

Private Function SelectBatchMacro() As Boolean

    Dim answer As VbMsgBoxResult
    Dim projectPath As String
    Dim macroName As String
    Dim waitIndex As Long

    selectedVbaProjectName = vbNullString
    selectedVbaMacroName = vbNullString
    selectedVbaProjectPath = vbNullString

    answer = MsgBox( _
        "Click OK to browse for the VBA project (.mvba) to run." & vbCrLf & vbCrLf & _
        "After selecting the project, choose its macro from the displayed list.", _
        vbOKCancel + vbInformation, _
        "Select Batch Macro")

    If answer <> vbOK Then Exit Function

    projectPath = PickVbaProjectFile()
    If Len(projectPath) = 0 Then Exit Function

    If Len(Dir$(projectPath)) = 0 Or _
       LCase$(Right$(projectPath, 5)) <> ".mvba" Then
        MsgBox "Select an existing .mvba VBA project file.", _
               vbExclamation, _
               "Invalid VBA Project"
        Exit Function
    End If

    selectedVbaProjectName = FileBaseName(projectPath)

    ' Load the project so its public macros can be discovered and selected.
    CadInputQueue.SendKeyin "VBA LOAD " & Chr$(34) & projectPath & Chr$(34)
    For waitIndex = 1 To 20
        DoEvents
    Next waitIndex

    macroName = SelectMacroFromVbaProject( _
        projectPath, _
        selectedVbaProjectName)

    If Len(macroName) = 0 Then Exit Function

    selectedVbaProjectPath = projectPath
    selectedVbaMacroName = macroName
    SaveBatchMacroSelection

    SelectBatchMacro = True

End Function

Private Function EnsureBatchMacroSelected() As Boolean

    If Len(selectedVbaProjectName) > 0 And _
       Len(selectedVbaMacroName) > 0 Then
        EnsureBatchMacroSelected = True
    ElseIf LoadBatchMacroSelection() Then
        LoadSelectedVbaProject
        EnsureBatchMacroSelected = True
    Else
        EnsureBatchMacroSelected = SelectBatchMacro()
    End If

End Function

Private Sub SaveBatchMacroSelection()

    Dim fileNumber As Integer

    On Error GoTo SaveFailed

    fileNumber = FreeFile
    Open BatchMacroSettingsPath() For Output As #fileNumber
    Print #fileNumber, selectedVbaProjectPath
    Print #fileNumber, selectedVbaProjectName
    Print #fileNumber, selectedVbaMacroName
    Close #fileNumber
    Exit Sub

SaveFailed:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    On Error GoTo 0

    MsgBox "The ORD batch macro selection could not be saved." & vbCrLf & _
           "Batch Process may ask for the macro again on each model.", _
           vbExclamation, _
           "Batch Selection Not Saved"

End Sub

Private Function LoadBatchMacroSelection() As Boolean

    Dim fileNumber As Integer
    Dim settingsPath As String

    On Error GoTo LoadFailed

    settingsPath = BatchMacroSettingsPath()
    If Len(Dir$(settingsPath)) = 0 Then Exit Function

    fileNumber = FreeFile
    Open settingsPath For Input As #fileNumber
    Line Input #fileNumber, selectedVbaProjectPath
    Line Input #fileNumber, selectedVbaProjectName
    Line Input #fileNumber, selectedVbaMacroName
    Close #fileNumber

    If Len(selectedVbaProjectPath) = 0 Then Exit Function
    If Len(Dir$(selectedVbaProjectPath)) = 0 Then Exit Function
    If Len(selectedVbaProjectName) = 0 Then Exit Function
    If Len(selectedVbaMacroName) = 0 Then Exit Function

    LoadBatchMacroSelection = True
    Exit Function

LoadFailed:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    On Error GoTo 0

    selectedVbaProjectPath = vbNullString
    selectedVbaProjectName = vbNullString
    selectedVbaMacroName = vbNullString

End Function

Private Sub LoadSelectedVbaProject()

    Dim waitIndex As Long

    If Len(selectedVbaProjectPath) = 0 Then Exit Sub

    CadInputQueue.SendKeyin _
        "VBA LOAD " & Chr$(34) & selectedVbaProjectPath & Chr$(34)

    For waitIndex = 1 To 20
        DoEvents
    Next waitIndex

End Sub

Private Function BatchMacroSettingsPath() As String

    BatchMacroSettingsPath = _
        Environ$("TEMP") & "\ORD_Batch_Macro_Selection.txt"

End Function

Private Sub ClearSavedBatchMacroSelection()

    On Error Resume Next

    If Len(Dir$(BatchMacroSettingsPath())) > 0 Then
        Kill BatchMacroSettingsPath()
    End If

    On Error GoTo 0

End Sub

Private Function SelectMacroFromVbaProject( _
    ByVal projectPath As String, _
    ByVal expectedProjectName As String) As String

    Dim vbeObject As Object
    Dim project As Object
    Dim targetProject As Object
    Dim component As Object
    Dim codeModule As Object
    Dim macroNames As Collection
    Dim macroName As String
    Dim lineText As String
    Dim lineNumber As Long
    Dim attempt As Long
    Const StandardModuleType As Long = 1

    On Error GoTo DiscoveryFailed

    Set vbeObject = Application.VBE

    ' The VBA LOAD key-in can take a few message cycles to finish.
    For attempt = 1 To 50
        For Each project In vbeObject.VBProjects
            If StrComp(project.Name, expectedProjectName, vbTextCompare) = 0 Then
                Set targetProject = project
                Exit For
            End If

            On Error Resume Next
            If StrComp(project.FileName, projectPath, vbTextCompare) = 0 Then
                Set targetProject = project
            End If
            On Error GoTo DiscoveryFailed

            If Not targetProject Is Nothing Then Exit For
        Next project

        If Not targetProject Is Nothing Then Exit For
        DoEvents
    Next attempt

    If targetProject Is Nothing Then
        MsgBox "The selected VBA project could not be found after it was loaded.", _
               vbExclamation, _
               "VBA Project Not Loaded"
        Exit Function
    End If

    Set macroNames = New Collection

    For Each component In targetProject.VBComponents
        If component.Type = StandardModuleType Then
            Set codeModule = component.CodeModule

            For lineNumber = 1 To codeModule.CountOfLines
                lineText = codeModule.Lines(lineNumber, 1)
                macroName = PublicNoArgumentSubName(lineText)

                If Len(macroName) > 0 Then
                    macroName = component.Name & "." & macroName

                    If Not CollectionContains(macroNames, macroName) Then
                        macroNames.Add macroName
                    End If
                End If
            Next lineNumber
        End If
    Next component

    If macroNames.Count = 0 Then
        MsgBox "No public macros without arguments were found in:" & vbCrLf & _
               projectPath, _
               vbExclamation, _
               "No Runnable Macros"
        Exit Function
    End If

    If macroNames.Count = 1 Then
        SelectMacroFromVbaProject = CStr(macroNames.Item(1))
    Else
        SelectMacroFromVbaProject = PickMacroFromList(macroNames)
    End If

    Exit Function

DiscoveryFailed:
    MsgBox "The macros in the selected project could not be read." & vbCrLf & vbCrLf & _
           "The project may be password-protected, or access to the VBA project " & _
           "object model may be disabled.", _
           vbExclamation, _
           "Macro Discovery Unavailable"

End Function

Private Function PublicNoArgumentSubName(ByVal declaration As String) As String

    Dim text As String
    Dim lowerText As String
    Dim remainder As String
    Dim openParen As Long
    Dim closeParen As Long
    Dim nameEnd As Long

    text = Trim$(declaration)
    If Len(text) = 0 Then Exit Function
    If Left$(text, 1) = "'" Then Exit Function

    lowerText = LCase$(text)
    If Left$(lowerText, 8) = "private " Then Exit Function
    If Left$(lowerText, 7) = "public " Then text = Trim$(Mid$(text, 8))

    lowerText = LCase$(text)
    If Left$(lowerText, 7) = "static " Then text = Trim$(Mid$(text, 8))

    If LCase$(Left$(text, 4)) <> "sub " Then Exit Function

    remainder = Trim$(Mid$(text, 5))
    openParen = InStr(1, remainder, "(", vbBinaryCompare)

    If openParen > 0 Then
        closeParen = InStr(openParen + 1, remainder, ")", vbBinaryCompare)
        If closeParen = 0 Then Exit Function
        If Len(Trim$(Mid$(remainder, openParen + 1, _
                         closeParen - openParen - 1))) > 0 Then Exit Function
        nameEnd = openParen - 1
    Else
        nameEnd = InStr(1, remainder, " ", vbBinaryCompare) - 1
        If nameEnd < 1 Then nameEnd = Len(remainder)
    End If

    PublicNoArgumentSubName = Trim$(Left$(remainder, nameEnd))

End Function

Private Function PickMacroFromList(ByVal macroNames As Collection) As String

    Dim prompt As String
    Dim selection As String
    Dim selectionNumber As Long
    Dim index As Long
    Dim page As Long
    Dim pageCount As Long
    Dim firstItem As Long
    Dim lastItem As Long
    Dim itemsOnPage As Long
    Const ItemsPerPage As Long = 8

    page = 1
    pageCount = (macroNames.Count + ItemsPerPage - 1) \ ItemsPerPage

    Do
        firstItem = ((page - 1) * ItemsPerPage) + 1
        lastItem = firstItem + ItemsPerPage - 1
        If lastItem > macroNames.Count Then lastItem = macroNames.Count
        itemsOnPage = lastItem - firstItem + 1

        prompt = "Page " & page & " of " & pageCount & ". " & _
                 "Enter 1-" & itemsOnPage & ", N = next, or P = previous." & _
                 vbCrLf & vbCrLf

        For index = firstItem To lastItem
            prompt = prompt & (index - firstItem + 1) & ". " & _
                     CStr(macroNames.Item(index)) & vbCrLf
        Next index

        selection = UCase$(Trim$(InputBox(prompt, "Select VBA Macro")))
        If Len(selection) = 0 Then Exit Function

        If selection = "N" And page < pageCount Then
            page = page + 1
        ElseIf selection = "P" And page > 1 Then
            page = page - 1
        ElseIf IsNumeric(selection) Then
            selectionNumber = CLng(selection)

            If selectionNumber >= 1 And selectionNumber <= itemsOnPage Then
                PickMacroFromList = _
                    CStr(macroNames.Item(firstItem + selectionNumber - 1))
                Exit Function
            End If

            MsgBox "Enter a number from 1 to " & itemsOnPage & ".", _
                   vbExclamation, _
                   "Invalid VBA Macro"
        Else
            MsgBox "Enter a displayed number, N for next, or P for previous.", _
                   vbExclamation, _
                   "Invalid VBA Macro"
        End If
    Loop

End Function

Private Function PickVbaProjectFile() As String

    Dim dialog As Object
    Dim initialFolder As String

    On Error GoTo UseFolderPicker

    Set dialog = CreateObject("UserAccounts.CommonDialog")

    initialFolder = ParentFolder(ActiveDesignFile.FullName)
    If Len(initialFolder) > 0 Then dialog.InitialDir = initialFolder

    dialog.Filter = "MicroStation VBA Projects|*.mvba|All Files|*.*"
    dialog.FilterIndex = 1

    If dialog.ShowOpen Then
        PickVbaProjectFile = dialog.FileName
    End If

    Exit Function

UseFolderPicker:
    Err.Clear
    PickVbaProjectFile = PickVbaProjectFromFolder()

End Function

Private Function PickVbaProjectFromFolder() As String

    Dim shellApp As Object
    Dim selectedFolder As Object
    Dim fileSystem As Object
    Dim folder As Object
    Dim file As Object
    Dim projectFiles As Collection
    Dim prompt As String
    Dim selection As String
    Dim selectionNumber As Long
    Dim item As Variant
    Dim index As Long
    Dim page As Long
    Dim pageCount As Long
    Dim firstItem As Long
    Dim lastItem As Long
    Dim itemsOnPage As Long
    Const ItemsPerPage As Long = 8

    On Error GoTo UsePathPrompt

    Set shellApp = CreateObject("Shell.Application")
    Set selectedFolder = shellApp.BrowseForFolder( _
        0, _
        "Select the folder containing the VBA project (.mvba).", _
        1)

    If selectedFolder Is Nothing Then Exit Function

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folder = fileSystem.GetFolder(selectedFolder.Self.Path)
    Set projectFiles = New Collection

    For Each file In folder.Files
        If LCase$(fileSystem.GetExtensionName(file.Name)) = "mvba" Then
            projectFiles.Add file.Name
        End If
    Next file

    If projectFiles.Count = 0 Then
        MsgBox "No .mvba files were found in the selected folder.", _
               vbExclamation, _
               "VBA Project Not Found"
        Exit Function
    End If

    If projectFiles.Count = 1 Then
        PickVbaProjectFromFolder = folder.Path & "\" & CStr(projectFiles.Item(1))
        Exit Function
    End If

    page = 1
    pageCount = (projectFiles.Count + ItemsPerPage - 1) \ ItemsPerPage

    Do
        firstItem = ((page - 1) * ItemsPerPage) + 1
        lastItem = firstItem + ItemsPerPage - 1
        If lastItem > projectFiles.Count Then lastItem = projectFiles.Count
        itemsOnPage = lastItem - firstItem + 1

        prompt = "Page " & page & " of " & pageCount & ". " & _
                 "Enter 1-" & itemsOnPage & ", N = next, or P = previous." & _
                 vbCrLf & vbCrLf

        For index = firstItem To lastItem
            prompt = prompt & (index - firstItem + 1) & ". " & _
                     CStr(projectFiles.Item(index)) & vbCrLf
        Next index

        selection = UCase$(Trim$(InputBox(prompt, "Select VBA Project")))
        If Len(selection) = 0 Then Exit Function

        If selection = "N" And page < pageCount Then
            page = page + 1
        ElseIf selection = "P" And page > 1 Then
            page = page - 1
        ElseIf IsNumeric(selection) Then
            selectionNumber = CLng(selection)

            If selectionNumber >= 1 And selectionNumber <= itemsOnPage Then
                PickVbaProjectFromFolder = folder.Path & "\" & _
                    CStr(projectFiles.Item(firstItem + selectionNumber - 1))
                Exit Function
            End If

            MsgBox "Enter a number from 1 to " & itemsOnPage & ".", _
                   vbExclamation, _
                   "Invalid VBA Project"
        Else
            MsgBox "Enter a displayed number, N for next, or P for previous.", _
                   vbExclamation, _
                   "Invalid VBA Project"
        End If
    Loop

UsePathPrompt:
    Err.Clear
    PickVbaProjectFromFolder = Trim$(InputBox( _
        "The folder browser was unavailable." & vbCrLf & _
        "Enter or paste the full path to an .mvba file.", _
        "Select VBA Project"))

End Function

Private Function FileBaseName(ByVal filePath As String) As String

    Dim fileName As String
    Dim extensionPosition As Long

    fileName = Mid$(filePath, InStrRev(filePath, "\") + 1)
    extensionPosition = InStrRev(fileName, ".")

    If extensionPosition > 1 Then
        FileBaseName = Left$(fileName, extensionPosition - 1)
    Else
        FileBaseName = fileName
    End If

End Function

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
