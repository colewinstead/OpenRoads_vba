Attribute VB_Name = "RestoreModifiedLevels"
Option Explicit

'Restores level attributes on every Sheet model to their source state.
'
'Reference levels (including nested references) are synchronized from the
'source DGN. Used library levels are synchronized from the configured DGNLIBs.
'This clears intentional overrides too, so test on a copy of the DGN first.
Public Sub RestoreModifiedLevels_AllSheets()

    Dim sheetNames As Collection
    Dim model As ModelReference
    Dim sheetName As Variant
    Dim originalModelName As String
    Dim processedCount As Long
    Dim answer As VbMsgBoxResult
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim errorMessage As String

    answer = MsgBox( _
        "Restore modified level settings on every Sheet model?" & _
        vbCrLf & vbCrLf & _
        "This resets reference and DGNLIB level attributes to their " & _
        "source values, including display, freeze, plot, color, style, " & _
        "weight, and other overrides." & vbCrLf & vbCrLf & _
        "Any intentional sheet-level overrides will also be removed.", _
        vbYesNo + vbExclamation + vbDefaultButton2, _
        "Restore Original Level State")

    If answer <> vbYes Then Exit Sub

    On Error Resume Next
    originalModelName = ActiveModelReference.Name
    On Error GoTo FatalError

    'Collect names first. Activating models while enumerating the Models
    'collection is unreliable in some OpenRoads/MicroStation builds.
    Set sheetNames = New Collection

    For Each model In ActiveDesignFile.Models
        If model.Type = msdModelTypeSheet Then
            sheetNames.Add model.Name
        End If
    Next model

    If sheetNames.Count = 0 Then
        MsgBox "No Sheet models were found in the active DGN.", _
               vbExclamation, "Restore Original Level State"
        Exit Sub
    End If

    For Each sheetName In sheetNames
        Set model = ActiveDesignFile.Models.Item(CStr(sheetName))
        model.Activate
        DoEvents

        'The final ALL targets every direct reference attachment. MicroStation
        'also drills into each attachment's nested references.
        CadInputQueue.SendKeyin "REFERENCE SYNCHRONIZE LEVELS ALL ALL"

        'Restore levels used from configured DGN level libraries as well.
        CadInputQueue.SendKeyin "DGNLIB UPDATE LEVELS ALL"

        DoEvents
        processedCount = processedCount + 1
    Next sheetName

    'Return to the model that was active when the macro started.
    On Error Resume Next
    If Len(originalModelName) > 0 Then
        ActiveDesignFile.Models.Item(originalModelName).Activate
    End If
    On Error GoTo FatalError

    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"

    MsgBox "Level restoration complete." & vbCrLf & vbCrLf & _
           "Sheet models processed: " & CStr(processedCount), _
           vbInformation, "Restore Original Level State"
    Exit Sub

FatalError:
    errorNumber = Err.Number
    errorDescription = Err.Description

    errorMessage = "The macro stopped."

    On Error Resume Next
    errorMessage = errorMessage & vbCrLf & _
                   "Active model: " & ActiveModelReference.Name
    On Error GoTo 0

    errorMessage = errorMessage & vbCrLf & vbCrLf & _
                   "Error " & CStr(errorNumber) & ": " & errorDescription

    MsgBox errorMessage, vbCritical, "Restore Original Level State"
End Sub
