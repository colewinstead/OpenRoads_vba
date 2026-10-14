Attribute VB_Name = "AttachTitleBlock_Rebuilt"
Option Explicit

'Attaches one selected title-block DGN to every sheet model.
'If that DGN is already referenced in a sheet, the sheet is skipped.
'For each missing sheet, Attachment Properties opens for the user to click OK.
Public Sub AttachTitleBlockToAllSheets_REBUILT()

    Const DEFAULT_TITLE_BLOCK As String = _
        "PW_WORKDIR:d0204989\TITLE_BLOCK.dgn"

    Dim titleBlockPath As String
    Dim sheetNames As Collection
    Dim sheetModel As ModelReference
    Dim sheetName As Variant
    Dim originalModelName As String
    Dim sheetCount As Long
    Dim attachedCount As Long
    Dim skippedCount As Long
    Dim cancelledCount As Long

    titleBlockPath = Trim$(InputBox( _
        "Paste the full path of the title-block DGN to attach.", _
        "Select Title Block", _
        DEFAULT_TITLE_BLOCK))

    If Len(titleBlockPath) = 0 Then
        MsgBox "Cancelled. No title-block path was entered.", _
               vbExclamation, "Attach Title Block"
        Exit Sub
    End If

    On Error Resume Next
    originalModelName = ActiveModelReference.Name
    On Error GoTo FatalError

    'Store the sheet names before activating or modifying any model.
    Set sheetNames = New Collection

    For Each sheetModel In ActiveDesignFile.Models
        If sheetModel.Type = msdModelTypeSheet Then
            sheetNames.Add sheetModel.Name
        End If
    Next sheetModel

    sheetCount = sheetNames.Count

    If sheetCount = 0 Then
        MsgBox "No sheet models were found in the active DGN.", _
               vbExclamation, "Attach Title Block"
        Exit Sub
    End If

    For Each sheetName In sheetNames
        Set sheetModel = ActiveDesignFile.Models.Item(CStr(sheetName))

        'Activate first so ORD loads the sheet's current reference cache.
        sheetModel.Activate

        If TitleBlockDgnIsAttached(ActiveModelReference, titleBlockPath) Then
            skippedCount = skippedCount + 1
        Else
            'Supplying only the file path intentionally opens Attachment
            'Properties. Select/confirm the model and click OK for this sheet.
            CadInputQueue.SendKeyin _
                "REFERENCE ATTACH " & QuoteText(titleBlockPath)

            'End the reference command after the dialog is accepted/cancelled.
            CadInputQueue.SendKeyin "NULL"

            'The dialog has closed, so verify the attachment by DGN filename.
            If TitleBlockDgnIsAttached(ActiveModelReference, titleBlockPath) Then
                attachedCount = attachedCount + 1
            Else
                cancelledCount = cancelledCount + 1
            End If
        End If
    Next sheetName

    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"

    'Return to the model that was active when the macro started.
    On Error Resume Next
    If Len(originalModelName) > 0 Then
        ActiveDesignFile.Models.Item(originalModelName).Activate
    End If
    On Error GoTo 0

    MsgBox "Title-block processing complete." & vbCrLf & vbCrLf & _
           "Sheet models: " & sheetCount & vbCrLf & _
           "Attached: " & attachedCount & vbCrLf & _
           "Already attached (skipped): " & skippedCount & vbCrLf & _
           "Cancelled or not attached: " & cancelledCount, _
           vbInformation, "Attach Title Block"

    Exit Sub

FatalError:
    MsgBox "The macro stopped on sheet model: " & _
           ActiveModelReference.Name & vbCrLf & vbCrLf & _
           "Error " & CStr(Err.Number) & ": " & Err.Description, _
           vbCritical, "Attach Title Block"
End Sub

Private Function TitleBlockDgnIsAttached( _
    ByVal sheetModel As ModelReference, _
    ByVal requestedPath As String) As Boolean

    On Error GoTo NotFound

    Dim attachment As Attachment
    Dim attachedPath As String
    Dim requestedFileName As String

    requestedFileName = FileNameFromPath(requestedPath)
    If Len(requestedFileName) = 0 Then Exit Function

    For Each attachment In sheetModel.Attachments
        attachedPath = ""

        'The property exposed for the stored reference path varies by ORD build.
        On Error Resume Next
        attachedPath = attachment.AttachName
        If Len(attachedPath) = 0 Then
            attachedPath = attachment.DesignFile.FullName
        End If
        Err.Clear
        On Error GoTo NotFound

        If StrComp(FileNameFromPath(attachedPath), _
                   requestedFileName, vbTextCompare) = 0 Then
            TitleBlockDgnIsAttached = True
            Exit Function
        End If
    Next attachment

NotFound:
    'False is the default return value.
End Function

Private Function FileNameFromPath(ByVal filePath As String) As String
    Dim normalizedPath As String
    Dim slashPosition As Long

    normalizedPath = Trim$(Replace(filePath, "/", "\"))

    Do While Len(normalizedPath) > 0 And _
             Right$(normalizedPath, 1) = "\"
        normalizedPath = Left$(normalizedPath, Len(normalizedPath) - 1)
    Loop

    slashPosition = InStrRev(normalizedPath, "\")

    If slashPosition > 0 Then
        FileNameFromPath = Mid$(normalizedPath, slashPosition + 1)
    Else
        FileNameFromPath = normalizedPath
    End If
End Function

Private Function QuoteText(ByVal value As String) As String
    QuoteText = Chr$(34) & value & Chr$(34)
End Function
