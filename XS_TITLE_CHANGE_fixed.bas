Attribute VB_Name = "ModXS_TITLE_CHANGE"
Option Explicit

Public Sub BmrXS_TITLE_CHANGE_AllSheets()

    Dim oldTitle As String
    Dim newTitle As String
    Dim targetNames As Collection
    Dim m As ModelReference
    Dim nm As Variant
    Dim attemptedCount As Long

    oldTitle = InputBox( _
        "Enter the existing title text to replace.", _
        "Existing Cross Section Title Text", _
        "83.6")

    If StrPtr(oldTitle) = 0 Then
        MsgBox "Cancelled.", vbExclamation
        Exit Sub
    End If

    oldTitle = Trim$(oldTitle)
    If Len(oldTitle) = 0 Then
        MsgBox "Cancelled. No existing text entered.", vbExclamation
        Exit Sub
    End If

    newTitle = InputBox( _
        "Enter the new title text to place on every Sheet model.", _
        "Cross Section Title Text", _
        "SR 8 SOUTH ENTRY")

    If StrPtr(newTitle) = 0 Then
        MsgBox "Cancelled.", vbExclamation
        Exit Sub
    End If

    newTitle = Trim$(newTitle)
    If Len(newTitle) = 0 Then
        MsgBox "Cancelled. No replacement text entered.", vbExclamation
        Exit Sub
    End If

    Set targetNames = New Collection

    ' Build the model list first, then activate one-by-one.
    For Each m In ActiveDesignFile.Models
        If m.Type = msdModelTypeSheet Then
            targetNames.Add m.Name
        End If
    Next m

    If targetNames.Count = 0 Then
        MsgBox "No Sheet models were found in the active DGN.", vbExclamation
        Exit Sub
    End If

    For Each nm In targetNames
        Set m = ActiveDesignFile.Models.Item(CStr(nm))
        m.Activate

        BmrXS_TITLE_CHANGE_Recorded oldTitle, newTitle
        attemptedCount = attemptedCount + 1
    Next nm

    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"

    MsgBox "Finished attempting title change on Sheet models." & vbCrLf & vbCrLf & _
           "Existing text: " & oldTitle & vbCrLf & _
           "Replacement text: " & newTitle & vbCrLf & _
           "Sheets processed: " & attemptedCount, _
           vbInformation, "Title Change Complete"

End Sub

Private Sub BmrXS_TITLE_CHANGE_Recorded(ByVal originalText As String, ByVal replacementText As String)

    Dim startPoint As Point3d
    Dim point As Point3d
    Dim i As Long

    startPoint.X = 2.73368890464626
    startPoint.Y = 0.417978090237162
    startPoint.Z = 0#

    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendKeyin "TEXTEDITOR MODIFY"

    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND CLEAR_ANCHOR_CARET"
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND SET_INSERT_CARET LINE 0 CHARACTER " & CStr(Len(originalText))
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND SET_ANCHOR_CARET LINE 0 CHARACTER " & CStr(Len(originalText))

    For i = Len(originalText) - 1 To 0 Step -1
        CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND SET_INSERT_CARET LINE 0 CHARACTER " & CStr(i)
    Next i

    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""s"""
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x02 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    CadInputQueue.SendKeyin _
        "TEXTEDITOR PLAYCOMMAND INSERT_TEXT """ & EscapeForKeyin(replacementText) & """"

    point.X = startPoint.X + 0.21825281696686
    point.Y = startPoint.Y + 0.334393919770742
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X + 0.222831547392739
    point.Y = startPoint.Y + 0.375620567413711
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CommandState.StartDefaultCommand

End Sub

Private Function EscapeForKeyin(ByVal s As String) As String
    EscapeForKeyin = Replace(s, """", """""")
End Function
