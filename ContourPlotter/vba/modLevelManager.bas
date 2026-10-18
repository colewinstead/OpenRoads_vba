Attribute VB_Name = "modLevelManager"
Option Explicit
' SPDX-License-Identifier: MIT
Public Sub CPFillLevels(ByVal box As MSForms.ComboBox, ByVal preference As String)
    Dim item As Level, previous As String, i As Long
    previous = GetSetting(CP_APP, "Levels", preference, "")
    box.Clear
    For Each item In ActiveDesignFile.Levels
        box.AddItem item.Name
    Next item
    For i = 0 To box.ListCount - 1
        If StrComp(box.List(i), previous, vbBinaryCompare) = 0 Then box.ListIndex = i
    Next i
End Sub

Public Function CPLevel(ByVal name As String) As Level
    Dim selected As Level
    CPRequire Len(name) > 0, "Select all required contour/label levels."
    Set selected = ActiveDesignFile.Levels.Find(name, Nothing)
    CPRequire Not selected Is Nothing, "Level unavailable in DGN or attached libraries: " & name
    CPRequire StrComp(selected.Name, name, vbBinaryCompare) = 0, "Level name must match exactly: " & name
    CPRequire Not selected.IsLocked And selected.ElementAccess = msdLevelElementAccessAll, "Level is not fully writable: " & name
    Set CPLevel = selected
End Function

Public Sub CPFillStyles(ByVal box As MSForms.ComboBox)
    Dim style As TextStyle, previous As String, i As Long
    previous = GetSetting(CP_APP, "Settings", "TextStyle", "")
    box.Clear
    For Each style In ActiveDesignFile.TextStyles
        If Len(style.Name) > 0 Then box.AddItem style.Name
    Next style
    For i = 0 To box.ListCount - 1
        If StrComp(box.List(i), previous, vbBinaryCompare) = 0 Then box.ListIndex = i
    Next i
End Sub

Public Function CPStyle(ByVal name As String) As TextStyle
    Dim style As TextStyle
    CPRequire Len(name) > 0, "Select an existing DGN/library text style."
    Set style = ActiveDesignFile.TextStyles.Find(name, Nothing)
    CPRequire Not style Is Nothing, "Text style unavailable: " & name
    CPRequire StrComp(style.Name, name, vbBinaryCompare) = 0, "Text style name must match exactly: " & name
    CPRequire style.Height > 0 And style.Width > 0, "Text style needs positive height and width."
    ' Cache the selected existing definition so element-local overrides are editable.
    If style.IsFromLibrary Then Set style = ActiveDesignFile.TextStyles.Add(style, name)
    Set CPStyle = style
End Function
