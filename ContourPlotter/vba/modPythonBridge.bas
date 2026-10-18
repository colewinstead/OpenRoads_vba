Attribute VB_Name = "modPythonBridge"
Option Explicit
' SPDX-License-Identifier: MIT
Private Type CPGuid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef value As CPGuid) As Long
Private Type CPOpenFileName
    size As Long
    owner As LongPtr
    instance As LongPtr
    filter As LongPtr
    customFilter As LongPtr
    maxCustomFilter As Long
    filterIndex As Long
    filename As LongPtr
    maxFile As Long
    fileTitle As LongPtr
    maxFileTitle As Long
    initialDirectory As LongPtr
    title As LongPtr
    flags As Long
    fileOffset As Integer
    extensionOffset As Integer
    defaultExtension As LongPtr
    customData As LongPtr
    hook As LongPtr
    templateName As LongPtr
    reserved As LongPtr
    reservedWord As Long
    extraFlags As Long
End Type
Private Type CPStartupInfo
    size As Long
    reserved As LongPtr
    desktop As LongPtr
    title As LongPtr
    x As Long
    y As Long
    width As Long
    height As Long
    xChars As Long
    yChars As Long
    fill As Long
    flags As Long
    showWindow As Integer
    reservedSize As Integer
    reservedBytes As LongPtr
    inputHandle As LongPtr
    outputHandle As LongPtr
    errorHandle As LongPtr
End Type
Private Type CPProcessInfo
    processHandle As LongPtr
    threadHandle As LongPtr
    processID As Long
    threadID As Long
End Type
Private Declare PtrSafe Function CPGetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (ByRef dialog As CPOpenFileName) As Long
Private Declare PtrSafe Function CPDialogError Lib "comdlg32.dll" Alias "CommDlgExtendedError" () As Long
Private Declare PtrSafe Function CPActiveWindow Lib "user32.dll" Alias "GetActiveWindow" () As LongPtr
Private Declare PtrSafe Function CPCreateProcess Lib "kernel32.dll" Alias "CreateProcessW" ( _
    ByVal applicationName As LongPtr, ByVal commandLine As LongPtr, ByVal processSecurity As LongPtr, _
    ByVal threadSecurity As LongPtr, ByVal inheritHandles As Long, ByVal flags As Long, _
    ByVal environment As LongPtr, ByVal directory As LongPtr, ByRef startup As CPStartupInfo, _
    ByRef process As CPProcessInfo) As Long
Private Declare PtrSafe Function CPWaitForProcess Lib "kernel32.dll" Alias "WaitForSingleObject" (ByVal handle As LongPtr, ByVal milliseconds As Long) As Long
Private Declare PtrSafe Function CPExitCode Lib "kernel32.dll" Alias "GetExitCodeProcess" (ByVal handle As LongPtr, ByRef code As Long) As Long
Private Declare PtrSafe Function CPCloseHandle Lib "kernel32.dll" Alias "CloseHandle" (ByVal handle As LongPtr) As Long

Public Function CPNewRunID() As String
    Dim value As CPGuid, i As Long, tail As String
    CPRequire CoCreateGuid(value) = 0, "Windows could not create a unique run ID."
    For i = 0 To 7
        tail = tail & Right$("00" & Hex$(value.Data4(i)), 2)
    Next i
    CPNewRunID = LCase$(Right$("00000000" & Hex$(value.Data1), 8) & "-" & _
        Right$("0000" & Hex$(value.Data2), 4) & "-" & Right$("0000" & Hex$(value.Data3), 4) & "-" & _
        Left$(tail, 4) & "-" & Mid$(tail, 5))
End Function

Public Function CPNewJob(ByVal runID As String) As String
    Dim fso As Object, root As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    root = Environ$("LOCALAPPDATA") & "\ContourPlotter"
    If Not fso.FolderExists(root) Then fso.CreateFolder root
    root = root & "\jobs"
    If Not fso.FolderExists(root) Then fso.CreateFolder root
    CPNewJob = root & "\" & runID
    CPRequire Not fso.FolderExists(CPNewJob), "Run folder already exists."
    fso.CreateFolder CPNewJob
End Function

Private Function CPQuote(ByVal value As String) As String
    CPRequire InStr(value, Chr$(34)) = 0 And InStr(value, vbCr) = 0 And InStr(value, vbLf) = 0, "Invalid executable or request path."
    CPQuote = Chr$(34) & value & Chr$(34)
End Function

Public Sub CPWriteUTF8(ByVal filename As String, ByVal value As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText value
    stream.SaveToFile filename, 2
    stream.Close
End Sub

Public Sub CPRunEngine(ByVal executable As String, ByVal job As String, ByVal request As String)
    Dim process As CPProcessInfo, startup As CPStartupInfo, fso As Object, message As String, f As Object
    Dim stopRequested As Boolean, commandLine As String, waitResult As Long, exitCode As Long
    On Error GoTo Failed
    executable = CPPath(executable)
    Set fso = CreateObject("Scripting.FileSystemObject")
    CPRequire fso.FileExists(executable) And LCase$(fso.GetExtensionName(executable)) = "exe", "Browse to the packaged ContourPlotter.exe."
    CPWriteUTF8 job & "\request.json", request
    startup.size = LenB(startup)
    CPRequire startup.size = 104 And LenB(process) = 24, "Unsupported Windows process structure layout."
    startup.flags = 1 ' STARTF_USESHOWWINDOW; showWindow defaults to SW_HIDE.
    commandLine = CPQuote(executable) & " --request " & CPQuote(job & "\request.json")
    ' CREATE_NO_WINDOW prevents a console from being created, even briefly.
    If CPCreateProcess(StrPtr(executable), StrPtr(commandLine), 0, 0, 0, &H8000000, 0, 0, startup, process) = 0 Then
        Err.Raise vbObjectError + 2812, "Contour engine", "Cannot start contour engine (Windows error " & CStr(Err.LastDllError) & ")."
    End If
    CPCloseHandle process.threadHandle: process.threadHandle = 0
    Do
        waitResult = CPWaitForProcess(process.processHandle, 75)
        If waitResult = 0 Then Exit Do
        CPRequire waitResult = 258, "Cannot monitor contour engine."
        DoEvents
        If CPCancelled And Not stopRequested Then
            Set f = fso.CreateTextFile(job & "\cancel", True)
            f.Close
            stopRequested = True
            CPStatus "Cancellation requested; waiting for the current library operation..."
        End If
    Loop
    CPRequire CPExitCode(process.processHandle, exitCode) <> 0, "Cannot read contour engine exit status."
    CPCloseHandle process.processHandle: process.processHandle = 0
    CPRequire Not CPCancelled, "Cancelled. No DGN elements have been inserted."
    If exitCode <> 0 Then
        If fso.FileExists(job & "\error.txt") Then message = CPReadUnicode(job & "\error.txt")
        If Len(message) = 0 Then message = "Contour engine exited with code " & CStr(exitCode) & ". See error.json in " & job
        Err.Raise vbObjectError + 2812, "Contour engine", message
    End If
    CPRequire fso.FileExists(job & "\result.json"), "Engine exited without a successful result."
    Exit Sub
Failed:
    message = Err.Description
    If process.threadHandle <> 0 Then CPCloseHandle process.threadHandle
    If process.processHandle <> 0 Then CPCloseHandle process.processHandle
    Err.Raise vbObjectError + 2812, "Contour engine", message
End Sub

Public Function CPRequest(ByVal action As String, ByVal runID As String, ByVal job As String, _
        ByVal source As String, ByVal spatial As String, ByVal settings As String) As String
    CPRequest = "{""version"":1,""action"":" & CPJson(action) & ",""run_id"":" & CPJson(runID) & _
        ",""input"":" & CPJson(CPPath(source)) & ",""output_dir"":" & CPJson(job) & _
        ",""spatial"":" & spatial & ",""settings"":" & settings & "}"
End Function

Public Function CPPath(ByVal value As String) As String
    value = Trim$(value)
    If Len(value) >= 2 And Left$(value, 1) = Chr$(34) And Right$(value, 1) = Chr$(34) Then
        value = Mid$(value, 2, Len(value) - 2)
    End If
    CPRequire InStr(value, Chr$(34)) = 0 And InStr(value, vbCr) = 0 And InStr(value, vbLf) = 0 And _
        InStr(value, vbNullChar) = 0 And InStr(value, vbTab) = 0, "Invalid file path: remove unmatched quotes or control characters."
    CPPath = value
End Function

Public Function CPBrowse(ByVal prompt As String, ByVal existing As String, _
        Optional ByVal fileFilter As String = "All files (*.*)|*.*") As String
    Dim dialog As CPOpenFileName, buffer As String, filter As String, folder As String
    Dim fso As Object, errorCode As Long
    On Error GoTo Failed
    existing = CPPath(existing)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(existing) Then
        folder = existing: existing = ""
    Else
        folder = fso.GetParentFolderName(existing)
        If Not fso.FolderExists(folder) Then folder = ""
        If Not fso.FileExists(existing) Then existing = ""
    End If
    CPRequire Len(existing) < 32768, "The existing file path exceeds the picker buffer."
    buffer = existing & String$(32768 - Len(existing), vbNullChar)
    filter = Replace(fileFilter, "|", vbNullChar) & vbNullChar & vbNullChar
    dialog.size = LenB(dialog)
    CPRequire dialog.size = 152, "Unsupported Windows file-picker structure layout."
    dialog.owner = CPActiveWindow()
    dialog.filter = StrPtr(filter): dialog.filterIndex = 1
    dialog.filename = StrPtr(buffer): dialog.maxFile = Len(buffer)
    dialog.title = StrPtr(prompt): dialog.initialDirectory = StrPtr(folder)
    dialog.flags = &H80000 Or &H1000 Or &H800 Or &H8 Or &H4
    If CPGetOpenFileName(dialog) <> 0 Then
        CPBrowse = CPPath(Left$(buffer, InStr(buffer, vbNullChar) - 1))
    Else
        errorCode = CPDialogError()
        CPRequire errorCode = 0, "Windows file picker failed (error 0x" & Hex$(errorCode) & ")."
    End If
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Contour Plotter"
End Function
