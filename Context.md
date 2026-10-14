# OpenRoads / MicroStation — Project Master Instructions
_Last updated: 2026-03-19_

## How to use this document
- This is the **single source of truth** for “what works” on this project.
- Add new workflows as **Task Runbooks** (Section 3). Keep them **copy/paste runnable**.
- Record anything that *didn’t* work in that task’s **Change log** so you don’t re‑test it later.
- When ORD builds differ, prefer **KEYIN-based** steps/macros and note the build that failed.

## Table of contents
1. Project info (fill in)
2. Global conventions (project-wide)
3. Task runbooks
4. Go-bys & examples (project-wide)
5. Project-wide change log

---

## 1) Project info (fill in)
- **Project name:** `____________________________`
- **Client / Agency:** `____________________________`
- **ORD version / build:** `____________________________`
- **Workspace / standards:** `____________________________`
- **Coordinate system / units:** `____________________________`
- **PW working directory:** `____________________________`
- **Folder conventions:** `____________________________`
- **Key contacts:** `____________________________`

---

## 2) Global conventions (project-wide)

### 2.1 Model naming (general)
- **Model naming convention:** `____________________________`
- **Station format in model names:** `####+##.####` (example: `434+50.0000`)
- **Sheet model handling:** **Do not** use Sheet models unless a task explicitly says so.
  - Sheet models often show `[Sheet]` in the name — treat that as the “skip” flag for automation.

### 2.2 Station token standard (used by automation)
Many workflows need to match by station even when prefixes differ, e.g.
- `SR42_SITE3 -434+50.0000`
- `ALL_SITE3_SR42 -434+50.0000`

**Standard station token regex (including optional negative):**
```regex
(-?\d+\+\d+(?:\.\d+)?)
```

**Standard conversion to a numeric station value:**
- Parse `major+minor` (and optional sign)
- Convert to `sign * (major*100 + minor)`
  - This assumes `+##` is feet within the 100‑ft station.
  - If your project uses a different convention, update this rule once here.

**Comparison tolerance:** `0.0001` (after conversion)

### 2.3 References & logical names
- Always attach references with a **stable Logical Name** so you can:
  - detect duplicates in automation
  - target/replace specific references later
- Logical name rules (safe across builds):
  - no spaces
  - avoid special characters
  - keep ≤ 40 chars

### 2.4 VBA / Macro conventions (OpenRoads-safe)
OpenRoads VBA is not “full Office VBA.” To keep macros portable across ORD builds:
- Avoid Office-only objects (e.g., `FileDialog`) and methods that aren’t exposed in ORD builds.
- Prefer **late binding**:
  - `CreateObject("Scripting.Dictionary")` (no reference needed)
  - `CreateObject("VBScript.RegExp")` for regex
  - `CreateObject("UserAccounts.CommonDialog")` for a file picker *if available*, otherwise use `InputBox`
- If a VBA attachment API fails on your ORD build, fall back to the **KEYIN** approach:
  - `REFERENCE ATTACH "path","model","logicalName"`
- When looping models:
  - build a list of model names first
  - then activate models one-by-one (safer than activating during enumeration)

---

## 3) Task runbooks

### Task template (copy/paste for new tasks)
#### Task: <TITLE>
**Purpose:**  
**When to use:**  
**Preconditions:**  
**Inputs:**  
**Output:**  

**Known-good steps:**  
1.  
2.  
3.  

**Working commands / keyins / VBA (if applicable):**
```text
<commands here>
```

**Notes / gotchas:**  
-  

**Troubleshooting checklist:**  
- [ ]  
- [ ]  

**Go-bys / examples:**  
-  

**Task change log:**  
- YYYY-MM-DD:  

---

### Task: Attach XS reference models into main XS models by station (Earthwork block workaround)

#### Purpose
Attach (reference) the correct XS model from a **reference DGN** into each XS model in the **main DGN**, matching models by **station token** (not full model name). **Skip all Sheet models.**

#### When to use
- Prefixes differ between main/ref model names, so exact-name matching fails.
- XS annotation/labels aren’t behaving and you need a consistent earthwork block workflow.
- You want a repeatable automation that works across ORD builds.

#### Preconditions
- You can open both DGNs (ProjectWise or local working directory).
- Reference DGN contains a matching XS model for each station you need.

#### Inputs
- **Main DGN:** current active file (the one you’re in)
- **Reference DGN:** browse/paste path (contains the “source” XS models)
- **Logical name prefix:** `XS_` (change if needed)

#### Output
For each **non-sheet** XS model in the main DGN:
- One matching XS model from the reference DGN is attached using `REFERENCE ATTACH`
- A stable logical name is assigned so duplicates can be detected

#### What made it work (key discoveries)
- **Match by station only** (extract token via regex, compare numeric station values)
- **Skip sheet models** using `[Sheet]` in the model name
- **Use KEYIN** for attaching (ORD build differences broke `Attachments.Add` / similar APIs)
- **Loop all models safely** by collecting names first, then activating one at a time

#### Known-good VBA (OpenRoads-safe, no Office references)
> Paste into a new VBA module in OpenRoads / MicroStation VBA, then run:
> `XS_AttachRefModels_ByStation_AllMainModels`

```vb
Option Explicit

'=============================
' MAIN ENTRY POINT
'=============================
Public Sub XS_AttachRefModels_ByStation_AllMainModels()

    Dim refDgnPath As String
    refDgnPath = GetRefDgnPath()
    If Len(Trim$(refDgnPath)) = 0 Then Exit Sub

    Dim logicalPrefix As String
    logicalPrefix = "XS_"

    'Open reference DGN read-only
    Dim refDgn As DesignFile
    Set refDgn = Application.OpenDesignFile(refDgnPath, True)

    'Build station -> model index for the reference DGN
    Dim refIndex As Object
    Set refIndex = CreateObject("Scripting.Dictionary") 'late-bound
    BuildRefStationIndex refDgn, refIndex

    'Collect main model names first (safer when activating)
    Dim mainNames As Collection
    Set mainNames = New Collection

    Dim mm As Model
    For Each mm In ActiveDesignFile.Models
        mainNames.Add mm.Name
    Next

    'Stats
    Dim attachedCount As Long, skippedSheet As Long, noStation As Long
    Dim noMatch As Long, already As Long, failed As Long

    Dim nm As Variant
    For Each nm In mainNames

        Dim m As Model
        Set m = ActiveDesignFile.Models.Item(CStr(nm))

        If IsSheetModelName(m.Name) Then
            skippedSheet = skippedSheet + 1
        Else
            m.Activate

            Dim stTok As String
            stTok = ExtractStationToken(m.Name)

            If Len(stTok) = 0 Then
                noStation = noStation + 1
            Else
                Dim stVal As Double
                stVal = StationTokenToValue(stTok)

                Dim key As String
                key = StationKey(stVal)

                If Not refIndex.Exists(key) Then
                    noMatch = noMatch + 1
                Else
                    Dim refModelName As String
                    refModelName = CStr(refIndex(key))

                    Dim logicalName As String
                    logicalName = BuildLogicalName(logicalPrefix, stTok)

                    If AttachmentLogicalNameExists(logicalName) Then
                        already = already + 1
                    Else
                        If AttachModelByKeyin(refDgnPath, refModelName, logicalName) Then
                            attachedCount = attachedCount + 1
                        Else
                            failed = failed + 1
                        End If
                    End If
                End If
            End If
        End If

    Next nm

    On Error Resume Next
    refDgn.Close
    On Error GoTo 0

    MsgBox "XS Attach by Station complete." & vbCrLf & _
           "Attached: " & attachedCount & vbCrLf & _
           "Skipped [Sheet]: " & skippedSheet & vbCrLf & _
           "No station token: " & noStation & vbCrLf & _
           "No match in ref: " & noMatch & vbCrLf & _
           "Already attached: " & already & vbCrLf & _
           "Failed: " & failed, vbInformation, "OpenRoads"

End Sub

'=============================
' FILE PATH PICKER (ORD-safe)
'=============================
Private Function GetRefDgnPath() As String
    Dim p As String
    p = TryPickFileCommonDialog("Pick reference XS DGN", "DGN Files|*.dgn|All Files|*.*")
    If Len(Trim$(p)) = 0 Then
        p = InputBox("Paste full path to the reference XS DGN:", "Reference DGN Path")
    End If
    GetRefDgnPath = Trim$(p)
End Function

'Late-bound Windows CommonDialog. If not available on your machine, it just returns "" and falls back to InputBox.
Private Function TryPickFileCommonDialog(ByVal title As String, ByVal filter As String) As String
    On Error GoTo EH
    Dim dlg As Object
    Set dlg = CreateObject("UserAccounts.CommonDialog")
    dlg.Filter = filter
    dlg.FilterIndex = 1
    dlg.InitialDir = CurDir$
    dlg.DialogTitle = title

    If dlg.ShowOpen Then
        TryPickFileCommonDialog = dlg.FileName
    Else
        TryPickFileCommonDialog = ""
    End If
    Exit Function
EH:
    TryPickFileCommonDialog = ""
End Function

'=============================
' INDEX REFERENCE MODELS
'=============================
Private Sub BuildRefStationIndex(ByVal refDgn As DesignFile, ByVal dict As Object)
    Dim m As Model
    For Each m In refDgn.Models
        If Not IsSheetModelName(m.Name) Then
            Dim tok As String
            tok = ExtractStationToken(m.Name)
            If Len(tok) > 0 Then
                Dim v As Double
                v = StationTokenToValue(tok)

                Dim k As String
                k = StationKey(v)

                'Keep first hit (change to overwrite if needed)
                If Not dict.Exists(k) Then dict.Add k, m.Name
            End If
        End If
    Next m
End Sub

'=============================
' STATION PARSING
'=============================
Private Function ExtractStationToken(ByVal modelName As String) As String
    On Error GoTo EH
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp") 'late-bound
    re.Pattern = "(-?\d+\+\d+(?:\.\d+)?)"
    re.Global = False
    re.IgnoreCase = True

    If re.Test(modelName) Then
        ExtractStationToken = re.Execute(modelName)(0).SubMatches(0)
    Else
        ExtractStationToken = ""
    End If
    Exit Function
EH:
    ExtractStationToken = ""
End Function

'Converts "-434+50.0000" => -43450.0000 (major*100 + minor)
Private Function StationTokenToValue(ByVal stTok As String) As Double
    Dim s As String: s = Trim$(stTok)

    Dim sign As Double: sign = 1#
    If Left$(s, 1) = "-" Then
        sign = -1#
        s = Mid$(s, 2)
    ElseIf Left$(s, 1) = "+" Then
        s = Mid$(s, 2)
    End If

    Dim parts() As String
    parts = Split(s, "+")

    If UBound(parts) < 1 Then
        StationTokenToValue = 0#
        Exit Function
    End If

    Dim major As Double: major = CDbl(parts(0))
    Dim minor As Double: minor = CDbl(parts(1))

    StationTokenToValue = sign * (major * 100# + minor)
End Function

Private Function StationKey(ByVal stVal As Double) As String
    StationKey = Format$(Round(stVal, 4), "0.0000")
End Function

'=============================
' SHEET SKIP + LOGICAL NAME
'=============================
Private Function IsSheetModelName(ByVal modelName As String) As Boolean
    IsSheetModelName = (InStr(1, modelName, "[Sheet]", vbTextCompare) > 0)
End Function

Private Function BuildLogicalName(ByVal prefix As String, ByVal stTok As String) As String
    Dim t As String: t = stTok
    t = Replace(t, " ", "")
    t = Replace(t, "+", "_")
    t = Replace(t, ".", "_")
    t = Replace(t, "-", "NEG_")
    BuildLogicalName = Left$(prefix & t, 40)
End Function

'=============================
' DUPLICATE CHECK (best effort)
'=============================
Private Function AttachmentLogicalNameExists(ByVal logicalName As String) As Boolean
    On Error GoTo EH
    Dim a As Attachment
    For Each a In ActiveModelReference.Attachments
        If StrComp(a.LogicalName, logicalName, vbTextCompare) = 0 Then
            AttachmentLogicalNameExists = True
            Exit Function
        End If
    Next a
    AttachmentLogicalNameExists = False
    Exit Function
EH:
    'If Attachments collection isn't exposed on your build, we can't check reliably.
    AttachmentLogicalNameExists = False
End Function

'=============================
' ATTACH VIA KEYIN (known-good)
'=============================
Private Function AttachModelByKeyin(ByVal refPath As String, ByVal refModel As String, ByVal logicalName As String) As Boolean
    On Error GoTo EH
    Dim k As String
    k = "REFERENCE ATTACH " & Quote(refPath) & "," & Quote(refModel) & "," & Quote(logicalName)
    CadInputQueue.SendKeyin k
    CadInputQueue.SendKeyin "NULL"
    AttachModelByKeyin = True
    Exit Function
EH:
    AttachModelByKeyin = False
End Function

Private Function Quote(ByVal s As String) As String
    Quote = Chr$(34) & s & Chr$(34)
End Function
```

#### Notes / gotchas
- If `pw:\...` paths give attach errors, try a **local working path** (e.g., `C:\pw_working\...`) for the reference DGN.
- If you have duplicate station models in the reference DGN, the current code keeps the **first** one it sees.
- Station conversion assumes `major*100 + minor`. If your stationing isn’t 100‑ft stations, change `StationTokenToValue`.

#### Troubleshooting checklist
- [ ] Are you running it from the **main** XS DGN (active file)?
- [ ] Do your model names actually contain a station token that matches the regex?
- [ ] Does the reference DGN contain the stations you need (non-sheet)?
- [ ] If nothing attaches, try using a local path instead of a PW path.

#### Go-bys / examples
- Main: `____________________________`
- Ref: `____________________________`

#### Task change log
- 2026-01-21: Added the complete known-good macro + project-wide conventions used to make it work.

---


---

### Task: Automation pattern that works in ORD (Record → Wrap → Batch)

#### Purpose
Repeat a UI-heavy action that works once (especially **Fields**, dialog tools, deletes) across:
- many **models** (Sheet vs Design), and/or
- many **DGNs** (multiple files)

#### When to use
- The task **works manually** (or as a recorded macro) but ORD VBA APIs are inconsistent across builds.
- Find/Replace doesn’t work because the displayed text differs per sheet (Fields).
- You must be precise about *where* to run (Sheet models vs Default/Design models).

#### Core pattern (known-good)
1. **Record** a VBA macro that does the action correctly in ONE model.
2. **Wrap** it in a loop that activates only the right model type:
   - `msdModelTypeSheet` for sheets
   - `msdModelTypeDesign` for design/default models
3. If you need to run across **many DGNs**, use **Batch Process** to open each file and run the wrapper.

#### Wrapper templates

**Run recorded Sub on all Sheet models (current DGN):**
```vb
Option Explicit

Sub <RecordedSub>_AllSheets()
    Dim m As ModelReference
    For Each m In ActiveDesignFile.Models
        If m.Type = msdModelTypeSheet Then
            m.Activate
            <RecordedSub>
        End If
    Next m
    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"
End Sub
```

**Run recorded Sub on all Design/Default models (current DGN):**
```vb
Option Explicit

Sub <RecordedSub>_AllDesignModels()
    Dim m As ModelReference
    For Each m In ActiveDesignFile.Models
        If m.Type = msdModelTypeDesign Then
            m.Activate
            <RecordedSub>
        End If
    Next m
    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"
End Sub
```

#### Batch Process command file (run wrapper across many DGNs)
Create a `.txt` command file like:
```text
vba load <YourMacros>.mvba
vba run <RecordedSub>_AllDesignModels
filedesign savesettings
```

#### Common gotchas (and fixes)
- **Expected variable or procedure, not module**  
  You called a module name instead of the recorded `Sub`. Call the actual recorded procedure name (example: `BmrCWsheetnopg`).
- **Invalid inside procedure**  
  You put `Option Explicit` (or a new `Sub`) *inside* another `Sub`. Move `Option Explicit` to the top of the module and put wrappers after an `End Sub`.
- Model names shown with `[Sheet]` in the UI: **don’t** include `[Sheet]` in your `Like` pattern. It’s not part of `Model.Name`.

#### Notes / gotchas
- For recorded `TEXTEDITOR` macros, the wrapper does **not** create a true find/replace. The recorded `Sub` is still **position-based** unless you explicitly code otherwise.
- If the macro uses `SET_INSERT_CARET LINE <n> CHARACTER <c>`, changing text length or starting at the wrong character can leave trailing text behind.
- Example failure seen on this project: replacing `SR 8 WEST ENTRY` with `SR 8 WEST EXIT` produced `SR 8 WEST EXITENTRY`.
- For full-line replacement, put the caret at `TITLE_START_CHAR + Len(oldText)`, then backspace `Len(oldText)` and insert the new string.
- When looping Sheet models in ORD, build the list of model names first, then activate them one-by-one before calling the recorded `Sub`.

#### Troubleshooting checklist
- [ ] Did the macro leave part of the old text behind, such as `EXITENTRY`?
- [ ] Is the title start character set correctly (`0` if the title begins at the first character of the line)?
- [ ] Did the old text input include an accidental leading or trailing space?
- [ ] Is the title actually on `LINE 0`, or does the recorded macro need the line number adjusted?
- [ ] Is the recorded tentative/data point opening the correct text element on every sheet?
- [ ] Test on one sheet first before running the all-sheets wrapper.

#### Task change log
- 2026-03-19: Recorded `TEXTEDITOR` title-block replacement left trailing text when a fixed caret position was used. Added end-of-old-string backspace guidance and troubleshooting notes.
- 2026-01-21: Added after successful use on Field edits ($PG$), model loops, and multi-file batch runs.

---

### Task: Attach Title Block reference to ONLY Sheet models (VBA; no Batch Process filtering)

#### Purpose
Attach `PW_WORKDIR:d0300586\Title Block.dgn` to **only** the Sheet models in the current DGN.

#### What made it work
- Use `m.Type = msdModelTypeSheet` to detect sheets (avoid name-based filtering).
- Do **not** include `[Sheet]` in model-name patterns (UI only).
- If you want only SR 32 sheets, filter by `Like "SR 32 -*"` (no `[Sheet]` suffix).

#### Known-good VBA
```vb
Option Explicit

Sub AttachTitleBlockToSheetModels()

    Const REF_PATH As String = "PW_WORKDIR:d0300586\Title Block.dgn"
    Const REF_LOGICAL As String = "TITLEBLOCK"
    Const NAME_FILTER As String = "SR 32 -*"   'Change to "*" for all sheet models

    Dim m As ModelReference
    Dim a As Attachment
    Dim already As Boolean

    For Each m In ActiveDesignFile.Models

        If m.Type = msdModelTypeSheet Then

            If m.Name Like NAME_FILTER Then

                already = False
                For Each a In m.Attachments
                    If UCase$(a.LogicalName) = REF_LOGICAL Then
                        already = True
                        Exit For
                    End If
                Next a

                If Not already Then
                    'Attach coincident to THIS sheet model
                    Call m.Attachments.AddCoincident(REF_PATH, "", REF_LOGICAL, "MDOT Title Block", True)
                End If

            End If
        End If

    Next m

    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"
    MsgBox "Title Block attached to qualifying sheet models.", vbInformation

End Sub
```

#### Notes / gotchas
- If nothing attaches, verify the **ref path resolves on this machine** (try a local working path if needed).
- If your title block lives in a specific model inside the ref, you may need a different attach method; this coincident attach works for typical MDOT title blocks.

#### Task change log
- 2026-01-21: Confirmed working in production (sheet-only attach).

---

### Task: Replace a per-sheet Field with literal `$PG$` on every sheet

#### Purpose
Convert a text **Field** (value differs per sheet) into the literal text `$PG$` on every Sheet model.

#### Known-good approach
- Record the manual edit once (click field → delete field content → type `$PG$` → accept).
- Wrap that recorded Sub in an “all sheets” loop.

#### Example (real project names)
Recorded Sub (example): `BmrCWsheetnopg`  
Wrapper to run on all sheets:
```vb
Option Explicit

Sub BmrCWsheetnopg_AllSheets()

    Dim m As ModelReference

    For Each m In ActiveDesignFile.Models
        If m.Type = msdModelTypeSheet Then
            m.Activate
            BmrCWsheetnopg
        End If
    Next m

    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"
    MsgBox "Replaced Field with $PG$ on all sheet models.", vbInformation

End Sub
```

#### Notes / gotchas
- The recorded Sub name is what you call (e.g., `BmrCWsheetnopg`), not the module name.
- Keep `Option Explicit` at the top of the module only.

#### Task change log
- 2026-01-21: Confirmed working using “record + wrapper” method.

---

### Task: Run the same recorded macro in Default/Design models across MANY files

#### Purpose
Apply a recorded action across multiple DGNs, running only in **Design/Default** models (not sheets).

#### Steps
1. Ensure you have a wrapper that runs the recorded action on design models:
```vb
Option Explicit

Sub <RecordedSub>_AllDesignModels()
    Dim m As ModelReference
    For Each m In ActiveDesignFile.Models
        If m.Type = msdModelTypeDesign Then
            m.Activate
            <RecordedSub>
        End If
    Next m
    CadInputQueue.SendCommand "FILEDESIGN SAVESETTINGS"
End Sub
```

2. Use Batch Process command file:
```text
vba load <YourMacros>.mvba
vba run <RecordedSub>_AllDesignModels
filedesign savesettings
```

#### Task change log
- 2026-01-21: Confirmed working to run recorded actions across multiple DGNs in Default/Design models.


### Task: Prevent clearzone/end-condition spike when EB-L must remain point-controlled to existing toe of ditch

#### Purpose
Prevent template spikes/fold-backs when the final left-side end-condition point (`EB-L`) must stay under **Point Control** to an existing toe-of-ditch feature, but the controlled toe elevation rises above the clearzone point.

#### When to use
- `EB-L` must remain the point-controlled point.
- The left-side end condition includes a clearzone branch.
- The corridor spikes or folds back when the controlled existing toe of ditch is **higher** than the clearzone point.
- A single end-condition branch is trying to solve both the “toe below clearzone” and “toe above clearzone” cases.

#### Preconditions
- ORD / OpenRoads Designer corridor with editable template or template drop.
- `EB-L` is already assigned as the controlled point in **Point Control**.
- You know the **clearzone point** name used for the comparison.
- You can duplicate the affected left-side branch in the template.

#### Inputs
- Controlled point: `EB-L`
- Comparison point: project-specific **clearzone point**
- Existing toe-of-ditch feature used by Point Control
- Threshold value: `0` vertical difference unless a tolerance is needed

#### Output
A stable template that:
- keeps `EB-L` under Point Control,
- shows the **normal clearzone branch** when `EB-L` is at or below the clearzone point,
- shows an **alternate high-toe branch** when `EB-L` is above the clearzone point,
- avoids the spike/fold-back condition.

#### Known-good steps
1. Leave **Point Control** on `EB-L`. Do **not** move the control to another point if project requirements say `EB-L` must stay controlled.
2. In **Create Template**, duplicate the left-side branch that runs from the clearzone side out to `EB-L`.
3. Keep both branches terminating at the same controlled point: `EB-L`.
4. Make **Branch A** the normal clearzone solution.
5. Add a **Display Rule** to Branch A so it only displays when `EB-L` is at or below the clearzone point.
   - Use a **Vertical** comparison.
   - Compare: `EB-L` to `<Clearzone Point>`
   - Expression: `<=`
   - Value: `0`
6. Make **Branch B** the alternate “high toe” solution that can rise to `EB-L` without folding back.
   - Typical fix: simplify the outer branch or tie more directly from the inboard hinge to `EB-L`.
7. Add the opposite **Display Rule** to Branch B.
   - Use a **Vertical** comparison.
   - Compare: `EB-L` to `<Clearzone Point>`
   - Expression: `>`
   - Value: `0`
8. Set correct **Parent Components** so hidden branches do not leave child geometry behind.
9. Test the template or template drop with both conditions:
   - `EB-L` below clearzone point → only Branch A should display.
   - `EB-L` above clearzone point → only Branch B should display.
10. If needed, add a very small tolerance so the switch is clean at the exact crossover elevation.

**Working commands / keyins / VBA (if applicable):**
```text
Corridors > Edit > Edits split button > Point Control
Create Template > double-click component > Display Rules
Create Template > double-click component > Parent Component
```

**Notes / gotchas:**
- This is a **two-branch solution**, not a one-branch fix.
- Keep `EB-L` point-controlled if the project requires it; solve the issue by switching which branch displays.
- The display-rule comparison is based on the vertical difference `EB-L.y - ClearzonePoint.y`.
- If the switch occurs right at equal elevation and flickers, use a small tolerance such as `<= 0.01` for one branch and `> 0.01` for the other.
- If this area is inside a **template transition**, verify behavior carefully; display-rule-driven switching can require extra review at transition limits.
- If only a short station range is bad, an **End Condition Exception** may still be useful, but the known-good fix for this case is the two-branch display-rule setup.

**Troubleshooting checklist:**
- [ ] Does `EB-L` remain the controlled point in Corridor Point Controls?
- [ ] Do both branches terminate at the same `EB-L`?
- [ ] Is the normal branch rule `EB-L <= Clearzone Point`?
- [ ] Is the alternate branch rule `EB-L > Clearzone Point`?
- [ ] Are Parent Components set so hidden child pieces also turn off?
- [ ] If both branches show or neither shows, did the vertical comparison point order get reversed?
- [ ] If the switch is unstable, did you add a small tolerance?

**Go-bys / examples:**
- Branch A (normal): show when `EB-L.y - ClearzonePoint.y <= 0`
- Branch B (high toe): show when `EB-L.y - ClearzonePoint.y > 0`
- Confirmed project result: this eliminated the spike while keeping `EB-L` point-controlled.

**Task change log:**
- 2026-03-11: Confirmed known-good fix for spike caused by point-controlled `EB-L` rising above the clearzone point. Added two-branch display-rule workflow that keeps `EB-L` under Point Control.

---

## 4) Go-bys & examples (project-wide)
- Training reference PDF in this project: `RWD_Workflow_Training_ORD.pdf` (use for MDOT/OpenRoads standards & baseline workflows)

---

## 5) Project-wide change log
- 2026-03-19: Added troubleshooting guidance for recorded `TEXTEDITOR` title-block macros. Fixed-caret replacements can leave trailing text; use end-of-old-string backspace logic and test on one sheet first.
- 2026-03-11: Added runbook for preventing clearzone/end-condition spikes when `EB-L` must remain point-controlled to an existing toe of ditch; confirmed two-branch Display Rule workflow.
- 2026-01-21: Added automation runbooks for **Record → Wrap → Batch** (sheet vs design; multi-file) + Title Block attach + `$PG$` field replacement.
- 2026-01-21: Fixed malformed `Quote()` helper in XS attach VBA snippet.
- 2026-01-21: Updated this master doc with OpenRoads-safe automation conventions and the known-good XS attach-by-station runbook.
