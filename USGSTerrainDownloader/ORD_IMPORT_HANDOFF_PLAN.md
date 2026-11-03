# Plan: open the generated LandXML in ORD's terrain import dialog

Status: proposed only. This document does not implement the handoff.

## Intended behavior

After a DEM job successfully produces `terrain.xml`, the MVBA opens ORD's
**Import Terrain Model(s)** dialog with that exact XML loaded. It stops before
**Import**. The user reviews the destination model, feature definition, units,
projection and append options, then chooses whether to import.

The automation must not create a terrain, append to an existing terrain, set an
active terrain, save a DGN, or click Import. Cancelling the dialog leaves the
exported XML and TIFF available for later use.

## Establish the supported ORD command first

The precise command and file-path argument have not been verified in this ORD
build. Do not implement this using a guessed key-in.

1. Use a disposable DGN and a small, already validated XML fixture.
2. Record the native **Terrain > Create > From File** operation and inspect the
   installed command resources/documentation to identify its supported key-in.
3. Determine whether it accepts an XML path and opens the review dialog without
   importing. Reject any command that immediately creates or modifies terrain.
4. If file selection is a separate step, test whether ORD exposes it through
   VBA's native modal-dialog handler interface. Verify the actual dialog type;
   do not assume a WPF/Civil dialog implements the MicroStation file-dialog API.
5. Prefer the verified command plus an explicitly scoped dialog handler. Avoid
   coordinate clicks, focus-dependent SendKeys, and automatic confirmation keys.
6. If VBA cannot supply the path reliably, investigate a small supported managed
   Civil UI adapter. Direct terrain-import SDK methods are not substitutes: they
   may build terrain instead of stopping at the review dialog. Do not add an
   adapter until the native VBA approach has been ruled out.

Success for this investigation means the expected XML is visible in the import
list, the Import button has not been pressed, and DGN element IDs are unchanged.
The screenshot's "Import Complete" state is not the desired stopping point;
the handoff should stop with the file loaded and awaiting import.

## Smallest implementation

### Completion hook

In `vba/frmUSGSTerrainDownloader.code.vba`, extend `RunWorker` after the worker
has exited successfully and normal busy-state cleanup has completed.

- Run only after DEM processing, never after Search, LiDAR download, failure or
  cancellation. Exit code zero alone is insufficient because LiDAR also succeeds.
- Use the exact current job folder, `mJob`, to locate `terrain.xml`. Never search
  for the newest XML in the destination or reuse a previous job's output.
- Require a nonempty, final `terrain.xml`; ignore partial XML files. Confirm the
  worker's completed DEM state using the existing status/report contract.
- Call a shared helper, provisionally `TerrainOpenImportReview(xmlPath)`, in
  `vba/modPythonBridge.bas`. The helper name is proposed, not an existing API.
- Hide the downloader only when handing focus to the native review dialog. Keep
  the output folder reachable and restore the downloader on handoff failure.
- Handle a handoff failure separately from a processing failure: report
  "Terrain files are ready; ORD import dialog could not be opened" with the XML
  path and reason. Do not mark a completed export failed or rerun its download.

### Model and file guards

- Recheck design-file/model identity, model units and GCS after processing using
  `TerrainCheckContext(True)` and the established session context.
- Verify the initial implementation in a suitable active 3D design model. If the
  job ran in 2D, retain the successful export and explain that the user must use
  an appropriate 3D model for import. Preserve the existing 2D export workflow.
- Do not create an arbitrary 3D model or switch models automatically. Associated
  Civil 2D/3D model creation is a separate, currently unverified integration task.
- Do not bypass stale-context checks when the user changes models. A future
  explicit handoff from another model would need fresh model/unit/GCS review.
- Quote the full XML path as data, using the bridge's command-quoting rules where
  applicable. Paths can contain spaces, apostrophes, ampersands and Unicode.
- Register any file-dialog handler immediately before launching the command and
  remove it on success, cancellation, timeout and error. Bound any wait. Never
  respond to an unrelated file dialog or leave a global handler installed.

### Review settings and stopping point

The desired **Build Terrain From** setting is **Definition**, preserving the
exported clipped faces and holes. Set it only through a verified native interface
that cannot initiate import. Otherwise leave the review dialog open with a clear
instruction for the user to select Definition.

[Bentley's import-option documentation](https://bentleysystems.service-now.com/community?id=kb_article&sysparm_article=KB0019731)
confirms that Definition uses stored faces while Source triangulates features.

The screenshot also shows **Append to existing Terrain Model** enabled. Do not
silently append to the previously selected terrain. Prefer opening the new-file
review with append disabled if the native interface supports it. If it retains
existing settings, explicitly tell the user to review that checkbox and target.
The user retains final control of the feature definition, projection, terrain
name, append choice, and Import button.

## Files expected to change during implementation

| File | Proposed change |
| --- | --- |
| `vba/frmUSGSTerrainDownloader.code.vba` | Successful DEM completion hook and separate handoff-error handling |
| `vba/modPythonBridge.bas` | Verified import-dialog launch and path validation helper |
| A dedicated VBA class, only if needed | Scoped native modal-dialog event handler |
| `tests/` native test module | Dialog handoff, cancellation, context and no-DGN-write assertions |
| `README.md`, `TEST_PLAN.md`, `API_NOTES.md` | Supported command, user workflow, verified limitations and evidence |
| `USGSTerrainDownloader.mvba` | Native compile/save of the matching source; no diagnostic modules shipped |

The Python exporter, geometry, request schema, run-folder naming and source cache
do not need to change. Before implementation, inspect the working tree and
establish the intended baseline; preserve unrelated edits and verify that the
expected worker files exist. Do not overwrite or restore files indiscriminately.

## Validation before delivery

Use a disposable DGN and separate test output directory. Record the ORD build and
command/interface actually used.

- Successful DEM completion opens the exact current XML in the review dialog.
- Search, LiDAR, failed/cancelled jobs, missing XML and partial XML do not launch it.
- Changed DGN/model/units/GCS and an unsuitable model produce an actionable message.
- Test paths containing spaces, apostrophes, ampersands and Unicode, plus file
  locks and a dialog already open for another source.
- Cancel from each file-selection/review stage; verify handler cleanup and no
  repeated launch, queued confirmation, or unexpected import on the next command.
- Compare DGN element IDs and relevant model/workspace settings before and after
  opening and cancelling the dialog. They must remain unchanged.
- Confirm the dialog is not in "Import Complete" state and that no terrain has
  been created or appended before the user clicks Import.
- Separately, let the tester manually import using Definition and check boundary,
  concavities, holes, units and sample elevations. This verifies the intended
  downstream workflow; it is not permission for the automation to press Import.
- Compile the final MVBA, remove test modules/handlers, save through Bentley's
  editor, and record the matching source and test evidence.

No production DGN edits, workspace/autoload changes, historical folder renames,
Git commit, or GitHub push are part of this documentation-only request.
