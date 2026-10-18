# ORD acceptance plan

**Pending:** VBA compile, form operation, all DGN writes, save/reopen, regeneration,
rollback, cancellation during insertion and PDF plotting. Only the API read/unsaved
preparation checks in API_NOTES.md have been run against a live ORD session.

Use a copy of a seed/test DGN with existing approved MDOT levels. Never run first
write acceptance against the user's open production DGN. Keep unrelated test lines
and text on the same selected contour levels to prove deletion is by ID/tag.

1. **Assembly and UI.** Import six modules and create the blank named form as
   in README. Debug > Compile, save the real MVBA and record ORD build. Exercise
   all tabs, keyboard/tab order, large display scaling, browse fallback, Close/X,
   source inventory, multiple/duplicate surface names and stale source selections.
2. **Context.** Record file/model, GCS including local modifiers, master unit
   numerator/denominator, UOR/master, source axes and vertical datum. Confirm a
   known source control point against the DGN. Test model/file switching during
   Python processing and batch insertion. Reject read-only, locked, sheet/drawing
   and reference targets without writes. Test 2D and 3D targets separately.
3. **Levels/styles.** Enumerate actual DGN and configured-DGNLIB names, including
   unused levels/styles. Missing preferences must be blank; no replacement names
   may appear. Place onto unused library levels and measure a library text style;
   preserve the original shared definitions. Lock/restrict a level or remove a
   style and check clear failure before graphics are written.
   Capture all shared level/style definitions and active settings before/after.
4. **Native line strings.** Run a known triangle and a square two-triangle plane.
   Check endpoints, connected vertices, elevations and coordinates independently.
   Use a closed pyramid and topological hole; verify closure/voids and no bridging.
   Check 1,024-point boundaries and repeated joins for long contours. Counts must
   match geometry.cp1 and elements.cpj.
5. **Symbology.** Verify actual stored ByLevel color/weight/style after insertion
   and Rewrite, and again after reopening. Ensure elements are primary-class,
   ungrouped and printable. Check optional explicit overrides and all active
   symbology combinations; do not change shared definitions.
6. **Text.** Select each relevant MDOT text style. At 20, 50 and 100 ft/in and
   0.10-inch paper height, measure physical model text heights 2, 5 and 10 ft.
   Check both 2D and 3D rotations/origins, centered justification, long and decimal
   elevations, upright orientation, collisions and gaps. Validate actual glyph
   outlines against reserved width/height. Repeat with different active annotation
   scales and annotation-scale propagation after generation. Shared style values
   must stay unchanged. If scale propagation changes output, regenerate/recheck
   rather than claiming static annotation independence.
7. **Preflight/integrity.** Corrupt a test CP1 header, run ID, unit, count, vertex,
   Z, label value, truncation or footer. Every case must fail before adding
   anything. Try modifying the file between passes: the deny-write lock must hold.
   Modify the XML after analysis and verify hash rejection.
8. **Regeneration.** Save/reopen, regenerate, decline old deletion (both survive),
   then approve exact old-member removal. Unrelated same-level graphics survive.
   Copy/move members within/across models; copied new IDs must remain outside the
   original journal. Edit original text/vertices/symbology and verify explicit
   manual-edit confirmation before any old removal. Strip a tag or use another
   DGN/model journal: refuse removal. Delete an old member and check host-specific
   absent-ID behavior. Record whether GetElementByID returns Nothing or raises.
9. **Failure/cancellation.** Cancel before worker completion, before insertion,
   during each insertion batch and during old cleanup. Force a write/Rewrite
   failure using a disposable target/level. Check that only new handles are rolled
   back and old successful output is retained. Interrupt ORD on a test copy,
   preserve partial journals, run ListTaggedRuns, and manually reconcile unflushed
   members. Do not advertise atomic transactions or one-step Undo.
10. **Real terrain and plotting.** Use the local private export, increasing synthetic
    TINs and a larger approved real terrain. Record worker timings, validation and
    AddElement/Rewrite times, element counts, DGN bytes before/after, save/reopen
    times, pan/zoom and PDF print times. Print actual MDOT pen tables at the
    intended scales. Inspect steep slopes, channels, tie-ins, tiny loops, line
    weights and readable physical gaps. Compare against manually checked contours.
    No production-ready performance claim until these gates pass.

Record each case with input source hash (kept private), settings, expected/measured
result, screenshots/PDF where allowed, ORD build and pass/fail. Do not commit
project LandXML, DGN, private logs, PDFs or coordinate reports.
