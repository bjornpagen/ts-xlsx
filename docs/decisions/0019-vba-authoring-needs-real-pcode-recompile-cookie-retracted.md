# ADR 0019: VBA authoring needs real, compiled p-code; the "recompile cookie" premise is retracted

**Status:** Accepted (2026-07-24) · **retracts** the core mechanism of ADR 0017 §2.3c and ADR 0018 ·
amends nothing in ADR 0016's read view.

## Context

ADR 0017 (author from scratch) and ADR 0018 (edit an existing module's source by splice) were both built
on one premise:

> Emit a module with `MODULEOFFSET = 0` (no p-code) and set `_VBA_PROJECT` to a fixed 7-byte
> "unmatchable version" cookie (`0xCC 0x61 0xFF 0xFF 0x00 0x00 0x00`); Excel will see the version
> mismatch, discard the (absent) PerformanceCache, and **recompile every module from source on open**.

**That premise is false.** It was never actually exercised: ADR 0017's Excel oracle opened files with
`AutomationSecurity = ForceDisable`, which blocks the VBA engine _before_ it loads or compiles anything;
ADR 0018's GUI probe used `open-verdict.ps1`, which dismisses prompts with the safe negative and **never
clicks Enable Content**. Neither ever ran the VBA compiler. So "opens clean, source preserved" only ever
proved the file passed through structurally, not that a single module ever recompiled.

### What Excel actually does (verified 2026-07-24)

Byte-surgery against a genuine, known-good Excel-authored file, each variant opened with **macros
enabled** and the authored macro **executed over COM** (the new `execute-verdict.ps1` probe):

| Mutation against a known-good file                                           | Result                                                                                                                             |
| ---------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| change only a module's source; real p-code and real `_VBA_PROJECT` untouched | loads and runs (the p-code still runs, so **the new source is ignored**)                                                           |
| `MODULEOFFSET → 0` (p-code removed); `_VBA_PROJECT` untouched                | **"Invalid data format"**, project deleted                                                                                         |
| p-code present but garbage (same length); `_VBA_PROJECT` untouched           | **"Invalid data format"**, project deleted                                                                                         |
| `_VBA_PROJECT → cookie` only; real p-code and `MODULEOFFSET` untouched       | loads and runs correctly (**the cookie does nothing**)                                                                             |
| real-but-unrelated p-code plus a wildly different new source                 | loads clean, but **silently runs the stale p-code**; the new source's macros don't exist and the placeholder's original macro runs |

Conclusions, now facts rather than hypotheses:

1. **Excel does not recompile from source on open.** A module runs the compiled p-code
   (PerformanceCache) it ships; the source text is only recompiled when a human opens the VBE, edits a
   line, or runs Debug ▸ Compile.
2. **The `_VBA_PROJECT` version cookie has no effect** on loading when real p-code is present. It is not
   the recompile trigger it was believed to be. Worse, resetting `_VBA_PROJECT` to the cookie on a
   project that _has_ real p-code **crashes the VBA load** (RPC-fatal), so it was not even a harmless
   no-op.
3. **"Loads clean" is not evidence of correctness** for a VBA-authoring feature. A module can load
   without any error and silently run the wrong (stale) code, which is strictly worse than a loud failure.

There is therefore **no synthesizable placeholder** that makes Excel both accept the file and run the
authored source. Genuinely valid, **source-matched p-code is a hard requirement**, and only a real Excel
can produce it.

## Decision

1. **Authoring and editing VBA module _source_ moves out of the shipped library into an offline build
   tool, `tools/vba-compiler`.** It drives a real, headless Excel through the VBIDE object model
   (`VBComponents.Add`, then `CodeModule.AddFromString`, then `SaveAs`) to produce genuinely compiled,
   source-matched p-code, emitting either a `vbaProject.bin` (from-scratch, procedural or class) or a whole
   edited `.xlsm` (in-place, any module kind including document code-behind). It is a **probe/build tool,
   never in CI** (same contract as `tools/excel-oracle`), needs a licensed Excel and
   `AccessVBOM = 1`, and its output seeds committed corpus fixtures. The **shipped library stays pure-TS
   with zero Office dependency**: consumers attach the emitted bytes via `Workbook.vbaProjectBytes`.

2. **The unsound pure-TS authoring API is removed.** It emitted files that either fail to load or
   silently run stale code:
   - `writeVbaProject` / `Workbook.setVbaProject` (from-scratch synthesis, ADR 0017 §2.3c);
   - `editVbaModuleSources` / `Workbook.setVbaModuleSource` / `editXlsxVbaModuleSource(s)`
     (edit-source-by-splice, ADR 0018);
   - `addVbaModule` / `Workbook.addVbaModule` / `editXlsxVbaAddModule` (add-module, which also wrote a
     no-p-code module).

   `src/vba/project-writer.ts` (the synthesizer) is gone; its still-shared record/name encoders live on
   in `src/vba/vba-encoding.ts`.

3. **The purely-structural edits are kept, and fixed.** `removeVbaModule` and `addVbaReference` (and
   their `Workbook`/`editXlsxVba*` wrappers) never rewrite a module's p-code. They only edit the `dir`
   stream (and, for a removal, `PROJECT`/`PROJECTwm`). They were being broken _solely_ by the
   `_VBA_PROJECT → cookie` reset they inherited; **that reset is removed**, leaving `_VBA_PROJECT`
   untouched. Verified: both now load and run correctly, because the surviving modules keep their real
   p-code and the `dir` stream, authoritative for the module/reference list, carries the change.

4. **A new oracle makes this class of defect visible.**
   `.claude/skills/excel-gui-automation/scripts/execute-verdict.ps1` opens a file with macros enabled and
   `Application.Run`s a named authored macro, reporting whether it actually runs and returns the expected
   value. Every VBA-authoring artifact is verified this way, never by "opens clean" alone.

## Consequences

- **Breaking API change** (welcome, per the constitution): the source-authoring functions above are gone
  from the public API. There is no deprecation shim, because they produced silently-wrong output.
- Authoring VBA now requires an offline Excel build step. This is an honest reflection of reality: no
  pure-code path can produce runnable VBA, because runnable VBA _is_ compiled p-code only Excel emits.
- ADR 0016's read view is unchanged. Byte-for-byte macro **preservation** on round-trip, the safety
  floor, is untouched and remains the default for any `.xlsm` the library reads and rewrites.
- The prior ADRs' "Verified against real Excel 365" claims are **withdrawn** as never having exercised
  the compile/load step; this ADR's table is the corrected record.
