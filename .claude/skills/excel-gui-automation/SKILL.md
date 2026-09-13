---
name: excel-gui-automation
description: Use on Windows with desktop Microsoft Excel installed when you need something the headless COM oracle (tools/excel-oracle) cannot give you - the INTERACTIVE open verdict (does Excel open a file clean, prompt to repair, auto-repair, warn on a format mismatch, or reject it) that DisplayAlerts=$false suppresses; the repair log Excel writes; a fixture for a feature with no COM surface authored by driving the real ribbon; or a visual check of how a file renders in the actual app. Escalation from tools/excel-oracle and the validate-ooxml skill, not a replacement - try those first.
metadata:
  # For working *on* ts-xlsx, not *with* it: it drives this repo's PowerShell scripts. `npx skills
  # add shbernal/ts-xlsx` walks .claude/skills/ too, and this flag is what keeps it out of the menu
  # a consumer sees. Set INSTALL_INTERNAL_SKILLS=1 to install it here.
  internal: true
---

# Excel GUI Automation

Drives the real, visible desktop Excel window: foreground control, ribbon
KeyTips over SendKeys, and UI Automation (`InvokePattern`/`TogglePattern`) for
anything keyboard and synthetic mouse can't reach. The foreground/screenshot/
KeyTip/UIA machinery is a port of the **verified** PowerPoint equivalent
(`ts-pptx/.agents/skills/powerpoint-gui-automation`); the Excel-specific
`open-verdict.ps1` and the repair-dialog patterns are **new and unverified until
their first real run** - treat the timings and English-locale dialog strings in
the scripts as starting points to confirm with `-Dump`, not gospel.

## When this is the right tool (and when it is not)

Three cheaper oracles come first; reach here only for what they structurally cannot do:

| Tool                                      | Answers                                                                                                        | Cannot                                                                                                                                             |
| ----------------------------------------- | -------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------- |
| `validate-ooxml` skill (OpenXmlValidator) | Is the XML ECMA-376 conformant?                                                                                | Whether _Excel_ accepts it - Excel tolerates many invalid files and rejects some valid ones.                                                       |
| `tools/excel-oracle` (headless COM)       | Formula recalc, canonical re-save, model readback; _detects_ that a repair happened.                           | The **interactive** repair experience - `DisplayAlerts=$false` suppresses the modal, so it never reproduces the dialog or captures the repair log. |
| **this skill** (interactive GUI)          | The open verdict a user actually sees; the repair log; no-COM-surface authoring; visual render.                | Nothing fast, deterministic, or CI-able - see below.                                                                                               |
| **this skill** (`execute-verdict.ps1`)    | Whether an authored/edited VBA module actually LOADS and RUNS its source once macros are enabled (Workflow C). | Same PROBE contract - never CI.                                                                                                                    |

**This is a PROBE, not a test** - same contract as `tools/excel-oracle` (ADR 0012,
seed+lock split). Its output is a _recorded fact that seeds a corpus case_. **Never
wire it into CI**: it needs a licensed Excel, an interactive desktop session, and
is inherently slow and flaky. The committed fixture + a Tier-2 seam fact are what
run in CI; the live GUI never does.

**Prefer COM.** For golden fixtures (Excel's own Save-As), formula recalc, and
model round-trips, `tools/excel-oracle` is strictly more reliable. Before ribbon-
driving a "no COM surface" feature, verify that claim against the real
`Application`/`Workbook`/`Worksheet`/`Shapes` object model - Excel's COM surface
is broad and most features are reachable.

## Prerequisites

- **Interactive console session, not session 0 or disconnected.** A headless/
  service session renders nothing, so `CopyFromScreen` returns black or stale
  pixels. Check:
  ```
  query session
  ```
  Look for `>console ... N  Active` (the `>` marks the current session). If it
  shows `services`/session 0 or a `Disc` state, GUI automation is impossible
  here - stop and say so rather than producing plausible-but-empty screenshots.
- **Desktop Excel installed and COM-registered** (verified present: `Excel.Application`
  = 16.0). `open-verdict.ps1` launches it with `/x` (a separate process) itself;
  the ribbon/screenshot scripts need a visible Excel window already open.
- **Clear stale recovery state first.** A Document Recovery pane intercepts
  keyboard input before the ribbon does. Delete
  `HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\{DocumentRecovery,StartupItems}`
  before ribbon-driving.
- **Run scripts through the PowerShell (pwsh 7) tool** with the call operator
  (`& '...\script.ps1' ...`). No `-ExecutionPolicy Bypass` (trips the sandbox's
  weaken-security classifier). `uia-lib.ps1` needs `UIAutomationClient`, which
  requires an STA apartment; if you invoke via a raw `powershell.exe`/`pwsh`
  call instead of the tool's default host, add `-STA`.
- **Keep script string literals ASCII-only.** An em-dash inside a quoted string
  in a `.ps1` run with `-File` corrupted the encoding and threw a spurious
  parse error far from the real problem. (All shipped scripts are ASCII;
  verified with the AST parser.)

## Workflow A - the open verdict (the headline Excel use)

The one thing no other oracle can produce. Read-only by default (dismisses any
prompt with the safe negative, rewriting nothing).

1. **Run the verdict, capturing a screenshot.** Emits one JSON blob on stdout.
   ```
   & '.claude\skills\excel-gui-automation\scripts\open-verdict.ps1' `
       -Path .\test\corpus\fixtures\suspect.xlsx -Shot .tmp\verdict.png -CloseAfter
   ```
   `verdict` is one of: `clean`, `repaired` (Excel silently fixed it - title
   carries `Repaired`), `repair-prompt` (the modal "We found a problem with
   some content..." - the case COM hides), `format-mismatch`, `rejected`
   (Excel refuses the file), or `timeout`. Read the PNG to confirm the classifier
   matched reality.
2. **If you need Excel's recovered output and the repair log,** accept the repair
   (this rewrites/creates files, so pass an explicit destination):
   ```
   & open-verdict.ps1 -Path .\broken.xlsx -AcceptRepair `
       -SaveRepairedTo .\test\corpus\fixtures\broken.recovered.xlsx -CloseAfter
   ```
   `repairLog` lists the log Excel wrote to `%TEMP%` (`error<pid><n>_01.xml`),
   whose `<removedRecord>` entries name what the repair dropped, by part only;
   `repairedPath` is Excel's recovery of your content, reached through the ROT
   (`rot-lib.ps1`) and saved - gold for a regression fixture. Read both: only the
   recovered package says which cells went and what Excel kept.
3. **Seed a corpus case** from the recorded verdict (see the `write-corpus-case`
   skill): the fixture + the expected `verdict` become the durable, CI-runnable
   artifact; the GUI run does not repeat in CI.

The classifier keys off Excel's English-locale dialog strings (encoded as regexes
at the top of `open-verdict.ps1`). On first real run, if a verdict comes back
`timeout` or `unknown` when the screenshot clearly shows a dialog, dump the real
text and widen the patterns - do not assume the wording.

> **`open-verdict.ps1` is structurally BLIND to VBA-authoring defects.** It never
> enables macros, so it only catches package-level (structural) repair - never a
> module that fails to compile, or one that loads clean but silently runs the
> wrong (stale) p-code, once macros are on. For anything that authors/edits VBA,
> use **Workflow C** below, not this one. (Verified 2026-07-24: a from-scratch VBA
> project passed `open-verdict.ps1` "clean" yet threw "Invalid data format" the
> instant Enable Content was clicked; another loaded clean but ran stale p-code
> from an unrelated module.)

## Workflow C - the VBA execute verdict (does the authored macro actually run?)

The verdict that proves a VBA-authoring feature works: not "opens clean" but "runs
the authored source." `execute-verdict.ps1` opens the file on a private, headless
Excel COM instance with `AutomationSecurity = Low` (the automation equivalent of a
user clicking Enable Content: the VBA project is loaded and COMPILED on open), then
runs a known macro and reports its result.

```
pwsh -File .claude\skills\excel-gui-automation\scripts\execute-verdict.ps1 `
    -Path .\authored.xlsm -RunMacro AddThem -RunArgs 4,5 -ExpectResult 9 -ReadModule Module1
```

- **`passed: true` with the expected `runResult`** - the authored source is live
  and correct.
- **`runError: "...macro may not be available..."`** - the module did not
  load/compile (bad p-code), or Excel is running stale p-code lacking that macro.
- **`openThrew: true`** - Excel rejected the file structurally on open.

Rules that matter:

- **Always pass `-RunMacro`.** `AutomationSecurity=Low` + `DisplayAlerts=$false`
  suppresses the interactive "Invalid data format" modal, so brokenness surfaces
  as a Run failure, NOT a popup. "Opened without error" alone proves nothing.
- **Pick a no-UI `Function`** (deterministic return, no `MsgBox`). A `MsgBox`
  blocks `Run` indefinitely and hangs the probe.
- Runs under pwsh 7 or Windows PowerShell (only needs `New-Object -ComObject`; no
  GetActiveObject, no UI Automation), and leaves no Trusted-Documents residue, so
  it is safely repeatable. Seed a corpus case from the recorded `passed`/`runResult`.

## Workflow B - author a fixture for a feature with no COM surface

Only after confirming the feature genuinely has no COM/VBA equivalent.

1. **Sanity check the window foregrounds** (a background/detached terminal
   otherwise stays in front; plain `SetForegroundWindow` is blocked by Windows'
   foreground lock - the lib uses the `AttachThreadInput` trick):
   ```
   & foreground-and-shoot.ps1 -Shot .tmp\shot0.png
   ```
   Read the PNG; confirm you're looking at Excel.
2. **Discover ribbon KeyTips.** Press Alt alone and screenshot - KeyTips are
   drawn as overlay badges on every tab/button, so read them off the image:
   ```
   & drive-ribbon.ps1 -KeyTips '{ESC};{ESC};%' -Shot .tmp\keytips.png
   ```
   Then Alt -> tab-letter to reveal that tab's KeyTips, and down into any dropdown.
3. **Drive the action + open the menu item, then DUMP the dialog** before
   guessing any control name (displayed text and accessible name differ, and
   Invoke silently no-ops on a wrong name):
   ```
   & drive-ribbon.ps1 -KeyTips '{ESC};{ESC};%;N' -UiaMenuItem '<item>' `
       -DialogTitleLike '<Dialog *>' -Dump -Shot .tmp\dlg.png
   ```
4. **Toggle/submit using the REAL names from the dump:**
   ```
   & drive-ribbon.ps1 -DialogTitleLike '<Dialog *>' -Toggle '<checkbox name>' `
       -InvokeButton 'OK' -Shot .tmp\done.png
   ```
5. **Verify against the actual OOXML, not the screenshot.** A screenshot proves
   the UI reacted, not that the XML is correct or wired up:
   ```
   & save-and-extract.ps1 -NameLike '<name>*' -DestDir .tmp\gui-extract
   ```
   then read the extracted `xl/worksheets/sheetN.xml` and `_rels` directly.

## Hygiene before committing anything GUI-authored

- A GUI-authored fixture's `docProps/core.xml` carries the interactively
  logged-in user's real name and may carry add-in residue. **Scrub before
  committing to the repo.**
- The scripts only ever act on an EXCEL.EXE **this run spawned** (PID diff
  against a pre-run snapshot) - cleanup never kills a session you didn't start.
  `open-verdict.ps1 -CloseAfter` closes only its own instance; without it, an
  EXCEL.EXE is left running for you to inspect.

## Why keyboard/mouse alone don't work on ribbon popups (do not re-try these)

Verified on the PowerPoint sibling; Office ribbons behave identically. Chorded
`SendKeys` (`%N`) fires the OLD Alt-accelerator, not the KeyTip overlay - KeyTips
need Alt pressed-and-released, then each letter as its own sequential
`SendKeys.SendWait`. Inside an open dropdown, arrow keys, `{ENTER}`, and
synthetic `mouse_event` clicks all register-but-do-not-activate. **Office ribbon
popups and their dialogs run on a separate input queue (UIPI integrity
isolation);** UIA's `InvokePattern`/`TogglePattern` operate on the accessibility
tree directly and bypass it - that is the only reliable path (`uia-lib.ps1`).

## Other gotchas

- **Full-screen captures and unscoped UIA enumeration both leak other windows'
  content.** `Save-WindowScreenshot` crops to the target window rect;
  `Find-UiaWindowsByPid` restricts the verdict scan to the spawned Excel PID;
  `Find-UiaDialog` scopes by title. Always pass a PID/title scope - never fall
  back to unscoped desktop search without saying so.
- **Loading `UIAutomationClient` flips the process DPI-aware mid-session,** which
  can jump screenshot resolution (e.g. 1280x800 -> 1920x1200). Harmless for UIA
  (coordinate-free); take any "before" screenshot before dot-sourcing `uia-lib.ps1`.
- **Timing** (starting points - re-tune if a step hasn't visibly settled in a
  screenshot): ~700ms after foregrounding before input; ~700-900ms between
  sequential KeyTips (600ms sometimes too short for a submenu to arm);
  ~1200-1500ms settle after a menu-opening Invoke; ~600ms per verdict poll.
- **An accepted repair leaves Excel modal** behind a `Repairs to '<file>'` dialog
  (class `bosa_sdm_XL9`) whose list and buttons UIA does not expose; every COM call
  is rejected (`RPC_E_CALL_REJECTED`) until it closes. `open-verdict.ps1` finds it
  by title and posts `WM_CLOSE`. Never click `Close` across an instance's windows:
  the workbook frame has one too, and its save prompt is a modal of its own.
- **PowerPoint starts Excel instances for its charts** at any moment. Track the
  process a script launched (`Start-Process -PassThru`), never "every EXCEL.EXE
  that appeared since", or cleanup kills a chart someone is editing.
- The sandbox's `Remove-Item` false-positive guard (trips when the command text
  also holds regex-like substrings such as `r:` or `\w+`, easy to hit near
  quoted XML rels) applies here; `save-and-extract.ps1` extracts to a fresh
  directory rather than deleting one.

## Scripts

- `scripts/xl-window-lib.ps1` - dot-source library: `Get-ExcelHwnd`,
  `Set-ExcelForeground` (the `AttachThreadInput` foreground-lock workaround, with
  `-TargetPid` to pin one instance), `Save-WindowScreenshot` (window-rect-scoped),
  `Send-KeyTipSequence`.
- `scripts/uia-lib.ps1` - dot-source library: `Find-UiaDialog`,
  `Find-UiaWindowsByPid`, `Get-UiaVisibleText`, `Find-UiaElementByName`,
  `Get-UiaControlDump`, `Invoke-UiaElement`, `Set-UiaToggleOn`.
- `scripts/open-verdict.ps1` - **the headline tool:** launch a file in
  interactive Excel, classify open verdict, screenshot, optionally accept the
  repair + capture the log, emit JSON. Read-only unless `-AcceptRepair`.
  **Blind to VBA-authoring defects** (never enables macros) - use
  `execute-verdict.ps1` for those.
- `scripts/execute-verdict.ps1` - **the VBA verdict (Workflow C):** open with
  macros enabled (headless COM, `AutomationSecurity=Low`), run a known authored
  macro, emit JSON `{openThrew, ran, runResult, runError, moduleLines, passed}`.
  The only oracle that proves an authored/edited module actually LOADS and RUNS
  its source. Runs under pwsh or Windows PowerShell.
- `scripts/foreground-and-shoot.ps1` - sanity check / "just look at current state".
- `scripts/drive-ribbon.ps1` - ribbon driver: KeyTips -> UIA menu invoke ->
  optional dump/toggle/button-invoke -> screenshot. See its comment-based help.
- `scripts/save-and-extract.ps1` - save the driven workbook, reached through the
  ROT, and extract the package, to read the real OOXML the GUI produced.
- `scripts/rot-lib.ps1` - dot-source library: `Get-RunningWorkbook` (an open
  workbook by path, in whichever Excel holds it), `Invoke-WhenExcelAccepts`
  (retry a COM call Excel rejects as busy), `Get-RunningObjectNames`. Replaces
  `Marshal.GetActiveObject`, which pwsh 7 does not have and which named the first
  Excel registered rather than the one a script spawned.
