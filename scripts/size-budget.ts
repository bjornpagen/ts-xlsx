// Bundle-size budgets for the publishable build.
//
// Two numbers, because they answer two different questions.
//
// The **total** is every emitted `dist/**/*.js`: what ships in the tarball, and the tripwire
// against accidental bloat (an errant dependency inlined, dead code shipped).
//
// The **per-entry** numbers are what a consumer actually loads. Each public subpath in
// `package.json`'s `exports` is walked transitively through its static imports; the closure is the
// set of modules that must be present for that entry to evaluate. It is a lower bound on any
// bundler's answer, since `sideEffects: false` lets a bundler prune *within* these modules and
// never add to them. It is also the only number that notices the failures that matter here:
// a codec acquiring a value-import of something it previously needed only as a type, or the model
// reaching into a parser. The total cannot see either; both leave it unchanged.
//
// Budgets are tripwires, not targets. Raise one deliberately, with the same eyes a dependency
// addition would get. When you do, say in the commit *what* the entry gained.
//
// Run by the `size` gate of `verify --full`, which builds first, as well as by `prepublishOnly`.
// It used to be the publish step's alone, and `/customui` spent a release 3 KB over its budget on a
// tree that was green everywhere anyone looked: a budget only the publish step checks is not a
// tripwire. `docs/architecture.md` ("The size budgets are per entry") carries the reasoning.
//
//   node scripts/size-budget.ts

import {statSync} from 'node:fs';
import {join, resolve} from 'node:path';

import {closure, importedPaths, sourceFiles} from './module-graph.ts';
import {readPackageJson, ROOT} from './repo.ts';
import {verdict} from './verdict.ts';

const DIST = join(ROOT, 'dist');
// Raised from 580 when a run of read and write correctness fixes (VBA byte splicing, splice-anchored
// values and page breaks, validation attributes, drawing preservation, checked error codes, table
// style names) had grown the shipped JavaScript to exactly this line, each step within its entry's
// budget but the sum with no room left.
//
// Raised again, from 584, when the stylesheet reader's one-slot-per-`<fill>` draft put the sum a tenth
// over.
//
// Raised again, from 585, when an authored pivot began refusing a field used twice and a repeated
// source header, and flagging a value field that is also an axis, three tenths over.
//
// Raised again, from 586, when the CSV reader began refusing a file that does not fit the grid with its
// own `CsvParseError`, and checking a date and a number's digits before coercing them, four tenths over.
//
// Raised again, from 587, when the CSV writer began keeping row positions and refusing delimiters it
// cannot round-trip, three tenths over.
//
// Raised again, from 588, when the MS-OVBA compressor began padding a raw chunk to the 4096 bytes the
// format fixes, a tenth over.
//
// Raised again, from 589, by the [MS-CFB] layout records on every entry that carries the VBA editors,
// described above the entry budgets.
//
// Raised again, from 591, when the BIFF12 reader began reading sheet protection, a tenth over.
//
// Raised again, from 592, when pictures began carrying alternative text, a title, a crop and a link
// through the drawing codec, 4.6 KB over.
//
// Raised again, from 597, when a row or column splice began moving the references in formula text, on every sheet and in defined names. `core/formula-references.ts` is a reader of formula text
// of its own, and `core/grid-edits.ts` the pass that hands every formula the splice. 11.1 KB over.
//
// Raised again, from 609, when a table column began keeping its calculated column formula, which left
// the total sitting on its budget.
//
// Raised again, from 610, when a splice began moving a table's column formulas and an authored
// pivot's source, 0.7 to 1.6 KB over.
//
// Raised again, from 612, when the writer began replaying a workbook's splices over the charts and
// pivot caches it preserves, which takes a rewriter of those parts and a scanner for every element
// of a name, 0.8 to 4.0 KB over.
//
// Raised again, from 617, when the XML reader began translating the parts a Strict workbook carries
// through into Transitional: the namespace pairs and the percentage attributes, both out of the
// schema graph, and the tag scanner that rewrites them in place, 6.5 to 7.0 KB over. The container
// keeps only the relationship-type rule, since `/xlsb` would otherwise pay 11.5 KB for XML
// machinery a binary workbook, which has no Strict form, never uses.
//
// Raised again, from 625, when a function a formula passes as a value began taking the `_xleta.`
// prefix Excel stores it under: `core/function-values.ts`, the 528 names Excel prefixes, and the pass
// in `core/formula.ts` that weighs each against the defined names in scope. 5.5 KB over.
//
// Raised again, from 632, when the BIFF12 function table took the 105 functions Excel 2007 added past
// `RTD`, 1.9 KB over.
//
// Raised again, from 634, when conditional format, data validation and table column formulas began
// taking the function prefixes a cell formula does, 1.7 KB over.
//
// Raised again, from 636, when hyperlinks moved out of the cell value into a sheet overlay,
// `core/hyperlink.ts`, which every entry carrying the model loads, 1.2 KB over.
//
// Raised again, from 638, when conditional formats Excel keeps in the worksheet extension began
// to be read and written, inline styles and custom icons included, 9.5 KB over.
//
// Raised again, from 648, when a data bar's 2010 facets (lengths, border, direction, axis,
// negative colours, automatic anchors) began to be read from and written to its extension, 1.9 KB
// over.
//
// Raised again, from 650, when a picture added beside a kept drawing began to be written into that
// drawing rather than refused, 1.3 KB over.
//
// Raised again, from 652, when an array formula began reading and writing its range and the cell metadata that marks a dynamic array.
//
// Raised again, from 662, when an error Excel has no literal for began to be read from and written
// to the value metadata and rich-value parts that name it, 8.5 KB over.
const TOTAL_BUDGET_BYTES = 672 * 1024;

// Roughly a tenth of headroom over the measured closure, per entry: enough that ordinary growth is
// not a chore, tight enough that a whole codec crossing a boundary cannot hide inside it.
//
// Every number below was halved when `build` split into two tsc passes and the JS pass started
// stripping comments. That is not a budget cut: nothing left the closure, and no consumer loads a
// byte less than they did before the prose was measured as part of it. It is the measurement
// finally being of code. The old figures were ~47% comment, which is what had made this tripwire
// soft: a codec crossing a boundary is the failure these numbers exist to catch, and at the old
// scale one could have arrived inside a release's ordinary comment churn without moving them.
//
// Re-baselined again when the browser boundary landed (ADR 0040), because two things moved at once
// and in opposite directions. `src/sha512.ts` and `toBase64` replaced `node:crypto` and `Buffer`,
// which puts about 5 KB into the bottom layer that every entry reaching `/core` now carries, and
// `/customui`, which reaches almost nothing, went over a budget with 0.7 KB left in it. Meanwhile
// the streaming writer left `.` and `/xlsx` for `/node`, so those two fell. The tenth of headroom
// this comment describes had been eaten to a rounding error on several entries (`/core` sat 0.1 KB
// under its number); the figures below restore it against today's measurement rather than
// grandfathering the drift. `/errors` and `/vba` keep theirs, which are already deliberate.
//
// Re-baselined once more by the write-path and error-taxonomy corrections. Four small primitives
// landed below several entries at once and each is on a path those entries actually take:
// `elementRange` (editing a part at scanner-found offsets instead of by regular expression),
// `assertWritableNumber` and `formulaNumberLiteral` (a formula literal has a serialisation of its
// own, and the BIFF12 codec produces it), `read-repair.ts` and `xml-chars.ts` (the reader no longer
// hands a file-derived value to a guard written about the caller). Nothing crossed a boundary; four
// entries had simply been left sitting at a few tenths of a percent of headroom, which is not the
// tripwire this comment describes. The numbers below restore it against today's measurement.
//
// Re-baselined by a series of hostile-input and read-path corrections, each measured as it landed and
// none crossing a boundary. The attribute scanner became a hand-written linear scan (+1.4 KB on every
// entry that parses XML). Read repair and the column budget moved below both codecs into
// `io/read-policy/`, which the binary reader now calls (`/xlsb` +2.7 KB and two modules). The BIFF12
// formula decoder bounds its text, the grid refuses an out-of-bounds edit before mutating
// (`/core` +0.8 KB), and a shared formula's clone shifts whole-column and whole-row ranges (+1.1 KB on
// `.`). `.` went over, `/core` and `/xlsb` were left under a kilobyte, and the total under four; the
// numbers below restore them against today's measurement.
//
// Re-baselined again when a row or column splice began moving the coordinates a cell's value carries
// (a hyperlink's clickable range, a data table's ranges) beside the shared-formula master it already
// moved. `core/grid-edits.ts` is on every entry that reaches the model, so the ~0.9 KB landed on five
// entries at once, none of which had a kilobyte left after the VBA byte-splicing and model-assignment
// fixes just before it. Each number below restores about a kilobyte against today's measurement.
//
// And again, on `.` and `/xlsx` alone, when data validations began carrying `showDropDown` and
// `imeMode`: the reader and writer for them sit in the XML codec, and a closed token guard and its
// union came to a few tenths of a kilobyte more than either entry had left.
//
// Re-baselined when the [MS-CFB] header and directory-entry offsets moved into named records in
// `vba/cfb-format.ts` that the reader and the writer both index, where each had written them as bare
// numbers. The measurement is of unminified code, so every `HEADER_FIELD.fatSectors` costs its full
// text: about 1.8 KB across the CFB modules, on every entry that reaches the VBA editors, which is all
// of them but `/customui` and `/errors`. That is the price of one transcription of the layout instead
// of two, paid deliberately. Each number below is the next whole kilobyte above today's measurement.
const ENTRY_BUDGETS_KB: Readonly<Record<string, number>> = {
  // Raised from 570 when the future-function registry took [MS-XLSX]'s full table and the names Excel
  // prefixes beyond it, fifty-odd names more, which put this entry seven tenths over.
  //
  // Raised again, from 571, when the date-format renderer began keeping a short meridiem's spelling
  // to render it as Excel does, which put this entry two tenths over.
  //
  // Raised again, from 572, when conditional formats gained an integer boundary for priority, rank
  // and stdDev, a range filter on read and a refusal of an empty range, a tenth over.
  //
  // Raised again, from 573, when a typed cell without a value began reading as blank, which needs the
  // run machine to report whether an `<is>` opened at all, a tenth over.
  //
  // Raised again, from 574, when an authored pivot began refusing a field used twice and a repeated
  // source header, and the writer began flagging a value field that is also an axis, which left this
  // entry sitting on its budget.
  //
  // Raised again, from 575, by the same CSV reader bounds as `/csv`, two tenths over.
  //
  // Raised again, from 576, by the same CSV writer row positions and delimiter refusals as `/csv`, which
  // left this entry sitting on its budget.
  //
  // Raised again, from 577, when the attribute scanner began reading a raw tab, LF or CR in a value as
  // the space XML 1.0 makes it, two tenths over.
  //
  // Raised again, from 580, by the same BIFF12 sheet protection reader as `/xlsb`, which left this
  // entry sitting on its budget.
  //
  // Raised again, from 581, by the same picture properties as `/xlsx`, 4.3 KB over.
  //
  // Raised again, from 586, by the same formula reference moves as `/core`, 10.9 KB over.
  //
  // Raised again, from 597, when `duplicateRow` began copying a formula filled down rather than as written, three tenths over.
  //
  // Raised again, from 598, when a splice began moving a table's column formulas and an authored
  // pivot's source, 0.7 to 1.6 KB over.
  //
  // Raised again, from 601, when the writer began replaying a workbook's splices over the charts
  // and pivot caches it preserves, which takes a rewriter of those parts and a scanner for every
  // element of a name, 0.8 to 4.0 KB over.
  //
  // Raised again, from 606, when the XML reader began translating the parts a Strict workbook
  // carries through into Transitional: the namespace pairs and the percentage attributes, both out
  // of the schema graph, and the tag scanner that rewrites them in place, 6.5 to 7.0 KB over. The
  // container keeps only the relationship-type rule, since `/xlsb` would otherwise pay 11.5 KB for
  // XML machinery a binary workbook, which has no Strict form, never uses.
  //
  // Raised again, from 614, when a function a formula passes as a value began taking the `_xleta.`
  // prefix: the list of the functions Excel prefixes and the pass that decides it, 5.0 KB over.
  //
  // Raised again, from 620, by the same BIFF12 function table as `/xlsb`, 2.4 KB over.
  //
  // Raised again, from 623, when conditional format, data validation and table column formulas began
  // taking the function prefixes a cell formula does, 1.2 KB over.
  //
  // Raised again, from 625, by the same hyperlink overlay as `/core`, 0.8 KB over.
  //
  // Raised again, from 627, by the same extension-form conditional formats as `/xlsx`, 9.0 KB over.
  //
  // Raised again, from 637, by the same data-bar facets as `/xlsx`, 1.5 KB over.
  //
  // Raised again, from 639, by the same picture merge as `/xlsx`, 0.9 KB over.
  //
  // Raised again, from 640, when an array formula began reading and writing its range and the cell metadata that marks a dynamic array.
  //
  // Raised again, from 650, by the same rich-value errors as `/xlsx`, 8.8 KB over.
  '.': 660,
  // Raised from 207 when the VBA editors started writing `dir` records and cutting `PROJECT` lines as
  // bytes, rather than as spread arrays and re-encoded text. `Workbook` reaches the VBA editors, so
  // every entry that carries the model pays for it, and this one went over by a tenth of a kilobyte.
  //
  // Raised again, from 210, by the same structured-reference spelling as `/xlsb` and `/csv`, which
  // put this entry four tenths over.
  //
  // Raised again, from 211, when merged ranges began to be stored canonically (a sheet prefix refused,
  // a declared range without its rectangle an `InternalError`) and the merge-over-table test moved into
  // `core/merge.ts` for both codecs to share, which put this entry a kilobyte over.
  //
  // Raised again, from 212, when `core/pivot-table.ts` began refusing a field placed on two axes or twice
  // on one, and a source header row that repeats a name, nine tenths over.
  //
  // Raised again, from 213, when the MS-OVBA compressor began storing a chunk raw only at the 4096 bytes
  // the format fixes, padding a short one. `Workbook` reaches the VBA editors, so this entry pays for it,
  // two tenths over.
  //
  // Raised again, from 216, when `core/address.ts` began stating which names Excel reads as a
  // reference, the rule that quotes a sheet named `R1C1` and refuses a table named `T1`. Every entry
  // carrying the model loads that module; seven tenths over.
  //
  // Raised again, from 217, when a splice began flagging a data table's deleted input cell rather than
  // leaving the reference to be read from whatever moved into its place, a tenth over.
  //
  // Raised again, from 218, when a row or column splice began moving the references in formula text, on every sheet and in defined names.
  // The model now reads formula text: `core/formula-references.ts`, and the opaque-region scanner it
  // shares with the function mangling, which moved into `core/formula-scan.ts` beside it so this entry
  // takes the scanner without the mangling or the future-function table. The shared-formula
  // translation moved in with them from `core/formula.ts`. 14.5 KB over.
  //
  // Raised again, from 233, when a splice began moving a table's column formulas and an authored
  // pivot's source, 0.7 to 1.6 KB over.
  //
  // Raised again, from 236, when a hyperlink moved out of the cell value into a range-bound overlay on
  // the sheet, `core/hyperlink.ts`, which moves with a splice and a row copy, 2.4 KB over.
  //
  // Raised again, from 239, when a conditional format gained custom icons, the 2009 icon families
  // and the form it is stored in, 0.4 KB over.
  //
  // Raised again, from 240, when an array formula began reading and writing its range and the cell metadata that marks a dynamic array.
  './core': 242,
  // Raised from 554 by the same future-function registry as `.`, which put this entry nine tenths over.
  //
  // Raised again, from 555, by the same conditional-format boundaries as `.`, eight tenths over.
  //
  // Raised again, from 556, when the stylesheet reader began committing exactly one fill slot per
  // `<fill>` from a per-fill draft, so a malformed fill table no longer shifts every later `fillId`.
  // Half a kilobyte over.
  //
  // Raised again, from 557, by the same pivot role refusals as `/core`, eight tenths over.
  //
  // Raised again, from 558, when the package-level VBA edit began finding parts through the reader's
  // case-folding accessors and removing a stale signature's references at scanner-found offsets rather
  // than by pattern, three tenths over.
  //
  // Raised again, from 561, by the same BIFF12 sheet protection reader as `/xlsb`: `readXlsx` hands an
  // `.xlsb` package to that codec, six tenths over.
  //
  // Raised again, from 562, by the same data-table input flag as `/core`, read and written here as
  // `del1`/`del2`, three tenths over.
  //
  // Raised again, from 563, when the drawing codec began reading and writing a picture's alternative
  // text and title, its crop in either spelling `ST_Percentage` allows, and its link through a hyperlink
  // relationship, with the checks that keep each writable. 4.0 KB over.
  //
  // Raised again, from 568, by the same formula reference moves as `/core`, 10.6 KB over.
  //
  // Raised again, from 579, when `duplicateRow` began copying a formula filled down rather than as written, which left this entry sitting on its budget.
  //
  // Raised again, from 580, when a splice began moving a table's column formulas and an authored
  // pivot's source, 0.7 to 1.6 KB over.
  //
  // Raised again, from 583, when the writer began replaying a workbook's splices over the charts
  // and pivot caches it preserves, which takes a rewriter of those parts and a scanner for every
  // element of a name, 0.8 to 4.0 KB over.
  //
  // Raised again, from 588, when the XML reader began translating the parts a Strict workbook
  // carries through into Transitional: the namespace pairs and the percentage attributes, both out
  // of the schema graph, and the tag scanner that rewrites them in place, 6.5 to 7.0 KB over. The
  // container keeps only the relationship-type rule, since `/xlsb` would otherwise pay 11.5 KB for
  // XML machinery a binary workbook, which has no Strict form, never uses.
  //
  // Raised again, from 596, by the same function-value prefix as `.`, 4.7 KB over.
  //
  // Raised again, from 602, by the same BIFF12 function table as `/xlsb`, which `readXlsx` reaches for
  // an `.xlsb` package, 2.1 KB over.
  //
  // Raised again, from 605, by the same rule and table formula prefixes as `.`, 0.9 KB over.
  //
  // Raised again, from 607, by the same hyperlink overlay as `/core`, 0.5 KB over.
  //
  // Raised again, from 609, when the conditional-format codec began reading and writing the rules
  // Excel keeps in the worksheet extension, their inline styles included, 8.7 KB over.
  //
  // Raised again, from 618, when the conditional-format codec began reading and writing a data bar's
  // 2010 facets, 2.1 KB over.
  //
  // Raised again, from 621, when a picture added to a sheet that kept its drawing began to join that
  // drawing, its anchors and relationships numbered past the drawing's own, 0.5 KB over.
  //
  // Raised again, from 622, when an array formula began reading and writing its range and the cell metadata that marks a dynamic array.
  //
  // Raised again, from 632, when an error Excel has no literal for began to be read from and written to
  // the value metadata and rich-value parts that name it, in `io/cell-metadata/` for the model both
  // codecs resolve a cell through and in the XML codec for the parts, 8.3 KB over.
  './xlsx': 642,
  // Raised from 282 when the style primitives gained real clone plans. A font, a border and a fill
  // were each copied with a spread, which shares everything one level down, so the plans and their
  // exhaustiveness proofs are the fix rather than an addition. They sit in `core/style.ts`, which
  // every entry carries, and this was the one entry whose headroom the ~3 KB exhausted. Restores it
  // against that measurement rather than granting the growth a permanent home in the margin.
  //
  // Raised again, from 295, by the same VBA byte-splicing change as `/core`, which put this entry
  // 0.7 KB over. The new figure restores the headroom against today's measurement.
  //
  // Raised again, from 299, when a table's totals row began spelling its structured reference by
  // Excel's escaping rules. The rule sits in its own module so no entry carries the modern-function
  // list for it, yet its few hundred bytes still put this entry a tenth over.
  //
  // Raised again, from 300, when the future-function registry took [MS-XLSX]'s full table and the
  // names Excel prefixes beyond it, fifty-odd names more, which put this entry four tenths over.
  //
  // Raised again, from 301, by the same canonical merge storage as `/core`, eight tenths over.
  //
  // Raised again, from 302, by the same pivot role refusals as `/core`, seven tenths over.
  //
  // Raised again, from 303, by the same MS-OVBA raw-chunk padding as `/core`, two tenths over.
  //
  // Raised again, from 306, by the same reference-name rule as `/core`, eight tenths over.
  //
  // Raised again, from 307, when the worksheet reader began reading `BrtSheetProtection` and
  // `BrtSheetProtectionIso` into the model, which also brings the base64 encoder into this entry for
  // the credential. A protected `.xlsb` sheet used to read as unprotected. 2.3 KB over.
  //
  // Raised again, from 310, by the model half of the picture properties (`core/image.ts`), two tenths
  // over.
  //
  // Raised again, from 311, by the same formula reference moves as `/core`, 10.7 KB over.
  //
  // Raised again, from 322, when `duplicateRow` began copying a formula filled down rather than as written, a tenth over.
  //
  // Raised again, from 323, when a splice began moving a table's column formulas and an authored
  // pivot's source, 0.7 to 1.6 KB over.
  //
  // Raised again, from 325, when the writer began replaying a workbook's splices over the charts
  // and pivot caches it preserves, which takes a rewriter of those parts and a scanner for every
  // element of a name, 0.8 to 4.0 KB over.
  //
  // Raised again, from 327, by the same function-value prefix as `.`, 5.3 KB over. This entry only
  // reads the prefix, which needs the defined names and not the list, but `core/formula.ts` keeps both
  // directions of the mangling in one module and the list comes with it.
  //
  // Raised again, from 333, when the function table took the 105 functions Excel 2007 added past `RTD`,
  // which an `.xlsb` cites by index, 2.8 KB over.
  //
  // Raised again, from 336, by the same hyperlink overlay as `/core`, 2.8 KB over.
  //
  // Raised again, from 339, by the same conditional-format model as `/core`, 0.7 KB over.
  //
  // Raised again, from 340, by the data-bar facet types and guards the conditional-format model
  // gained, 0.2 KB over.
  //
  // Raised again, from 341, when an array formula began reading and writing its range and the cell metadata that marks a dynamic array.
  './xlsb': 344,
  // Raised from 210 when the CSV writer's private moment.js-style date table was replaced by a real
  // Excel number-format renderer (ADR 0041). It is the one entry that pays for it: the renderer sits
  // in `core/date-format.ts` apart from `core/date.ts` precisely so the four entries that never
  // render a date do not carry four kilobytes of month names and grammar.
  //
  // Raised again, from 216, when a worksheet model assignment began rehearsing on a scratch sheet. The
  // entry sat at exactly its budget, and those few lines in `core/worksheet.ts` put it a tenth over;
  // the new figure restores a kilobyte of headroom against today's measurement.
  //
  // Raised again, from 219, by the same structured-reference spelling as `/xlsb`, which put this
  // entry two tenths over.
  //
  // Raised again, from 220, by the conditional-format overlay's refusal of an empty range, which
  // every entry carrying the model pays for, a tenth over.
  //
  // Raised again, from 221, by the same canonical merge storage as `/core`, three tenths over.
  //
  // Raised again, from 222, by the same pivot role refusals as `/core`, two tenths over.
  //
  // Raised again, from 223, when the reader began refusing a record wider than the grid, or more
  // records than it has rows, with `CsvParseError` rather than the grid's `RangeError`, and checking an
  // ISO date and a number's significant digits before coercing. Three tenths over.
  //
  // Raised again, from 224, when the writer began placing row N on line N and refusing a quote, CR or
  // LF as a delimiter and a row delimiter it cannot keep apart from the data. Two tenths over.
  //
  // Raised again, from 227, by the same reference-name rule as `/core`, a kilobyte over.
  //
  // Raised again, from 229, by the same formula reference moves as `/core`, 14.8 KB over.
  //
  // Raised again, from 244, when `duplicateRow` began copying a formula filled down rather than as written, two tenths over.
  //
  // Raised again, from 245, when a splice began moving a table's column formulas and an authored
  // pivot's source, 0.7 to 1.6 KB over.
  //
  // Raised again, from 247, by the same hyperlink overlay as `/core`, 2.6 KB over.
  //
  // Raised again, from 250, by the same conditional-format model as `/core`, 0.6 KB over.
  //
  // Raised again, from 251, by the same data-bar facet model as `/xlsb`, 0.1 KB over.
  //
  // Raised again, from 252, when an array formula began reading and writing its range and the cell metadata that marks a dynamic array.
  './csv': 254,
  // The streaming writer and the write half it rides on, and nothing of the reader: a jump here is
  // the read path arriving, which would mean the entry had stopped being about one thing.
  //
  // Raised from 395 when the reader stopped handing file-derived names straight to the model's
  // authoring guards. `io/read-policy/read-repair.ts` and `xml/xml-chars.ts` are the two new modules, ~2 KB
  // between them, and both are on the untrusted-input path rather than beside it: without them a
  // corrupt package raised an `AuthoringError` blaming the caller, or a native `SyntaxError` the
  // taxonomy cannot see at all. This entry had 2.2 KB left in it, which is a rounding error and not
  // the headroom described above; the new figure restores it against today's measurement.
  //
  // Raised again, from 405, by the module seams: `worksheet-merges.ts`, `workbook-media.ts`,
  // `font-xml.ts` and `xml-attrs.ts` are four slices lifted out of files that had grown past what
  // anyone can read, and the code inside them did not change. What a module costs that a block of a
  // larger file does not is its import statements and its export keywords, which came to about
  // 2.4 KB across the four. Paying that for four seams is the trade this project takes; noticing it
  // is what the tripwire is for.
  //
  // Raised again, from 413, when page breaks joined the splice participants and kept their manual
  // flag through a save. The model this entry writes carries `core/grid-edits.ts`, and the new step
  // put it 0.7 KB over; the new figure restores the headroom against today's measurement. And to 417
  // when the row writer began checking a cell error's code against the closed set before writing it,
  // which left the entry sitting exactly on its budget.
  //
  // Raised again, from 417, by the same conditional-format boundaries as `.`, seven tenths over.
  //
  // Raised again, from 418, by the same pivot role refusals as `/core`, three tenths over.
  //
  // Raised again, from 421, by the same picture properties as `/xlsx`, which this entry's writer carries,
  // 3.6 KB over.
  //
  // Raised again, from 425, by the same formula reference moves as `/core`, 11.1 KB over.
  //
  // Raised again, from 437, when the table codec began reading and writing a calculated column
  // formula, which left the entry sitting on its budget.
  //
  // Raised again, from 438, when a splice began moving a table's column formulas and an authored
  // pivot's source, 0.7 to 1.6 KB over.
  //
  // Raised again, from 440, when the writer began replaying a workbook's splices over the charts
  // and pivot caches it preserves, which takes a rewriter of those parts and a scanner for every
  // element of a name, 0.8 to 4.0 KB over.
  //
  // Raised again, from 445, by the same function-value prefix as `.`, 4.7 KB over.
  //
  // Raised again, from 451, by the same rule and table formula prefixes as `.`, which this entry's writer
  // carries, 0.4 KB over.
  //
  // Raised again, from 452, by the same hyperlink overlay as `/core`, 0.9 KB over.
  //
  // Raised again, from 454, by the same extension-form conditional formats as `/xlsx`, 9.1 KB over.
  //
  // Raised again, from 464, by the same data-bar facets as `/xlsx`, 1.5 KB over.
  //
  // Raised again, from 466, by the same picture merge as `/xlsx`, 0.9 KB over.
  //
  // Raised again, from 467, when an array formula began reading and writing its range and the cell metadata that marks a dynamic array.
  //
  // Raised again, from 475, by the same rich-value errors as `/xlsx`, 8.6 KB over.
  './node': 485,
  // Raised from 50 when the MS-OVBA encoder stopped rescanning its whole back-window for every
  // output byte. The hash chain that replaced the rescan is the cost, and it buys a time bound on a
  // path an untrusted `.xlsm` reaches through `removeVbaModule`; the CFB and `dir` guards landed
  // alongside it are the rest. Restores this entry's tenth of headroom against that measurement.
  './vba': 57,
  './customui': 16,
  // The taxonomy reaches nothing but itself, and that is the point: classifying a failure must
  // not cost a parser. A jump here means an error class started importing the layer it describes.
  './errors': 4,
};

// Emitted JS, not source, and the emitter picks its own quoting: double under TypeScript 6 and single
// under 7. Both forms are matched, by the shared walker, which is the reason it is shared: two of the
// four gates matched one quote style only, so the same flip in the source formatter would have made
// them report a clean graph.
const staticImports = (file: string): string[] => importedPaths(file);

function bytes(files: Iterable<string>): number {
  let total = 0;
  for (const file of files) total += statSync(file).size;
  return total;
}

const kb = (n: number) => `${(n / 1024).toFixed(1)} KB`;

const pkg = readPackageJson();
const over: string[] = [];

const all = sourceFiles(DIST, '.js');
const total = bytes(all);
console.log(
  `total runtime JS: ${kb(total)} across ${all.length} file(s); budget ${kb(TOTAL_BUDGET_BYTES)}`,
);
if (total > TOTAL_BUDGET_BYTES) {
  over.push(`total is over by ${kb(total - TOTAL_BUDGET_BYTES)}`);
}

console.log('\nper entry, the module closure a consumer of that subpath loads:\n');
for (const [subpath, target] of Object.entries(pkg.exports)) {
  const emitted = typeof target === 'string' ? undefined : target.default;
  if (emitted === undefined || !emitted.endsWith('.js')) continue;

  const budgetKb = ENTRY_BUDGETS_KB[subpath];
  if (budgetKb === undefined) {
    over.push(`"${subpath}" is published with no budget in ENTRY_BUDGETS_KB`);
    continue;
  }
  const reached = closure(resolve(ROOT, emitted), staticImports);
  const size = bytes(reached);
  const budget = budgetKb * 1024;
  const state = size > budget ? `OVER by ${kb(size - budget)}` : 'ok';
  console.log(
    `  ${subpath.padEnd(12)} ${kb(size).padStart(9)}  ${String(reached.size).padStart(3)} modules   budget ${kb(budget).padStart(9)}   ${state}`,
  );
  if (size > budget) over.push(`"${subpath}" is over by ${kb(size - budget)}`);
}

// Through `verdict` like every sibling, and `process.exitCode` rather than `process.exit`. This is
// a publish-blocking gate whose output `verify.ts` captures through a pipe, which is the exact
// configuration `verdict.ts`'s header describes: `process.exit` truncates an unflushed pipe, so the
// gate that most needs its diagnostic read was the one that could lose it.
verdict({
  gate: 'size',
  problems: over.map((line) => `  ${line}`),
  ok: 'every entry within its budget',
  failure: 'entry closure(s) over budget; investigate the growth or raise the budget deliberately',
});
