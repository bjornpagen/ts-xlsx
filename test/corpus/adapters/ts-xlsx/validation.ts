// Data validation rules: authoring them, reading them back, and their serialized XML.

import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {partMapOf} from './package-facts.ts';
import {readFixture, readXlsx, Workbook, writeXlsx, XlsxError} from './runtime.ts';
import {attrsOf, expandSqref} from './xml-probes.ts';

export const validation = {
  // Author one list validation that hides its dropdown arrow and sets an input mode, in the standard
  // form on A1 and the extended form on B1, write it, read it back and write it again →
  // { written, read, rewritten }: `written`/`rewritten` give each form's `{showDropDown, imeMode}`
  // attributes (null when absent), and `read` gives each cell's rule as `{suppressDropDown, imeMode}`.
  dropdownAndImeModeReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    const rule = {
      type: 'list' as const,
      formulae: ['"a,b"'],
      suppressDropDown: true,
      imeMode: 'off' as const,
    };
    sheet.addDataValidation('A1', rule);
    sheet.addDataValidation('B1', rule, {extended: true});
    const attributes = (bytes: Uint8Array) => {
      const xml = partMapOf(bytes)['xl/worksheets/sheet1.xml'] ?? '';
      const pick = (pattern: RegExp) => {
        const attrs = attrsOf(xml.match(pattern)?.[0] ?? '<none/>');
        return {showDropDown: attrs.showDropDown ?? null, imeMode: attrs.imeMode ?? null};
      };
      return {
        standard: pick(/<dataValidation\b[^>]*>/),
        extended: pick(/<x14:dataValidation\b[^>]*>/),
      };
    };
    const written = writeXlsx(wb);
    const read = readXlsx(written);
    const cell = (ref: string) => {
      const back = read.getWorksheet('S')?.dataValidationAt(ref);
      return {suppressDropDown: back?.suppressDropDown ?? null, imeMode: back?.imeMode ?? null};
    };
    return {
      written: attributes(written),
      read: {A1: cell('A1'), B1: cell('B1')},
      rewritten: attributes(writeXlsx(read)),
    };
  },

  // Author a date-type validation whose operand coerces to a serial (or fails to), write, and report
  // the emitted first bound → { formula1, hasNaN, refused }. A real date writes a numeric serial; a
  // non-coercible operand is a number the format cannot spell, so the write is refused (`refused`,
  // with no package and therefore no "NaN") rather than serialising the literal "NaN".
  authorDateValidation(operand: string) {
    const serial = (() => {
      const ms = Date.parse(operand);
      return Number.isNaN(ms) ? Number.NaN : ms / 86_400_000 + 25_569; // Unix epoch → 1900 serial
    })();
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.addDataValidation('A1', {type: 'date', operator: 'greaterThan', formulae: [serial]});
    let bytes: Uint8Array;
    try {
      bytes = writeXlsx(workbook);
    } catch (error) {
      if (!(error instanceof XlsxError)) throw error;
      return {formula1: null, hasNaN: false, refused: true};
    }
    const sheetXml = partMapOf(bytes)['xl/worksheets/sheet1.xml'] || '';
    const formula1 = (sheetXml.match(/<formula1>([\s\S]*?)<\/formula1>/) || [])[1] ?? null;
    return {formula1, hasNaN: /NaN/.test(sheetXml), refused: false};
  },

  // Apply one list validation with a cross-sheet source over a vertical span, round-trip, and report
  // the per-row source references plus how many dataValidation blocks were emitted → { source,
  // formulae, allIdentical, sqrefBlocks }. Every row must keep the exact source (no relative drift),
  // and the identical rules collapse into one spanning sqref.
  listValidationSourceRangeAcrossRows(rows: Untyped, source: Untyped) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.addDataValidation(`A1:A${rows}`, {type: 'list', formulae: [source]}, {extended: true});
    const buffer = writeXlsx(workbook);
    const reloaded = readXlsx(buffer).getWorksheet('S');
    const formulae = [];
    for (let r = 1; r <= rows; r++)
      formulae.push(reloaded!.dataValidationAt(`A${r}`)?.formulae?.[0] ?? null);
    const allIdentical = formulae.every((f) => f === source);
    const dvXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const sqrefBlocks = (dvXml.match(/<(?:x14:)?dataValidation[\s>]/g) || []).length;
    return {source, formulae, allIdentical, sqrefBlocks};
  },

  // Read a fixture and report each covered cell's data validation → { cells: { 'Sheet!Ref': rule },
  // count }. A rule authored over a multi-cell range must be reported on EVERY covered cell (the
  // range form is resolved per cell), and a reference/name operand must survive as its string.
  readFixtureValidations(rel: string) {
    const wb = readFixture(rel);
    const cells: Record<string, Untyped> = {};
    for (const sheet of wb.worksheets) {
      for (const {sqref} of sheet.dataValidations) {
        for (const ref of expandSqref(sqref)) {
          const dv = sheet.dataValidationAt(ref);
          if (dv) cells[`${sheet.name}!${ref}`] = JSON.parse(JSON.stringify(dv));
        }
      }
    }
    return {cells, count: Object.keys(cells).length};
  },

  // Read a fixture and report each sheet's data-validation rules, de-duplicated by content with a
  // per-rule coverage count → { sheets: { name: { rules: [{rule, coverageCount}], ruleCount } } }.
  // Reads the worksheet overlay (not populated cells), so a rule over an empty range is still seen,
  // and surfaces a reference source (a defined name, a cross-sheet range) as its verbatim formula
  // text rather than "[object Object]".
  readFixtureValidationRules(rel: string) {
    const wb = readFixture(rel);
    const sheets: Record<string, Untyped> = {};
    for (const sheet of wb.worksheets) {
      const byContent = new Map();
      for (const {sqref, rule} of sheet.dataValidations) {
        const key = JSON.stringify(rule);
        const entry = byContent.get(key) || {rule: JSON.parse(key), coverageCount: 0};
        entry.coverageCount += expandSqref(sqref).length;
        byContent.set(key, entry);
      }
      sheets[sheet.name] = {rules: [...byContent.values()], ruleCount: byContent.size};
    }
    return {sheets};
  },

  // Read a fixture, write it back, and report the data-validation facts of the *re-serialized*
  // package: both the standard `<dataValidation>` entries and the extended `<x14:dataValidation>`
  // form (2009 extension schema, carried in `<extLst>`, used for cross-sheet / whole-column list
  // sources). Lets a case assert a template's validation survives a read→write round-trip rather than
  // being silently dropped because only the standard form was understood.
  roundtripFixtureValidationXml(rel: string) {
    const parts = partMapOf(writeXlsx(readFixture(rel)));
    const sheetParts = Object.keys(parts)
      .filter((p) => /^xl\/worksheets\/sheet\d+\.xml$/.test(p))
      .sort();

    const sheets: Record<string, Untyped> = {};
    let totalStandard = 0;
    let totalExt = 0;
    for (const p of sheetParts) {
      const xml = parts[p] || '';
      // `[ >]` after the tag name separates an individual entry from its `<dataValidations>` /
      // `<x14:dataValidations>` container (whose next char is `s`).
      const standardCount = [...xml.matchAll(/<dataValidation[ >]/g)].length;
      const extCount = [...xml.matchAll(/<x14:dataValidation[ >]/g)].length;
      const extSqrefs = [...xml.matchAll(/<xm:sqref>([^<]*)<\/xm:sqref>/g)].map((m) => m[1]);
      const standardRules = [
        ...xml.matchAll(/<dataValidation\b([^>]*)>([\s\S]*?)<\/dataValidation>/g),
      ].map((m) => {
        const a = attrsOf(`<x ${m[1]}>`);
        const f1 = (m[2]!.match(/<formula1>([\s\S]*?)<\/formula1>/) || [])[1] ?? null;
        return {
          type: a.type ?? null,
          sqref: a.sqref ?? null,
          errorTitle: a.errorTitle ?? null,
          error: a.error ?? null,
          formula1:
            f1 == null
              ? null
              : f1
                  .replace(/&quot;/g, '"')
                  .replace(/&lt;/g, '<')
                  .replace(/&gt;/g, '>')
                  .replace(/&amp;/g, '&'),
        };
      });
      sheets[p] = {
        standardCount,
        extCount,
        extSqrefs,
        standardRules,
        hasExtLst: /<extLst\b/.test(xml),
      };
      totalStandard += standardCount;
      totalExt += extCount;
    }
    return {sheets, totalStandard, totalExt, totalValidations: totalStandard + totalExt};
  },

  // Attach a list validation whose source formula is supplied (possibly with a leading '='), write,
  // and report the serialized formula1 text → { formula1, hasLeadingEquals }. OOXML formula1 carries
  // no leading '='; the writer must strip exactly one so the app applies the validation immediately.
  dvFormulaLeadingEquals(formula = '=$AA$1:$AA$2') {
    const wb = new Workbook();
    wb.addWorksheet('S').addDataValidation('A1', {
      type: 'list',
      allowBlank: true,
      formulae: [formula],
    });
    const sheetXml = partMapOf(writeXlsx(wb))['xl/worksheets/sheet1.xml'] || '';
    const formula1 = (sheetXml.match(/<formula1>([\s\S]*?)<\/formula1>/) || [])[1] ?? null;
    return {formula1, hasLeadingEquals: formula1?.startsWith('=')};
  },

  // Attach one validation over a whole range, write, and report the serialized facts → { writeOk,
  // writeError, sqrefs, count, reloadOk }. A range-form validation must emit exactly ONE
  // dataValidation whose sqref is the requested range, not one entry per covered cell.
  roundtripRangeValidation({range, type = 'list', formula = '"a,b,c"'}: Untyped = {}) {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    let buffer: Uint8Array;
    try {
      sheet.addDataValidation(range, {type, allowBlank: true, formulae: [formula]});
      buffer = writeXlsx(wb);
    } catch (e) {
      return {
        writeOk: false,
        writeError: messageOf(e),
        sqrefs: [],
        count: 0,
        reloadOk: null,
      };
    }
    const xml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const sqrefs = [...xml.matchAll(/<dataValidation\b[^>]*sqref="([^"]*)"/g)].map((m) => m[1]);
    const count = [...xml.matchAll(/<dataValidation[ >]/g)].length;
    let reloadOk = true;
    try {
      readXlsx(buffer);
    } catch {
      reloadOk = false;
    }
    return {writeOk: true, writeError: null, sqrefs, count, reloadOk};
  },

  // Author list validations on a 'Main' sheet from the two source forms an author uses: an inline
  // quoted literal ("Male,Female") and a cross-sheet range reference (Levels!$A$2:$A$9999), write,
  // read back, and report both the per-cell rule the reader hands back and the serialized
  // `<dataValidations>` facts (count, well-formedness, the verbatim formula1 texts). Lets a case
  // assert BOTH forms survive a write→read round-trip and that inline lists stay quoted while range
  // references stay unquoted, without the case knowing how validations are shaped internally.
  authorListValidations(validations: Untyped = []) {
    const wb = new Workbook();
    const main = wb.addWorksheet('Main');
    const levels = wb.addWorksheet('Levels');
    levels.getCell('A2').value = 'X';
    for (const v of validations) {
      const rule: Untyped = {
        type: 'list',
        allowBlank: v.allowBlank !== false,
        formulae: [v.formula],
      };
      if (v.error !== undefined) {
        rule.showErrorMessage = true;
        rule.error = v.error;
      }
      main.addDataValidation(v.ref, rule);
    }
    const buffer = writeXlsx(wb);

    const reread = readXlsx(buffer).getWorksheet('Main');
    const readBack: Record<string, Untyped> = {};
    for (const v of validations) {
      const dv = reread?.dataValidationAt(v.ref);
      readBack[v.ref] = dv ? {type: dv.type, formulae: dv.formulae ?? null} : null;
    }

    const xml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const block = (xml.match(/<dataValidations[\s\S]*?<\/dataValidations>/) || [])[0] || '';
    return {
      readBack,
      xml: {
        count: [...xml.matchAll(/<dataValidation[ >]/g)].length,
        // Cheap structural check: a strict consumer chokes on a raw & that is not an entity.
        wellFormed: !/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(block),
        formula1: [...block.matchAll(/<formula1>([\s\S]*?)<\/formula1>/g)].map((m) =>
          m[1]!
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'")
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&amp;/g, '&'),
        ),
      },
    };
  },
};
