// Formulas: shared formulas, data tables, and the values a formula cell reports.

import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {partOf, patchedPackage, roundtrip} from './package-facts.ts';
import {readXlsx, Workbook, writeXlsx} from './runtime.ts';
import {buildFrom, isoOrNull} from './spec-model.ts';

export const formulas = {
  // Inject a `<f t="dataTable">` into a written sheet, read it back, and re-write → { reloadOk,
  // readShareType, readRef, readResult, outHasDataTable }. The reader must surface the data-table
  // kind/range/result, and a read-modify-write must re-emit t="dataTable" rather than dropping it.
  dataTableFormulaRoundtrip() {
    const seed = new Workbook();
    const seedSheet = seed.addWorksheet('S');
    seedSheet.getCell('A1').value = 1;
    seedSheet.getCell('B1').value = 2;
    seedSheet.getCell('B2').value = 99;
    const injected = patchedPackage(writeXlsx(seed), {
      edit: {
        'xl/worksheets/sheet1.xml': (xml) =>
          xml.replace(
            /<c r="B2"[^>]*>[\s\S]*?<\/c>/,
            '<c r="B2"><f t="dataTable" ref="B2:B5" dt2D="0" dtr="1" r1="A1"/><v>99</v></c>',
          ),
      },
    });

    let reloadOk: boolean;
    let readShareType = null;
    let readRef = null;
    let readResult = null;
    let outHasDataTable = false;
    try {
      const reload = readXlsx(injected);
      const value = reload.getWorksheet('S')!.getCell('B2').value;
      if (value && typeof value === 'object') {
        // reaches past the value union for the data-table formula's parsed fields
        readShareType = (value as Untyped).shareType ?? null;
        readRef = (value as Untyped).ref ?? null;
        readResult = (value as Untyped).result ?? null;
      }
      reloadOk = true;
      const outXml = partOf(writeXlsx(reload), 'xl/worksheets/sheet1.xml');
      outHasDataTable = /t="dataTable"/.test(outXml);
    } catch {
      reloadOk = false;
    }
    return {readShareType, readRef, readResult, reloadOk, outHasDataTable};
  },

  // Round-trip formula cells whose cached results are truthy and falsy (2, 0, false, '') and report
  // each recovered result → { truthy, zero, boolFalse, emptyString } of { hasResult, result }. A falsy
  // result (0, false, empty string) must survive, not be dropped as if the formula had no cached value.
  formulaFalsyResultReport() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = {formula: '1+1', result: 2};
    sheet.getCell('A2').value = {formula: 'B1-B1', result: 0};
    sheet.getCell('A3').value = {formula: 'FALSE()', result: false};
    sheet.getCell('A4').value = {formula: 'T("")', result: ''};
    const back = roundtrip(workbook).getWorksheet('S')!;
    const probe = (ref: string) => {
      const value = back.getCell(ref).value;
      const hasResult = !!value && typeof value === 'object' && 'result' in value;
      return {hasResult, result: hasResult ? (value as Untyped).result : undefined};
    };
    return {
      truthy: probe('A1'),
      zero: probe('A2'),
      boolFalse: probe('A3'),
      emptyString: probe('A4'),
    };
  },

  // Round-trip a formula whose cached result is a Date → { isValidDate, resultIso, keepsFormula }. The
  // date result reads back as a valid Date (the default date format survives), and the cell stays a
  // formula cell rather than collapsing to a bare value.
  formulaDateResultReport() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = {formula: 'TODAY()', result: new Date(Date.UTC(2021, 0, 2))};
    const value = roundtrip(workbook).getWorksheet('S')!.getCell('A1').value;
    const result = value && typeof value === 'object' ? (value as Untyped).result : undefined;
    const isValidDate = result instanceof Date && !Number.isNaN(result.getTime());
    return {
      isValidDate,
      resultIso: result instanceof Date ? isoOrNull(result) : String(result),
      keepsFormula:
        !!value && typeof value === 'object' && typeof (value as Untyped).formula === 'string',
    };
  },

  // Build a formula-bearing spec, write it, read it back, and report each cell as
  // { formula, sharedFormula, result }, mirroring the oracle. A shared-formula clone reads back a
  // concrete formula (the master's, translated to the clone's address) while retaining its master
  // reference under `sharedFormula`; a plain formula master carries no `sharedFormula`.
  roundtripFormulas(spec: Untyped) {
    const reloaded = roundtrip(buildFrom(spec));
    const out: Record<string, Untyped> = {};
    for (const s of spec.sheets || []) {
      const sheet = reloaded.getWorksheet(s.name);
      for (const c of s.cells || []) {
        const v = sheet ? sheet.getCell(c.ref).value : null;
        const obj = v && typeof v === 'object';
        out[c.ref] = {
          formula: obj && 'formula' in v ? v.formula : null,
          sharedFormula: obj && 'sharedFormula' in v ? v.sharedFormula : null,
          result: obj && 'result' in v ? (v.result ?? null) : null,
        };
      }
    }
    return out;
  },

  // Build a shared-formula sheet (master B1 filled down to B2/B3), then report two things: whether a
  // read → write round-trip preserves the dependents as formula cells, and whether splicing a column
  // into the loaded sheet writes without throwing. The clone's master reference is an address the
  // rewrite does not yet re-anchor on a structural edit, so the splice is the known-open here.
  sharedFormulaRoundtripAndSplice() {
    const build = () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S');
      sheet.getCell('A1').value = 1;
      sheet.getCell('A2').value = 2;
      sheet.getCell('A3').value = 3;
      sheet.getCell('B1').value = {formula: 'A1*2', result: 2};
      sheet.getCell('B2').value = {sharedFormula: 'B1', result: 4};
      sheet.getCell('B3').value = {sharedFormula: 'B1', result: 6};
      return wb;
    };
    const buffer = writeXlsx(build());

    let roundtripError = null;
    let preservedFormulas = null;
    try {
      const reread = readXlsx(buffer);
      writeXlsx(reread);
      const s = reread.getWorksheet('S')!;
      preservedFormulas = ['B2', 'B3'].every((ref) => {
        const v = s.getCell(ref).value;
        return !!(v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v));
      });
    } catch (e) {
      roundtripError = messageOf(e);
    }

    let spliceError = null;
    try {
      const reread = readXlsx(buffer);
      reread.getWorksheet('S')!.spliceColumns(1, 0, []);
      writeXlsx(reread);
    } catch (e) {
      spliceError = messageOf(e);
    }

    return {
      roundtripOk: roundtripError === null,
      roundtripError,
      preservedFormulas,
      spliceOk: spliceError === null,
      spliceError,
    };
  },
};
