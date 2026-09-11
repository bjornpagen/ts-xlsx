// Formulas: shared formulas, data tables, and the values a formula cell reports.

import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {partOf, patchedPackage, roundtrip} from './package-facts.ts';
import {
  encodeAddress,
  readFixture,
  readXlsx,
  Workbook,
  type WorkbookInstance,
  writeXlsx,
} from './runtime.ts';
import {buildFrom, isoOrNull} from './spec-model.ts';

export const formulas = {
  // Two readings of a data table whose input cell was deleted → { excel, spliced }. `excel` is the cell
  // Excel itself saved after deleting the input row (`<dir>/input-row-deleted.xlsx`, B3), `spliced` the
  // same table read from `<dir>/before.xlsx` (B4) with row 1 spliced out here and the package written.
  // Each is { ref, r1, r1Deleted, written }, where `written` is the attributes of the `<f>` the cell is
  // written with, as a map.
  dataTableInputDeletionReport(dir: string) {
    const factsOf = (workbook: WorkbookInstance, address: string) => {
      const value: Untyped = workbook.worksheets[0]?.getCell(address).value;
      const xml = partOf(writeXlsx(workbook), 'xl/worksheets/sheet1.xml');
      const tag = new RegExp(`<c r="${address}"[^>]*><f ([^>]*?)/>`).exec(xml)?.[1] ?? '';
      return {
        ref: value?.ref ?? null,
        r1: value?.r1 ?? null,
        r1Deleted: value?.r1Deleted ?? null,
        written: Object.fromEntries([...tag.matchAll(/(\w+)="([^"]*)"/g)].map((m) => [m[1], m[2]])),
      };
    };
    const spliced = readFixture(`${dir}/before.xlsx`);
    spliced.worksheets[0]?.spliceRows(1, 1);
    return {
      excel: factsOf(readFixture(`${dir}/input-row-deleted.xlsx`), 'B3'),
      spliced: factsOf(spliced, 'B3'),
    };
  },

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

  // Author formulas on a workbook of sheets S1, S2 and S3, make one row or column edit of S1, and read
  // each formula back → a map of key → formula text after the edit, spelled with a leading `=` as
  // Excel's `Range.Formula` spells it (empty for a cell holding none). A key is `<sheet>!<cell>` for a
  // cell formula, `name:<name>` for a workbook-scoped defined name, or `dv:<sheet>!<cell>` for a list
  // validation's source on that cell. `read` names where to look after the edit, which for a cell or a
  // validation the edit moved is where it went.
  formulasAfterSplice({
    formulas,
    edit,
    read,
  }: {
    formulas: Record<string, string>;
    edit: {op: 'insert' | 'delete'; axis: 'row' | 'column'; start: number; count: number};
    read: string[];
  }) {
    const workbook = new Workbook();
    for (const name of ['S1', 'S2', 'S3']) workbook.addWorksheet(name);
    const bare = (formula: string) => formula.replace(/^=/, '');
    const place = (key: string) => {
      const [sheet = '', cell = ''] = key.replace(/^dv:/, '').split('!');
      return {sheet: workbook.requireWorksheet(sheet), cell};
    };
    for (const [key, formula] of Object.entries(formulas)) {
      if (key.startsWith('name:')) {
        workbook.defineName({name: key.slice(5), refersTo: bare(formula)});
      } else if (key.startsWith('dv:')) {
        const {sheet, cell} = place(key);
        sheet.addDataValidation(cell, {type: 'list', formulae: [bare(formula)]});
      } else {
        const {sheet, cell} = place(key);
        sheet.getCell(cell).value = {formula: bare(formula)};
      }
    }
    const spliced = workbook.requireWorksheet('S1');
    const inserted = edit.op === 'insert' ? Array.from({length: edit.count}, () => []) : [];
    const removed = edit.op === 'delete' ? edit.count : 0;
    if (edit.axis === 'row') spliced.spliceRows(edit.start, removed, ...inserted);
    else spliced.spliceColumns(edit.start, removed, ...inserted);

    const readBack = (key: string): string => {
      if (key.startsWith('name:')) {
        const refersTo = workbook.definedNames.find((name) => name.name === key.slice(5))?.refersTo;
        return refersTo === undefined ? '' : `=${refersTo}`;
      }
      const {sheet, cell} = place(key);
      if (key.startsWith('dv:')) {
        const source = sheet.dataValidations.find((entry) => entry.sqref === cell)?.rule
          .formulae?.[0];
        return typeof source === 'string' ? `=${source}` : '';
      }
      const value: Untyped = sheet.getCell(cell).value;
      return typeof value?.formula === 'string' ? `=${value.formula}` : '';
    };
    return Object.fromEntries(read.map((key) => [key, readBack(key)]));
  },

  // Author formulas and defined names on a workbook of sheets S1 and S2, each holding 1, 2 and 3 in
  // A1:A3, and report the text each formula is stored as in the written package → a map of key → that
  // text, '' for one the package does not carry. A key is `<sheet>!<cell>` for a cell formula,
  // `name:<name>` for a workbook-level defined name, or `name:<sheet>!<name>` for a name scoped to a
  // sheet.
  storedFormulas(formulas: Record<string, string>) {
    const workbook = new Workbook();
    for (const name of ['S1', 'S2']) {
      const sheet = workbook.addWorksheet(name);
      for (const row of [1, 2, 3]) sheet.getCell(`A${row}`).value = row;
    }
    for (const [key, formula] of Object.entries(formulas)) {
      const target = formulaKey(key);
      if (target.kind === 'name') {
        workbook.defineName({
          name: target.name,
          refersTo: formula,
          ...(target.sheet === undefined ? {} : {scope: target.sheet}),
        });
      } else {
        workbook.requireWorksheet(target.sheet).getCell(target.cell).value = {formula};
      }
    }
    return storedFormulaTexts(writeXlsx(workbook), Object.keys(formulas));
  },

  // Read a fixture, an `.xlsx` or an `.xlsb`, and report every cell formula and defined name it holds,
  // keyed as `storedFormulas` keys them → { read, written }: the formula text the model reads, and the
  // text the package written back from that model stores.
  fixtureFormulaSpellings(rel: string) {
    const workbook = readFixture(rel);
    const cells = workbook.worksheets.flatMap((sheet) =>
      sheet.model.cells.flatMap((cell) => {
        const value: Untyped = cell.value;
        return typeof value?.formula === 'string'
          ? [[`${sheet.name}!${encodeAddress(cell.col, cell.row)}`, value.formula] as const]
          : [];
      }),
    );
    const names = workbook.definedNames.map(
      (name) =>
        [
          name.scope === undefined ? `name:${name.name}` : `name:${name.scope}!${name.name}`,
          name.refersTo,
        ] as const,
    );
    const read = Object.fromEntries([...cells, ...names]);
    return {read, written: storedFormulaTexts(writeXlsx(workbook), Object.keys(read))};
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

type FormulaTarget =
  | {readonly kind: 'cell'; readonly sheet: string; readonly cell: string}
  | {readonly kind: 'name'; readonly name: string; readonly sheet: string | undefined};

// A `storedFormulas` key, taken apart.
function formulaKey(key: string): FormulaTarget {
  if (key.startsWith('name:')) {
    const bang = key.lastIndexOf('!');
    return bang === -1
      ? {kind: 'name', name: key.slice(5), sheet: undefined}
      : {kind: 'name', name: key.slice(bang + 1), sheet: key.slice(5, bang)};
  }
  const [sheet = '', cell = ''] = key.split('!');
  return {kind: 'cell', sheet, cell};
}

// The formula text a written package stores for each key, read off the parts rather than through the
// reader, since the reader is what normalises the spelling away. The writer names the nth sheet's part
// `sheet<n>.xml`, in the order `xl/workbook.xml` lists the sheets.
function storedFormulaTexts(buffer: Uint8Array, keys: readonly string[]) {
  const workbookXml = partOf(buffer, 'xl/workbook.xml');
  const sheetNames = [...workbookXml.matchAll(/<sheet\b[^>]*\bname="([^"]*)"/g)].map((match) =>
    xmlText(match[1] ?? ''),
  );
  const stored = (key: string): string => {
    const target = formulaKey(key);
    if (target.kind === 'name') {
      const scope = target.sheet === undefined ? -1 : sheetNames.indexOf(target.sheet);
      for (const match of workbookXml.matchAll(
        /<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g,
      )) {
        const attrs = match[1] ?? '';
        const name = xmlText(/\bname="([^"]*)"/.exec(attrs)?.[1] ?? '');
        const localSheetId = Number(/\blocalSheetId="(\d+)"/.exec(attrs)?.[1] ?? -1);
        if (name === target.name && localSheetId === scope) return xmlText(match[2] ?? '');
      }
      return '';
    }
    const index = sheetNames.indexOf(target.sheet);
    if (index === -1) return '';
    const sheetXml = partOf(buffer, `xl/worksheets/sheet${index + 1}.xml`);
    const body = new RegExp(`<c r="${target.cell}"[^>]*>([\\s\\S]*?)</c>`).exec(sheetXml)?.[1];
    return xmlText(/<f[^>]*>([\s\S]*?)<\/f>/.exec(body ?? '')?.[1] ?? '');
  };
  return Object.fromEntries(keys.map((key) => [key, stored(key)] as const));
}

// Element text or an attribute value as written, with the five predefined entities decoded; `&amp;`
// last, so an escaped entity decodes to the entity rather than to its character.
function xmlText(text: string): string {
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}
