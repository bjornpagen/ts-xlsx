// Formulas: shared formulas, data tables, and the values a formula cell reports.

import {strToU8} from 'fflate';

import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {partNamesOf, partOf, patchedPackage, roundtrip} from './package-facts.ts';
import {
  encodeAddress,
  fixtureBytes,
  readFixture,
  readXlsx,
  Workbook,
  type WorkbookInstance,
  writeXlsx,
} from './runtime.ts';
import {buildFrom, isoOrNull} from './spec-model.ts';

// A cell value as a case compares it: a formula of any kind as its kind, range, dynamic-array mark,
// text and cached result, and anything else as it is.
function formulaFacts(value: Untyped): Untyped {
  if (value === null || typeof value !== 'object' || typeof value.formula !== 'string')
    return value;
  return {
    kind: value.shareType ?? 'formula',
    ref: value.ref ?? null,
    dynamic: value.dynamic === true,
    formula: value.formula,
    result: value.result ?? null,
  };
}

// Each formula a package's first sheet stores, keyed by cell → {t, ref, dynamic}: the `<f>` element's
// type and range, and whether the cell's `cm` resolves, through the cell metadata part the workbook's
// relationships name, to dynamic-array properties with `fDynamic` set. Read with patterns rather than
// with the library, so a reader and a writer agreeing with each other cannot pass for agreeing with
// the package Excel saved.
function storedFormulas(pkg: Uint8Array) {
  const rels = partOf(pkg, 'xl/_rels/workbook.xml.rels');
  const relationship = /<Relationship [^>]*\/sheetMetadata"[^>]*>/.exec(rels)?.[0] ?? '';
  const target = /Target="([^"]*)"/.exec(relationship)?.[1];
  const path =
    target === undefined ? undefined : target.startsWith('/') ? target.slice(1) : `xl/${target}`;
  const metadata = path !== undefined && partNamesOf(pkg).includes(path) ? partOf(pkg, path) : '';
  const blocks = (body: string | undefined) =>
    [...(body ?? '').matchAll(/<bk>([\s\S]*?)<\/bk>/g)].map((match) => match[1] ?? '');
  const typeNames = [...metadata.matchAll(/<metadataType [^>]*?name="([^"]*)"/g)].map((m) => m[1]);
  const dynamicBlocks = blocks(
    /<futureMetadata name="XLDAPR"[^>]*>([\s\S]*?)<\/futureMetadata>/.exec(metadata)?.[1],
  ).map((block) => /fDynamic="1"/.test(block));
  const cellBlocks = blocks(
    /<cellMetadata[^>]*>([\s\S]*?)<\/cellMetadata>/.exec(metadata)?.[1],
  ).map((block) =>
    [...block.matchAll(/<rc t="(\d+)" v="(\d+)"\/>/g)].map((rc) => ({
      t: Number(rc[1]),
      v: Number(rc[2]),
    })),
  );
  const marked = (cm: string | undefined) =>
    cm !== undefined &&
    (cellBlocks[Number(cm) - 1] ?? []).some(
      ({t, v}) => typeNames[t - 1] === 'XLDAPR' && dynamicBlocks[v] === true,
    );
  const attr = (tag: string, name: string) => new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1];
  const sheet = partOf(pkg, 'xl/worksheets/sheet1.xml');
  return Object.fromEntries(
    [...sheet.matchAll(/<c r="([A-Z]+\d+)"([^>]*)><f(\s[^>]*)?>/g)].map((match) => {
      const f = match[3] ?? '';
      return [
        match[1] ?? '',
        {
          t: attr(f, 't') ?? null,
          ref: attr(f, 'ref') ?? null,
          dynamic: marked(attr(match[2] ?? '', 'cm')),
        },
      ];
    }),
  );
}

export const formulas = {
  // Excel's workbook of array formulas at `rel`, read and written back → {read, excel, written}.
  // `read` maps each non-empty cell of the first sheet to its value as the model holds it, a formula
  // as {kind, ref, dynamic, formula, result}. `excel` and `written` are what the two packages store,
  // as `storedFormulas` reports it.
  arrayFormulaReport(rel: string) {
    const bytes = fixtureBytes(rel);
    const workbook = readXlsx(bytes);
    const cells: Untyped[] = workbook.worksheets[0]?.model.cells ?? [];
    const read = Object.fromEntries(
      cells
        .filter((cell) => cell.value !== null)
        .map((cell) => [encodeAddress(cell.col, cell.row), formulaFacts(cell.value)]),
    );
    return {read, excel: storedFormulas(bytes), written: storedFormulas(writeXlsx(workbook))};
  },

  // A dynamic-array formula over B1:B3 and a Ctrl+Shift+Enter one over D1:D3, authored and written →
  // what the package stores, as `arrayFormulaReport` reports it.
  authoredArrayFormulas() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('B1').value = {
      shareType: 'array',
      formula: 'SEQUENCE(3)',
      ref: 'B1:B3',
      dynamic: true,
    };
    sheet.getCell('D1').value = {shareType: 'array', formula: 'A1:A3*2', ref: 'D1:D3'};
    return storedFormulas(writeXlsx(workbook));
  },

  // Excel's rules for an array formula's range, asked of the library → {spills, authored, edits, foreign}.
  // `spills` maps each edit bringing a formula into a dynamic array filling B1:B3 to {B1, B2, written}:
  // both cells as `arrayFormulaReport` reports a formula, and whether the sheet then writes.
  // `authored` maps each shape to the writer's refusal message, or null when it writes it. `edits` maps
  // each edit of a sheet holding one array formula (B1:B3, or B1:C3 for the column edits) to {refused,
  // ref}: whether the edit threw, and the range the array formula holds after it, null when it is gone.
  // `foreign` is a written package patched so B1:B3 holds a formula in B2, read → {B1, B2, rewrites}: each
  // cell as `arrayFormulaReport` reports a formula, and whether the reading writes back.
  arrayRangeReport() {
    const array = (ref: string, dynamic = false): Untyped => ({
      shareType: 'array',
      formula: 'A1:A3*2',
      ref,
      ...(dynamic ? {dynamic: true} : {}),
    });
    const refusal = (cells: Record<string, Untyped>) => {
      const workbook = new Workbook();
      const sheet = workbook.addWorksheet('S');
      for (const [address, value] of Object.entries(cells)) sheet.getCell(address).value = value;
      try {
        writeXlsx(workbook);
        return null;
      } catch (error) {
        return messageOf(error);
      }
    };
    const edit = (value: Untyped, perform: (sheet: Untyped) => void) => {
      const sheet: Untyped = new Workbook().addWorksheet('S');
      sheet.getCell(String(value.ref).split(':')[0]).value = value;
      let refused = false;
      try {
        perform(sheet);
      } catch {
        refused = true;
      }
      const kept = sheet.model.cells.find((cell: Untyped) => cell.value?.shareType === 'array');
      return {refused, ref: kept?.value.ref ?? null};
    };
    const written = new Workbook();
    const source = written.addWorksheet('S');
    source.getCell('A1').value = 1;
    source.getCell('B1').value = array('B1:B3');
    source.getCell('B2').value = 4;
    const bytes = writeXlsx(written);
    const foreign = patchedPackage(bytes, {
      put: {
        'xl/worksheets/sheet1.xml': strToU8(
          partOf(bytes, 'xl/worksheets/sheet1.xml').replace(
            '<c r="B2"><v>4</v></c>',
            '<c r="B2"><f>A2*10</f></c>',
          ),
        ),
      },
    });
    const read = readXlsx(foreign);
    let rewrites = true;
    try {
      writeXlsx(read);
    } catch {
      rewrites = false;
    }
    const spill = (perform: (sheet: Untyped) => void) => {
      const workbook = new Workbook();
      const sheet: Untyped = workbook.addWorksheet('S');
      for (let row = 1; row <= 3; row++) sheet.getCell(`A${row}`).value = row;
      sheet.getCell('B1').value = array('B1:B3', true);
      perform(sheet);
      let written = true;
      try {
        writeXlsx(workbook);
      } catch {
        written = false;
      }
      return {
        B1: formulaFacts(sheet.getCell('B1').value),
        B2: formulaFacts(sheet.getCell('B2').value),
        written,
      };
    };
    return {
      spills: {
        formulaInserted: spill((sheet) => sheet.insertRow(2, [null, {formula: 'A2*10'}])),
        rowCopied: spill((sheet) => sheet.duplicateRow(1)),
      },
      authored: {
        formulaInsideLegacy: refusal({B1: array('B1:B3'), B2: {formula: 'A2*10'}}),
        formulaInsideDynamic: refusal({B1: array('B1:B3', true), B2: {formula: 'A2*10'}}),
        cornerOverlap: refusal({B1: array('B1:C2'), A2: array('A2:B3')}),
        valuesInside: refusal({B1: array('B1:B3'), B2: 4, B3: 6}),
      },
      edits: {
        insertRowInside: edit(array('B1:B3'), (sheet) => sheet.insertRow(2, [])),
        deleteRowInside: edit(array('B1:B3'), (sheet) => sheet.spliceRows(2, 1)),
        deleteRowsCrossing: edit(array('B1:B3'), (sheet) => sheet.spliceRows(3, 2)),
        duplicateAnchorRow: edit(array('B1:B3'), (sheet) => sheet.duplicateRow(1)),
        insertRowAbove: edit(array('B1:B3'), (sheet) => sheet.insertRow(1, [])),
        insertRowBelow: edit(array('B1:B3'), (sheet) => sheet.insertRow(4, [])),
        deleteWholeRange: edit(array('B1:B3'), (sheet) => sheet.spliceRows(1, 3)),
        insertColumnInside: edit(array('B1:C3'), (sheet) => sheet.insertColumn(3, [])),
        insertColumnLeft: edit(array('B1:C3'), (sheet) => sheet.insertColumn(2, [])),
        deleteColumnPart: edit(array('B1:C3'), (sheet) => sheet.spliceColumns(3, 1)),
        dynamicInsertInside: edit(array('B1:B3', true), (sheet) => sheet.insertRow(2, [])),
      },
      foreign: {
        B1: formulaFacts(read.worksheets[0]?.getCell('B1').value),
        B2: formulaFacts(read.worksheets[0]?.getCell('B2').value),
        rewrites,
      },
    };
  },

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

  // Read a fixture and report the formulas its first sheet carries outside its cells → { read, written,
  // writtenExtension }: the text the model reads, the text the package written back from that model
  // stores, and every `<xm:f>` the written sheet's extension carries, in document order. A key is
  // `cf:<sqref>:<n>` for a conditional format's nth formula, `cf:<sqref>:cfvo<n>` for its nth `formula`
  // anchor, `dv:<sqref>:<n>` for a validation's nth operand, and `table:<name>:<column>:calculated` or
  // `…:totals` for a table column's formulas.
  ruleFormulaSpellings(rel: string) {
    return ruleFormulaSpellingsOf(readFixture(rel));
  },

  // The same report for a workbook authored here: on sheet S1 of S1 and S2, an expression calling
  // XLOOKUP over A1:A3, a colour scale over E1:E3 and a gradient data bar over F1:F3 anchored at MINIFS
  // and MAXIFS, custom validations calling XLOOKUP on D1 and, reaching S2, on D3 (extended), and a table
  // T at G1 whose column `b` is calculated from XLOOKUP and LET and totalled by SUBTOTAL plus XLOOKUP.
  // `read` is the model after a round trip.
  authoredRuleFormulaSpellings() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S1');
    workbook.addWorksheet('S2');
    const bounds = [
      {type: 'formula', value: 'MINIFS($A$1:$A$3,$A$1:$A$3,">0")'},
      {type: 'formula', value: 'MAXIFS($A$1:$A$3,$A$1:$A$3,">0")'},
    ] as const;
    sheet.addConditionalFormatting({
      ref: 'A1:A3',
      rules: [{type: 'expression', formulae: ['XLOOKUP(A1,$B$1:$B$3,$B$1:$B$3)=1']}],
    });
    sheet.addConditionalFormatting({
      ref: 'E1:E3',
      rules: [
        {type: 'colorScale', cfvo: [...bounds], colors: [{argb: 'FFFF0000'}, {argb: 'FF0000FF'}]},
      ],
    });
    sheet.addConditionalFormatting({
      ref: 'F1:F3',
      rules: [{type: 'dataBar', gradient: true, cfvo: [...bounds]}],
    });
    sheet.addDataValidation('D1', {
      type: 'custom',
      formulae: ['ISNUMBER(XLOOKUP(D1,$A$1:$A$3,$A$1:$A$3))'],
    });
    sheet.addDataValidation(
      'D3',
      {type: 'custom', formulae: ['ISNUMBER(XLOOKUP(D3,S2!$A$1:$A$3,S2!$A$1:$A$3))']},
      {extended: true},
    );
    sheet.addTable({
      name: 'T',
      ref: 'G1',
      rowCount: 3,
      totalsRow: true,
      columns: [
        {name: 'a'},
        {
          name: 'b',
          calculatedColumnFormula: 'XLOOKUP(T[[#This Row],[a]],$A$1:$A$3,$B$1:$B$3)+LET(x,1,x)',
          totalsRowFunction: 'custom',
          totalsRowFormula: 'SUBTOTAL(109,T[b])+XLOOKUP(1,$A$1:$A$3,$B$1:$B$3)',
        },
      ],
    });
    const {written, writtenExtension} = ruleFormulaSpellingsOf(workbook);
    return {read: ruleFormulaSpellingsOf(roundtrip(workbook)).read, written, writtenExtension};
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

// The `ruleFormulaSpellings` report for a workbook: its first sheet's formulas outside cells as the
// model holds them, and as the package the workbook writes stores them.
function ruleFormulaSpellingsOf(workbook: WorkbookInstance) {
  const sheet = workbook.worksheets[0];
  const read: (readonly [string, string])[] = [];
  for (const {ref, rules} of sheet?.conditionalFormattings ?? []) {
    const formulae = rules.flatMap((rule) => rule.formulae ?? []);
    formulae.forEach((formula, n) => read.push([`cf:${ref}:${n}`, String(formula)]));
    const anchors = rules
      .flatMap((rule) => rule.cfvo ?? [])
      .filter((cfvo) => cfvo.type === 'formula');
    anchors.forEach((cfvo, n) => read.push([`cf:${ref}:cfvo${n}`, String(cfvo.value)]));
  }
  for (const {sqref, rule} of sheet?.dataValidations ?? []) {
    (rule.formulae ?? []).forEach((formula, n) => read.push([`dv:${sqref}:${n}`, String(formula)]));
  }
  for (const table of sheet?.tables ?? []) {
    for (const column of table.columns) {
      const key = `table:${table.name}:${column.name}`;
      if (column.calculatedColumnFormula !== undefined) {
        read.push([`${key}:calculated`, column.calculatedColumnFormula]);
      }
      if (column.totalsRowFormula !== undefined)
        read.push([`${key}:totals`, column.totalsRowFormula]);
    }
  }

  const buffer = writeXlsx(workbook);
  const sheetXml = partOf(buffer, 'xl/worksheets/sheet1.xml');
  const written: (readonly [string, string])[] = [];
  const numbered = (key: string, texts: Iterable<RegExpMatchArray>) =>
    [...texts].forEach((match, n) => written.push([`${key}${n}`, xmlText(match[1] ?? '')]));
  for (const block of sheetXml.matchAll(
    /<conditionalFormatting sqref="([^"]*)">([\s\S]*?)<\/conditionalFormatting>/g,
  )) {
    const ref = xmlText(block[1] ?? '');
    numbered(`cf:${ref}:`, (block[2] ?? '').matchAll(/<formula>([\s\S]*?)<\/formula>/g));
    numbered(`cf:${ref}:cfvo`, (block[2] ?? '').matchAll(/<cfvo type="formula" val="([^"]*)"\/>/g));
  }
  for (const element of sheetXml.matchAll(
    /<dataValidation\b[^>]*\bsqref="([^"]*)"[^>]*>([\s\S]*?)<\/dataValidation>/g,
  )) {
    numbered(
      `dv:${xmlText(element[1] ?? '')}:`,
      (element[2] ?? '').matchAll(/<formula[12]>([\s\S]*?)<\/formula[12]>/g),
    );
  }
  for (const element of sheetXml.matchAll(
    /<x14:dataValidation\b[^>]*>([\s\S]*?)<\/x14:dataValidation>/g,
  )) {
    const sqref = xmlText(/<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(element[1] ?? '')?.[1] ?? '');
    numbered(
      `dv:${sqref}:`,
      (element[1] ?? '').matchAll(/<x14:formula[12]><xm:f>([\s\S]*?)<\/xm:f><\/x14:formula[12]>/g),
    );
  }
  for (const part of partNamesOf(buffer).filter((name) =>
    /^xl\/tables\/table\d+\.xml$/.test(name),
  )) {
    const tableXml = partOf(buffer, part);
    const name = xmlText(/\bdisplayName="([^"]*)"/.exec(tableXml)?.[1] ?? '');
    for (const column of tableXml.matchAll(
      /<tableColumn\b([^>]*?)(?:\/>|>([\s\S]*?)<\/tableColumn>)/g,
    )) {
      const key = `table:${name}:${xmlText(/\bname="([^"]*)"/.exec(column[1] ?? '')?.[1] ?? '')}`;
      const body = column[2] ?? '';
      const calculated = /<calculatedColumnFormula>([\s\S]*?)<\/calculatedColumnFormula>/.exec(
        body,
      );
      if (calculated !== null) written.push([`${key}:calculated`, xmlText(calculated[1] ?? '')]);
      const totals = /<totalsRowFormula>([\s\S]*?)<\/totalsRowFormula>/.exec(body);
      if (totals !== null) written.push([`${key}:totals`, xmlText(totals[1] ?? '')]);
    }
  }
  const writtenExtension = [...sheetXml.matchAll(/<xm:f>([\s\S]*?)<\/xm:f>/g)].map((match) =>
    xmlText(match[1] ?? ''),
  );
  return {read: Object.fromEntries(read), written: Object.fromEntries(written), writtenExtension};
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
