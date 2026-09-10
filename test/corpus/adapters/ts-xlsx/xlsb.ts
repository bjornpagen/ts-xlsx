// Reading the binary `.xlsb` serialisation. The headline capability compares the two readings of one
// workbook Excel saved in both forms, which is why the fixture is a *pair*: the XML twin is an
// independent oracle for what the binary must decode to, not something this library produced.

import {strToU8, zipSync} from 'fflate';

import {canonicalJson} from '../../canonical-json.ts';
import type {Untyped} from '../../untyped.ts';
import {partBytesOf, patchedPackage} from './package-facts.ts';
import {encodeAddress, fixtureBytes, readXlsb, readXlsx, type WorkbookInstance} from './runtime.ts';

const FIXTURE = 'xlsb-binary-workbook-reads-like-its-xlsx-twin';
const FORMULAS = 'xlsb-formula-token-streams-decode-to-formula-text';

export const xlsb = {
  // Read the same workbook from its binary and its XML serialisation and report whether the two
  // models agree: sheet names, order and visibility; every cell's value and style facets; row and
  // column geometry; merges. On disagreement, name the first field that differs so a failure is
  // legible without a debugger.
  //
  // Formula text is compared like everything else; only the shared-formula *grouping* is projected
  // away, because the binary form does not record it (see `xlsbFilledFormulaColumn`).
  xlsbModelMatchesXlsxTwin() {
    const binary = snapshot(readXlsx(fixtureBytes(`${FIXTURE}/source.xlsb`)));
    const xml = snapshot(readXlsx(fixtureBytes(`${FIXTURE}/source.xlsx`)));
    if (binary === xml) return {identical: true, firstDifference: null};
    const binaryLines = binary.split('\n');
    const xmlLines = xml.split('\n');
    for (let index = 0; index < Math.max(binaryLines.length, xmlLines.length); index++) {
      if (binaryLines[index] !== xmlLines[index]) {
        return {
          identical: false,
          firstDifference: `xlsb: ${binaryLines[index] ?? '(end)'} | xlsx: ${xmlLines[index] ?? '(end)'}`,
        };
      }
    }
    return {identical: false, firstDifference: 'lengths differ'};
  },

  // The sheets a binary workbook declares, in tab order, with their visibility.
  xlsbSheets() {
    return readXlsb(fixtureBytes(`${FIXTURE}/source.xlsb`)).worksheets.map((sheet) => ({
      name: sheet.name,
      state: sheet.state,
    }));
  },

  // One cell of the binary reading, as JSON-serializable facts: its decoded value (a Date rendered
  // as an ISO string) and the style facets it resolved to.
  xlsbCell(sheetName: string, reference: string) {
    const workbook = readXlsb(fixtureBytes(`${FIXTURE}/source.xlsb`));
    const sheet = workbook.getWorksheet(sheetName);
    if (sheet === undefined) return {found: false};
    const cell = sheet.getCell(reference);
    return {
      found: true,
      value: normalize(cell.value),
      formula: formulaOf(cell.value),
      type: typeof cell.value,
      numFmt: cell.numFmt ?? null,
      font: cell.font ?? null,
      fill: cell.fill ?? null,
      border: cell.border ?? null,
      alignment: cell.alignment ?? null,
      protection: cell.protection ?? null,
    };
  },

  // A sheet's row/column geometry and merged ranges, from the binary reading.
  xlsbGrid(sheetName: string) {
    const workbook = readXlsb(fixtureBytes(`${FIXTURE}/source.xlsb`));
    const sheet = workbook.getWorksheet(sheetName);
    if (sheet === undefined) return {found: false};
    const model = sheet.model;
    return {
      found: true,
      columns: model.columns,
      rows: model.rows,
      merges: model.merges,
    };
  },

  // Every formula cell of the binary reading, compared against the same cell of its XML twin, where
  // the formula is *text* rather than a token stream. Reports each disagreement, so a failure names
  // the token class that broke rather than a count.
  xlsbFormulaTextMatchesXlsxTwin() {
    const binary = formulaTexts(readXlsb(fixtureBytes(`${FORMULAS}/source.xlsb`)));
    const xml = formulaTexts(readXlsx(fixtureBytes(`${FORMULAS}/source.xlsx`)));
    const addresses = [...new Set([...binary.keys(), ...xml.keys()])].sort();
    return {
      compared: xml.size,
      differences: addresses
        .filter((address) => binary.get(address) !== xml.get(address))
        .map((address) => `${address}: xlsb ${binary.get(address)} | xlsx ${xml.get(address)}`),
    };
  },

  // One formula cell of the binary reading: the decoded text and the result Excel cached beside it.
  xlsbFormula(sheetName: string, reference: string) {
    const sheet = readXlsb(fixtureBytes(`${FORMULAS}/source.xlsb`)).getWorksheet(sheetName);
    const value = sheet?.getCell(reference).value;
    return {formula: formulaOf(value), result: normalize(value)};
  },

  // The defined names the binary workbook part declares.
  xlsbDefinedNames() {
    return readXlsb(fixtureBytes(`${FORMULAS}/source.xlsb`)).definedNames;
  },

  // A column filled with one formula, as the binary reading sees it: address, formula text, and
  // whether the cell claims membership of a shared-formula group.
  xlsbFilledFormulaColumn() {
    const sheet = readXlsb(fixtureBytes(`${FORMULAS}/source.xlsb`)).getWorksheet('Calc');
    return ['D1', 'D2', 'D3', 'D4', 'D5'].map((address) => {
      const value = sheet?.getCell(address).value as Untyped;
      return {
        address,
        formula: formulaOf(value),
        shared: value !== null && typeof value === 'object' && 'sharedFormula' in value,
      };
    });
  },

  // Damage one cell's token stream so its last token declares an operand past the end of `rgce`, then
  // read the whole workbook -> { threw, errorName, sheets, formulas, damagedCell, siblingFormula }.
  // The module contract says a formula this decoder cannot read costs the formula and leaves the
  // cached result: that held for an *unrecognised* token and not for a *truncated* one, which threw
  // out of the record reader uncaught and discarded an otherwise fully recoverable workbook. The two
  // are the same damage.
  //
  // The offsets are the fixture's own, and the guard below is what says so: `sheet1.bin`'s first
  // `BrtFmlaNum` record holds `C1`, whose eleven `rgce` bytes end the token stream. Overwriting the
  // last of them with `PtgInt` (0x1E), which reads a two-byte operand that is no longer there, is a
  // truncation expressed without changing a single length field, so the record framing stays exactly
  // as valid as it was and only the formula runs off its own end.
  xlsbTruncatedFormulaToken() {
    const source = fixtureBytes(`${FORMULAS}/source.xlsb`);
    const original = partBytesOf(source, 'xl/worksheets/sheet1.bin');
    const RGCE_LAST = 266;
    if (original[RGCE_LAST] !== 0x03) {
      throw new Error(
        `fixture moved: expected the last rgce byte of C1 at ${RGCE_LAST}, found 0x${(original[RGCE_LAST] ?? 0).toString(16)}`,
      );
    }
    const sheet = Uint8Array.from(original);
    sheet[RGCE_LAST] = 0x1e;
    const archive = patchedPackage(source, {put: {'xl/worksheets/sheet1.bin': sheet}});
    try {
      const workbook = readXlsb(archive);
      const first = workbook.getWorksheet('Calc');
      return {
        threw: false,
        errorName: null,
        sheets: workbook.worksheets.length,
        formulas: countFormulas(workbook),
        damagedCell: normalize(first?.getCell('C1').value),
        damagedFormula: formulaOf(first?.getCell('C1').value),
        siblingFormula: formulaOf(first?.getCell('D1').value),
      };
    } catch (error) {
      const failure = error as Untyped;
      return {
        threw: true,
        errorName: String(failure?.name ?? ''),
        sheets: 0,
        formulas: 0,
        damagedCell: null,
        damagedFormula: null,
        siblingFormula: null,
      };
    }
  },

  // A ZIP whose office document is a binary workbook that does not conform to the record framing:
  // classification succeeded (it *is* an `.xlsb`), so the failure must be a typed parse error rather
  // than a crash, a hang, or a silently empty workbook.
  xlsbMalformedBinaryWorkbook() {
    // A record header claiming a payload far longer than the part holds: the lever a hostile file
    // would pull.
    const workbook = Uint8Array.of(0x81, 0x00, 0x02, 0xff, 0xff, 0xff, 0x7f, 0x01);
    const archive = zipSync({
      '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types/>'),
      'xl/workbook.bin': workbook,
    });
    try {
      readXlsx(archive);
      return {threw: false, errorName: null, message: null};
    } catch (error) {
      const failure = error as Untyped;
      return {
        threw: true,
        errorName: failure?.name ?? null,
        message: String(failure?.message ?? ''),
      };
    }
  },
};

// A Date is not JSON-serializable in a way a case can compare, and a formula cell's *value* is the
// result it computed; both are flattened the same way on either side of a comparison.
function normalize(value: Untyped): Untyped {
  if (value instanceof Date) return value.toISOString();
  if (value !== null && typeof value === 'object' && 'formula' in value) {
    return normalize(value.result ?? null);
  }
  return value;
}

// A cell value as the two serialisations can honestly be compared: formula text and cached result
// both kept, but the shared-formula *grouping* dropped. A spreadsheet fills a formula down a column
// by storing it once and marking the rest as clones; the XML form records that grouping and the
// binary form does not: Excel writes each cell's own formula out in full. So a clone reads back with
// the same formula text either way, and only the `sharedFormula` pointer back to the master differs.
function comparable(value: Untyped): Untyped {
  const formula = formulaOf(value);
  if (formula === null) return normalize(value);
  return {formula, result: normalize(value.result ?? null)};
}

// The formula text a cell carries, or null for a cell that is not a formula.
function formulaOf(value: Untyped): Untyped {
  return value !== null && typeof value === 'object' && typeof value.formula === 'string'
    ? value.formula
    : null;
}

// Every formula cell of a workbook, keyed `Sheet!A1`.
function formulaTexts(workbook: WorkbookInstance): Map<string, string> {
  const texts = new Map<string, string>();
  for (const sheet of workbook.worksheets) {
    for (const cell of sheet.model.cells) {
      const formula = formulaOf(cell.value);
      if (formula !== null)
        texts.set(`${sheet.name}!${encodeAddress(cell.col, cell.row)}`, formula);
    }
  }
  return texts;
}

// A stable, order-independent rendering of everything a worksheet holds. Key order is an artefact of
// which parser filled the object, so it is normalised away; nothing else is.
function snapshot(workbook: WorkbookInstance): string {
  return JSON.stringify(
    canonicalJson(
      workbook.worksheets.map((sheet) => {
        const model = sheet.model;
        return {
          name: sheet.name,
          state: model.state,
          columns: model.columns,
          rows: model.rows,
          merges: model.merges,
          cells: model.cells.map((cell) => ({...cell, value: comparable(cell.value)})),
        };
      }),
    ),
    null,
    1,
  );
}

// How many formula cells a workbook holds, across every sheet: the "and nothing else was lost" half
// of a damaged-formula assertion.
function countFormulas(workbook: WorkbookInstance): number {
  let total = 0;
  for (const sheet of workbook.worksheets) {
    for (const {cells} of sheet.rows()) {
      for (const cell of cells) if (formulaOf(cell.value) !== null) total += 1;
    }
  }
  return total;
}
