// Hyperlinks, including the internal (same-workbook) form and how it serializes.

import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {partMapOf} from './package-facts.ts';
import {readFixture, readXlsx, Workbook, writeXlsx} from './runtime.ts';
import {attrsOf, reloadPatched} from './xml-probes.ts';

export const hyperlinks = {
  // Read a fixture and report every hyperlink on its first sheet as { <ref>: { hyperlink, text } }:
  // the link's target, and the value of the cell at the top-left of the cells it covers, with a rich
  // label flattened to its concatenated text. A foreign file's links, and the rejoining of an external
  // URL's fragment carried in the location attribute, must be read faithfully.
  async readFixtureHyperlinks(rel: string) {
    const flatten = (t: Untyped) =>
      t == null
        ? null
        : typeof t === 'string'
          ? t
          : Array.isArray(t.richText)
            ? t.richText.map((r: Untyped) => r.text).join('')
            : t;
    const sheet = readFixture(rel).worksheets[0];
    const out: Record<string, Untyped> = {};
    for (const link of sheet?.hyperlinks ?? []) {
      const anchor = link.ref.split(':')[0] ?? link.ref;
      out[link.ref] = {hyperlink: link.target, text: flatten(sheet?.getCell(anchor).value)};
    }
    return out;
  },

  // Build a workbook with one internal ('#'-prefixed) hyperlink, write it, and report how the link
  // serialized → { writeOk, hasLocation, location, hasExternalRel, hasRid, reloadOk }. An internal
  // target must ride in a `location` attribute with no external-mode relationship, and the package
  // must reload.
  async internalHyperlinkReport(target = "#'Target'!A1") {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Main');
    wb.addWorksheet('Target');
    sheet.getCell('A1').value = 'go';
    sheet.addHyperlink({ref: 'A1', target});
    let buffer: Uint8Array;
    try {
      buffer = writeXlsx(wb);
    } catch (e) {
      return {writeOk: false, writeError: messageOf(e)};
    }
    const parts = partMapOf(buffer);
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const relsXml = parts['xl/worksheets/_rels/sheet1.xml.rels'] || '';
    const a = attrsOf((sheetXml.match(/<hyperlink\b[^>]*\/?>/) || [''])[0]);
    let reloadOk = true;
    try {
      readXlsx(buffer);
    } catch {
      reloadOk = false;
    }
    return {
      writeOk: true,
      hasLocation: a.location != null,
      location: a.location ?? null,
      hasExternalRel: /TargetMode="External"/.test(relsXml),
      hasRid: a['r:id'] != null,
      reloadOk,
    };
  },

  // Build a workbook with an internal '#Sheet2!A1' hyperlink (plus a tooltip), write it, and report
  // the serialized distinctions → { hasWorksheetRels, hyperlinkHasRid, hyperlinkLocation,
  // relTargetMode, reReadHyperlink }. The internal form must carry a location and NO external
  // relationship, and the target must survive a reload.
  async internalHyperlinkSerializationReport() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Sheet1');
    wb.addWorksheet('Sheet2');
    ws.getCell('A1').value = 'go';
    ws.addHyperlink({ref: 'A1', target: '#Sheet2!A1', tooltip: 'tt'});
    const buffer = writeXlsx(wb);
    const parts = partMapOf(buffer);
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const relsXml = parts['xl/worksheets/_rels/sheet1.xml.rels'] || '';
    const hyperlinkEl = (sheetXml.match(/<hyperlink\b[^>]*\/?>/) || [''])[0];
    const relEl = (relsXml.match(/<Relationship\b[^>]*hyperlink[^>]*\/?>/) || [''])[0];
    const reReadHyperlink =
      readXlsx(buffer).getWorksheet('Sheet1')?.hyperlinkAt('A1')?.target ?? null;
    return {
      hasWorksheetRels: /Type="[^"]*\/hyperlink"/.test(relsXml),
      hyperlinkHasRid: /r:id="/.test(hyperlinkEl),
      hyperlinkLocation: (hyperlinkEl.match(/location="([^"]*)"/) || [null, null])[1],
      relTargetMode: (relEl.match(/TargetMode="([^"]*)"/) || [null, null])[1],
      reReadHyperlink,
    };
  },

  // Write a sheet holding a number (A1), a boolean (B1), a formula (C1), a date (D1), a text label (E1)
  // and an empty F1, patch an in-workbook hyperlink onto each of the six, read it, and write and read
  // it again → { read, rewritten }, each { cells, links }: `cells` maps the six refs to { kind, value },
  // `kind` being 'formula', 'date', 'null' or the value's `typeof`, and `links` lists each hyperlink the
  // sheet holds as { ref, target }.
  hyperlinkOverNonTextCellsReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('A1').value = 42;
    sheet.getCell('B1').value = true;
    sheet.getCell('C1').value = {formula: '1+1', result: 2};
    sheet.getCell('D1').value = new Date(Date.UTC(2024, 0, 15));
    sheet.getCell('E1').value = 'label';
    const refs = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1'];
    const links = refs.map((ref) => `<hyperlink ref="${ref}" location="S!H1"/>`).join('');
    const report = (workbook: ReturnType<typeof readXlsx>) => {
      const s = workbook.getWorksheet('S')!;
      return {
        cells: Object.fromEntries(refs.map((ref) => [ref, describeValue(s.getCell(ref).value)])),
        links: s.hyperlinks.map(({ref, target}) => ({ref, target})),
      };
    };
    const read = reloadPatched(writeXlsx(wb), {
      'xl/worksheets/sheet1.xml': (xml) =>
        xml.replace('</sheetData>', `</sheetData><hyperlinks>${links}</hyperlinks>`),
    });
    return {read: report(read), rewritten: report(readXlsx(writeXlsx(read)))};
  },

  // Read a fixture and report its first sheet's links and the value of each named cell, as read and
  // after a write and a second read → { read, rewritten }, each { cells, links }: `cells` maps each
  // named cell to { kind, value } as `hyperlinkOverNonTextCellsReport` describes one (a formula's Date
  // result as an ISO string), `links` lists every link as { ref, target, tooltip } (null when none).
  fixtureHyperlinksReport(rel: string, refs: string[]) {
    const report = (workbook: ReturnType<typeof readXlsx>) => {
      const sheet = workbook.worksheets[0];
      return {
        cells: Object.fromEntries(
          refs.map((ref) => [ref, describeValue(sheet?.getCell(ref).value ?? null)]),
        ),
        links: (sheet?.hyperlinks ?? []).map(({ref, target, tooltip}) => ({
          ref,
          target,
          tooltip: tooltip ?? null,
        })),
      };
    };
    const workbook = readFixture(rel);
    return {read: report(workbook), rewritten: report(readXlsx(writeXlsx(workbook)))};
  },

  // Put links through the edits Excel 16.0 was seen to make, recorded in
  // test/corpus/fixtures/excel-oracle/hyperlinks-through-edits.json → { overlap, splice, copyAndColumn }.
  // `overlap` is { refs, b1, c1 } after a link over A1:C1 and then one over B1 are written and read
  // back: the links' refs in order and the target each of B1 and C1 opens. `splice` is the refs of a
  // link over A2:A4 after a row is inserted at 3 (`afterInsertInside`), then row 2 is deleted
  // (`afterTopRowDelete`), then rows 2 to 4 are (`afterWholeDelete`). `copyAndColumn` is the refs of a
  // link on A1 after row 1 is duplicated once (`afterRowCopy`) and then a column is inserted before A
  // (`afterColumnInsert`).
  hyperlinksThroughEdits() {
    const refs = (sheet: {hyperlinks: readonly {ref: string}[]}) =>
      sheet.hyperlinks.map((link) => link.ref);

    const layered = new Workbook();
    const top = layered.addWorksheet('S');
    top.addHyperlink({ref: 'A1:C1', target: 'https://one.example/'});
    top.addHyperlink({ref: 'B1', target: 'https://two.example/'});
    const reread = readXlsx(writeXlsx(layered)).getWorksheet('S')!;

    const spliced = new Workbook().addWorksheet('S');
    spliced.addHyperlink({ref: 'A2:A4', target: 'https://range.example/'});
    spliced.spliceRows(3, 0, []);
    const afterInsertInside = refs(spliced);
    spliced.spliceRows(2, 1);
    const afterTopRowDelete = refs(spliced);
    spliced.spliceRows(2, 3);
    const afterWholeDelete = refs(spliced);

    const copied = new Workbook().addWorksheet('S');
    copied.getCell('A1').value = 7;
    copied.addHyperlink({ref: 'A1', target: 'https://copy.example/'});
    copied.duplicateRow(1);
    const afterRowCopy = refs(copied);
    copied.spliceColumns(1, 0, []);

    return {
      overlap: {
        refs: refs(reread),
        b1: reread.hyperlinkAt('B1')?.target ?? null,
        c1: reread.hyperlinkAt('C1')?.target ?? null,
      },
      splice: {afterInsertInside, afterTopRowDelete, afterWholeDelete},
      copyAndColumn: {afterRowCopy, afterColumnInsert: refs(copied)},
    };
  },

  // Put a link over merged D1:H1, whose top-left cell holds its label, insert a row above, write and
  // read it back → { merges, linkRef, link }: the merged ranges read back, the `ref` of the written
  // `<hyperlink>` (or null), and the link covering D2 as { hyperlink, text, range } (its target, D2's
  // value, and the cells it covers), or null when none covers D2.
  rangedHyperlinkAfterInsertRow() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('D1').value = 'go';
    sheet.addHyperlink({ref: 'D1:H1', target: 'https://example.com/'});
    sheet.mergeCells('D1:H1');
    sheet.insertRow(1, ['header']);
    const bytes = writeXlsx(wb);
    const element = (partMapOf(bytes)['xl/worksheets/sheet1.xml'] ?? '').match(
      /<hyperlink\b[^>]*\/?>/,
    )?.[0];
    const read = readXlsx(bytes).getWorksheet('S');
    const link = read?.hyperlinkAt('D2');
    return {
      merges: read === undefined ? [] : [...read.merges],
      linkRef: element === undefined ? null : (attrsOf(element).ref ?? null),
      link:
        link === undefined
          ? null
          : {hyperlink: link.target, text: read?.getCell('D2').value ?? null, range: link.ref},
    };
  },
};

// A cell value as { kind, value }: `kind` is 'date', 'formula', 'null' or the value's `typeof`, and a
// Date, a formula's Date result included, is an ISO string so the report stays plain JSON.
function describeValue(value: Untyped) {
  const plain = (v: Untyped) => (v instanceof Date ? v.toISOString() : v);
  if (value instanceof Date) return {kind: 'date', value: value.toISOString()};
  if (value && typeof value === 'object' && 'formula' in value) {
    return {kind: 'formula', value: {formula: value.formula, result: plain(value.result ?? null)}};
  }
  return {kind: value === null ? 'null' : typeof value, value: value ?? null};
}
