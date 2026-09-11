// The CSV reader and writer, including encoding behaviour.

import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {
  normalizeCsvValue,
  specCsvValue,
  translateCsvReadOptions,
  translateCsvWriteOptions,
} from './csv-support.ts';
import {readCsv, Workbook, writeCsv, writeCsvText} from './runtime.ts';

export const csv = {
  csvRead({csv, options}: Untyped = {}) {
    try {
      const wb = readCsv(csv, translateCsvReadOptions(options));
      const sheet = wb.worksheets[0];
      const rows = [];
      if (sheet) {
        for (const {cells} of sheet.rows()) {
          let width = 0;
          for (const cell of cells) if (cell.col > width) width = cell.col;
          const fields = new Array(width).fill(null);
          for (const cell of cells) fields[cell.col - 1] = normalizeCsvValue(cell.value);
          rows.push(fields);
        }
      }
      return {ok: true, error: null, rows};
    } catch (e) {
      return {ok: false, error: messageOf(e), rows: []};
    }
  },

  // Both halves of the codec against the same delimiter -> one row per candidate
  // { delimiter, wroteOk, writeError, text, readOk, readError, rows }. A codec whose two halves
  // disagree about what a delimiter is cannot round-trip its own output, and the disagreement was
  // silent in the worst direction: `field.includes('')` holds for every field, so an empty delimiter
  // quoted every field and produced a file with no separators in it that parses as one column.
  csvDelimiterAgreement() {
    const candidates = [',', ';', '\t', '||', '', '"', '\r', '\n'];
    return candidates.map((delimiter) => {
      const written = this.csvWrite({
        spec: {
          rows: [
            ['a', 'b'],
            ['c', 'd'],
          ],
        },
        options: {formatterOptions: {delimiter}},
      });
      if (!written.ok) {
        return {
          delimiter,
          wroteOk: false,
          writeError: written.error,
          text: null,
          readOk: null,
          readError: null,
          rows: null,
        };
      }
      const back = this.csvRead({csv: written.text, options: {parserOptions: {delimiter}}});
      return {
        delimiter,
        wroteOk: true,
        writeError: null,
        text: written.text,
        readOk: back.ok,
        readError: back.error,
        rows: back.rows,
      };
    });
  },

  csvWrite({spec = {}, options}: Untyped = {}) {
    try {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S');
      for (const row of spec.rows || []) sheet.addRow((row || []).map(specCsvValue));
      const text = writeCsvText(wb, translateCsvWriteOptions(options));
      return {ok: true, error: null, text};
    } catch (e) {
      return {ok: false, error: messageOf(e), text: null};
    }
  },

  // Write a sheet from addressed cells, optionally with formatting-only rows, then read the text back
  // -> { text, cells: {address: value} } over the read-back cells that hold a value. Addresses rather
  // than a row array, because a row array is exactly the shape that cannot show a row moving up.
  csvRowPositionsRoundTrip({
    cells,
    hiddenRows = [],
  }: {
    readonly cells: Readonly<Record<string, string>>;
    readonly hiddenRows?: readonly number[];
  }) {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    for (const [address, value] of Object.entries(cells)) sheet.getCell(address).value = value;
    for (const row of hiddenRows) sheet.getRow(row).hidden = true;
    const text = writeCsvText(wb);
    const back = readCsv(text).worksheets[0];
    const readBack = Object.fromEntries(
      [...(back?.rows() ?? [])].flatMap((row) =>
        row.cells
          .filter((cell) => cell.value !== null)
          .map((cell) => [cell.address, normalizeCsvValue(cell.value)]),
      ),
    );
    return {text, cells: readBack};
  },

  // `undefined` is a case the corpus exercises, not an oversight: it asks the writer to pick a sheet
  // with no selector, so the signature has to admit it.
  csvWriteSheetSelection(sheetName: string | undefined) {
    const wb = new Workbook();
    wb.addWorksheet('First').addRow(['a', 1]);
    const second = wb.addWorksheet('Second');
    second.addRow(['b', 2]);
    second.addRow(['c', 3]);
    let error = null;
    let text = null;
    try {
      text = writeCsvText(wb, sheetName === undefined ? {} : {sheetName});
    } catch (e) {
      error = messageOf(e);
    }
    return {
      ok: error === null,
      error,
      text,
      rowCount: text ? text.split(/\r?\n/).filter(Boolean).length : 0,
    };
  },

  csvReadMapReport() {
    const csv = 'id,amount\n007,32.5\n008,40';
    const read = (map: Untyped) => {
      const wb = readCsv(csv, map ? {map} : {});
      const sheet = wb.worksheets[0];
      const a = sheet ? sheet.getCell('A2').value : null;
      const b = sheet ? sheet.getCell('B2').value : null;
      return {a, aType: typeof a, b, bType: typeof b};
    };
    return {default: read(null), identity: read((v: Untyped) => v)};
  },

  csvWriteEncodingReport({encoding = 'utf16le', text = 'café'}: Untyped = {}) {
    const EMOJI = '😀🎉';
    const CJK = '日本語テスト';

    const fidelityWb = new Workbook();
    fidelityWb.addWorksheet('S').addRow([EMOJI, CJK]);
    const reread = readCsv(writeCsv(fidelityWb)).worksheets[0];
    const emojiRoundtrips =
      !!reread && reread.getCell('A1').value === EMOJI && reread.getCell('B1').value === CJK;

    const encodedWb = new Workbook();
    encodedWb.addWorksheet('S').addRow([text]);
    const encodedBuffer = Buffer.from(writeCsv(encodedWb, {encoding}));
    const decodesAsRequested = encodedBuffer.toString(encoding).replace(/\r?\n$/, '') === text;
    const decodesAsUtf8 = encodedBuffer.toString('utf8').replace(/\r?\n$/, '') === text;

    return {emojiRoundtrips, requestedEncoding: encoding, decodesAsRequested, decodesAsUtf8};
  },

  csvNonAsciiEncodingReport(text = 'שלום') {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('A1').value = text;
    sheet.getCell('B1').value = 'world';
    const buffer = Buffer.from(writeCsv(wb));
    const hasBom = buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf;
    const body = hasBom ? buffer.subarray(3) : buffer;
    return {hasBom, bytesDecodeToText: body.toString('utf8').startsWith(text)};
  },
};
