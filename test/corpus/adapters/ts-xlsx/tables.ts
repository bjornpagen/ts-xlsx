// Worksheet tables: their columns, styles, display names, and what editing one does to the
// cells underneath it.

import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {type PartMap, partMapOf, patchedPackage, roundtrip} from './package-facts.ts';
import {
  fixtureBytes,
  readFixture,
  readXlsx,
  Workbook,
  type WorksheetInstance,
  writeXlsx,
} from './runtime.ts';
import {buildFrom} from './spec-model.ts';

export const tables = {
  // Find a table by name across a loaded fixture's sheets and report its column names and data-row
  // count. The reader reconstructs the table from its part, deriving the data-row count from the
  // stored range (height minus the header and totals rows), so a loaded table exposes its rows.
  readFixtureTable(rel: string, tableName: string) {
    const wb = readFixture(rel);
    for (const s of wb.worksheets) {
      const table = s.tables.find((t) => t.name === tableName);
      if (table)
        return {
          found: true,
          columns: table.columns.map((c) => c.name),
          rowCount: table.options.rowCount,
        };
    }
    return {found: false, columns: null, rowCount: null};
  },

  // Load a fixture and report a named table's column count and names, used to prove a table with a
  // calculated column (a <calculatedColumnFormula> child of the column) does not truncate the column
  // list or crash the read.
  loadFixtureTableColumns(rel: string, tableName: string) {
    try {
      const wb = readFixture(rel);
      for (const s of wb.worksheets) {
        const table = s.tables.find((t) => t.name === tableName);
        if (table) {
          return {
            loaded: true,
            error: null,
            columnCount: table.columns.length,
            columnNames: table.columns.map((c) => c.name),
          };
        }
      }
      return {loaded: true, error: null, columnCount: 0, columnNames: []};
    } catch (e) {
      return {
        loaded: false,
        error: messageOf(e),
        columnCount: null,
        columnNames: null,
      };
    }
  },

  // Build a table-bearing spec, report the full ref written into each table part, then read the
  // package back and re-write it, reporting the ref and well-formedness of each re-emitted part, so
  // a degenerate (empty-body or single-row) table is proven to survive a load→save round-trip.
  roundtripSpecTableFacts(spec: Untyped) {
    const tableFacts = (parts: PartMap) =>
      Object.keys(parts)
        .filter((n) => /^xl\/tables\/table\d+\.xml$/.test(n))
        .sort((a, b) => Number(a.match(/\d+/)![0]) - Number(b.match(/\d+/)![0]))
        .map((n) => parts[n] ?? '')
        .map((xml) => ({
          ref: (xml.match(/<table\b[^>]*\bref="([^"]*)"/) || [])[1] ?? null,
          wellFormed: !/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(xml),
        }));
    const write = tableFacts(partMapOf(writeXlsx(buildFrom(spec))));
    let loadOk = true;
    let loadError = null;
    let rewritten: Untyped[] = [];
    try {
      rewritten = tableFacts(partMapOf(writeXlsx(roundtrip(buildFrom(spec)))));
    } catch (e) {
      loadOk = false;
      loadError = messageOf(e);
    }
    return {write, roundtrip: rewritten, loadOk, loadError};
  },

  // Author a five-column table, round-trip it, and report the loaded column count and names: the
  // reader must expose every column in order, not truncate to a fixed cap.
  wideTableColumnReadReport() {
    const wb = new Workbook();
    wb.addWorksheet('S').addTable({
      name: 'Wide',
      ref: 'A1',
      columns: [{name: 'C1'}, {name: 'C2'}, {name: 'C3'}, {name: 'C4'}, {name: 'C5'}],
      rowCount: 2,
    });
    const table = roundtrip(wb).getWorksheet('S')!.tables[0]!;
    return {colCount: table.columns.length, colNames: table.columns.map((c) => c.name)};
  },

  // Add a table and a list validation to each of five sheets, then report that the package writes
  // with unique table part ids, reloads with every table present, and keeps the first sheet's
  // validation: the loop over many sheets must not collide table ids or strip validations.
  multiSheetTableReport() {
    const wb = new Workbook();
    for (let i = 1; i <= 5; i++) {
      const s = wb.addWorksheet(`Sheet${i}`);
      s.addTable({name: `Table_${i}`, ref: 'A1', columns: [{name: 'Col'}], rowCount: 1});
      s.addDataValidation('C1', {type: 'list', allowBlank: true, formulae: ['"a,b,c"']});
    }
    let writeOk = true;
    let writeError = null;
    let buffer: Uint8Array | null = null;
    try {
      buffer = writeXlsx(wb);
    } catch (e) {
      writeOk = false;
      writeError = messageOf(e);
    }
    if (!writeOk)
      return {
        writeOk,
        writeError,
        reloadOk: false,
        tableCount: null,
        idsUnique: false,
        firstSheetDvSurvives: false,
      };

    const parts = partMapOf(buffer!);
    const ids = Object.keys(parts)
      .filter((n) => /^xl\/tables\/table\d+\.xml$/.test(n))
      .map((n) => (parts[n]!.match(/<table\b[^>]*\bid="([^"]*)"/) || [])[1]);
    const idsUnique =
      ids.length > 0 && ids.every((x) => x != null) && new Set(ids).size === ids.length;

    let reloadOk = true;
    let tableCount = 0;
    let firstSheetDvSurvives = false;
    try {
      const back = readXlsx(buffer!);
      tableCount = back.worksheets.reduce((n, s) => n + s.tables.length, 0);
      firstSheetDvSurvives = !!back.getWorksheet('Sheet1')?.dataValidationAt('C1');
    } catch {
      reloadOk = false;
    }
    return {writeOk, writeError, reloadOk, tableCount, idsUnique, firstSheetDvSurvives};
  },

  // Write three tables, one with a real built-in style name, one with no style, one with the
  // sentinel theme "None", and report the tableStyleInfo name attribute (or null when absent) each
  // emits, plus whether the "None" table kept its showRowStripes flag. Theme "None" must mean an
  // unstyled table (no name attribute), not a bogus name="None" referencing a non-existent style.
  tableStyleThemeReport() {
    const styleInfoOf = (style: Untyped) => {
      const wb = new Workbook();
      wb.addWorksheet('S').addTable({
        name: 'T',
        ref: 'A1',
        columns: [{name: 'A'}],
        rowCount: 1,
        style,
      });
      let ok = true;
      let tag: string | null;
      try {
        const parts = partMapOf(writeXlsx(wb));
        const part = Object.keys(parts).find((n) => /^xl\/tables\/table\d+\.xml$/.test(n));
        tag = (parts[part!]!.match(/<tableStyleInfo[^>]*\/?>/) || [])[0] ?? null;
      } catch (e) {
        ok = false;
        tag = messageOf(e);
      }
      const name = tag && ok ? ((tag.match(/\bname="([^"]*)"/) || [])[1] ?? null) : null;
      const hasStripes = !!(tag && ok && /\bshowRowStripes="1"/.test(tag));
      return {ok, name, hasStripes, tag};
    };
    return {
      real: styleInfoOf({name: 'TableStyleMedium2'}),
      // An explicit style object with no name is OOXML's "unstyled", distinct from omitting the
      // style entirely, which a freshly-authored table fills with Excel's default (TableStyleMedium2).
      nullTheme: styleInfoOf({}),
      none: styleInfoOf({name: 'None', showRowStripes: true}),
    };
  },

  // Author a table from a header-name list (which may contain collisions), write it, and report the
  // column names emitted into the table part plus whether they are unique. OOXML requires unique
  // tableColumn names; colliding inputs must be disambiguated deterministically, not written verbatim
  // into a corrupt file → { ok, writtenNames, uniqueNames }.
  tableDuplicateColumnNamesReport(headers: string[]) {
    const wb = new Workbook();
    let ok = true;
    let writtenNames: string | (string | undefined)[];
    try {
      wb.addWorksheet('S').addTable({
        name: 'T',
        ref: 'A1',
        columns: headers.map((name: Untyped) => ({name})),
        rowCount: 1,
      });
      const parts = partMapOf(writeXlsx(wb));
      const part = Object.keys(parts).find((n) => /^xl\/tables\/table\d+\.xml$/.test(n));
      writtenNames = [...parts[part!]!.matchAll(/<tableColumn\b[^>]*\bname="([^"]*)"/g)].map(
        (m) => m[1],
      );
    } catch (e) {
      ok = false;
      writtenNames = messageOf(e);
    }
    const uniqueNames =
      Array.isArray(writtenNames) &&
      new Set(writtenNames.map((n) => n!.toLowerCase())).size === writtenNames.length;
    return {ok, writtenNames, uniqueNames};
  },

  // Author a two-column table whose first column carries a numFmt style, append two data rows, then
  // round-trip and report the numFmt read back on each column's body cells → { writeOk, reloadOk,
  // writeError, styledBody, unstyledBody }. The per-column style must bake into the styled column's
  // body cells and leave the unstyled column untouched.
  tableColumnStyleReport(numFmt: string) {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    const table = s.addTable({
      name: 'T',
      ref: 'A1',
      columns: [{name: 'Amount', style: {numFmt}}, {name: 'Label'}],
      rowCount: 0,
    });
    s.getCell('A1').value = 'Amount';
    s.getCell('B1').value = 'Label';
    table.addRow([1234.5, 'x']);
    table.addRow([6789, 'y']);

    let writeOk = true;
    let writeError = null;
    let buffer: Uint8Array | null = null;
    try {
      buffer = writeXlsx(wb);
    } catch (e) {
      writeOk = false;
      writeError = messageOf(e);
    }
    if (!writeOk)
      return {writeOk, writeError, reloadOk: false, styledBody: null, unstyledBody: null};

    let reloadOk = true;
    let styledBody = null;
    let unstyledBody = null;
    try {
      const back = readXlsx(buffer!).getWorksheet('S')!;
      styledBody = [back.getCell('A2').numFmt ?? null, back.getCell('A3').numFmt ?? null];
      unstyledBody = [back.getCell('B2').numFmt ?? null, back.getCell('B3').numFmt ?? null];
    } catch (e) {
      reloadOk = false;
      writeError = messageOf(e);
    }
    return {writeOk, writeError, reloadOk, styledBody, unstyledBody};
  },

  // Build a table-bearing spec, round-trip it through a write→read, fetch the named table on the
  // reloaded model, append the requested rows, then re-write and re-read to report the final row
  // count. A table read from a file must expose its data rows and accept appends exactly like a
  // freshly-created one → { hasTable, loadedRowCount, addError, committed, finalRowCount }.
  roundtripTableAppend(spec: Untyped, {tableName, appendRows}: Untyped) {
    const reloaded = roundtrip(buildFrom(spec));
    let table: Untyped = null;
    for (const s of reloaded.worksheets) {
      const found = s.getTable(tableName);
      if (found) {
        table = found;
        break;
      }
    }
    const hasTable = table !== null;
    if (!hasTable) {
      return {
        hasTable,
        loadedRowCount: null,
        addError: null,
        committed: false,
        finalRowCount: null,
      };
    }
    const loadedRowCount = table.rowCount;

    let addError = null;
    for (const row of appendRows) {
      try {
        table.addRow(row);
      } catch (e) {
        addError = messageOf(e);
        break;
      }
    }

    let committed = false;
    let finalRowCount = null;
    if (addError === null) {
      try {
        const out = writeXlsx(reloaded);
        committed = true;
        const back = readXlsx(out);
        for (const s of back.worksheets) {
          const found = s.getTable(tableName);
          if (found) {
            finalRowCount = found.rowCount;
            break;
          }
        }
      } catch (e) {
        addError = messageOf(e);
      }
    }
    return {hasTable, loadedRowCount, addError, committed, finalRowCount};
  },

  // Author a table over A1:B3, populate its cells, load the package, edit a body cell (B2 → 999),
  // and re-write, reporting that both writes and the reload succeed, the table part and its unique
  // worksheet relationship survive, and the edited value reads back. Editing a cell inside a table's
  // range must not truncate or corrupt the table part or its rels.
  tableCellEditRoundtrip() {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    s.addTable({name: 'T', ref: 'A1', columns: [{name: 'H1'}, {name: 'H2'}], rowCount: 2});
    s.getCell('A1').value = 'H1';
    s.getCell('B1').value = 'H2';
    s.getCell('A2').value = 'a';
    s.getCell('B2').value = 1;
    s.getCell('A3').value = 'b';
    s.getCell('B3').value = 2;

    let writeOk = true;
    let writeError = null;
    let firstBuffer: Uint8Array | null = null;
    try {
      firstBuffer = writeXlsx(wb);
    } catch (e) {
      writeOk = false;
      writeError = messageOf(e);
    }
    if (!writeOk) {
      return {
        writeOk,
        writeError,
        reloadOk: false,
        hasTablePart: false,
        tablePresent: false,
        editedValue: null,
        relUnique: false,
      };
    }

    let reloadOk = true;
    let hasTablePart = false;
    let tablePresent = false;
    let editedValue = null;
    let relUnique = false;
    try {
      const reloaded = readXlsx(firstBuffer!);
      const sheet = reloaded.getWorksheet('S')!;
      sheet.getCell('B2').value = 999;
      const out = writeXlsx(reloaded);
      const parts = partMapOf(out);
      const tablePart = Object.keys(parts).find((n) => /^xl\/tables\/table\d+\.xml$/.test(n));
      hasTablePart = tablePart !== undefined;
      const relPart = Object.keys(parts).find((n) =>
        /xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(n),
      );
      const relIds = relPart ? [...parts[relPart]!.matchAll(/Id="([^"]*)"/g)].map((m) => m[1]) : [];
      relUnique = relIds.length > 0 && new Set(relIds).size === relIds.length;
      const back = readXlsx(out);
      const backSheet = back.getWorksheet('S')!;
      tablePresent = backSheet.tables.some((t) => t.name === 'T');
      editedValue = backSheet.getCell('B2').value;
    } catch (e) {
      reloadOk = false;
      writeError = messageOf(e);
    }
    return {writeOk, writeError, reloadOk, hasTablePart, tablePresent, editedValue, relUnique};
  },

  // Write a table whose first column name embeds CR/LF line breaks, then report the first
  // <tableColumn> tag and whether it carries a raw (unescaped) CR or LF. A raw control char in an
  // attribute value is not preserved by XML normalisation (a CR reparses as a space) and makes the
  // package suspect, so the name must be emitted XML-escaped (&#13;&#10;), not raw.
  tableColumnNameControlChars() {
    const wb = new Workbook();
    wb.addWorksheet('S').addTable({
      name: 'T',
      ref: 'A1',
      columns: [{name: 'Test\r\nmultiple\r\nlines'}, {name: 'Plain'}],
      rowCount: 1,
    });
    let writeOk = true;
    let writeError = null;
    let firstColumnTag = null;
    let rawControlChars = null;
    try {
      const parts = partMapOf(writeXlsx(wb));
      const part = Object.keys(parts).find((n) => /^xl\/tables\/table\d+\.xml$/.test(n));
      firstColumnTag = (parts[part!]!.match(/<tableColumn\b[^>]*\/?>/) || [])[0] ?? null;
      rawControlChars = firstColumnTag === null ? null : /[\r\n]/.test(firstColumnTag);
    } catch (e) {
      writeOk = false;
      writeError = messageOf(e);
    }
    return {writeOk, writeError, firstColumnTag, rawControlChars};
  },

  // Read every table part of a fixture, then load→save it and read the re-emitted parts, reporting
  // each table's autoFilter / header-row / totals-row / column-count facts before and after, and each
  // column's calculated column formula as a map of column name → formula text (columns without one
  // are absent). A no-op round-trip of a table that has no autoFilter must not inject one, flip the
  // header row off, or turn totalsRowShown on; a table that does have one must keep its ref and column
  // count.
  roundtripFixtureTableXml(rel: string) {
    const facts = (xml: string) => ({
      hasAutoFilter: /<(?:\w+:)?autoFilter\b/.test(xml),
      autoFilterRef: (xml.match(/<(?:\w+:)?autoFilter\b[^>]*\bref="([^"]*)"/) || [])[1] ?? null,
      headerRowCount: (xml.match(/\bheaderRowCount="([^"]*)"/) || [])[1] ?? null,
      totalsRowShown: (xml.match(/\btotalsRowShown="([^"]*)"/) || [])[1] ?? null,
      columnCount: (xml.match(/<tableColumns\b[^>]*\bcount="([^"]*)"/) || [])[1] ?? null,
      calculatedColumnFormulas: Object.fromEntries(
        [
          ...xml.matchAll(
            /<tableColumn\b[^>]*\bname="([^"]*)"[^>]*>\s*<calculatedColumnFormula\b[^>]*>([\s\S]*?)<\/calculatedColumnFormula>/g,
          ),
        ].map((match) => [match[1] ?? '', match[2] ?? '']),
      ),
    });
    const tablePartsInOrder = (parts: PartMap) =>
      Object.keys(parts)
        .filter((n) => /^xl\/tables\/table\d+\.xml$/.test(n))
        .sort((a, b) => Number(a.match(/\d+/)![0]) - Number(b.match(/\d+/)![0]))
        .map((n) => parts[n] ?? '');
    const buffer = fixtureBytes(rel);
    const source = tablePartsInOrder(partMapOf(buffer));
    const rewritten = tablePartsInOrder(partMapOf(writeXlsx(readXlsx(buffer))));
    return {
      tables: source.map((xml, i) => {
        const after = rewritten[i];
        return {
          name: (xml.match(/<table\b[^>]*\bname="([^"]*)"/) || [])[1] ?? null,
          source: facts(xml),
          rewritten: after === undefined ? null : facts(after),
        };
      }),
    };
  },

  // A table's two name attributes, both ways → { written, readDiffering, readNameOnly }. `written` is
  // the `name` and `displayName` of the part a table authored as `Sales` is written with. The other two
  // are the model name read back from that part patched to say `name="Internal" displayName="Shown"`,
  // and patched to carry `name="OnlyName"` and no `displayName` at all.
  tableNamesReport() {
    const workbook = new Workbook();
    workbook
      .addWorksheet('S')
      .addTable({name: 'Sales', ref: 'A1', columns: [{name: 'h'}], rowCount: 1});
    const bytes = writeXlsx(workbook);
    const PART = 'xl/tables/table1.xml';
    const part = partMapOf(bytes)[PART] ?? '';
    const attribute = (name: string) =>
      new RegExp(`<table [^>]* ${name}="([^"]*)"`).exec(part)?.[1] ?? null;
    const readAs = (patch: (xml: string) => string) =>
      readXlsx(patchedPackage(bytes, {edit: {[PART]: patch}})).getWorksheet('S')?.tables[0]?.name ??
      null;
    return {
      written: {name: attribute('name'), displayName: attribute('displayName')},
      readDiffering: readAs((xml) =>
        xml.replace('name="Sales" displayName="Sales"', 'name="Internal" displayName="Shown"'),
      ),
      readNameOnly: readAs((xml) =>
        xml.replace('name="Sales" displayName="Sales"', 'name="OnlyName"'),
      ),
    };
  },
  // Splice columns against a two-column table and report what is left → { wholeDelete, leftInsert,
  // rightEdge }. `shiftRows` asks whether the deleted span swallowed the table and returns false so
  // the caller prunes it; `shiftColumns` did neither that nor the clamp every other coordinate gets,
  // so a table survived a splice that deleted its every column (still naming columns that no longer
  // exist, declared over whatever slid left, and emitted that way), and an insert could push its
  // anchor past the last column, where `range` throws on read.
  tableThroughColumnSplice() {
    const build = () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S');
      sheet.getCell('B1').value = 'h1';
      sheet.getCell('C1').value = 'h2';
      sheet.getCell('B2').value = 1;
      sheet.getCell('C2').value = 2;
      sheet.addTable({name: 'T', ref: 'B1', columns: [{name: 'h1'}, {name: 'h2'}], rowCount: 1});
      return sheet;
    };

    const deleted = build();
    deleted.spliceColumns(2, 2);

    const shifted = build();
    shifted.spliceColumns(1, 0, ['inserted']);

    // A table anchored on the very last column, where an unclamped anchor would leave the grid.
    const edgeWb = new Workbook();
    const edge = edgeWb.addWorksheet('E');
    edge.getCell('XFD1').value = 'h';
    edge.getCell('XFD2').value = 1;
    edge.addTable({name: 'T', ref: 'XFD1', columns: [{name: 'h'}], rowCount: 1});
    edge.spliceColumns(1, 0, ['inserted']);

    return {
      wholeDelete: {tableCount: deleted.tables.length},
      leftInsert: {tableCount: shifted.tables.length, range: shifted.tables[0]?.range ?? null},
      // Reported as data either way: reading `range` on an anchor pushed off the grid throws, and a
      // thrown range is exactly the failure the clamp exists to prevent.
      rightEdge: {range: readRange(edge)},
    };
  },
  // Ask a table to occupy a rectangle the grid does not have, four ways → { pastRightEdge,
  // pastBottomEdge, onTheEdge, splicedPastRightEdge, writtenRef }. A table's anchor was validated and
  // its far corner, derived from the anchor plus the column count and the row count, was not. So a
  // table could be declared over columns past XFD or rows past 1048576: `range` then threw when
  // anything read it back, and a `rowCount` large enough (which a reader derives from a file's own
  // stored range, not only an author) put a `<table ref>` naming impossible rows into the package.
  // Author a one-column table under each name and report what construction made of it → a map of
  // name → { refused, error }. The error is the refusal's message, or null for a name accepted.
  tableNameRefusals(names: string[]) {
    return Object.fromEntries(
      names.map((name) => {
        try {
          new Workbook()
            .addWorksheet('S')
            .addTable({name, ref: 'A1', columns: [{name: 'h'}], rowCount: 1});
          return [name, {refused: false, error: null}];
        } catch (error) {
          return [name, {refused: true, error: messageOf(error)}];
        }
      }),
    );
  },

  tableOutsideTheGrid() {
    const refused = (build: () => void) => {
      try {
        build();
        return null;
      } catch (error) {
        return messageOf(error);
      }
    };
    const sheetWith = (name: string) => new Workbook().addWorksheet(name);

    const pastRightEdge = refused(() => {
      sheetWith('A').addTable({
        name: 'T',
        ref: 'XFC1',
        columns: [{name: 'a'}, {name: 'b'}, {name: 'c'}],
        rowCount: 1,
      });
    });
    const pastBottomEdge = refused(() => {
      sheetWith('B').addTable({name: 'T', ref: 'A1', columns: [{name: 'a'}], rowCount: 5_000_000});
    });

    // Two columns anchored at XFC reach XFD exactly: the last legal rectangle, which must be accepted
    // and must round-trip. Inserting a column to its left then has nowhere to put its right edge.
    const edgeWb = new Workbook();
    const edge = edgeWb.addWorksheet('S');
    edge.getCell('XFC1').value = 'a';
    edge.getCell('XFD1').value = 'b';
    edge.addTable({name: 'T', ref: 'XFC1', columns: [{name: 'a'}, {name: 'b'}], rowCount: 1});
    const onTheEdge = readRange(edge);
    const writtenRef =
      /<table\b[^>]*\bref="([^"]*)"/.exec(
        partMapOf(writeXlsx(edgeWb))['xl/tables/table1.xml'] ?? '',
      )?.[1] ?? null;

    edge.spliceColumns(1, 0, ['inserted']);
    const splicedPastRightEdge = {tableCount: edge.tables.length, range: readRange(edge)};

    return {pastRightEdge, pastBottomEdge, onTheEdge, writtenRef, splicedPastRightEdge};
  },

  // Read a fixture that already carries preserved pivot parts, author a NEW pivot onto it, write, and
  // report the package → { pivotTableParts, cacheDefinitionParts, cacheRecordsParts, duplicateOverrides,
  // authoredPartChanged, reloadOk, reloadError }. The writer generates a pivot table part and both cache
  // parts numbered globally from 1, while a preserved pivot kept its ORIGINAL path on the theory that
  // the writer never generates one, so both landed on `pivotTable1.xml`. The preserved bytes were
  // emitted last and won: the authored pivot's part was overwritten, its sheet relationship pointed at
  // the old pivot's data, and the content types declared the same PartName twice, which violates OPC
  // M2.5 and makes Excel repair the package.
  pivotPreservedAndAuthoredCoexist(rel: string) {
    const wb = readFixture(rel);
    const source = wb.addWorksheet('PivotSrc');
    source.getCell('A1').value = 'Region';
    source.getCell('B1').value = 'Quarter';
    source.getCell('C1').value = 'Amount';
    source.getCell('A2').value = 'West';
    source.getCell('B2').value = 'Q1';
    source.getCell('C2').value = 10;
    source.getCell('A3').value = 'East';
    source.getCell('B3').value = 'Q2';
    source.getCell('C3').value = 20;
    wb.addWorksheet('NewPivot').addPivotTable({
      source,
      rows: ['Region'],
      columns: ['Quarter'],
      values: ['Amount'],
    });

    const sourceParts = partMapOf(fixtureBytes(rel));
    const written = partMapOf(writeXlsx(wb));
    const names = Object.keys(written);
    const matching = (pattern: RegExp) => names.filter((name) => pattern.test(name)).sort();

    const overrides = [
      ...(written['[Content_Types].xml'] ?? '').matchAll(/<Override PartName="([^"]*)"/g),
    ].map((match) => match[1] ?? '');

    return {
      pivotTableParts: matching(/^xl\/pivotTables\/pivotTable\d+\.xml$/),
      cacheDefinitionParts: matching(/^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/),
      cacheRecordsParts: matching(/^xl\/pivotCache\/pivotCacheRecords\d+\.xml$/),
      duplicateOverrides: overrides.filter((name, i) => overrides.indexOf(name) !== i).sort(),
      // The generated pivot takes number 1, so this part must NOT be the fixture's own bytes.
      authoredPartChanged:
        written['xl/pivotTables/pivotTable1.xml'] !== sourceParts['xl/pivotTables/pivotTable1.xml'],
      sourcePivotTableParts: Object.keys(sourceParts)
        .filter((name) => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(name))
        .sort(),
    };
  },

  // Author a pivot over `source` (header → the five data values, null for a missing one) plus a numeric
  // `Amount` value field, with the named fields on each axis, and report → a map of header → the
  // attributes of its `<sharedItems>` in the written cache definition, in written order, `count` included
  // for a field that carries a catalogue. A field the source did not produce is absent.
  pivotCacheFieldKinds(spec: {
    source: Readonly<Record<string, readonly (string | number | null)[]>>;
    rows: readonly string[];
    columns: readonly string[];
  }) {
    const wb = new Workbook();
    const source = wb.addWorksheet('Data');
    const headers = Object.keys(spec.source);
    source.addRow([...headers, 'Amount']);
    for (let row = 0; row < 5; row++) {
      source.addRow([...headers.map((header) => spec.source[header]?.[row] ?? null), row + 1]);
    }
    wb.addWorksheet('Pivot').addPivotTable({
      source,
      rows: spec.rows,
      columns: spec.columns,
      values: ['Amount'],
    });
    const cache = partMapOf(writeXlsx(wb))['xl/pivotCache/pivotCacheDefinition1.xml'] ?? '';
    return Object.fromEntries(
      [...cache.matchAll(/<cacheField name="([^"]*)"[^>]*><sharedItems ?([^>]*?)\/?>/g)].map(
        (match) => [match[1] ?? '', match[2] ?? ''],
      ),
    );
  },

  // Author a pivot over a small source whose header row is `headers`, naming fields for each role, and
  // report → { refusal, pivotFields, rowFields, columnFields, dataFields }. A refused pivot carries
  // `refusal: {code, message}` and nulls elsewhere; a written one carries `refusal: null`, each
  // `<pivotField>` as `{axis, dataField}`, and the row, column and data fields as indices into the cache
  // fields. Data rows fill every header column with a short string or number.
  pivotFieldWiring(spec: {
    headers: readonly string[];
    rows: readonly string[];
    columns: readonly string[];
    values: readonly string[];
  }) {
    const wb = new Workbook();
    const source = wb.addWorksheet('Data');
    const data = [
      ['a', 'x', 1],
      ['b', 'y', 2],
      ['a', 'y', 3],
    ];
    spec.headers.forEach((header, col) => {
      const letter = String.fromCharCode(65 + col);
      source.getCell(`${letter}1`).value = header;
      data.forEach((row, i) => {
        source.getCell(`${letter}${i + 2}`).value = row[col % row.length] ?? null;
      });
    });
    try {
      wb.addWorksheet('Pivot').addPivotTable({
        source,
        rows: spec.rows,
        columns: spec.columns,
        values: spec.values,
      });
    } catch (e) {
      return {
        refusal: {code: String((e as Untyped)?.code ?? ''), message: messageOf(e)},
        pivotFields: null,
        rowFields: null,
        columnFields: null,
        dataFields: null,
      };
    }
    const xml = partMapOf(writeXlsx(wb))['xl/pivotTables/pivotTable1.xml'] ?? '';
    const indicesIn = (element: string, child: string, attribute: string) =>
      [
        ...(new RegExp(`<${element}\\b[^>]*>(.*?)</${element}>`).exec(xml)?.[1] ?? '').matchAll(
          new RegExp(`<${child}\\b[^>]*\\b${attribute}="(\\d+)"`, 'g'),
        ),
      ].map((match) => Number(match[1]));
    return {
      refusal: null,
      pivotFields: [...xml.matchAll(/<pivotField\b([^>]*)>/g)].map((match) => ({
        axis: /\baxis="([^"]*)"/.exec(match[1] ?? '')?.[1] ?? null,
        dataField: /\bdataField="(1|true)"/.test(match[1] ?? ''),
      })),
      rowFields: indicesIn('rowFields', 'field', 'x'),
      columnFields: indicesIn('colFields', 'field', 'x'),
      dataFields: indicesIn('dataFields', 'dataField', 'fld'),
    };
  },
};

// A table's range, or the message reading it threw. An anchor moved past the last column makes
// `range`, `autoFilterRef` and `region` all throw, so a case has to be able to see that as a value.
function readRange(sheet: WorksheetInstance): string | null {
  try {
    return sheet.tables[0]?.range ?? null;
  } catch (e) {
    return messageOf(e);
  }
}
