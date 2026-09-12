// Core model and whole-package behaviour: address decoding, reader input classification,
// package inspection, and the workbook-level round-trips that are not about one feature.

import {canonicalJson} from '../../canonical-json.ts';
import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {packageFacts} from '../ooxml-facts.ts';
import {packagePartFacts, partMapOf, roundtrip} from './package-facts.ts';
import {
  type CellInstance,
  decodeAddress,
  decodeRange,
  detectValueType,
  encodeAddress,
  fixtureBytes,
  readFixture,
  readWorkbookStream,
  readXlsx,
  Workbook,
  type WorkbookInstance,
  writeXlsx,
  XlsxError,
} from './runtime.ts';
import {
  buildFrom,
  isoOrNull,
  normalizeRewriteCell,
  normalizeStreamValue,
  UnsupportedSpecError,
} from './spec-model.ts';
import {
  buildReadInput,
  classifyReadError,
  mainNamespacePrefixReport,
  type ReadInputKind,
  relationshipPrefixReport,
  reloadPatched,
  reloadWithMarkupSubstituted,
  themePrefixReport,
  selfClosingDefinedNameReport,
} from './xml-probes.ts';

// Each error cell a package's first sheet stores, keyed by address → {v, errorType}: the `<v>` text of
// the `t="e"` cell, and the `errorType` its `vm` names, through the value metadata and the rich-value
// parts, or null when it points at none. Read with patterns rather than with the library, so a reader
// and a writer agreeing with each other cannot pass for agreeing with the package Excel saved.
function storedErrors(pkg: Uint8Array) {
  const parts = partMapOf(pkg);
  const metadata = parts['xl/metadata.xml'] ?? '';
  const typeNames = [...metadata.matchAll(/<metadataType [^>]*?name="([^"]*)"/g)].map((m) => m[1]);
  const richBlocks = /<futureMetadata name="XLRICHVALUE"[^>]*>([\s\S]*?)<\/futureMetadata>/.exec(
    metadata,
  )?.[1];
  const richValueOfBlock = [...(richBlocks ?? '').matchAll(/<xlrd:rvb i="(\d+)"\/>/g)].map((m) =>
    Number(m[1]),
  );
  const valueBlocks = [
    ...(/<valueMetadata[^>]*>([\s\S]*?)<\/valueMetadata>/.exec(metadata)?.[1] ?? '').matchAll(
      /<bk><rc t="(\d+)" v="(\d+)"\/><\/bk>/g,
    ),
  ].map((m) => ({type: Number(m[1]), value: Number(m[2])}));
  const structureKeys = [
    ...(parts['xl/richData/rdrichvaluestructure.xml'] ?? '').matchAll(/<s [^>]*>([\s\S]*?)<\/s>/g),
  ].map((m) => [...(m[1] ?? '').matchAll(/<k n="([^"]*)"/g)].map((k) => k[1]));
  const richValues = [
    ...(parts['xl/richData/rdrichvalue.xml'] ?? '').matchAll(/<rv s="(\d+)">([\s\S]*?)<\/rv>/g),
  ].map((m) => ({
    structure: Number(m[1]),
    values: [...(m[2] ?? '').matchAll(/<v>([^<]*)<\/v>/g)].map((v) => v[1]),
  }));
  const errorTypeOf = (vm: string | undefined): number | null => {
    const block = vm === undefined ? undefined : valueBlocks[Number(vm) - 1];
    if (block === undefined || typeNames[block.type - 1] !== 'XLRICHVALUE') return null;
    const rich = richValues[richValueOfBlock[block.value] ?? -1];
    const key =
      rich === undefined ? -1 : (structureKeys[rich.structure] ?? []).indexOf('errorType');
    return rich === undefined || key < 0 ? null : Number(rich.values[key]);
  };
  const sheet = parts['xl/worksheets/sheet1.xml'] ?? '';
  return Object.fromEntries(
    [...sheet.matchAll(/<c r="([A-Z]+\d+)"([^>]*?)(?<!\/)>([\s\S]*?)<\/c>/g)]
      .filter((m) => /\st="e"/.test(m[2] ?? ''))
      .map((m) => [
        m[1] ?? '',
        {
          v: /<v>([^<]*)<\/v>/.exec(m[3] ?? '')?.[1] ?? null,
          errorType: errorTypeOf(/\svm="(\d+)"/.exec(m[2] ?? '')?.[1]),
        },
      ]),
  );
}

// The characters XML 1.0 cannot carry, restated here rather than imported: a case asserts on
// behaviour, and reading the rule out of the code under test would make the assertion circular.
// oxlint-disable-next-line eslint/no-control-regex -- naming the control characters is the point
const UNREPRESENTABLE = /[\u{0}-\u{8}\u{B}\u{C}\u{E}-\u{1F}\u{FFFE}\u{FFFF}\u{D800}-\u{DFFF}]/u;

export const core = {
  // Classify a reader input by format family and report the typed error (or success) it produces:
  // `{threw, errorName, code, format, message, leaksZipInternals, leaksAbsolutePath}`. A
  // non-`.xlsx` blob must fail with a clear, catchable, typed error rather than a raw zip crash.
  classifyReadInput(kind: ReadInputKind) {
    return classifyReadError(() => {
      readXlsx(buildReadInput(kind));
    });
  },

  // Author one sheet carrying all five kinds of content the reader takes out of a worksheet part (a
  // hyperlink, a standard data validation, an extended x14 validation, a classic conditional format,
  // and a data bar whose gradient and negative-fill facets can only live in the x14 extension), write
  // it, and report what comes back → { hyperlink, cellText, standardValidation, extendedValidation,
  // classicRuleType, dataBarGradient, dataBarNegativeFill }. The reader takes all five from a single
  // parse of the part, so this is where one kind's handler swallowing another kind's events would
  // show: every fact below comes from a different reader over the same event stream.
  worksheetPartCarryingEveryKind() {
    const workbook = new Workbook();
    const other = workbook.addWorksheet('Other');
    other.getCell('A1').value = 'listed';
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 'go';
    sheet.addHyperlink({ref: 'A1', target: 'https://example.com/'});
    for (let row = 1; row <= 3; row++) sheet.getCell(`D${row}`).value = row / 3;
    sheet.addDataValidation('B1:B5', {type: 'list', formulae: ['"one,two,three"']});
    sheet.addDataValidation(
      'C1:C5',
      {type: 'list', formulae: ['Other!$A$1:$A$1']},
      {extended: true},
    );
    sheet.addConditionalFormatting({
      ref: 'D1:D3',
      rules: [{type: 'cellIs', operator: 'greaterThan', formulae: ['0.5'], priority: 1}],
    });
    sheet.addConditionalFormatting({
      ref: 'E1:E3',
      rules: [
        {
          type: 'dataBar',
          priority: 2,
          gradient: false,
          color: {argb: 'FF638EC6'},
          negativeFillColor: {argb: 'FFFF0000'},
          cfvo: [
            {type: 'num', value: 0},
            {type: 'num', value: 1},
          ],
        },
      ],
    });

    const back = roundtrip(workbook).getWorksheet('S')!;
    const value = back.getCell('A1').value as Untyped;
    const validations = back.dataValidations as Untyped[];
    const rules = (back.conditionalFormattings as Untyped[]).flatMap(
      (cf: Untyped) => cf.rules as Untyped[],
    );
    const dataBar = rules.find((rule: Untyped) => rule.type === 'dataBar');
    return {
      hyperlink: back.hyperlinkAt('A1')?.target ?? null,
      cellText: value ?? null,
      standardValidation:
        validations.find((entry) => entry.sqref === 'B1:B5' && entry.extended !== true)?.rule
          ?.formulae?.[0] ?? null,
      extendedValidation:
        validations.find((entry) => entry.extended === true)?.rule?.formulae?.[0] ?? null,
      classicRuleType: rules.find((rule: Untyped) => rule.type === 'cellIs')?.type ?? null,
      dataBarGradient: dataBar?.gradient ?? null,
      dataBarNegativeFill: dataBar?.negativeFillColor?.argb ?? null,
    };
  },

  decodeAddress(reference: string) {
    return decodeAddress(reference);
  },

  decodeRange(reference: string) {
    return decodeRange(reference);
  },

  inspectPackage(spec: Untyped) {
    return packageFacts(spec, partMapOf(writeXlsx(buildFrom(spec))));
  },

  // Same package facts as `inspectPackage`, but after a full write→read→write cycle: the spec is
  // written, loaded back into a fresh model, and re-emitted. Lets a case assert that content the
  // writer materializes (e.g. a table's totals row) survives a round-trip unchanged, neither dropped
  // on read nor duplicated/clobbered when the reloaded model is written again.
  roundtripInspectPackage(spec: Untyped) {
    return packageFacts(spec, partMapOf(writeXlsx(roundtrip(buildFrom(spec)))));
  },

  // Author a pivot table over source data containing XML-special characters (& < > " ') and a
  // missing field value, write, and report whether the emitted pivotCacheDefinition is well-formed
  // and free of raw unescaped ampersands. Mirrors the oracle's shape → { ok, writeError,
  // cacheWellFormed, hasRawUnescapedAmp }.
  pivotCacheSpecialCharsReport() {
    try {
      const wb = new Workbook();
      const src = wb.addWorksheet('Data');
      src.getCell('A1').value = 'Name';
      src.getCell('B1').value = 'Region';
      src.getCell('C1').value = 'Amount';
      src.getCell('A2').value = 'Smith & Co';
      src.getCell('B2').value = '<West>';
      src.getCell('C2').value = 10;
      src.getCell('B3').value = 'East';
      src.getCell('C3').value = 20;
      src.getCell('A4').value = 'It\'s "best"';
      src.getCell('B4').value = 'West';
      src.getCell('C4').value = 30;
      wb.addWorksheet('Pivot').addPivotTable({
        source: src,
        rows: ['Name'],
        columns: ['Region'],
        values: ['Amount'],
        metric: 'sum',
      });
      const parts = partMapOf(writeXlsx(wb));
      const key = Object.keys(parts).find((n) => /pivotCacheDefinition\d*\.xml$/.test(n));
      const cacheXml = key ? parts[key]! : '';
      const hasRawUnescapedAmp = /&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[0-9a-fA-F]+;)/.test(
        cacheXml,
      );
      return {
        ok: true,
        writeError: null,
        cacheWellFormed: cacheXml ? !hasRawUnescapedAmp : false,
        hasRawUnescapedAmp,
      };
    } catch (e) {
      return {
        ok: false,
        writeError: messageOf(e),
        cacheWellFormed: null,
        hasRawUnescapedAmp: null,
      };
    }
  },

  // Set an autofilter over a range, write, and report the sheet's autoFilter ref plus whether the
  // workbook declares the hidden, sheet-scoped `_xlnm._FilterDatabase` defined name portable consumers
  // (LibreOffice) rely on → { autoFilterRef, hasFilterDatabase, filterDatabaseHidden,
  // filterDatabaseFormula }. Mirrors the oracle's shape.
  autoFilterDefinedNameReport(ref = 'A1:B2') {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 'H1';
    sheet.getCell('B1').value = 'H2';
    sheet.getCell('A2').value = 1;
    sheet.getCell('B2').value = 2;
    sheet.autoFilter = ref;
    const parts = partMapOf(writeXlsx(workbook));
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const wbXml = parts['xl/workbook.xml'] || '';
    const autoFilterRef = (sheetXml.match(/<autoFilter\b[^>]*ref="([^"]*)"/) || [])[1] ?? null;
    const filterDb = wbXml.match(
      /<definedName\b([^>]*)name="_xlnm._FilterDatabase"([^>]*)>([\s\S]*?)<\/definedName>/,
    );
    return {
      autoFilterRef,
      hasFilterDatabase: !!filterDb,
      filterDatabaseHidden: filterDb ? /hidden="1"/.test(filterDb[1]! + filterDb[2]!) : false,
      filterDatabaseFormula: filterDb ? filterDb[3] : null,
    };
  },

  // Write one string cell with the buffered writer's useSharedStrings option and report how it was
  // stored → { hasSharedStringsPart, isSharedRef, isInline }. The option must actually control
  // storage: enabled emits a sharedStrings part and a t="s" cell reference; disabled keeps the string
  // inline with no such part.
  sharedStringsOption(useSharedStrings: boolean) {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 'shared-me';
    const parts = partMapOf(writeXlsx(wb, {useSharedStrings}));
    const sheet = parts['xl/worksheets/sheet1.xml'] || '';
    return {
      hasSharedStringsPart: 'xl/sharedStrings.xml' in parts,
      isSharedRef: /t="s"><v>\d+<\/v>/.test(sheet),
      isInline: /t="inlineStr"><is><t>shared-me<\/t>/.test(sheet),
    };
  },

  // Read a fixture, write it back unchanged, and report package-part facts before/after →
  // { source, rewritten }. Use it to assert a no-op round-trip PRESERVES parts the reader does
  // not model (a vector-shape drawing, a header/footer image and its VML) instead of dropping them.
  roundtripFixturePackageParts(rel: string) {
    const source = packagePartFacts(partMapOf(fixtureBytes(rel)));
    const rewritten = packagePartFacts(partMapOf(writeXlsx(readXlsx(fixtureBytes(rel)))));
    return {source, rewritten};
  },

  // Author each code as a cell's error value, write, and read back → a map of code → { written,
  // errorType, readBack, refused }. `written` is the `<v>` of the typed cell, `errorType` the one its `vm`
  // names through the rich-value parts (null for none), `readBack` the value the package reads as, and
  // `refused` the writer's message when it declined to write the cell at all.
  errorLiteralReport(codes: string[]) {
    return Object.fromEntries(
      codes.map((code) => {
        const workbook = new Workbook();
        // Untyped on purpose: the case names codes the model may not accept, to see what the writer does.
        workbook.addWorksheet('S').getCell('A1').value = {error: code} as Untyped;
        try {
          const bytes = writeXlsx(workbook);
          const stored = storedErrors(bytes).A1;
          return [
            code,
            {
              written: stored?.v ?? null,
              errorType: stored?.errorType ?? null,
              readBack: readXlsx(bytes).worksheets[0]?.getCell('A1').value ?? null,
              refused: null,
            },
          ];
        } catch (error) {
          return [
            code,
            {written: null, errorType: null, readBack: null, refused: messageOf(error)},
          ];
        }
      }),
    );
  },

  // Excel's workbook of computed and pasted errors at `rel`, read and written back → {read, excel,
  // written}. `read` maps each cell of the first sheet holding an error, as its value or as its formula's
  // cached result, to that error's code; `excel` and `written` are what the two packages store for each
  // error cell, as {v, errorType}.
  richValueErrorReport(rel: string) {
    const bytes = fixtureBytes(rel);
    const workbook = readXlsx(bytes);
    const cells: Untyped[] = workbook.worksheets[0]?.model.cells ?? [];
    const read = Object.fromEntries(
      cells
        .map((cell) => [
          encodeAddress(cell.col, cell.row),
          cell.value?.error ?? cell.value?.result?.error ?? null,
        ])
        .filter(([, error]) => error !== null),
    );
    return {read, excel: storedErrors(bytes), written: storedErrors(writeXlsx(workbook))};
  },

  // Write a non-finite numeric cell (NaN / Infinity / -Infinity) and report whether the sheet XML
  // carries a bare token in a <v> → { hasNonFiniteToken, token }. A non-finite value has no OOXML
  // representation, so it must serialize as a valueless cell, never a literal "NaN"/"Infinity".
  nonFiniteCellReport(kind: 'NaN' | 'Infinity' | '-Infinity') {
    const value =
      kind === 'NaN'
        ? Number.NaN
        : kind === '-Infinity'
          ? Number.NEGATIVE_INFINITY
          : Number.POSITIVE_INFINITY;
    const workbook = new Workbook();
    workbook.addWorksheet('S').getCell('A1').value = value;
    const sheetXml = partMapOf(writeXlsx(workbook))['xl/worksheets/sheet1.xml'] || '';
    const token = (sheetXml.match(/<c r="A1"[^>]*>\s*<v>([\s\S]*?)<\/v>/) || [])[1] ?? null;
    return {hasNonFiniteToken: /<v>[^<]*(NaN|Infinity)[^<]*<\/v>/.test(sheetXml), token};
  },

  // Author into cell text, a cached string formula result, or a sheet name a string that XML 1.0
  // cannot carry verbatim, meaning a C0 control, a noncharacter or a lone surrogate, and report what
  // the writer did → { writeOk, writeError, partsWithRawChar, emittedText }. Cell values have the
  // SpreadsheetML `_xHHHH_` convention and must use it; a sheet name has none, so a refusal is the
  // only honest outcome there. Neither may put the raw character into an emitted part, which would
  // make the package malformed XML. `readValue` closes the loop: reading the package back must give
  // the character the author asked for, since the escape and the unescape are inverses.
  xmlCharacterSafetyReport(where: 'cell-text' | 'formula-result' | 'sheet-name', text: string) {
    // oxlint-disable-next-line eslint/no-control-regex -- matching the control characters is the check: this asks whether an emitted part carries one
    const raw = /[\u{0}-\u{8}\u{B}\u{C}\u{E}-\u{1F}\u{FFFE}\u{FFFF}\u{D800}-\u{DFFF}]/u;
    let writeOk = true;
    let writeError: string | null = null;
    let partsWithRawChar: string[] = [];
    let emittedText: string | null = null;
    let readValue: unknown = null;
    try {
      const workbook = new Workbook();
      const sheet = workbook.addWorksheet(where === 'sheet-name' ? text : 'S');
      if (where === 'cell-text') sheet.getCell('A1').value = text;
      if (where === 'formula-result') sheet.getCell('A1').value = {formula: 'B1', result: text};
      const bytes = writeXlsx(workbook);
      const reloaded = readXlsx(bytes).worksheets[0]?.getCell('A1').value;
      readValue =
        reloaded !== null && typeof reloaded === 'object' && 'result' in reloaded
          ? reloaded.result
          : reloaded;
      const parts = partMapOf(bytes);
      partsWithRawChar = Object.keys(parts)
        .filter((name) => raw.test(parts[name] ?? ''))
        .sort();
      const cellXml = (parts['xl/worksheets/sheet1.xml'] ?? '').match(/<c r="A1"[\s\S]*?<\/c>/);
      const body =
        cellXml === null ? null : /<t[^>]*>([\s\S]*?)<\/t>|<v>([\s\S]*?)<\/v>/.exec(cellXml[0]);
      emittedText = body === null ? null : (body[1] ?? body[2] ?? null);
    } catch (error) {
      writeOk = false;
      writeError = messageOf(error);
    }
    return {writeOk, writeError, partsWithRawChar, emittedText, readValue};
  },

  // Put an entity reference no DTD defines into a cell's text and report what the reader makes of
  // each → a map of entity name → { value, type }. An OOXML package carries no DTD, so there is
  // nothing an unrecognised `&name;` could expand to and the only safe reading is the literal text.
  // The names worth asking about are the ones a JavaScript object answers for free.
  unknownEntityReport(names: string[]) {
    const marker = 'ENTITYSLOT';
    return Object.fromEntries(
      names.map((name) => {
        const workbook = new Workbook();
        workbook.addWorksheet('S').getCell('A1').value = marker;
        const reloaded = reloadWithMarkupSubstituted(writeXlsx(workbook), marker, `&${name};`);
        const value = reloaded.worksheets[0]?.getCell('A1').value ?? null;
        return [name, {value, type: typeof value}];
      }),
    );
  },

  // Read a fixture whose cells hold `_xHHHH_` escapes, three ways → { eager, streaming, roundtrip }.
  // Each is a map of A1 reference → the cell's decoded text (a formula cell reports its cached
  // result). `eager` and `streaming` must agree, since a file read row-by-row must decode identically
  // to the same file read whole; `roundtrip` reads back what our own writer re-emitted, which is what
  // holds the escape and the unescape to being inverses of each other.
  escapeDecodeReport(rel: string, refs: string[]) {
    const textOf = (value: unknown): string | null => {
      if (typeof value === 'string') return value;
      if (value !== null && typeof value === 'object' && 'result' in value) {
        const {result} = value;
        return typeof result === 'string' ? result : null;
      }
      return null;
    };
    const eagerOf = (workbook: WorkbookInstance) => {
      const sheet = workbook.worksheets[0];
      return Object.fromEntries(
        refs.map((ref) => [ref, sheet === undefined ? null : textOf(sheet.getCell(ref).value)]),
      );
    };
    const streaming: Record<string, string | null> = Object.fromEntries(
      refs.map((ref) => [ref, null]),
    );
    for (const sheet of readWorkbookStream(fixtureBytes(rel))) {
      for (const row of sheet.rows()) {
        for (const cell of row.cells) {
          if (cell.address in streaming) streaming[cell.address] = textOf(cell.value);
        }
      }
      break; // first worksheet only
    }
    return {
      eager: eagerOf(readFixture(rel)),
      streaming,
      roundtrip: eagerOf(roundtrip(readFixture(rel))),
    };
  },

  // Read a fixture, write it, and report how the written package spells what the source spelled in
  // ISO/IEC 29500 Strict → { strictParts, relationshipTypes, attributes }. `strictParts` names every
  // written part whose text still mentions a `purl.oclc.org/ooxml/` URI; `relationshipTypes` is every
  // distinct Type the written `.rels` parts declare, sorted; `attributes` maps each requested
  // `<part>#<element> <attribute>` (the element by its qualified name as written) to the values that
  // attribute takes in that written part, in document order.
  strictSpellingsAfterWrite(rel: string, attributes: string[] = []) {
    const parts = partMapOf(writeXlsx(readFixture(rel)));
    const names = Object.keys(parts).sort();
    return {
      strictParts: names.filter((name) => (parts[name] ?? '').includes('purl.oclc.org/ooxml/')),
      relationshipTypes: [
        ...new Set(
          names
            .filter((name) => name.endsWith('.rels'))
            .flatMap((name) =>
              [...(parts[name] ?? '').matchAll(/\bType="([^"]*)"/g)].map((match) => match[1] ?? ''),
            ),
        ),
      ].sort(),
      attributes: Object.fromEntries(
        attributes.map((key) => {
          const [part = '', selector = ''] = key.split('#');
          const [element = '', attribute = ''] = selector.split(' ');
          const pattern = new RegExp(String.raw`<${element}\b[^>]*\s${attribute}="([^"]*)"`, 'g');
          return [key, [...(parts[part] ?? '').matchAll(pattern)].map((match) => match[1] ?? '')];
        }),
      ),
    };
  },

  // Read a fixture and report which date system it declares plus the requested cells three ways →
  // { epoch, eager, streaming, roundtrip }, each a map of A1 reference → { type, value }. A workbook
  // declaring the 1904 system counts its serials from 1904-01-01 rather than 1900-01-01, so all three
  // readings must land on the same calendar date, and the round trip must still declare the system:
  // dropping the flag leaves the serials where they were for a consumer to re-read four years out.
  dateSystemReport(rel: string, refs: string[] = []) {
    // Typed as the sibling cell reports are: a cell's normalised value is a union of every shape a
    // cell can hold, so an inferred map would make every case narrow it before reading the one field
    // the case is about.
    const cellsOf = (workbook: WorkbookInstance): Record<string, Untyped> => {
      const sheet = workbook.worksheets[0];
      return Object.fromEntries(
        refs.map((ref) => {
          const cell = sheet === undefined ? null : sheet.getCell(ref);
          return [
            ref,
            cell === null
              ? null
              : {type: detectValueType(cell.value), value: normalizeStreamValue(cell.value)},
          ];
        }),
      );
    };
    // The row streamer is built on the XML worksheet parser and has no BIFF12 path, so a binary
    // fixture reports no streamed reading rather than a failure: `null` is the honest answer to
    // "what does the streaming reader make of this", and a case compares the two only where both run.
    const streaming: Record<string, Untyped> = Object.fromEntries(refs.map((ref) => [ref, null]));
    if (rel.endsWith('.xlsx')) {
      for (const sheet of readWorkbookStream(fixtureBytes(rel))) {
        for (const row of sheet.rows()) {
          for (const cell of row.cells) {
            if (cell.address in streaming) {
              streaming[cell.address] = {
                type: detectValueType(cell.value),
                value: normalizeStreamValue(cell.value),
              };
            }
          }
        }
        break; // first worksheet only
      }
    }
    const before = readFixture(rel);
    const after = roundtrip(before);
    return {
      epoch: before.dateEpoch,
      roundtripEpoch: after.dateEpoch,
      eager: cellsOf(before),
      streaming,
      roundtrip: cellsOf(after),
    };
  },

  // Read a fixture and report the VBA code names it carries, before and after a write/read round trip
  // → { before, after }, each { workbook, sheets: {<sheet name>: {codeName, tabColor}} }. A code name
  // is what a macro project means by `ThisWorkbook` or `Sheet1`, so dropping one on a round trip
  // severs the binding between a `.xlsm`'s macros and the document they act on. The tab colour rides
  // along because it is the other thing `<sheetPr>` carries: the attribute and the children are
  // written by one function, so a case can hold it to emitting both rather than one or the other.
  codeNamesReport(rel: string) {
    const namesOf = (
      workbook: WorkbookInstance,
    ): {workbook: string | null; sheets: Record<string, Untyped>} => ({
      workbook: workbook.codeName ?? null,
      sheets: Object.fromEntries(
        workbook.worksheets.map((sheet) => [
          sheet.name,
          {codeName: sheet.codeName ?? null, tabColor: sheet.tabColor ?? null},
        ]),
      ),
    });
    const before = readFixture(rel);
    return {before: namesOf(before), after: namesOf(roundtrip(before))};
  },

  // Read a fixture, write it back, and parse the requested cells straight from the re-emitted sheet
  // XML → { hasNaNToken, cells }. Each cell is { t, formula, value } read off the raw `<c>`. Guards
  // that a string-typed formula result under a date format is not coerced to a numeric/NaN cell.
  roundtripFixtureCellXml(rel: string, refs: string[] = []) {
    const parts = partMapOf(writeXlsx(readXlsx(fixtureBytes(rel))));
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const cells: Record<string, Untyped> = {};
    for (const ref of refs) {
      const match = sheetXml.match(new RegExp(`<c r="${ref}"([^>]*)>([\\s\\S]*?)</c>`));
      if (!match) continue;
      const t = (match[1]!.match(/\bt="([^"]*)"/) || [])[1] ?? null;
      const formula = (match[2]!.match(/<f[^>]*>([\s\S]*?)<\/f>/) || [])[1] ?? null;
      const rawValue = (match[2]!.match(/<v>([\s\S]*?)<\/v>/) || [])[1] ?? null;
      // A t="str" (or shared-string) cell holds text; anything else with a bare <v> is numeric.
      const value = rawValue === null ? null : t === 'str' ? rawValue : Number(rawValue);
      cells[ref] = {t, formula, value};
    }
    return {hasNaNToken: /<v>[^<]*NaN[^<]*<\/v>/.test(sheetXml), cells};
  },

  // Add a sheet, then probe name lookup and uniqueness for case-consistency → { foundExact,
  // foundVariant, addVariantThrew }. Lookup and add must agree on identity: a case-variant name is
  // found by getWorksheet AND rejected by addWorksheet (both case-insensitive), so no absent-yet-
  // unaddable surprise exists.
  worksheetNameLookupReport() {
    const workbook = new Workbook();
    workbook.addWorksheet('Sheet');
    const foundExact = workbook.getWorksheet('Sheet') !== undefined;
    const foundVariant = workbook.getWorksheet('sheet') !== undefined;
    let addVariantThrew = false;
    try {
      workbook.addWorksheet('sheet');
    } catch {
      addVariantThrew = true;
    }
    return {foundExact, foundVariant, addVariantThrew};
  },

  // Read a real fixture `.xlsx` and report only whether it loaded, any error, its sheet names, and
  // a couple of core properties → { ok, error, sheetNames, lastModifiedBy, creator }. The read error
  // is captured and returned as data (never propagated) so a case asserts on a crash rather than the
  // runner blowing up. Exercises the reader against foreign generators and schema-valid corners Excel
  // never emits (namespace-prefixed roots, a leading BOM, unusual part order, missing optional parts).
  readFixtureReport(rel: string) {
    try {
      const wb = readFixture(rel);
      return {
        ok: true,
        error: null,
        sheetNames: wb.worksheets.map((s) => s.name),
        lastModifiedBy: wb.properties.lastModifiedBy ?? null,
        creator: wb.properties.creator ?? null,
      };
    } catch (e) {
      return {ok: false, error: messageOf(e), sheetNames: null};
    }
  },

  // Author a sheet named "History" (a name Excel blocks in its UI but that is a valid OOXML name)
  // through the public API, round-trip it, and separately confirm a genuinely-illegal name is still
  // rejected → { addThrew, addError, roundtripName, invalidRejected }. The reserved-name UI nicety
  // is not a document-model rule, so construction must not throw; the character-legality guard is a
  // real rule and must stay.
  addReservedSheetNameReport() {
    const wb = new Workbook();
    let addThrew = false;
    let addError = null;
    try {
      wb.addWorksheet('History');
    } catch (e) {
      addThrew = true;
      addError = messageOf(e);
    }
    const reloaded = roundtrip(wb);
    const roundtripName =
      reloaded.worksheets.map((sheet) => sheet.name).find((name) => name === 'History') ?? null;
    let invalidRejected = false;
    try {
      wb.addWorksheet('a/b');
    } catch {
      invalidRejected = true;
    }
    return {addThrew, addError, roundtripName, invalidRejected};
  },

  // Author a workbook whose sheets carry each visibility state (a valid workbook keeps one visible),
  // then report the state read back after a round-trip and the state attribute the workbook.xml
  // sheet-list entry carries → { readStates, xmlStates }, each keyed by sheet name. veryHidden, a
  // format-only state, must survive both as the model state and as the sheet-list attribute, never
  // degrading to hidden or visible.
  worksheetStateReport() {
    const wb = new Workbook();
    wb.addWorksheet('Visible');
    wb.addWorksheet('Hid', {state: 'hidden'});
    wb.addWorksheet('VeryHid', {state: 'veryHidden'});
    const buffer = writeXlsx(wb);
    const readStates: Record<string, Untyped> = {};
    for (const sheet of readXlsx(buffer).worksheets) readStates[sheet.name] = sheet.state;
    const workbookXml = partMapOf(buffer)['xl/workbook.xml'] ?? '';
    const xmlStates: Record<string, Untyped> = {};
    for (const m of workbookXml.matchAll(/<sheet\b([^>]*)\/?>/g)) {
      const attrs = m[1]!;
      const name = attrs.match(/\bname="([^"]*)"/)?.[1];
      if (name === undefined) continue;
      xmlStates[name] = attrs.match(/\bstate="([^"]*)"/)?.[1] ?? 'visible';
    }
    return {readStates, xmlStates};
  },

  // Read a fixture and report its defined names as { <name>: [refersTo…] }, mirroring the oracle.
  // The model retains every name as its own entry rather than keying by name, so two same-named
  // names scoped to different sheets both survive: the scope collision that drops one on the
  // oracle's name-keyed reader.
  readFixtureDefinedNames(rel: string) {
    const wb = readFixture(rel);
    const names: Record<string, Untyped> = {};
    for (const dn of wb.definedNames) {
      names[dn.name] ||= [];
      names[dn.name].push(dn.refersTo);
    }
    for (const k of Object.keys(names)) names[k]!.sort();
    return {names, count: Object.keys(names).length, modelCount: wb.definedNames.length};
  },

  // Read a real fixture `.xlsx` and report the fill and font colour the reader surfaces for each
  // requested `<sheet>!<address>` cell → { [key]: { fill, fontColor } | null }. Mirrors the oracle:
  // a solid-pattern fill's visible colour lives on fgColor while bgColor is the automatic indexed
  // placeholder, and the font colour is a wholly separate facet; the two are never conflated.
  // Read a real fixture `.xlsx` and report each requested cell's observable type, value, number
  // format, and note → { <addr>: {type, value, numFmt, note} | null }, on the first sheet. Mirrors
  // the oracle: a date-formatted numeric serial surfaces as a Date (value { date: iso }), not a raw
  // number, honouring the 1900 date-system leap-year quirk. `type` is a stable label.
  readFixtureCells(rel: string, cells: Untyped = []) {
    const wb = readFixture(rel);
    const sheet = wb.worksheets[0];
    const out: Record<string, Untyped> = {};
    for (const addr of cells) {
      const cell = sheet ? sheet.getCell(addr) : null;
      out[addr] = cell
        ? {
            type: cell.type,
            value: normalizeStreamValue(cell.value),
            numFmt: cell.numFmt ?? null,
            note: cell.note !== undefined ? cell.note : undefined,
          }
        : null;
    }
    return out;
  },

  // ── Streaming reader (readWorkbookStream) ──────────────────────────────────────────────────────
  // The corpus's streaming-read contract, bound to the rewrite's generator-based reader. Each method
  // mirrors its oracle sibling in workbook-io.mjs; where a case compares the streaming path to the
  // eager one, BOTH come from the rewrite, so the assertion checks that streaming and buffered reads
  // agree cell-for-cell. The rewrite's reader is a synchronous generator, so the "without race" and
  // "chunk boundary" hazards the ExcelJS stream reader faces are structurally absent.

  // Read a real styled template, write it straight back out, read the result, and report whether
  // its sheet names, custom column widths, and per-cell fill/font/numFmt/alignment/border survived
  // that no-op read→write→read → { sheetNames(Before), columns(Before), styleSurvival }. This is the
  // mainstream "open a styled template, fill it in, save it" path, which must be format-preserving.
  // Style comparison is key-order-insensitive so a case asserts on content survival, not
  // serialization incidentals. In the rewrite's model a column stores a width only when it is a
  // custom width, so "has a width" is exactly "is a custom width".
  roundtripFixture(rel: string) {
    const before = readFixture(rel);
    const after = roundtrip(before);

    const hasStyle = (cell: CellInstance) =>
      !!(cell.numFmt || cell.fill?.type || cell.font || cell.alignment || cell.border);
    const styleKey = (cell: CellInstance) =>
      JSON.stringify(
        canonicalJson({
          numFmt: cell.numFmt || null,
          fill: cell.fill?.type ? cell.fill : null,
          font: cell.font || null,
          alignment: cell.alignment || null,
          border: cell.border || null,
        }),
      );
    const columnsWithWidth = (wb: WorkbookInstance) => {
      const out: Record<string, Untyped> = {};
      for (const sheet of wb.worksheets) {
        const cols: Record<string, Untyped> = {};
        for (const {index, properties} of sheet.columns()) {
          // `properties` is absent for a column carrying no formatting record at all. `columns()`
          // yields those too, and reading `.width` straight off it would throw rather than report.
          if (properties?.width !== undefined)
            cols[index] = {width: properties.width, customWidth: true};
        }
        out[sheet.name] = cols;
      }
      return out;
    };

    let checked = 0;
    let mismatches = 0;
    let sample = null;
    for (const sheet of before.worksheets) {
      const other = after.getWorksheet(sheet.name);
      for (const {cells} of sheet.rows()) {
        for (const cell of cells) {
          // Resolve both sides through getCell so a merged-range slave redirects to its master on
          // each: comparing the row-iterated slave (its own style) against getCell (the master) would
          // report a phantom drift that is only an access asymmetry, not a lost style.
          const beforeCell = sheet.getCell(cell.address);
          if (!hasStyle(beforeCell)) continue;
          checked += 1;
          const beforeKey = styleKey(beforeCell);
          const afterKey = other ? styleKey(other.getCell(cell.address)) : '(sheet missing)';
          if (beforeKey !== afterKey) {
            mismatches += 1;
            if (!sample)
              sample = {cell: `${sheet.name}!${cell.address}`, before: beforeKey, after: afterKey};
          }
        }
      }
    }

    return {
      sheetNamesBefore: before.worksheets.map((s) => s.name),
      sheetNames: after.worksheets.map((s) => s.name),
      columnsBefore: columnsWithWidth(before),
      columns: columnsWithWidth(after),
      styleSurvival: {checked, mismatches, sample},
    };
  },

  // Build → write → read back through the rewrite's own reader, then normalize to the
  // same JSON model current.mjs reports, so every write→read round-trip case runs
  // unchanged. Facets the writer/reader do not materialize yet come back empty/null;
  // the writer's feature-gate keeps a case whose spec needs those from ever running here.
  roundtripWorkbook(spec: Untyped) {
    const reloaded = roundtrip(buildFrom(spec));
    const sheets: Record<string, Untyped> = {};
    for (const s of spec.sheets || []) {
      const sheet = reloaded.getWorksheet(s.name);
      if (!sheet) {
        sheets[s.name] = null;
        continue;
      }
      const cells: Record<string, Untyped> = {};
      for (const c of s.cells || []) {
        cells[c.ref] = normalizeRewriteCell(sheet.getCell(c.ref), sheet);
      }
      const columns: Record<string, Untyped> = {};
      for (const col of s.columns || []) {
        const p = sheet.getColumn(col.index);
        columns[col.index] = {width: p.width ?? null, hidden: !!p.hidden, numFmt: p.numFmt ?? null};
      }
      const rows: Record<string, Untyped> = {};
      for (const row of s.rows || []) {
        const p = sheet.getRow(row.index);
        rows[row.index] = {height: p.height ?? null, hidden: !!p.hidden};
      }
      const margins = Object.keys(sheet.pageMargins).length > 0 ? {...sheet.pageMargins} : null;
      const ps = sheet.pageSetup;
      sheets[s.name] = {
        cells,
        columns,
        rows,
        margins,
        pageSetup: {
          fitToPage: !!ps.fitToPage,
          fitToWidth: ps.fitToWidth ?? null,
          fitToHeight: ps.fitToHeight ?? null,
          scale: ps.scale ?? null,
          paperSize: ps.paperSize ?? null,
        },
        autoFilter: sheet.autoFilter?.ref ?? null,
        merges: [...sheet.merges],
        rowCount: sheet.rowCount,
        actualRowCount: sheet.actualRowCount,
      };
    }
    const props = reloaded.properties;
    const definedNames: Record<string, Untyped> = {};
    for (const dn of reloaded.definedNames) {
      definedNames[dn.name] ||= [];
      definedNames[dn.name].push(dn.refersTo);
    }
    for (const k of Object.keys(definedNames)) definedNames[k]!.sort();
    return {
      properties: {
        creator: props.creator ?? null,
        lastModifiedBy: props.lastModifiedBy ?? null,
        created: isoOrNull(props.created),
        modified: isoOrNull(props.modified),
      },
      sheets,
      definedNames,
    };
  },

  // Load a fixture and try to write it back → { loadOk, loadError, writeOk, writeError,
  // sheetNames }. Use it to assert a foreign construct round-trips without the writer crashing.
  roundtripFixtureWriteReport(rel: string) {
    let workbook: WorkbookInstance;
    try {
      workbook = readFixture(rel);
    } catch (e) {
      return {
        loadOk: false,
        loadError: messageOf(e),
        writeOk: false,
        writeError: null,
        sheetNames: [],
      };
    }
    let writeOk = false;
    let writeError = null;
    try {
      writeXlsx(workbook);
      writeOk = true;
    } catch (e) {
      writeError = messageOf(e);
    }
    return {
      loadOk: true,
      loadError: null,
      writeOk,
      writeError,
      sheetNames: workbook.worksheets.map((w) => w.name),
    };
  },

  // Clone a worksheet through its model export/import: build a source sheet with cells and merges,
  // read its model, and assign that model onto a fresh sheet. Reports the merges the source model
  // exposed and the merges the destination carries afterwards → { srcMerges, dstMerges, error }.
  // The historical bug this measures is an asymmetric model contract that dropped merges on import;
  // the rewrite's getter and setter cover the same fields, so the round-trip is lossless.
  copyWorksheetModel({merges = ['A1:C1'], cells = [{ref: 'A1', value: 'merged'}]}: Untyped = {}) {
    const workbook = new Workbook();
    const src = workbook.addWorksheet('Src');
    for (const c of cells) src.getCell(c.ref).value = c.value;
    for (const m of merges) src.mergeCells(m);
    const dst = workbook.addWorksheet('Dst');

    let error = null;
    // Declared `string[]` rather than `Untyped[]`: these are A1 range strings, and typing them is
    // what lets the sort below be checked at all.
    let dstMerges: string[] = [];
    const srcMerges: string[] = [...src.model.merges];
    try {
      dst.model = {...src.model, name: 'Dst'} as Untyped;
      dstMerges = [...dst.model.merges];
    } catch (e) {
      error = messageOf(e);
    }
    return {srcMerges: srcMerges.sort(), dstMerges: dstMerges.sort(), error};
  },

  mutateWorksheet({cells = [], ops = [], read = [], readStyles = []}: Untyped = {}) {
    const sheet = new Workbook().addWorksheet('S');
    for (const c of cells) {
      const cell = sheet.getCell(c.ref);
      cell.value = c.value;
      // Optional per-cell style so a case can assert a structural edit carries a cell's style to its
      // shifted position rather than blanking it.
      if (c.font) cell.font = c.font;
      if (c.fill) cell.fill = c.fill;
      if (c.numFmt) cell.numFmt = c.numFmt;
    }

    let error = null;
    try {
      for (const op of ops) {
        const inserts = op.inserts || [];
        if (op.op === 'spliceRows') sheet.spliceRows(op.start, op.count, ...inserts);
        else if (op.op === 'spliceColumns') sheet.spliceColumns(op.start, op.count, ...inserts);
        else if (op.op === 'mergeCells') sheet.mergeCells(op.range);
        else if (op.op === 'insertRow') sheet.insertRow(op.pos, op.value || []);
        else if (op.op === 'duplicateRow')
          sheet.duplicateRow(op.start, {count: op.count ?? 1, insert: op.insert !== false});
        else throw new Error(`unknown mutation op: ${op.op}`);
      }
    } catch (e) {
      error = messageOf(e);
    }

    const readCells: Record<string, Untyped> = {};
    for (const ref of read) readCells[ref] = sheet.getCell(ref).value ?? null;

    // Per-cell style facets after the mutations. The style a cell carried before a splice must
    // still describe the (possibly shifted) cell afterward, rather than being lost.
    const styles: Record<string, Untyped> = {};
    for (const ref of readStyles) {
      const cell = sheet.getCell(ref);
      styles[ref] = {
        value: cell.value ?? null,
        font: cell.font ? JSON.parse(JSON.stringify(cell.font)) : null,
        fill: cell.fill?.type ? JSON.parse(JSON.stringify(cell.fill)) : null,
        numFmt: cell.numFmt ?? null,
      };
    }

    // The last POPULATED row and its column-1 value, derived from the row iterator (ascending, so
    // the final populated row wins); a delete-splice must leave this on the true last row, never a
    // trailing empty slot.
    let lastRow = null;
    for (const {number, cells: rowCells} of sheet.rows()) {
      if (rowCells.some((c) => c.value !== null && c.value !== undefined)) {
        const first = rowCells.find((c) => c.col === 1);
        lastRow = {number, value: first?.value ?? null};
      }
    }

    return {
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      cells: readCells,
      styles,
      merges: [...sheet.merges],
      lastRow,
      error,
    };
  },

  tryWriteWorkbook(spec: Untyped) {
    let workbook: WorkbookInstance;
    try {
      workbook = buildFrom(spec);
    } catch (error) {
      if (error instanceof UnsupportedSpecError) throw error;
      return {ok: false, phase: 'build', error: messageOf(error)};
    }
    let buffer: Uint8Array;
    try {
      buffer = writeXlsx(workbook);
    } catch (error) {
      if (error instanceof UnsupportedSpecError) throw error;
      return {ok: false, phase: 'write', error: messageOf(error)};
    }
    // Report which cells survived the round-trip, so a case can prove a bad cell (e.g. an Invalid
    // Date, written value-less) did not drop its siblings.
    const reread = readXlsx(buffer);
    const survivingCells: Record<string, Untyped> = {};
    for (const s of spec.sheets || []) {
      const sheet = reread.getWorksheet(s.name);
      survivingCells[s.name] = (s.cells || [])
        .filter((c: Untyped) => sheet && sheet.getCell(c.ref).value !== null)
        .map((c: Untyped) => c.ref);
    }
    return {ok: true, byteLength: buffer.byteLength ?? buffer.length, survivingCells};
  },
  // Replace a written package's shared-string pool with one whose empty entries are spelled the two
  // legal ways (`<si/>` and `<si><t/></si>`), and point a row of `t="s"` cells at every slot plus one
  // whose `<v>` is present and empty -> { pooled, emptyValueCell }. Both halves are about a *wrong*
  // value rather than a missing one: an empty entry that commits no slot shifts every later index, so
  // a cell resolves to its neighbour string, and an empty `<v/>` read as a number is index 0, so the
  // cell resolves to the first pooled string. Neither is visible to anything downstream.
  pooledStringIndexReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    // Four distinct strings so the writer emits a pool part at all; its content is replaced below.
    for (const [index, ref] of ['A1', 'B1', 'C1', 'D1'].entries()) {
      sheet.getCell(ref).value = `seed${index}`;
    }
    const cells = ['A1', 'B1', 'C1', 'D1'].map(
      (ref, index) => `<c r="${ref}" t="s"><v>${index}</v></c>`,
    );
    const back = reloadPatched(writeXlsx(wb, {useSharedStrings: true}), {
      'xl/sharedStrings.xml': () =>
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="4">' +
        '<si><t>first</t></si><si/><si><t/></si><si><t>fourth</t></si></sst>',
      'xl/worksheets/sheet1.xml': (xml) =>
        xml.replace(
          /<sheetData>[\s\S]*<\/sheetData>/,
          `<sheetData><row r="1">${cells.join('')}<c r="E1" t="s"><v/></c></row></sheetData>`,
        ),
    });
    const s1 = back.getWorksheet('S')!;
    return {
      pooled: ['A1', 'B1', 'C1', 'D1'].map((ref) => s1.getCell(ref).value),
      emptyValueCell: s1.getCell('E1').value,
    };
  },

  // Point a Strict-mode (`t="d"`) date cell at text that is not a date, beside one that is
  // -> { unparseable, unparseableIsDate, parsed }. An unguarded `new Date(text)` yields a Date whose
  // time is NaN: it satisfies every guard, reaches the writer, and serialises as the literal string
  // `Invalid Date`, so the cell reads back as a value nobody stored.
  strictModeDateCellReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('A1').value = 'a';
    sheet.getCell('B1').value = 'b';
    const back = reloadPatched(writeXlsx(wb), {
      'xl/worksheets/sheet1.xml': (xml) =>
        xml.replace(
          /<sheetData>[\s\S]*<\/sheetData>/,
          '<sheetData><row r="1">' +
            '<c r="A1" t="d"><v>not-a-date</v></c>' +
            '<c r="B1" t="d"><v>2024-03-04T05:06:07Z</v></c>' +
            '</row></sheetData>',
        ),
    });
    const s1 = back.getWorksheet('S')!;
    const unparseable = s1.getCell('A1').value;
    const parsed = s1.getCell('B1').value;
    return {
      unparseable,
      unparseableIsDate: unparseable instanceof Date,
      parsed: parsed instanceof Date ? parsed.toISOString() : null,
    };
  },

  // Patch one reference-bearing attribute of a written package with a value no cell, area or region
  // can have, then read the result back -> one row per mutation
  // { mutation, threw, isXlsxError, errorName, keptSiblingCell, sheetsRead }. A `.xlsx` the library
  // did not write is allowed to be wrong: a malformed `r`/`ref`/`sqref` must cost the element that
  // carried it, never the sheet, and must never surface as a native RangeError/SyntaxError from
  // outside the XlsxError taxonomy one `catch` clause is supposed to answer.
  malformedReferenceReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('A1').value = 'keep';
    sheet.getCell('B2').value = 'target';
    sheet.getCell('D1').value = 'go';
    sheet.addHyperlink({ref: 'D1', target: 'https://example.com/'});
    sheet.getCell('H1').note = 'a note';
    sheet.addDataValidation('E1:E5', {type: 'list', allowBlank: true, formulae: ['"a,b"']});
    sheet.addTable({name: 'T', ref: 'J1', columns: [{name: 'H'}], rowCount: 1});
    sheet.autoFilter = 'A1:B2';
    const bytes = writeXlsx(wb);

    // Each mutation names the attribute it corrupts and the part that carries it. The values are the
    // three ways a reference can be wrong that the tolerant decoders now answer for: off the grid on
    // the row axis (`A0`), off it on the column axis (`ZZZZ1`), and unparseable (`junk!!`).
    const mutations: {name: string; part: string; patch: (xml: string) => string}[] = [
      {
        name: 'cell r off the grid (A0)',
        part: 'xl/worksheets/sheet1.xml',
        patch: (xml) => xml.replace('r="B2"', 'r="A0"'),
      },
      {
        name: 'cell r past the last column (ZZZZ1)',
        part: 'xl/worksheets/sheet1.xml',
        patch: (xml) => xml.replace('r="B2"', 'r="ZZZZ1"'),
      },
      {
        name: 'cell r unparseable (junk!!)',
        part: 'xl/worksheets/sheet1.xml',
        patch: (xml) => xml.replace('r="B2"', 'r="junk!!"'),
      },
      {
        name: 'dataValidation sqref off the grid',
        part: 'xl/worksheets/sheet1.xml',
        patch: (xml) => xml.replace('sqref="E1:E5"', 'sqref="ZZZZ0"'),
      },
      {
        name: 'autoFilter ref off the grid',
        part: 'xl/worksheets/sheet1.xml',
        patch: (xml) => xml.replace(/<autoFilter ref="[^"]*"/, '<autoFilter ref="ZZZZ0"'),
      },
      {
        name: 'hyperlink ref unparseable',
        part: 'xl/worksheets/sheet1.xml',
        patch: (xml) => xml.replace(/(<hyperlink ref=")[^"]*/, '$1junk!!'),
      },
      {
        name: 'table ref unparseable',
        part: 'xl/tables/table1.xml',
        patch: (xml) => xml.replace(/(ref=")[^"]*/, '$1ZZZZ0:!!'),
      },
      {
        name: 'comment ref off the grid',
        part: 'xl/comments1.xml',
        patch: (xml) => xml.replace(/(<comment ref=")[^"]*/, '$1ZZZZ0'),
      },
    ];

    return mutations.map((mutation) => {
      try {
        const back = reloadPatched(bytes, {[mutation.part]: mutation.patch});
        const s = back.getWorksheet('S')!;
        return {
          mutation: mutation.name,
          threw: false,
          isXlsxError: null,
          errorName: null,
          keptSiblingCell: s.getCell('A1').value,
          sheetsRead: back.worksheets.length,
          // One census per read, so a case can name both halves of the contract: the corrupted
          // element is absent, and every feature it did not touch is still there.
          model: {
            targetCell: s.getCell('B2').value,
            validations: s.dataValidations.length,
            tables: s.tables.length,
            autoFilter: s.autoFilter?.ref ?? null,
            hyperlink: s.hyperlinkAt('D1')?.target ?? null,
            note: s.getCell('H1').note ?? null,
          },
        };
      } catch (error) {
        return {
          mutation: mutation.name,
          threw: true,
          isXlsxError: error instanceof XlsxError,
          errorName: (error as Error)?.constructor?.name ?? null,
          keptSiblingCell: null,
          sheetsRead: 0,
          model: null,
        };
      }
    });
  },
  // Patch one *name* of a written package to something the model's own authoring guards refuse, then
  // read the result back and write it out again -> one row per mutation
  // { mutation, threw, isXlsxError, errorName, sheetNames, tables, definedNames, rewrote }. The
  // taxonomy says `AuthoringError` is always the caller's fault and never the file's, and native
  // `RangeError`/`SyntaxError` are outside it altogether: a file naming a sheet twice, or a table
  // `1 bad`, must therefore cost that name or that feature, never the read.
  hostileNameReport() {
    const wb = new Workbook();
    const alpha = wb.addWorksheet('Alpha');
    alpha.getCell('A1').value = 'keep';
    alpha.addTable({name: 'T', ref: 'C1', columns: [{name: 'H'}], rowCount: 1});
    wb.addWorksheet('Beta').getCell('A1').value = 'beta';
    wb.defineName({name: 'Rate', refersTo: 'Alpha!$A$1'});
    const bytes = writeXlsx(wb);

    const SHEET = 'xl/workbook.xml';
    const TABLE = 'xl/tables/table1.xml';
    const renameSecondSheet = (to: string) => (xml: string) =>
      xml.replace('name="Beta"', `name="${to}"`);
    const renameTable = (to: string) => (xml: string) =>
      xml.replace('name="T"', `name="${to}"`).replace('displayName="T"', `displayName="${to}"`);

    const mutations: {name: string; part: string; patch: (xml: string) => string}[] = [
      {name: 'duplicate sheet name', part: SHEET, patch: renameSecondSheet('Alpha')},
      {name: 'empty sheet name', part: SHEET, patch: renameSecondSheet('')},
      {name: 'sheet name with a forbidden character', part: SHEET, patch: renameSecondSheet('a/b')},
      {
        name: 'sheet name over the length limit',
        part: SHEET,
        patch: renameSecondSheet('B'.repeat(32)),
      },
      {
        name: 'sheet name edged with an apostrophe',
        part: SHEET,
        patch: renameSecondSheet("'Beta'"),
      },
      {name: 'table name that is not an identifier', part: TABLE, patch: renameTable('1 bad')},
      {name: 'table name over the length limit', part: TABLE, patch: renameTable('T'.repeat(300))},
      {name: 'table name that reads as a cell', part: TABLE, patch: renameTable('T1')},
      {
        name: 'defined name with no name at all',
        part: SHEET,
        patch: (xml) => xml.replace('<definedName name="Rate"', '<definedName name=""'),
      },
    ];

    return mutations.map((mutation) => {
      try {
        const back = reloadPatched(bytes, {[mutation.part]: mutation.patch});
        // Written out again, because half of what "the read succeeded" has to mean is that the model
        // it produced is one this library can still serialise.
        let rewrote = true;
        try {
          writeXlsx(back);
        } catch {
          rewrote = false;
        }
        return {
          mutation: mutation.name,
          threw: false,
          isXlsxError: null,
          errorName: null,
          sheetNames: back.worksheets.map((sheet) => sheet.name),
          keptSiblingCell: back.worksheets[0]?.getCell('A1').value ?? null,
          tables: back.worksheets[0]?.tables.length ?? 0,
          definedNames: back.definedNames.length,
          rewrote,
        };
      } catch (error) {
        return {
          mutation: mutation.name,
          threw: true,
          isXlsxError: error instanceof XlsxError,
          errorName: (error as Error)?.constructor?.name ?? null,
          sheetNames: [],
          keptSiblingCell: null,
          tables: 0,
          definedNames: 0,
          rewrote: false,
        };
      }
    });
  },

  // Put a character XML 1.0 cannot carry in front of the reader, as the reference a hostile file
  // spells it with and as the raw byte, in each of the two places a structural string lives, then
  // read AND re-write -> one row per case
  // { where, spelling, threw, rewrote, sheetName, definedName, carriesUnrepresentable }. The reader
  // is only allowed to produce values the writer can serialise; decoding one it cannot turns a
  // hostile file into an `AuthoringError` blaming the caller on the next save.
  unrepresentableCharacterReport() {
    const wb = new Workbook();
    wb.addWorksheet('Alpha').getCell('A1').value = 'keep';
    wb.defineName({name: 'Rate', refersTo: 'Alpha!$A$1'});
    const bytes = writeXlsx(wb);

    const cases: {where: string; spelling: string; patch: (xml: string) => string}[] = [
      {
        where: 'sheet name',
        spelling: '&#1;',
        patch: (xml) => xml.replace('name="Alpha"', 'name="A&#1;lpha"'),
      },
      {
        where: 'sheet name',
        spelling: '&#xD800;',
        patch: (xml) => xml.replace('name="Alpha"', 'name="A&#xD800;lpha"'),
      },
      {
        where: 'sheet name',
        spelling: 'a raw control character',
        patch: (xml) => xml.replace('name="Alpha"', 'name="A\u0001lpha"'),
      },
      {
        where: 'defined name text',
        spelling: '&#1;',
        patch: (xml) => xml.replace('Alpha!$A$1', 'Alpha!$A$1&#1;'),
      },
      {
        where: 'defined name text',
        spelling: 'a raw control character',
        patch: (xml) => xml.replace('Alpha!$A$1', 'Alpha!$A$1\u0001'),
      },
    ];

    return cases.map(({where, spelling, patch}) => {
      try {
        const back = reloadPatched(bytes, {'xl/workbook.xml': patch});
        const sheetName = back.worksheets[0]?.name ?? null;
        const definedName = back.definedNames[0]?.refersTo ?? null;
        let rewrote = true;
        try {
          writeXlsx(back);
        } catch {
          rewrote = false;
        }
        return {
          where,
          spelling,
          threw: false,
          rewrote,
          sheetName,
          definedName,
          // What matters is not which spelling survived but that nothing unrepresentable did.
          carriesUnrepresentable: UNREPRESENTABLE.test(`${sheetName ?? ''}${definedName ?? ''}`),
        };
      } catch {
        return {
          where,
          spelling,
          threw: true,
          rewrote: false,
          sheetName: null,
          definedName: null,
          carriesUnrepresentable: false,
        };
      }
    });
  },

  // Write a workbook whose `created`/`modified` document property is each shape a `Date` can take at
  // this boundary -> one row per shape { property, kind, threw, isXlsxError, errorName, message,
  // stamp }. `toISOString()` answers three of the four with something no `dcterms:W3CDTF` can carry:
  // a throw from outside the taxonomy, or ISO 8601's expanded-year form that Excel repairs.
  documentDateReport() {
    const shapes: {kind: string; date: Date}[] = [
      {kind: 'an ordinary timestamp', date: new Date(Date.UTC(2026, 0, 2, 3, 4, 5, 678))},
      {kind: 'an Invalid Date', date: new Date(Number.NaN)},
      {kind: 'a year past 9999', date: new Date(Date.UTC(275760, 8, 12))},
      {kind: 'a year before 0000', date: new Date(-62200000000000)},
    ];
    return shapes.flatMap(({kind, date}) =>
      (['created', 'modified'] as const).map((property) => {
        const wb = new Workbook();
        wb.addWorksheet('S');
        wb.properties[property] = date;
        try {
          const core = partMapOf(writeXlsx(wb))['docProps/core.xml'] ?? '';
          return {
            property,
            kind,
            threw: false,
            isXlsxError: null,
            errorName: null,
            message: null,
            stamp: new RegExp(`<dcterms:${property}[^>]*>([^<]*)<`).exec(core)?.[1] ?? null,
          };
        } catch (error) {
          return {
            property,
            kind,
            threw: true,
            isXlsxError: error instanceof XlsxError,
            errorName: (error as Error)?.constructor?.name ?? null,
            message: messageOf(error),
            stamp: null,
          };
        }
      }),
    );
  },

  // Namespace-prefix independence: a file may bind an OOXML namespace to any prefix it likes, and the
  // reader must read it the same either way. Built in `xml-probes.ts` beside the other crafted reader
  // inputs; named here so a case can reach them.
  relationshipPrefixReport,
  mainNamespacePrefixReport,
  themePrefixReport,
  selfClosingDefinedNameReport,
};
