// XML- and byte-level probes: small readers for the corners a model-level assertion cannot
// reach (drawing anchors, sqref expansion, patched-part reloads), plus the synthetic reader
// inputs used to prove the reader classifies foreign formats instead of crashing on them.

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';

import type {Untyped} from '../../untyped.ts';
import {patchedPackage} from './package-facts.ts';
import {
  decodeRange,
  encodeAddress,
  readXlsx,
  Workbook,
  writeCompoundFile,
  writeXlsx,
} from './runtime.ts';
import {buildFrom} from './spec-model.ts';

export const hexBytes = (hex: string) =>
  Uint8Array.from((hex.match(/../g) ?? []).map((h) => Number.parseInt(h, 16)));

// Parse the integer children of an <xdr:from>/<xdr:to> block, mirroring the oracle so a drawing anchor
// reports the same plain-number geometry from either adapter.
export const intAt = (xml: string, tag: string) => {
  const m = xml.match(new RegExp(`<${tag}>(-?\\d+)</${tag}>`));
  return m ? Number(m[1]) : null;
};
export const parseAnchorSide = (block: string | null | undefined) =>
  block
    ? {
        col: intAt(block, 'xdr:col'),
        colOff: intAt(block, 'xdr:colOff'),
        row: intAt(block, 'xdr:row'),
        rowOff: intAt(block, 'xdr:rowOff'),
      }
    : null;
// An `&` that opens no entity reference is the one escaping mistake a writer can make that renders a
// whole part unparseable, so it is the well-formedness question every XML-level probe here asks.
export const xmlWellFormed = (xml: string) =>
  !/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(xml);
export const imageXmlWellFormed = xmlWellFormed;

// The inverse of the writer's attribute escaping, so a probe can compare what came back out against
// what went in. Escaping applied twice survives this and shows up as a value still carrying `&amp;`.
export const decodeXmlEntities = (text: string) =>
  text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');

// Expand an OOXML sqref (space-separated ranges) into its covered cell references, bounded by a cap so
// a whole-column range never balloons, used to check that a range-form validation is reported on
// every covered cell. An unbounded whole-row/column part is skipped rather than expanded.
export function expandSqref(sqref: string, cap = 4096) {
  const refs: string[] = [];
  for (const part of String(sqref).split(/\s+/).filter(Boolean)) {
    const {left, right, top, bottom} = decodeRange(part);
    if (left == null || right == null || top == null || bottom == null) continue;
    for (let c = left; c <= right && refs.length < cap; c++) {
      for (let r = top; r <= bottom && refs.length < cap; r++) refs.push(encodeAddress(c, r));
    }
  }
  return refs;
}

// Rewrite named parts of a written package and read the result back: the way to feed the
// reader the hand-edited OOXML forms real producers emit but the writer itself never generates
// (an explicit-false boolean flag `<b val="0"/>`, an alignment element carrying only `wrapText="0"`,
// an injected xf). `edits` maps a part path to a (xml) => xml transform; unlisted parts pass through.
export function reloadPatched(buffer: Uint8Array, edits: Record<string, (xml: string) => string>) {
  return readXlsx(patchedPackage(buffer, {edit: edits}));
}

// Substitute a marker the author planted in a cell for arbitrary markup, in every text part that
// carries it, then read the package back. The writer escapes `&`, so an entity reference is markup
// it can never emit: rewriting the written package after the fact is the only way to put one in
// front of the reader, and it is exactly what a hostile file does.
export function reloadWithMarkupSubstituted(buffer: Uint8Array, marker: string, markup: string) {
  const files = unzipSync(buffer);
  for (const [name, bytes] of Object.entries(files)) {
    if (!name.endsWith('.xml')) continue;
    const xml = strFromU8(bytes);
    if (xml.includes(marker)) files[name] = strToU8(xml.split(marker).join(markup));
  }
  return readXlsx(zipSync(files));
}

// Parse an XML tag's attributes into a plain { name: value } map. base64 salt/hash values use
// only XML-safe characters, so a naive quoted-value scan is sufficient here.
export function attrsOf(tag: string) {
  const out: Record<string, string> = {};
  const re = /([\w:]+)="([^"]*)"/g;
  let m = re.exec(tag);
  while (m !== null) {
    out[m[1]!] = m[2]!;
    m = re.exec(tag);
  }
  return out;
}

// ── Macro-enabled fixture builder ──────────────────────────────────────────────────────────────────
// Assemble a genuine, navigable vbaProject.bin (via the production CFB writer + MS-OVBA compressor, so
// its storage tree is walkable exactly as Excel's is) inside a minimal .xlsm package. This is the only
// way to produce an edit-in-place *input* without an interactive VBA editor: the writer cannot author a
// project from a model (no reference support, document-module linkage is host-coupled), but the editor
// splices new module source into an existing bin. The fixture carries a hand-crafted PROJECTREFERENCES
// record and a `document` code-behind module: the two things splice-editing must preserve that
// re-synthesis structurally cannot.

// Build a reader input of a given format family, to probe the reader's typed-error classification: a
// genuine `.xlsx` (the control that must still read), a legacy `.xls` (an OLE2/CFB compound file, via the
// production CFB writer), a binary `.xlsb` (a real ZIP whose office document is `xl/workbook.bin`),
// non-ZIP text (a CSV handed to the wrong reader), and a ZIP-headed-but-corrupt archive.
/**
 * The format families {@link buildReadInput} can synthesise: a closed set, so a case naming one that
 * does not exist fails to compile instead of reaching the `default` branch at runtime.
 */
export type ReadInputKind = 'xlsx' | 'xls' | 'xlsb' | 'garbage' | 'corrupt-zip';

export function buildReadInput(kind: ReadInputKind): Uint8Array {
  switch (kind) {
    case 'xlsx':
      return writeXlsx(buildFrom({sheets: [{name: 'S', cells: [{ref: 'A1', value: 42}]}]}));
    case 'xls':
      return writeCompoundFile([{name: 'Workbook', data: strToU8('legacy biff bytes')}]);
    case 'xlsb':
      return zipSync({
        '[Content_Types].xml': strToU8(
          '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>' +
            '</Types>',
        ),
        'xl/workbook.bin': Uint8Array.of(0, 1, 2, 3),
      });
    case 'garbage':
      return strToU8('name,amount\nwidget,10\n');
    case 'corrupt-zip': {
      // A genuine package cut off mid-stream: a half-downloaded file, which is what a corrupt archive
      // looks like in the wild. The bytes a zip layer actually rejects are the point: a few hand-made
      // `PK` bytes are quietly skipped by a streaming unzip rather than reported as a failure.
      const good = writeXlsx(buildFrom({sheets: [{name: 'S', cells: [{ref: 'A1', value: 42}]}]}));
      return good.subarray(0, good.length >> 1);
    }
    default:
      throw new Error(`unknown read-input kind: ${String(kind)}`);
  }
}

// Turn a reader call into JSON-serializable classification facts: whether it threw, the error's `name`,
// `code` and `format` branch fields (the typed contract a caller keys on: `code` says what kind of
// failure it is, `format` which unsupported input it was), and whether the message leaks the zip
// layer's internals or an absolute filesystem path (the anti-leak contract).
export function classifyReadError(run: () => void): Untyped {
  try {
    run();
    return {threw: false, errorName: null, code: null, format: null, message: null};
  } catch (e) {
    const err = e as Untyped;
    const message = String(err?.message ?? '');
    return {
      threw: true,
      errorName: err?.name ?? null,
      code: err?.code ?? null,
      format: err?.format ?? null,
      message,
      leaksZipInternals: /central directory|is this a zip|invalid zip data|unexpected EOF/i.test(
        message,
      ),
      leaksAbsolutePath: /[A-Za-z]:\\|\/(?:Users|home)\//.test(message),
    };
  }
}

// ── Namespace-prefix independence ─────────────────────────────────────────────────────────────────
// A file is free to bind an OOXML namespace to any prefix it likes, and real toolchains do. The reader
// is prefix-agnostic almost everywhere (`localName` strips whatever prefix a file chose), which is
// what made the handful of places testing a prefix STRING dangerous: the rest of the file reads
// perfectly and one whole feature disappears with no error.
//
// These craft the packages in-memory rather than committing binaries, because the whole point of each
// is one namespace declaration and its consequence, and a diffable string says that where a zip cannot.

const SPREADSHEETML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const DRAWINGML = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/**
 * Rebind the relationships namespace in `xl/workbook.xml` from `r` to another prefix, then read →
 * `{sheetNames, a1}`. A workbook whose `<sheet>` elements carry a differently-prefixed relationship id
 * used to resolve no relationship at all, so every sheet loaded permanently empty.
 */
export function relationshipPrefixReport(prefix: string) {
  const wb = new Workbook();
  wb.addWorksheet('S1').getCell('A1').value = 'hello';
  // `r` is the control: the package exactly as written, which is the prefix the writer already chose.
  const written = writeXlsx(wb);
  const back =
    prefix === 'r'
      ? readXlsx(written)
      : reloadPatched(written, {
          'xl/workbook.xml': (xml) =>
            xml.replace('xmlns:r=', `xmlns:${prefix}=`).replaceAll('r:id=', `${prefix}:id=`),
        });
  return {
    sheetNames: back.worksheets.map((sheet) => sheet.name),
    a1: back.getWorksheet('S1')?.getCell('A1').value ?? null,
  };
}

/**
 * Rewrite a worksheet part so the MAIN namespace is bound to a prefix and every element carries it,
 * then read → `{a1, validations, conditionalFormats}`. Every element name then has a colon in it, so a
 * reader using "has a prefix" to mean "is an extension element" discarded every data validation and
 * every conditional format in the file.
 */
export function mainNamespacePrefixReport(prefix: string) {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 1;
  ws.addDataValidation('A1', {type: 'list', formulae: ['"a,b,c"']});
  ws.addConditionalFormatting({ref: 'A1:A5', rules: [{type: 'duplicateValues', priority: 1}]});

  const back = reloadPatched(writeXlsx(wb), {
    'xl/worksheets/sheet1.xml': (xml) =>
      xml
        .replace(
          `<worksheet xmlns="${SPREADSHEETML}"`,
          `<${prefix}:worksheet xmlns:${prefix}="${SPREADSHEETML}"`,
        )
        // Every remaining unprefixed element takes the same prefix; already-prefixed ones (the x14
        // extension block) are left exactly as they are, which is the point of the comparison.
        .replace(/<(\/?)([a-zA-Z][\w]*)(?=[ />])/g, (match, slash: string, tag: string) =>
          tag === 'worksheet' ? match : `<${slash}${prefix}:${tag}`,
        )
        .replace(/<\/worksheet>/, `</${prefix}:worksheet>`),
  }).getWorksheet('S');
  return {
    a1: back?.getCell('A1').value ?? null,
    validations: (back?.dataValidations ?? []).map((entry) => entry.sqref),
    conditionalFormats: (back?.conditionalFormattings ?? []).flatMap((block) =>
      block.rules.map((rule) => rule.type),
    ),
  };
}

/**
 * Replace a package's theme with one written under `prefix` (`''` for the default DrawingML
 * namespace), read it back and report the resolved scheme plus a theme-indexed cell colour →
 * `{colors, fonts, cellColor}`. A theme under any prefix but `a` read as no theme at all, so the
 * workbook fell back to the library's default and every theme-indexed colour in the file resolved to
 * the wrong RGB: a whole-document visual divergence with nothing reported.
 *
 * The `lt1` slot carries an XML comment between it and its colour, which the old scanner's
 * whitespace-only pattern also lost even under the conventional prefix.
 */
export function themePrefixReport(prefix: string) {
  const q = prefix === '' ? '' : `${prefix}:`;
  const ns = prefix === '' ? `xmlns="${DRAWINGML}"` : `xmlns:${prefix}="${DRAWINGML}"`;
  const theme =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<${q}theme ${ns} name="Branded"><${q}themeElements>` +
    `<${q}clrScheme name="Branded">` +
    `<${q}dk1><${q}sysClr val="windowText" lastClr="111111"/></${q}dk1>` +
    `<${q}lt1><!-- a comment the old scanner could not pass --><${q}srgbClr val="222222"/></${q}lt1>` +
    `<${q}dk2><${q}srgbClr val="333333"/></${q}dk2>` +
    `<${q}lt2><${q}srgbClr val="444444"/></${q}lt2>` +
    `<${q}accent1><${q}srgbClr val="AABBCC"/></${q}accent1>` +
    `<${q}accent2><${q}srgbClr val="555555"/></${q}accent2>` +
    `<${q}accent3><${q}srgbClr val="666666"/></${q}accent3>` +
    `<${q}accent4><${q}srgbClr val="777777"/></${q}accent4>` +
    `<${q}accent5><${q}srgbClr val="888888"/></${q}accent5>` +
    `<${q}accent6><${q}srgbClr val="999999"/></${q}accent6>` +
    `<${q}hlink><${q}srgbClr val="0000FF"/></${q}hlink>` +
    `<${q}folHlink><${q}srgbClr val="800080"/></${q}folHlink>` +
    `</${q}clrScheme>` +
    `<${q}fontScheme name="Branded">` +
    `<${q}majorFont><${q}latin typeface="Major &amp; Co"/></${q}majorFont>` +
    `<${q}minorFont><${q}latin typeface="Minor"/></${q}minorFont>` +
    `</${q}fontScheme>` +
    `</${q}themeElements></${q}theme>`;

  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  // A cell whose font colour is theme slot 4 (accent1), so the resolved RGB says which theme won.
  ws.getCell('A1').value = 'x';
  ws.getCell('A1').font = {color: {theme: 4}};
  const back = reloadPatched(writeXlsx(wb), {'xl/theme/theme1.xml': () => theme});
  return {
    colors: {...back.themeColors},
    fonts: {...back.themeFonts},
    cellColor: back.getWorksheet('S')?.getCell('A1').font?.color ?? null,
  };
}

/**
 * Patch a workbook part so its `<definedNames>` opens with a self-closing `<definedName name="Empty"/>`
 * before a real one, read → `{names}`. A defined name with no `refersTo` is legal markup that fires no
 * close event, and the hand-rolled text capture that read these committed only on the close: the name
 * was dropped entirely, and the capture stayed latched on whatever text came next.
 */
export function selfClosingDefinedNameReport() {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 1;
  wb.defineName({name: 'Later', refersTo: 'S!$A$1'});
  const back = reloadPatched(writeXlsx(wb), {
    'xl/workbook.xml': (xml) =>
      xml.replace('<definedNames>', '<definedNames><definedName name="Empty"/>'),
  });
  return {
    names: back.definedNames.map((name) => ({
      name: name.name,
      refersTo: name.refersTo,
    })),
  };
}
