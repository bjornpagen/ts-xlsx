// Workbook, worksheet and cell protection.

import {strToU8, zipSync} from 'fflate';

import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {partMapOf} from './package-facts.ts';
import {decodeAddress, readXlsx, Workbook, writeXlsx} from './runtime.ts';
import {attrsOf, decodeXmlEntities, reloadPatched, xmlWellFormed} from './xml-probes.ts';

// An attribute's value as the author wrote it, or null when the element or attribute is missing.
const decode = (raw: string | null | undefined) =>
  raw === null || raw === undefined ? null : decodeXmlEntities(raw);

export const protection = {
  // Build a workbook, inject workbook-level structure protection into its workbook.xml (reproducing a
  // file locked with <workbookProtection lockStructure="1">), then read it back and write it out again.
  // Report whether the protection survives the read→write round-trip rather than being silently dropped.
  workbookProtectionRoundtrip() {
    const base = new Workbook();
    base.addWorksheet('S').getCell('A1').value = 'x';
    const parts = partMapOf(writeXlsx(base));
    const injectedXml = parts['xl/workbook.xml']!.replace(
      /<sheets>/,
      '<workbookProtection lockStructure="1" lockWindows="0"/><sheets>',
    );
    const zipFiles: Record<string, Untyped> = {};
    for (const [name, xml] of Object.entries(parts)) {
      zipFiles[name] = strToU8(name === 'xl/workbook.xml' ? injectedXml : xml);
    }
    const injected = zipSync(zipFiles);

    const rewrittenBuffer = writeXlsx(readXlsx(injected));
    const rewrittenXml = partMapOf(rewrittenBuffer)['xl/workbook.xml'] ?? '';
    return {
      sourceHadProtection: /workbookProtection/.test(injectedXml),
      rewrittenHasProtection: /workbookProtection/.test(rewrittenXml),
      rewrittenLocksStructure: /lockStructure="1"/.test(rewrittenXml),
    };
  },

  // Author per-cell protection (and optionally protect the sheet), write, then read back →
  // { readBack, hasApplyProtection, sheetProtection, sheetProtectionAttrs }. Reports whether an
  // explicitly-unlocked cell round-trips as locked=false, whether the style record carries the
  // flag (applyProtection + <protection> in cellXfs), and the emitted <sheetProtection> that
  // makes the locked flags enforceable.
  authorCellProtection(
    cells: Untyped = [],
    protect: Untyped = null,
    {rows = [], columns = []}: Untyped = {},
  ) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    for (const c of cells) {
      const cell = sheet.getCell(c.ref);
      cell.value = c.value ?? c.ref;
      if (c.protection !== undefined) cell.protection = c.protection;
    }
    // Whole-column / whole-row protection: the model carries protection per cell, so realize an
    // unlocked band by stamping its flag onto each listed cell that falls in the band: the same
    // end-state a per-cell override yields (column-scope inheritance is a separate capability).
    // Applied after the per-cell settings so the band-level flag is what the case asserts.
    for (const col of columns) {
      for (const c of cells)
        if (decodeAddress(c.ref).col === col.index)
          sheet.getCell(c.ref).protection = col.protection;
    }
    for (const r of rows) {
      for (const c of cells)
        if (decodeAddress(c.ref).row === r.index) sheet.getCell(c.ref).protection = r.protection;
    }
    if (protect) sheet.protect(protect.password ?? undefined, protect.options ?? {});
    const buffer = writeXlsx(workbook);
    const parts = partMapOf(buffer);
    const styles = parts['xl/styles.xml'] || '';
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const sheetProtection = (sheetXml.match(/<sheetProtection\b[^>]*\/?>/) || [])[0] || null;

    const reread = readXlsx(buffer);
    const sheet2 = reread.getWorksheet('S')!;
    const readBack: Record<string, Untyped> = {};
    for (const c of cells) {
      const p = sheet2.getCell(c.ref).protection;
      readBack[c.ref] = p ? {locked: p.locked ?? null} : null;
    }
    return {
      readBack,
      hasApplyProtection: /applyProtection="1"/.test(styles) && /<protection\b/.test(styles),
      sheetProtection,
      sheetProtectionAttrs: sheetProtection ? attrsOf(sheetProtection) : null,
    };
  },

  // Password-protect a worksheet twice under Node and report the emitted protection facts →
  // { threw, algorithm, hasHash, hasSalt, spinCount, selectLockedCells, selectUnlockedCells,
  // saltsDiffer }. Proves protect succeeds without a browser-random error, emits a well-formed
  // password credential, honors the requested options, and salts with real randomness (two
  // protects with the same password differ).
  worksheetPasswordProtectionReport(password = 'secret') {
    const protectOnce = () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      ws.protect(password, {selectLockedCells: false, selectUnlockedCells: false});
      const xml = partMapOf(writeXlsx(wb))['xl/worksheets/sheet1.xml'] || '';
      return (xml.match(/<sheetProtection\b[^>]*\/>/) || [''])[0];
    };
    let first: string;
    let second: string;
    try {
      first = protectOnce();
      second = protectOnce();
    } catch (e) {
      return {
        threw: messageOf(e),
        algorithm: null,
        hasHash: false,
        hasSalt: false,
        spinCount: null,
        selectLockedCells: null,
        selectUnlockedCells: null,
        saltsDiffer: false,
      };
    }
    const a = attrsOf(first);
    const b = attrsOf(second);
    return {
      threw: null,
      algorithm: a.algorithmName ?? null,
      hasHash: !!a.hashValue,
      hasSalt: !!a.saltValue,
      spinCount: a.spinCount ?? null,
      selectLockedCells: a.selectLockedCells ?? null,
      selectUnlockedCells: a.selectUnlockedCells ?? null,
      saltsDiffer: !!a.saltValue && !!b.saltValue && a.saltValue !== b.saltValue,
    };
  },

  // Protect a worksheet, write it, read it back, then write the reloaded workbook again, reporting
  // the <sheetProtection> attributes from BOTH writes → { first, second }. Proves the reader carries
  // sheet-level protection back into the model rather than silently dropping it on a passthrough
  // save: the second write must still emit protection, with the agile credential preserved verbatim
  // (no plaintext password survives to re-hash) and the permissive flags intact.
  sheetProtectionRoundtrip(
    password = 'secret',
    options = {sort: true, autoFilter: true, selectLockedCells: false},
  ) {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'x';
    ws.protect(password ?? undefined, options ?? {});

    const buf1 = writeXlsx(wb);
    const buf2 = writeXlsx(readXlsx(buf1));
    const protAttrs = (buf: Untyped) => {
      const xml = partMapOf(buf)['xl/worksheets/sheet1.xml'] || '';
      const el = (xml.match(/<sheetProtection\b[^>]*\/?>/) || [])[0];
      return el ? attrsOf(el) : null;
    };
    return {first: protAttrs(buf1), second: protAttrs(buf2)};
  },

  // Copy-on-write style aliasing family. Each cell owns its facet fields and every setter REPLACES
  // the field (the readonly facet types forbid in-place mutation of a shared record), so mutating
  // one cell's facet, even a cell that shared a style with siblings on disk, cannot bleed onto a
  // sibling. These methods prove that end-to-end through the real write→read path.

  // Inject an XML special into each of the three foreign values a package can carry into a place the
  // writer re-emits verbatim: a <sheetProtection> agile-hash credential, a preserved part's path and
  // content type, and a preserved workbook relationship's target. Then read the package and write it
  // back → { source: {...the values put in}, rewritten: {wellFormed, algorithmName, spinCount,
  // overridePartName, overrideContentType, relTarget} }. Use it to assert a value taken off an
  // untrusted package survives the round-trip as itself: escaped once on the way out (so the part
  // still parses), never twice (so a preserved relationship still resolves to the part it names).
  foreignValueRewriteReport() {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 'x';
    const parts = partMapOf(writeXlsx(wb));

    // A slicer cache is preserved wholesale by relationship type, which is what carries an arbitrary
    // foreign path and content type through the model and back out again.
    const relType = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache';
    const algorithmName = 'SHA-512 & "friends"';
    const slicerPath = 'xl/slicerCaches/s&1.xml';
    const slicerContentType = 'application/vnd.ms-excel.slicerCache+xml; q="1" & p=<2>';

    // `<sheetProtection>` follows `<sheetData>` in CT_Worksheet order.
    parts['xl/worksheets/sheet1.xml'] = parts['xl/worksheets/sheet1.xml']!.replace(
      '</sheetData>',
      `</sheetData><sheetProtection algorithmName="${algorithmName.replace(/&/g, '&amp;').replace(/"/g, '&quot;')}"` +
        ' hashValue="aGFzaA==" saltValue="c2FsdA==" spinCount="100000" sheet="1"/>',
    );
    parts['xl/_rels/workbook.xml.rels'] = parts['xl/_rels/workbook.xml.rels']!.replace(
      '</Relationships>',
      `<Relationship Id="rIdSlicer" Type="${relType}" Target="slicerCaches/s&amp;1.xml"/></Relationships>`,
    );
    parts['[Content_Types].xml'] = parts['[Content_Types].xml']!.replace(
      '</Types>',
      `<Override PartName="/${slicerPath.replace(/&/g, '&amp;')}" ContentType="${slicerContentType
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/"/g, '&quot;')}"/></Types>`,
    );
    parts[slicerPath] = '<slicerCacheDefinition/>';

    const zipFiles: Record<string, Untyped> = {};
    for (const [name, text] of Object.entries(parts)) zipFiles[name] = strToU8(text);

    const rewritten = partMapOf(writeXlsx(readXlsx(zipSync(zipFiles))));
    const contentTypes = rewritten['[Content_Types].xml'] ?? '';
    const override = (contentTypes.match(/<Override PartName="[^"]*slicerCache[^"]*"[^>]*>/) ??
      [])[0];
    const overrideAttrs = override === undefined ? null : attrsOf(override);
    const sheetXml = rewritten['xl/worksheets/sheet1.xml'] ?? '';
    const protectionEl = (sheetXml.match(/<sheetProtection\b[^>]*\/?>/) ?? [])[0];
    const workbookRels = rewritten['xl/_rels/workbook.xml.rels'] ?? '';
    const relEl = (workbookRels.match(/<Relationship[^>]*slicerCache[^>]*>/) ?? [])[0];

    return {
      source: {algorithmName, slicerPath, slicerContentType},
      rewritten: {
        wellFormed: [contentTypes, sheetXml, workbookRels].every((xml) => xmlWellFormed(xml)),
        algorithmName: decode(
          protectionEl === undefined ? null : attrsOf(protectionEl).algorithmName,
        ),
        spinCount: protectionEl === undefined ? null : (attrsOf(protectionEl).spinCount ?? null),
        overridePartName: decode(overrideAttrs?.PartName),
        overrideContentType: decode(overrideAttrs?.ContentType),
        relTarget: decode(relEl === undefined ? null : attrsOf(relEl).Target),
      },
    };
  },

  // Put `<sheetProtection ${attributes}/>` into a written sheet, read it, and write it again →
  // { protected, legacyPasswordHash, rewritten }, where `rewritten` is the attributes of the
  // `<sheetProtection>` the second write emitted, or null when it emitted none.
  legacySheetProtectionReport(attributes: string) {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 'x';
    const read = reloadPatched(writeXlsx(wb), {
      // `<sheetProtection>` follows `<sheetData>` in CT_Worksheet order.
      'xl/worksheets/sheet1.xml': (xml) =>
        xml.replace('</sheetData>', `</sheetData><sheetProtection ${attributes}/>`),
    });
    const sheetProtection = read.getWorksheet('S')?.protection;
    const rewrittenXml = partMapOf(writeXlsx(read))['xl/worksheets/sheet1.xml'] ?? '';
    const element = rewrittenXml.match(/<sheetProtection\b[^>]*\/>/)?.[0];
    return {
      protected: sheetProtection !== undefined,
      legacyPasswordHash: sheetProtection?.legacyPasswordHash ?? null,
      rewritten: element === undefined ? null : attrsOf(element),
    };
  },
};
