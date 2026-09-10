// Threaded comments, legacy notes, and rich text.

import {canonicalJson} from '../../canonical-json.ts';
import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {
  commentThreadFacts,
  packagePartFacts,
  partBytesOf,
  partMapOf,
  partNamesOf,
  patchedPackage,
  roundtrip,
} from './package-facts.ts';
import {
  isRichTextValue,
  readFixture,
  readXlsx,
  Workbook,
  type WorkbookInstance,
  writeXlsx,
} from './runtime.ts';
import {reloadPatched} from './xml-probes.ts';

export const comments = {
  // Author a control character into each of the other places cell text is carried, a legacy note's
  // body and a rich-text run, then write and read back → { note, runs }. These reach the same `<t>`
  // as a plain string cell, through different readers, so each needs its own decode wired.
  escapedTextCarrierRoundtrip(text: string) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 'x';
    sheet.getCell('A1').note = text;
    sheet.getCell('A2').value = {richText: [{text}, {text: `${text}!`, font: {bold: true}}]};
    const reloaded = roundtrip(workbook).worksheets[0];
    const richValue = reloaded?.getCell('A2').value;
    return {
      note: reloaded?.getCell('A1').note ?? null,
      runs:
        richValue !== undefined && isRichTextValue(richValue)
          ? richValue.richText.map((run) => run.text)
          : null,
    };
  },

  // Read a fixture whose threaded-comment bodies hold `_xHHHH_` escapes → { eager, roundtrip }. Each is
  // a map of a thread's anchor to its single message's text. `eager` is the fixture as read; `roundtrip`
  // is the same model written back through our own writer and re-read, which is what holds the escape and
  // the decode to being inverses rather than two independently plausible transformations.
  threadedCommentEscapeReport(rel: string, refs: string[]) {
    const textAt = (workbook: WorkbookInstance) => {
      const sheet = workbook.worksheets[0];
      return Object.fromEntries(
        refs.map((ref) => [ref, sheet?.commentThreadAt(ref)?.comments[0]?.text ?? null]),
      );
    };
    const source = readFixture(rel);
    return {eager: textAt(source), roundtrip: textAt(roundtrip(source))};
  },

  // Author `text` into a threaded comment, write, and read back → { emittedText, readText, rawInPart }.
  // `emittedText` is the body exactly as it reached the 2018 part, `rawInPart` says whether any emitted
  // part carries the character verbatim, which would make the package malformed XML, and `readText` is
  // what the reader gives back. The legacy fallback the writer builds beside the thread carries the same
  // text through a `<t>`, so a package that escaped one and not the other would fail `rawInPart`.
  authoredThreadedCommentEscape(text: string) {
    // oxlint-disable-next-line eslint/no-control-regex -- matching the control characters is the check: this asks whether an emitted part carries one
    const raw = /[\u{0}-\u{8}\u{B}\u{C}\u{E}-\u{1F}\u{FFFE}\u{FFFF}\u{D800}-\u{DFFF}]/u;
    const AUTHOR = '{39236F6F-643D-4654-8264-DD21C8472F7F}';
    const workbook = new Workbook();
    workbook.addPerson({id: AUTHOR, displayName: 'Ada Lovelace', providerId: 'AD'});
    const sheet = workbook.addWorksheet('Review');
    sheet.addCommentThread({
      ref: 'A1',
      resolved: false,
      comments: [
        {id: '{11111111-2222-3333-4444-555555555555}', personId: AUTHOR, text, mentions: []},
      ],
    });
    const bytes = writeXlsx(workbook);
    const parts = partMapOf(bytes);
    const threadPart =
      Object.entries(parts).find(([name]) => name.includes('threadedComments/'))?.[1] ?? '';
    return {
      emittedText: /<text[^>]*>([\s\S]*?)<\/text>/.exec(threadPart)?.[1] ?? null,
      rawInPart: Object.keys(parts)
        .filter((name) => raw.test(parts[name] ?? ''))
        .sort(),
      readText: readXlsx(bytes).worksheets[0]?.commentThreadAt('A1')?.comments[0]?.text ?? null,
    };
  },

  // Write a noted cell, relocate its comments part to a non-canonical path (xl/sheet1_comments.xml)
  // reachable only through the worksheet rels, and reload → { ok, error, note }. The reader locates the
  // comments part by relationship *type*, not by filename glob, so the moved part still loads and its
  // note reads back.
  nonCanonicalCommentsPartReport() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'x';
    ws.getCell('A1').note = 'hi';
    const written = writeXlsx(wb);
    const commentPart = partNamesOf(written).find((n) => /^xl\/comments\d*\.xml$/.test(n));
    if (commentPart === undefined) throw new Error('expected a comments part for the note on A1');
    const buffer = patchedPackage(written, {
      edit: {
        'xl/worksheets/_rels/sheet1.xml.rels': (xml) =>
          xml.replace(/Target="[^"]*comments\d*\.xml"/i, 'Target="../sheet1_comments.xml"'),
      },
      drop: [commentPart],
      put: {'xl/sheet1_comments.xml': partBytesOf(written, commentPart)},
    });
    let ok = false;
    let error = null;
    let note = null;
    try {
      note = readXlsx(buffer).worksheets[0]!.getCell('A1').note ?? null;
      ok = true;
    } catch (e) {
      error = messageOf(e);
    }
    return {ok, error, note};
  },

  // Attach then clear a cell note and report whether the written package still carries a comment/VML
  // artifact → { commentPartPresent, vmlPartPresent, readNoteAfter, neighborNoteIntact,
  // cleanHasCommentPart }. Clearing a note (note = undefined) is a genuine removal: the sole-noted
  // cell's comment part disappears and reads back null, while a neighbor that kept its note is intact,
  // and a workbook that never had a note emits no comment part at all.
  removeCellNoteReport() {
    const partNames = (wb: WorkbookInstance) => Object.keys(partMapOf(writeXlsx(wb)));
    // The note being cleared is the only one, so a lingering comment part would prove the removal
    // merely emptied the note rather than dropping it.
    const solo = new Workbook();
    const soloSheet = solo.addWorksheet('S');
    soloSheet.getCell('A1').value = 'x';
    soloSheet.getCell('A1').note = 'remove me';
    soloSheet.getCell('A1').note = undefined;
    const soloParts = partNames(solo);
    const commentPartPresent = soloParts.some((f) => /comments\d*\.xml$/.test(f));
    const vmlPartPresent = soloParts.some((f) => /vmlDrawing\d*\.vml$/.test(f));
    const readNoteAfter = roundtrip(solo).worksheets[0]!.getCell('A1').note ?? null;

    // Clearing one note must not disturb another cell's note.
    const pair = new Workbook();
    const pairSheet = pair.addWorksheet('S');
    pairSheet.getCell('A1').value = 'x';
    pairSheet.getCell('A1').note = 'remove me';
    pairSheet.getCell('B1').value = 'y';
    pairSheet.getCell('B1').note = 'keep me';
    pairSheet.getCell('A1').note = undefined;
    const reloaded = roundtrip(pair).worksheets[0]!;
    const neighborNoteIntact = !!reloaded.getCell('B1').note;

    const clean = new Workbook();
    clean.addWorksheet('S').getCell('A1').value = 'x';
    const cleanHasCommentPart = partNames(clean).some((f) => /comments\d*\.xml$/.test(f));

    return {
      commentPartPresent,
      vmlPartPresent,
      readNoteAfter,
      neighborNoteIntact,
      cleanHasCommentPart,
    };
  },

  // Write a rich-text cell, then read it back, and report how its runs serialized and survived →
  // { emptyTextRunInXml, runCount, runs: [{text, bold, italic, underline}] }. Mirrors the oracle. A
  // zero-length run must never emit an empty <t> (Excel flags it corrupt); the surviving runs keep
  // their text and per-run formatting. The rewrite writes rich text inline, so both the empty-<t>
  // scan and the read-back target the worksheet XML (there is no shared-strings part).
  async richTextRoundtripReport(runs: Untyped) {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = {richText: runs};
    const buffer = writeXlsx(wb);
    const parts = partMapOf(buffer);
    const xml = parts['xl/sharedStrings.xml'] || parts['xl/worksheets/sheet1.xml'] || '';
    const emptyTextRunInXml = /<(?:\w+:)?t\b[^>]*\/>|<(?:\w+:)?t\b[^>]*><\/(?:\w+:)?t>/.test(xml);
    const value = readXlsx(buffer).getWorksheet('S')!.getCell('A1').value;
    // The library's own guard rather than a duck-type on `.richText`: what a case wants to know is
    // whether the value reads back as rich text *by the same test a consumer would apply*.
    const readRuns = value !== null && isRichTextValue(value) ? value.richText : [];
    return {
      emptyTextRunInXml,
      runCount: readRuns.length,
      runs: readRuns.map((r) => ({
        text: r.text ?? null,
        bold: r.font ? (r.font.bold ?? false) : false,
        italic: r.font ? (r.font.italic ?? false) : false,
        underline: r.font ? (r.font.underline ?? false) : false,
      })),
    };
  },

  // Read a fixture and report the modern threaded conversations the reader reconstructs, see
  // {@link commentThreadFacts} for the shape.
  readFixtureCommentThreads(rel: string, refs: string[] = []) {
    return commentThreadFacts(readFixture(rel), refs);
  },

  // Author a conversation in the MODEL, never read from a file, write it, and read the package back:
  // → { parts, deterministic, model }. Nothing here comes from preserved bytes, so this is what proves the
  // writer serialises the feature rather than merely carrying it: the thread part, the workbook person
  // registry, and the legacy fallback comment are all built from the model on the way out and reassembled
  // into threads on the way in. `deterministic` re-writes the same model and compares the bytes: the
  // writer holds no clock and no id generator, so every guid and timestamp comes from the caller.
  //
  // The conversation deliberately mixes every facet at once (a reply by a second author, an @mention
  // resolving through a third registry entry, a resolved thread, and a genuine legacy note on another
  // cell) because those are the facets that interact: the fallback text folds the reply in, the mention
  // spans the message text, and the note must not be mistaken for the thread's fallback.
  authoredCommentThreadRoundtrip() {
    const ADA = '{39236F6F-643D-4654-8264-DD21C8472F7F}';
    const GRACE = '{1B2C3D4E-5F60-4A71-8B92-0C1D2E3F4A5B}';
    const GRACE_MENTIONED = '{BA397017-DD76-4496-AA75-59ADB199950C}';
    const workbook = new Workbook();
    workbook.addPerson({id: ADA, displayName: 'Ada Lovelace', providerId: 'AD'});
    workbook.addPerson({id: GRACE, displayName: 'Grace Hopper', providerId: 'AD'});
    // The same human as GRACE under a second id: how Excel registers a *mentioned* identity.
    workbook.addPerson({
      id: GRACE_MENTIONED,
      displayName: 'Grace Hopper',
      providerId: 'PeoplePicker',
    });
    const sheet = workbook.addWorksheet('Review');
    sheet.getCell('B2').value = 1234;
    sheet.getCell('D4').note = 'an ordinary note beside the conversation';
    sheet.addCommentThread({
      // Absolute on purpose: the anchor is canonicalised, so `$B$2` and `B2` are one cell.
      ref: '$B$2',
      resolved: true,
      comments: [
        {
          id: '{11111111-2222-3333-4444-555555555555}',
          personId: ADA,
          date: '2026-07-26T12:00:00.00',
          text: '@Grace Hopper is this gross or net?',
          mentions: [
            {
              personId: GRACE_MENTIONED,
              mentionId: '{3F2C1A9E-5B84-4D67-9C2E-71A0D4E8B531}',
              startIndex: 0,
              length: 13,
            },
          ],
        },
        {
          id: '{66666666-7777-8888-9999-AAAAAAAAAAAA}',
          personId: GRACE,
          date: '2026-07-26T12:05:00.00',
          text: 'Gross. Confirmed with finance.',
          mentions: [],
        },
      ],
    });
    const bytes = writeXlsx(workbook);
    return {
      parts: packagePartFacts(partMapOf(bytes)),
      deterministic: Buffer.compare(Buffer.from(bytes), Buffer.from(writeXlsx(workbook))) === 0,
      model: commentThreadFacts(readXlsx(bytes), ['B2', '$B$2', 'D4']),
    };
  },
  // Read one and the same rich string twice, once pooled in a `<si>` and once inline in an `<is>`,
  // and report each reading -> { pooled, inline, identical }. The two are separate grammars in the
  // format and were separate copies of one grammar in the reader, kept in step only by a comment
  // saying they agreed. Nothing measured it: each reader is exercised through its own path, so a
  // drift would surface as the same cell content reading two ways, decided by an encoding choice the
  // author never made.
  richStringPooledVersusInlineReport() {
    const runs =
      '<r><rPr><b/><sz val="14"/><color rgb="FFFF0000"/><rFont val="Arial"/></rPr><t>bold red</t></r>' +
      '<r><t xml:space="preserve"> plain </t></r>' +
      '<r><rPr><i/><u/></rPr><t>italic</t></r>';

    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 'placeholder';
    // Pooled, so A1 is already `t="s"` at ordinal 0 and the pool part exists with its relationship.
    const bytes = writeXlsx(wb, {useSharedStrings: true});

    // Both packages are the same bytes with one part rewritten, so anything that differs between the
    // two readings came from the `<si>`/`<is>` split and nothing else.
    const pooled = reloadPatched(bytes, {
      'xl/sharedStrings.xml': () =>
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">' +
        `<si>${runs}</si></sst>`,
    });
    const inline = reloadPatched(bytes, {
      'xl/worksheets/sheet1.xml': (xml) =>
        xml.replace(/<c r="A1"[^>]*>[\s\S]*?<\/c>/, `<c r="A1" t="inlineStr"><is>${runs}</is></c>`),
    });

    const readingOf = (wbk: WorkbookInstance) =>
      JSON.stringify(canonicalJson(wbk.getWorksheet('S')!.getCell('A1').value));
    const pooledReading = readingOf(pooled);
    const inlineReading = readingOf(inline);
    return {
      pooled: pooledReading,
      inline: inlineReading,
      identical: pooledReading === inlineReading,
    };
  },

  // Put a phonetic run (`<rPh>`, the furigana Japanese Excel stores beside a string's base text) in
  // each place a `CT_Rst` carries one: a pooled plain string on A1, a pooled rich string on A2, an
  // inline string on A3 and the note on A1. Read it, then write and read again → { read, rewritten },
  // each { pooledPlain, pooledRichRuns, inline, note }, where the base text of every one is `漢字`
  // except the note's, which is `note`.
  phoneticRunReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('A1').value = 'one';
    sheet.getCell('A2').value = 'two';
    sheet.getCell('A3').value = 'three';
    sheet.getCell('A1').note = 'note';
    const written = writeXlsx(wb, {useSharedStrings: true});
    const notesPart = partNamesOf(written).find((name) => /^xl\/comments\d*\.xml$/.test(name));
    if (notesPart === undefined) throw new Error('expected a comments part for the note on A1');
    const phonetic =
      '<rPh sb="0" eb="2"><t>かんじ</t></rPh><phoneticPr fontId="0" type="noConversion"/>';
    const patched = patchedPackage(written, {
      edit: {
        'xl/sharedStrings.xml': () =>
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">' +
          `<si><t>漢字</t>${phonetic}</si>` +
          `<si><r><rPr><b/></rPr><t>漢字</t></r>${phonetic}</si></sst>`,
        'xl/worksheets/sheet1.xml': (xml) =>
          xml.replace(
            /<c r="A3"[^>]*>[\s\S]*?<\/c>/,
            `<c r="A3" t="inlineStr"><is><t>漢字</t>${phonetic}</is></c>`,
          ),
        // Appended to the note's own body, whatever runs the writer spelled that body in.
        [notesPart]: (xml) => xml.replace('</text>', `${phonetic}</text>`),
      },
    });
    const readings = (workbook: WorkbookInstance) => {
      const s = workbook.getWorksheet('S')!;
      const rich = s.getCell('A2').value;
      return {
        pooledPlain: s.getCell('A1').value,
        pooledRichRuns: isRichTextValue(rich) ? rich.richText.map((run) => run.text) : null,
        inline: s.getCell('A3').value,
        note: s.getCell('A1').note ?? null,
      };
    };
    const reread = readXlsx(patched);
    return {read: readings(reread), rewritten: readings(roundtrip(reread))};
  },
};
