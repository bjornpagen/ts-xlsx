// Cluster: xlsx-io
//
// Real-world scenario: a producer writes a `<c>` whose `t` names a type and then leaves the payload
// out, or leaves it empty: `<c r="A1" t="b"/>`, `<c t="s"><v></v></c>`, `<c t="inlineStr"/>`. Excel
// opens such a file without comment and shows every one of those cells as blank. A reader that
// decodes by `t` alone invents a value instead: `FALSE` for the boolean, an empty string for the
// others, and a save then writes the invention back as real data that `COUNTA` and `ISBLANK` see.
//
// The line Excel draws is not "no `<v>`". The two string types are text as soon as their carrier is
// present, so `<c t="str"><v/></c>` and `<c t="inlineStr"><is/></c>` are empty strings, while the same
// cells without the carrier are blank. Every other type is only its `<v>` text, and an empty `<v>` is
// the same blank as a missing one. The fixture holds each shape once; its expected readings are Excel
// Desktop's own, recorded in `test/corpus/fixtures/excel-oracle/typed-cell-without-value.json`.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'typed-cell-without-value-reads-blank/typed-cells.xlsx';

// Rows 1-14 are the shapes Excel shows as blank, rows 15-19 the empty strings, 20-21 the controls.
const BLANK = Array.from({length: 14}, (_, i) => `A${i + 1}`);
const EMPTY_TEXT = ['A15', 'A16', 'A17', 'A18', 'A19'];
const ALL = [...BLANK, ...EMPTY_TEXT, 'A20', 'A21'];

export default {
  id: 'typed-cell-without-value-reads-blank',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/typed-cell-without-value.json',
  },
  cluster: 'xlsx-io',
  description:
    'A `<c>` whose `t` names a type but whose payload is missing or empty reads the way Excel shows it: ' +
    'blank for every type whose `<v>` is missing or empty, and an empty string only for a `str` or ' +
    '`inlineStr` cell whose `<v>` or `<is>` is present. Both readers agree.',

  behavior: [
    {
      name: 'a typed cell with no `<v>`, or an empty one, reads as no value rather than FALSE or ""',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.readFixtureCells(FIXTURE, BLANK);
        for (const ref of BLANK) {
          assert.strictEqual(cells[ref].value, null, `${ref} is blank in Excel`);
        }
      },
    },
    {
      name: 'a string cell whose `<v>` or `<is>` is present but empty reads as the empty string',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.readFixtureCells(FIXTURE, EMPTY_TEXT);
        for (const ref of EMPTY_TEXT) {
          assert.strictEqual(cells[ref].value, '', `${ref} is empty text in Excel`);
        }
      },
    },
    {
      name: 'a typed cell that carries its value still reads it',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.readFixtureCells(FIXTURE, ['A20', 'A21']);
        assert.strictEqual(cells.A20.value, false);
        assert.strictEqual(cells.A21.value, 'pooled');
      },
    },
    {
      name: 'the streaming reader yields no cell for the blank shapes and agrees on the rest',
      async expect(api: CorpusApi, assert: Assert) {
        const streamed = await api.streamReadFixture(FIXTURE, ALL);
        const buffered = await api.readFixtureCells(FIXTURE, ALL);
        for (const ref of BLANK) {
          assert.strictEqual(streamed[ref], null, `${ref} is not a data cell`);
        }
        for (const ref of [...EMPTY_TEXT, 'A20', 'A21']) {
          assert.deepStrictEqual(streamed[ref]?.value, buffered[ref].value, `${ref} agrees`);
        }
      },
    },
  ],
} satisfies Case;
