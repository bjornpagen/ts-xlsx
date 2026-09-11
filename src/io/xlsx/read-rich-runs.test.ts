import assert from 'node:assert/strict';
import {test} from 'node:test';

import {RunAccumulator} from './read-rich-runs.ts';

// Drive the machine the way a parser does, so a test states a document rather than a call sequence.
// `<t>` and `<r>` are written non-self-closing unless a case is about the self-closing form, since a
// self-closing element fires no close in the SAX layer.
function feed(runs: RunAccumulator, events: readonly (readonly [string, ...unknown[]])[]): void {
  for (const [kind, ...rest] of events) {
    if (kind === 'open') runs.open(rest[0] as string, {}, (rest[1] as boolean) ?? false);
    else if (kind === 'text') runs.text(rest[0] as string);
    else runs.close(rest[0] as string);
  }
}

// A tiny document reader over one container, so a case reads as XML rather than as a call log.
function read(container: 'si' | 'is' | 'text', xml: string, readRuns = true) {
  const runs = new RunAccumulator({container, readRuns});
  // Deliberately not a real parser: a hand-rolled scan keeps this test independent of `parseXml`, so
  // a break in the machine cannot be masked by a break in the tokenizer (or vice versa).
  for (const token of xml.match(/<[^>]+>|[^<]+/g) ?? []) {
    if (token.startsWith('</')) runs.close(token.slice(2, -1));
    else if (token.startsWith('<')) {
      const selfClosing = token.endsWith('/>');
      const [name = ''] = token.slice(1, selfClosing ? -2 : -1).split(/\s/);
      runs.open(name, {}, selfClosing);
    } else runs.text(token);
  }
  return runs;
}

// One invariant carries this class, and it is invisible from any single caller: the accumulator is
// emptied when a container opens, so every reader that opens one has to say so. These pin it directly,
// rather than only through the two readers that happen to obey it today.

test('a second container discards the first container runs', () => {
  const runs = read('si', '<si><r><t>first</t></r></si><si><r><t>second</t></r></si>');
  assert.deepStrictEqual(
    runs.runs.map((run) => run.text),
    ['second'],
  );
});

test('a value built from one container keeps its own array across the next', () => {
  // Why beginContainer installs a new array instead of emptying the one it has: the consumer reads the
  // runs after the container closes, and a shared array would be cleared out from under it.
  const runs = new RunAccumulator({container: 'si', readRuns: true});
  feed(runs, [
    ['open', 'si'],
    ['open', 'r'],
    ['open', 't'],
    ['text', 'kept'],
    ['close', 't'],
    ['close', 'r'],
    ['close', 'si'],
  ]);
  const captured = runs.runs;

  feed(runs, [['open', 'si']]);
  assert.deepStrictEqual(
    captured.map((run) => run.text),
    ['kept'],
  );
});

test('a truncated container does not bleed its unterminated run into the next', () => {
  // A damaged `<is>` leaves a run open. The next container must not inherit it, or the first cell of
  // a broken sheet bleeds into the second.
  const runs = read('is', '<is><r><t>orphan</t></is><is><t>loose</t></is>');
  assert.deepStrictEqual(runs.runs, [], 'the fresh container has no run open');
  assert.strictEqual(runs.plainText, 'loose', 'so its text is the container own');
  assert.strictEqual(runs.isRich, false);
});

test('a t left open by truncation does not carry its text into the next container', () => {
  // One layer under the case above: the `<t>` itself is what never closes, so the capture stays armed
  // across the container boundary. The stray `</t>` in the second container is then answered with the
  // first container's buffer, and one cell's text surfaces as another's.
  const runs = read('is', '<is><t>orphan</is><is>loose</t></is>');
  assert.strictEqual(runs.plainText, '');
});

test('a self-closing container leaves no text latched onto the next element', () => {
  // `<si/>` is legal and fires no close, so a machine that only unlatched on `</si>` would still
  // believe it was inside a container. The stray `<t>` that follows would then be absorbed as that
  // container plain text and surface as a pooled string nobody wrote.
  const runs = read('si', '<si/><t>stray</t>');
  assert.strictEqual(runs.plainText, '', 'text outside every container belongs to no container');
  assert.strictEqual(runs.isRich, false);
});

test('a run keeps only the font facets its own rPr set', () => {
  const runs = read('si', '<si><r><rPr><b/></rPr><t>bold</t></r><r><t>plain</t></r></si>');
  assert.deepStrictEqual(runs.runs, [{text: 'bold', font: {bold: true}}, {text: 'plain'}]);
});

test('an rPr that sets no facet leaves the run without a font', () => {
  const runs = read('si', '<si><r><rPr></rPr></r></si>');
  assert.deepStrictEqual(runs.runs, [{text: ''}]);
});

// The old shape treated every unrecognised element as a run property. Refusing one outside an open
// `<rPr>` is what lets the machine tell the caller "not mine" instead of silently swallowing it.
test('an rPr child outside a run is refused rather than fabricating one', () => {
  const runs = new RunAccumulator({container: 'si', readRuns: true});
  feed(runs, [['open', 'si']]);
  assert.strictEqual(runs.open('b', {}, true), false, 'no rPr is open, so <b/> is not ours');
  assert.deepStrictEqual(runs.runs, []);
});

// The font draft outlives its `</rPr>` (the run commits it at `</r>`), so "a draft exists" is not
// the same question as "an rPr is open". Asking the wrong one absorbs the next element in the
// document as a font facet, which for a worksheet body means eating the `<c>` after a rich `<is>`
// and giving that cell the previous one's address and style.
test('the element after a closed rPr is not absorbed as a font facet', () => {
  const runs = new RunAccumulator({container: 'is', readRuns: true});
  feed(runs, [
    ['open', 'is'],
    ['open', 'r'],
    ['open', 'rPr'],
    ['open', 'b', true],
    ['close', 'rPr'],
    ['open', 't'],
    ['text', 'x'],
    ['close', 't'],
    ['close', 'r'],
    ['close', 'is'],
  ]);
  assert.deepStrictEqual(runs.runs, [{text: 'x', font: {bold: true}}]);
  assert.strictEqual(runs.open('c', {}, true), false, 'the next cell is the caller own');
});

test('an r outside the container is not a rich-text run', () => {
  const runs = new RunAccumulator({container: 'is', readRuns: true});
  assert.strictEqual(runs.open('r', {}, false), false);
  assert.strictEqual(runs.close('r'), 'other');
});

test('a self-closing t contributes no text and arms no capture for the next element', () => {
  // `<t/>` is legal and fires no close. A raw capture flag would stay armed and swallow whatever
  // character data came next; the capture this machine owns cannot.
  const runs = read('si', '<si><r><t/><t>real</t></r></si>');
  assert.deepStrictEqual(runs.runs, [{text: 'real'}]);
});

test('an empty t contributes an empty string rather than being skipped', () => {
  const runs = read('si', '<si><t></t></si>');
  assert.strictEqual(runs.plainText, '');
  assert.strictEqual(runs.isRich, false);
});

test('a t inside a run beats the container plain text', () => {
  const runs = read('is', '<is><t>bare</t><r><t>in-run</t></r></is>');
  assert.deepStrictEqual(runs.runs, [{text: 'in-run'}]);
  assert.strictEqual(runs.plainText, 'bare');
  assert.strictEqual(runs.isRich, true);
});

test('the _xHHHH_ escape is decoded once per whole t, not per character-data chunk', () => {
  // An entity splits a SAX text run, so `_x00` and `01_` arrive separately. Decoding per chunk would
  // miss the escape in exactly the strings that contain an `&`.
  const runs = new RunAccumulator({container: 'si', readRuns: true});
  feed(runs, [
    ['open', 'si'],
    ['open', 't'],
    ['text', '_x00'],
    ['text', '01_'],
    ['close', 't'],
    ['close', 'si'],
  ]);
  assert.strictEqual(runs.plainText, '\u0001');
});

test('close reports the container so a caller knows when its string is complete', () => {
  const runs = new RunAccumulator({container: 'si', readRuns: true});
  runs.open('si', {}, false);
  assert.strictEqual(runs.close('t'), 'claimed');
  assert.strictEqual(runs.close('si'), 'container');
  assert.strictEqual(runs.close('row'), 'other');
});

// ── Phonetic runs ────────────────────────────────────────────────────────────────────────────────────
// `<rPh>` holds the furigana Japanese Excel stores beside a string's base text, in the same container
// as that text. Gathering every `<t>` read `漢字` as `漢字かんじ`, pooled, inline and in a note alike.

const PHONETIC =
  '<rPh sb="0" eb="2"><t>かんじ</t></rPh><phoneticPr fontId="0" type="noConversion"/>';

test('a pooled plain string reads its base text without the phonetic run beside it', () => {
  const runs = read('si', `<si><t>漢字</t>${PHONETIC}</si>`);
  assert.strictEqual(runs.plainText, '漢字');
  assert.strictEqual(runs.isRich, false);
});

test("a rich string's runs carry the base text and its plain text gains nothing", () => {
  const runs = read('si', `<si><r><rPr><b/></rPr><t>漢字</t></r>${PHONETIC}</si>`);
  assert.deepStrictEqual(runs.runs, [{text: '漢字', font: {bold: true}}]);
  assert.strictEqual(runs.plainText, '');
});

test('with runs off, the text after a phonetic run still reads and the run itself does not', () => {
  const runs = read(
    'is',
    '<is><r><t>漢</t></r><rPh sb="0" eb="1"><t>かん</t></rPh><r><t>字</t></r></is>',
    false,
  );
  assert.strictEqual(runs.plainText, '漢字');
});

test('a self-closing rPh opens no phonetic span, so the text after it still reads', () => {
  assert.strictEqual(read('si', '<si><rPh/><t>base</t></si>').plainText, 'base');
});

test("a note's body is a container like any other, and leaves its phonetic run out", () => {
  const runs = read('text', `<text><r><t>note</t></r>${PHONETIC}</text>`, false);
  assert.strictEqual(runs.plainText, 'note');
});

// A note's body reads no runs, because a note is plain text by model, so its runs flatten to their
// concatenated text.
test('with runs off, every t falls through to the container plain text', () => {
  const runs = read('is', '<is><r><rPr><b/></rPr><t>bo</t></r><r><t>ld</t></r></is>', false);
  assert.deepStrictEqual(runs.runs, []);
  assert.strictEqual(runs.isRich, false);
  assert.strictEqual(runs.plainText, 'bold');
});
