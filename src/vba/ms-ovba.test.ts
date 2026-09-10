// The MS-OVBA compression container: the decompressor against an independent vector and at its bounds,
// and the compressor through a round trip and through the whole parse pipeline.

import {strict as assert} from 'node:assert';
import {test} from 'node:test';

import {strToU8} from 'fflate';

import {writeCompoundFile} from './cfb-writer.ts';
import {VbaParseError} from './errors.ts';
import {compressContainer, decompressContainer} from './ms-ovba.ts';
import {parseVbaProject} from './project.ts';
import {buildDirStream, buildVbaProjectBin, CODE_PAGE, storeCompress} from './vba.test-support.ts';

// ── Decompressor: independent ground-truth vector ───────────────────────────────────────────────────

test('decompressContainer expands a hand-verified copy-token vector', () => {
  // Container for "abcabc": literals a,b,c then CopyToken(offset 3, length 3). Encoded by hand from
  // [MS-OVBA] 2.4.1.3.19.3, not by this suite's storeCompress, which never emits copy tokens.
  const container = Uint8Array.from([0x01, 0x05, 0xb0, 0x08, 0x61, 0x62, 0x63, 0x00, 0x20]);
  assert.equal(new TextDecoder('latin1').decode(decompressContainer(container)), 'abcabc');
});

test('decompressContainer round-trips the store encoder for multi-chunk data', () => {
  const data = new Uint8Array(5000).map((_, i) => (i * 31 + 7) & 0xff); // > 2048 → several chunks
  assert.deepEqual(decompressContainer(storeCompress(data)), data);
});

test('decompressContainer rejects a bad signature byte', () => {
  assert.throws(() => decompressContainer(Uint8Array.from([0x00, 0x01])), VbaParseError);
});

test('decompressContainer caps output to guard a decompression bomb', () => {
  const data = new Uint8Array(4096);
  assert.throws(() => decompressContainer(storeCompress(data), 0, 1024), VbaParseError);
});

test('decompressContainer admits exactly the ceiling and refuses the byte after it', () => {
  // The ceiling bounds the buffer the decompressor grows, so its boundary is a real allocation
  // limit, not a post-hoc check: a container producing exactly `maxOutput` bytes must still expand.
  const data = new Uint8Array(5000).map((_, i) => (i * 17 + 3) & 0xff);
  assert.deepEqual(decompressContainer(storeCompress(data), 0, 5000), data);
  assert.throws(() => decompressContainer(storeCompress(data), 0, 4999), VbaParseError);
});

test('decompressContainer refuses a chunk that decodes past the 4096-byte window MS-OVBA defines', () => {
  // [MS-OVBA] 2.4.1.3.6 caps a chunk at 4096 decompressed bytes. Seven bytes reach past it: one literal
  // to seed the window, then a single CopyToken whose length field is at its maximum, which run-length
  // expands that one byte 4098 times. Past the window `copyTokenHelp` widens the offset field until the
  // length mask is seven bits, so every later token in the chunk decodes under a bit split no producer
  // emits, which is bytes nobody wrote rather than merely bytes nobody wanted.
  const flags = 0b0000_0010; // the first token is a literal, the second a copy
  const body = [flags, 0x41, 0xff, 0x0f]; // 'A', then CopyToken 0x0fff: offset 1, length 4098
  const header = 0xb000 | (body.length - 1);
  const container = Uint8Array.from([0x01, header & 0xff, (header >> 8) & 0xff, ...body]);

  assert.throws(() => decompressContainer(container), VbaParseError);
  assert.throws(() => decompressContainer(container), /4096/);
});

test('a project budget bounds what every module decompresses between them, not each call alone', () => {
  // `decompressContainer`'s own ceiling is per call, which bounds one bomb and nothing else: a project
  // is one `dir` plus one stream per module, all retained at once. A project whose modules each fit
  // comfortably under the per-call ceiling must still be refused once they exceed the project's.
  const sources = ['A', 'B', 'C'].map((name) => ({
    name,
    documentType: false,
    sourceBytes: new Array(3000).fill(0x41),
    pcodePrefixLen: 0,
  }));
  const bin = buildVbaProjectBin(CODE_PAGE, sources);

  // Comfortably above any one stream, comfortably below their sum plus the dir stream.
  assert.throws(() => parseVbaProject(bin, 4000), VbaParseError);
  assert.ok(parseVbaProject(bin, 64 * 1024).modules.length === 3);
});

// ── Compressor (§2.3b): compressContainer ────────────────────────────────────────────────────────────

test('compressContainer round-trips arbitrary data across the chunk boundary', () => {
  // Sizes straddling the 4096-byte chunk window and the 8-token flag group catch off-by-one framing.
  for (const n of [0, 1, 2, 3, 7, 8, 9, 100, 4095, 4096, 4097, 5000, 12000]) {
    const data = new Uint8Array(n).map((_, i) => (i * 131 + 7) & 0xff);
    assert.deepEqual(
      decompressContainer(compressContainer(data)),
      data,
      `round-trip failed at n=${n}`,
    );
  }
});

test('compressContainer emits copy tokens, shrinking repetitive data via run-length overlap', () => {
  const runs = new Uint8Array(4096).fill(0x41); // one byte repeated → a single overlapping back-reference
  const packed = compressContainer(runs);
  assert.ok(
    packed.length < 32,
    `4096 identical bytes should collapse to a tiny container, got ${packed.length}`,
  );
  assert.deepEqual(decompressContainer(packed), runs);

  const abab = Uint8Array.from({length: 6000}, (_, i) => (i % 2 ? 0x62 : 0x61));
  assert.ok(compressContainer(abab).length < abab.length / 4);
  assert.deepEqual(decompressContainer(compressContainer(abab)), abab);
});

// `removeVbaModule` and `addVbaReference` recompress the `dir` stream of a workbook the caller did
// not write, so the encoder's cost is a number an untrusted file chooses. Data with no matches is its
// worst case, and the one this bounds: the exhaustive back-window rescan this replaced spent 1.2 s on
// 128 KiB of it, so ~10 s on the megabyte below, against 0.14 s for the hash chain. The ceiling is
// wall-clock because the blow-up was a constant factor inside a fixed 4096-byte window rather than a
// growth rate, so no ratio between two input sizes can see it; the margin is what makes it sound.
test('the compressor stays fast on data that offers it no matches', () => {
  let state = 0x9e37_79b9;
  const noMatches = new Uint8Array(1024 * 1024).map(() => {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    return state & 0xff;
  });

  const started = performance.now();
  const packed = compressContainer(noMatches);
  const elapsed = performance.now() - started;

  assert.deepEqual(decompressContainer(packed), noMatches);
  assert.ok(elapsed < 3000, `compressing 1 MiB of unmatched data took ${elapsed.toFixed(0)} ms`);
});

test('compressContainer output re-parses as a real module through the whole pipeline', () => {
  // Compress genuine VBA source, wrap it as a module stream at offset 0, and read it back through the
  // production CFB writer + parser: the compressor feeding the reader end to end, no store-mode fixture.
  const source = 'Sub Demo()\r\n    MsgBox "hi"\r\n    MsgBox "hi"\r\nEnd Sub';
  const compressed = compressContainer(strToU8(source));
  const dir = compressContainer(
    Uint8Array.from(
      buildDirStream(1252, [
        {name: 'Demo', documentType: false, sourceBytes: [], pcodePrefixLen: 0},
      ]),
    ),
  );
  const bin = writeCompoundFile([
    {name: 'PROJECT', data: strToU8('Module=Demo\r\n')},
    {
      name: 'VBA',
      children: [
        {name: 'dir', data: dir},
        {name: 'Demo', data: compressed},
      ],
    },
  ]);
  const project = parseVbaProject(bin);
  assert.equal(
    project.modules[0]!.source,
    source,
    'the module source survives compress → write → parse',
  );
});
