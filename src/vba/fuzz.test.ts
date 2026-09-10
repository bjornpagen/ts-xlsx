import assert from 'node:assert/strict';
import {test} from 'node:test';

import {xorshift32} from '../fuzz.test-support.ts';
import {type CfbNode, writeCompoundFile} from './cfb-writer.ts';
import {CompoundFile} from './cfb.ts';
import {VbaAuthorError, VbaParseError} from './errors.ts';
import {compressContainer} from './ms-ovba.ts';
import {addVbaReference, removeVbaModule} from './project-editor.ts';
import {parseVbaProject} from './project.ts';
import {ascii, rec, u16le, u32le} from './vba.test-support.ts';

// An adversarial pass over the CFB container reader and the MS-OVBA decompressor.
//
// The same contract `io/xlsb/fuzz.test.ts` holds the BIFF12 reader to, for the same reason: a mutated
// container has no correct reading, so what is under test is that whatever the bytes say, the parser
// either produces a project or fails closed with a typed error, in bounded time and bounded memory.
//
// The three shapes each hide a bug inside a crash:
//
//   - a `RangeError` or `TypeError` escaping means a count out of the header became an index, a length
//     or a recursion depth without being checked first;
//   - a run that does not finish means a count out of the header became a loop bound;
//   - a run that finishes but allocates without limit means a count out of the header became an
//     allocation size, and the only thing standing between it and an OOM was the size of the input.
//
// This file is a `vbaProject.bin` shaped like the real thing, written by the production writer and
// compressed by the production compressor, so a mutation lands inside a structure real enough to reach
// the FAT, DIFAT, directory tree and decompressor rather than being turned away at the signature.

// The wall-clock ceiling any single parse must stay under. Generous by two orders of magnitude against
// a healthy parse of this seed, so it fires only on a genuine unbounded loop and never on a slow
// machine or a cold JIT.
const PARSE_BUDGET_MS = 4_000;

// ── The seed ────────────────────────────────────────────────────────────────────────────────────────

const MODULE_NAMES = ['ThisWorkbook', 'Module1', 'Class1'] as const;

// The dir stream [MS-OVBA] 2.3.4.2: enough of the record grammar for `parseVbaProject` to discover
// every module and its stream, which is what puts the deeper parsers in a mutation's reach.
function dirStream(): Uint8Array {
  const records: number[] = [];
  records.push(...rec(0x0003, u16le(1252))); // PROJECTCODEPAGE
  records.push(...rec(0x0009, u32le(4)), ...u16le(10)); // PROJECTVERSION, uncounted minor
  records.push(...rec(0x000f, u16le(MODULE_NAMES.length))); // MODULES_COUNT
  records.push(...rec(0x0013, u16le(0xffff))); // PROJECTCOOKIE
  for (const [index, name] of MODULE_NAMES.entries()) {
    records.push(...rec(0x0019, ascii(name))); // MODULENAME
    records.push(...rec(0x001a, ascii(name))); // MODULESTREAMNAME
    records.push(...rec(0x0031, u32le(16))); // MODULEOFFSET
    records.push(...rec(index === 0 ? 0x0022 : 0x0021, [])); // MODULETYPE
    records.push(...rec(0x002b, [])); // MODULETERMINATOR
  }
  records.push(...rec(0x0010, [])); // dir terminator
  return compressContainer(Uint8Array.from(records));
}

// A module's CFB stream: a zeroed stand-in for the p-code prefix, then the compressed source.
function moduleStream(source: string): Uint8Array {
  const compressed = compressContainer(Uint8Array.from(ascii(source)));
  const out = new Uint8Array(16 + compressed.length);
  out.set(compressed, 16);
  return out;
}

const PROJECT_STREAM = ascii(
  [
    'ID="{00000000-0000-0000-0000-000000000000}"',
    'Document=ThisWorkbook/&H00000000',
    'Module=Module1',
    'Class=Class1',
    '',
  ].join('\r\n'),
);

function seed(): Uint8Array {
  const children: CfbNode[] = [
    {name: 'dir', data: dirStream()},
    {name: '_VBA_PROJECT', data: Uint8Array.of(0x61, 0xcc, 0x5e, 0x00, 0x00, 0x01, 0x02, 0x03)},
    ...MODULE_NAMES.map((name) => ({name, data: moduleStream(`Sub ${name}_Demo()\r\nEnd Sub`)})),
  ];
  return writeCompoundFile([
    {name: 'PROJECT', data: Uint8Array.from(PROJECT_STREAM)},
    {name: 'VBA', children},
  ]);
}

const SEED = seed();

// ── The harness ─────────────────────────────────────────────────────────────────────────────────────

/**
 * Drive every entry point the container reaches through, asserting only that each fails the way a
 * parser is allowed to fail, and within the budget.
 *
 * `tree()` is called explicitly because it is the only reader path `parseVbaProject` does not touch,
 * and it is the one that recurses; the editors are called because they are the public surface that
 * reaches it (`Workbook.removeVbaModule`, `editXlsxVbaRemoveModule`).
 */
function parsesOrFailsClosed(bin: Uint8Array, label: string): void {
  const started = performance.now();
  for (const drive of [
    () => parseVbaProject(bin),
    () => new CompoundFile(bin).tree(),
    () => removeVbaModule(bin, 'Module1'),
    () =>
      addVbaReference(bin, {
        name: 'Scripting',
        guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
        majorVersion: 1,
        minorVersion: 0,
        path: 'C:\\Windows\\System32\\scrrun.dll',
      }),
  ]) {
    try {
      drive();
    } catch (error) {
      if (error instanceof VbaParseError || error instanceof VbaAuthorError) continue;
      // The stack, not just the message: a failure here names a line in a parser, and reproducing it
      // from the label alone would mean re-running the whole pass to find out where it landed.
      assert.ok(
        error instanceof Error && !(error instanceof TypeError) && !(error instanceof RangeError),
        `${label}: expected a typed, closed failure but got ${error instanceof Error ? error.stack : String(error)}`,
      );
    }
  }
  const elapsed = performance.now() - started;
  assert.ok(elapsed < PARSE_BUDGET_MS, `${label}: took ${elapsed.toFixed(0)}ms, over the budget`);
}

// ── The passes ──────────────────────────────────────────────────────────────────────────────────────

test('the seed itself parses, so a mutation of it is a mutation of something real', () => {
  const project = parseVbaProject(SEED);
  assert.deepEqual(
    project.modules.map((module) => module.name),
    [...MODULE_NAMES],
  );
  assert.equal(new CompoundFile(SEED).tree().length, 2);
});

test('single-byte mutations anywhere in the container never escape the typed failure modes', () => {
  const next = xorshift32(0xc0ffee);
  for (let round = 0; round < 400; round++) {
    const mutated = Uint8Array.from(SEED);
    const offset = next() % mutated.length;
    mutated[offset] = next() & 0xff;
    parsesOrFailsClosed(mutated, `byte@${offset} round ${round}`);
  }
});

test('a container truncated at any point fails closed rather than reading half a project into a crash', () => {
  const next = xorshift32(0x7ac6);
  for (let round = 0; round < 120; round++) {
    parsesOrFailsClosed(SEED.subarray(0, next() % SEED.length), `cut round ${round}`);
  }
});

test('every header count driven to its maximum is rejected, not believed', () => {
  // The header's counts are the file's most dangerous numbers: each is a `u32` that becomes a loop
  // bound, an allocation size or a sector index. Every one is struck at both its real offset and at
  // neighbouring alignments, so a field read at the wrong place is covered too.
  for (const offset of [30, 32, 44, 48, 56, 60, 68, 72]) {
    for (const value of [0xffffffff, 0x7fffffff, 0x0000ffff, 0xfffffffe]) {
      const mutated = Uint8Array.from(SEED);
      mutated.set(u32le(value), offset);
      parsesOrFailsClosed(mutated, `header@${offset}=${value.toString(16)}`);
    }
  }
});

// The sweep above asks only that a struck header fails closed. These two ask for the stronger thing on
// the two layout fields the reader used to believe: that it *rejects* them, rather than reading the
// container through the wrong allocator or through half of a declared size. Both destinations are
// bounds-checked, so without these the crafted file yields a module's source read back as bytes nobody
// wrote, which is a worse outcome than a refusal, and a quieter one.
test('a mini-stream cutoff other than the one MS-CFB fixes is rejected, not believed', () => {
  for (const cutoff of [0, 64, 8192, 0xffffffff]) {
    const mutated = Uint8Array.from(SEED);
    mutated.set(u32le(cutoff), 56);
    assert.throws(() => new CompoundFile(mutated), {
      name: 'VbaParseError',
      message: /mini-stream cutoff/,
    });
  }
});

test('a directory entry declaring a stream past 4 GiB is rejected rather than read truncated', () => {
  const mutated = Uint8Array.from(SEED);
  // The root entry is the first; the second is the first real stream, whose size is a u64 at +120.
  mutated.set(u32le(1), findDirectory(mutated) + 128 + 124); // the high half of the declared size
  assert.throws(() => new CompoundFile(mutated), {
    name: 'VbaParseError',
    message: /larger than 4 GiB/,
  });
});

test('a stream whose chain ends before its declared size is rejected, not returned short', () => {
  // The sibling of the 4-GiB refusal above, and the same failure it guards against: a size the chain
  // cannot deliver reads back as a well-formed prefix, so every parser downstream sees a shorter
  // module than the file declares and none of them can tell. The root entry is the first, and its
  // stream is the mini-stream the container reads while opening.
  const mutated = Uint8Array.from(SEED);
  const rootSize = new DataView(mutated.buffer, mutated.byteOffset, mutated.byteLength).getUint32(
    findDirectory(mutated) + 120,
    true,
  );
  mutated.set(u32le(rootSize + 8192), findDirectory(mutated) + 120);
  assert.throws(() => new CompoundFile(mutated), {
    name: 'VbaParseError',
    message: /ends after \d+ bytes, but \d+ were declared/,
  });
});

test('a directory whose sibling links form a chain does not recurse once per entry', () => {
  // A red-black sibling tree is balanced by construction, so its depth is logarithmic and recursion is
  // safe. A hostile file is under no such obligation: linking every entry as the left child of the one
  // before it makes depth equal to entry count. The directory is bounded only by its FAT chain, so
  // "entry count" is "as many as the attacker cares to write".
  const entries = 20_000;
  const nodes: CfbNode[] = Array.from({length: entries}, (_, index) => ({
    name: `S${index}`,
    data: Uint8Array.of(index & 0xff),
  }));
  const container = writeCompoundFile(nodes);

  // Relink: every entry's left sibling is its predecessor, its right sibling absent. The writer emits a
  // balanced tree, so this is a rewrite of the links rather than something it can be asked for.
  const directory = findDirectory(container);
  for (let index = 1; index < entries; index++) {
    const base = directory + index * 128;
    if (base + 128 > container.length) break;
    container.set(u32le(index - 1), base + 68); // left sibling
    container.set(u32le(0xffffffff), base + 72); // right sibling: none
  }

  parsesOrFailsClosed(container, 'left-linked sibling chain');
});

test('a FAT assembled from repeated sector ids does not amplify the input into heap', () => {
  // The FAT is assembled by pushing one entry per `u32` of every sector a DIFAT id names, and ids are
  // free to repeat, so the cost is the header's *claim* rather than the bytes behind it. Before the
  // bound, this container cost 272 MB of heap and 1.3 s: linear in input size, so a 10 MB project was
  // ~2.7 GB and a fatal OOM instead of a typed failure.
  const container = amplifyingContainer();
  const before = process.memoryUsage().heapUsed;
  assert.throws(() => new CompoundFile(container), VbaParseError);
  const grew = process.memoryUsage().heapUsed - before;
  // Two orders of magnitude under what the unbounded version cost, and far above anything a bounded
  // parse of a 1 MB file can legitimately need, so this measures the bound rather than the allocator.
  assert.ok(
    grew < 16_000_000,
    `parsing ${container.length} bytes grew the heap by ${(grew / 1e6).toFixed(1)} MB`,
  );
});

/**
 * A container whose header claims far more FAT and DIFAT sectors than the file has sectors to hold,
 * with every DIFAT slot naming the same sector so the ids repeat without limit.
 */
function amplifyingContainer(): Uint8Array {
  const SECTOR = 512;
  const SECTORS = 2000;
  const buf = new Uint8Array(SECTOR + SECTORS * SECTOR);
  const view = new DataView(buf.buffer);
  view.setUint32(0, 0xe011cfd0, true);
  view.setUint32(4, 0xe11ab1a1, true);
  view.setUint16(30, 9, true); // 512-byte sectors
  view.setUint16(32, 6, true);
  view.setUint32(44, 0x00ffffff, true); // FAT sectors: a lie
  view.setUint32(48, 0xfffffffe, true); // directory: end of chain
  view.setUint32(56, 4096, true);
  view.setUint32(60, 0xfffffffe, true);
  view.setUint32(68, 0, true); // first DIFAT sector: real, so the chain is walkable
  view.setUint32(72, 0x00ffffff, true); // DIFAT sectors: a lie
  for (let slot = 0; slot < 109; slot++) view.setUint32(76 + slot * 4, 1, true);
  for (let sector = 0; sector < SECTORS - 1; sector++) {
    const base = SECTOR + sector * SECTOR;
    for (let slot = 0; slot < 127; slot++) view.setUint32(base + slot * 4, 1, true);
    view.setUint32(base + 127 * 4, sector + 1, true); // chain on to the next
  }
  view.setUint32(SECTOR + (SECTORS - 1) * SECTOR + 127 * 4, 0xfffffffe, true);
  return buf;
}

/** The byte offset of the first directory sector, resolved from the header the way the reader does. */
function findDirectory(container: Uint8Array): number {
  const view = new DataView(container.buffer, container.byteOffset, container.byteLength);
  const sectorSize = 1 << view.getUint16(30, true);
  return (view.getUint32(48, true) + 1) * sectorSize;
}
