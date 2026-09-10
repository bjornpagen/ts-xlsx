// Tests for the compound-file encoder itself, addressed through its own entry point rather than
// through a VBA project that happens to contain one. What a host does with the container is a
// property of bytes at fixed offsets: the FAT count in the header, a directory entry's size field,
// the order of the sibling tree. Reading those back is the only way to assert them, since a
// round-trip through this library's own reader would agree with the writer about a shared mistake.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strToU8} from 'fflate';

import {type CfbNode, writeCompoundFile} from './cfb-writer.ts';
import {CompoundFile} from './cfb.ts';
import {VbaAuthorError} from './errors.ts';
import {parseVbaProject} from './project.ts';
import {
  buildDirStream,
  buildModuleStream,
  CODE_PAGE,
  MODULES,
  PROJECT_STREAM,
  storeCompress,
} from './vba.test-support.ts';

const SECTOR = 512;
const NOSTREAM = 0xffffffff;
const TYPE_STORAGE = 1;
const TYPE_STREAM = 2;

interface RawDirEntry {
  readonly name: string;
  readonly type: number;
  readonly left: number;
  readonly right: number;
  readonly child: number;
  readonly size: number;
}

/** A view over the container's raw header and directory, read at the offsets [MS-CFB] fixes. */
function inspect(bin: Uint8Array): {
  fatSectors: number;
  totalSectors: number;
  entry: (index: number) => RawDirEntry;
} {
  const dv = new DataView(bin.buffer, bin.byteOffset, bin.byteLength);
  const dirStart = dv.getUint32(48, true);
  const at = (sector: number): number => (sector + 1) * SECTOR;
  return {
    fatSectors: dv.getUint32(44, true),
    // The header occupies the first sector and is not counted among them.
    totalSectors: bin.length / SECTOR - 1,
    entry: (index) => {
      const off = at(dirStart) + index * 128;
      const nameChars = dv.getUint16(off + 64, true) / 2 - 1;
      let name = '';
      for (let k = 0; k < nameChars; k++)
        name += String.fromCharCode(dv.getUint16(off + k * 2, true));
      return {
        name,
        type: bin[off + 66] as number,
        left: dv.getUint32(off + 68, true),
        right: dv.getUint32(off + 72, true),
        child: dv.getUint32(off + 76, true),
        size: dv.getUint32(off + 120, true),
      };
    },
  };
}

// Navigate the directory as a host does, from the Root Entry's child down each storage's balanced
// tree, collecting stream paths in traversal order. Independent of CompoundFile, whose own tree walk
// would agree with this writer about a mistake the two shared; this asserts the tree Excel actually
// walks is a valid, acyclic search tree that reaches every entry.
function treeReachableStreams(bin: Uint8Array): string[] {
  const {entry} = inspect(bin);
  const out: string[] = [];
  const seen = new Set<number>();
  const walk = (idx: number, prefix: string): void => {
    if (idx === NOSTREAM) return;
    if (seen.has(idx)) throw new Error('cycle in directory tree');
    seen.add(idx);
    const e = entry(idx);
    walk(e.left, prefix);
    if (e.type === TYPE_STREAM) out.push(`${prefix}/${e.name}`);
    if (e.type === TYPE_STORAGE) walk(e.child, `${prefix}/${e.name}`);
    walk(e.right, prefix);
  };
  walk(entry(0).child, '');
  return out;
}

const bytes = (length: number, seed: number): Uint8Array =>
  new Uint8Array(length).map((_, i) => (i * seed + 3) & 0xff);

test('writeCompoundFile round-trips mixed small, large, and empty streams through the reader', () => {
  const small = bytes(100, 1);
  const large = bytes(9000, 7); // > 4096 cutoff → regular FAT
  const empty = new Uint8Array(0);
  const bin = writeCompoundFile([
    {name: 'small', data: small},
    {name: 'big', data: large},
    {name: 'empty', data: empty},
  ]);
  const cfb = new CompoundFile(bin);
  assert.deepEqual(cfb.readStream(['small']), small);
  assert.deepEqual(
    cfb.readStream(['big']),
    large,
    'a stream past the mini cutoff round-trips via the regular FAT',
  );
  assert.deepEqual(cfb.readStream(['empty']), empty);
});

test('writeCompoundFile nests streams inside a storage and keeps the tree navigable', () => {
  const bin = writeCompoundFile([
    {name: 'PROJECT', data: strToU8('ID="x"')},
    {
      name: 'VBA',
      children: [
        {name: 'dir', data: Uint8Array.from([1, 2, 3])},
        {name: 'Module1', data: Uint8Array.from([4, 5, 6])},
      ],
    },
  ]);
  const cfb = new CompoundFile(bin);
  assert.deepEqual(cfb.readStream(['VBA', 'dir']), Uint8Array.from([1, 2, 3]));
  assert.deepEqual(cfb.readStream(['VBA', 'Module1']), Uint8Array.from([4, 5, 6]));

  assert.deepEqual(
    treeReachableStreams(bin).sort(),
    ['/PROJECT', '/VBA/Module1', '/VBA/dir'],
    'modules resolve under the VBA storage by tree navigation, not only by linear scan',
  );
});

test('the mini cutoff is the boundary the spec puts it at: 4095 mini, 4096 whole sectors', () => {
  const justUnder = bytes(4095, 5);
  const exactly = bytes(4096, 11);
  const bin = writeCompoundFile([
    {name: 'under', data: justUnder},
    {name: 'atCutoff', data: exactly},
  ]);
  const {entry} = inspect(bin);
  const root = entry(0);
  assert.equal(
    root.size,
    4096,
    'the mini stream holds the sub-cutoff stream alone, padded to whole 64-byte mini sectors',
  );

  const cfb = new CompoundFile(bin);
  assert.deepEqual(cfb.readStream(['under']), justUnder);
  assert.deepEqual(
    cfb.readStream(['atCutoff']),
    exactly,
    'a stream at the cutoff reads back whole',
  );
});

test("a storage's children are linked in the [MS-CFB] name order, not lexicographic order", () => {
  // Two ways this order differs from a plain sort: shorter names sort first whatever their letters
  // ('Z' before 'aa'), and ties compare uppercased ('a' before 'B', which ASCII would reverse).
  const bin = writeCompoundFile([
    {name: 'aa', data: bytes(8, 1)},
    {name: 'Z', data: bytes(8, 2)},
    {name: 'B', data: bytes(8, 3)},
    {name: 'a', data: bytes(8, 4)},
  ]);
  assert.deepEqual(
    treeReachableStreams(bin),
    ['/a', '/B', '/Z', '/aa'],
    'an in-order walk of the sibling tree yields the search order a host binary-searches on',
  );
});

test('a container with no children at all is still a valid, readable compound file', () => {
  const bin = writeCompoundFile([]);
  const {entry} = inspect(bin);
  assert.equal(entry(0).child, NOSTREAM, 'the Root Entry links no children');
  assert.deepEqual(new CompoundFile(bin).tree(), []);
});

test('the FAT sizes itself against a total that includes the FAT, at every crossing', () => {
  // The fixpoint loop is the part worth pinning: each FAT sector it adds enlarges the container, which
  // can demand another. Sweeping stream sizes across the 128-entries-per-sector boundary walks the
  // loop through the sizes where one more sector changes the answer.
  for (let sectors = 120; sectors <= 136; sectors++) {
    const data = bytes(sectors * SECTOR, 3);
    const bin = writeCompoundFile([{name: 'big', data}]);
    const {fatSectors, totalSectors} = inspect(bin);
    assert.equal(
      fatSectors,
      Math.ceil(totalSectors / 128),
      `a ${sectors}-sector stream must leave the FAT sized for the whole container, itself included`,
    );
    assert.deepEqual(new CompoundFile(bin).readStream(['big']), data, `round-trip at ${sectors}`);
  }
});

test('a project past the 109 header DIFAT slots is refused rather than silently truncated', () => {
  // 109 FAT sectors address 109 × 128 sectors, so ~7 MB is the wall. The writer emits no DIFAT sectors,
  // and a container whose FAT pointers do not all fit in the header would be unreadable if written.
  const tooBig = new Uint8Array(109 * 128 * SECTOR);
  assert.throws(
    () => writeCompoundFile([{name: 'huge', data: tooBig}]),
    (error: unknown) =>
      error instanceof VbaAuthorError && /109-sector single-header bound/.test(error.message),
  );
});

test('writeCompoundFile is deterministic: the same hierarchy encodes to the same bytes', () => {
  const build = (): CfbNode[] => [
    {name: 'PROJECT', data: strToU8('ID="x"')},
    {
      name: 'VBA',
      children: [
        {name: 'dir', data: bytes(200, 9)},
        {name: 'M', data: bytes(5000, 4)},
      ],
    },
  ];
  assert.deepEqual(writeCompoundFile(build()), writeCompoundFile(build()));
});

test('writeCompoundFile refuses the names a directory entry cannot carry, fail-closed', () => {
  assert.throws(() => writeCompoundFile([{name: '', data: new Uint8Array(1)}]), VbaAuthorError);
  assert.throws(
    () => writeCompoundFile([{name: 'x'.repeat(32), data: new Uint8Array(1)}]),
    VbaAuthorError,
  );
  assert.throws(
    () =>
      writeCompoundFile([
        {name: 'dup', data: new Uint8Array(1)},
        {name: 'dup', data: new Uint8Array(2)},
      ]),
    VbaAuthorError,
  );
  assert.throws(
    () =>
      writeCompoundFile([
        {
          name: 'VBA',
          children: [
            {name: 'dup', data: new Uint8Array(1)},
            {name: 'dup', data: new Uint8Array(2)},
          ],
        },
      ]),
    VbaAuthorError,
    'a collision nested inside a storage is caught too: the check recurses',
  );
});

test('a name that collides only across storages is legal, since uniqueness is per parent', () => {
  const bin = writeCompoundFile([
    {name: 'A', children: [{name: 'shared', data: bytes(16, 1)}]},
    {name: 'B', children: [{name: 'shared', data: bytes(16, 2)}]},
  ]);
  assert.deepEqual(treeReachableStreams(bin).sort(), ['/A/shared', '/B/shared']);
});

// ── The join: a project packaged through this writer ─────────────────────────────────────────────────
// The encoder's own properties are asserted above, against the bytes. What belongs here too is the
// join: that a project packaged through the production writer is one the production parser reads back.

test('writeCompoundFile produces a container parseVbaProject decodes', () => {
  // Build the VBA-project stream set from the same fixture bytes, but package it through the production
  // writer (proper VBA-storage hierarchy) rather than the fixture's own container, proving the writer
  // yields a parseable project, not merely a reader-round-trippable blob.
  const dir = storeCompress(Uint8Array.from(buildDirStream(CODE_PAGE, MODULES)));
  const vbaChildren: CfbNode[] = [
    {name: 'dir', data: dir},
    ...MODULES.map((m) => ({name: m.name, data: buildModuleStream(m)})),
  ];
  const bin = writeCompoundFile([
    {name: 'PROJECT', data: strToU8(PROJECT_STREAM)},
    {name: 'VBA', children: vbaChildren},
  ]);

  const project = parseVbaProject(bin);
  assert.deepEqual(
    project.modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
  );
  assert.equal(project.codePage, 1251);
});
