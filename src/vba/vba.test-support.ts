// Builders for the VBA tests: the bytes of a genuine, spec-valid `vbaProject.bin` constructed from
// scratch, and the macro-enabled packages around one.
//
// An MS-OVBA "store" encoder (literal-only chunks: valid compression that happens not to compress) and a
// minimal MS-CFB writer. No third-party bytes, no Excel. This exercises the real parse pipeline; the
// decoder is additionally pinned against an independent hand-verified vector (the first test in
// `ms-ovba.test.ts`) so the fixture and the code under test are not a closed loop on the copy-token path.

import {strict as assert} from 'node:assert';

import {strToU8, zipSync} from 'fflate';

import {type CfbNode, writeCompoundFile} from './cfb-writer.ts';

const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const FATSECT = 0xfffffffd;

/** Encode bytes as an MS-OVBA CompressedContainer using literal-only tokens (no back-references). */
export function storeCompress(data: Uint8Array): Uint8Array {
  const out: number[] = [0x01];
  const CHUNK = 2048; // keep encoded chunk-data ≤ 4096 (the 12-bit size field) after 1/8 flag overhead
  for (let c = 0; c < data.length; c += CHUNK) {
    const slice = data.subarray(c, Math.min(c + CHUNK, data.length));
    const body: number[] = [];
    for (let i = 0; i < slice.length; i += 8) {
      body.push(0x00); // one flag byte, all-literal for the next up-to-8 bytes
      for (let j = i; j < Math.min(i + 8, slice.length); j++) body.push(slice[j] as number);
    }
    const header = 0xb000 | ((body.length - 1) & 0x0fff); // compressed + 0b011 sig + (size-1)
    out.push(header & 0xff, (header >> 8) & 0xff, ...body);
  }
  return Uint8Array.from(out);
}

export function u16le(n: number): number[] {
  return [n & 0xff, (n >> 8) & 0xff];
}

export function u32le(n: number): number[] {
  return [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >> 24) & 0xff];
}

/** One TLV record: `Id(u16) Size(u32) data[Size]`. */
export function rec(id: number, data: readonly number[]): number[] {
  return [...u16le(id), ...u32le(data.length), ...data];
}

// Code units, not code points. `[...s]` iterates code points, so a surrogate pair would arrive as a
// single character and `charCodeAt(0)` would keep only its high half. A CFB stream name is UTF-16
// code units on the wire, so the encoders below must count the same way the format does.
export function ascii(s: string): number[] {
  return Array.from({length: s.length}, (_, i) => s.charCodeAt(i));
}

function utf16le(s: string): number[] {
  return ascii(s).flatMap(u16le);
}

interface ModuleSpec {
  name: string;
  documentType: boolean; // MODULETYPE: false→procedural (0x21), true→non-procedural (0x22)
  sourceBytes: number[];
  pcodePrefixLen: number;
}

export function buildDirStream(codePage: number, modules: ModuleSpec[]): number[] {
  const records: number[] = [];
  records.push(...rec(0x0003, u16le(codePage))); // PROJECTCODEPAGE
  // PROJECTVERSION: Size=4 counts only VersionMajor; the trailing 2-byte VersionMinor is uncounted,
  // the exact record that misaligns a naive TLV walk. Its presence proves the parser skips it.
  records.push(...rec(0x0009, u32le(0x04)), ...u16le(0x000a));
  records.push(...rec(0x000f, u16le(modules.length))); // MODULES_COUNT
  records.push(...rec(0x0013, u16le(0xffff))); // PROJECTCOOKIE
  for (const m of modules) {
    records.push(...rec(0x0019, ascii(m.name))); // MODULENAME
    records.push(...rec(0x001a, ascii(m.name))); // MODULESTREAMNAME (MBCS)
    records.push(...rec(0x0032, utf16le(m.name))); // Reserved: Unicode stream name, must be skipped
    records.push(...rec(0x0031, u32le(m.pcodePrefixLen))); // MODULEOFFSET
    records.push(...rec(m.documentType ? 0x0022 : 0x0021, [])); // MODULETYPE (Reserved u32 = Size 0)
    records.push(...rec(0x002b, [])); // MODULETERMINATOR
  }
  records.push(...rec(0x0010, [])); // dir Terminator: closes PROJECTMODULES, ends the dir stream
  return records;
}

/** Bytes stored in a module's CFB stream: a fake p-code prefix, then the compressed source. */
export function buildModuleStream(m: ModuleSpec): Uint8Array {
  const compressed = storeCompress(Uint8Array.from(m.sourceBytes));
  const out = new Uint8Array(m.pcodePrefixLen + compressed.length);
  out.set(compressed, m.pcodePrefixLen); // prefix left as zeros: stand-in for the PerformanceCache
  return out;
}

/** One stream of a fixture container. */
interface FixtureStream {
  readonly name: string;
  readonly data: Uint8Array;
}

// A directory entry under construction: its index on disk, and the sibling-tree links a reader walks.
interface FixtureEntry {
  readonly index: number;
  readonly name: string;
  readonly type: number;
  start: number;
  size: number;
  right: number;
  child: number;
}

/**
 * A minimal MS-CFB container holding `rootStreams` at the root and `vbaStreams` in a `VBA` storage, the
 * layout [MS-OVBA] 2.2 fixes. Every stream here is < 4096 bytes, so all live in the mini stream. Each
 * storage links its children as a chain of right siblings: a tree a reader can walk, though not the
 * balanced, name-ordered one a host binary-searches, which `cfb-writer.ts` builds and this deliberately
 * does not share.
 */
function buildCfb(
  rootStreams: readonly FixtureStream[],
  vbaStreams: readonly FixtureStream[],
): Uint8Array {
  const SEC = 512;
  const MINI = 64;

  const entries: FixtureEntry[] = [];
  const add = (name: string, type: number, start = 0, size = 0): FixtureEntry => {
    const entry = {
      index: entries.length,
      name,
      type,
      start,
      size,
      right: FREESECT,
      child: FREESECT,
    };
    entries.push(entry);
    return entry;
  };
  // The first child hangs off the storage, and each after it off the right of the one before.
  const link = (storage: FixtureEntry, children: readonly FixtureEntry[]): void => {
    let previous: FixtureEntry | undefined;
    for (const child of children) {
      if (previous === undefined) storage.child = child.index;
      else previous.right = child.index;
      previous = child;
    }
  };

  // Pack each stream into whole mini-sectors and chain them in the mini-FAT.
  const miniBytes: number[] = [];
  const miniFat: number[] = [];
  const addStream = (s: FixtureStream): FixtureEntry => {
    assert.ok(s.data.length < 4096, 'fixture streams must be mini-stream sized');
    const startMini = miniBytes.length / MINI;
    const numMini = Math.max(1, Math.ceil(s.data.length / MINI));
    for (let k = 0; k < numMini; k++)
      miniFat.push(k < numMini - 1 ? startMini + k + 1 : ENDOFCHAIN);
    miniBytes.push(...s.data, ...new Array<number>(numMini * MINI - s.data.length).fill(0));
    return add(s.name, 2, startMini, s.data.length);
  };
  const root = add('Root Entry', 5);
  const vba = add('VBA', 1);
  link(root, [...rootStreams.map(addStream), vba]);
  link(vba, vbaStreams.map(addStream));

  const dirSectors = Math.max(1, Math.ceil((entries.length * 128) / SEC));
  const miniFatSectors = miniFat.length > 0 ? Math.ceil((miniFat.length * 4) / SEC) : 0;
  const miniStreamSectors = Math.ceil(miniBytes.length / SEC);

  let next = 0;
  const fatSectorIdx = next++;
  const dirStart = next;
  next += dirSectors;
  const miniFatStart = miniFatSectors > 0 ? next : ENDOFCHAIN;
  next += miniFatSectors;
  const miniStreamStart = next;
  next += miniStreamSectors;
  const totalSectors = next;
  assert.ok(totalSectors + 1 <= 128, 'fixture must fit a single FAT sector');

  const fat = new Array<number>(128).fill(FREESECT);
  fat[fatSectorIdx] = FATSECT;
  const chain = (from: number, count: number): void => {
    for (let k = 0; k < count; k++) fat[from + k] = k < count - 1 ? from + k + 1 : ENDOFCHAIN;
  };
  chain(dirStart, dirSectors);
  if (miniFatSectors > 0) chain(miniFatStart, miniFatSectors);
  chain(miniStreamStart, miniStreamSectors);

  root.start = miniStreamStart;
  root.size = miniBytes.length;

  const buf = new Uint8Array((totalSectors + 1) * SEC);
  const dv = new DataView(buf.buffer);
  const sectorOffset = (idx: number): number => (idx + 1) * SEC;

  // Header
  dv.setUint32(0, 0xe011cfd0, true);
  dv.setUint32(4, 0xe11ab1a1, true);
  dv.setUint16(24, 0x003e, true); // minor version
  dv.setUint16(26, 0x0003, true); // major version (v3)
  dv.setUint16(28, 0xfffe, true); // byte order
  dv.setUint16(30, 9, true); // sector shift → 512
  dv.setUint16(32, 6, true); // mini sector shift → 64
  dv.setUint32(44, 1, true); // number of FAT sectors
  dv.setUint32(48, dirStart, true); // first directory sector
  dv.setUint32(56, 4096, true); // mini-stream cutoff
  dv.setUint32(60, miniFatStart, true); // first mini-FAT sector
  dv.setUint32(64, miniFatSectors, true); // number of mini-FAT sectors
  dv.setUint32(68, ENDOFCHAIN, true); // first DIFAT sector
  dv.setUint32(72, 0, true); // number of DIFAT sectors
  for (let i = 0; i < 109; i++) dv.setUint32(76 + i * 4, i === 0 ? fatSectorIdx : FREESECT, true);

  // FAT sector
  for (let i = 0; i < 128; i++)
    dv.setUint32(sectorOffset(fatSectorIdx) + i * 4, fat[i] as number, true);

  // Directory
  for (const e of entries) {
    const off = sectorOffset(dirStart) + e.index * 128;
    const name16 = utf16le(e.name);
    name16.forEach((b, j) => {
      buf[off + j] = b;
    });
    dv.setUint16(off + 64, name16.length + 2, true); // name length incl. null terminator
    buf[off + 66] = e.type;
    dv.setUint32(off + 68, FREESECT, true); // left sibling: every chain here runs right
    dv.setUint32(off + 72, e.right, true); // right sibling
    dv.setUint32(off + 76, e.child, true); // child
    dv.setUint32(off + 116, e.start, true);
    dv.setUint32(off + 120, e.size, true);
  }

  // Mini-FAT
  if (miniFatSectors > 0) {
    for (let i = 0; i < miniFatSectors * (SEC / 4); i++) {
      dv.setUint32(sectorOffset(miniFatStart) + i * 4, miniFat[i] ?? FREESECT, true);
    }
  }

  // Mini stream
  miniBytes.forEach((b, i) => {
    buf[sectorOffset(miniStreamStart) + i] = b;
  });

  return buf;
}

export const PROJECT_STREAM = [
  'ID="{00000000-0000-0000-0000-000000000000}"',
  'Document=ThisWorkbook/&H00000000',
  'Module=Module1',
  'Class=Class1',
  '',
].join('\r\n');

export function buildVbaProjectBin(codePage: number, modules: ModuleSpec[]): Uint8Array {
  const dir = storeCompress(Uint8Array.from(buildDirStream(codePage, modules)));
  return buildCfb(
    [{name: 'PROJECT', data: strToU8(PROJECT_STREAM)}],
    [{name: 'dir', data: dir}, ...modules.map((m) => ({name: m.name, data: buildModuleStream(m)}))],
  );
}

// A three-module project: a document code-behind, a procedural .bas, and a class module. The last two
// share MODULETYPE 0x22/0x21 but are told apart by the PROJECT stream. Module1's source carries byte
// 0xC0, which is 'А' (U+0410) in code page 1251, proving code-page-aware decoding, not latin1.
export const MODULES: ModuleSpec[] = [
  {
    name: 'ThisWorkbook',
    documentType: true,
    sourceBytes: ascii('Private Sub Workbook_Open()\r\nEnd Sub'),
    pcodePrefixLen: 16,
  },
  {
    name: 'Module1',
    documentType: false,
    sourceBytes: [...ascii('Sub Test() '), 0xc0],
    pcodePrefixLen: 24,
  },
  {name: 'Class1', documentType: true, sourceBytes: ascii('Public X As Long'), pcodePrefixLen: 8},
];
export const CODE_PAGE = 1251;

// One VBA-signature generation to wire into a package: the full relationship Type (carrying the year
// segment the generation really uses: 2006 legacy, 2014 agile, 2020 V3), the sibling part's file name
// and content type, and its raw bytes. Office attaches up to three of these over one project.
interface SignatureGeneration {
  readonly relType: string;
  readonly fileName: string;
  readonly contentType: string;
  readonly bytes: Uint8Array;
}

export const legacySignature = (bytes: Uint8Array): SignatureGeneration => ({
  relType: 'http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature',
  fileName: 'vbaProjectSignature.bin',
  contentType: 'application/vnd.ms-office.vbaProjectSignature',
  bytes,
});
export const agileSignature = (bytes: Uint8Array): SignatureGeneration => ({
  relType: 'http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile',
  fileName: 'vbaProjectSignatureAgile.bin',
  contentType: 'application/vnd.ms-office.vbaProjectSignatureAgile',
  bytes,
});
export const v3Signature = (bytes: Uint8Array): SignatureGeneration => ({
  relType: 'http://schemas.microsoft.com/office/2020/relationships/vbaProjectSignatureV3',
  fileName: 'vbaProjectSignatureV3.bin',
  contentType: 'application/vnd.ms-office.vbaProjectSignatureV3',
  bytes,
});

/**
 * A minimal macro-enabled package around a vbaProject.bin, or the same package with no project at all
 * when `vbaBin` is `undefined`.
 *
 * `signatures` additionally wires digital signatures over the project: sibling parts reached from the
 * project part's own rels, the shape an authoring replace must invalidate (a signature over old bytes
 * cannot vouch for new ones). `documentDir` moves the office document out of `xl/`, the shape a producer
 * other than Excel may emit: OPC fixes nothing but `_rels/.rels`, so `xl/` is a convention rather than a
 * rule and every target below it resolves relative to its referrer.
 */
export function xlsmPackage(
  vbaBin: Uint8Array | undefined,
  options: {
    readonly signatures?: readonly SignatureGeneration[];
    readonly documentDir?: string;
  } = {},
): Uint8Array {
  const {signatures = [], documentDir: dir = 'xl'} = options;
  assert.ok(vbaBin !== undefined || signatures.length === 0, 'a signature needs a project to sign');
  const ms = 'http://schemas.microsoft.com/office/2006/relationships';
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        (vbaBin === undefined
          ? ''
          : `<Override PartName="/${dir}/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>`) +
        signatures
          .map((s) => `<Override PartName="/${dir}/${s.fileName}" ContentType="${s.contentType}"/>`)
          .join('') +
        `<Override PartName="/${dir}/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
        `<Override PartName="/${dir}/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
        '</Types>',
    ),
    '_rels/.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="${dir}/workbook.xml"/>` +
        '</Relationships>',
    ),
    [`${dir}/workbook.xml`]: strToU8(
      '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
    ),
    [`${dir}/_rels/workbook.xml.rels`]: strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        (vbaBin === undefined
          ? ''
          : `<Relationship Id="rId2" Type="${ms}/vbaProject" Target="vbaProject.bin"/>`) +
        '</Relationships>',
    ),
    [`${dir}/worksheets/sheet1.xml`]: strToU8(
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    ),
  };
  if (vbaBin !== undefined) files[`${dir}/vbaProject.bin`] = vbaBin;
  if (signatures.length > 0) {
    files[`${dir}/_rels/vbaProject.bin.rels`] = strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        signatures
          .map(
            (s, i) => `<Relationship Id="rId${i + 1}" Type="${s.relType}" Target="${s.fileName}"/>`,
          )
          .join('') +
        '</Relationships>',
    );
    for (const s of signatures) files[`${dir}/${s.fileName}`] = s.bytes;
  }
  return zipSync(files);
}

// The PROJECTwm stream pairs each module's MBCS name with its UTF-16 name, both NUL-terminated, ending
// with an empty pair, one record per module the dir/PROJECT streams declare, so removeVbaModule's
// splice (which counts existing records against the parsed module count) has a real structure to shrink.
function buildProjectwmStream(names: readonly string[]): Uint8Array {
  const b: number[] = [];
  for (const name of names) b.push(...ascii(name), 0x00, ...utf16le(name), 0x00, 0x00);
  b.push(0x00, 0x00);
  return Uint8Array.from(b);
}

// Package a project through the *production* CFB writer so it has the navigable red-black sibling tree
// the removeVbaModule/addVbaReference editors walk (buildVbaProjectBin leaves those links null, fine for
// the linear-scan reader, but the editors rebuild the tree). Optionally append raw dir records (e.g. a
// PROJECTREFERENCES entry) and a distinctive _VBA_PROJECT, so a test can prove both survive / are
// replaced as intended.
export function buildNavigableProjectBin(
  codePage: number,
  modules: ModuleSpec[],
  extraDirRecords: number[] = [],
  craftDir: (records: number[]) => number[] = (records) => records,
): Uint8Array {
  const dir = storeCompress(
    Uint8Array.from(craftDir([...buildDirStream(codePage, modules), ...extraDirRecords])),
  );
  const vbaChildren: CfbNode[] = [
    {name: 'dir', data: dir},
    {name: '_VBA_PROJECT', data: Uint8Array.from([0x61, 0xcc, 0x5e, 0x00, 0x00, 0x01, 0x02, 0x03])},
    ...modules.map((m) => ({name: m.name, data: buildModuleStream(m)})),
  ];
  return writeCompoundFile([
    {name: 'PROJECT', data: strToU8(PROJECT_STREAM)},
    {name: 'PROJECTwm', data: buildProjectwmStream(modules.map((m) => m.name))},
    {name: 'VBA', children: vbaChildren},
  ]);
}

export function indexOfBytes(haystack: Uint8Array, needle: Uint8Array): number {
  outer: for (let i = 0; i + needle.length <= haystack.length; i++) {
    for (let j = 0; j < needle.length; j++) if (haystack[i + j] !== needle[j]) continue outer;
    return i;
  }
  return -1;
}
