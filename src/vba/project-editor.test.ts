// The structural edits on a `vbaProject.bin`: removeVbaModule and addVbaReference, which splice the
// dir, PROJECT and PROJECTwm streams and rebuild the container around them.

import {strict as assert} from 'node:assert';
import {test} from 'node:test';

import {strFromU8, strToU8} from 'fflate';

import {type CfbNode, writeCompoundFile} from './cfb-writer.ts';
import {CompoundFile} from './cfb.ts';
import {VbaAuthorError, VbaParseError} from './errors.ts';
import {compressContainer, decompressContainer} from './ms-ovba.ts';
import {addVbaReference, removeVbaModule} from './project-editor.ts';
import {parseVbaProject} from './project.ts';
import {
  ascii,
  buildDirStream,
  buildNavigableProjectBin,
  CODE_PAGE,
  indexOfBytes,
  MODULES,
  rec,
} from './vba.test-support.ts';

// Flatten what `CompoundFile.tree()` reaches into `/storage/stream` paths. `tree()` walks the sibling
// tree a host navigates rather than the linear directory scan, so a stream the editors unlinked but left
// in place does not appear here.
function streamPaths(nodes: readonly CfbNode[], prefix = ''): string[] {
  return nodes.flatMap((node) =>
    'data' in node
      ? [`${prefix}/${node.name}`]
      : streamPaths(node.children, `${prefix}/${node.name}`),
  );
}

// The dir stream's MODULES_COUNT field ([MS-OVBA] 2.3.4.2.3.2), not cross-checked by parseVbaProject
// (which discovers modules by MODULETERMINATOR markers regardless of the count), but Excel relies on it,
// so removeVbaModule must keep it in sync. Reads it directly out of the decompressed dir bytes.
function readModulesCount(dir: Uint8Array): number {
  let pos = 0;
  while (pos + 6 <= dir.length) {
    const id = (dir[pos] as number) | ((dir[pos + 1] as number) << 8);
    const size =
      ((dir[pos + 2] as number) |
        ((dir[pos + 3] as number) << 8) |
        ((dir[pos + 4] as number) << 16) |
        ((dir[pos + 5] as number) << 24)) >>>
      0;
    const dataStart = pos + 6;
    if (id === 0x000f) return (dir[dataStart] as number) | ((dir[dataStart + 1] as number) << 8);
    pos = dataStart + size;
    if (id === 0x0009) pos += 2; // PROJECTVERSION's uncounted VersionMinor
  }
  throw new Error('MODULES_COUNT not found');
}

// ── Structural edit: removeVbaModule ────────────────────────────────────────────────────────────────

test('removeVbaModule removes a procedural module, preserving references and untouched modules', () => {
  const refPayload = ascii('*\\Gstdole2.tlb#OLE Automation#REF-MARKER-42');
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES, rec(0x000d, refPayload));

  const removed = removeVbaModule(bin, 'Module1');

  const project = parseVbaProject(removed);
  assert.deepEqual(
    project.modules.map((m) => [m.name, m.kind]),
    [
      ['ThisWorkbook', 'document'],
      ['Class1', 'class'],
    ],
    'Module1 is gone; the other modules and their order survive',
  );

  const before = new CompoundFile(bin);
  const after = new CompoundFile(removed);
  for (const name of ['ThisWorkbook', 'Class1']) {
    assert.deepEqual(
      after.readStream(['VBA', name]),
      before.readStream(['VBA', name]),
      `${name} rides through unchanged`,
    );
  }
  assert.equal(after.readStream(['VBA', 'Module1']), undefined, "Module1's stream is gone");

  // _VBA_PROJECT is preserved untouched: Excel runs the modules' existing p-code, and resetting the
  // cookie would crash the load.
  assert.deepEqual(
    after.readStream(['VBA', '_VBA_PROJECT']),
    before.readStream(['VBA', '_VBA_PROJECT']),
  );

  const dirBefore = decompressContainer(before.readStream(['VBA', 'dir'])!);
  const dirAfter = decompressContainer(after.readStream(['VBA', 'dir'])!);
  assert.ok(
    indexOfBytes(dirAfter, Uint8Array.from(refPayload)) >= 0,
    'the PROJECTREFERENCES record is preserved',
  );
  assert.equal(readModulesCount(dirBefore), 3);
  assert.equal(
    readModulesCount(dirAfter),
    2,
    'MODULES_COUNT is decremented for the removed module',
  );

  const projectText = strFromU8(after.readStream(['PROJECT'])!);
  assert.doesNotMatch(projectText, /^Module=Module1$/m, "Module1's declaration line is gone");
  assert.match(projectText, /^Document=ThisWorkbook\/&H00000000$/m, 'ThisWorkbook line survives');
  assert.match(projectText, /^Class=Class1$/m, 'Class1 line survives');

  assert.equal(
    indexOfBytes(after.readStream(['PROJECTwm'])!, Uint8Array.from(ascii('Module1'))),
    -1,
    'PROJECTwm no longer carries the removed module name',
  );

  assert.deepEqual(
    streamPaths(new CompoundFile(removed).tree()).sort(),
    ['/PROJECT', '/PROJECTwm', '/VBA/Class1', '/VBA/ThisWorkbook', '/VBA/_VBA_PROJECT', '/VBA/dir'],
    'the removed module no longer resolves under the VBA storage by tree navigation',
  );
});

test('removeVbaModule matches the module name case-insensitively', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  const removed = removeVbaModule(bin, 'module1');
  assert.deepEqual(
    parseVbaProject(removed).modules.map((m) => m.name),
    ['ThisWorkbook', 'Class1'],
  );
});

test('removeVbaModule rejects an unknown module name', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  assert.throws(() => removeVbaModule(bin, 'NoSuchModule'), VbaAuthorError);
});

test('removeVbaModule rejects removing a document module fail-closed', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  assert.throws(() => removeVbaModule(bin, 'ThisWorkbook'), VbaAuthorError);
});

// MODULES_COUNT sits ahead of every module block in a dir Excel wrote, which is what makes patching it
// after a splice safe. These two craft a dir where that is false, the way a hostile `.xlsm` would: one
// moves the record behind the block being cut (so the patch would write two bytes into an unrelated
// record's payload), the other declares zero modules while carrying one (so the decrement would store
// 0xffff). Both are refused rather than produced.

// Lift the 8-byte MODULES_COUNT record out of its place and re-seat it just before the dir terminator,
// behind every module block.
function modulesCountAfterTheModules(records: number[]): number[] {
  const at = records.findIndex(
    (byte, i) => byte === 0x0f && records[i + 1] === 0x00 && records[i + 2] === 0x02,
  );
  const record = records.slice(at, at + 8);
  const rest = [...records.slice(0, at), ...records.slice(at + 8)];
  const terminatorAt = rest.length - 6; // the dir Terminator record closes the stream
  return [...rest.slice(0, terminatorAt), ...record, ...rest.slice(terminatorAt)];
}

test('removeVbaModule refuses a dir stream whose MODULES_COUNT sits behind a module block', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES, [], modulesCountAfterTheModules);
  assert.throws(() => removeVbaModule(bin, 'Module1'), {
    name: 'VbaParseError',
    message: /MODULES_COUNT after a module block/,
  });
});

test('removeVbaModule refuses a dir stream that declares zero modules but carries one', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES, [], (records) => {
    const at = records.findIndex(
      (byte, i) => byte === 0x0f && records[i + 1] === 0x00 && records[i + 2] === 0x02,
    );
    const crafted = [...records];
    crafted[at + 6] = 0;
    crafted[at + 7] = 0;
    return crafted;
  });
  assert.throws(() => removeVbaModule(bin, 'Module1'), {
    name: 'VbaParseError',
    message: /zero modules/,
  });
});

test('removeVbaModule rejects a malformed container as a parse error', () => {
  assert.throws(() => removeVbaModule(Uint8Array.from([1, 2, 3, 4]), 'Module1'), VbaParseError);
});

// A module may be named `PROJECT`, the name of the root stream that declares every module. Found by
// name, whichever came first in directory order won: the parse read the root text as a compressed
// module and failed, or, the other way round, a removal wrote the module's bytes over the root text.
test('a module named PROJECT parses in either directory order, and removing another keeps both streams', () => {
  const source = 'Sub P()\r\nEnd Sub';
  const dir = compressContainer(
    Uint8Array.from(
      buildDirStream(1252, [
        {name: 'PROJECT', documentType: false, sourceBytes: [], pcodePrefixLen: 0},
        {name: 'Module1', documentType: false, sourceBytes: [], pcodePrefixLen: 0},
      ]),
    ),
  );
  const declarations: CfbNode = {
    name: 'PROJECT',
    data: strToU8('Module=PROJECT\r\nModule=Module1\r\n'),
  };
  const vba: CfbNode = {
    name: 'VBA',
    children: [
      {name: 'dir', data: dir},
      {name: 'PROJECT', data: compressContainer(strToU8(source))},
      {name: 'Module1', data: compressContainer(strToU8('Sub M()\r\nEnd Sub'))},
    ],
  };

  for (const order of [
    [declarations, vba],
    [vba, declarations],
  ]) {
    const bin = writeCompoundFile(order);
    assert.equal(parseVbaProject(bin).modules.find((m) => m.name === 'PROJECT')?.source, source);

    const removed = removeVbaModule(bin, 'Module1');
    const text = strFromU8(new CompoundFile(removed).readStream(['PROJECT'])!);
    assert.match(text, /^Module=PROJECT$/m, 'the root PROJECT still declares the kept module');
    assert.doesNotMatch(text, /^Module=Module1$/m);
    assert.deepEqual(
      parseVbaProject(removed).modules.map((m) => [m.name, m.source]),
      [['PROJECT', source]],
      'and the module named PROJECT keeps its own source',
    );
  }
});

// ── Structural edit: addVbaReference ─────────────────────────────────────────────────────────────────

// Microsoft Scripting Runtime's real GUID/path: the exact reference this splice was verified against on
// a genuine Excel-authored project (2026-07-23, excel-gui-automation probe, ADR 0012/0013 provenance).
const SCRIPTING_REF = {
  name: 'Scripting',
  displayName: 'Microsoft Scripting Runtime',
  guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
  majorVersion: 1,
  minorVersion: 0,
  path: 'C:\\Windows\\System32\\scrrun.dll',
};
const SCRIPTING_LIBID =
  '*\\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\\Windows\\System32\\scrrun.dll#Microsoft Scripting Runtime';

test('addVbaReference adds a registered reference to a project with no existing reference, without touching PROJECT/PROJECTwm', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  const added = addVbaReference(bin, SCRIPTING_REF);

  // Modules are unaffected: same set, same order.
  assert.deepEqual(
    parseVbaProject(added).modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
  );

  const before = new CompoundFile(bin);
  const after = new CompoundFile(added);
  const dirAfter = decompressContainer(after.readStream(['VBA', 'dir'])!);
  assert.ok(
    indexOfBytes(dirAfter, Uint8Array.from(ascii(SCRIPTING_LIBID))) >= 0,
    'the assembled Libid string is present in the dir stream',
  );
  assert.ok(
    indexOfBytes(dirAfter, Uint8Array.from(ascii('Scripting'))) >= 0,
    'the REFERENCENAME name is present',
  );
  assert.equal(readModulesCount(dirAfter), 3, 'MODULES_COUNT is untouched by adding a reference');

  // No real Excel-authored PROJECT stream carries a Reference= line for a registered library reference
  // (verified against a genuine Excel-authored project) so neither PROJECT nor PROJECTwm changes here.
  assert.deepEqual(after.readStream(['PROJECT']), before.readStream(['PROJECT']));
  assert.deepEqual(after.readStream(['PROJECTwm']), before.readStream(['PROJECTwm']));

  // _VBA_PROJECT is preserved untouched: Excel runs the modules' existing p-code, and resetting the
  // cookie would crash the load.
  assert.deepEqual(
    after.readStream(['VBA', '_VBA_PROJECT']),
    before.readStream(['VBA', '_VBA_PROJECT']),
  );
});

test('addVbaReference adds a reference to an existing project, preserving an existing reference and every module byte-for-byte', () => {
  const existingRefPayload = ascii('*\\Gstdole2.tlb#OLE Automation#REF-MARKER-42');
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES, rec(0x000d, existingRefPayload));

  const added = addVbaReference(bin, SCRIPTING_REF);

  const before = new CompoundFile(bin);
  const after = new CompoundFile(added);
  for (const name of ['ThisWorkbook', 'Module1', 'Class1']) {
    assert.deepEqual(
      after.readStream(['VBA', name]),
      before.readStream(['VBA', name]),
      `${name} rides through unchanged`,
    );
  }

  const dirBefore = decompressContainer(before.readStream(['VBA', 'dir'])!);
  const dirAfter = decompressContainer(after.readStream(['VBA', 'dir'])!);
  assert.ok(
    indexOfBytes(dirAfter, Uint8Array.from(existingRefPayload)) >= 0,
    'the pre-existing reference is preserved',
  );
  assert.ok(
    indexOfBytes(dirAfter, Uint8Array.from(ascii(SCRIPTING_LIBID))) >= 0,
    'the new reference is present',
  );
  assert.equal(readModulesCount(dirBefore), 3);
  assert.equal(readModulesCount(dirAfter), 3, 'MODULES_COUNT is unaffected by adding a reference');

  assert.deepEqual(
    parseVbaProject(added).modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
    'the module set and order are unaffected',
  );
});

test('addVbaReference rejects an invalid reference name', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  assert.throws(() => addVbaReference(bin, {...SCRIPTING_REF, name: '1Bad'}), VbaAuthorError);
});

test('addVbaReference rejects a malformed GUID', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  assert.throws(() => addVbaReference(bin, {...SCRIPTING_REF, guid: 'not-a-guid'}), VbaAuthorError);
});

test('addVbaReference rejects an out-of-range version or LCID', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  assert.throws(() => addVbaReference(bin, {...SCRIPTING_REF, majorVersion: -1}), VbaAuthorError);
  assert.throws(
    () => addVbaReference(bin, {...SCRIPTING_REF, minorVersion: 0x10000}),
    VbaAuthorError,
  );
  assert.throws(() => addVbaReference(bin, {...SCRIPTING_REF, lcid: -1}), VbaAuthorError);
});

test('addVbaReference rejects an invalid path', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  assert.throws(() => addVbaReference(bin, {...SCRIPTING_REF, path: ''}), VbaAuthorError);
  assert.throws(
    () => addVbaReference(bin, {...SCRIPTING_REF, path: 'C:\\has#hash.dll'}),
    VbaAuthorError,
  );
});

test('addVbaReference rejects display text the code page cannot represent', () => {
  const bin = buildNavigableProjectBin(1252, MODULES);
  assert.throws(
    () => addVbaReference(bin, {...SCRIPTING_REF, displayName: '你好'}),
    VbaAuthorError,
  );
});

test('addVbaReference rejects a malformed container as a parse error', () => {
  assert.throws(() => addVbaReference(Uint8Array.from([1, 2, 3, 4]), SCRIPTING_REF), VbaParseError);
});
