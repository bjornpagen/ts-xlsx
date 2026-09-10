// The workbook's VBA surface: the read-only project view, attaching and replacing the raw project and
// what that does to its signatures, and the structural edits reached through the Workbook.

import {strict as assert} from 'node:assert';
import {test} from 'node:test';

import {optionalPartText, partBytes, partText, roundtrip} from '../io/xlsx/package.test-support.ts';
import {readXlsx} from '../io/xlsx/read.ts';
import {writeXlsx} from '../io/xlsx/write.ts';
import {CompoundFile} from '../vba/cfb.ts';
import {VbaAuthorError, VbaParseError} from '../vba/errors.ts';
import {decompressContainer} from '../vba/ms-ovba.ts';
import {
  agileSignature,
  ascii,
  buildNavigableProjectBin,
  buildVbaProjectBin,
  CODE_PAGE,
  indexOfBytes,
  legacySignature,
  MODULES,
  rec,
  v3Signature,
  xlsmPackage,
} from '../vba/vba.test-support.ts';
import {Workbook} from './workbook.ts';

// ── Workbook integration + round-trip non-regression ─────────────────────────────────────────────────

test('Workbook.vbaProject decodes macros from a read .xlsm and memoises', () => {
  const wb = readXlsx(xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES)));
  const project = wb.vbaProject;
  assert.ok(project, 'a macro-enabled workbook exposes its vbaProject');
  assert.deepEqual(
    project.modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
  );
  assert.equal(wb.vbaProject, project, 'the parsed project is memoised, not re-decoded');
});

test('Workbook.vbaProject is undefined for a macro-free workbook', () => {
  assert.equal(readXlsx(xlsmPackage(undefined)).vbaProject, undefined);
});

test('reading vbaProject does not regress byte-for-byte macro preservation on write', () => {
  const vbaBin = buildVbaProjectBin(CODE_PAGE, MODULES);
  const wb = readXlsx(xlsmPackage(vbaBin));
  // Force the read-only projection before writing; it must not perturb the preserved bytes.
  assert.ok(wb.vbaProject);
  assert.deepEqual(
    partBytes(writeXlsx(wb), 'xl/vbaProject.bin'),
    vbaBin,
    'the macro blob is re-emitted byte-for-byte',
  );
  // And it still parses from the re-emitted package.
  assert.deepEqual(
    roundtrip(wb).vbaProject?.modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
  );
});

// ── Attach-blob authoring (§2.1): Workbook.vbaProjectBytes get/set ───────────────────────────────────

test('attaching vbaProjectBytes turns a plain workbook macro-enabled and embeds the blob verbatim', () => {
  const bin = buildVbaProjectBin(CODE_PAGE, MODULES);
  const wb = new Workbook();
  wb.addWorksheet('Sheet1');
  wb.vbaProjectBytes = bin;

  const out = writeXlsx(wb);
  assert.deepEqual(
    partBytes(out, 'xl/vbaProject.bin'),
    bin,
    'the attached blob is embedded byte-for-byte',
  );

  const ct = partText(out, '[Content_Types].xml');
  assert.match(
    ct,
    /application\/vnd\.ms-excel\.sheet\.macroEnabled\.main\+xml/,
    'the workbook part is declared macro-enabled',
  );
  assert.match(ct, /application\/vnd\.ms-office\.vbaProject/, 'the .bin is typed as a vbaProject');

  // The re-read package exposes the same macros.
  assert.deepEqual(
    roundtrip(wb).vbaProject?.modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
  );
});

test('vbaProjectBytes copies a macro project from one workbook to another', () => {
  const source = readXlsx(xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES)));
  const target = new Workbook();
  target.addWorksheet('Sheet1');

  const bytes = source.vbaProjectBytes;
  assert.ok(bytes, 'the source workbook exposes its raw macro blob');
  target.vbaProjectBytes = bytes;

  assert.deepEqual(
    roundtrip(target).vbaProject?.modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
    'the copied project decodes from the target package',
  );
});

test('the lazily-parsed project is re-decoded after the bytes are replaced', () => {
  const wb = new Workbook();
  wb.addWorksheet('Sheet1');
  // Read through a closure, not directly: `assert.equal(wb.vbaProject, undefined)` narrows the
  // getter for the rest of the test, and the whole point here is that a later read differs.
  const moduleNames = (): string[] | undefined => wb.vbaProject?.modules.map((m) => m.name);

  // Reading first is what makes this a test. A macro-free workbook parses to `undefined`, and "no
  // macros" has to stay distinguishable from "not yet decoded", which is why the flag and the value
  // are two fields, and why a write has to clear both.
  assert.equal(moduleNames(), undefined, 'precondition: no macros, and that miss is now cached');

  wb.vbaProjectBytes = buildVbaProjectBin(CODE_PAGE, MODULES);
  assert.deepEqual(
    moduleNames(),
    ['ThisWorkbook', 'Module1', 'Class1'],
    'attaching a project invalidates the cached miss',
  );

  const renamed = MODULES.map((m) => (m.name === 'Module1' ? {...m, name: 'Module2'} : m));
  wb.vbaProjectBytes = buildVbaProjectBin(CODE_PAGE, renamed);
  assert.deepEqual(
    moduleNames(),
    ['ThisWorkbook', 'Module2', 'Class1'],
    'and a replacement is re-decoded, not served from the previous parse',
  );
});

test('the vbaProjectBytes getter returns a defensive copy', () => {
  const wb = readXlsx(xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES)));
  const first = wb.vbaProjectBytes;
  assert.ok(first);
  first.fill(0); // scribble on the returned copy
  const second = wb.vbaProjectBytes;
  assert.ok(second);
  assert.notDeepEqual(second, first, 'mutating a returned copy does not corrupt the stored blob');
  // The stored blob still round-trips and parses.
  assert.ok(roundtrip(wb).vbaProject);
});

test('assigning undefined removes the macro project, reverting to a plain package', () => {
  const wb = readXlsx(xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES)));
  assert.ok(wb.vbaProject, 'precondition: the workbook has macros');
  wb.vbaProjectBytes = undefined;

  assert.equal(wb.vbaProject, undefined, 'the read view reflects the removal');
  assert.equal(wb.vbaProjectBytes, undefined, 'no blob remains attached');

  const out = writeXlsx(wb);
  assert.equal(
    optionalPartText(out, 'xl/vbaProject.bin'),
    undefined,
    'the package no longer carries a vbaProject.bin',
  );
  const ct = partText(out, '[Content_Types].xml');
  assert.doesNotMatch(ct, /macroEnabled/, 'the workbook is no longer declared macro-enabled');
});

test('attaching a malformed blob is rejected fail-closed and leaves the workbook untouched', () => {
  const wb = readXlsx(xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES)));
  const original = wb.vbaProjectBytes;

  assert.throws(() => {
    wb.vbaProjectBytes = Uint8Array.from([1, 2, 3, 4]); // not a CFB container
  }, VbaParseError);

  // The reject happens before the old project is cleared, so the workbook is unchanged.
  assert.deepEqual(
    wb.vbaProjectBytes,
    original,
    'a rejected attach does not disturb the existing blob',
  );
});

test('replacing the project drops a now-stale signature over the old bytes', () => {
  const oldBin = buildVbaProjectBin(CODE_PAGE, MODULES);
  const sig = Uint8Array.from([0xde, 0xad, 0xbe, 0xef]);
  const wb = readXlsx(xlsmPackage(oldBin, {signatures: [legacySignature(sig)]}));
  // Precondition: the signature part is present in the read package.
  assert.deepEqual(
    partBytes(writeXlsx(wb), 'xl/vbaProjectSignature.bin'),
    sig,
    'signature present before replace',
  );
  assert.equal(wb.vbaProjectSigned, true, 'the accessor sees the signature before replace');

  const newBin = buildVbaProjectBin(1252, [
    {
      name: 'Module1',
      documentType: false,
      sourceBytes: ascii('Sub Fresh()\r\nEnd Sub'),
      pcodePrefixLen: 8,
    },
  ]);
  wb.vbaProjectBytes = newBin;

  const out = writeXlsx(wb);
  assert.deepEqual(partBytes(out, 'xl/vbaProject.bin'), newBin, 'the new blob is embedded');
  assert.equal(
    optionalPartText(out, 'xl/vbaProjectSignature.bin'),
    undefined,
    'the stale signature over the old bytes is dropped, not left to advertise a broken signature',
  );
  assert.equal(
    wb.vbaProjectSigned,
    false,
    'the drop is visible through the accessor: a replaced project reads unsigned',
  );
  assert.deepEqual(wb.vbaProjectSignatures, [], 'no signatures remain after replace');
  assert.deepEqual(
    roundtrip(wb).vbaProject?.modules.map((m) => m.name),
    ['Module1'],
    'the replacement project decodes',
  );
});

test('Workbook.vbaProjectSigned reports a signed project and exposes the raw signature bytes', () => {
  const sig = Uint8Array.from([0xde, 0xad, 0xbe, 0xef, 4, 5]);
  const wb = readXlsx(
    xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES), {signatures: [legacySignature(sig)]}),
  );

  assert.equal(wb.vbaProjectSigned, true);
  assert.equal(wb.vbaProjectSignatures.length, 1);
  const only = wb.vbaProjectSignatures[0];
  assert.ok(only);
  assert.equal(only.kind, 'legacy');
  assert.deepEqual(only.bytes, sig, 'the raw signature blob is passed through verbatim');

  // The getter hands back a defensive copy; mutating it must not corrupt the preserved bytes.
  only.bytes[0] = 0;
  assert.equal(
    wb.vbaProjectSignatures[0]?.bytes[0],
    0xde,
    'the stored signature bytes are unchanged',
  );
});

test('Workbook.vbaProjectSigned reads false for an unsigned macro project', () => {
  const wb = readXlsx(xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES)));
  assert.equal(wb.vbaProjectSigned, false);
  assert.deepEqual(wb.vbaProjectSignatures, []);
});

test('Workbook.vbaProjectSigned reads false for a workbook with no macros', () => {
  const wb = new Workbook();
  assert.equal(wb.vbaProjectSigned, false);
  assert.deepEqual(wb.vbaProjectSignatures, []);
});

test('Workbook.vbaProjectSignatures reports every generation, across the years each URI carries', () => {
  const wb = readXlsx(
    xlsmPackage(buildVbaProjectBin(CODE_PAGE, MODULES), {
      signatures: [
        legacySignature(Uint8Array.from([1])),
        agileSignature(Uint8Array.from([2])),
        v3Signature(Uint8Array.from([3])),
      ],
    }),
  );

  assert.equal(wb.vbaProjectSigned, true);
  // Detection keys off the relationship Type's final segment, so all three are found even though their
  // URIs carry different year segments (2006 / 2014 / 2020); the closure walk preserves each part.
  assert.deepEqual(
    wb.vbaProjectSignatures.map((s) => s.kind),
    ['legacy', 'agile', 'v3'],
  );
  assert.deepEqual(
    wb.vbaProjectSignatures.map((s) => [...s.bytes]),
    [[1], [2], [3]],
  );
});

// ── Structural edits through the public surface: Workbook.removeVbaModule / addVbaReference ───────────

test('Workbook.removeVbaModule removes a module from a read workbook and preserves references end-to-end', () => {
  const refPayload = ascii('*\\Gstdole2.tlb#OLE Automation#REF-MARKER-42');
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES, rec(0x000d, refPayload));
  const wb = readXlsx(xlsmPackage(bin));

  wb.removeVbaModule('Module1');

  const reDir = decompressContainer(
    new CompoundFile(partBytes(writeXlsx(wb), 'xl/vbaProject.bin')).readStream('dir')!,
  );
  assert.ok(
    indexOfBytes(reDir, Uint8Array.from(refPayload)) >= 0,
    'the PROJECTREFERENCES record survives the whole package round-trip',
  );

  const reread = roundtrip(wb);
  assert.deepEqual(
    reread.vbaProject?.modules.map((m) => [m.name, m.kind]),
    [
      ['ThisWorkbook', 'document'],
      ['Class1', 'class'],
    ],
  );
});

test('Workbook.removeVbaModule on a macro-free workbook throws', () => {
  const wb = new Workbook();
  wb.addWorksheet('Sheet1');
  assert.throws(() => wb.removeVbaModule('Module1'), VbaAuthorError);
});

test('Workbook.removeVbaModule rejects an unknown module without disturbing the existing project', () => {
  const wb = readXlsx(xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES)));
  const before = wb.vbaProjectBytes;

  assert.throws(() => wb.removeVbaModule('Nope'), VbaAuthorError);
  assert.deepEqual(wb.vbaProjectBytes, before, 'a rejected removal leaves the workbook untouched');
});

test('Workbook.removeVbaModule rejects removing a document module fail-closed', () => {
  const wb = readXlsx(xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES)));
  assert.throws(() => wb.removeVbaModule('ThisWorkbook'), VbaAuthorError);
});

test('Workbook.addVbaReference adds a reference to a read workbook, preserving modules', () => {
  const bin = buildNavigableProjectBin(CODE_PAGE, MODULES);
  const wb = readXlsx(xlsmPackage(bin));

  wb.addVbaReference({
    name: 'Scripting',
    guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
    majorVersion: 1,
    minorVersion: 0,
    path: 'C:\\Windows\\System32\\scrrun.dll',
  });

  const reDir = decompressContainer(
    new CompoundFile(partBytes(writeXlsx(wb), 'xl/vbaProject.bin')).readStream('dir')!,
  );
  assert.ok(
    indexOfBytes(reDir, Uint8Array.from(ascii('Scripting'))) >= 0,
    'the new reference is present after a full package round-trip',
  );
  assert.deepEqual(
    roundtrip(wb).vbaProject?.modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
    'the module set is unaffected',
  );
});

test('Workbook.addVbaReference on a macro-free workbook throws', () => {
  const wb = new Workbook();
  wb.addWorksheet('Sheet1');
  assert.throws(
    () =>
      wb.addVbaReference({
        name: 'Scripting',
        guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
        majorVersion: 1,
        minorVersion: 0,
        path: 'C:\\Windows\\System32\\scrrun.dll',
      }),
    VbaAuthorError,
  );
});

test('Workbook.addVbaReference rejects an invalid reference without disturbing the existing project', () => {
  const wb = readXlsx(xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES)));
  const before = wb.vbaProjectBytes;

  assert.throws(
    () =>
      wb.addVbaReference({
        name: '1Bad',
        guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
        majorVersion: 1,
        minorVersion: 0,
        path: 'C:\\Windows\\System32\\scrrun.dll',
      }),
    VbaAuthorError,
  );
  assert.deepEqual(wb.vbaProjectBytes, before, 'a rejected add leaves the workbook untouched');
});
