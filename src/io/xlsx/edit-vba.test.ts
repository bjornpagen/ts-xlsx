// The package-preserving VBA edits: editXlsxVbaRemoveModule and editXlsxVbaAddReference splice the
// project inside an `.xlsm` and leave every other part of the package as it was.

import {strict as assert} from 'node:assert';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {CompoundFile} from '../../vba/cfb.ts';
import {VbaAuthorError} from '../../vba/errors.ts';
import {decompressContainer} from '../../vba/ms-ovba.ts';
import {parseVbaProject} from '../../vba/project.ts';
import {
  ascii,
  buildNavigableProjectBin,
  CODE_PAGE,
  indexOfBytes,
  legacySignature,
  MODULES,
  rec,
  xlsmPackage,
} from '../../vba/vba.test-support.ts';
import {PackageReadError} from '../opc/errors.ts';
import {editXlsxVbaAddReference, editXlsxVbaRemoveModule} from './edit-vba.ts';
import {optionalPartText, partBytes, partsOf} from './package.test-support.ts';
import {writeXlsx} from './write.ts';

test('the package-level edits enforce the same inflate ceiling every other reader does', () => {
  // Both take raw caller-supplied bytes off the public entry, which makes them readers, and a reader
  // that unzips without a bound believes whatever the archive expands to. They reach the shared
  // inflater now, so the bomb guard fires here on the same terms and with the same message it does for
  // `readXlsx`, and the ceiling is settable on the same option.
  const pkg = xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES));

  for (const edit of [
    () => editXlsxVbaRemoveModule(pkg, 'Module1', {maxUncompressedBytes: 512}),
    () =>
      editXlsxVbaAddReference(
        pkg,
        {
          name: 'Scripting',
          guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
          majorVersion: 1,
          minorVersion: 0,
          path: 'C:\\Windows\\System32\\scrrun.dll',
        },
        {maxUncompressedBytes: 512},
      ),
  ]) {
    assert.throws(edit, PackageReadError);
    assert.throws(edit, /possible zip bomb/);
  }

  // And the default ceiling leaves a real package alone.
  assert.ok(editXlsxVbaRemoveModule(pkg, 'Module1').length > 0);
});

test('editXlsxVbaRemoveModule removes a module and preserves every other package part byte-for-byte', () => {
  const refPayload = ascii('*\\Gstdole2.tlb#OLE Automation#REF-MARKER-42');
  const pkg = xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES, rec(0x000d, refPayload)));
  const names = Object.keys(partsOf(pkg)).sort();

  const edited = editXlsxVbaRemoveModule(pkg, 'Module1');

  assert.deepEqual(Object.keys(partsOf(edited)).sort(), names, 'no parts are added or removed');
  for (const name of names) {
    if (name === 'xl/vbaProject.bin') continue;
    assert.deepEqual(
      partBytes(edited, name),
      partBytes(pkg, name),
      `${name} is preserved byte-for-byte`,
    );
  }

  const bin = partBytes(edited, 'xl/vbaProject.bin');
  const project = parseVbaProject(bin);
  assert.deepEqual(
    project.modules.map((m) => m.name),
    ['ThisWorkbook', 'Class1'],
  );
  const reDir = decompressContainer(new CompoundFile(bin).readStream(['VBA', 'dir'])!);
  assert.ok(
    indexOfBytes(reDir, Uint8Array.from(refPayload)) >= 0,
    'the PROJECTREFERENCES record survives the splice',
  );
});

test('editXlsxVbaRemoveModule re-zips with pinned timestamps, so the same edit is reproducible', () => {
  const pkg = xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES));

  assert.deepEqual(
    editXlsxVbaRemoveModule(pkg, 'Module1'),
    editXlsxVbaRemoveModule(pkg, 'Module1'),
    'the same edit twice produces the same bytes',
  );

  // The archive opens with a local file header, whose DOS date/time dword sits at offset 10 with the
  // year (less 1980) in its top bits. Asserting the year alone is enough: from the clock it would be
  // this one, and 2001 can only come from the pin. The full decode lives in write-determinism.test.ts.
  const edited = editXlsxVbaRemoveModule(pkg, 'Module1');
  const dos = new DataView(edited.buffer, edited.byteOffset, edited.byteLength).getUint32(10, true);
  assert.equal(((dos >>> 25) & 0x7f) + 1980, 2001, 'entries are stamped, not clocked');
});

test('editXlsxVbaRemoveModule drops a stale signature part', () => {
  const pkg = xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES), {
    signatures: [legacySignature(Uint8Array.from([1, 2, 3]))],
  });
  assert.deepEqual(
    partBytes(pkg, 'xl/vbaProjectSignature.bin'),
    Uint8Array.from([1, 2, 3]),
    'precondition: the package is signed',
  );

  assert.equal(
    optionalPartText(editXlsxVbaRemoveModule(pkg, 'Module1'), 'xl/vbaProjectSignature.bin'),
    undefined,
    'the signature part is dropped',
  );
});

test('editXlsxVbaRemoveModule finds a project whose workbook is not at xl/workbook.xml', () => {
  // OPC fixes nothing but `_rels/.rels`, so a producer other than Excel may put the office document
  // anywhere, and every target below it resolves relative to its referrer.
  const pkg = xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES), {documentDir: 'xl/sub'});

  const edited = editXlsxVbaRemoveModule(pkg, 'Module1');

  const project = parseVbaProject(partBytes(edited, 'xl/sub/vbaProject.bin'));
  assert.deepEqual(
    project.modules.map((m) => m.name),
    ['ThisWorkbook', 'Class1'],
    'the relationship target resolves against the workbook that declares it, not against a fixed xl/',
  );
});

test('editXlsxVbaRemoveModule throws for a macro-free package', () => {
  const wb = new Workbook();
  wb.addWorksheet('Sheet1');
  assert.throws(() => editXlsxVbaRemoveModule(writeXlsx(wb), 'Module1'), VbaAuthorError);
});

test('editXlsxVbaRemoveModule propagates VbaAuthorError for an unknown module', () => {
  const pkg = xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES));
  assert.throws(() => editXlsxVbaRemoveModule(pkg, 'Nope'), VbaAuthorError);
});

test('editXlsxVbaRemoveModule propagates VbaAuthorError for a document module', () => {
  const pkg = xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES));
  assert.throws(() => editXlsxVbaRemoveModule(pkg, 'ThisWorkbook'), VbaAuthorError);
});

test('editXlsxVbaAddReference adds a reference and preserves every other package part byte-for-byte', () => {
  const pkg = xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES));

  const edited = editXlsxVbaAddReference(pkg, {
    name: 'Scripting',
    guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
    majorVersion: 1,
    minorVersion: 0,
    path: 'C:\\Windows\\System32\\scrrun.dll',
  });

  for (const name of Object.keys(partsOf(pkg))) {
    if (name === 'xl/vbaProject.bin') continue;
    assert.deepEqual(
      partBytes(edited, name),
      partBytes(pkg, name),
      `${name} is preserved byte-for-byte`,
    );
  }

  const bin = partBytes(edited, 'xl/vbaProject.bin');
  const reDir = decompressContainer(new CompoundFile(bin).readStream(['VBA', 'dir'])!);
  assert.ok(
    indexOfBytes(reDir, Uint8Array.from(ascii('Scripting'))) >= 0,
    'the new reference is present in the spliced project',
  );
  assert.deepEqual(
    parseVbaProject(bin).modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
    'the module set is unaffected',
  );
});

test('editXlsxVbaAddReference throws for a macro-free package', () => {
  const wb = new Workbook();
  wb.addWorksheet('Sheet1');
  assert.throws(
    () =>
      editXlsxVbaAddReference(writeXlsx(wb), {
        name: 'Scripting',
        guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
        majorVersion: 1,
        minorVersion: 0,
        path: 'C:\\Windows\\System32\\scrrun.dll',
      }),
    VbaAuthorError,
  );
});

test('editXlsxVbaAddReference propagates VbaAuthorError for an invalid reference', () => {
  const pkg = xlsmPackage(buildNavigableProjectBin(CODE_PAGE, MODULES));
  assert.throws(
    () =>
      editXlsxVbaAddReference(pkg, {
        name: '1Bad',
        guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
        majorVersion: 1,
        minorVersion: 0,
        path: 'C:\\Windows\\System32\\scrrun.dll',
      }),
    VbaAuthorError,
  );
});
