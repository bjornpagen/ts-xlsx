// The project parse: a `vbaProject.bin` into its modules, code page and kinds.

import {strict as assert} from 'node:assert';
import {test} from 'node:test';

import {strToU8} from 'fflate';

import {writeCompoundFile} from './cfb-writer.ts';
import {VbaParseError} from './errors.ts';
import {compressContainer} from './ms-ovba.ts';
import {parseVbaProject, vbaProjectSignatureKind} from './project.ts';
import {buildDirStream, buildVbaProjectBin, CODE_PAGE, MODULES} from './vba.test-support.ts';

test('parseVbaProject decodes modules, code page, kinds, and source past the p-code', () => {
  const project = parseVbaProject(buildVbaProjectBin(CODE_PAGE, MODULES));

  assert.equal(project.codePage, 1251);
  assert.deepEqual(
    project.modules.map((m) => m.name),
    ['ThisWorkbook', 'Module1', 'Class1'],
  );
  assert.deepEqual(
    project.modules.map((m) => m.kind),
    ['document', 'procedural', 'class'], // Class1 (MODULETYPE 0x22) refined to 'class' via PROJECT
  );
  assert.match(project.modules[0]!.source, /Workbook_Open/);
  assert.ok(
    project.modules[1]!.source.includes('А'),
    'code page 1251 byte 0xC0 decodes to Cyrillic А',
  );
});

test('parseVbaProject throws VbaParseError on a corrupt dir stream', () => {
  const bin = buildVbaProjectBin(CODE_PAGE, MODULES);
  assert.throws(() => parseVbaProject(bin.subarray(0, 900)), VbaParseError);
});

test('a PROJECT keyword naming an Object.prototype member does not become a module kind', () => {
  // The keyword table is indexed by text taken straight off the PROJECT stream, so a table that
  // inherits Object.prototype answers `constructor=` with the Object function, and that value is
  // published as VbaModuleKind. An unrecognised keyword must fall back to the dir stream's coarser
  // MODULETYPE, exactly as `Reference=` and `ID=` already do.
  const dir = compressContainer(
    Uint8Array.from(
      buildDirStream(1252, [
        {name: 'Demo', documentType: false, sourceBytes: [], pcodePrefixLen: 0},
      ]),
    ),
  );
  const bin = writeCompoundFile([
    {name: 'PROJECT', data: strToU8('constructor=Demo\r\n')},
    {
      name: 'VBA',
      children: [
        {name: 'dir', data: dir},
        {name: 'Demo', data: compressContainer(strToU8('Sub Demo()\r\nEnd Sub'))},
      ],
    },
  ]);

  const kind = parseVbaProject(bin).modules[0]!.kind;
  assert.equal(typeof kind, 'string');
  assert.equal(kind, 'procedural');
});

test('vbaProjectSignatureKind reports undefined for a rel type naming an Object.prototype member', () => {
  // The final segment comes off a relationship Type in the package, so the table it indexes must
  // not answer with an inherited function that `workbook-vba` would then publish as a kind.
  assert.equal(
    vbaProjectSignatureKind('http://example.invalid/relationships/constructor'),
    undefined,
  );
  assert.equal(
    vbaProjectSignatureKind(
      'http://schemas.microsoft.com/office/2020/relationships/vbaProjectSignatureV3',
    ),
    'v3',
  );
});
