// The .xlsm VBA project: round-tripping it untouched, and the two structural edits the
// project editor supports.

import {strFromU8} from 'fflate';

import type {Untyped} from '../../untyped.ts';
import {partBytesOf, partNamesOf, partOf, patchedPackage} from './package-facts.ts';
import {
  addVbaReference,
  CompoundFile,
  decompressContainer,
  parseVbaProject,
  readXlsx,
  removeVbaModule,
  Workbook,
  writeXlsx,
} from './runtime.ts';
import {
  buildVbaFixtureBin,
  buildVbaFixturePackage,
  VBA_FIXTURE_REF_MARKER,
  vbaAscii,
  vbaBytesIdentical,
  vbaIndexOfBytes,
} from './vba-fixtures.ts';

export const vba = {
  // Splice a synthetic vbaProject.bin + its workbook relationship + content-type override into an
  // otherwise-plain written package, since the writer cannot author a VBA project itself (its
  // bytes are opaque, never modeled) and this is the only way to produce a macro-enabled-shaped
  // package to round-trip. Then read it back and write it again → { originalHasVba,
  // reloadedPreservedCount,
  // rewrittenHasVba, rewrittenIsMacroEnabled }. For asserting a macro-enabled workbook's VBA project
  // and its macro-enabled content-type survive a read/write cycle rather than being silently dropped
  // (an unrecognised workbook relationship is otherwise discarded: real-world data loss on any
  // .xlsm a caller loads and re-saves).
  xlsmVbaProjectRoundtrip() {
    const VBA_REL_TYPE =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vbaProject';
    const wb = new Workbook();
    wb.addWorksheet('S');
    const relId = 'rIdVba';
    const macroPackage = patchedPackage(writeXlsx(wb), {
      edit: {
        'xl/_rels/workbook.xml.rels': (xml) =>
          xml.replace(
            '</Relationships>',
            `<Relationship Id="${relId}" Type="${VBA_REL_TYPE}" Target="vbaProject.bin"/></Relationships>`,
          ),
        '[Content_Types].xml': (xml) =>
          xml.replace(
            '</Types>',
            '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
          ),
      },
      put: {'xl/vbaProject.bin': 'FAKE-VBA-PROJECT-BYTES'},
    });
    const originalHasVba = partNamesOf(macroPackage).includes('xl/vbaProject.bin');

    const loaded = readXlsx(macroPackage);
    const reloadedPreservedCount = (loaded.preservedReferences as Untyped[]).filter((r: Untyped) =>
      r.relType.endsWith('/vbaProject'),
    ).length;

    const rewritten = writeXlsx(loaded);
    const rewrittenHasVba = partNamesOf(rewritten).includes('xl/vbaProject.bin');
    const rewrittenIsMacroEnabled = /macroEnabled\.main\+xml/.test(
      partOf(rewritten, '[Content_Types].xml'),
    );

    return {originalHasVba, reloadedPreservedCount, rewrittenHasVba, rewrittenIsMacroEnabled};
  },

  // Add a registered (COM type-library) reference to an existing macro project's vbaProject.bin via the
  // project-editor primitive addVbaReference. Adds "Microsoft Scripting Runtime" (the real GUID/path,
  // verified against a genuine Excel-authored project) to the three-module, one-reference fixture, and
  // asserts the new REFERENCENAME + REFERENCEREGISTERED records read back correctly, the hand-crafted
  // pre-existing reference and every module survive byte-for-byte, PROJECT/PROJECTwm are untouched (no
  // real Excel-authored PROJECT stream carries a Reference= line for a registered reference), and
  // _VBA_PROJECT is left byte-for-byte unchanged: Excel runs the modules' existing p-code, so resetting
  // the cookie would crash the load (ADR 0019).
  xlsmVbaAddReference() {
    const originalBin = buildVbaFixtureBin();
    const newRef = {
      name: 'Scripting',
      displayName: 'Microsoft Scripting Runtime',
      guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
      majorVersion: 1,
      minorVersion: 0,
      path: 'C:\\Windows\\System32\\scrrun.dll',
    };
    const newLibid =
      '*\\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\\Windows\\System32\\scrrun.dll#Microsoft Scripting Runtime';

    const addedBin = addVbaReference(originalBin, newRef);

    const project = parseVbaProject(addedBin);
    const moduleNames = project.modules.map((m) => m.name);

    const originalCfb = new CompoundFile(originalBin);
    const addedCfb = new CompoundFile(addedBin);
    const untouchedModuleByteIdentical =
      vbaIndexOfBytes(addedCfb.readStream('Module1')!, originalCfb.readStream('Module1')!) === 0 &&
      addedCfb.readStream('Module1')!.length === originalCfb.readStream('Module1')!.length;

    const dirAfter = decompressContainer(addedCfb.readStream('dir')!);
    const existingReferencePreserved =
      vbaIndexOfBytes(dirAfter, Uint8Array.from(vbaAscii(VBA_FIXTURE_REF_MARKER))) >= 0;
    const newReferencePresent = vbaIndexOfBytes(dirAfter, Uint8Array.from(vbaAscii(newLibid))) >= 0;

    const projectStreamUnchanged =
      vbaIndexOfBytes(addedCfb.readStream('PROJECT')!, originalCfb.readStream('PROJECT')!) === 0 &&
      addedCfb.readStream('PROJECT')!.length === originalCfb.readStream('PROJECT')!.length;

    const vbaProjectStreamPreserved = vbaBytesIdentical(
      addedCfb.readStream('_VBA_PROJECT')!,
      originalCfb.readStream('_VBA_PROJECT')!,
    );

    return {
      moduleNames,
      untouchedModuleByteIdentical,
      existingReferencePreserved,
      newReferencePresent,
      projectStreamUnchanged,
      vbaProjectStreamPreserved,
    };
  },

  // Remove a standard module from an existing macro project's vbaProject.bin via the project-editor
  // primitive removeVbaModule. Removes the procedural Module1 from the three-module, one-reference
  // fixture, and asserts: Module1 is gone from the module list; the untouched document module
  // (ThisWorkbook) and the hand-crafted PROJECTREFERENCES record survive byte-for-byte; Module1's
  // declaration line is gone from PROJECT while the other modules' lines survive; Module1's name pair is
  // gone from PROJECTwm; and _VBA_PROJECT is left byte-for-byte unchanged: Excel runs the surviving
  // modules' existing p-code, so resetting the cookie would crash the load (ADR 0019).
  xlsmVbaRemoveModule() {
    const originalBin = buildVbaFixtureBin();

    const removedBin = removeVbaModule(originalBin, 'Module1');

    const project = parseVbaProject(removedBin);
    const moduleNames = project.modules.map((m) => m.name);
    const moduleKinds = project.modules.map((m) => [m.name, m.kind]);

    const originalCfb = new CompoundFile(originalBin);
    const removedCfb = new CompoundFile(removedBin);
    const untouchedModuleByteIdentical =
      vbaIndexOfBytes(
        removedCfb.readStream('ThisWorkbook')!,
        originalCfb.readStream('ThisWorkbook')!,
      ) === 0 &&
      removedCfb.readStream('ThisWorkbook')!.length ===
        originalCfb.readStream('ThisWorkbook')!.length;
    const removedModuleStreamGone = removedCfb.readStream('Module1') === undefined;

    const referencePreserved =
      vbaIndexOfBytes(
        decompressContainer(removedCfb.readStream('dir')!),
        Uint8Array.from(vbaAscii(VBA_FIXTURE_REF_MARKER)),
      ) >= 0;

    const projectText = strFromU8(removedCfb.readStream('PROJECT')!);
    const removedDeclLineGone = !/^Module=Module1$/m.test(projectText);
    const otherDeclLinesSurvive =
      /^Document=ThisWorkbook\/&H00000000$/m.test(projectText) &&
      /^Class=Class1$/m.test(projectText);

    const projectwmNoLongerHasModule1 =
      vbaIndexOfBytes(removedCfb.readStream('PROJECTwm')!, Uint8Array.from(vbaAscii('Module1'))) <
      0;

    const vbaProjectStreamPreserved = vbaBytesIdentical(
      removedCfb.readStream('_VBA_PROJECT')!,
      originalCfb.readStream('_VBA_PROJECT')!,
    );

    return {
      moduleNames,
      moduleKinds,
      untouchedModuleByteIdentical,
      removedModuleStreamGone,
      referencePreserved,
      removedDeclLineGone,
      otherDeclLinesSurvive,
      projectwmNoLongerHasModule1,
      vbaProjectStreamPreserved,
    };
  },

  // Chain the two structural edits through the *public, package-level* surface: Workbook.removeVbaModule,
  // Workbook.addVbaReference, rather than calling the project-editor primitives directly, the way
  // xlsmVbaRemoveModule/xlsmVbaAddReference do. Those cases lock the splice primitives; this one locks
  // that the primitives are actually wired to Workbook and survive a real readXlsx -> edit -> writeXlsx ->
  // readXlsx package round-trip, not only a bare-bin call. Removes the pre-existing Module1 and adds a
  // reference to the three-module, one-reference fixture, then asserts: the final module set/kinds are as
  // expected; the hand-crafted PROJECTREFERENCES record AND the newly added reference are both present;
  // the untouched Class1 module survives byte-for-byte; the package stays macro-enabled; and
  // _VBA_PROJECT is left byte-for-byte unchanged: Excel runs the surviving modules' existing p-code, so
  // resetting the cookie would crash the load (ADR 0019).
  xlsmVbaWorkbookStructuralEdits() {
    const originalBin = buildVbaFixtureBin();
    const pkg = buildVbaFixturePackage(originalBin);

    const wb = readXlsx(pkg);
    wb.removeVbaModule('Module1');
    const newLibid =
      '*\\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\\Windows\\System32\\scrrun.dll#Microsoft Scripting Runtime';
    wb.addVbaReference({
      name: 'Scripting',
      displayName: 'Microsoft Scripting Runtime',
      guid: '{420B2830-E718-11CF-893D-00A0C9054228}',
      majorVersion: 1,
      minorVersion: 0,
      path: 'C:\\Windows\\System32\\scrrun.dll',
    });

    const written = writeXlsx(wb);
    const rewrittenIsMacroEnabled = /macroEnabled\.main\+xml/.test(
      partOf(written, '[Content_Types].xml'),
    );

    const reread = readXlsx(written);
    const modules = reread.vbaProject?.modules ?? [];
    const moduleNames = modules.map((m) => m.name);
    const moduleKinds = modules.map((m) => [m.name, m.kind]);

    // A project the writer dropped is a reported fact here, not a crash, so absence is allowed.
    const rewrittenBin = partNamesOf(written).includes('xl/vbaProject.bin')
      ? partBytesOf(written, 'xl/vbaProject.bin')
      : undefined;
    const rewrittenCfb = rewrittenBin ? new CompoundFile(rewrittenBin) : undefined;
    const originalCfb = new CompoundFile(originalBin);
    const untouchedModuleByteIdentical =
      rewrittenCfb !== undefined &&
      vbaIndexOfBytes(rewrittenCfb.readStream('Class1')!, originalCfb.readStream('Class1')!) ===
        0 &&
      rewrittenCfb.readStream('Class1')!.length === originalCfb.readStream('Class1')!.length;

    const rewrittenDir = rewrittenCfb
      ? decompressContainer(rewrittenCfb.readStream('dir')!)
      : undefined;
    const originalReferencePreserved =
      rewrittenDir !== undefined &&
      vbaIndexOfBytes(rewrittenDir, Uint8Array.from(vbaAscii(VBA_FIXTURE_REF_MARKER))) >= 0;
    const newReferencePresent =
      rewrittenDir !== undefined &&
      vbaIndexOfBytes(rewrittenDir, Uint8Array.from(vbaAscii(newLibid))) >= 0;

    const vbaProjectStreamPreserved =
      rewrittenCfb !== undefined &&
      vbaBytesIdentical(
        rewrittenCfb.readStream('_VBA_PROJECT')!,
        originalCfb.readStream('_VBA_PROJECT')!,
      );

    return {
      moduleNames,
      moduleKinds,
      rewrittenIsMacroEnabled,
      untouchedModuleByteIdentical,
      originalReferencePreserved,
      newReferencePresent,
      vbaProjectStreamPreserved,
    };
  },
};
