// Entry barrel for `@shbernal/ts-xlsx/vba`: the native VBA read view and the structural edits.
//
// A deliberately narrower face than `src/vba/index.ts`, which is the *internal* barrel and also
// carries the error classes, the part-path and relationship constants, and the signature-kind reader
// that `Workbook` needs. Those are implementation, not API.

export {
  parseVbaProject,
  type VbaModule,
  type VbaModuleKind,
  type VbaProject,
  type VbaProjectSignature,
  type VbaProjectSignatureKind,
} from '../vba/project.ts';
export {addVbaReference, removeVbaModule, type VbaLibraryReference} from '../vba/project-editor.ts';
