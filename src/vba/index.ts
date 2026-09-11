// Native VBA read, plus pure-TS structural edits: decode a macro-enabled workbook's `vbaProject.bin`
// into readable module source, and remove a module or add a library reference from an existing project
// by splicing the original bytes (every module's compiled p-code preserved untouched).
//
// Authoring or editing module SOURCE is NOT here. Excel runs a module's compiled p-code, not its source,
// and only a real Excel can produce source-matched p-code, so that lives in the offline
// `tools/vba-compiler` (VBIDE), whose output is attached via `Workbook.vbaProjectBytes`. The read path
// is a projection over bytes the model preserves opaquely on round-trip. For the design invariants and
// the wider VBA feature map see the ADRs: read view (docs/decisions/0016), authoring (0017/0019), and
// structural edits (0018/0019), plus docs/knowledge/specs/xlsm-macro-preservation.md.

export {VbaAuthorError, VbaParseError} from './errors.ts';
export {
  parseVbaProject,
  VBA_PROJECT_CONTENT_TYPE,
  VBA_PROJECT_PART_PATH,
  VBA_PROJECT_REL_TYPE,
  type VbaModule,
  type VbaModuleKind,
  type VbaProject,
  type VbaProjectSignature,
  type VbaProjectSignatureKind,
  vbaProjectSignatureKind,
} from './project.ts';
export {addVbaReference, removeVbaModule, type VbaLibraryReference} from './project-editor.ts';
