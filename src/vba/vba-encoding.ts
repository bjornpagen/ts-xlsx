// VBA name validation: the identifier rules every name the structural splices in `project-editor.ts`
// write must meet. The `dir` records those splices emit are written by `dir-records.ts`, beside the walk
// that reads them.
//
// There is no from-scratch `vbaProject.bin` synthesizer here. Excel does not recompile VBA from source
// on open (a module runs the compiled p-code it ships) so authoring/editing module SOURCE is done by
// the offline `tools/vba-compiler` (VBIDE), which produces genuinely compiled p-code. This module holds
// only what the pure-TS structural edits (remove module, add reference) still need (ADR 0019).

import {quoted} from '../errors.ts';
import {MAX_NAME_CHARS} from './cfb-format.ts';
import {VbaAuthorError} from './errors.ts';

const IDENTIFIER = /^[A-Za-z][A-Za-z0-9_]*$/;

/**
 * Validate a library reference's name against the VBA identifier contract: a valid identifier, at
 * most 31 characters (the CFB stream-name limit, which VBA also applies to its names). The reference
 * is the one new name {@link project-editor.ts | project-editor}'s structural edits write.
 *
 * @throws {VbaAuthorError} if `name` is not a valid VBA identifier or exceeds 31 characters.
 */
export function validateReferenceName(name: string): void {
  if (!IDENTIFIER.test(name) || name.length > MAX_NAME_CHARS) {
    throw new VbaAuthorError(
      `invalid reference name ${quoted(name)} (must be a VBA identifier ≤ 31 chars)`,
    );
  }
}
