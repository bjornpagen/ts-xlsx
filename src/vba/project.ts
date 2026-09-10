// Turn a raw `vbaProject.bin` into readable module source.
//
// Pipeline ([MS-OVBA] 2.3.4.2): the CFB container holds a `VBA/dir` stream (itself MS-OVBA compressed)
// that lists each module's stream name, its code page, and the byte offset in that module's stream
// where the compressed source text begins, past the p-code / PerformanceCache. We decompress `dir`,
// read those facts, then decompress each module stream from its text offset and decode with the
// project code page. The p-code is version-specific and deliberately not exposed; a reader wants source.

import {quoted} from '../errors.ts';
import {readU16, readU32} from './bytes.ts';
import {CompoundFile} from './cfb.ts';
import {type Decoder, decoderForCodePage} from './codepage.ts';
import {
  dirRecords,
  REC_MODULE_NAME,
  REC_MODULE_OFFSET,
  REC_MODULE_STREAMNAME,
  REC_MODULE_TERMINATOR,
  REC_MODULE_TYPE_DOCUMENT,
  REC_MODULE_TYPE_PROCEDURAL,
  REC_PROJECT_CODEPAGE,
} from './dir-records.ts';
import {VbaParseError} from './errors.ts';
import {decompressContainer} from './ms-ovba.ts';

/** How a module participates in the project: the classification the VBA editor shows. */
export type VbaModuleKind = 'procedural' | 'document' | 'class' | 'designer';

export interface VbaModule {
  /** The module's code name as seen in the VBA editor, e.g. `ThisWorkbook`, `JsonConverter`. */
  readonly name: string;
  /** The CFB stream the module's bytes live in, usually equal to {@link name}. */
  readonly streamName: string;
  /** Procedural (`.bas`), document code-behind, class module, or designer (UserForm). */
  readonly kind: VbaModuleKind;
  /** The decompressed VBA source (p-code and PerformanceCache are not included). */
  readonly source: string;
}

export interface VbaProject {
  /** The project code page (`PROJECTCODEPAGE`) used to decode module names and source. */
  readonly codePage: number;
  /** The project's modules, in declaration order. */
  readonly modules: readonly VbaModule[];
}

// The OPC packaging identity of a workbook's VBA project: the relationship Type URI Excel wires it
// under, the content type its `.bin` part declares, and the canonical package path. These are the
// facts an authoring path needs to attach a macro blob so the writer emits a valid macro-enabled
// package; the reader matches the same relationship type by suffix (`isPreservedWorkbookRelType`).
export const VBA_PROJECT_REL_TYPE =
  'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
export const VBA_PROJECT_CONTENT_TYPE = 'application/vnd.ms-office.vbaProject';
export const VBA_PROJECT_PART_PATH = 'xl/vbaProject.bin';

/**
 * Which generation of VBA project signature a part is. Office emits up to three sibling signature
 * parts off `vbaProject.bin`'s own rels over the same project bytes ([MS-OFFMACRO2]): the original
 * `legacy` signature, the `agile` (V2) successor, and the `v3` scheme that closes a tampering hole
 * the earlier two left open (KB5000676). All three can coexist in one package.
 */
export type VbaProjectSignatureKind = 'legacy' | 'agile' | 'v3';

/** One digital signature over a workbook's VBA project: its generation and its raw signature bytes. */
export interface VbaProjectSignature {
  readonly kind: VbaProjectSignatureKind;
  /**
   * The raw signature part bytes (a PKCS#7/CMS blob), passed through verbatim. This library does not
   * parse or cryptographically verify them. Their presence means "a signature is attached," never
   * "this signature is valid."
   */
  readonly bytes: Uint8Array;
}

// A signature part is matched by the FINAL segment of its relationship Type, not the whole URI: the
// segment (`vbaProjectSignature{,Agile,V3}`) names the generation, while the year the URI carries
// (2006 / 2014 / 2020) is incidental and varies. The closure walk already carries any such part
// through verbatim regardless of type, so recognising the generation token is all detection needs:
// no exact URI to pin, and a future scheme this map does not know stays preserved but unreported.
const SIGNATURE_KIND_BY_REL_SEGMENT: ReadonlyMap<string, VbaProjectSignatureKind> = new Map([
  ['vbaProjectSignature', 'legacy'],
  ['vbaProjectSignatureAgile', 'agile'],
  ['vbaProjectSignatureV3', 'v3'],
]);

/**
 * The VBA-signature generation a relationship Type names, or `undefined` if it is not a signature
 * relationship. Keys off the Type's final path segment, so it is independent of the year the URI
 * carries (`.../office/2006/...` vs `.../2014/...` vs `.../2020/...`).
 */
export function vbaProjectSignatureKind(relType: string): VbaProjectSignatureKind | undefined {
  return SIGNATURE_KIND_BY_REL_SEGMENT.get(relType.slice(relType.lastIndexOf('/') + 1));
}

interface PendingModule {
  name?: Uint8Array;
  streamName?: Uint8Array;
  offset?: number;
  documentType?: boolean;
}

/**
 * A {@link PendingModule} that reached its MODULE_TERMINATOR carrying both records a module cannot be
 * read without. Narrowing at the push is what lets the map below read `offset` and `streamName`
 * directly; without it the only way to spend them was an assertion restating a check thirty lines
 * away, which is exactly the kind that survives the check being changed.
 */
type CompleteModule = PendingModule & {streamName: Uint8Array; offset: number};

function isComplete(module: PendingModule): module is CompleteModule {
  return module.streamName !== undefined && module.offset !== undefined;
}

// Where [MS-OVBA] 2.2 puts each stream this layer reads: `PROJECT` and `PROJECTwm` at the root, `dir`
// and every module stream in the `VBA` storage. Each is read by its path, because a name alone repeats
// across storages (see `CompoundFile.readStream`).
export const VBA_STORAGE = 'VBA';
export const DIR_PATH: readonly string[] = [VBA_STORAGE, 'dir'];
export const PROJECT_PATH: readonly string[] = ['PROJECT'];
export const PROJECTWM_PATH: readonly string[] = ['PROJECTwm'];

export function parseVbaProject(
  bin: Uint8Array,
  maxOutput = DEFAULT_MAX_PROJECT_OUTPUT,
): VbaProject {
  return parseVbaProjectIn(new CompoundFile(bin), maxOutput);
}

/**
 * As {@link parseVbaProject}, over a container the caller has already opened.
 *
 * Both project editors parse fail-closed first and then need the container itself, and each was
 * building a second `CompoundFile` over the same bytes: a second walk of the DIFAT, FAT, mini-FAT,
 * directory and mini-stream, which is roughly double the cost and double the peak allocation of every
 * edit. Opening it once and handing it in is the whole of the fix.
 */
export function parseVbaProjectIn(
  cfb: CompoundFile,
  maxOutput = DEFAULT_MAX_PROJECT_OUTPUT,
): VbaProject {
  const budget = new DecompressionBudget(maxOutput);

  const dirCompressed = cfb.readStream(DIR_PATH);
  if (!dirCompressed) throw new VbaParseError("VBA project has no 'dir' stream");
  const dir = budget.spend(dirCompressed, 0);

  let codePage = 1252; // Western-European default until PROJECTCODEPAGE says otherwise.
  const rawModules: CompleteModule[] = [];
  let pending: PendingModule = {};

  for (const {id, dataStart, size} of dirRecords(dir, 'overruns stream')) {
    switch (id) {
      case REC_PROJECT_CODEPAGE:
        if (size >= 2) codePage = readU16(dir, dataStart);
        break;
      case REC_MODULE_NAME:
        pending.name = dir.subarray(dataStart, dataStart + size);
        break;
      case REC_MODULE_STREAMNAME:
        pending.streamName = dir.subarray(dataStart, dataStart + size);
        break;
      case REC_MODULE_TYPE_PROCEDURAL:
        pending.documentType = false;
        break;
      case REC_MODULE_TYPE_DOCUMENT:
        pending.documentType = true;
        break;
      case REC_MODULE_OFFSET:
        if (size >= 4) pending.offset = readU32(dir, dataStart);
        break;
      case REC_MODULE_TERMINATOR:
        if (isComplete(pending)) rawModules.push(pending);
        pending = {};
        break;
      default:
        break;
    }
  }

  const decoder = decoderForCodePage(codePage);
  const kindByName = readProjectStreamKinds(cfb, decoder);

  const modules: VbaModule[] = rawModules.map((m) => {
    const streamName = decoder.decode(m.streamName);
    const name = m.name !== undefined ? decoder.decode(m.name) : streamName;
    const kind = kindByName.get(name) ?? (m.documentType ? 'document' : 'procedural');
    return {
      name,
      streamName,
      kind,
      source: readModuleSource(cfb, streamName, m.offset, decoder, budget),
    };
  });

  return {codePage, modules};
}

function readModuleSource(
  cfb: CompoundFile,
  streamName: string,
  textOffset: number,
  decoder: Decoder,
  budget: DecompressionBudget,
): string {
  const stream = cfb.readStream([VBA_STORAGE, streamName]);
  if (!stream) {
    throw new VbaParseError(
      `module stream ${quoted(streamName)} not found in the ${quoted(VBA_STORAGE)} storage`,
    );
  }
  return decoder.decode(budget.spend(stream, textOffset));
}

/**
 * The ceiling on everything one project decompresses, across its `dir` stream and every module.
 *
 * `decompressContainer`'s own ceiling is per call, which bounds a single bomb and nothing else: a
 * project is one `dir` plus one stream per module, all retained at once, and neither the module count
 * nor the total was bounded. Measured amplification is around 230:1, so a 30 MB `vbaProject.bin` with a
 * hundred module streams reached roughly 6.4 GB of retained bytes, doubled again as JS strings, while
 * every individual call stayed comfortably under its own limit.
 *
 * `io/opc/inflate.ts` answers the same question for the zip container with one running counter across
 * all parts. This is that counter for the VBA subsystem, which had the guard but not the accounting.
 * 64 MiB is far above any real project (a large one is a few hundred KB) and small enough that the
 * worst case is a typed failure rather than an OOM.
 */
const DEFAULT_MAX_PROJECT_OUTPUT = 64 * 1024 * 1024;

class DecompressionBudget {
  #remaining: number;

  constructor(total: number) {
    this.#remaining = total;
  }

  /** Decompress one container against what is left, and charge the output to the budget. */
  spend(stream: Uint8Array, start: number): Uint8Array {
    const out = decompressContainer(stream, start, this.#remaining);
    this.#remaining -= out.length;
    return out;
  }
}

// A Map, not an object literal: the keyword is whatever text precedes the `=` on a PROJECT line, so
// an object would answer `constructor=Foo` with the Object function and publish it as a module kind.
const MODULE_KIND_BY_PROJECT_KEYWORD: ReadonlyMap<string, VbaModuleKind> = new Map([
  ['Document', 'document'],
  ['Module', 'procedural'],
  ['Class', 'class'],
  ['BaseClass', 'designer'],
]);

// The `PROJECT` stream (uncounted, plain MBCS text) declares each module's exact kind by keyword:
// `Document=`, `Module=` (procedural .bas), `Class=`, `BaseClass=` (designer/UserForm). The `dir`
// stream's MODULETYPE only distinguishes procedural from non-procedural, so PROJECT refines it. Absent
// or unparsable PROJECT just falls back to that coarser MODULETYPE classification.
function readProjectStreamKinds(cfb: CompoundFile, decoder: Decoder): Map<string, VbaModuleKind> {
  const kinds = new Map<string, VbaModuleKind>();
  const stream = cfb.readStream(PROJECT_PATH);
  if (!stream) return kinds;
  const text = decoder.decode(stream);
  for (const line of text.split(/\r\n|\r|\n/)) {
    const eq = line.indexOf('=');
    if (eq < 0) continue;
    const kind = MODULE_KIND_BY_PROJECT_KEYWORD.get(line.slice(0, eq));
    if (kind === undefined) continue;
    // Value is `Name` or `Name/&H00000000` (document modules carry a cookie); take the name.
    const name = (line.slice(eq + 1).split('/')[0] ?? '').trim();
    if (name) kinds.set(name, kind);
  }
  return kinds;
}
