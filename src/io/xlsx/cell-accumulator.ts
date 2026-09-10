// Cell-gathering state machine for the worksheet-body reader. It owns the still-textual pieces of the
// one `<c>` currently being read (its address/type/style, formula, cached value, inline text, and
// rich-text runs) plus the shared-formula master map that spans the whole sheet. Each `<c>` resets
// the per-cell state on {@link beginCell} and commits it on {@link finalize}, so this class is the
// single owner of "what has this cell gathered so far": to a cell what {@link RunAccumulator} is to a
// rich string. Value *decoding* stays in `cell-value.ts`; this class only gathers the raw pieces.
//
// It also drives itself. {@link openElement}/{@link appendChunk}/{@link closeElement} are the
// `<c>`/`<is>`/`<f>`/`<v>`/`<t>` element machine both worksheet readers used to spell out
// identically, each with its own copy of the three flags it runs on (inside-an-inline-string,
// capturing, captured text). Two readers keeping one machine in step by convention is not a
// mechanism, and the flags are exactly the state that drifts: an element the two disagree about
// silently reads a different value on one path than the other. Each reader now contributes only
// what is genuinely its own -- what committing a cell means, and whether rich runs are read at all
// -- and falls through to this for the rest.

import {encodeAddress, MAX_COLUMN, tryDecodeCellRef} from '../../core/address.ts';
import type {DateEpoch} from '../../core/date.ts';
import {translateFormula, unmangleFunctions} from '../../core/formula.ts';
import type {
  CellValue,
  DataTableFormulaValue,
  FormulaResult,
  SharedFormulaValue,
} from '../../core/value.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {numInteger} from '../../xml/xml-attrs.ts';
import {boolStrict, type XmlAttributes} from '../../xml/xml-scan.ts';
import {applyXfToCell, type XfStyle} from '../style/xf-style.ts';
import {
  decodeCellContent,
  decodeFormulaResult,
  type RawCell,
  type SharedString,
} from './cell-value.ts';
import {RunAccumulator} from './read-rich-runs.ts';

// The declaration attributes of a `<f t="dataTable">`, held from the `<f>` open until the cell
// finalises. A data-table formula is preserved by declaration, not evaluated.
interface DataTableDeclaration {
  ref: string;
  dt2D: string | undefined;
  dtr: string | undefined;
  r1: string | undefined;
  r2: string | undefined;
}

export class CellAccumulator {
  #ref = '';
  #type = '';
  #style = -1;
  #col = -1;
  #row = -1;
  // Where a `<c>` with no `r` of its own sits: the column after the last one placed in this row.
  #nextCol = 1;
  #formula = '';
  // Shared-formula bookkeeping. A master `<f t="shared" ref si>TEXT</f>` seeds the group; every clone
  // `<f t="shared" si/>` in the sheet references it by `si` and carries no text of its own.
  #formulaShared = false;
  #formulaSi = -1;
  #sharedClone = false;
  #dataTable: DataTableDeclaration | null = null;
  #valueText = '';
  #hasFormula = false;
  #hasValue = false;
  readonly #runs: RunAccumulator;
  // Masters always precede their clones (Excel keeps the master top-left), so a clone resolves against
  // a map filled as the sheet streams: the master's formula translated to the clone's position.
  readonly #masters = new Map<number, {formula: string; col: number; row: number}>();

  // The element machine's own state: whether the current element's text is being gathered, and what
  // has been gathered of it. The `<is>`/`<r>`/`<rPr>`/`<t>` half of the grammar, and the capture
  // that goes with it, belongs to {@link RunAccumulator}, which the pooled-string reader drives too.
  #capture = false;
  #text = '';

  // The workbook's date system: not a fact about any cell, but an input to every cell's decode, so it
  // is held for the sheet rather than passed through `finalize`/`decode`/`cachedResult` three times.
  readonly #dateEpoch: DateEpoch;

  constructor(options: {readonly richRuns: boolean; readonly dateEpoch: DateEpoch}) {
    this.#runs = new RunAccumulator({container: 'is', readRuns: options.richRuns});
    this.#dateEpoch = options.dateEpoch;
  }

  /** This cell's `<c r>` address (`"B3"`), or '' when it carried none. */
  get ref(): string {
    return this.#ref;
  }

  /** This cell's own `<c s>` style index, or -1 when it carries none. */
  get styleIndex(): number {
    return this.#style;
  }

  /** This cell's 1-based column, or -1 when its address was unparseable. */
  get col(): number {
    return this.#col;
  }

  /**
   * Open a `<row>`: the cells that follow belong to it, and the next one with no `r` of its own is
   * its first column. Both readers call this where they already tell the style resolver a row opened.
   */
  openRow(number: number): void {
    this.#row = number;
    this.#nextCol = 1;
  }

  // Begin a new `<c>`: record its address/type/style and clear every per-cell gathered field so the
  // last cell's formula, value, runs, or shared/data-table declaration cannot bleed into this one.
  #beginCell(attrs: XmlAttributes): void {
    this.#type = attrs.t ?? '';
    this.#style = numInteger(attrs.s, 0) ?? -1;
    this.#placeCell(attrs.r);
    this.#formula = '';
    this.#valueText = '';
    this.#runs.beginContainer();
    this.#hasFormula = false;
    this.#hasValue = false;
    this.#formulaShared = false;
    this.#formulaSi = -1;
    this.#sharedClone = false;
    this.#dataTable = null;
  }

  /**
   * Resolve where this `<c>` sits, from its `r` or from its position in the row.
   *
   * The two absences are different and used to be conflated. An `r` that names no cell that can exist
   * (`A0`, `ZZZZ1`, `junk!!`) is malformed, and the cell is dropped the way any unreadable foreign
   * attribute is; -1 is the sentinel the shared-formula translation reads for that state. An `r` that
   * is simply *absent* is not malformed at all: `r` is optional on `sml:CT_Cell`, and a producer
   * relying on document position is emitting a legal file that this reader used to lose every cell of.
   */
  #placeCell(ref: string | undefined): void {
    const decoded = ref === undefined ? undefined : tryDecodeCellRef(ref);
    if (decoded !== undefined) {
      this.#ref = ref ?? '';
      this.#col = decoded.col;
      this.#row = decoded.row;
    } else if (ref === undefined && this.#row > 0 && this.#nextCol <= MAX_COLUMN) {
      this.#col = this.#nextCol;
      this.#ref = encodeAddress(this.#col, this.#row);
    } else {
      this.#ref = '';
      this.#col = -1;
      // The row is left alone: it belongs to the open `<row>`, not to this cell, and a malformed `r`
      // must not cost the cells after it their position.
    }
    // A declared `r` re-anchors the count the way it does for rows, so a file mixing the two
    // spellings resumes counting from wherever it last said it was.
    if (this.#col > 0) this.#nextCol = this.#col + 1;
  }

  // Begin an `<f>`: record its shared-formula grouping and any data-table declaration. A self-closing
  // `<f t="shared" si/>` is a clone, firing no close and carrying no text, so mark it here to
  // resolve against its master when the cell finalises.
  #beginFormula(attrs: XmlAttributes, selfClosing: boolean): void {
    this.#formulaShared = attrs.t === 'shared';
    this.#formulaSi = numInteger(attrs.si, 0) ?? -1;
    if (selfClosing && this.#formulaShared) this.#sharedClone = true;
    if (attrs.t === 'dataTable' && attrs.ref !== undefined) {
      this.#dataTable = {
        ref: attrs.ref,
        dt2D: attrs.dt2D,
        dtr: attrs.dtr,
        r1: attrs.r1,
        r2: attrs.r2,
      };
    }
  }

  #setFormula(text: string): void {
    this.#formula = text;
    this.#hasFormula = true;
  }

  #setValue(text: string): void {
    this.#valueText = text;
    this.#hasValue = true;
  }

  /**
   * Drive one element open, and return whether it was one of the cell machine's own. Every open
   * resets the capture state first, which is true of both readers and of every element, not just
   * these. A caller gathering text of its own uses a `TextCapture`, not this machine's buffer.
   */
  openElement(local: string, attrs: XmlAttributes, selfClosing: boolean): boolean {
    this.#text = '';
    this.#capture = false;
    // The string half of the grammar first: `<is>` and everything under it is the run machine's.
    if (this.#runs.open(local, attrs, selfClosing)) return true;
    switch (local) {
      case 'c':
        this.#beginCell(attrs);
        return true;
      case 'f':
        // A self-closing `<f/>` fires no close, so nothing will consume the capture it just armed.
        this.#capture = !selfClosing;
        this.#beginFormula(attrs, selfClosing);
        return true;
      case 'v':
        this.#capture = !selfClosing;
        return true;
      default:
        return false;
    }
  }

  /** Feed one run of character data. Ignored unless something is capturing. */
  appendChunk(chunk: string): void {
    this.#runs.text(chunk);
    if (this.#capture) this.#text += chunk;
  }

  /**
   * Drive one element close. `'cell'` means a `</c>` closed and the caller should commit the
   * gathered cell, which is the one step the two readers do differently; `'claimed'` means the
   * machine handled it; `'other'` leaves it to the caller. Capture always ends here, as it does on
   * every close in both readers.
   */
  closeElement(local: string): 'cell' | 'claimed' | 'other' {
    const verdict = this.#closeElement(local);
    this.#capture = false;
    return verdict;
  }

  #closeElement(local: string): 'cell' | 'claimed' | 'other' {
    // A closing `</is>` completes the inline string, but a cell reads it at `</c>` rather than here,
    // so 'container' is nothing more to this caller than 'claimed'.
    if (this.#runs.close(local) !== 'other') return 'claimed';
    switch (local) {
      case 'f':
        this.#setFormula(this.#text);
        return 'claimed';
      case 'v':
        this.#setValue(this.#text);
        return 'claimed';
      case 'c':
        return 'cell';
      default:
        return 'other';
    }
  }

  // Commit the gathered cell to the sheet with its already-resolved style (the caller applies the
  // cell → row → column inheritance order). A data-table cell surfaces its declaration; a shared-formula
  // master seeds the group before finalising, and a clone resolves to the master translated to its own
  // position. Everything else decodes as an ordinary `<c>` payload. An address-less cell is a no-op.
  finalize(
    sheet: Worksheet,
    sharedStrings: readonly SharedString[],
    style: XfStyle | undefined,
  ): void {
    if (this.#ref === '') return;
    const value = this.#resolveValue(sharedStrings, style);
    const cell = sheet.getCell(this.#ref);
    applyXfToCell(cell, style);
    cell.value = value;
  }

  // Which of the four readings of a `<c>` applies, in the order the format makes them exclusive.
  // Two of the branches are not pure: a shared-formula master seeds the group here as a side effect
  // and then falls through to the ordinary decode, because a master cell *is* an ordinary cell that
  // happens to be shared. A clone whose master is missing falls through too: that is the reading
  // for a file whose shared-formula group is broken, and it must stay a fallthrough rather than a
  // failure.
  #resolveValue(
    sharedStrings: readonly SharedString[],
    style: XfStyle | undefined,
  ): CellValue | DataTableFormulaValue | SharedFormulaValue {
    if (this.#dataTable !== null) {
      return {
        shareType: 'dataTable',
        ref: this.#dataTable.ref,
        ...(boolStrict(this.#dataTable.dt2D) ? {dataTable2D: true} : {}),
        ...(boolStrict(this.#dataTable.dtr) ? {dataTableRow: true} : {}),
        ...(this.#dataTable.r1 !== undefined ? {r1: this.#dataTable.r1} : {}),
        ...(this.#dataTable.r2 !== undefined ? {r2: this.#dataTable.r2} : {}),
        ...this.cachedResult(style),
      };
    }
    if (this.#hasFormula && this.#formulaShared && this.#formulaSi >= 0) {
      this.#masters.set(this.#formulaSi, {formula: this.#formula, col: this.#col, row: this.#row});
    } else if (this.#sharedClone && this.#formulaSi >= 0) {
      const master = this.#masters.get(this.#formulaSi);
      if (master !== undefined) {
        const translated = translateFormula(
          master.formula,
          this.#col - master.col,
          this.#row - master.row,
        );
        return {
          sharedFormula: encodeAddress(master.col, master.row),
          formula: unmangleFunctions(translated),
          // A clone's cached result honours the cell's date format the same way a plain formula's does.
          ...this.cachedResult(style),
        };
      }
    }
    return this.decode(sharedStrings, style);
  }

  // The `result` property of a formula-shaped value, or nothing. A cached `<v>` the decoder cannot
  // read is no cached result, so the key is omitted rather than set to `undefined`: the value is
  // then indistinguishable from a formula cell that carried no `<v>` at all, which is what the
  // writer will emit for it either way.
  private cachedResult(style: XfStyle | undefined): {result?: FormulaResult} {
    if (!this.#hasValue) return {};
    const result = decodeFormulaResult(this.#type, this.#valueText, style?.numFmt, this.#dateEpoch);
    return result === undefined ? {} : {result};
  }

  // Decode the gathered pieces into a plain cell value, resolving the shared pool and date formats but
  // NOT the shared-formula / data-table declarations {@link finalize} handles. This is what a data
  // read (the streaming reader) wants: the cell's own value, with a shared-formula clone surfacing its
  // cached result rather than a translated formula it will not evaluate.
  decode(sharedStrings: readonly SharedString[], style: XfStyle | undefined): CellValue {
    const raw: RawCell = {
      type: this.#type,
      hasFormula: this.#hasFormula,
      formula: this.#formula,
      hasValue: this.#hasValue,
      valueText: this.#valueText,
      inlineText: this.#runs.plainText,
      richTextRuns: this.#runs.runs,
    };
    return decodeCellContent(raw, sharedStrings, style?.numFmt, this.#dateEpoch);
  }
}
