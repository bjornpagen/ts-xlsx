// Where a `<row>` sits, when the file does not say.
//
// `r` is optional on `sml:CT_Row` (and on `sml:CT_Cell`), because SpreadsheetML lets a producer rely
// on document position instead: the nth `<row>` of `<sheetData>` is row n, and the nth `<c>` of a row
// is that row's nth column. Excel always writes the attribute, so a reader tested only against Excel
// output never meets the other spelling; several non-Excel writers emit it, and against them a reader
// that requires `r` loses the whole sheet.
//
// The inference is one line, which is exactly why it belongs in one place: the two worksheet readers
// each had a copy, and they had already drifted. One inferred positionally and the other dropped the
// row's formatting outright, so the same file read one way through `readXlsx` and another through the
// row streamer. `CellAccumulator` and `CellStyleResolver` exist to make that class of divergence
// structurally impossible, and this is the third decision that belongs beside them.

import {MAX_ROW} from '../../core/address.ts';
import {numInteger} from '../../xml/xml-attrs.ts';
import {type XmlAttributes} from '../../xml/xml-scan.ts';

/** One `<row>`, as the reader that opened it needs to see it. */
export interface RowPosition {
  /** The row's 1-based number, inferred from position when the element declares no `r`. */
  readonly number: number;
  /**
   * Whether that number names a row the grid has. A row past the ceiling is dropped whole by both
   * readers rather than clamped: an `r` names one row, so there is nothing to fold it onto, and
   * clamping would silently move its formatting to the last row of the sheet. Its cells go with it,
   * even one whose own `r` is in the grid, because `CellAccumulator` places no cell outside an
   * in-grid row.
   */
  readonly inGrid: boolean;
}

/** Tracks where the reader is in `<sheetData>`, so a `<row>` with no `r` still has a number. */
export class RowPositionTracker {
  #last = 0;

  /** Open the next `<row>`. A declared `r` wins and re-anchors the count, exactly as position does. */
  open(attrs: XmlAttributes): RowPosition {
    const number = numInteger(attrs.r, 1) ?? this.#last + 1;
    this.#last = number;
    return {number, inGrid: number <= MAX_ROW};
  }
}
