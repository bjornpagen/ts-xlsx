// Which columns a `<col min max>` element reaches: the XML half of the column rule, reading the two
// attributes, in front of the budget every reader shares.

import {numInteger} from '../../xml/xml-attrs.ts';
import type {XmlAttributes} from '../../xml/xml-scan.ts';
import {type ColumnRecordBudget, clampColumnSpan} from '../read-policy/column-budget.ts';

/**
 * The columns a `<col min max>` element actually applies to, or `undefined` for one that applies to
 * none: unreadable bounds, or anything {@link clampColumnSpan} refuses. Both XML worksheet readers
 * ask this, so the two cannot read one `<col>` as covering different columns.
 */
export function takeColumnSpan(
  attrs: XmlAttributes,
  budget: ColumnRecordBudget,
): {first: number; last: number} | undefined {
  const first = numInteger(attrs.min, 1);
  const last = numInteger(attrs.max, 1);
  if (first === undefined || last === undefined) return undefined;
  return clampColumnSpan(first, last, budget);
}
