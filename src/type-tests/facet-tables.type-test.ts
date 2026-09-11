// Type-level tests over the internal facet tables: the declarations that exist so a reader and a
// writer cannot fall out of step over the same set of fields.
//
// Each table pairs with an `AssertNever<Exclude<keyof T, table keys>>` alias beside it, which fails
// to resolve the moment the type gains a field the table does not name. That alias is already a
// compile error where it is declared; asserting it here is what states, in one place, which tables
// carry the guarantee, so a table added without one is visible as an absence rather than assumed.
//
// Which is worth exactly as much as the list is complete, and the list had drifted to six of eleven:
// `EveryColumnPropertyIsMirrored` was asserted and its literal twin `EveryRowPropertyIsMirrored` was
// not, and all four `page-setup.ts` proofs were missing. Nothing was unproven, because the alias
// fires where it is declared. What was wrong was the register's own claim: read it and you concluded
// `RowProperties` and `PageSetup` carry no exhaustiveness proof. `scripts/check-facet-register.ts`
// now holds the two in step, so the next proof added is either listed here or a failed gate.

import type {EveryColumnPropertyIsMirrored} from '../core/column.ts';
import type {
  EveryDifferentialStyleFieldIsCloned,
  EveryRuleFieldIsCloned,
} from '../core/conditional-formatting.ts';
import type {EveryDataValidationFieldIsCloned} from '../core/data-validation.ts';
import type {
  EveryHeaderFooterElementIsDeclared,
  EveryMarginSideIsDeclared,
  EveryPageSetupFacetIsDeclared,
  EveryPrintOptionFlagIsDeclared,
} from '../core/page-setup.ts';
import type {EverySheetProtectionFlagIsDeclared} from '../core/protection.ts';
import type {EveryRowPropertyIsMirrored} from '../core/row.ts';
import type {
  EveryAlignmentFacetIsDeclared,
  EveryBorderEdgeFieldIsCloned,
  EveryFillPatternIsOrdered,
  EveryBorderFieldIsCloned,
  EveryFontFieldIsCloned,
  EveryGradientFillFieldIsCloned,
  EveryGradientStopFieldIsCloned,
  EveryPatternFillFieldIsCloned,
} from '../core/style.ts';
import type {
  EveryTableStyleFlagIsDeclared,
  EveryTableStyleInfoFieldIsCloned,
} from '../core/table.ts';
import type {EveryWorksheetModelFieldHasAFacet} from '../core/worksheet-model.ts';
import type {EveryProtectionFlagHasARecordSlot} from '../io/xlsb/sheet-protection.ts';
import type {EverySlotIsInSchemaOrder} from '../io/xlsx/theme-xml.ts';
import type {Equal, Expect} from './expect.ts';

export type FacetTableExhaustiveness = [
  // `<alignment>`'s seven facets, read and written from one declaration.
  Expect<Equal<EveryAlignmentFacetIsDeclared, never>>,
  // The two axis handles' properties, each mirrored between the handle and the stored record. Both,
  // because `axis-handle.ts` names them as a pair and only one of them used to be asserted here.
  Expect<Equal<EveryColumnPropertyIsMirrored, never>>,
  Expect<Equal<EveryRowPropertyIsMirrored, never>>,
  // Every worksheet model field, covered by a getter/setter facet.
  Expect<Equal<EveryWorksheetModelFieldHasAFacet, never>>,
  // The print-layout tables, all four driven from `page-setup.ts` by both the reader and the writer.
  Expect<Equal<EveryPageSetupFacetIsDeclared, never>>,
  Expect<Equal<EveryPrintOptionFlagIsDeclared, never>>,
  Expect<Equal<EveryMarginSideIsDeclared, never>>,
  Expect<Equal<EveryHeaderFooterElementIsDeclared, never>>,
  // The sheet-protection flags, turned into `<sheetProtection>` attributes and back from one table.
  Expect<Equal<EverySheetProtectionFlagIsDeclared, never>>,
  // A table's banding flags, turned into `<tableStyleInfo>` attributes and back from one table.
  Expect<Equal<EveryTableStyleFlagIsDeclared, never>>,
  // A list, not a record, so it owes the other half of its proof explicitly: BIFF12 indexes the fill
  // patterns, so the enumeration is needed in order and an ordered list can omit a member silently.
  Expect<Equal<EveryFillPatternIsOrdered, never>>,
  // The other ordered list, and the one whose order is not the one the model indexes by: a theme
  // colour slot missing from the schema order is dropped from every rewritten `<a:clrScheme>`.
  Expect<Equal<EverySlotIsInSchemaOrder, never>>,
  // The same flags as BIFF12 stores them, by position: one left out shifts every flag after it.
  Expect<Equal<EveryProtectionFlagHasARecordSlot, never>>,
  // The defensive copies: a field added to one of these types without a copy strategy would
  // otherwise be carried by reference into the model, aliasing the caller's object silently.
  Expect<Equal<EveryRuleFieldIsCloned, never>>,
  Expect<Equal<EveryDifferentialStyleFieldIsCloned, never>>,
  Expect<Equal<EveryDataValidationFieldIsCloned, never>>,
  Expect<Equal<EveryTableStyleInfoFieldIsCloned, never>>,
  // The style primitives a differential style is made of. These are the level the copy used to stop
  // one short of: `'record'` on a font, a border or a fill is a spread, and each of the three nests
  // something a spread carries by reference.
  Expect<Equal<EveryFontFieldIsCloned, never>>,
  Expect<Equal<EveryBorderFieldIsCloned, never>>,
  Expect<Equal<EveryBorderEdgeFieldIsCloned, never>>,
  Expect<Equal<EveryPatternFillFieldIsCloned, never>>,
  Expect<Equal<EveryGradientFillFieldIsCloned, never>>,
  Expect<Equal<EveryGradientStopFieldIsCloned, never>>,
];
