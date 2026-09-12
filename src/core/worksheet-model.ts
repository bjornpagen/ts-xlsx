// The field table behind `Worksheet.model`'s round-trip.
//
// Both directions used to be hand-written statement lists, one in the getter and one in the setter,
// each enumerating the same fields. Nothing but a comment asked the next editor to touch
// both, and a field exported but not imported loses data silently: the merge-loss failure the
// model contract exists to prevent. Here each field declares both directions in one place, and the
// registry is proved exhaustive over `keyof WorksheetModel` at compile time, so adding a field
// without wiring it is an error naming the field rather than a review catch.

import {cellToModel, copyCellContent} from './cell.ts';
import {cloneConditionalFormatting} from './conditional-formatting.ts';
import {overwrite, replaceContents} from './containers.ts';
import {cloneDataValidation} from './data-validation.ts';
import {type AssertNever, INTERNAL} from './internal.ts';
import type {PageBreak} from './page-setup.ts';
import type {CellModel, Worksheet, WorksheetModel} from './worksheet.ts';

/** One field of a {@link WorksheetModel}, with both directions of its round-trip declared together. */
interface ModelFacet<K extends keyof WorksheetModel = keyof WorksheetModel> {
  readonly key: K;
  /**
   * Produce the field's value, in a form a caller cannot mutate the sheet through.
   *
   * Copying is one way to get there and the type is the other. `tabColor`, `autoFilter` and
   * `protection` are handed back by reference precisely because their types are readonly all the way
   * down, so there is nothing to defend against and a clone would only be one more shape to keep in
   * step with its declaration. Every other field is a mutable record or array and is copied, which is
   * where the spreads below come from.
   */
  readonly read: (sheet: Worksheet) => WorksheetModel[K];
  /**
   * Apply the field to a sheet whose content has already been reset. Takes the whole model rather
   * than the field so that a loop over the registry needs no correlation between `key` and the
   * field's type. {@link facet} does that projection once, where the key is still a single type.
   *
   * The obvious shape, `write(sheet, value: WorksheetModel[K])` stored as-is, cannot be called from
   * a loop: over a union of `ModelFacet<K>` the parameter is contravariant under
   * `strictFunctionTypes`, which breaks the correlation. Declaring `write` with method syntax makes
   * it compile, by making the position bivariant, which buys the call back by switching the check
   * off. Projecting inside the helper is correlated *and* sound; do not "simplify" it back.
   */
  readonly write: (sheet: Worksheet, model: WorksheetModel) => void;
}

function facet<K extends keyof WorksheetModel>(
  key: K,
  read: (sheet: Worksheet) => WorksheetModel[K],
  write: (sheet: Worksheet, value: WorksheetModel[K]) => void,
): ModelFacet<K> {
  return {key, read, write: (sheet, model) => write(sheet, model[key])};
}

/**
 * A field that is a flat record of optional properties, held live on the sheet under the same name.
 *
 * Seven fields share this shape, and written out per field it was five lines each, of which the
 * spread on the read side is the smallest part and the easiest to leave out. A field read without it
 * hands back the sheet's own object, and `model.pageSetup.orientation = 'landscape'` then edits the
 * sheet through what {@link ModelFacet.read} documents as a copy.
 */
function recordFacet<K extends RecordField>(key: K): ModelFacet<K> {
  return {
    key,
    read: (sheet) => ({...sheet[key]}),
    write: (sheet, model) => overwrite<WorksheetModel[K]>(sheet[key], model[key]),
  };
}

type RecordField =
  | 'properties'
  | 'outline'
  | 'view'
  | 'pageSetup'
  | 'printOptions'
  | 'pageMargins'
  | 'headerFooter';

/** {@link recordFacet} for a field that is a *list* of flat records, replaced in place. */
function recordsFacet<K extends 'rowBreaks' | 'columnBreaks'>(key: K): ModelFacet<K> {
  const copy = (breaks: readonly PageBreak[]): PageBreak[] => breaks.map((brk) => ({...brk}));
  return {
    key,
    read: (sheet) => copy(sheet[key]),
    write: (sheet, model) => replaceContents(sheet[key], copy(model[key])),
  };
}

/**
 * Every field of a {@link WorksheetModel}, in the order a model assignment applies them. Order is
 * load-bearing: cells are placed at their exact positions before any merge exists, so a covered
 * cell's value lands where the model says instead of being routed to a region master mid-load.
 */
export const WORKSHEET_MODEL_FACETS = [
  facet(
    'state',
    (sheet) => sheet.state,
    (sheet, value) => {
      sheet.state = value;
    },
  ),
  facet(
    'tabColor',
    (sheet) => sheet.tabColor,
    (sheet, value) => {
      sheet.tabColor = value;
    },
  ),
  recordFacet('properties'),
  recordFacet('outline'),
  recordFacet('view'),
  recordFacet('pageSetup'),
  recordFacet('printOptions'),
  recordFacet('pageMargins'),
  recordFacet('headerFooter'),
  recordsFacet('rowBreaks'),
  recordsFacet('columnBreaks'),
  facet(
    'columns',
    (sheet) =>
      [...sheet.columns()].map(({index, properties}) => ({index, properties: {...properties}})),
    (sheet, value) => {
      for (const {index, properties} of value) Object.assign(sheet.getColumn(index), properties);
    },
  ),
  facet(
    'rows',
    (sheet) => {
      const rows: WorksheetModel['rows'] = [];
      // A row appears here only for its formatting; the cells it holds are the `cells` field's
      // business, and a row carrying nothing but cells has no properties to round-trip.
      for (const {number, properties} of sheet.rows()) {
        if (properties !== undefined) rows.push({number, properties: {...properties}});
      }
      return rows;
    },
    (sheet, value) => {
      for (const {number, properties} of value) Object.assign(sheet.getRow(number), properties);
    },
  ),
  facet(
    'cells',
    (sheet) => {
      const cells: CellModel[] = [];
      for (const row of sheet.rows()) for (const cell of row.cells) cells.push(cellToModel(cell));
      return cells;
    },
    (sheet, value) => {
      for (const cell of value) copyCellContent(cell, sheet[INTERNAL].cellAt(cell.row, cell.col));
    },
  ),
  facet(
    'merges',
    (sheet) => [...sheet.merges],
    (sheet, value) => {
      for (const range of value) sheet.mergeCells(range);
    },
  ),
  facet(
    'hyperlinks',
    // A link's fields are strings, so a spread is a whole copy.
    (sheet) => sheet.hyperlinks.map((link) => ({...link})),
    (sheet, value) => {
      for (const link of value) sheet.addHyperlink(link);
    },
  ),
  facet(
    'dataValidations',
    (sheet) =>
      sheet.dataValidations.map(({sqref, rule, extended}) => ({
        sqref,
        rule: cloneDataValidation(rule),
        ...(extended ? {extended: true} : {}),
      })),
    (sheet, value) => {
      for (const {sqref, rule, extended} of value) {
        sheet.addDataValidation(sqref, rule, extended ? {extended: true} : {});
      }
    },
  ),
  facet(
    'conditionalFormattings',
    (sheet) => sheet.conditionalFormattings.map(cloneConditionalFormatting),
    (sheet, value) => {
      for (const formatting of value) sheet.addConditionalFormatting(formatting);
    },
  ),
  facet(
    'tables',
    (sheet) => sheet.tables.map((table) => table.options),
    (sheet, value) => {
      for (const options of value) sheet.addTable(options);
    },
  ),
  facet(
    'autoFilter',
    (sheet) => sheet.autoFilter,
    // Through the public setter, which re-canonicalises the range and, on `undefined`, clears any
    // autofilter the destination held. That clearing is why the field is applied even when absent.
    (sheet, value) => {
      sheet.autoFilter = value;
    },
  ),
  facet(
    'protection',
    (sheet) => sheet.protection,
    // A loaded credential is already hashed, so it is reinstated verbatim rather than re-derived;
    // `protect` cannot express that, which is why this goes through the internal channel.
    (sheet, value) => {
      if (value === undefined) sheet.unprotect();
      else sheet[INTERNAL].restoreProtection(value);
    },
  ),
];

/**
 * Compile-time proof that {@link WORKSHEET_MODEL_FACETS} covers every {@link WorksheetModel} field.
 * A field added without a facet resolves this to that field's name, which does not satisfy `never`,
 * so the error names what is missing. This is the guarantee the registry exists to provide: the two
 * directions of the round-trip can no longer drift apart without the build saying so.
 */
export type EveryWorksheetModelFieldHasAFacet = AssertNever<
  Exclude<keyof WorksheetModel, (typeof WORKSHEET_MODEL_FACETS)[number]['key']>
>;
