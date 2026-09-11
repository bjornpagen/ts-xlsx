// Data validation I/O: the sheet-level `<dataValidations>` element and the reader that folds each
// rule back onto its range.
//
// A validation stores its operands in `<formula1>`/`<formula2>` child elements and its target in a
// `sqref` attribute. The operand text carries NO leading '=' (the '=' is an authoring convention, not
// part of the stored formula), so the writer strips one and the reader keeps whatever it finds. A
// numeric-typed rule's literal operand parses to a number; a cell reference, defined name, or list
// source keeps its string, so a reference is never coerced to NaN and lost.
//
// The extended `<x14:dataValidation>` form (2009 extension schema) carries the validations a legacy
// element cannot express, chiefly a list whose source lives on another sheet. It lives in the
// worksheet `<extLst>`, keeps its target in a `<xm:sqref>` child rather than a `sqref` attribute, and
// wraps each operand in an `<xm:f>` under `<x14:formula1>`/`<x14:formula2>`. A rule read from that
// form is tagged `extended` so it is written back there; the two forms are parsed and serialised by
// prefix so neither reader mistakes one for the other.

import {
  type DataValidation,
  type DataValidationEntry,
  isDataValidationErrorStyle,
  isDataValidationImeMode,
  isDataValidationOperator,
  isDataValidationType,
} from '../../core/data-validation.ts';
import {mangleFormula, stripFormulaEquals, unmangleFunctions} from '../../core/formula.ts';
import {decodeSqrefRects} from '../../core/merge.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {coerceNumericLiteral, enumToken} from '../../xml/xml-attrs.ts';
import {type CollectingPass, type SaxHandlers, TextCapture} from '../../xml/xml-read.ts';
import {boolStrict, localName} from '../../xml/xml-scan.ts';
import {checkedToken, escapeAttr, escapeText, textAttr} from '../../xml/xml.ts';
import {admitting} from '../read-policy/read-repair.ts';
// The x14/xm extension namespaces and `DATA_VALIDATION_EXT_URI` are declared inline on the elements
// that need them, exactly as Excel writes them, so the block is self-contained and the worksheet root
// needs no extra namespace declaration.
import {
  DATA_VALIDATION_EXT_URI,
  isExtensionElement,
  isMainNamespaceElement,
  XM_NS,
} from './namespaces.ts';
import {x14Ext} from './x14-ext.ts';

// The typed validations whose literal operands are numbers; `list`/`custom` operands stay strings.
const TYPED = new Set<string>(['whole', 'decimal', 'date', 'time', 'textLength']);

/** The standard `<dataValidations>` element for the rules stored in the legacy form, or '' when the
 * sheet has none of them, so a sheet with only extended (or no) validations stays byte-clean here.
 * The extended rules are emitted separately by {@link dataValidationsExtXml}.
 *
 * An operand is stored with the function prefixes a cell formula takes, as Excel stores it;
 * `formulaNames` is `formulaNamesInScope` for the sheet. */
export function dataValidationsXml(
  entries: readonly DataValidationEntry[],
  formulaNames: ReadonlySet<string>,
): string {
  const standard = entries.filter((entry) => !entry.extended);
  if (standard.length === 0) return '';
  const items = standard
    .map(({sqref, rule}) => dataValidationXml(sqref, rule, formulaNames))
    .join('');
  return `<dataValidations count="${standard.length}">${items}</dataValidations>`;
}

/** The `<ext>` carrying the extended (`<x14:dataValidation>`) rules, or '' when the sheet declares
 * none. Emitted bare (no `<extLst>` wrapper) so the worksheet serialiser can gather it into a single
 * `<extLst>` beside the conditional-formatting extension: a worksheet may carry at most one. */
export function dataValidationsExtXml(
  entries: readonly DataValidationEntry[],
  formulaNames: ReadonlySet<string>,
): string {
  const extended = entries.filter((entry) => entry.extended);
  if (extended.length === 0) return '';
  const items = extended
    .map(({sqref, rule}) => extendedDataValidationXml(sqref, rule, formulaNames))
    .join('');
  return x14Ext(
    DATA_VALIDATION_EXT_URI,
    `<x14:dataValidations count="${extended.length}" xmlns:xm="${XM_NS}">${items}</x14:dataValidations>`,
  );
}

// The shared attributes of a validation, in CT_DataValidation order: type, errorStyle, imeMode,
// operator, allowBlank, showDropDown, showInputMessage, showErrorMessage, errorTitle, error,
// promptTitle, prompt. The target range differs between the two forms (a `sqref` attribute vs an
// `<xm:sqref>` child), so it is not part of this shared prefix.
function ruleAttrs(rule: DataValidation): string {
  return (
    ` type="${checkedToken(rule.type, isDataValidationType, 'data validation type')}"` +
    (rule.errorStyle === undefined
      ? ''
      : ` errorStyle="${checkedToken(rule.errorStyle, isDataValidationErrorStyle, 'data validation error style')}"`) +
    (rule.imeMode === undefined
      ? ''
      : ` imeMode="${checkedToken(rule.imeMode, isDataValidationImeMode, 'data validation IME mode')}"`) +
    (rule.operator === undefined
      ? ''
      : ` operator="${checkedToken(rule.operator, isDataValidationOperator, 'data validation operator')}"`) +
    (rule.allowBlank ? ' allowBlank="1"' : '') +
    (rule.suppressDropDown ? ' showDropDown="1"' : '') +
    (rule.showInputMessage ? ' showInputMessage="1"' : '') +
    (rule.showErrorMessage ? ' showErrorMessage="1"' : '') +
    textAttr('errorTitle', rule.errorTitle) +
    textAttr('error', rule.error) +
    textAttr('promptTitle', rule.promptTitle) +
    textAttr('prompt', rule.prompt)
  );
}

// The standard element: shared attributes, then `sqref` last, then `<formula1>`/`<formula2>` bodies.
//
// A non-finite numeric operand (a date bound that failed to coerce to a serial, say) is refused by
// `stripFormulaEquals` in both forms, as a conditional format's operand is: a number the format cannot
// spell is refused where it becomes bytes. It used to be dropped, which quietly turned a `between`
// into a rule with one bound.
function dataValidationXml(
  sqref: string,
  rule: DataValidation,
  formulaNames: ReadonlySet<string>,
): string {
  const [f1, f2] = rule.formulae ?? [];
  const body =
    (f1 !== undefined ? `<formula1>${operandText(f1, formulaNames)}</formula1>` : '') +
    (f2 !== undefined ? `<formula2>${operandText(f2, formulaNames)}</formula2>` : '');
  return `<dataValidation${ruleAttrs(rule)} sqref="${escapeAttr(sqref)}">${body}</dataValidation>`;
}

// The extended element: same shared attributes, but each operand wraps in `<x14:formula1><xm:f>…` and
// the target range is an `<xm:sqref>` child that follows the formulae. The `xr:uid` Excel adds is
// revision metadata it regenerates freely, so it is not modelled or re-emitted.
function extendedDataValidationXml(
  sqref: string,
  rule: DataValidation,
  formulaNames: ReadonlySet<string>,
): string {
  const [f1, f2] = rule.formulae ?? [];
  const body =
    (f1 !== undefined
      ? `<x14:formula1><xm:f>${operandText(f1, formulaNames)}</xm:f></x14:formula1>`
      : '') +
    (f2 !== undefined
      ? `<x14:formula2><xm:f>${operandText(f2, formulaNames)}</xm:f></x14:formula2>`
      : '') +
    `<xm:sqref>${escapeText(sqref)}</xm:sqref>`;
  return `<x14:dataValidation${ruleAttrs(rule)}>${body}</x14:dataValidation>`;
}

// One operand as element text, the same in both forms: without the author's `=`, with the function
// prefixes Excel stores, escaped.
function operandText(operand: string | number, formulaNames: ReadonlySet<string>): string {
  return escapeText(mangleFormula(stripFormulaEquals(operand), formulaNames));
}

// The two operand elements a validation carries, either of which may be absent.
const FORMULA_ELEMENTS: ReadonlySet<string> = new Set(['formula1', 'formula2']);

/** A pass gathering the standard `<dataValidation>` elements of a worksheet part, for a caller
 * reading the part alongside its other readers in one parse. An operand sheds its function prefixes
 * against `definedNames`, the workbook's names as `definedNameKeys` spells them. */
export function dataValidationPass(
  definedNames: ReadonlySet<string>,
): CollectingPass<DataValidationEntry[]> {
  const entries: DataValidationEntry[] = [];
  let current: {attrs: Record<string, string>; formulae: string[]} | undefined;
  // Through the shared machine, like its extended sibling thirty lines below already was. The
  // hand-rolled latch stayed set on a self-closing `<formula1/>`, so the text of whatever element came
  // next was appended to the formula it should have ended.
  const formula = new TextCapture(FORMULA_ELEMENTS);

  const handlers: SaxHandlers = {
    onOpen(name, attrs, selfClosing, scope) {
      const ln = localName(name);
      // Only the one in the MAIN namespace: an `x14:dataValidation` is left for the extended path.
      // Tested by namespace, not by whether the name has a prefix: a worksheet that binds the main
      // namespace to a prefix, which is legal and which real toolchains emit, has a colon in every
      // element name, so the prefix test discarded every standard validation in such a file.
      if (ln === 'dataValidation' && isMainNamespaceElement(scope, name)) {
        current = {attrs, formulae: []};
      } else if (current !== undefined) {
        formula.open(ln, selfClosing);
      }
    },
    onText(text) {
      formula.text(text);
    },
    onClose(name) {
      const ln = localName(name);
      const text = formula.close(ln);
      if (text !== undefined) {
        if (current !== undefined) current.formulae[ln === 'formula1' ? 0 : 1] = text;
        return;
      }
      if (ln === 'dataValidation' && current !== undefined) {
        const built = buildEntry(current.attrs, current.formulae, definedNames);
        if (built !== undefined) entries.push(built);
        current = undefined;
      }
    },
  };
  return {handlers, result: () => entries};
}

function buildEntry(
  attrs: Record<string, string>,
  formulae: readonly string[],
  definedNames: ReadonlySet<string>,
): DataValidationEntry | undefined {
  const {sqref} = attrs;
  if (sqref === undefined) return undefined;
  const rule = buildRule(attrs, formulae, definedNames);
  return rule === undefined ? undefined : {sqref, rule};
}

// The rule carried by a validation element of either form: its attributes decide the type, operator,
// flags, and messages; its `<formula1>`/`<formula2>` operands become `formulae`. The target range is
// supplied separately by each form's caller, so it is not read here.
function buildRule(
  attrs: Record<string, string>,
  formulae: readonly string[],
  definedNames: ReadonlySet<string>,
): DataValidation | undefined {
  const {type} = attrs;
  // A type outside `ST_DataValidationType` does not name a constraint this model can hold, and the
  // type is the rule's identity: there is nothing left to keep, so the whole entry is dropped.
  if (type === undefined || !isDataValidationType(type)) return undefined;

  const rule: DataValidation = {type};
  if (attrs.operator !== undefined) {
    // An operator, unlike the type, is a facet of a rule that stands without it. An unrecognised one
    // is dropped and the rule kept, the same "store only what the source carried" rule the boolean
    // and integer readers follow.
    if (isDataValidationOperator(attrs.operator)) rule.operator = attrs.operator;
  } else if (TYPED.has(type)) {
    // Excel omits `operator="between"` because it is the default for a typed rule; restore it so a
    // reader sees the operator the rule actually enforces.
    rule.operator = 'between';
  }
  if (boolStrict(attrs.allowBlank)) rule.allowBlank = true;
  if (boolStrict(attrs.showDropDown)) rule.suppressDropDown = true;
  if (boolStrict(attrs.showInputMessage)) rule.showInputMessage = true;
  if (boolStrict(attrs.showErrorMessage)) rule.showErrorMessage = true;
  // Likewise a facet: an unrecognised errorStyle costs the rule its alert level, not its identity.
  if (attrs.errorStyle !== undefined && isDataValidationErrorStyle(attrs.errorStyle))
    rule.errorStyle = attrs.errorStyle;
  // And the input mode: a foreign one costs the rule only how the IME behaves in its cells.
  const imeMode = enumToken(attrs.imeMode, isDataValidationImeMode);
  if (imeMode !== undefined) rule.imeMode = imeMode;
  if (attrs.error !== undefined) rule.error = attrs.error;
  if (attrs.errorTitle !== undefined) rule.errorTitle = attrs.errorTitle;
  if (attrs.prompt !== undefined) rule.prompt = attrs.prompt;
  if (attrs.promptTitle !== undefined) rule.promptTitle = attrs.promptTitle;

  // A `list`/`custom` operand is always a string (a source list or an expression); every other
  // type's literal operand coerces to a number when it is one, so a numeric bound reads back as a
  // number while a cell reference or defined name survives as its verbatim string.
  const parsed = formulae
    .filter((f): f is string => f !== undefined)
    .map((f) => unmangleFunctions(f, definedNames))
    .map((f) => (type === 'list' || type === 'custom' ? f : coerceNumericLiteral(f)));
  if (parsed.length > 0) rule.formulae = parsed;

  return rule;
}

/** A pass gathering the extended `<x14:dataValidation>` elements of a worksheet's `<extLst>`, so a
 * cross-sheet or whole-column list validation Excel stored only in the 2009 extension form is read
 * back rather than dropped. The standard pass ignores these (they are prefixed); this one,
 * symmetrically, handles only the prefixed elements. An operand sheds its function prefixes against
 * `definedNames`, as the standard pass's does. */
export function extendedDataValidationPass(
  definedNames: ReadonlySet<string>,
): CollectingPass<DataValidationEntry[]> {
  const entries: DataValidationEntry[] = [];
  let current: {attrs: Record<string, string>; formulae: string[]; sqref: string} | undefined;
  // Which operand an `<xm:f>` feeds, set by the enclosing `<x14:formula1>`/`<x14:formula2>`.
  let slot: number | undefined;
  const capture = new TextCapture(['f', 'sqref']);

  const handlers: SaxHandlers = {
    onOpen(name, attrs, selfClosing, scope) {
      const ln = localName(name);
      // In the x14 extension namespace, rather than merely carrying a prefix: the prefix test was true
      // of every element in a worksheet that prefixes the main namespace, so such a file had its
      // standard validations read as extended ones and dropped by both passes.
      const extension = isExtensionElement(scope, name);
      // A `<x14:dataValidation>`; its attributes (type, flags, messages) build the rule.
      if (ln === 'dataValidation' && extension) {
        current = {attrs, formulae: [], sqref: ''};
      } else if (current !== undefined && extension && ln === 'formula1') {
        slot = 0;
      } else if (current !== undefined && extension && ln === 'formula2') {
        slot = 1;
      } else if (current !== undefined) {
        capture.open(ln, selfClosing);
      }
    },
    onText(chunk) {
      capture.text(chunk);
    },
    onClose(name, scope) {
      const ln = localName(name);
      const extension = isExtensionElement(scope, name);
      const text = capture.close(ln);
      if (text !== undefined) {
        if (ln === 'f') {
          if (current !== undefined && slot !== undefined) current.formulae[slot] = text;
        } else if (current !== undefined) {
          current.sqref = text;
        }
      } else if ((ln === 'formula1' || ln === 'formula2') && extension) {
        slot = undefined;
      } else if (ln === 'dataValidation' && extension && current !== undefined) {
        const built = buildExtendedEntry(
          current.attrs,
          current.formulae,
          current.sqref,
          definedNames,
        );
        if (built !== undefined) entries.push(built);
        current = undefined;
      }
    },
  };
  return {handlers, result: () => entries};
}

function buildExtendedEntry(
  attrs: Record<string, string>,
  formulae: readonly string[],
  sqref: string,
  definedNames: ReadonlySet<string>,
): DataValidationEntry | undefined {
  if (sqref === '') return undefined;
  const rule = buildRule(attrs, formulae, definedNames);
  return rule === undefined ? undefined : {sqref, rule, extended: true};
}

/** Fold parsed validations onto a sheet, each bound to its original range and carrying its form: an
 * `extended` entry is re-attached as extended so a round-trip writes it back to the x14 block. */
export function applyDataValidations(
  sheet: Worksheet,
  entries: readonly DataValidationEntry[],
): void {
  for (const {sqref, rule, extended} of entries) {
    // A `sqref` no area of which decodes names no cells to validate, and re-emitting it would put the
    // file's own unreadable text back on the wire. Dropped here, at the reader's boundary, so the
    // authoring guard behind `addDataValidation` stays a guard rather than a control-flow path.
    if (decodeSqrefRects(sqref).length === 0) continue;
    admitting(() => {
      sheet.addDataValidation(sqref, rule, extended ? {extended: true} : {});
    });
  }
}
