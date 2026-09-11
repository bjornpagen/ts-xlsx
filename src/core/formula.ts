// OOXML froze its formula-function grammar around Excel 2007. Every function Microsoft has added
// since (the dynamic-array family, LAMBDA and its helpers, the newer text and logical functions)
// is persisted in the sheet XML under an `_xlfn.` name-mangling prefix, which a worksheet-only
// function extends to `_xlfn._xlws.`. The prefix is purely an
// on-disk convention: the model only ever holds the plain, readable name, the writer applies the
// prefix on the way out, and the reader strips it back on the way in. A writer that omits it emits
// a formula current Excel silently drops, because the function is unknown under its bare name. This
// module is the single place that knows the mangling, shared by the xlsx writer and reader like
// address.ts and date.ts own their domains.
//
// Moving the references a formula makes, for a shared-formula clone or a splice, is
// `core/formula-references.ts`. Both passes skip the text a reference cannot sit in through
// `core/formula-scan.ts`; `mangleParams` runs its own forward walk, because LET/LAMBDA parameter scope
// opens and closes at paren boundaries, state `scanFormula`'s per-run transform cannot carry, but it
// still defers to `skipOpaque`.

import {assertWritableNumber} from '../errors.ts';
import {nameReadsAsReference} from './address.ts';
import {scanFormula, skipOpaque} from './formula-scan.ts';
import {FUTURE_FUNCTION_PREFIXES} from './future-functions.ts';

const XLPM = '_xlpm.';

/**
 * Quote a sheet name for use in a reference exactly when Excel would: a name that is not a plain
 * identifier, or that would read as a reference, is wrapped in single quotes with its internal
 * quotes doubled, and a simple name is left bare. Shared by everything that *builds* a qualified
 * reference: the `_FilterDatabase` name the writer derives from an autofilter, and the `.xlsb`
 * reader's Ptg decoder, which has only a sheet index to work from and must spell the prefix itself.
 *
 * `last` names the far end of a 3-D span (`Data:More!A1`). A span is quoted as a whole or not at all,
 * because the quotes delimit the sheet *reference* rather than either name, so one awkward endpoint
 * puts both inside the quotes.
 */
export function quoteSheetName(name: string, last?: string): string {
  const names = last === undefined ? [name] : [name, last];
  const joined = names.join(':');
  // A sheet name cannot itself contain a colon, so joining first and quoting the result is
  // unambiguous.
  return names.every(isBareSheetName) ? joined : `'${joined.replace(/'/g, "''")}'`;
}

// A name that reads as a reference is quoted even though every character in it is a name character:
// bare, `R1C1!A1` and `TRUE!A1` are not sheet prefixes. Testing only for an A1 cell left those two,
// `R`, `C` and `R1X` bare, and quoted `XFE1`, which names no column.
function isBareSheetName(name: string): boolean {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !nameReadsAsReference(name);
}

// LET and LAMBDA are the only functions that bind names. Their parameter identifiers are persisted
// under an `_xlpm.` prefix, at the declaration site and at every in-body reference, exactly as the
// modern functions themselves carry `_xlfn.`. The prefix is scoped: a name bound by one LET/LAMBDA
// is only prefixed inside that call, so a same-named defined-name reference elsewhere is untouched.
const SCOPING_FUNCTIONS: ReadonlySet<string> = new Set(['LET', 'LAMBDA']);

// A function call is an identifier (dots included, so a dotted name like NORM.DIST is matched whole
// rather than by its tail) immediately followed by '('. The negative lookbehind rejects a name
// preceded by an identifier character or '.', so an already-qualified name (`_xlfn.XLOOKUP`) is
// consumed as a single token whose uppercased form is absent from the set, and is therefore never
// double-prefixed. Lookbehind rather than a consumed boundary char so adjacent calls
// (SUM(FILTER(…))) both match.
const FUNCTION_CALL = /(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_.]*)(\s*\()/g;
const PREFIX = /_xlfn\.(?:_xlws\.)?|_xlpm\./g;

/**
 * Prefix every future function called by its plain name with the prefix Excel stores it under
 * (`_xlfn.`, or `_xlfn._xlws.` for the worksheet-only FILTER, SORT and PY) so Excel accepts the stored
 * formula. Names already prefixed are left alone (never doubled), unknown/legacy functions pass
 * through untouched, and opaque regions (string literals, sheet names, structured references) are
 * preserved verbatim. No other rewriting occurs: in particular no `@` implicit-intersection operator
 * is ever introduced.
 */
export function mangleFunctions(formula: string): string {
  return scanFormula(formula, (code) =>
    code.replace(FUNCTION_CALL, (whole, name: string, open: string) => {
      const prefix = FUTURE_FUNCTION_PREFIXES.get(name.toUpperCase());
      return prefix === undefined ? whole : `${prefix}${name}${open}`;
    }),
  );
}

/**
 * Strip the `_xlfn.` function prefix (with the `_xlws.` a worksheet-only function adds after it) and
 * the `_xlpm.` LET-parameter prefix back to the plain names, so the model holds the readable form
 * regardless of how a file stored it. Opaque regions (string literals, sheet names, structured
 * references) are left untouched.
 */
export function unmangleFunctions(formula: string): string {
  return scanFormula(formula, (code) => code.replace(PREFIX, ''));
}

const NAME_START = /[A-Za-z_]/;
const NAME_CHAR = /[A-Za-z0-9_.]/;
const WHITESPACE = /\s/;

// Advance past an identifier (dots included, matching FUNCTION_CALL) starting at `i`, or return `i`
// unchanged when no identifier begins there.
function readName(formula: string, i: number): number {
  if (!NAME_START.test(formula[i] ?? '')) return i;
  let j = i + 1;
  while (j < formula.length && NAME_CHAR.test(formula[j] ?? '')) j += 1;
  return j;
}

/**
 * From the index of a call's opening paren, find the matching close and the `[start, end)` ranges of
 * its top-level, comma-separated arguments. Nested parens are tracked by depth; opaque regions (string
 * literals, sheet names, structured references) are skipped whole so their commas do not split an
 * argument.
 */
function parseCall(formula: string, open: number): {close: number; args: [number, number][]} {
  const args: [number, number][] = [];
  const n = formula.length;
  let depth = 0;
  let argStart = open + 1;
  let i = open;
  while (i < n) {
    const past = skipOpaque(formula, i);
    if (past > i) {
      i = past;
      continue;
    }
    const ch = formula[i];
    if (ch === '(') {
      depth += 1;
      i += 1;
    } else if (ch === ')') {
      depth -= 1;
      if (depth === 0) {
        args.push([argStart, i]);
        return {close: i, args};
      }
      i += 1;
    } else if (ch === ',' && depth === 1) {
      args.push([argStart, i]);
      argStart = i + 1;
      i += 1;
    } else {
      i += 1;
    }
  }
  args.push([argStart, n]);
  return {close: n, args};
}

// Extract the single, unprefixed identifier occupying an argument range, or `undefined` when the
// range is not one clean name (whitespace-trimmed): a malformed binding we decline to touch.
function boundName(formula: string, [start, end]: [number, number]): string | undefined {
  let s = start;
  let e = end;
  while (s < e && WHITESPACE.test(formula[s] ?? '')) s += 1;
  while (e > s && WHITESPACE.test(formula[e - 1] ?? '')) e -= 1;
  if (s >= e || readName(formula, s) !== e) return undefined;
  const name = formula.slice(s, e);
  return name.startsWith(XLPM) ? undefined : name;
}

// The parameter names a LET/LAMBDA call binds, keyed uppercased to the spelling each was declared with.
// LAMBDA binds every argument but its last (the body); LET binds the even-indexed arguments up to but
// excluding its last (the calculation).
function parameterNames(
  formula: string,
  keyword: string,
  args: [number, number][],
): ReadonlyMap<string, string> {
  const names = new Map<string, string>();
  const isLambda = keyword === 'LAMBDA';
  for (let a = 0; a < args.length - 1; a += 1) {
    if (isLambda || a % 2 === 0) {
      const name = boundName(formula, args[a] as [number, number]);
      if (name !== undefined && !names.has(name.toUpperCase())) names.set(name.toUpperCase(), name);
    }
  }
  return names;
}

/**
 * Prefix every LET/LAMBDA parameter identifier with `_xlpm.`, at its declaration and at each
 * reference within the binding call's parentheses, so Excel accepts the stored formula. The prefix
 * is lexically scoped: a name is only rewritten inside the call that binds it, opaque regions are
 * copied verbatim, and a lambda-valued parameter used as a call (`f(…)`) is prefixed too. Formulas
 * with no LET/LAMBDA pass through unchanged. A parameter matches case-insensitively, as Excel's names
 * do, and every reference to it is written in the spelling it was declared with, as Excel writes it.
 */
export function mangleParams(formula: string): string {
  // This is the one pass scanFormula cannot serve: it must know when a paren opens a LET/LAMBDA scope
  // and when the matching paren closes it, so it carries frame state across the code between opaque
  // regions rather than transforming each run in isolation. It still defers to skipOpaque, keeping the
  // opaque-skipping rule in one place even though the forward walk here is bespoke.
  let out = '';
  let i = 0;
  const n = formula.length;
  // A stack of active bindings, each expiring exactly at its owner call's close paren. A name is looked
  // up uppercased: matched case-sensitively, `LET(x,1,X+1)` left `X` bare, which Excel then reads as a
  // defined name. The innermost binding is searched first, because a shadowing name decides the
  // spelling its references are written in.
  const frames: {end: number; names: ReadonlyMap<string, string>}[] = [];
  const declared = (name: string): string | undefined => {
    const key = name.toUpperCase();
    for (let f = frames.length - 1; f >= 0; f -= 1) {
      const spelling = frames[f]?.names.get(key);
      if (spelling !== undefined) return spelling;
    }
    return undefined;
  };

  while (i < n) {
    const top = frames[frames.length - 1];
    if (top !== undefined && i >= top.end) {
      frames.pop();
      continue;
    }

    const past = skipOpaque(formula, i);
    if (past > i) {
      out += formula.slice(i, past);
      i = past;
      continue;
    }

    const ch = formula[i] as string;
    const nameEnd = readName(formula, i);
    if (nameEnd === i) {
      out += ch;
      i += 1;
      continue;
    }

    const name = formula.slice(i, nameEnd);
    let k = nameEnd;
    while (k < n && WHITESPACE.test(formula[k] ?? '')) k += 1;
    const heads = formula[k] === '(';

    if (heads && SCOPING_FUNCTIONS.has(name.toUpperCase()) && declared(name) === undefined) {
      const {close, args} = parseCall(formula, k);
      // The keyword stays at the outer scope; its parameters take effect inside the parens.
      out += formula.slice(i, k + 1);
      frames.push({end: close, names: parameterNames(formula, name.toUpperCase(), args)});
      i = k + 1;
      continue;
    }

    // Any other identifier: a bare reference, an ordinary call, or a lambda-valued parameter call.
    // In-scope names (declaration sites included, as they lie inside their own binding's parens) take
    // the prefix and their declared spelling; the rest pass through. Call arguments are covered by the
    // continuing scan, so a nested LET/LAMBDA within them is still seen.
    const spelling = declared(name);
    out += spelling === undefined ? name : `${XLPM}${spelling}`;
    i = nameEnd;
  }
  return out;
}

/**
 * Mangle a model formula into its on-disk form: LET/LAMBDA parameter names first (`_xlpm.`), then the
 * modern-function prefix (`_xlfn.`). Ordering matters: parameter mangling reads the plain LET/LAMBDA
 * names before the function pass qualifies them. The inverse for both prefixes is unmangleFunctions.
 */
export function mangleFormula(formula: string): string {
  return mangleFunctions(mangleParams(formula));
}

/**
 * A number as it appears *inside formula text*, which is not the same serialisation as an
 * attribute's.
 *
 * Two divergences, both found as drift rather than designed. The exponent's case: Excel writes
 * `1E+21` where JavaScript writes `1e+21`, so the same literal read back through the two codecs
 * produced two different formula strings depending on which file it came from -- exactly the
 * asymmetry `cell-value.ts` and `cell-accumulator.ts` were extracted to prevent. And the finiteness
 * guard: the BIFF12 decoder had a private copy of this function with none, so a `PtgNum` whose eight
 * bytes decode to an infinity produced the formula text `INFINITY`, which the writer then escaped as
 * ordinary text into a package Excel reports as damaged.
 *
 * Here rather than in `xml/xml.ts` beside the attribute form, because formula text is not XML: the
 * BIFF12 codec produces it too, and reaching the XML serialiser for it put the whole write half of
 * that module into the closure of an entry that has no XML in it.
 *
 * @throws {AuthoringError} when the value is not finite.
 */
export function formulaNumberLiteral(value: number): string {
  assertWritableNumber(value);
  // Only `e` can appear in a JavaScript number's decimal form besides digits and a sign, so an
  // unconditional fold cannot touch anything else.
  return String(value).toUpperCase();
}

/**
 * Render a formula operand for serialisation: a number becomes its literal, a string is stripped of
 * the single optional leading '=' an author may write (OOXML stores the expression without it, e.g.
 * `=A1>0` on disk is `A1>0`). The result is unescaped: the caller escapes it for its target, whether
 * that is element text or an attribute value.
 */
export function stripFormulaEquals(value: string | number): string {
  if (typeof value === 'number') return formulaNumberLiteral(value);
  return value.startsWith('=') ? value.slice(1) : value;
}
