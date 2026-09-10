// OOXML froze its formula-function grammar around Excel 2007. Every function Microsoft has added
// since (the dynamic-array family, LAMBDA and its helpers, the newer text and logical functions)
// is persisted in the sheet XML under an `_xlfn.` name-mangling prefix. The prefix is purely an
// on-disk convention: the model only ever holds the plain, readable name, the writer applies the
// prefix on the way out, and the reader strips it back on the way in. A writer that omits it emits
// a formula current Excel silently drops, because the function is unknown under its bare name. This
// module is the single place that knows the mangling, shared by the xlsx writer and reader like
// address.ts and date.ts own their domains.
//
// It also owns formula *translation*: a spreadsheet fills a formula down or across a range by storing
// it once on a master cell and marking the rest as shared clones. Reading a clone means recovering the
// master's formula shifted to the clone's position: relative references move by the row/column
// offset, absolute (`$`-anchored) parts stay put. That relative-reference arithmetic lives here too.
//
// Every pass over a formula shares one hazard: a comma, paren, function name, or cell reference is
// mere text when it sits inside a string literal, a single-quoted sheet name, or a bracketed
// structured reference. `skipOpaque` is the single owner of skipping those regions: every pass drives
// its string/quote/bracket handling through it, so the rule lives in one place. The stateless passes
// (function-name and cell-reference rewriting) ride `scanFormula`, which copies the opaque regions
// verbatim and hands each code run between them to a transform. `mangleParams` is the deliberate
// exception: LET/LAMBDA parameter scope opens and closes at paren boundaries, state `scanFormula`'s
// per-run transform cannot carry, so it runs its own forward walk, still deferring to `skipOpaque`.

import {assertWritableNumber} from '../errors.ts';
import {MAX_COLUMN, MAX_ROW, numberToColumn, tryColumnToNumber} from './address.ts';
import {MODERN_FUNCTIONS} from './modern-functions.ts';
import {REF_ERROR} from './value.ts';

const XLFN = '_xlfn.';
const XLPM = '_xlpm.';

/**
 * Quote a sheet name for use in a reference exactly when Excel would: a name that is not a plain
 * identifier, or that would read as a cell address, is wrapped in single quotes with its internal
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

function isBareSheetName(name: string): boolean {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name);
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
const PREFIX = /_xlfn\.|_xlpm\./g;

// Advance past the opaque region opened at `index`: a double-quoted string literal or a single-quoted
// sheet name, both honouring the doubled-quote escape (`""`, `''`), or a bracketed structured
// reference, which may nest (`Table[[#Data],[Col]]`) and escapes a bracket inside a column name with a
// leading `'` (`T[[a'[b]]`). Returns the index just past the region, or
// `index` unchanged when no opaque region opens there. Inside any of the three a comma, paren, function
// name, or cell reference is inert, so every pass over a formula skips them through this one function.
function skipOpaque(formula: string, index: number): number {
  const opener = formula[index];
  const n = formula.length;
  if (opener === '"' || opener === "'") {
    let j = index + 1;
    while (j < n) {
      if (formula[j] === opener) {
        if (formula[j + 1] === opener) {
          j += 2;
          continue;
        }
        return j + 1;
      }
      j += 1;
    }
    return n;
  }
  if (opener === '[') {
    let depth = 0;
    let j = index;
    while (j < n) {
      const ch = formula[j];
      // An escaped character is part of the name, whatever it is: counted as a bracket, an escaped
      // `[` left the region open to the end of the formula and an escaped `]` could close it early.
      if (ch === "'") {
        j += 2;
        continue;
      }
      if (ch === '[') depth += 1;
      else if (ch === ']') {
        depth -= 1;
        if (depth === 0) return j + 1;
      }
      j += 1;
    }
    return n;
  }
  return index;
}

// Rewrite a formula's code while copying its opaque regions (string literals, single-quoted sheet
// names, bracketed structured references) verbatim. `transform` sees each maximal run of code between
// those regions and returns its replacement; the opaque text is never handed to it, so a literal like
// `"FILTER("` is never mistaken for a call and a `,` inside a structured reference never reads as a
// separator. Concatenating the transformed runs with the copied regions reproduces the formula.
function scanFormula(formula: string, transform: (code: string) => string): string {
  let out = '';
  let codeStart = 0;
  let i = 0;
  const n = formula.length;
  while (i < n) {
    const past = skipOpaque(formula, i);
    if (past > i) {
      out += transform(formula.slice(codeStart, i));
      out += formula.slice(i, past);
      i = past;
      codeStart = past;
    } else {
      i += 1;
    }
  }
  return out + transform(formula.slice(codeStart));
}

/**
 * Prefix every modern function called by its plain name with `_xlfn.` so Excel accepts the stored
 * formula. Names already prefixed are left alone (never doubled), unknown/legacy functions pass
 * through untouched, and opaque regions (string literals, sheet names, structured references) are
 * preserved verbatim. No other rewriting occurs: in particular no `@` implicit-intersection operator
 * is ever introduced.
 */
export function mangleFunctions(formula: string): string {
  return scanFormula(formula, (code) =>
    code.replace(FUNCTION_CALL, (whole, name: string, open: string) =>
      MODERN_FUNCTIONS.has(name.toUpperCase()) ? `${XLFN}${name}${open}` : whole,
    ),
  );
}

/**
 * Strip the `_xlfn.` function prefix and the `_xlpm.` LET-parameter prefix back to the plain names,
 * so the model holds the readable form regardless of how a file stored it. Opaque regions (string
 * literals, sheet names, structured references) are left untouched.
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

// The parameter names a LET/LAMBDA call binds. LAMBDA binds every argument but its last (the body);
// LET binds the even-indexed arguments up to but excluding its last (the calculation).
function parameterNames(
  formula: string,
  keyword: string,
  args: [number, number][],
): ReadonlySet<string> {
  const names = new Set<string>();
  const isLambda = keyword === 'LAMBDA';
  for (let a = 0; a < args.length - 1; a += 1) {
    if (isLambda || a % 2 === 0) {
      const name = boundName(formula, args[a] as [number, number]);
      if (name !== undefined) names.add(name);
    }
  }
  return names;
}

/**
 * Prefix every LET/LAMBDA parameter identifier with `_xlpm.`, at its declaration and at each
 * reference within the binding call's parentheses, so Excel accepts the stored formula. The prefix
 * is lexically scoped: a name is only rewritten inside the call that binds it, opaque regions are
 * copied verbatim, and a lambda-valued parameter used as a call (`f(…)`) is prefixed too. Formulas
 * with no LET/LAMBDA pass through unchanged.
 */
export function mangleParams(formula: string): string {
  // This is the one pass scanFormula cannot serve: it must know when a paren opens a LET/LAMBDA scope
  // and when the matching paren closes it, so it carries frame state across the code between opaque
  // regions rather than transforming each run in isolation. It still defers to skipOpaque, keeping the
  // opaque-skipping rule in one place even though the forward walk here is bespoke.
  let out = '';
  let i = 0;
  const n = formula.length;
  // A stack of active bindings, each expiring exactly at its owner call's close paren. Nested
  // LET/LAMBDA push inner frames that pop first, so shadowing resolves to the same prefix anyway.
  const frames: {end: number; names: ReadonlySet<string>}[] = [];
  const inScope = (name: string): boolean => frames.some((frame) => frame.names.has(name));

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

    if (heads && SCOPING_FUNCTIONS.has(name.toUpperCase()) && !inScope(name)) {
      const {close, args} = parseCall(formula, k);
      // The keyword stays at the outer scope; its parameters take effect inside the parens.
      out += formula.slice(i, k + 1);
      frames.push({end: close, names: parameterNames(formula, name.toUpperCase(), args)});
      i = k + 1;
      continue;
    }

    // Any other identifier: a bare reference, an ordinary call, or a lambda-valued parameter call.
    // In-scope names (declaration sites included, as they lie inside their own binding's parens) take
    // the prefix; the rest pass through. Call arguments are covered by the continuing scan, so a
    // nested LET/LAMBDA within them is still seen.
    out += inScope(name) ? `${XLPM}${name}` : name;
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

// The three reference shapes a shift moves, tried at each position in this order: a whole-column
// range (`A:C`), a whole-row range (`1:5`), then a single cell (`A1`). Each axis takes an optional `$`.
// Column letters are uppercase-only because Excel stores them that way, so a lowercase defined name is
// never mistaken for one, and row digits stop at seven (Excel's last row is 1048576).
//
// The lookbehind rejects a reference glued to a preceding name character or '.', so the `A1` inside
// `_xlfn.A1` or a defined name `FOO_A1` is left alone. The lookahead rejects one continued by a name
// character, opening a call `(`, or preceding a sheet `!`: a token before `!` is the sheet name
// (`Q1!A1`), not a cell. A sheet-qualified reference still shifts, because the `!` before it is not a
// name character. Applied per code run, where opaque regions such as a `"10:30"` literal are gone.
//
// One pattern rather than a pass per shape, so a shifted result is never scanned again, and `A1:B2`
// still resolves as two cells: a range alternative needs a bare letter run or a bare digit run on both
// sides of its colon. Cells used to be the only shape, so `A:A` and `1:1` never moved.
const SHIFTABLE_REFERENCE =
  /(?<![A-Za-z0-9_.])(?:(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})|(\$?)([0-9]{1,7}):(\$?)([0-9]{1,7})|(\$?)([A-Z]{1,3})(\$?)([0-9]{1,7}))(?![A-Za-z0-9_.!(])/g;

/**
 * Shift every relative reference in a formula by `colDelta` columns and `rowDelta` rows, leaving
 * absolute (`$`-anchored) axes fixed. This is how a shared-formula clone recovers its own formula from
 * the master's: a master `A1*2` shared one row down reads back as `A2*2`, `$A$1*B1` shared one row and
 * one column across as `$A$1*C2`, and `COUNTIF(A:A,A1)` shared one column right as `COUNTIF(B:B,B1)`.
 * Whole-column and whole-row ranges shift at each relative end, as cells do. String literals,
 * single-quoted sheet names, and bracketed structured references are copied verbatim, and a
 * sheet-qualified reference shifts while its sheet name is untouched. Function names and defined names
 * are not reference-shaped, so they pass through.
 */
export function translateFormula(formula: string, colDelta: number, rowDelta: number): string {
  if (colDelta === 0 && rowDelta === 0) return formula;
  // A reference the grid cannot hold becomes `#REF!` on either axis, which is what Excel writes for the
  // same shift; the decode is tolerant so that answer is reachable at all. The deltas come from a
  // file's own shared-formula geometry, so this is a read path, and an odd file aborting the whole
  // sheet with an error outside the library's taxonomy is not an answer. Both axes used to do exactly
  // that, in opposite ways. The column went through `columnToNumber`, which threw a bare `RangeError`
  // before any guard could speak: the pattern matches three letters, so `ZZZ1` (column 18278) reached
  // it, and a reference past XFD failed where a shift past XFD resolved. The row axis just did the
  // arithmetic and emitted `A0` or `A-4`, which is not a reference at all.
  const column = (anchor: string, letters: string): string | undefined => {
    const decoded = tryColumnToNumber(letters);
    if (decoded === undefined) return undefined;
    const col = anchor === '$' ? decoded : decoded + colDelta;
    return col < 1 || col > MAX_COLUMN ? undefined : `${anchor}${numberToColumn(col)}`;
  };
  const row = (anchor: string, digits: string): string | undefined => {
    const index = anchor === '$' ? Number(digits) : Number(digits) + rowDelta;
    return index < 1 || index > MAX_ROW ? undefined : `${anchor}${index}`;
  };
  const range = (first: string | undefined, last: string | undefined): string =>
    first === undefined || last === undefined ? REF_ERROR : `${first}:${last}`;

  return scanFormula(formula, (code) =>
    code.replace(
      SHIFTABLE_REFERENCE,
      (
        _match,
        firstColumnAnchor: string | undefined,
        firstColumn: string | undefined,
        lastColumnAnchor: string | undefined,
        lastColumn: string | undefined,
        firstRowAnchor: string | undefined,
        firstRow: string | undefined,
        lastRowAnchor: string | undefined,
        lastRow: string | undefined,
        cellColumnAnchor: string | undefined,
        cellColumn: string | undefined,
        cellRowAnchor: string | undefined,
        cellRow: string | undefined,
      ) => {
        // An anchor group is `''`, never `undefined`, whenever its alternative matched, so the
        // `?? ''` only ever meets an alternative that did not.
        if (firstColumn !== undefined && lastColumn !== undefined) {
          return range(
            column(firstColumnAnchor ?? '', firstColumn),
            column(lastColumnAnchor ?? '', lastColumn),
          );
        }
        if (firstRow !== undefined && lastRow !== undefined) {
          return range(row(firstRowAnchor ?? '', firstRow), row(lastRowAnchor ?? '', lastRow));
        }
        const shiftedColumn = column(cellColumnAnchor ?? '', cellColumn ?? '');
        const shiftedRow = row(cellRowAnchor ?? '', cellRow ?? '');
        return shiftedColumn === undefined || shiftedRow === undefined
          ? REF_ERROR
          : `${shiftedColumn}${shiftedRow}`;
      },
    ),
  );
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
