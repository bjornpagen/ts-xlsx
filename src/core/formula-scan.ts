// Every pass over formula text shares one hazard: a comma, paren, function name, or cell reference is
// mere text when it sits inside a string literal, a single-quoted sheet name, or a bracketed structured
// reference. `skipOpaque` is the single owner of skipping those regions, so the rule lives in one
// place, and the stateless passes (function-name and reference rewriting) ride `scanFormula`, which
// copies the opaque regions verbatim and hands each code run between them to a transform.

// Advance past the opaque region opened at `index`: a double-quoted string literal or a single-quoted
// sheet name, both honouring the doubled-quote escape (`""`, `''`), or a bracketed structured
// reference, which may nest (`Table[[#Data],[Col]]`) and escapes a bracket inside a column name with a
// leading `'` (`T[[a'[b]]`). Returns the index just past the region, or
// `index` unchanged when no opaque region opens there. Inside any of the three a comma, paren, function
// name, or cell reference is inert, so every pass over a formula skips them through this one function.
export function skipOpaque(formula: string, index: number): number {
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
export function scanFormula(formula: string, transform: (code: string) => string): string {
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
