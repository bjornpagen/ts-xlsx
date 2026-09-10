// Spelling a structured reference to a table column. This lives apart from `formula.ts` because the
// table model needs it and `formula.ts` carries the modern-function list every entry would then pay for.

// The characters that make Excel spell a column specifier in the double-bracket form, `Table[[Col]]`,
// and the subset of them it escapes with a leading `'` inside one. Both lists are Microsoft's own, from
// "Using structured references with Excel tables" (support.microsoft.com). A name holding none of them,
// spaces included, keeps the single-bracket form.
const STRUCTURED_REFERENCE_SPECIAL = /[\t\n\r,:.[\]#'"{}$^&*+=></@\\!()%?`;~_-]/;
const STRUCTURED_REFERENCE_ESCAPED = /[[\]#'@]/g;

/**
 * A structured reference to one column of a table, spelled as Excel spells it: `Table[Column]` for a
 * plain name, and `Table[[Column]]` for a name holding a character from Microsoft's special list, with
 * `[`, `]`, `#`, `'` and `@` escaped by a leading `'`. Interpolating the name bare wrote
 * `SUBTOTAL(109,T[Price [USD]])`, a formula whose brackets close in the middle of the name.
 */
export function structuredColumnReference(tableName: string, columnName: string): string {
  if (!STRUCTURED_REFERENCE_SPECIAL.test(columnName)) return `${tableName}[${columnName}]`;
  return `${tableName}[[${columnName.replace(STRUCTURED_REFERENCE_ESCAPED, "'$&")}]]`;
}
