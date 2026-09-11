// What a module imports, and everything a module reaches, for the four gates that ask.
//
// `check-layering`, `check-browser-safe`, `size-budget` and `smoke-dist` each walk the module graph,
// and each had grown its own copy of the walk: four file listers, four specifier regexes, two
// byte-identical path-segment collapse loops and two byte-identical closure functions.
//
// The four regexes did not agree, and the disagreement was live rather than theoretical.
// `size-budget` matched either quote style and recorded *why* in a comment: the emitter's quoting
// flipped between TypeScript 6 and 7. `check-layering` and `check-browser-safe` matched single quotes
// only, so the same flip in the *source* formatter would have made both gates see zero imports and
// report a clean graph. A gate that silently stops checking is worse than no gate, and these two are
// named in CLAUDE.md as part of what makes the machine-checkable net the primary guarantor of
// correctness. Four answers to "what does this module import" is four definitions of the graph that
// net is guaranteeing.
//
// None of the four stripped comments before matching either, so a commented-out import read as a real
// one: a false layering violation, from prose.
//
// Paths here are always `/`-separated, whatever the platform, because a graph keyed by path needs one
// spelling per file: half these callers build repo-relative strings by hand and half use `node:path`,
// and on Windows the two disagree on the separator, which would count one module twice.

import {readdirSync, readFileSync} from 'node:fs';

/** A path with `/` separators, whatever `node:path` produced on this platform. */
export function toPosix(path: string): string {
  return path.replaceAll('\\', '/');
}

/**
 * Every file under `dir` (recursively) whose name ends with `suffix`, as `/`-separated paths built
 * from `dir` as given. Tests are not source: a `.test.ts` is excluded from a `.ts` listing, because
 * every caller here asks about what ships.
 */
export function sourceFiles(dir: string, suffix: string): string[] {
  return readdirSync(dir, {withFileTypes: true}).flatMap((entry) => {
    const path = `${toPosix(dir)}/${entry.name}`;
    if (entry.isDirectory()) return sourceFiles(path, suffix);
    if (!entry.name.endsWith(suffix)) return [];
    return suffix === '.ts' && entry.name.endsWith('.test.ts') ? [] : [path];
  });
}

/**
 * Whether a `/` at this point opens a regex literal rather than dividing.
 *
 * The previous significant character is the whole test. After a value -- an identifier, a number, a
 * closing paren or bracket -- a slash divides; everywhere else (after `=`, `(`, `,`, `:`, `return`,
 * the start of the file) it opens a literal.
 */
function startsRegex(previous: string): boolean {
  return previous !== '' && !/[\w$)\]]/.test(previous);
}

// Both spellings, and both quote styles. `… from '…'` is the ordinary form; the bare `import '…'` has
// no `from` to anchor on and should never appear here, since the package declares `"sideEffects":
// false` and an import kept only for its effect is a lie to every bundler. But a rule that cannot see
// it would report a clean graph while a layer was being crossed by the one form it was blind to.
// The dynamic form is matched too, and `src/` has none today. That is the point: a graph that could
// not see `import('…')` would report a clean answer the day one appeared, and the four gates built on
// this would all be wrong at once. A hole nothing records is not a known limitation.
const SPECIFIER = /\b(?:from|import)\s+["']([^"']*)["']|\bimport\s*\(\s*["']([^"']*)["']/dg;

/**
 * Every specifier a module imports or re-exports from, in source order, comments excluded.
 *
 * The pattern runs over the source with string contents blanked too, so `from` or `import` followed by
 * a quote only counts where both are code: `'… from ' + '…'` used to read as an import of ` + `. The
 * specifier is then read at the same offsets from the source with only comments blanked, which is
 * where its text still is.
 */
export function specifiers(source: string): string[] {
  const text = withoutComments(source);
  return [...blanked(source, true).matchAll(SPECIFIER)].map((match) => {
    const [start, end] = match.indices?.[1] ?? match.indices?.[2] ?? [0, 0];
    return text.slice(start, end);
  });
}

/** Only the relative specifiers: a bare one names a dependency, which is not part of this graph. */
export function relativeSpecifiers(source: string): string[] {
  return specifiers(source).filter((specifier) => specifier.startsWith('.'));
}

/**
 * Resolve a relative specifier against the module that wrote it, collapsing `.` and `..`.
 *
 * Deliberately not `node:path`: these paths are graph keys rather than filesystem arguments, and
 * `join`/`normalize` would hand back `\` on Windows for callers that built their side with `/`.
 */
export function resolveSpecifier(fromFile: string, specifier: string): string {
  const from = toPosix(fromFile);
  const dir = from.slice(0, from.lastIndexOf('/'));
  const out: string[] = [];
  for (const segment of `${dir}/${specifier}`.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') out.pop();
    else out.push(segment);
  }
  // A leading `/` (or a drive-relative Windows path) survives the split as an empty first segment,
  // which the loop drops; put it back so an absolute path stays absolute.
  return (from.startsWith('/') ? '/' : '') + out.join('/');
}

/** Every relative import of `file`, read off disk and resolved. */
export function importedPaths(file: string): string[] {
  return relativeSpecifiers(readFileSync(file, 'utf8')).map((specifier) =>
    resolveSpecifier(file, specifier),
  );
}

/**
 * Every module that has to be present for `entry` to evaluate, `entry` included.
 *
 * `imports` is a parameter because the two source gates read repo-relative paths through a root and
 * the two `dist` gates read real ones; what they share is the walk, not the lookup.
 */
export function closure(entry: string, imports: (file: string) => string[]): Set<string> {
  const reached = new Set<string>();
  const pending = [toPosix(entry)];
  while (pending.length > 0) {
    const file = pending.pop() as string;
    if (reached.has(file)) continue;
    reached.add(file);
    pending.push(...imports(file));
  }
  return reached;
}

/**
 * A module's code with its comments blanked out, newlines kept so a match still reports its own line.
 *
 * Strings are walked rather than skipped so a `//` inside one is not mistaken for the start of a
 * comment. Exported because `check-browser-safe` scans the same blanked source for the Node globals
 * it forbids, for the same reason: prose is where those identifiers legitimately appear.
 *
 * Regex literals are walked for the same reason one step further out. `const a = /don't/;` used to
 * open a phantom string at the apostrophe, so every `//` until the next quote stayed unblanked and a
 * commented-out import read as a real one. That degrades toward a false alarm, which is the safe
 * direction, but a gate that can fail for reasons unrelated to the code is how a gate loses its
 * reader. Whether a `/` opens a regex or divides is decided by the previous significant character,
 * the usual approximation: after an identifier, a number, a `)` or a `]` it divides, and otherwise it
 * opens one. `if (x) /re/.test(y)` is the shape that defeats it, and there is none in this tree.
 */
export function withoutComments(source: string): string {
  return blanked(source, false);
}

// The scanner behind both views of a module. With `strings` set, a single- or double-quoted string
// keeps its quotes and loses its contents, so offsets still line up with the source. A template literal
// keeps its contents either way: blanking one would also blank an `import('…')` in a `${}` inside it,
// and a graph that misses an import reports a clean answer, where a false one only raises an alarm.
function blanked(source: string, strings: boolean): string {
  let out = '';
  // The last character that was not whitespace, which is all the context deciding regex-or-division
  // needs. Comments are already blanked by the time it is read, so a comment between the operand and
  // the slash does not confuse it.
  let previous = '';
  for (let i = 0; i < source.length; i++) {
    const two = source.slice(i, i + 2);
    if (two === '//') {
      const end = source.indexOf('\n', i);
      const stop = end === -1 ? source.length : end;
      out += ' '.repeat(stop - i);
      i = stop - 1;
      continue;
    }
    if (two === '/*') {
      const end = source.indexOf('*/', i + 2);
      const stop = end === -1 ? source.length : end + 2;
      out += source.slice(i, stop).replace(/[^\n]/g, ' ');
      i = stop - 1;
      continue;
    }
    const char = source[i] as string;
    out += char;
    if (char === '/' && startsRegex(previous)) {
      previous = '/';
      // Inside a regex literal: copy to its close, honouring escapes and a character class, inside
      // which an unescaped `/` is an ordinary character rather than the terminator.
      let inClass = false;
      for (i += 1; i < source.length; i++) {
        const inner = source[i] as string;
        out += inner;
        if (inner === '\\') {
          out += source[i + 1] ?? '';
          i += 1;
          continue;
        }
        if (inner === '[') inClass = true;
        else if (inner === ']') inClass = false;
        else if (inner === '/' && !inClass) break;
        else if (inner === '\n') break; // Unterminated: not a regex after all, do not eat the file.
      }
      continue;
    }
    if (!/\s/.test(char)) previous = char;
    if (char !== "'" && char !== '"' && char !== '`') continue;
    // Inside a string literal: copy to its close, honouring backslash escapes.
    const blank = strings && char !== '`';
    const copy = (text: string) => (blank && text !== '\n' ? ' '.repeat(text.length) : text);
    for (i += 1; i < source.length; i++) {
      const inner = source[i] as string;
      if (inner === '\\') {
        out += copy(inner) + copy(source[i + 1] ?? '');
        i += 1;
        continue;
      }
      if (inner === char) {
        out += inner;
        break;
      }
      out += copy(inner);
    }
    previous = char;
  }
  return out;
}
