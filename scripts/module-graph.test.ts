// The harness's one shared definition of "what does this module import", tested.
//
// `module-graph.ts` is what `check-layering`, `check-browser-safe`, `size-budget` and `smoke-dist`
// all mean by an import, and nothing under `scripts/` had a test. Its comment-stripper is a
// hand-rolled character scanner, which is the right shape (a parser for four gates would cost more
// than it saves) and exactly the shape that gets one case wrong quietly.
//
// The cases below are that scanner's edges: what a `//` can hide behind, what a quote can hide
// behind, and the two import spellings that carry no `from` to anchor on.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {relativeSpecifiers, resolveSpecifier, specifiers, withoutComments} from './module-graph.ts';

test('a commented-out import is not an import', () => {
  assert.deepEqual(specifiers("// import {x} from './gone.ts';\nimport {y} from './here.ts';"), [
    './here.ts',
  ]);
  assert.deepEqual(specifiers("/* import {x} from './gone.ts'; */\nimport {y} from './here.ts';"), [
    './here.ts',
  ]);
});

test('a quote inside a regex literal does not open a string', () => {
  // The live failure this was written for: the apostrophe opened a phantom string, so the `//` on
  // the next line was never blanked and the specifier inside it read as a real import.
  const source = "const apostrophe = /don't/;\n// import {x} from './gone.ts';\n";
  assert.deepEqual(specifiers(source), []);
});

test('a regex holding a slash in a character class runs to its real end', () => {
  const source = "const path = /[/\\\\]+/;\n// import {x} from './gone.ts';\n";
  assert.deepEqual(specifiers(source), []);
});

test('a division is not a regex, however many quotes follow it', () => {
  // `total / 2` divides; treating the slash as a literal would swallow the rest of the file, and
  // with it every real import after this line.
  const source = "const half = total / 2;\nimport {x} from './here.ts';\n";
  assert.deepEqual(specifiers(source), ['./here.ts']);
});

test('a `//` inside a string is text, not a comment', () => {
  const source = "const url = 'https://example.com';\nimport {x} from './here.ts';\n";
  assert.deepEqual(specifiers(source), ['./here.ts']);
});

test('a template literal with an interpolation is walked to its close', () => {
  const source = 'const t = `a ${b} // c`;\nimport {x} from "./here.ts";\n';
  assert.deepEqual(specifiers(source), ['./here.ts']);
});

test('a `*/` inside a string does not close a block comment early', () => {
  // The stripper scans for `*/` textually, which is what the language does too: a block comment has
  // no string literals inside it. The case is here to record that this is the answer, not an
  // oversight, because it reads like one.
  const source = "/* a '*/ const s = '*/';\nimport {x} from './here.ts';\n";
  assert.deepEqual(specifiers(source), ['./here.ts']);
});

test('the word `from` or `import` inside a string is not an import', () => {
  // A corpus case split a sentence right after "from", and `'… from ' + '…'` read as an import of
  // ` + `: the keyword and the quote after it were both there, only inside two strings.
  assert.deepEqual(
    specifiers("const s = 'reads from ' + 'a book';\nimport {x} from './here.ts';\n"),
    ['./here.ts'],
  );
  assert.deepEqual(specifiers('const s = "call import " + "(x)";\nexport * from "./d.ts";\n'), [
    './d.ts',
  ]);
});

test('every spelling that carries a specifier is seen', () => {
  const source = [
    "import a from './a.ts';",
    "import './b.ts';",
    "export {c} from './c.ts';",
    "export * from './d.ts';",
    "const e = await import('./e.ts');",
    'import f from "./f.ts";',
  ].join('\n');
  assert.deepEqual(specifiers(source), [
    './a.ts',
    './b.ts',
    './c.ts',
    './d.ts',
    './e.ts',
    './f.ts',
  ]);
});

test('a bare specifier names a dependency and is not part of this graph', () => {
  const source = "import {zip} from 'fflate';\nimport {x} from './here.ts';";
  assert.deepEqual(specifiers(source), ['fflate', './here.ts']);
  assert.deepEqual(relativeSpecifiers(source), ['./here.ts']);
});

test('blanking a comment keeps the line numbering it was found on', () => {
  const blanked = withoutComments('a\n// two\n/* three\nfour */\nfive');
  assert.equal(blanked.split('\n').length, 5);
  assert.equal(blanked.split('\n')[4], 'five');
});

test('a specifier resolves against its importer, in the spelling the graph uses', () => {
  // Every path this module hands back is `/`-separated, whatever the platform's separator is: the
  // gates built on it compare against repo-relative keys and print them, and a mixed-separator key
  // matches nothing on Windows.
  const from = 'C:/repo/src/io/xlsx/read.ts';
  assert.equal(resolveSpecifier(from, './write.ts'), 'C:/repo/src/io/xlsx/write.ts');
  assert.equal(resolveSpecifier(from, '../opc/read-opc.ts'), 'C:/repo/src/io/opc/read-opc.ts');
  assert.equal(resolveSpecifier('C:\\repo\\src\\a.ts', './b.ts'), 'C:/repo/src/b.ts');
});
