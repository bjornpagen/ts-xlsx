#!/usr/bin/env node
// A corpus case reaches the library only through the adapter.
//
// Implementation-blindness is why the corpus outlived the rewrite: a case states a behaviour in terms
// of the `CorpusApi` it is handed, so the same case judges whatever implementation stands behind the
// adapter. A case that imports a src module, or the adapter's own runtime, has quietly stopped being
// that. It still typechecks, still passes every gate, and still reads like a corpus case; only the
// import line says otherwise, and nobody reads import lines. So this reads them.
//
// A case may import `test/corpus/case.ts` and `test/corpus/untyped.ts`, which are the vocabulary a case
// is written in, and `node:` built-ins. Anything else is a violation, a bare package specifier
// included: a case that imports `fflate` to unzip a package is doing the adapter's job.
//
//   node scripts/check-corpus-blind.ts

import {readFileSync} from 'node:fs';

import {resolveSpecifier, sourceFiles, toPosix, withoutComments} from './module-graph.ts';
import {ROOT as REPO_ROOT} from './repo.ts';
import {reportCrash, verdict} from './verdict.ts';

const ROOT = toPosix(REPO_ROOT);
const CASES = 'test/corpus/cases';
const ALLOWED = new Set(['test/corpus/case.ts', 'test/corpus/untyped.ts']);

// Anchored to the start of a statement, where `module-graph.ts`'s `specifiers` matches `from '…'`
// anywhere. Cases are mostly prose in string literals, and one of them splits a sentence across a
// concatenation right after the word "from": `'… from ' + '…'` reads to the unanchored pattern as an
// import of ` + `. The dynamic form keeps no anchor, since `import(` is a call wherever it appears.
const IMPORT =
  /(?:^|;)\s*(?:import|export)\b[^;'"]*?\bfrom\s*["']([^"']*)["']|(?:^|;)\s*import\s*["']([^"']*)["']|\bimport\s*\(\s*["']([^"']*)["']/gm;

function importsOf(source: string): string[] {
  return [...withoutComments(source).matchAll(IMPORT)].map(
    (match) => match[1] ?? match[2] ?? match[3] ?? '',
  );
}

const repoRelative = (path: string): string => toPosix(path).slice(ROOT.length + 1);

function main(): void {
  const problems: string[] = [];
  const files = sourceFiles(`${ROOT}/${CASES}`, '.ts');
  for (const file of files) {
    for (const specifier of importsOf(readFileSync(file, 'utf8'))) {
      if (specifier.startsWith('node:')) continue;
      const target = specifier.startsWith('.')
        ? repoRelative(resolveSpecifier(file, specifier))
        : specifier;
      if (ALLOWED.has(target)) continue;
      problems.push(
        `  ${repoRelative(file)}\n    imports ${target}\n    a case may import only ${[...ALLOWED].join(', ')} and node: built-ins;\n` +
          '    reach the library through the CorpusApi the adapter hands the case',
      );
    }
  }

  verdict({
    gate: 'corpus blindness',
    problems,
    ok: `${files.length} cases import nothing but the case vocabulary`,
    failure: 'import(s) break a corpus case’s blindness',
  });
}

try {
  main();
} catch (error) {
  reportCrash('corpus blindness', error);
}
