// Cluster: formulas
//
// Real-world scenario: a user writes a cell formula using a modern Excel function that post-dates
// the original OOXML function grammar, FILTER, XLOOKUP, LET, SEQUENCE, by its plain, readable name
// (e.g. `=FILTER(A1:D1, A2:D2 = N1)`). When the file is opened in current desktop Excel, the formula
// is silently dropped or shown as removed, even though typing the identical formula into Excel works
// fine. The cause is an OOXML storage convention: functions introduced after the frozen legacy
// grammar must be stored in the sheet XML with a `_xlfn.` name-mangling prefix (and LET's defined
// names with a `_xlpm.` prefix). The prefix is an on-disk detail: the reader must strip it back to
// the plain name, but if the writer never applies it, Excel rejects the formula.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'modern-function-xlfn-prefix-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'formulas',
  description:
    'A modern function written by its plain name (FILTER, XLOOKUP, …) is persisted in the sheet XML ' +
    'with the required _xlfn. name-mangling prefix so Excel accepts it, while a formula whose input ' +
    'already carries the prefix is not double-prefixed.',

  behavior: [
    {
      // FILTER is one of the worksheet-only functions, which Excel stores with `_xlws.` after `_xlfn.`.
      name: 'a modern function written by its plain name is stored with its future-function prefix on disk',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheets} = await api.inspectPackage({
          sheets: [{name: 'S', cells: [{ref: 'A1', formula: 'FILTER(B1:D1,B2:D2=1)', result: 0}]}],
        });
        assert.ok(
          /_xlfn\._xlws\.FILTER/.test(sheets.S!.formulas.A1 || ''),
          `FILTER must be stored as _xlfn._xlws.FILTER, as Excel stores it; got: ${sheets.S!.formulas.A1}`,
        );
      },
    },
    {
      name: 'a formula whose input already carries an explicit _xlfn. prefix is not double-prefixed',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheets} = await api.inspectPackage({
          sheets: [
            {name: 'S', cells: [{ref: 'A1', formula: '_xlfn.XLOOKUP(1,B:B,C:C)', result: 0}]},
          ],
        });
        const f = sheets.S!.formulas.A1 || '';
        assert.ok(/_xlfn\.XLOOKUP/.test(f), 'the explicit prefix survives');
        assert.ok(!/_xlfn\._xlfn/.test(f), `the prefix must not be doubled; got: ${f}`);
      },
    },
    {
      // A separate corruption reported alongside the missing prefix: a spurious `@`
      // implicit-intersection operator injected in front of references the author never marked,
      // which by itself makes Excel flag the formula. Writing a modern-function formula in plain
      // syntax must not fabricate an `@` on its range references.
      name: 'writing a modern-function formula does not inject a spurious @ implicit-intersection operator',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheets} = await api.inspectPackage({
          sheets: [
            {name: 'S', cells: [{ref: 'A1', formula: 'IFS(B1>0,"pos",B1<0,"neg")', result: 'pos'}]},
          ],
        });
        const f = sheets.S!.formulas.A1 || '';
        assert.ok(
          !/(^|[^A-Za-z0-9_])@/.test(f),
          `no @ implicit-intersection operator should be injected; got: ${f}`,
        );
      },
    },
    {
      // LET and the lambda family (LAMBDA/BYROW) are modern functions too, so they need the _xlfn.
      // prefix; additionally LET's own parameter names are stored with a _xlpm. prefix. A real-world
      // formula combining them (LET + BYROW + LAMBDA + FILTER) that is written verbatim, with none of
      // these prefixes, is rejected by Excel as corrupt.
      name: 'a LET/LAMBDA formula is stored with the _xlfn. function prefix Excel requires',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheets} = await api.inspectPackage({
          sheets: [
            {
              name: 'S',
              cells: [
                {
                  ref: 'A1',
                  formula: 'LET(a,B2:B9,b,BYROW(a,LAMBDA(r,SUM(r))),COUNTA(UNIQUE(FILTER(a,b=1))))',
                  result: 0,
                },
              ],
            },
          ],
        });
        const f = sheets.S!.formulas.A1 || '';
        assert.ok(
          /_xlfn\.LET/.test(f),
          `LET must be stored as _xlfn.LET for Excel to accept it; got: ${f}`,
        );
      },
    },
    {
      // The 2010 statistical-consistency rename family (NORM.DIST, T.DIST.2T, …) also post-dates the
      // frozen grammar and needs the `_xlfn.` prefix, but its names carry an internal '.'. The whole
      // dotted name must be prefixed once, `_xlfn.NORM.DIST`, not its trailing segment.
      name: 'a dotted statistical function is stored whole with the _xlfn. prefix, not on its tail segment',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheets} = await api.inspectPackage({
          sheets: [
            {name: 'S', cells: [{ref: 'A1', formula: 'NORM.DIST(A2,0,1,TRUE)', result: 0.5}]},
          ],
        });
        const f = sheets.S!.formulas.A1 || '';
        assert.ok(
          /_xlfn\.NORM\.DIST/.test(f),
          `NORM.DIST must be stored as _xlfn.NORM.DIST; got: ${f}`,
        );
        assert.ok(
          !/_xlfn\.DIST/.test(f),
          `the tail segment must not be prefixed on its own; got: ${f}`,
        );
      },
    },
    {
      // Not every post-2007 function is dotted or a dynamic array: the 2010/2013 releases also added
      // plain-named trigonometric, bitwise and engineering functions (SEC, BITAND, AGGREGATE, XOR, …)
      // that carry the same `_xlfn.` requirement. A pre-2007 cousin of the same shape (SIN) must be
      // left bare, so this pins that the prefix follows the function's vintage, not its spelling.
      name: 'a bare-name post-2007 function is prefixed while its pre-2007 cousin is left bare',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheets} = await api.inspectPackage({
          sheets: [{name: 'S', cells: [{ref: 'A1', formula: 'SEC(A2)+SIN(A2)', result: 0}]}],
        });
        const f = sheets.S!.formulas.A1 || '';
        assert.ok(/_xlfn\.SEC/.test(f), `SEC must be stored as _xlfn.SEC; got: ${f}`);
        assert.ok(!/_xlfn\.SIN/.test(f), `the pre-2007 SIN must not be prefixed; got: ${f}`);
      },
    },
  ],
} satisfies Case;
