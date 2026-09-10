import assert from 'node:assert/strict';
import {test} from 'node:test';

import {structuredColumnReference} from './structured-reference.ts';

test('a plain column name keeps the single-bracket structured reference, spaces included', () => {
  assert.equal(structuredColumnReference('T', 'Amount'), 'T[Amount]');
  assert.equal(structuredColumnReference('T', 'Unit Price'), 'T[Unit Price]');
});

test("[ ] # ' and @ are escaped with ' inside the double-bracket form", () => {
  assert.equal(structuredColumnReference('T', 'Price [USD]'), "T[[Price '[USD']]]");
  assert.equal(structuredColumnReference('T', "it's #1"), "T[[it''s '#1]]");
  assert.equal(structuredColumnReference('T', 'a@b'), "T[[a'@b]]");
});

test("every other character on Microsoft's special list takes the double-bracket form unescaped", () => {
  const unescaped = ['\t', '\n', '\r', ',', ':', '.', '"', '{', '}', '$', '^', '&', '*', '+', '='];
  unescaped.push('-', '>', '<', '/', '\\', '!', '(', ')', '%', '?', '`', ';', '~', '_');
  for (const ch of unescaped) {
    assert.equal(structuredColumnReference('T', `a${ch}b`), `T[[a${ch}b]]`, JSON.stringify(ch));
  }
});
