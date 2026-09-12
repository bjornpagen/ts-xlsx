import assert from 'node:assert/strict';
import {test} from 'node:test';

import {AuthoringError} from '../errors.ts';
import {HyperlinkOverlay} from './hyperlink.ts';

test('a link is stored under the canonical spelling of the cells it covers', () => {
  const links = new HyperlinkOverlay();
  links.add({ref: '$B$2', target: 'https://example.com/'});
  links.add({ref: 'D1:D1', target: 'https://example.com/'});
  links.add({ref: 'H1:D1', target: 'https://example.com/'});
  assert.deepEqual(
    links.entries.map((link) => link.ref),
    ['B2', 'D1', 'D1:H1'],
  );
});

test('a link over the same cells replaces the one there, and counts as added last', () => {
  const links = new HyperlinkOverlay();
  links.add({ref: 'A1', target: 'https://first.example/'});
  links.add({ref: 'A1:C1', target: 'https://wide.example/'});
  links.add({ref: 'A1', target: 'https://second.example/', tooltip: 'again'});

  assert.deepEqual(links.entries, [
    {ref: 'A1:C1', target: 'https://wide.example/'},
    {ref: 'A1', target: 'https://second.example/', tooltip: 'again'},
  ]);
  assert.equal(links.at(1, 1)?.target, 'https://second.example/');
  assert.equal(links.at(2, 1)?.target, 'https://wide.example/');
  assert.equal(links.at(4, 1), undefined);
});

test('removing takes only the link over exactly those cells', () => {
  const links = new HyperlinkOverlay();
  links.add({ref: 'A1:C1', target: 'https://wide.example/'});
  links.add({ref: 'B1', target: 'https://narrow.example/'});

  assert.equal(links.remove('A1'), false, 'no link covers exactly A1');
  assert.equal(links.remove('$B$1'), true);
  assert.equal(links.remove('junk!!'), false);
  assert.deepEqual(links.entries, [{ref: 'A1:C1', target: 'https://wide.example/'}]);
});

test('a link naming a whole row or column, or nothing, is refused', () => {
  const links = new HyperlinkOverlay();
  assert.throws(() => links.add({ref: 'A:A', target: 'x'}), AuthoringError);
  assert.throws(() => links.add({ref: '3:3', target: 'x'}), AuthoringError);
  assert.throws(() => links.add({ref: 'junk!!', target: 'x'}), SyntaxError);
  assert.deepEqual(links.entries, []);
});

test('a link keeps its own fields only, copied', () => {
  const links = new HyperlinkOverlay();
  const given = {ref: 'A1', target: 'https://example.com/', label: 'not a link field'};
  links.add(given);
  given.target = 'https://changed.example/';
  assert.deepEqual(links.entries, [{ref: 'A1', target: 'https://example.com/'}]);
});

test("a row copy gives the destination the source row's own links, and not a taller one", () => {
  const links = new HyperlinkOverlay();
  links.add({ref: 'A1', target: 'https://a.example/'});
  links.add({ref: 'B1:C1', target: 'https://b.example/'});
  links.add({ref: 'D1:D2', target: 'https://tall.example/'});
  links.copyRow(1, 3);
  assert.deepEqual(
    links.entries.map((link) => link.ref),
    ['A1', 'B1:C1', 'D1:D2', 'A3', 'B3:C3'],
  );
});
