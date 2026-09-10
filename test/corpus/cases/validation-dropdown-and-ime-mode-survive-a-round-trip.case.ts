// Cluster: data-validation
//
// Real-world scenario: an order form validates a column against a list without showing the dropdown
// arrow (`showDropDown="1"`, the usual way to enforce a list without offering a picker), and a name
// field on a Japanese form switches the input method to hiragana (`imeMode`). The reader modelled
// neither attribute, so one load and save brought the arrow back and dropped the input mode.

import type {Assert, Case, CorpusApi} from '../case.ts';

const BOTH = {showDropDown: '1', imeMode: 'off'};

export default {
  id: 'validation-dropdown-and-ime-mode-survive-a-round-trip',
  cluster: 'data-validation',
  provenance: {source: 'audit'},
  description:
    'A list validation that hides its dropdown arrow and sets an input mode writes both attributes in ' +
    'the standard and the extended form, reads both back on each cell, and writes both again.',

  behavior: [
    {
      name: 'both attributes are written, in the standard form and the extended one',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.dropdownAndImeModeReport().written, {standard: BOTH, extended: BOTH});
      },
    },
    {
      name: 'a read keeps the hidden arrow and the input mode on each cell',
      expect(api: CorpusApi, assert: Assert) {
        const cell = {suppressDropDown: true, imeMode: 'off'};
        assert.deepEqual(api.dropdownAndImeModeReport().read, {A1: cell, B1: cell});
      },
    },
    {
      name: 'a second save writes both attributes again',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.dropdownAndImeModeReport().rewritten, {
          standard: BOTH,
          extended: BOTH,
        });
      },
    },
  ],
} satisfies Case;
