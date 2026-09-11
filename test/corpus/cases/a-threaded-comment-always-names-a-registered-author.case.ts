// Cluster: comment
//
// Real-world scenario: a generator leaves review comments without knowing who wrote them, so it gives
// a message no author, or it names an author id it never registered in the workbook's person list. A
// third producer leaves a package whose thread points at a person its registry does not hold. The 2018
// schema requires `personId` on every message, and Excel wants the id to name a registered person:
// asked on 2026-09-11 (Excel 16.0 build 20326, open verdict), it offers to repair a package with either
// shape and opens one whose author is registered clean.
//
// Excel's own repair of the unregistered id is the answer to what an unknown author looks like. It
// rewrites the id to the null GUID, `{00000000-0000-0000-0000-000000000000}`, shows the message as
// written by "Author", and opens that package clean, with or without a registry beside it. So a message
// with no author is written under the null GUID, an authored id nobody registered is refused where the
// workbook is written, and a file's unregistered id is read as no recorded author, which writes back as
// the null GUID the way Excel's repair would.

import type {Assert, Case, CorpusApi} from '../case.ts';

const UNKNOWN_AUTHOR = '{00000000-0000-0000-0000-000000000000}';

export default {
  id: 'a-threaded-comment-always-names-a-registered-author',
  provenance: {source: 'excel-desktop-verification'},
  cluster: 'comment',
  description:
    'Every threaded comment is written with a personId Excel accepts: a message with no author under ' +
    'the null GUID Excel itself uses, an authored id missing from the person registry refused at write, ' +
    'and an unregistered id read from a file treated as no recorded author.',

  behavior: [
    {
      name: 'a message authored with no author is written under the null GUID and reads back authorless',
      expect(api: CorpusApi, assert: Assert) {
        const {authorless} = api.threadAuthorReport();
        assert.strictEqual(authorless.writeError, null);
        assert.deepStrictEqual(authorless.writtenPersonIds, [UNKNOWN_AUTHOR]);
        assert.deepStrictEqual(authorless.readPersonIds, [null]);
      },
    },
    {
      name: 'an authored author id the person registry does not hold is refused when written',
      expect(api: CorpusApi, assert: Assert) {
        const {unregistered} = api.threadAuthorReport();
        assert.match(unregistered.writeError ?? '', /not a registered person/);
      },
    },
    {
      name: 'a file whose thread names an unregistered author reads it as no author and writes back',
      expect(api: CorpusApi, assert: Assert) {
        const {foreign} = api.threadAuthorReport();
        assert.deepStrictEqual(foreign.readAs, [null], 'the dangling id is not kept');
        assert.strictEqual(foreign.writeError, null, 'so the read workbook can be saved');
        assert.deepStrictEqual(foreign.writtenPersonIds, [UNKNOWN_AUTHOR]);
      },
    },
  ],
} satisfies Case;
