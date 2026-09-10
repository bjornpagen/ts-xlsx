// Cluster: security
//
// Real-world scenario: pre-2010 Excel, XlsxWriter, openpyxl and LibreOffice protect a sheet with the
// legacy 16-bit `password="XXXX"` hash rather than the agile SHA-512 credential. A reader with no place
// for that hash writes the sheet back still marked protected but with no password at all, so anyone
// can lift the protection its author set. The same element's `sheet` attribute defaults to false, and
// a reader that takes its absence as protected locks, on every save, a sheet its file left open.

import type {Assert, Case, CorpusApi} from '../case.ts';

const LEGACY = 'password="CC3D" sheet="1" objects="1" scenarios="1"';

export default {
  id: 'legacy-sheet-password-survives-a-round-trip',
  provenance: {source: 'audit'},
  cluster: 'security',
  description:
    'A sheet protected with the legacy 16-bit password hash is still protected by that same hash ' +
    'after a read and a save, a malformed hash is dropped rather than written back, and a ' +
    '<sheetProtection> element without sheet="1" leaves the sheet unprotected.',

  behavior: [
    {
      name: 'a sheet protected by a legacy hash reads as protected by that hash',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.legacySheetProtectionReport(LEGACY);
        assert.equal(report.protected, true);
        assert.equal(report.legacyPasswordHash, 'CC3D');
      },
    },
    {
      name: 'a save writes the same hash back beside sheet="1"',
      expect(api: CorpusApi, assert: Assert) {
        const {rewritten} = api.legacySheetProtectionReport(LEGACY);
        assert.equal(rewritten?.password, 'CC3D');
        assert.equal(rewritten?.sheet, '1');
      },
    },
    {
      name: 'a malformed legacy hash is dropped and the protection it came with is kept',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.legacySheetProtectionReport('password="not-a-hash" sheet="1"');
        assert.equal(report.protected, true);
        assert.equal(report.legacyPasswordHash, null);
        assert.equal(report.rewritten?.password, undefined);
      },
    },
    {
      name: 'an element without sheet="1" leaves the sheet unprotected, and a save does not lock it',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.legacySheetProtectionReport('objects="1" scenarios="1"');
        assert.equal(report.protected, false);
        assert.equal(report.rewritten, null);
      },
    },
  ],
} satisfies Case;
