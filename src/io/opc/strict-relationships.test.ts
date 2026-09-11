import assert from 'node:assert/strict';
import {test} from 'node:test';

import {transitionalRelationshipType} from './strict-relationships.ts';

test('a Strict relationship type is spelled as Transitional spells it, renamed segments included', () => {
  const strict = 'http://purl.oclc.org/ooxml/officeDocument/relationships/';
  const transitional = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  assert.equal(transitionalRelationshipType(`${strict}chart`), `${transitional}chart`);
  assert.equal(
    transitionalRelationshipType(`${strict}extendedProperties`),
    `${transitional}extended-properties`,
  );
  assert.equal(
    transitionalRelationshipType(`${strict}customProperties`),
    `${transitional}custom-properties`,
  );
  for (const type of [
    `${transitional}chart`,
    'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
    'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
  ]) {
    assert.equal(transitionalRelationshipType(type), type);
  }
});
