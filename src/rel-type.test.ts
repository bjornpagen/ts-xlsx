import assert from 'node:assert/strict';
import {test} from 'node:test';

import {isAnyRelType, isRelType, relTypeSegment} from './rel-type.ts';

const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

test('isRelType matches a whole trailing segment, never a suffix of one', () => {
  assert.equal(isRelType(`${OFFICE}/table`, 'table'), true);
  assert.equal(isRelType(`${OFFICE}/pivotTable`, 'table'), false);
  assert.equal(isRelType(`${OFFICE}/slicerCache`, 'slicer'), false);
});

test('isRelType matches a part class whose name spans segments', () => {
  assert.equal(
    isRelType(
      'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility',
      'ui/extensibility',
    ),
    true,
  );
});

test('isAnyRelType matches when any one of the classes does', () => {
  assert.equal(isAnyRelType(`${OFFICE}/chart`, 'drawing', 'chart'), true);
  assert.equal(isAnyRelType(`${OFFICE}/chartsheet`, 'drawing', 'chart'), false);
});

test('relTypeSegment is the last segment, or the whole Type when there is no slash', () => {
  assert.equal(
    relTypeSegment(
      'http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile',
    ),
    'vbaProjectSignatureAgile',
  );
  assert.equal(relTypeSegment('vbaProject'), 'vbaProject');
  assert.equal(relTypeSegment(`${OFFICE}/`), '');
});
