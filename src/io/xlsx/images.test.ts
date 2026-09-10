import assert from 'node:assert/strict';
import {test} from 'node:test';

import {isOneCellAnchor} from '../../core/image.ts';
import {Workbook} from '../../core/workbook.ts';
import {imageContentType} from './images.ts';
import {
  partIn,
  partMatching,
  partsWritten,
  partText,
  patchParts,
  roundtrip,
} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

// A 1×1 transparent PNG: enough bytes to prove the media round-trips verbatim.
const ONE_PX_PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
);

function anchored(): Workbook {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addImage(id, {tl: {col: 0, row: 5}, br: {col: 2, row: 8}});
  return wb;
}

test('an anchored image survives the write/read round-trip with its bytes intact', () => {
  const back = roundtrip(anchored()).getWorksheet('S');
  const images = back?.images ?? [];
  assert.strictEqual(images.length, 1);
  const anchor = images[0]?.anchor;
  assert.ok(anchor && !isOneCellAnchor(anchor));
  assert.deepStrictEqual(anchor.from, {col: 0, row: 5, colOff: 0, rowOff: 0});
  assert.deepStrictEqual(anchor.to, {col: 2, row: 8, colOff: 0, rowOff: 0});
  const media = roundtrip(anchored()).getImage(images[0]?.imageId ?? -1);
  assert.strictEqual(media?.extension, 'png');
  assert.deepStrictEqual(media?.data, ONE_PX_PNG);
});

test('a noted-free workbook with an image emits a drawing, a media part, and a <drawing> reference', () => {
  const parts = partsWritten(anchored());
  assert.ok(parts['xl/drawings/drawing1.xml'], 'a drawing part is written');
  assert.ok(parts['xl/media/image1.png'], 'the media bytes are written');
  const sheetXml = partIn(parts, 'xl/worksheets/sheet1.xml');
  assert.match(sheetXml, /<drawing r:id="[^"]+"\/>/);
  const drawingXml = partIn(parts, 'xl/drawings/drawing1.xml');
  assert.match(drawingXml, /<xdr:from>[\s\S]*?<xdr:row>5<\/xdr:row>/);
  const contentTypes = partIn(parts, '[Content_Types].xml');
  assert.match(contentTypes, /Extension="png" ContentType="image\/png"/);
  assert.match(contentTypes, /PartName="\/xl\/drawings\/drawing1\.xml"/);
});

test('an image-free workbook writes no drawing or media parts', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'plain';
  const names = Object.keys(partsWritten(wb));
  assert.ok(!names.some((n) => /drawing\d+\.xml$/.test(n)));
  assert.ok(!names.some((n) => n.startsWith('xl/media/')));
});

test('inserting a row above an anchored image shifts its anchor down a row', () => {
  const wb = anchored();
  wb.getWorksheet('S')?.spliceRows(1, 0, ['inserted']);
  const back = roundtrip(wb).getWorksheet('S');
  const anchor = back?.images[0]?.anchor;
  assert.ok(anchor && !isOneCellAnchor(anchor));
  assert.strictEqual(anchor.from.row, 6, 'from-anchor row 5 shifts to 6');
  assert.strictEqual(anchor.to.row, 9, 'to-anchor row 8 shifts to 9');
  assert.strictEqual(anchor.from.col, 0, 'columns are untouched by a row splice');
});

test('a one-cell anchor emits a oneCellAnchor with a pixel extent converted to EMUs', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addImage(id, {tl: {col: 2, row: 3}, ext: {width: 191, height: 47}});
  const parts = partsWritten(wb);
  const drawing = partIn(parts, 'xl/drawings/drawing1.xml');
  assert.match(drawing, /<xdr:oneCellAnchor>/);
  assert.doesNotMatch(drawing, /<xdr:to>/, 'a one-cell anchor has no bottom-right point');
  assert.match(drawing, new RegExp(`<xdr:ext cx="${191 * 9525}" cy="${47 * 9525}"/>`));
});

test('a two-cell anchor honors its editAs mode and defaults to oneCell', () => {
  const wb = new Workbook();
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  wb.addWorksheet('A').addImage(id, {
    tl: {col: 1, row: 1},
    br: {col: 4, row: 6},
    editAs: 'absolute',
  });
  wb.addWorksheet('B').addImage(id, {tl: {col: 1, row: 1}, br: {col: 4, row: 6}});
  const parts = partsWritten(wb);
  assert.match(partIn(parts, 'xl/drawings/drawing1.xml'), /<xdr:twoCellAnchor editAs="absolute">/);
  assert.match(partIn(parts, 'xl/drawings/drawing2.xml'), /<xdr:twoCellAnchor editAs="oneCell">/);
});

// A written package whose one drawing part `edit` rewrites, beside the rewritten drawing's text.
function withDrawing(edit: (xml: string) => string): {pkg: Uint8Array; drawing: string} {
  const written = writeXlsx(anchored());
  const drawing = edit(partText(written, 'xl/drawings/drawing1.xml'));
  return {pkg: patchParts(written, {'xl/drawings/drawing1.xml': () => drawing}), drawing};
}

// The schema default for a two-cell anchor's `editAs` is `twoCell`, which is what a file omitting the
// attribute means. Read as absent, it was written back with the authoring default, `oneCell`, and a
// picture that resized with its cells only moved with them after a save.
test('a two-cell anchor a file wrote without editAs reads as twoCell and is written back so', () => {
  const {pkg} = withDrawing((xml) => xml.replace(' editAs="oneCell"', ''));
  const read = readXlsx(pkg);
  const anchor = read.getWorksheet('S')?.images[0]?.anchor;
  assert.ok(anchor && !isOneCellAnchor(anchor));
  assert.equal(anchor.editAs, 'twoCell');
  assert.match(
    partText(writeXlsx(read), 'xl/drawings/drawing1.xml'),
    /<xdr:twoCellAnchor editAs="twoCell">/,
  );
});

// A drawing holding anything the image model cannot write back is kept whole. The list of such content
// missed an absolute anchor and a linked picture, so a drawing holding either beside a modeled picture
// was written back from that picture alone.
test('a drawing with an absolutely anchored picture beside a modeled one is preserved whole', () => {
  const absolute =
    '<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="9525" cy="9525"/>' +
    '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="9" name="Absolute"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
    '<xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
    '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>' +
    '<xdr:clientData/></xdr:absoluteAnchor>';
  const {pkg, drawing} = withDrawing((xml) => xml.replace('</xdr:wsDr>', `${absolute}</xdr:wsDr>`));

  const read = readXlsx(pkg);
  assert.equal(read.getWorksheet('S')?.images.length, 0, 'no picture is modeled from it');
  assert.equal(partMatching(writeXlsx(read), /^xl\/drawings\/drawing\d+\.xml$/), drawing);
});

test('a drawing with a linked picture beside an embedded one is preserved whole', () => {
  const {pkg, drawing} = withDrawing((xml) => {
    const anchor = xml.match(/<xdr:twoCellAnchor\b[\s\S]*<\/xdr:twoCellAnchor>/)?.[0] ?? '';
    const linked = anchor.replace('r:embed="rId1"', 'r:link="rId1"');
    return xml.replace('</xdr:wsDr>', `${linked}</xdr:wsDr>`);
  });

  const read = readXlsx(pkg);
  assert.equal(read.getWorksheet('S')?.images.length, 0, 'no picture is modeled from it');
  assert.equal(partMatching(writeXlsx(read), /^xl\/drawings\/drawing\d+\.xml$/), drawing);
});

test('an anchored picture carries no absolute spPr transform that would override the anchor', () => {
  const parts = partsWritten(anchored());
  const drawing = partIn(parts, 'xl/drawings/drawing1.xml');
  assert.doesNotMatch(drawing, /<a:xfrm/, 'the geometry is the anchor, not a zeroed transform');
});

test('a one-cell anchor round-trips through the reader as an extent, not a to-point', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addImage(id, {tl: {col: 2, row: 3}, ext: {width: 191, height: 47}});
  const anchor = roundtrip(wb).getWorksheet('S')?.images[0]?.anchor;
  assert.ok(anchor && isOneCellAnchor(anchor));
  assert.deepStrictEqual(anchor.from, {col: 2, row: 3, colOff: 0, rowOff: 0});
  assert.deepStrictEqual(anchor.ext, {cx: 191 * 9525, cy: 47 * 9525});
});

test('a dirty or missing image extension is sanitised to a well-formed media name', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const dirty = wb.addImage({buffer: ONE_PX_PNG, extension: '.png?alt=media&token=abc'});
  const missing = wb.addImage({buffer: ONE_PX_PNG});
  ws.addImage(dirty, {tl: {col: 0, row: 0}, br: {col: 1, row: 1}});
  ws.addImage(missing, {tl: {col: 2, row: 2}, br: {col: 3, row: 3}});
  const parts = partsWritten(wb);
  const media = Object.keys(parts).filter((n) => n.startsWith('xl/media/'));
  assert.deepStrictEqual(media.sort(), ['xl/media/image1.png', 'xl/media/image2.png']);
  const contentTypes = partIn(parts, '[Content_Types].xml');
  const defaults = [...contentTypes.matchAll(/<Default Extension="([^"]*)"/g)].map((m) => m[1]);
  assert.ok(
    defaults.every((e) => /^[A-Za-z0-9]+$/.test(e as string)),
    'every Default extension is a bare token',
  );
  assert.doesNotMatch(contentTypes, /image\/undefined/, 'no bogus media type');
  assert.strictEqual(roundtrip(wb).getWorksheet('S')?.images.length, 2);
});

test('removeImage drops exactly the targeted anchor and omits its now-orphaned media', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const id1 = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  const id2 = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addImage(id1, {tl: {col: 0, row: 0}, br: {col: 1, row: 1}});
  ws.addImage(id2, {tl: {col: 2, row: 0}, br: {col: 3, row: 1}});
  ws.removeImage(id1);
  assert.deepStrictEqual(
    ws.images.map((i) => i.imageId),
    [id2],
  );
  const media = Object.keys(partsWritten(wb)).filter((n) => n.startsWith('xl/media/'));
  assert.strictEqual(media.length, 1, 'the orphaned image is not written');
});

test('a fractional anchor floors to its cell with a sub-cell offset scaled by column width', () => {
  const fromCol = (width: number): {col: number; colOff: number} => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getColumn(4).width = width;
    const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
    ws.addImage(id, {tl: {col: 3.5, row: 0}, br: {col: 5, row: 2}});
    const anchor = ws.images[0]?.anchor;
    assert.ok(anchor && !isOneCellAnchor(anchor));
    return {col: anchor.from.col, colOff: anchor.from.colOff ?? 0};
  };
  const narrow = fromCol(5);
  const wide = fromCol(50);
  assert.strictEqual(narrow.col, 3, 'col 3.5 floors to cell column 3');
  assert.ok(
    narrow.colOff > 0 && wide.colOff > narrow.colOff,
    'a wider column yields a larger offset',
  );
});

test('a picture rotation survives the write/read round-trip on a rot-only transform', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addImageAnchor(id, {from: {col: 1, row: 1}, to: {col: 3, row: 3}, rotation: 2700000});
  const drawing = partText(writeXlsx(wb), 'xl/drawings/drawing1.xml');
  assert.match(drawing, /<a:xfrm rot="2700000"\/>/);
  assert.doesNotMatch(drawing, /<a:off|<a:ext/, 'the rot rides alone, no zeroed offset/extent');
  assert.strictEqual(roundtrip(wb).getWorksheet('S')?.images[0]?.anchor.rotation, 2700000);
});

test('a sheet background image writes a <picture>, an image relationship, and its media, and round-trips', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addBackgroundImage(id);
  const parts = partsWritten(wb);
  const sheetXml = partIn(parts, 'xl/worksheets/sheet1.xml');
  const picture = sheetXml.match(/<picture r:id="([^"]+)"\/>/);
  assert.ok(picture, 'the sheet references a background picture');
  const relsXml = partIn(parts, 'xl/worksheets/_rels/sheet1.xml.rels');
  assert.match(
    relsXml,
    new RegExp(
      `<Relationship Id="${picture[1]}"[^>]*Type="[^"]*/image"[^>]*Target="\\.\\./media/image1\\.png"`,
    ),
  );
  assert.ok(parts['xl/media/image1.png'], 'the background bytes are written once');
  const back = roundtrip(wb).getWorksheet('S');
  assert.strictEqual(
    back?.backgroundImageId !== undefined,
    true,
    'the background survives the round-trip',
  );
  assert.deepStrictEqual(
    back && roundtrip(wb).getImage(back.backgroundImageId ?? -1)?.data,
    ONE_PX_PNG,
  );
});

test('a background image, a note, and an anchored image on one sheet keep unique relationship ids', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('B2').note = 'a note';
  const anchoredId = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addImage(anchoredId, {tl: {col: 0, row: 0}, br: {col: 1, row: 1}});
  const bgId = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addBackgroundImage(bgId);
  const parts = partsWritten(wb);
  const relsXml = partIn(parts, 'xl/worksheets/_rels/sheet1.xml.rels');
  const ids = [...relsXml.matchAll(/<Relationship Id="([^"]+)"/g)].map((m) => m[1]);
  assert.strictEqual(new Set(ids).size, ids.length, 'no two worksheet relationships share an id');
  // The note (comments + VML), the drawing, and the background all resolve to distinct part classes.
  const types = [...relsXml.matchAll(/Type="[^"]*\/(\w+)"/g)].map((m) => m[1]);
  assert.ok(types.includes('image'), 'the background rides an image relationship');
  assert.ok(types.includes('drawing'), 'the anchored image rides a drawing relationship');
  assert.ok(
    types.includes('comments') && types.includes('vmlDrawing'),
    'the note rides comments + VML',
  );
});

test('a background image id the workbook never registered is refused by sheet and role', () => {
  const wb = new Workbook();
  wb.addWorksheet('Sales').addBackgroundImage(999);
  assert.throws(
    () => writeXlsx(wb),
    // The message names the sheet and says which of the two roles the id was used in: a workbook
    // whose sheets both anchor and background images is otherwise a search for which one is broken.
    {
      message:
        'sheet "Sales" sets background image id 999, which is not registered on the workbook',
    },
  );
});

test('removeBackgroundImage clears the background and omits its now-orphaned media', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addBackgroundImage(id);
  ws.removeBackgroundImage();
  assert.strictEqual(ws.backgroundImageId, undefined);
  const parts = partsWritten(wb);
  assert.ok(
    !Object.keys(parts).some((n) => n.startsWith('xl/media/')),
    'the orphaned background is not written',
  );
  assert.doesNotMatch(partIn(parts, 'xl/worksheets/sheet1.xml'), /<picture\b/);
});

test('one image anchored on two sheets is stored as a single media part', () => {
  const wb = new Workbook();
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  wb.addWorksheet('A').addImage(id, {tl: {col: 0, row: 0}, br: {col: 1, row: 1}});
  wb.addWorksheet('B').addImage(id, {tl: {col: 3, row: 3}, br: {col: 4, row: 4}});
  const parts = partsWritten(wb);
  const mediaParts = Object.keys(parts).filter((n) => n.startsWith('xl/media/'));
  assert.strictEqual(mediaParts.length, 1, 'the shared image is written once');
});

test('imageContentType falls back to a string for an extension naming an Object.prototype member', () => {
  // The extension comes off a media part's name, so the table it indexes must not answer with
  // an inherited function: that value goes straight into a [Content_Types].xml attribute.
  const type = imageContentType('constructor');
  assert.strictEqual(typeof type, 'string');
  assert.strictEqual(type, 'image/constructor');
  assert.strictEqual(imageContentType('PNG'), 'image/png', 'the known table still answers');
});
