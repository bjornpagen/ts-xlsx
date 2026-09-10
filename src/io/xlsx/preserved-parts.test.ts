import {strict as assert} from 'node:assert';
import {test} from 'node:test';

import {INTERNAL} from '../../core/internal.ts';
import {Workbook} from '../../core/workbook.ts';
import {AuthoringError} from '../../errors.ts';
import {
  msRelationship,
  partBytes,
  partMatching,
  partsOf,
  partText,
  relationship,
  relationshipsPart as rels,
  SHEET1,
  typedContentTypes as contentTypes,
  typedForeignPackage,
  typedWorksheet as worksheet,
} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

// Each scenario overlays the worksheet, its rels, and the unmodeled parts (a shape drawing, a
// header/footer VML + image) that exercise the passthrough onto a package whose relationships carry
// their real types, since the type is what marks a part for preservation.

const SHAPE_DRAWING =
  '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor>' +
  '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
  '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
  '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="Rectangle 1"/><xdr:cNvSpPr/></xdr:nvSpPr>' +
  '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>' +
  '<xdr:txBody><a:bodyPr/><a:p/></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>';

// A single drawing part holding BOTH a modeled picture (an <xdr:pic> with a blip embed) AND a chart
// (an <xdr:graphicFrame> naming a chart part by r:id): the shape Excel produces for a sheet that
// carries an image and a chart together. Modeling only the picture and re-serialising the drawing
// from it would drop the chart; the whole part must ride through preservation instead.
const MIXED_DRAWING =
  '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<xdr:oneCellAnchor>' +
  '<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
  '<xdr:ext cx="100" cy="100"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="p"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
  '<xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/>' +
  '</xdr:oneCellAnchor>' +
  '<xdr:twoCellAnchor>' +
  '<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
  '<xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
  '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
  '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/>' +
  '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>' +
  '</xdr:twoCellAnchor></xdr:wsDr>';

const CHART =
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart/></c:plotArea></c:chart></c:chartSpace>';

const HF_VML =
  '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">' +
  '<v:shape id="RH" type="#_x0000_t75"><v:imagedata o:relid="rId1" o:title="pic"/></v:shape></xml>';

const WORKBOOK_RELS = 'xl/_rels/workbook.xml.rels';
const SHEET1_RELS = 'xl/worksheets/_rels/sheet1.xml.rels';

// A picture's bytes. Nothing here decodes them, so a PNG signature is enough.
const PNG = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);
const PICTURE_AT_A1 = {tl: {col: 0, row: 0}, ext: {width: 10, height: 10}};

function partNames(pkg: Uint8Array): string[] {
  return Object.keys(partsOf(pkg));
}

// Sheet `S`, whose one drawing holds a vector shape and nothing the model draws.
function shapeDrawingPackage(): Uint8Array {
  return typedForeignPackage({
    '[Content_Types].xml': contentTypes(
      '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>',
    ),
    [SHEET1]: worksheet('<drawing r:id="rId1"/>'),
    [SHEET1_RELS]: rels(relationship('rId1', 'drawing', '../drawings/drawing1.xml')),
    'xl/drawings/drawing1.xml': SHAPE_DRAWING,
  });
}

test('a worksheet drawing holding only a vector shape survives read→write', () => {
  const out = writeXlsx(readXlsx(shapeDrawingPackage()));

  assert.ok(
    partNames(out).some((n) => /xl\/drawings\/drawing\d+\.xml$/.test(n)),
    'the drawing part survives',
  );
  assert.match(
    partText(out, SHEET1),
    /<drawing r:id="[^"]+"\/>/,
    'the worksheet still references it',
  );
  assert.match(
    partText(out, 'xl/drawings/drawing1.xml'),
    /<xdr:sp\b/,
    'the vector shape survives inside',
  );
  // The rewritten package must re-read without error, and re-writing it must keep preserving the shape.
  assert.match(
    partText(writeXlsx(readXlsx(out)), 'xl/drawings/drawing1.xml'),
    /<xdr:sp\b/,
    'idempotent across a second round-trip',
  );
});

// A worksheet references one drawing. A kept one and a picture the model draws would need two, and the
// writer referenced only the new one, so the chart or shape was dropped from the file without a word.
test('a picture cannot be added beside a kept drawing, and the refusal leaves the sheet as it was', () => {
  const workbook = readXlsx(shapeDrawingPackage());
  const sheet = workbook.requireWorksheet('S');
  const refusal = (error: unknown): boolean =>
    error instanceof AuthoringError &&
    error.message.includes('sheet "S"') &&
    error.message.includes('does not model');
  const id = workbook.addImage({buffer: PNG, extension: 'png'});
  assert.throws(() => sheet.addImage(id, PICTURE_AT_A1), refusal);

  // An import replaces a sheet's pictures and background, so it must refuse before clearing either.
  sheet.addBackgroundImage(id);
  const source = new Workbook();
  const pictured = source.addWorksheet('P');
  pictured.addImage(source.addImage({buffer: PNG, extension: 'png'}), PICTURE_AT_A1);
  assert.throws(() => workbook.importImages(sheet, source.exportImages(pictured)), refusal);
  assert.equal(sheet.backgroundImageId, id, 'the refused import cleared nothing');

  assert.equal(sheet.images.length, 0);
  assert.match(partText(writeXlsx(workbook), 'xl/drawings/drawing1.xml'), /<xdr:sp\b/);
});

test('the writer refuses a picture beside a kept drawing, however the two came to share a sheet', () => {
  const kept = readXlsx(shapeDrawingPackage())
    .requireWorksheet('S')
    .preservedReferences.find((reference) => reference.element === 'drawing');
  assert.ok(kept);
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('S');
  sheet.addImage(workbook.addImage({buffer: PNG, extension: 'png'}), PICTURE_AT_A1);
  // Only a codec reaches this channel. It stands in for any path that skips the authoring refusal.
  sheet[INTERNAL].addPreservedReference(kept);
  assert.throws(() => writeXlsx(workbook), AuthoringError);
});

test('a drawing holding both a picture and a chart preserves the chart across read→write', () => {
  const png = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);
  const src = typedForeignPackage({
    '[Content_Types].xml': contentTypes(
      '<Default Extension="png" ContentType="image/png"/>' +
        '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
        '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>',
    ),
    [SHEET1]: worksheet('<drawing r:id="rId1"/>'),
    [SHEET1_RELS]: rels(relationship('rId1', 'drawing', '../drawings/drawing1.xml')),
    'xl/drawings/drawing1.xml': MIXED_DRAWING,
    'xl/drawings/_rels/drawing1.xml.rels': rels(
      relationship('rId1', 'image', '../media/image1.png') +
        relationship('rId2', 'chart', '../charts/chart1.xml'),
    ),
    'xl/media/image1.png': png,
    'xl/charts/chart1.xml': CHART,
  });

  const out = writeXlsx(readXlsx(src));
  const names = partNames(out);

  assert.match(
    partText(out, SHEET1),
    /<drawing r:id="[^"]+"\/>/,
    'the worksheet still references the drawing',
  );
  const drawing = partMatching(out, /^xl\/drawings\/drawing\d+\.xml$/);
  assert.match(drawing, /<xdr:graphicFrame\b/, 'the chart anchor survives inside the drawing');
  assert.match(drawing, /<xdr:pic\b/, 'the picture anchor survives alongside it, not dropped');
  assert.ok(
    names.some((n) => /xl\/charts\/chart\d+\.xml$/.test(n)),
    'the chart part survives',
  );
  assert.ok(
    names.some((n) => /xl\/media\/.+\.png$/.test(n)),
    'the picture media survives',
  );
  // The whole mixed drawing rides through preservation, so a second round-trip keeps the chart too.
  assert.ok(
    partNames(writeXlsx(readXlsx(out))).some((n) => /xl\/charts\/chart\d+\.xml$/.test(n)),
    'idempotent across a second round-trip',
  );
});

test('a header/footer image (legacyDrawingHF VML + media) and its &G token survive read→write', () => {
  const image = Uint8Array.from([0xff, 0xd8, 0xff, 0xe0, 1, 2, 3, 4]);
  const src = typedForeignPackage({
    '[Content_Types].xml': contentTypes(
      '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
        '<Default Extension="jpeg" ContentType="image/jpeg"/>',
    ),
    [SHEET1]: worksheet(
      '<headerFooter><oddHeader>&amp;R&amp;G</oddHeader></headerFooter><legacyDrawingHF r:id="rId1"/>',
    ),
    [SHEET1_RELS]: rels(relationship('rId1', 'vmlDrawing', '../drawings/vmlDrawing1.vml')),
    'xl/drawings/vmlDrawing1.vml': HF_VML,
    'xl/drawings/_rels/vmlDrawing1.vml.rels': rels(
      relationship('rId1', 'image', '../media/image1.jpeg'),
    ),
    'xl/media/image1.jpeg': image,
  });

  const out = writeXlsx(readXlsx(src));
  const ws = partText(out, SHEET1);

  assert.match(ws, /<legacyDrawingHF r:id="[^"]+"\/>/, 'the legacyDrawingHF reference survives');
  assert.match(ws, /&amp;R&amp;G/, 'the &G header/footer picture token survives');
  assert.ok(
    partNames(out).some((n) => /vmlDrawing\d+\.vml$/.test(n)),
    'the VML drawing survives',
  );
  assert.ok(
    partNames(out).some((n) => /xl\/media\//.test(n)),
    'the image media survives',
  );
  // The VML's image relationship must re-resolve to the media part's new path.
  assert.match(
    partMatching(out, /^xl\/drawings\/_rels\/vmlDrawing\d+\.vml\.rels$/),
    /Target="\.\.\/media\/image\d+\.jpeg"/,
    'the VML relinks its image',
  );
});

test('a preserved header/footer VML is numbered clear of a modeled anchored image', () => {
  // A sheet that anchors a real image (modeled → drawing1.xml, image1.jpeg, media #1) AND carries a
  // preserved header/footer VML whose own image would also be image1.jpeg in the source. The writer
  // must renumber the preserved media past the modeled one rather than clobbering it.
  const modeledPng = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 9, 9]);
  const hfJpeg = Uint8Array.from([0xff, 0xd8, 0xff, 5, 6]);
  const src = typedForeignPackage({
    '[Content_Types].xml': contentTypes(
      '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
        '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>' +
        '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>',
    ),
    [SHEET1]: worksheet('<drawing r:id="rId1"/><legacyDrawingHF r:id="rId2"/>'),
    [SHEET1_RELS]: rels(
      relationship('rId1', 'drawing', '../drawings/drawing1.xml') +
        relationship('rId2', 'vmlDrawing', '../drawings/vmlDrawing1.vml'),
    ),
    'xl/drawings/drawing1.xml':
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:oneCellAnchor>' +
      '<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
      '<xdr:ext cx="100" cy="100"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="p"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
      '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/></xdr:blipFill>' +
      '<xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>',
    'xl/drawings/_rels/drawing1.xml.rels': rels(
      relationship('rId1', 'image', '../media/image1.png'),
    ),
    'xl/media/image1.png': modeledPng,
    'xl/drawings/vmlDrawing1.vml': HF_VML,
    'xl/drawings/_rels/vmlDrawing1.vml.rels': rels(
      relationship('rId1', 'image', '../media/image1.jpeg'),
    ),
    'xl/media/image1.jpeg': hfJpeg,
  });

  const out = writeXlsx(readXlsx(src));
  const names = partNames(out);
  const media = names.filter((n) => /xl\/media\//.test(n));
  assert.equal(new Set(media).size, media.length, 'no two media parts share a path');
  assert.ok(
    media.some((n) => n.endsWith('.png')) && media.some((n) => n.endsWith('.jpeg')),
    'both the modeled png and the preserved jpeg survive',
  );
  assert.ok(
    names.some((n) => /vmlDrawing\d+\.vml$/.test(n)),
    'the header/footer VML survives',
  );
  assert.match(
    partText(out, SHEET1),
    /<legacyDrawingHF r:id="[^"]+"\/>/,
    'the legacyDrawingHF reference survives',
  );
});

// A workbook overlay that adds a `<pivotCaches>` registration and a workbook relationship reaching a
// pivot cache: the wiring a real pivot-bearing workbook carries, which the base package omits.
const workbookWithPivotCache =
  '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>' +
  '<pivotCaches><pivotCache cacheId="42" r:id="rId2"/></pivotCaches></workbook>';

test('a pivot table and its pivot cache survive read→write, cacheId wiring intact', () => {
  const src = typedForeignPackage({
    '[Content_Types].xml': contentTypes(
      '<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>' +
        '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>' +
        '<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>',
    ),
    'xl/workbook.xml': workbookWithPivotCache,
    [WORKBOOK_RELS]: rels(
      relationship('rId1', 'worksheet', 'worksheets/sheet1.xml') +
        relationship('rId2', 'pivotCacheDefinition', 'pivotCache/pivotCacheDefinition1.xml'),
    ),
    // The pivot table is discovered through a sheet relationship: there is no worksheet child.
    [SHEET1]: worksheet(),
    [SHEET1_RELS]: rels(relationship('rId1', 'pivotTable', '../pivotTables/pivotTable1.xml')),
    'xl/pivotTables/pivotTable1.xml': '<pivotTableDefinition cacheId="42"/>',
    'xl/pivotTables/_rels/pivotTable1.xml.rels': rels(
      relationship('rId1', 'pivotCacheDefinition', '../pivotCache/pivotCacheDefinition1.xml'),
    ),
    'xl/pivotCache/pivotCacheDefinition1.xml': '<pivotCacheDefinition/>',
    'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': rels(
      relationship('rId1', 'pivotCacheRecords', 'pivotCacheRecords1.xml'),
    ),
    'xl/pivotCache/pivotCacheRecords1.xml': '<pivotCacheRecords/>',
  });

  const out = writeXlsx(readXlsx(src));
  const names = partNames(out);

  assert.ok(names.includes('xl/pivotTables/pivotTable1.xml'), 'the pivot table part survives');
  assert.equal(
    names.filter((n) => /xl\/pivotCache\/.+\.xml$/.test(n)).length,
    2,
    'both the pivot cache definition and its records survive',
  );
  assert.match(
    partText(out, SHEET1_RELS),
    /pivotTable/,
    'the sheet still references the pivot table',
  );

  // The <pivotCaches> registration is re-emitted with its cacheId, wired to the workbook relationship
  // that reaches the (surviving) cache definition, so a pivot table can resolve its cache on reopen.
  const wb = partText(out, 'xl/workbook.xml');
  const cache = /<pivotCache cacheId="42" r:id="(rId\d+)"\/>/.exec(wb);
  assert.ok(cache, `workbook registers the pivot cache with its cacheId; got ${wb}`);
  assert.match(
    partText(out, WORKBOOK_RELS),
    new RegExp(`Id="${cache[1]}"[^>]*Target="pivotCache/pivotCacheDefinition1\\.xml"`),
    'the pivotCaches relationship id resolves to the cache definition',
  );

  // Re-reading and re-writing the output keeps everything, so the passthrough is idempotent.
  const again = partNames(writeXlsx(readXlsx(out)));
  assert.ok(
    again.includes('xl/pivotTables/pivotTable1.xml'),
    'idempotent across a second round-trip',
  );
  assert.equal(
    again.filter((n) => /xl\/pivotCache\/.+\.xml$/.test(n)).length,
    2,
    'the cache survives a second round-trip',
  );
});

// A macro-enabled workbook's VBA project and a *digital signature* over it: the signature is a sibling
// package part (`xl/vbaProjectSignature.bin`) reached by a `.../vbaProjectSignature` relationship on the
// vbaProject part's own rels, so the closure walk carries it through. It shares the `.bin` extension
// with `vbaProject.bin` but has a DIFFERENT content type: the case a single per-extension `<Default>`
// mis-types unless the writer emits a per-part `<Override>` for the odd one out.
test('a signed VBA project keeps distinct content types for vbaProject.bin and its signature', () => {
  const vbaBytes = Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3]);
  const sigBytes = Uint8Array.from([0xde, 0xad, 0xbe, 0xef, 4, 5]);
  const src = typedForeignPackage({
    '[Content_Types].xml': contentTypes(
      '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>' +
        '<Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/>',
    ),
    [WORKBOOK_RELS]: rels(
      relationship('rId1', 'worksheet', 'worksheets/sheet1.xml') +
        msRelationship('rId2', '2006/relationships/vbaProject', 'vbaProject.bin'),
    ),
    'xl/vbaProject.bin': vbaBytes,
    'xl/_rels/vbaProject.bin.rels': rels(
      msRelationship('rId1', '2006/relationships/vbaProjectSignature', 'vbaProjectSignature.bin'),
    ),
    'xl/vbaProjectSignature.bin': sigBytes,
  });

  const out = writeXlsx(readXlsx(src));

  assert.deepEqual(partBytes(out, 'xl/vbaProject.bin'), vbaBytes, 'the VBA project blob survives');
  assert.deepEqual(
    partBytes(out, 'xl/vbaProjectSignature.bin'),
    sigBytes,
    'the signature blob survives',
  );

  const ct = partText(out, '[Content_Types].xml');
  // The signature part must keep its own content type: a lone `.bin` Default would mis-type it as a
  // second vbaProject, which Excel reads as a corrupt/duplicate project rather than a signature.
  assert.match(
    ct,
    /PartName="\/xl\/vbaProjectSignature\.bin" ContentType="application\/vnd\.ms-office\.vbaProjectSignature"/,
    'the signature part is typed as a vbaProjectSignature, not collapsed into the vbaProject default',
  );
  // The workbook is still declared macro-enabled and the project part still types as a vbaProject.
  assert.match(ct, /ContentType="application\/vnd\.ms-excel\.sheet\.macroEnabled\.main\+xml"/);
  assert.match(ct, /ContentType="application\/vnd\.ms-office\.vbaProject"/);

  // Idempotent: the corrected typing survives a second round-trip.
  const ct2 = partText(writeXlsx(readXlsx(out)), '[Content_Types].xml');
  assert.match(ct2, /application\/vnd\.ms-office\.vbaProjectSignature/);
});

test('slicer and slicer-cache parts survive read→write', () => {
  const src = typedForeignPackage({
    '[Content_Types].xml': contentTypes(
      '<Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>' +
        '<Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>',
    ),
    [WORKBOOK_RELS]: rels(
      relationship('rId1', 'worksheet', 'worksheets/sheet1.xml') +
        msRelationship('rId2', '2007/relationships/slicerCache', 'slicerCaches/slicerCache1.xml'),
    ),
    [SHEET1_RELS]: rels(
      msRelationship('rId1', '2007/relationships/slicer', '../slicers/slicer1.xml'),
    ),
    'xl/slicers/slicer1.xml': '<slicers/>',
    'xl/slicerCaches/slicerCache1.xml': '<slicerCacheDefinition/>',
  });

  const out = writeXlsx(readXlsx(src));
  const names = partNames(out);

  assert.ok(names.includes('xl/slicers/slicer1.xml'), 'the slicer part survives');
  assert.ok(names.includes('xl/slicerCaches/slicerCache1.xml'), 'the slicer cache part survives');
  assert.match(partText(out, SHEET1_RELS), /\/slicer"/, 'the sheet still references the slicer');
  assert.match(
    partText(out, WORKBOOK_RELS),
    /slicerCaches\/slicerCache1\.xml/,
    'the workbook still references the slicer cache',
  );

  // Parts surviving is not enough: Excel only rediscovers a slicer through its x14 wiring, which
  // references the relationship ids the writer reassigns, so the ext blocks must name the *new* ids.
  const slicerRelId = partText(out, SHEET1_RELS).match(/Id="(rId\d+)"[^>]*\/slicer"/)?.[1];
  assert.ok(slicerRelId, 'the re-emitted slicer rel has an id');
  assert.match(
    partText(out, SHEET1),
    new RegExp(`<x14:slicerList><x14:slicer r:id="${slicerRelId}"/></x14:slicerList>`),
    'the sheet body reactivates the slicer through a slicerList extension wired to its rel',
  );
  const cacheRelId = partText(out, WORKBOOK_RELS).match(/Id="(rId\d+)"[^>]*\/slicerCache"/)?.[1];
  assert.ok(cacheRelId, 'the re-emitted slicer-cache rel has an id');
  assert.match(
    partText(out, 'xl/workbook.xml'),
    new RegExp(`<x14:slicerCaches><x14:slicerCache r:id="${cacheRelId}"/></x14:slicerCaches>`),
    'the workbook registers the slicer cache in its x14 slicerCaches extension',
  );
});

// The ribbon-customisation parts (customUI.xml / customUI14.xml, the buttons a macro workbook adds to
// the ribbon) hang off the *package root* `_rels/.rels`, not the workbook part's rels, so the
// workbook-rels closure walk never reaches them. Preserving them needs the root rels themselves to be
// captured on read and re-declared on write.
const CUSTOM_UI_2007 =
  '<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">' +
  '<ribbon><tabs><tab id="t07" label="Legacy"><group id="g07" label="G">' +
  '<button id="b07" label="Run" onAction="LegacyMacro"/></group></tab></tabs></ribbon></customUI>';

const CUSTOM_UI_2009 =
  '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">' +
  '<ribbon><tabs><tab id="t14" label="Macros"><group id="g14" label="G">' +
  '<button id="b14" label="Run" onAction="MyMacro"/></group></tab></tabs></ribbon></customUI>';

test('customUI ribbon parts referenced from the package root rels survive read→write', () => {
  const src = typedForeignPackage({
    '_rels/.rels': rels(
      relationship('rId1', 'officeDocument', 'xl/workbook.xml') +
        msRelationship('rId4', '2006/relationships/ui/extensibility', 'customUI/customUI.xml') +
        msRelationship('rId5', '2007/relationships/ui/extensibility', 'customUI/customUI14.xml'),
    ),
    [SHEET1_RELS]: rels(''),
    'customUI/customUI.xml': CUSTOM_UI_2007,
    'customUI/customUI14.xml': CUSTOM_UI_2009,
  });

  const out = writeXlsx(readXlsx(src));
  const names = partNames(out);

  assert.ok(names.includes('customUI/customUI.xml'), 'the 2007 ribbon part survives');
  assert.ok(names.includes('customUI/customUI14.xml'), 'the 2009 ribbon part survives');
  assert.match(
    partText(out, 'customUI/customUI.xml'),
    /onAction="LegacyMacro"/,
    'the 2007 ribbon body survives intact',
  );
  assert.match(
    partText(out, 'customUI/customUI14.xml'),
    /onAction="MyMacro"/,
    'the 2009 ribbon body survives intact',
  );

  // Surviving as bytes is not enough: Excel only loads the ribbon through the root-rels
  // relationships, so both must be re-declared there with their Microsoft-namespaced types.
  const rootRels = partText(out, '_rels/.rels');
  assert.match(
    rootRels,
    /office\/2006\/relationships\/ui\/extensibility/,
    'root rels keep the 2007 ribbon relationship',
  );
  assert.match(
    rootRels,
    /office\/2007\/relationships\/ui\/extensibility/,
    'root rels keep the 2010 (customUI14) ribbon relationship',
  );
  assert.match(
    rootRels,
    /Target="customUI\/customUI\.xml"/,
    'the 2007 relationship still targets its part',
  );
  assert.match(
    rootRels,
    /Target="customUI\/customUI14\.xml"/,
    'the 2009 relationship still targets its part',
  );

  // Re-reading the rewritten package and writing it again must keep preserving both.
  assert.match(
    partText(writeXlsx(readXlsx(out)), 'customUI/customUI14.xml'),
    /onAction="MyMacro"/,
    'idempotent across a second round-trip',
  );
});
