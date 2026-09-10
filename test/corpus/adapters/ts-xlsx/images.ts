// Images and their drawing anchors: placement, enumeration, removal, and what happens to an
// anchor when the rows around it move.

import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {
  type PartMap,
  partBytesOf,
  partMapOf,
  partNamesOf,
  partOf,
  roundtrip,
} from './package-facts.ts';
import {fixtureBytes, readFixture, readXlsx, Workbook, writeXlsx} from './runtime.ts';
import {anchorSpecImage, buildFrom, ONE_PX_PNG} from './spec-model.ts';
import {
  attrsOf,
  hexBytes,
  imageXmlWellFormed,
  parseAnchorSide,
  reloadPatched,
} from './xml-probes.ts';

export const images = {
  // Build a workbook whose sheets place images at the spec's ranges, write it, and report the
  // serialized drawing-anchor geometry (type, editAs, from/to, one-cell extent, spPr transform) as
  // plain numbers: the surface a case asserts against for anchor correctness.
  inspectImageAnchors(spec: Untyped) {
    const parts = partMapOf(writeXlsx(buildFrom(spec)));
    const drawingParts = Object.keys(parts)
      .filter((f) => /^xl\/drawings\/drawing\d+\.xml$/.test(f))
      .sort();
    const anchors = [];
    let xmlOk = true;
    for (const p of drawingParts) {
      const xml = parts[p]!;
      if (!imageXmlWellFormed(xml)) xmlOk = false;
      for (const m of xml.matchAll(
        /<xdr:(oneCellAnchor|twoCellAnchor)\b([^>]*)>([\s\S]*?)<\/xdr:\1>/g,
      )) {
        const body = m[3]!;
        const fromBlock = (body.match(/<xdr:from>([\s\S]*?)<\/xdr:from>/) || [])[1];
        const toBlock = (body.match(/<xdr:to>([\s\S]*?)<\/xdr:to>/) || [])[1];
        const extTag = body.match(/<xdr:ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);
        const editAs = (m[2]!.match(/editAs="([^"]*)"/) || [])[1] || null;
        const sppr = (body.match(/<xdr:spPr\b[\s\S]*?<\/xdr:spPr>/) || [])[0] || '';
        const offTag = sppr.match(/<a:off\b[^>]*x="(-?\d+)"[^>]*y="(-?\d+)"/);
        const spExtTag = sppr.match(/<a:ext\b[^>]*cx="(-?\d+)"[^>]*cy="(-?\d+)"/);
        const off = offTag ? {x: Number(offTag[1]), y: Number(offTag[2])} : null;
        const spExt = spExtTag ? {cx: Number(spExtTag[1]), cy: Number(spExtTag[2])} : null;
        anchors.push({
          anchorType: m[1] === 'oneCellAnchor' ? 'oneCell' : 'twoCell',
          editAs,
          from: parseAnchorSide(fromBlock),
          to: parseAnchorSide(toBlock),
          ext: extTag ? {cx: Number(extTag[1]), cy: Number(extTag[2])} : null,
          spPr: {
            hasXfrm: /<a:xfrm\b/.test(sppr),
            off,
            ext: spExt,
            zeroedTransform: !!(
              off &&
              spExt &&
              off.x === 0 &&
              off.y === 0 &&
              spExt.cx === 0 &&
              spExt.cy === 0
            ),
          },
        });
      }
    }
    return {anchors, drawingCount: drawingParts.length, xmlWellFormed: xmlOk};
  },

  // Anchor two images to single-cell ranges (C2, C3), interleaving cell writes, and report each
  // anchor's serialized from col/row → { anchorCount, froms }. A cell-range anchor must resolve to
  // that exact cell with no off-by-one row drift.
  cellAnchoredImagePositionReport() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'r1';
    anchorSpecImage(ws, wb.addImage({buffer: ONE_PX_PNG, extension: 'png'}), 'C2:C2');
    ws.getCell('A2').value = 'r3';
    anchorSpecImage(ws, wb.addImage({buffer: ONE_PX_PNG, extension: 'png'}), 'C3:C3');
    const drawingXml = partMapOf(writeXlsx(wb))['xl/drawings/drawing1.xml'] || '';
    const froms = [
      ...drawingXml.matchAll(
        /<xdr:from>\s*<xdr:col>(\d+)<\/xdr:col>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/g,
      ),
    ].map((m) => ({
      col: Number(m[1]),
      row: Number(m[2]),
    }));
    return {anchorCount: froms.length, froms};
  },

  // Anchor a two-cell and a one-cell image, round-trip, and report each read-back image's top-left
  // cell and whether its media survived → { count, images, mediaCount }.
  enumerateImagesAfterRoundtrip() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.addImage(wb.addImage({buffer: ONE_PX_PNG, extension: 'png'}), {
      tl: {col: 1, row: 1},
      br: {col: 3, row: 3},
    });
    ws.addImage(wb.addImage({buffer: ONE_PX_PNG, extension: 'png'}), {
      tl: {col: 5, row: 5},
      ext: {width: 50, height: 50},
    });
    const reread = roundtrip(wb);
    const images = (reread.getWorksheet('S')?.images || []).map((im) => {
      const from = im.anchor.from;
      return {
        tl: from ? {col: from.col, row: from.row} : null,
        hasMedia: !!reread.getImage(im.imageId),
      };
    });
    return {count: images.length, images, mediaCount: reread.media.length};
  },

  // Register two DISTINCT images (told apart by byte length) and place them interleaved (default
  // B, A, A). Resolve, per anchor, which media part its embed actually references → so a case can
  // assert every anchor renders the image it was placed with, including a reused image.
  interleavedImageAnchors(placement = 'BAA') {
    const PNG_A = hexBytes(
      '89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da6360000002000001e5273db40000000049454e44ae426082',
    );
    const PNG_B = hexBytes(
      '89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000a49444154789c6300010000050001',
    );
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    const ids = {
      A: wb.addImage({buffer: PNG_A, extension: 'png'}),
      B: wb.addImage({buffer: PNG_B, extension: 'png'}),
    };
    const placed = placement.split('');
    placed.forEach((letter, i) => {
      const col = i * 2;
      sheet.addImage((ids as Record<string, number>)[letter]!, {
        tl: {col, row: 0},
        br: {col: col + 2, row: 2},
      });
    });
    const buffer = writeXlsx(wb);
    const relsXml = partOf(buffer, 'xl/drawings/_rels/drawing1.xml.rels');
    const relTarget: Record<string, string | undefined> = {};
    for (const t of relsXml.matchAll(/<Relationship\b[^>]*\/?>/g)) {
      const a = attrsOf(t[0]);
      relTarget[a.Id!] = (a.Target || '').split('/').pop();
    }
    const drawingXml = partOf(buffer, 'xl/drawings/drawing1.xml');
    const embedOrder = [...drawingXml.matchAll(/r:embed="([^"]*)"/g)].map((m) => m[1]);
    const resolvedMedia = embedOrder.map((rid) => relTarget[rid!] ?? null);
    const mediaSizes: Record<string, number> = {};
    for (const name of partNamesOf(buffer)) {
      const m = name.match(/^xl\/media\/(image\d+\.png)$/);
      if (m) mediaSizes[m[1]!] = partBytesOf(buffer, name).length;
    }
    const resolvedLetter = resolvedMedia.map((media) => {
      const size = mediaSizes[media!];
      if (size === PNG_A.length) return 'A';
      if (size === PNG_B.length) return 'B';
      return '?';
    });
    return {
      placed,
      embedOrder,
      resolvedMedia,
      resolvedLetter,
      distinctMediaCount: Object.keys(mediaSizes).length,
      distinctRelTargets: new Set(Object.values(relTarget)).size,
    };
  },

  // Carry an image-bearing sheet into a DIFFERENT workbook, taking the model for its content and
  // exportImages/importImages for its pictures, then write the destination and report whether it
  // arrived → { dstMediaCount, allEmbedsResolve, reReadImageCount, reReadAnchor, reReadHasBackground,
  // srcStillShowsImage, danglingAnchorError }. A drawing whose embed resolves to no media part is the
  // silently-broken-image failure; `danglingAnchorError` is the other leg: an anchor carrying a
  // foreign workbook's media id must be refused loudly at write time, never emitted as a dangling rel.
  carrySheetImagesAcrossWorkbooks() {
    const source = new Workbook();
    const src = source.addWorksheet('Src');
    src.getCell('A1').value = 'content';
    const logo = source.addImage({buffer: ONE_PX_PNG, extension: 'png'});
    src.addImage(logo, {tl: {col: 1, row: 1}, br: {col: 3, row: 4}});
    src.addBackgroundImage(source.addImage({buffer: hexBytes('474946383961'), extension: 'gif'}));

    const destination = new Workbook();
    const dst = destination.addWorksheet('Dst');
    dst.model = src.model;
    destination.importImages(dst, source.exportImages(src));

    const buffer = writeXlsx(destination);
    const names = partNamesOf(buffer);
    const mediaParts = new Set(names.filter((f) => f.startsWith('xl/media/')));
    const relTarget: Record<string, string | undefined> = {};
    for (const t of partOf(buffer, 'xl/drawings/_rels/drawing1.xml.rels').matchAll(
      /<Relationship\b[^>]*\/?>/g,
    )) {
      const a = attrsOf(t[0]);
      relTarget[a.Id!] = (a.Target || '').replace(/^\.\.\//, 'xl/');
    }
    const embeds = [
      ...partOf(buffer, 'xl/drawings/drawing1.xml').matchAll(/r:embed="([^"]*)"/g),
    ].map((m) => m[1]!);

    const back = readXlsx(buffer).getWorksheet('Dst')!;
    const anchor = back.images[0]?.anchor as Untyped;

    // The other leg of the contract: an anchor holding an id the destination never registered.
    let danglingAnchorError: string | null = null;
    const orphaned = new Workbook();
    orphaned.addWorksheet('O').addImageAnchor(41, {from: {col: 0, row: 0}, ext: {cx: 1, cy: 1}});
    try {
      writeXlsx(orphaned);
    } catch (e) {
      danglingAnchorError = messageOf(e);
    }

    return {
      dstMediaCount: mediaParts.size,
      allEmbedsResolve:
        embeds.length > 0 && embeds.every((id) => mediaParts.has(relTarget[id] ?? '')),
      reReadImageCount: back.images.length,
      reReadAnchor: anchor?.from ? {col: anchor.from.col, row: anchor.from.row} : null,
      reReadHasBackground: back.backgroundImageId !== undefined,
      srcStillShowsImage: src.images.length === 1,
      danglingAnchorError,
    };
  },

  // Load a workbook from bytes, register and anchor an image on the loaded worksheet, re-serialize,
  // and report the media/drawing presence and re-read image count → locks that adding an image to a
  // *loaded* (not freshly created) worksheet persists.
  addImageToLoadedWorksheetReport(range = 'B2:C4') {
    const base = new Workbook();
    base.addWorksheet('S').getCell('A1').value = 'x';
    const loaded = roundtrip(base);
    anchorSpecImage(
      loaded.getWorksheet('S'),
      loaded.addImage({buffer: ONE_PX_PNG, extension: 'png'}),
      range,
    );
    const out = writeXlsx(loaded);
    const files = Object.keys(partMapOf(out));
    const reloadImages = readXlsx(out).getWorksheet('S')?.images || [];
    return {
      hasMedia: files.some((f) => /xl\/media\//.test(f)),
      hasDrawing: files.some((f) => /xl\/drawings\/drawing/.test(f)),
      reloadImageCount: reloadImages.length,
    };
  },

  // Read a fixture and report each image's normalized anchor range → for asserting a file whose
  // drawing anchors were authored as cell ranges reads without throwing and exposes integer cell
  // coordinates, never a raw string.
  readFixtureImageAnchors(rel: string) {
    const workbook = readFixture(rel);
    const images = [];
    for (const sheet of workbook.worksheets) {
      for (const im of sheet.images) {
        const from = im.anchor.from;
        // reaches past the public ImageAnchor type for the two-cell `to` side
        const to = (im.anchor as Untyped).to;
        images.push({
          sheet: sheet.name,
          // reaches past the public ImageAnchor type for `editAs`
          editAs: (im.anchor as Untyped).editAs ?? null,
          tl: from ? {col: from.col, row: from.row} : null,
          br: to ? {col: to.col, row: to.row} : null,
        });
      }
    }
    return {images, count: images.length};
  },

  // Add one image whose extension may carry a leading dot or a query string, write, and report the
  // media filenames and re-read image count → a dirty extension must sanitise to a well-formed media
  // name the reader still recognises, not a doubled-separator name that drops the image.
  imageExtensionRoundtrip(extension = 'png') {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    const id = wb.addImage({buffer: ONE_PX_PNG, extension});
    ws.addImage(id, {tl: {col: 1, row: 1}, br: {col: 3, row: 3}});
    const buffer = writeXlsx(wb);
    const mediaParts = Object.keys(partMapOf(buffer))
      .filter((n) => /^xl\/media\/.+/.test(n))
      .map((n) => n.replace(/^xl\/media\//, ''));
    const images = readXlsx(buffer).getWorksheet('S')?.images || [];
    return {
      mediaParts,
      doubledSeparator: mediaParts.some((n) => /\.\./.test(n)),
      reloadedImageCount: images.length,
    };
  },

  // Anchor two images, then remove one by its media id → { supported, before, after, removedGone,
  // othersSurvive }. Removal must drop exactly the targeted image and leave the rest anchored.
  removeImageReport(range = 'A1:B2') {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    const id1 = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
    const id2 = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
    anchorSpecImage(ws, id1, range);
    anchorSpecImage(ws, id2, 'C1:D2');
    const supported = typeof ws.removeImage === 'function';
    const before = ws.images.length;
    let after = before;
    let removedGone = false;
    let othersSurvive = false;
    if (supported) {
      ws.removeImage(id1);
      const ids = ws.images.map((i) => i.imageId);
      after = ws.images.length;
      removedGone = !ids.includes(id1);
      othersSurvive = ids.includes(id2);
    }
    return {supported, before, after, removedGone, othersSurvive};
  },

  // Write sheet `S` with one picture, put a vector shape into its drawing, and read it back, so the
  // drawing is kept whole with no picture modelled. Then try to anchor a picture there → { keptShape,
  // imagesModeled, addError, shapeAfterAttempt }: whether a plain save keeps the shape, how many
  // pictures the read modelled, the add's error message or null, and whether a save after the
  // attempt still carries the shape. "Carries" means in the drawing the saved sheet's `<drawing>`
  // element references: a drawing part left in the package with nothing pointing at it is lost too.
  pictureBesideKeptDrawingReport() {
    const shape =
      '<xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row>' +
      '<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff>' +
      '<xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro="" textlink="">' +
      '<xdr:nvSpPr><xdr:cNvPr id="99" name="Rectangle 1"/><xdr:cNvSpPr/></xdr:nvSpPr>' +
      '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>' +
      '<xdr:clientData/></xdr:twoCellAnchor>';
    const source = new Workbook();
    const pictured = source.addWorksheet('S');
    anchorSpecImage(pictured, source.addImage({buffer: ONE_PX_PNG, extension: 'png'}), 'A1:B2');
    const read = reloadPatched(writeXlsx(source), {
      'xl/drawings/drawing1.xml': (xml) => xml.replace('</xdr:wsDr>', `${shape}</xdr:wsDr>`),
    });
    const hasShape = (): boolean => {
      const parts = partMapOf(writeXlsx(read));
      const relId = attrsOf(
        (parts['xl/worksheets/sheet1.xml'] ?? '').match(/<drawing\b[^>]*\/>/)?.[0] ?? '<drawing/>',
      )['r:id'];
      const relationship = [
        ...(parts['xl/worksheets/_rels/sheet1.xml.rels'] ?? '').matchAll(
          /<Relationship\b[^>]*\/>/g,
        ),
      ]
        .map(([element]) => attrsOf(element))
        .find((attrs) => attrs.Id === relId);
      const target = relationship?.Target?.replace(/^\.\.\//, 'xl/');
      return target !== undefined && /<xdr:sp\b/.test(parts[target] ?? '');
    };
    const sheet = read.getWorksheet('S');
    if (sheet === undefined) throw new Error('the patched package lost sheet S');
    const keptShape = hasShape();
    const imagesModeled = sheet.images.length;
    let addError: string | null = null;
    try {
      anchorSpecImage(sheet, read.addImage({buffer: ONE_PX_PNG, extension: 'png'}), 'D1:E2');
    } catch (error) {
      addError = messageOf(error);
    }
    return {keptShape, imagesModeled, addError, shapeAfterAttempt: hasShape()};
  },

  // Anchor an image and append rows in both orders → { imageFirst, rowsFirst }, each { rowCount,
  // firstDataCell }. Anchoring a floating image is metadata overlay, not row insertion: it must not
  // advance the row-append cursor, so the layout is identical regardless of add order.
  imageAnchorRowAppendReport() {
    const run = (order: Untyped) => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S');
      const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
      if (order === 'image-first') {
        anchorSpecImage(sheet, id, 'A1:B3');
        sheet.addRows([['a'], ['b'], ['c']]);
      } else {
        sheet.addRows([['a'], ['b'], ['c']]);
        anchorSpecImage(sheet, id, 'A1:B3');
      }
      return {rowCount: sheet.rowCount, firstDataCell: sheet.getCell('A1').value ?? null};
    };
    return {imageFirst: run('image-first'), rowsFirst: run('rows-first')};
  },

  // Read a fixture, load-and-rewrite it, and report the picture's drawing-anchor rotation before and
  // after → { sourceRot, rewrittenRot }. An image rotation (rot on <a:xfrm>) must survive the round-trip.
  roundtripFixtureImageRotation(rel: string) {
    const rotOf = (xml: string) => {
      const m = xml.match(/<a:xfrm\b[^>]*\brot="(-?\d+)"/);
      return m ? Number(m[1]) : null;
    };
    const drawingXml = (parts: PartMap) => {
      const name = Object.keys(parts).find((f) => /^xl\/drawings\/drawing\d+\.xml$/.test(f));
      return name === undefined ? null : (parts[name] ?? null);
    };
    const srcDrawing = drawingXml(partMapOf(fixtureBytes(rel)));
    const sourceRot = srcDrawing === null ? null : rotOf(srcDrawing);
    const outDrawing = drawingXml(partMapOf(writeXlsx(readXlsx(fixtureBytes(rel)))));
    const rewrittenRot = outDrawing === null ? null : rotOf(outDrawing);
    return {sourceRot, rewrittenRot};
  },
};
