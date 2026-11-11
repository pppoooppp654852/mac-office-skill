// Rasterize a PDF and report its text lines, using PDFKit/AppKit through JXA.
// Usage: osascript -l JavaScript pdf_render.js <pdf> <out-dir> <dpi> <png-prefix>
// Output: {"pages":N,"png":["/abs/<prefix>-page-1.png"],"lines":[{"page","text","x","y","w","h"}]}
// Coordinates are PDF points with a top-left origin; PNG pixel size is round(pt * dpi / 72).
ObjC.import('PDFKit');
ObjC.import('AppKit');

function fail(message) {
  throw new Error(message);
}

function run(argv) {
  const pdfPath = argv[0];
  const outDir = argv[1];
  const dpi = Number(argv[2]);
  const prefix = argv[3];
  if (!pdfPath || !outDir || !(dpi > 0) || !prefix) {
    fail('usage: pdf_render.js <pdf> <out-dir> <dpi> <png-prefix>');
  }
  const doc = $.PDFDocument.alloc.initWithURL($.NSURL.fileURLWithPath(pdfPath));
  if (!doc || doc.isNil()) {
    fail('cannot open PDF: ' + pdfPath);
  }
  const pageCount = Number(doc.pageCount);
  const scale = dpi / 72;
  const result = {pages: pageCount, png: [], lines: []};

  for (let index = 0; index < pageCount; index++) {
    const page = doc.pageAtIndex(index);
    const bounds = page.boundsForBox($.kPDFDisplayBoxMediaBox);
    const width = Math.round(bounds.size.width * scale);
    const height = Math.round(bounds.size.height * scale);

    const rep = $.NSBitmapImageRep.alloc.initWithBitmapDataPlanesPixelsWidePixelsHighBitsPerSampleSamplesPerPixelHasAlphaIsPlanarColorSpaceNameBytesPerRowBitsPerPixel(
      null, width, height, 8, 4, true, false, $.NSDeviceRGBColorSpace, 0, 0);
    const context = $.NSGraphicsContext.graphicsContextWithBitmapImageRep(rep);
    $.NSGraphicsContext.saveGraphicsState;
    $.NSGraphicsContext.setCurrentContext(context);
    $.NSColor.whiteColor.setFill;
    $.NSRectFill($.NSMakeRect(0, 0, width, height));
    const cgContext = context.CGContext;
    $.CGContextScaleCTM(cgContext, scale, scale);
    page.drawWithBoxToContext($.kPDFDisplayBoxMediaBox, cgContext);
    $.NSGraphicsContext.restoreGraphicsState;

    const data = rep.representationUsingTypeProperties($.NSBitmapImageFileTypePNG, $({}));
    const pngPath = outDir + '/' + prefix + '-page-' + (index + 1) + '.png';
    if (!data.writeToFileAtomically(pngPath, true)) {
      fail('cannot write ' + pngPath);
    }
    result.png.push(pngPath);

    const text = page.string.js || '';
    const pageHeight = bounds.size.height;
    let offset = 0;
    for (const line of text.split('\n')) {
      if (line.trim()) {
        const selection = page.selectionForRange($.NSMakeRange(offset, line.length));
        const rect = selection.boundsForPage(page);
        result.lines.push({
          page: index + 1,
          text: line,
          x: Number(rect.origin.x.toFixed(1)),
          y: Number((pageHeight - rect.origin.y - rect.size.height).toFixed(1)),
          w: Number(rect.size.width.toFixed(1)),
          h: Number(rect.size.height.toFixed(1)),
        });
      }
      offset += line.length + 1;
    }
  }
  return JSON.stringify(result);
}
