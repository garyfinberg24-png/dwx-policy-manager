/**
 * blankOfficeDocuments.ts
 *
 * Creates minimal valid Office Open XML documents (.docx, .xlsx, .pptx) entirely
 * at runtime with zero external dependencies. Each Office file is a ZIP archive
 * (Store method, no compression) built from raw XML strings using Uint8Array.
 *
 * The resulting Blobs are valid enough for SharePoint Online's Office Online
 * (WopiFrame.aspx) to open and edit.
 */

// ---------------------------------------------------------------------------
// ZIP builder — Store method (compression = 0), no external libs
// ---------------------------------------------------------------------------

interface ZipEntry {
  path: string;
  content: string;
}

/**
 * Encode a JS string as UTF-8 bytes.
 */
function stringToUtf8(str: string): Uint8Array {
  const out: number[] = [];
  for (let i = 0; i < str.length; i++) {
    let c: number = str.charCodeAt(i);
    if (c < 0x80) {
      out.push(c);
    } else if (c < 0x800) {
      out.push(0xc0 | (c >> 6), 0x80 | (c & 0x3f));
    } else if (c >= 0xd800 && c <= 0xdbff) {
      // surrogate pair
      const hi: number = c;
      const lo: number = str.charCodeAt(++i);
      c = 0x10000 + ((hi - 0xd800) << 10) + (lo - 0xdc00);
      out.push(
        0xf0 | (c >> 18),
        0x80 | ((c >> 12) & 0x3f),
        0x80 | ((c >> 6) & 0x3f),
        0x80 | (c & 0x3f)
      );
    } else {
      out.push(0xe0 | (c >> 12), 0x80 | ((c >> 6) & 0x3f), 0x80 | (c & 0x3f));
    }
  }
  return new Uint8Array(out);
}

/**
 * CRC-32 (ISO 3309 / ITU-T V.42) used by ZIP.
 */
function crc32(data: Uint8Array): number {
  let crc: number = 0xffffffff;
  for (let i = 0; i < data.length; i++) {
    crc ^= data[i];
    for (let j = 0; j < 8; j++) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/**
 * Write a 16-bit little-endian unsigned integer into a DataView.
 */
function writeU16(view: DataView, offset: number, value: number): void {
  view.setUint16(offset, value, true);
}

/**
 * Write a 32-bit little-endian unsigned integer into a DataView.
 */
function writeU32(view: DataView, offset: number, value: number): void {
  view.setUint32(offset, value, true);
}

/**
 * Build a valid ZIP file (Store method, compression=0) from an array of
 * { path, content } entries. Returns the raw bytes of the ZIP.
 */
function buildZip(entries: ZipEntry[]): Uint8Array {
  // Pre-encode all entries
  const encoded: Array<{
    pathBytes: Uint8Array;
    dataBytes: Uint8Array;
    crc: number;
  }> = entries.map((e) => {
    const pathBytes: Uint8Array = stringToUtf8(e.path);
    const dataBytes: Uint8Array = stringToUtf8(e.content);
    const crc: number = crc32(dataBytes);
    return { pathBytes, dataBytes, crc };
  });

  // Calculate sizes
  // Local file header: 30 + pathLen + dataLen per entry
  // Central directory header: 46 + pathLen per entry
  // End of central directory: 22

  let localSize: number = 0;
  let centralSize: number = 0;
  for (const enc of encoded) {
    localSize += 30 + enc.pathBytes.length + enc.dataBytes.length;
    centralSize += 46 + enc.pathBytes.length;
  }
  const totalSize: number = localSize + centralSize + 22;

  const buffer: ArrayBuffer = new ArrayBuffer(totalSize);
  const view: DataView = new DataView(buffer);
  const bytes: Uint8Array = new Uint8Array(buffer);

  let localOffset: number = 0;
  const offsets: number[] = [];

  // DOS date/time for 2025-01-01 00:00:00
  const dosTime: number = 0x0000; // 00:00:00
  const dosDate: number = 0x5a21; // 2025-01-01

  // ---- Local file headers + data ----
  for (let i = 0; i < encoded.length; i++) {
    const enc = encoded[i];
    offsets.push(localOffset);

    // Local file header signature
    writeU32(view, localOffset, 0x04034b50);
    // Version needed to extract (2.0)
    writeU16(view, localOffset + 4, 20);
    // General purpose bit flag
    writeU16(view, localOffset + 6, 0);
    // Compression method (0 = stored)
    writeU16(view, localOffset + 8, 0);
    // Last mod file time
    writeU16(view, localOffset + 10, dosTime);
    // Last mod file date
    writeU16(view, localOffset + 12, dosDate);
    // CRC-32
    writeU32(view, localOffset + 14, enc.crc);
    // Compressed size
    writeU32(view, localOffset + 18, enc.dataBytes.length);
    // Uncompressed size
    writeU32(view, localOffset + 22, enc.dataBytes.length);
    // File name length
    writeU16(view, localOffset + 26, enc.pathBytes.length);
    // Extra field length
    writeU16(view, localOffset + 28, 0);

    localOffset += 30;

    // File name
    bytes.set(enc.pathBytes, localOffset);
    localOffset += enc.pathBytes.length;

    // File data (uncompressed)
    bytes.set(enc.dataBytes, localOffset);
    localOffset += enc.dataBytes.length;
  }

  // ---- Central directory headers ----
  let centralOffset: number = localOffset;
  for (let i = 0; i < encoded.length; i++) {
    const enc = encoded[i];

    // Central directory file header signature
    writeU32(view, centralOffset, 0x02014b50);
    // Version made by (2.0)
    writeU16(view, centralOffset + 4, 20);
    // Version needed to extract (2.0)
    writeU16(view, centralOffset + 6, 20);
    // General purpose bit flag
    writeU16(view, centralOffset + 8, 0);
    // Compression method (0 = stored)
    writeU16(view, centralOffset + 10, 0);
    // Last mod file time
    writeU16(view, centralOffset + 12, dosTime);
    // Last mod file date
    writeU16(view, centralOffset + 14, dosDate);
    // CRC-32
    writeU32(view, centralOffset + 16, enc.crc);
    // Compressed size
    writeU32(view, centralOffset + 20, enc.dataBytes.length);
    // Uncompressed size
    writeU32(view, centralOffset + 24, enc.dataBytes.length);
    // File name length
    writeU16(view, centralOffset + 28, enc.pathBytes.length);
    // Extra field length
    writeU16(view, centralOffset + 30, 0);
    // File comment length
    writeU16(view, centralOffset + 32, 0);
    // Disk number start
    writeU16(view, centralOffset + 34, 0);
    // Internal file attributes
    writeU16(view, centralOffset + 36, 0);
    // External file attributes
    writeU32(view, centralOffset + 38, 0);
    // Relative offset of local header
    writeU32(view, centralOffset + 42, offsets[i]);

    centralOffset += 46;

    // File name
    bytes.set(enc.pathBytes, centralOffset);
    centralOffset += enc.pathBytes.length;
  }

  // ---- End of central directory record ----
  const eocdOffset: number = centralOffset;
  // Signature
  writeU32(view, eocdOffset, 0x06054b50);
  // Number of this disk
  writeU16(view, eocdOffset + 4, 0);
  // Disk where central directory starts
  writeU16(view, eocdOffset + 6, 0);
  // Number of central directory records on this disk
  writeU16(view, eocdOffset + 8, encoded.length);
  // Total number of central directory records
  writeU16(view, eocdOffset + 10, encoded.length);
  // Size of central directory
  writeU32(view, eocdOffset + 12, centralOffset - localSize);
  // Offset of start of central directory
  writeU32(view, eocdOffset + 16, localSize);
  // Comment length
  writeU16(view, eocdOffset + 20, 0);

  return bytes;
}

// ---------------------------------------------------------------------------
// Base64 helper
// ---------------------------------------------------------------------------

/**
 * Decode a base64 string into a Blob with the given MIME type.
 */
export function base64ToBlob(base64: string, mimeType: string): Blob {
  const raw: string = atob(base64);
  const bytes: Uint8Array = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) {
    bytes[i] = raw.charCodeAt(i);
  }
  return new Blob([bytes], { type: mimeType });
}

// ---------------------------------------------------------------------------
// Minimal OOXML content
// ---------------------------------------------------------------------------

// ===== DOCX =====

const DOCX_CONTENT_TYPES: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '</Types>';

const DOCX_RELS: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>';

const DOCX_DOCUMENT: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ' +
  'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'mc:Ignorable="w14 wp14">' +
  '<w:body><w:p><w:r><w:t></w:t></w:r></w:p></w:body>' +
  '</w:document>';

const DOCX_DOCUMENT_RELS: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '</Relationships>';

// ===== XLSX =====

const XLSX_CONTENT_TYPES: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
  '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
  '</Types>';

const XLSX_RELS: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
  '</Relationships>';

const XLSX_WORKBOOK: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
  '</workbook>';

const XLSX_WORKBOOK_RELS: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
  '</Relationships>';

const XLSX_SHEET1: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<sheetData/>' +
  '</worksheet>';

// ===== PPTX =====

const PPTX_CONTENT_TYPES: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
  '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
  '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
  '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
  '</Types>';

const PPTX_RELS: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>' +
  '</Relationships>';

const PPTX_PRESENTATION: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
  '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>' +
  '<p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>' +
  '<p:notesSz cx="6858000" cy="9144000"/>' +
  '</p:presentation>';

const PPTX_PRESENTATION_RELS: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>' +
  '</Relationships>';

const PPTX_SLIDE1: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:cSld><p:spTree>' +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
  '</p:spTree></p:cSld>' +
  '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
  '</p:sld>';

const PPTX_SLIDE1_RELS: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
  '</Relationships>';

const PPTX_SLIDE_LAYOUT1: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">' +
  '<p:cSld name="Blank"><p:spTree>' +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
  '</p:spTree></p:cSld>' +
  '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
  '</p:sldLayout>';

const PPTX_SLIDE_LAYOUT1_RELS: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>' +
  '</Relationships>';

const PPTX_SLIDE_MASTER1: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
  '<p:spTree>' +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
  '</p:spTree></p:cSld>' +
  '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
  '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
  '</p:sldMaster>';

const PPTX_SLIDE_MASTER1_RELS: string =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
  '</Relationships>';

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Creates a minimal valid .docx file (Word document) as a Blob.
 */
export function createBlankDocx(): Blob {
  const zipBytes: Uint8Array = buildZip([
    { path: '[Content_Types].xml', content: DOCX_CONTENT_TYPES },
    { path: '_rels/.rels', content: DOCX_RELS },
    { path: 'word/document.xml', content: DOCX_DOCUMENT },
    { path: 'word/_rels/document.xml.rels', content: DOCX_DOCUMENT_RELS },
  ]);
  return new Blob([zipBytes], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

/**
 * Creates a minimal valid .xlsx file (Excel workbook) as a Blob.
 */
export function createBlankXlsx(): Blob {
  const zipBytes: Uint8Array = buildZip([
    { path: '[Content_Types].xml', content: XLSX_CONTENT_TYPES },
    { path: '_rels/.rels', content: XLSX_RELS },
    { path: 'xl/workbook.xml', content: XLSX_WORKBOOK },
    { path: 'xl/_rels/workbook.xml.rels', content: XLSX_WORKBOOK_RELS },
    { path: 'xl/worksheets/sheet1.xml', content: XLSX_SHEET1 },
  ]);
  return new Blob([zipBytes], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

/**
 * Creates a minimal valid .pptx file (PowerPoint presentation) as a Blob.
 */
export function createBlankPptx(): Blob {
  const zipBytes: Uint8Array = buildZip([
    { path: '[Content_Types].xml', content: PPTX_CONTENT_TYPES },
    { path: '_rels/.rels', content: PPTX_RELS },
    { path: 'ppt/presentation.xml', content: PPTX_PRESENTATION },
    { path: 'ppt/_rels/presentation.xml.rels', content: PPTX_PRESENTATION_RELS },
    { path: 'ppt/slides/slide1.xml', content: PPTX_SLIDE1 },
    { path: 'ppt/slides/_rels/slide1.xml.rels', content: PPTX_SLIDE1_RELS },
    { path: 'ppt/slideLayouts/slideLayout1.xml', content: PPTX_SLIDE_LAYOUT1 },
    { path: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', content: PPTX_SLIDE_LAYOUT1_RELS },
    { path: 'ppt/slideMasters/slideMaster1.xml', content: PPTX_SLIDE_MASTER1 },
    { path: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', content: PPTX_SLIDE_MASTER1_RELS },
  ]);
  return new Blob([zipBytes], {
    type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
}
