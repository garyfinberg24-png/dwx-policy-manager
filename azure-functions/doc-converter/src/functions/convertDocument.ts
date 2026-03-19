/**
 * convertDocument — Azure Function (HTTP Trigger)
 *
 * Converts Office documents to clean HTML for the Policy Manager reader.
 *
 * Supported formats:
 *   - .docx/.doc  → mammoth.js (semantic HTML from Word)
 *   - .pptx/.ppt  → JSZip + XML parsing (slide-by-slide HTML)
 *   - .xlsx/.xls  → SheetJS (worksheet-by-worksheet HTML tables)
 *
 * Flow:
 *   1. Client sends: { siteUrl, documentUrl, policyId }
 *   2. Function downloads the file from SharePoint via MSAL
 *   3. Routes to the correct converter based on file extension
 *   4. Applies Forest Teal styling
 *   5. Returns { html, rawHtml, messages, characterCount, policyId }
 *
 * Why server-side?
 *   - mammoth.js / JSZip need Node.js (can't run in SPFx browser bundle)
 *   - One-time conversion at publish, not per-read
 *   - Clean HTML stored in SP list field = instant reader rendering
 */

import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import * as mammoth from 'mammoth';
import * as JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { ConfidentialClientApplication } from '@azure/msal-node';

const TENANT_ID = process.env.AZURE_TENANT_ID || '';
const CLIENT_ID = process.env.AZURE_CLIENT_ID || '';
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET || '';

interface ConvertRequest {
  siteUrl: string;
  documentUrl: string;
  policyId: number;
}

interface ConvertResult {
  html: string;
  rawHtml: string;
  messages: { type: string; message: string }[];
}

// ============================================================================
// SHARED: Authentication & File Download
// ============================================================================

async function getGraphToken(context?: InvocationContext): Promise<string> {
  const scope = 'https://graph.microsoft.com/.default';

  context?.log(`MSAL: tenant=${TENANT_ID}, clientId=${CLIENT_ID}, scope=${scope}`);

  const cca = new ConfidentialClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`,
      clientSecret: CLIENT_SECRET
    }
  });

  try {
    const result = await cca.acquireTokenByClientCredential({
      scopes: [scope]
    });

    if (!result?.accessToken) throw new Error('MSAL returned null token');
    context?.log(`MSAL: Graph token acquired, length=${result.accessToken.length}`);
    return result.accessToken;
  } catch (msalError: any) {
    context?.error(`MSAL error: ${msalError.message}`);
    throw new Error(`MSAL authentication failed: ${msalError.message}`);
  }
}

/**
 * Normalize a document URL to a server-relative path.
 * Handles both absolute URLs (https://tenant.sharepoint.com/sites/X/doc.docx)
 * and server-relative paths (/sites/X/doc.docx).
 */
function toServerRelativeUrl(documentUrl: string, siteUrl: string): string {
  if (documentUrl.startsWith('http://') || documentUrl.startsWith('https://')) {
    try {
      const parsed = new URL(documentUrl);
      return decodeURIComponent(parsed.pathname);
    } catch {
      // Fall through
    }
  }
  return documentUrl;
}

async function downloadFile(siteUrl: string, documentUrl: string, token: string, context?: InvocationContext): Promise<Buffer> {
  const serverRelativeUrl = toServerRelativeUrl(documentUrl, siteUrl);
  const siteHostname = new URL(siteUrl).hostname;
  const sitePath = new URL(siteUrl).pathname; // e.g. /sites/PolicyManager

  // Use Graph's SharePoint endpoint to get file by server-relative URL
  // Step 1: Get the site ID
  const siteInfoUrl = `https://graph.microsoft.com/v1.0/sites/${siteHostname}:${sitePath}`;
  context?.log(`Getting site ID from: ${siteInfoUrl}`);

  const siteResponse = await fetch(siteInfoUrl, {
    headers: { 'Authorization': `Bearer ${token}` }
  });

  if (!siteResponse.ok) {
    let errorBody = '';
    try { errorBody = await siteResponse.text(); } catch { /* ignore */ }
    throw new Error(`Failed to resolve site: ${siteResponse.status} — ${errorBody.substring(0, 300)}`);
  }

  const siteInfo = await siteResponse.json() as { id: string };
  const siteId = siteInfo.id;
  context?.log(`Site ID: ${siteId}`);

  // Step 2: Get the file path relative to site — split into library name and file path
  // serverRelativeUrl: /sites/PolicyManager/PM_PolicyDocuments/Code of Conduct Policy.docx
  // We need: library = PM_PolicyDocuments, file = Code of Conduct Policy.docx
  const pathAfterSite = serverRelativeUrl.startsWith(sitePath)
    ? serverRelativeUrl.substring(sitePath.length)  // /PM_PolicyDocuments/Code of Conduct Policy.docx
    : serverRelativeUrl;

  const pathParts = pathAfterSite.split('/').filter(Boolean); // ['PM_PolicyDocuments', 'Code of Conduct Policy.docx']
  const libraryName = pathParts[0]; // PM_PolicyDocuments
  const filePath = pathParts.slice(1).join('/'); // Code of Conduct Policy.docx

  // Step 3: Get the drive (document library) by name
  const drivesUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`;
  context?.log(`Getting drives from: ${drivesUrl}`);

  const drivesResponse = await fetch(drivesUrl, {
    headers: { 'Authorization': `Bearer ${token}` }
  });

  if (!drivesResponse.ok) {
    let errorBody = '';
    try { errorBody = await drivesResponse.text(); } catch { /* ignore */ }
    throw new Error(`Failed to list drives: ${drivesResponse.status} — ${errorBody.substring(0, 300)}`);
  }

  const drivesData = await drivesResponse.json() as { value: { id: string; name: string }[] };
  const drive = drivesData.value.find(d => d.name === libraryName);

  if (!drive) {
    throw new Error(`Document library '${libraryName}' not found. Available: ${drivesData.value.map(d => d.name).join(', ')}`);
  }

  // Step 4: Download the file content
  const fileUrl = `https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${encodeURI(filePath)}:/content`;
  context?.log(`Downloading file from: ${fileUrl}`);

  const response = await fetch(fileUrl, {
    headers: {
      'Authorization': `Bearer ${token}`
    },
    redirect: 'follow'
  });

  if (!response.ok) {
    let errorBody = '';
    try { errorBody = await response.text(); } catch { /* ignore */ }
    throw new Error(`Failed to download document: ${response.status} ${response.statusText} — Body: ${errorBody.substring(0, 500)}`);
  }

  const arrayBuffer = await response.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

// ============================================================================
// SHARED: Forest Teal Styling
// ============================================================================

function applyPolicyStyles(html: string): string {
  return `
<div class="pm-policy-content" style="font-family: 'Segoe UI', -apple-system, sans-serif; color: #334155; line-height: 1.8; font-size: 15px; max-width: 800px; margin: 0 auto;">
  <style>
    .pm-policy-content h1 { font-size: 24px; font-weight: 700; color: #0f172a; margin: 32px 0 16px; padding-bottom: 8px; border-bottom: 2px solid #0d9488; }
    .pm-policy-content h2 { font-size: 20px; font-weight: 700; color: #0f172a; margin: 28px 0 12px; padding-bottom: 6px; border-bottom: 1px solid #e2e8f0; }
    .pm-policy-content h3 { font-size: 17px; font-weight: 600; color: #0f172a; margin: 24px 0 8px; }
    .pm-policy-content h4 { font-size: 15px; font-weight: 600; color: #334155; margin: 20px 0 8px; }
    .pm-policy-content p { margin-bottom: 12px; }
    .pm-policy-content ul, .pm-policy-content ol { margin: 8px 0 16px 24px; }
    .pm-policy-content li { margin-bottom: 6px; }
    .pm-policy-content table { border-collapse: collapse; width: 100%; margin: 16px 0; }
    .pm-policy-content th { background: #f0fdfa; color: #0f766e; font-weight: 600; text-align: left; padding: 10px 12px; border: 1px solid #e2e8f0; }
    .pm-policy-content td { padding: 10px 12px; border: 1px solid #e2e8f0; }
    .pm-policy-content tr:nth-child(even) td { background: #fafafa; }
    .pm-policy-content img { max-width: 100%; height: auto; border-radius: 4px; margin: 12px 0; }
    .pm-policy-content a { color: #0d9488; text-decoration: none; }
    .pm-policy-content a:hover { text-decoration: underline; }
    .pm-policy-content strong { color: #0f172a; }
    .pm-policy-content blockquote { border-left: 4px solid #0d9488; padding: 12px 20px; margin: 16px 0; background: #f0fdfa; color: #0f766e; font-style: italic; }
    .pm-slide { background: #fff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 32px; margin-bottom: 24px; position: relative; }
    .pm-slide-number { position: absolute; top: 12px; right: 16px; font-size: 12px; font-weight: 600; color: #94a3b8; }
    .pm-slide h1 { margin-top: 0; }
    .pm-slide h2 { margin-top: 0; border-bottom: none; padding-bottom: 0; }
    .pm-slide-notes { margin-top: 16px; padding-top: 12px; border-top: 1px dashed #e2e8f0; font-size: 13px; color: #64748b; font-style: italic; }
    .pm-slide-notes::before { content: 'Speaker Notes: '; font-weight: 600; font-style: normal; color: #94a3b8; }
    .pm-sheet { margin-bottom: 32px; }
    .pm-sheet-title { font-size: 18px; font-weight: 600; color: #0f766e; margin: 0 0 12px; padding: 8px 16px; background: #f0fdfa; border-left: 4px solid #0d9488; border-radius: 0 4px 4px 0; }
    .pm-policy-content .pm-sheet table { font-size: 14px; }
    .pm-policy-content .pm-sheet td.pm-cell-number { text-align: right; font-variant-numeric: tabular-nums; }
    .pm-policy-content .pm-sheet tr:first-child th,
    .pm-policy-content .pm-sheet tr:first-child td { font-weight: 600; background: #f0fdfa; color: #0f766e; }
  </style>
  ${html}
</div>`;
}

// ============================================================================
// CONVERTER: Word (.docx)
// ============================================================================

async function convertDocx(fileBuffer: Buffer): Promise<ConvertResult> {
  const result = await mammoth.convertToHtml(
    { buffer: fileBuffer },
    {
      styleMap: [
        "p[style-name='Heading 1'] => h1:fresh",
        "p[style-name='Heading 2'] => h2:fresh",
        "p[style-name='Heading 3'] => h3:fresh",
        "p[style-name='Heading 4'] => h4:fresh",
        "p[style-name='Title'] => h1.document-title:fresh",
        "p[style-name='Subtitle'] => p.document-subtitle:fresh",
        "p[style-name='Quote'] => blockquote:fresh",
        "p[style-name='Intense Quote'] => blockquote.intense:fresh",
        "r[style-name='Strong'] => strong",
        "r[style-name='Emphasis'] => em",
      ],
      convertImage: (mammoth.images as any).inline(function(element: any) {
        return element.read('base64').then(function(imageBuffer: string) {
          return { src: 'data:' + element.contentType + ';base64,' + imageBuffer };
        });
      })
    }
  );

  return {
    html: result.value,
    rawHtml: result.value,
    messages: result.messages.map(m => ({ type: m.type, message: m.message }))
  };
}

// ============================================================================
// CONVERTER: PowerPoint (.pptx)
// ============================================================================

/**
 * Extract text content from a PowerPoint XML text body node.
 * Walks <a:p> paragraphs → <a:r> runs → <a:t> text nodes.
 * Detects bullet lists and heading-level text by font size.
 */
function extractTextFromBody(bodyXml: string): { html: string; plainText: string } {
  const paragraphs: string[] = [];
  const plainLines: string[] = [];

  // Extract each <a:p>...</a:p> paragraph
  const pRegex = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g;
  let pMatch: RegExpExecArray | null;

  while ((pMatch = pRegex.exec(bodyXml)) !== null) {
    const pContent = pMatch[1];

    // Extract all text runs <a:r>...<a:t>text</a:t>...</a:r>
    const textParts: string[] = [];
    const boldParts: boolean[] = [];
    const runRegex = /<a:r\b[^>]*>([\s\S]*?)<\/a:r>/g;
    let runMatch: RegExpExecArray | null;

    while ((runMatch = runRegex.exec(pContent)) !== null) {
      const runContent = runMatch[1];
      // Check for bold
      const isBold = /<a:rPr[^>]*\bb="1"/.test(runContent);
      // Extract text
      const textRegex = /<a:t[^>]*>([\s\S]*?)<\/a:t>/g;
      let textMatch: RegExpExecArray | null;
      while ((textMatch = textRegex.exec(runContent)) !== null) {
        const text = textMatch[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
        if (text.trim()) {
          textParts.push(text);
          boldParts.push(isBold);
        }
      }
    }

    // Also check for text directly in <a:fld> (field codes like slide numbers)
    const fldRegex = /<a:fld[^>]*>[\s\S]*?<a:t[^>]*>([\s\S]*?)<\/a:t>[\s\S]*?<\/a:fld>/g;
    let fldMatch: RegExpExecArray | null;
    while ((fldMatch = fldRegex.exec(pContent)) !== null) {
      // Skip field codes (slide numbers, dates)
    }

    if (textParts.length === 0) continue;

    const fullText = textParts.join('');
    plainLines.push(fullText);

    // Detect heading by font size (>= 2400 hundredths of a point = 24pt)
    const fontSizeMatch = /<a:rPr[^>]*\bsz="(\d+)"/.exec(pContent)
      || /<a:defRPr[^>]*\bsz="(\d+)"/.exec(pContent);
    const fontSize = fontSizeMatch ? parseInt(fontSizeMatch[1], 10) : 0;

    // Check for bullet/numbering
    const hasBullet = /<a:buChar/.test(pContent) || /<a:buAutoNum/.test(pContent) || /<a:buNone/.test(pContent) === false && /<a:pPr[^>]*\blvl="/.test(pContent);

    // Build HTML for this paragraph
    let lineHtml = textParts.map((t, i) => {
      const escaped = t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      return boldParts[i] ? `<strong>${escaped}</strong>` : escaped;
    }).join('');

    if (fontSize >= 2400) {
      paragraphs.push(`<h2>${lineHtml}</h2>`);
    } else if (fontSize >= 2000) {
      paragraphs.push(`<h3>${lineHtml}</h3>`);
    } else if (hasBullet) {
      paragraphs.push(`<li>${lineHtml}</li>`);
    } else {
      paragraphs.push(`<p>${lineHtml}</p>`);
    }
  }

  // Wrap consecutive <li> items in <ul>
  let html = '';
  let inList = false;
  for (const p of paragraphs) {
    if (p.startsWith('<li>')) {
      if (!inList) { html += '<ul>'; inList = true; }
      html += p;
    } else {
      if (inList) { html += '</ul>'; inList = false; }
      html += p;
    }
  }
  if (inList) html += '</ul>';

  return { html, plainText: plainLines.join('\n') };
}

/**
 * Extract speaker notes from a slide's notes XML.
 */
function extractNotes(notesXml: string): string {
  const bodyMatch = /<p:txBody>([\s\S]*?)<\/p:txBody>/g;
  const notes: string[] = [];

  let match: RegExpExecArray | null;
  while ((match = bodyMatch.exec(notesXml)) !== null) {
    const { plainText } = extractTextFromBody(match[1]);
    // Skip the placeholder slide number text (usually just a number)
    const cleaned = plainText.trim();
    if (cleaned && !/^\d+$/.test(cleaned)) {
      notes.push(cleaned);
    }
  }

  return notes.join(' ').trim();
}

async function convertPptx(fileBuffer: Buffer): Promise<ConvertResult> {
  const zip = await JSZip.loadAsync(fileBuffer);
  const messages: { type: string; message: string }[] = [];
  const slides: { index: number; html: string; notes: string }[] = [];

  // Find all slide files (ppt/slides/slide1.xml, slide2.xml, ...)
  const slideFiles = Object.keys(zip.files)
    .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => {
      const numA = parseInt(a.match(/slide(\d+)/)?.[1] || '0', 10);
      const numB = parseInt(b.match(/slide(\d+)/)?.[1] || '0', 10);
    return numA - numB;
    });

  if (slideFiles.length === 0) {
    throw new Error('No slides found in .pptx file');
  }

  messages.push({ type: 'info', message: `Found ${slideFiles.length} slides` });

  // Extract images from ppt/media/ for inline embedding
  const mediaMap = new Map<string, string>();
  const mediaFiles = Object.keys(zip.files).filter(f => f.startsWith('ppt/media/'));
  for (const mediaPath of mediaFiles) {
    try {
      const data = await zip.files[mediaPath].async('base64');
      const ext = mediaPath.split('.').pop()?.toLowerCase() || 'png';
      const mimeMap: Record<string, string> = {
        'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
        'gif': 'image/gif', 'svg': 'image/svg+xml', 'bmp': 'image/bmp',
        'emf': 'image/emf', 'wmf': 'image/wmf', 'tiff': 'image/tiff', 'tif': 'image/tiff'
      };
      const mime = mimeMap[ext] || 'image/png';
      mediaMap.set(mediaPath.split('/').pop() || '', `data:${mime};base64,${data}`);
    } catch {
      // Skip unreadable media
    }
  }

  for (const slideFile of slideFiles) {
    const slideNum = parseInt(slideFile.match(/slide(\d+)/)?.[1] || '0', 10);
    const slideXml = await zip.files[slideFile].async('string');

    // Extract text from all shape text bodies
    let slideHtml = '';
    const bodyRegex = /<p:txBody>([\s\S]*?)<\/p:txBody>/g;
    let bodyMatch: RegExpExecArray | null;

    while ((bodyMatch = bodyRegex.exec(slideXml)) !== null) {
      const { html } = extractTextFromBody(bodyMatch[1]);
      if (html.trim()) {
        slideHtml += html;
      }
    }

    // Extract images referenced in this slide
    // Parse relationships file for this slide
    const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
    const relsFile = zip.files[relsPath];
    const imageRefs: string[] = [];

    if (relsFile) {
      const relsXml = await relsFile.async('string');
      const relRegex = /<Relationship[^>]*Target="([^"]*)"[^>]*Type="[^"]*\/image"/g;
      let relMatch: RegExpExecArray | null;
      while ((relMatch = relRegex.exec(relsXml)) !== null) {
        const target = relMatch[1];
        const filename = target.split('/').pop() || '';
        const dataUri = mediaMap.get(filename);
        if (dataUri) {
          imageRefs.push(dataUri);
        }
      }
    }

    // Add images to slide HTML
    for (const imgSrc of imageRefs) {
      slideHtml += `<img src="${imgSrc}" alt="Slide ${slideNum} image" />`;
    }

    // Extract speaker notes
    const notesPath = `ppt/notesSlides/notesSlide${slideNum}.xml`;
    let notes = '';
    if (zip.files[notesPath]) {
      const notesXml = await zip.files[notesPath].async('string');
      notes = extractNotes(notesXml);
    }

    if (slideHtml.trim() || notes) {
      slides.push({ index: slideNum, html: slideHtml, notes });
    } else {
      messages.push({ type: 'warning', message: `Slide ${slideNum} has no extractable text content` });
    }
  }

  // Build final HTML with slide cards
  const rawHtml = slides.map(slide => {
    let card = `<div class="pm-slide">`;
    card += `<span class="pm-slide-number">Slide ${slide.index} of ${slideFiles.length}</span>`;
    card += slide.html;
    if (slide.notes) {
      card += `<div class="pm-slide-notes">${slide.notes.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</div>`;
    }
    card += `</div>`;
    return card;
  }).join('\n');

  return { html: rawHtml, rawHtml, messages };
}

// ============================================================================
// CONVERTER: Excel (.xlsx)
// ============================================================================

/**
 * Escape a cell value for safe HTML rendering.
 */
function escapeCell(value: any): string {
  if (value === null || value === undefined) return '';
  const str = String(value);
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

async function convertXlsx(fileBuffer: Buffer): Promise<ConvertResult> {
  const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
  const messages: { type: string; message: string }[] = [];
  const sheetHtmlParts: string[] = [];

  if (workbook.SheetNames.length === 0) {
    throw new Error('No worksheets found in Excel file');
  }

  messages.push({ type: 'info', message: `Found ${workbook.SheetNames.length} worksheet(s)` });

  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet || !worksheet['!ref']) {
      messages.push({ type: 'info', message: `Skipping empty sheet: ${sheetName}` });
      continue;
    }

    // Get data as 2D array (header: 1 = raw rows, no key mapping)
    const rows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    // Skip sheets where all rows are empty
    const hasContent = rows.some(row => row.some((cell: any) => cell !== '' && cell !== null && cell !== undefined));
    if (!hasContent) {
      messages.push({ type: 'info', message: `Skipping empty sheet: ${sheetName}` });
      continue;
    }

    // Build merged cell lookup: { "r,c" → true } for cells that are covered (not the top-left)
    const merges = worksheet['!merges'] || [];
    const mergeMap = new Map<string, { rowSpan: number; colSpan: number }>();
    const coveredCells = new Set<string>();

    for (const merge of merges) {
      const rowSpan = merge.e.r - merge.s.r + 1;
      const colSpan = merge.e.c - merge.s.c + 1;
      mergeMap.set(`${merge.s.r},${merge.s.c}`, { rowSpan, colSpan });

      // Mark all covered cells (except the top-left origin)
      for (let r = merge.s.r; r <= merge.e.r; r++) {
        for (let c = merge.s.c; c <= merge.e.c; c++) {
          if (r !== merge.s.r || c !== merge.s.c) {
            coveredCells.add(`${r},${c}`);
          }
        }
      }
    }

    // Determine max columns across all rows
    const maxCols = rows.reduce((max, row) => Math.max(max, row.length), 0);

    // Build HTML table
    let tableHtml = '<table>';

    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      // Use <th> for the first row (header row)
      const cellTag = r === 0 ? 'th' : 'td';

      tableHtml += '<tr>';
      for (let c = 0; c < maxCols; c++) {
        const key = `${r},${c}`;

        // Skip cells covered by a merge
        if (coveredCells.has(key)) continue;

        const cellValue = c < row.length ? row[c] : '';
        const escaped = escapeCell(cellValue);

        // Check for merge spans
        const merge = mergeMap.get(key);
        let attrs = '';
        if (merge) {
          if (merge.rowSpan > 1) attrs += ` rowspan="${merge.rowSpan}"`;
          if (merge.colSpan > 1) attrs += ` colspan="${merge.colSpan}"`;
        }

        // Detect numeric cells for right-alignment
        const isNumber = typeof cellValue === 'number';
        const cssClass = isNumber && r > 0 ? ' class="pm-cell-number"' : '';

        tableHtml += `<${cellTag}${attrs}${cssClass}>${escaped}</${cellTag}>`;
      }
      tableHtml += '</tr>';
    }

    tableHtml += '</table>';

    // Wrap in sheet section
    const escapedName = escapeCell(sheetName);
    const sheetSection = `<div class="pm-sheet">
<h3 class="pm-sheet-title">${escapedName}</h3>
${tableHtml}
</div>`;

    sheetHtmlParts.push(sheetSection);
  }

  if (sheetHtmlParts.length === 0) {
    throw new Error('All worksheets are empty');
  }

  const rawHtml = sheetHtmlParts.join('\n');
  messages.push({ type: 'info', message: `Rendered ${sheetHtmlParts.length} non-empty sheet(s)` });

  return { html: rawHtml, rawHtml, messages };
}

// ============================================================================
// MAIN HANDLER
// ============================================================================

export async function convertDocument(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  context.log('Document conversion request received');

  // CORS
  const origin = request.headers.get('origin') || '';
  const corsHeaders: Record<string, string> = {
    'Access-Control-Allow-Origin': origin.includes('.sharepoint.com') ? origin : '',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization'
  };

  if (request.method === 'OPTIONS') {
    return { status: 204, headers: corsHeaders };
  }

  try {
    const body = await request.json() as ConvertRequest;
    const { siteUrl, documentUrl, policyId } = body;

    if (!siteUrl || !documentUrl) {
      return {
        status: 400,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
        body: JSON.stringify({ error: 'siteUrl and documentUrl are required' })
      };
    }

    // Determine file type
    const ext = documentUrl.split('.').pop()?.toLowerCase() || '';
    const supportedFormats = ['docx', 'doc', 'pptx', 'ppt', 'xlsx', 'xls'];

    if (!supportedFormats.includes(ext)) {
      return {
        status: 400,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
        body: JSON.stringify({ error: `Unsupported file format: .${ext}. Supported: ${supportedFormats.join(', ')}` })
      };
    }

    context.log(`Converting ${ext} document: ${documentUrl} for policy ${policyId}`);

    // 1. Get Graph API token
    const token = await getGraphToken(context);

    // 2. Download the file via Microsoft Graph
    const fileBuffer = await downloadFile(siteUrl, documentUrl, token, context);
    context.log(`Downloaded ${fileBuffer.length} bytes`);

    // 3. Route to correct converter
    let result: ConvertResult;

    if (ext === 'docx' || ext === 'doc') {
      result = await convertDocx(fileBuffer);
    } else if (ext === 'pptx' || ext === 'ppt') {
      result = await convertPptx(fileBuffer);
    } else if (ext === 'xlsx' || ext === 'xls') {
      result = await convertXlsx(fileBuffer);
    } else {
      throw new Error(`No converter for .${ext}`);
    }

    // 4. Apply Forest Teal styling
    const styledHtml = applyPolicyStyles(result.html);

    context.log(`Conversion complete (${ext}): ${styledHtml.length} chars, ${result.messages.length} messages`);

    // 5. Return
    return {
      status: 200,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        html: styledHtml,
        rawHtml: result.rawHtml,
        messages: result.messages,
        characterCount: styledHtml.length,
        policyId,
        sourceFormat: ext
      })
    };

  } catch (error: any) {
    context.error('Document conversion failed:', error);
    return {
      status: 500,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        error: 'Document conversion failed',
        details: error.message || 'Unknown error'
      })
    };
  }
}

app.http('convertDocument', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'function',
  handler: convertDocument
});
