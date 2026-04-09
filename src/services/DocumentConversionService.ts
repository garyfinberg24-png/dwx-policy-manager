// @ts-nocheck
/**
 * DocumentConversionService
 *
 * Converts Office documents to clean HTML at publish time.
 * Uses the dwx-pm-docconv Azure Function (server-side conversion).
 *
 * Supported formats:
 *   - .docx/.doc  → mammoth.js (semantic HTML from Word)
 *   - .pptx/.ppt  → JSZip + XML parsing (slide-by-slide HTML)
 *   - .xlsx/.xls  → SheetJS (worksheet-by-worksheet HTML tables)
 *
 * Flow:
 *   1. Author publishes a policy with a linked document
 *   2. This service calls the Azure Function to convert → HTML
 *   3. The resulting HTML is saved to the PolicyContent field on PM_Policies
 *   4. The reader renders PolicyContent as native HTML (no iframe needed)
 *
 * Fallback: If the Azure Function is unavailable, the iframe viewer is used.
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { logger } from './LoggingService';

/** File extensions supported by the doc converter Azure Function */
const CONVERTIBLE_EXTENSIONS = ['docx', 'doc', 'pptx', 'ppt', 'xlsx', 'xls'];

export class DocumentConversionService {
  private sp: SPFI;
  private functionUrl: string;

  constructor(sp: SPFI, functionUrl?: string) {
    this.sp = sp;
    // Function URL from admin config or localStorage fallback
    this.functionUrl = functionUrl || '';
  }

  /**
   * Load the function URL from PM_Configuration if not provided
   */
  public async initialize(): Promise<void> {
    if (this.functionUrl) return;

    try {
      // Try localStorage first (instant)
      const cached = localStorage.getItem('PM_DocConverter_FunctionUrl');
      if (cached) {
        this.functionUrl = cached;
        return;
      }

      // Try PM_Configuration
      const items = await this.sp.web.lists
        .getByTitle('PM_Configuration')
        .items.filter("ConfigKey eq 'Integration.DocConverter.FunctionUrl'")
        .select('ConfigValue')
        .top(1)();

      if (items.length > 0 && items[0].ConfigValue) {
        this.functionUrl = items[0].ConfigValue;
        localStorage.setItem('PM_DocConverter_FunctionUrl', this.functionUrl);
      }
    } catch {
      // PM_Configuration may not exist
    }
  }

  /**
   * Check if a file extension is supported for conversion.
   */
  public static isConvertible(documentUrl: string): boolean {
    const ext = documentUrl.split('.').pop()?.toLowerCase() || '';
    return CONVERTIBLE_EXTENSIONS.includes(ext);
  }

  /**
   * Convert an Office document to HTML.
   * Call this when publishing a policy that has a linked document.
   *
   * Supported: .docx, .doc, .pptx, .ppt
   *
   * @returns The converted HTML string, or null if conversion failed/unavailable
   */
  public async convertToHtml(
    siteUrl: string,
    documentUrl: string,
    policyId: number
  ): Promise<string | null> {
    await this.initialize();

    if (!this.functionUrl) {
      logger.warn('DocumentConversionService', 'No function URL configured — skipping conversion');
      return null;
    }

    const ext = documentUrl.split('.').pop()?.toLowerCase() || '';
    if (!CONVERTIBLE_EXTENSIONS.includes(ext)) {
      logger.info('DocumentConversionService', `Unsupported format (.${ext}) — skipping conversion`);
      return null;
    }

    try {
      logger.info('DocumentConversionService', `Converting .${ext} document: ${documentUrl} for policy ${policyId}`);

      const response = await fetch(this.functionUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ siteUrl, documentUrl, policyId })
      });

      if (!response.ok) {
        const errText = await response.text();
        logger.error('DocumentConversionService', `Conversion failed: ${response.status} — ${errText}`);
        return null;
      }

      const result = await response.json();

      if (result.html) {
        logger.info('DocumentConversionService', `Conversion successful (${result.sourceFormat || ext}): ${result.characterCount} chars, ${result.messages?.length || 0} warnings`);
        return result.html;
      }

      return null;
    } catch (error) {
      logger.error('DocumentConversionService', 'Document conversion failed:', error);
      return null;
    }
  }

  /**
   * Legacy alias — calls convertToHtml internally.
   * @deprecated Use convertToHtml() instead.
   */
  public async convertDocxToHtml(
    siteUrl: string,
    documentUrl: string,
    policyId: number
  ): Promise<string | null> {
    return this.convertToHtml(siteUrl, documentUrl, policyId);
  }

  /**
   * Convert and save — converts the document AND updates the policy's PolicyContent field.
   * One-stop method for the publish flow.
   */
  public async convertAndSave(
    siteUrl: string,
    documentUrl: string,
    policyId: number
  ): Promise<boolean> {
    console.log(`[DocConvert] Starting conversion for policy ${policyId}: ${documentUrl}`);

    // Try Azure Function conversion first
    let html = await this.convertToHtml(siteUrl, documentUrl, policyId);

    // Fallback: client-side extraction if Azure Function not available
    if (!html) {
      console.log('[DocConvert] Azure Function conversion failed/unavailable — trying client-side extraction');
      html = await this.clientSideExtract(documentUrl);
    }

    if (!html) {
      console.warn(`[DocConvert] No HTML produced for policy ${policyId}`);
      return false;
    }

    try {
      // Write to BOTH PolicyContent and HTMLContent so the reader finds it
      await this.sp.web.lists
        .getByTitle('PM_Policies')
        .items.getById(policyId)
        .update({ PolicyContent: html, HTMLContent: html });

      console.log(`[DocConvert] ✓ PolicyContent + HTMLContent updated for policy ${policyId} (${html.length} chars)`);
      return true;
    } catch (error) {
      logger.error('DocumentConversionService', `Failed to save content for policy ${policyId}:`, error);
      return false;
    }
  }

  /**
   * Client-side fallback: fetch the document via PnP and extract text.
   * Works for .docx (basic XML extraction) — limited but better than nothing.
   * Does NOT require Azure Function.
   */
  private async clientSideExtract(documentUrl: string): Promise<string | null> {
    const ext = documentUrl.split('.').pop()?.toLowerCase() || '';
    if (!['docx', 'doc'].includes(ext)) return null; // Only Word supported client-side

    try {
      const serverRelPath = documentUrl.replace(/^https?:\/\/[^/]+/i, '');
      const buffer: ArrayBuffer = await this.sp.web.getFileByServerRelativePath(serverRelPath).getBuffer();

      // Basic DOCX extraction: DOCX is a ZIP containing XML
      // Extract text from word/document.xml
      const bytes = new Uint8Array(buffer);
      let xmlContent = '';

      // Find PK header (ZIP signature)
      if (bytes[0] === 0x50 && bytes[1] === 0x4B) {
        // Decode as text and extract paragraph content from XML tags
        const decoder = new TextDecoder('utf-8');
        const fullText = decoder.decode(buffer);

        // Extract content between <w:t> tags (Word XML text nodes)
        const textMatches = fullText.match(/<w:t[^>]*>([^<]*)<\/w:t>/g);
        if (textMatches) {
          const paragraphs: string[] = [];
          let currentPara = '';

          // Also detect paragraph breaks
          const allContent = fullText;
          let pos = 0;
          const paraBreak = '</w:p>';
          const textTag = /<w:t[^>]*>([^<]*)<\/w:t>/g;

          let match;
          while ((match = textTag.exec(allContent)) !== null) {
            // Check if there was a paragraph break since last match
            const between = allContent.substring(pos, match.index);
            if (between.includes(paraBreak) && currentPara) {
              paragraphs.push(currentPara.trim());
              currentPara = '';
            }
            currentPara += match[1];
            pos = match.index + match[0].length;
          }
          if (currentPara.trim()) paragraphs.push(currentPara.trim());

          xmlContent = paragraphs
            .filter(p => p.length > 0)
            .map(p => `<p>${p}</p>`)
            .join('\n');
        }
      }

      if (xmlContent && xmlContent.length > 20) {
        console.log(`[DocConvert] Client-side extraction: ${xmlContent.length} chars from ${ext}`);
        return `<div class="converted-content"><p><em>This content was extracted from a ${ext.toUpperCase()} document.</em></p>\n${xmlContent}</div>`;
      }

      return null;
    } catch (err) {
      console.warn('[DocConvert] Client-side extraction failed:', err);
      return null;
    }
  }
}
