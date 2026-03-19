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
    const html = await this.convertToHtml(siteUrl, documentUrl, policyId);

    if (!html) return false;

    try {
      await this.sp.web.lists
        .getByTitle('PM_Policies')
        .items.getById(policyId)
        .update({ PolicyContent: html });

      logger.info('DocumentConversionService', `PolicyContent updated for policy ${policyId}`);
      return true;
    } catch (error) {
      logger.error('DocumentConversionService', `Failed to save PolicyContent for policy ${policyId}:`, error);
      return false;
    }
  }
}
