// @ts-nocheck
/**
 * DocumentConversionService
 *
 * Converts .docx files to clean HTML at publish time.
 * Uses the dwx-pm-docconv Azure Function (mammoth.js server-side).
 *
 * Flow:
 *   1. Author publishes a policy with a .docx document
 *   2. This service calls the Azure Function to convert .docx → HTML
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
   * Convert a .docx document to HTML.
   * Call this when publishing a policy that has a linked .docx document.
   *
   * @returns The converted HTML string, or null if conversion failed/unavailable
   */
  public async convertDocxToHtml(
    siteUrl: string,
    documentUrl: string,
    policyId: number
  ): Promise<string | null> {
    await this.initialize();

    if (!this.functionUrl) {
      logger.warn('DocumentConversionService', 'No function URL configured — skipping conversion');
      return null;
    }

    const ext = documentUrl.split('.').pop()?.toLowerCase();
    if (ext !== 'docx' && ext !== 'doc') {
      logger.info('DocumentConversionService', `Not a Word document (${ext}) — skipping conversion`);
      return null;
    }

    try {
      logger.info('DocumentConversionService', `Converting ${documentUrl} for policy ${policyId}`);

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
        logger.info('DocumentConversionService', `Conversion successful: ${result.characterCount} chars, ${result.messages?.length || 0} warnings`);
        return result.html;
      }

      return null;
    } catch (error) {
      logger.error('DocumentConversionService', 'Document conversion failed:', error);
      return null;
    }
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
    const html = await this.convertDocxToHtml(siteUrl, documentUrl, policyId);

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
