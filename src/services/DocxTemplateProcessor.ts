// @ts-nocheck
/**
 * DOCX Template Processor Service
 * Uses docxtemplater to process Word documents with merge fields
 */

import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import { ITemplatePlaceholder, IJmlProcess } from '../models';
import { logger } from './LoggingService';

/**
 * Result from template processing
 */
export interface IDocxProcessingResult {
  success: boolean;
  blob?: Blob;
  fileName?: string;
  error?: string;
}

/**
 * Template data context for merge field replacement
 */
export interface ITemplateDataContext {
  /** Employee information */
  employee: {
    name: string;
    email: string;
    id: string;
    department: string;
    jobTitle: string;
    startDate: string;
    endDate?: string;
  };
  /** Manager information */
  manager: {
    name: string;
    email: string;
    title: string;
  };
  /** Company information */
  company: {
    name: string;
    address: string;
    phone: string;
    logo?: string;
  };
  /** Process information */
  process: {
    type: string;
    id: string;
    status: string;
    createdDate: string;
  };
  /** Document metadata */
  document: {
    generatedDate: string;
    generatedBy: string;
    version: string;
  };
  /** Custom placeholder values */
  custom: { [key: string]: string | number | boolean | Date };
}

/**
 * DOCX Template Processor using docxtemplater
 */
export class DocxTemplateProcessor {
  /**
   * Process a DOCX template with the provided data context
   * @param templateBlob The template file as a Blob
   * @param dataContext The data to merge into the template
   * @param fileName Optional output file name
   * @returns Processing result with the generated document
   */
  public async processTemplate(
    templateBlob: Blob,
    dataContext: Partial<ITemplateDataContext>,
    fileName?: string
  ): Promise<IDocxProcessingResult> {
    try {
      // Read the template as array buffer
      const arrayBuffer = await templateBlob.arrayBuffer();

      // Create a PizZip instance from the array buffer
      const zip = new PizZip(arrayBuffer);

      // Create docxtemplater instance
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        delimiters: {
          start: '{',
          end: '}'
        }
      });

      // Flatten the data context for template processing
      const flattenedData = this.flattenDataContext(dataContext);

      // Set the template data
      doc.setData(flattenedData);

      // Render the document
      doc.render();

      // Generate output
      const output = doc.getZip().generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });

      const outputFileName = fileName || this.generateFileName(dataContext);

      logger.info('DocxTemplateProcessor', `Successfully processed template: ${outputFileName}`);

      return {
        success: true,
        blob: output,
        fileName: outputFileName
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error during template processing';
      logger.error('DocxTemplateProcessor', 'Failed to process template:', error);

      return {
        success: false,
        error: errorMessage
      };
    }
  }

  /**
   * Process a template from a URL
   * @param templateUrl URL to the template file
   * @param dataContext The data to merge into the template
   * @param fileName Optional output file name
   */
  public async processTemplateFromUrl(
    templateUrl: string,
    dataContext: Partial<ITemplateDataContext>,
    fileName?: string
  ): Promise<IDocxProcessingResult> {
    try {
      // Fetch the template
      const response = await fetch(templateUrl);

      if (!response.ok) {
        throw new Error(`Failed to fetch template: ${response.status} ${response.statusText}`);
      }

      const templateBlob = await response.blob();

      // Verify it's a DOCX file
      if (!this.isDocxFile(templateBlob.type, templateUrl)) {
        throw new Error('Invalid template format. Expected a DOCX file.');
      }

      return this.processTemplate(templateBlob, dataContext, fileName);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Failed to process template from URL';
      logger.error('DocxTemplateProcessor', 'Failed to process template from URL:', error);

      return {
        success: false,
        error: errorMessage
      };
    }
  }

  /**
   * Build a data context from a JML process and custom values
   * @param process The JML process data
   * @param customValues Custom placeholder values
   * @param companyInfo Optional company information
   */
  public buildDataContext(
    process: IJmlProcess,
    customValues: { [key: string]: string },
    companyInfo?: { name: string; address: string; phone: string }
  ): ITemplateDataContext {
    const now = new Date();

    return {
      employee: {
        name: process.EmployeeName,
        email: process.EmployeeEmail || '',
        id: process.EmployeeID || '',
        department: process.Department,
        jobTitle: process.JobTitle || '',
        startDate: process.StartDate ? this.formatDate(process.StartDate) : '',
        endDate: process.ActualCompletionDate ? this.formatDate(process.ActualCompletionDate) : ''
      },
      manager: {
        name: process.Manager?.Title || '',
        email: process.Manager?.EMail || '',
        title: 'Manager'
      },
      company: {
        name: companyInfo?.name || 'Company Name',
        address: companyInfo?.address || '',
        phone: companyInfo?.phone || ''
      },
      process: {
        type: process.ProcessType,
        id: String(process.Id),
        status: process.ProcessStatus,
        createdDate: this.formatDate(process.Created || now)
      },
      document: {
        generatedDate: this.formatDate(now),
        generatedBy: 'JML Document Builder',
        version: '1.0'
      },
      custom: customValues
    };
  }

  /**
   * Extract placeholder keys from a DOCX template
   * @param templateBlob The template file as a Blob
   * @returns Array of placeholder keys found in the template
   */
  public async extractPlaceholders(templateBlob: Blob): Promise<string[]> {
    try {
      const arrayBuffer = await templateBlob.arrayBuffer();
      const zip = new PizZip(arrayBuffer);

      // Get the document.xml content
      const documentXml = zip.file('word/document.xml');
      if (!documentXml) {
        throw new Error('Invalid DOCX file: document.xml not found');
      }

      const content = documentXml.asText();

      // Find all placeholders using regex
      const placeholderRegex = /\{([^}]+)\}/g;
      const placeholders: Set<string> = new Set();
      let match: RegExpExecArray | null;

      while ((match = placeholderRegex.exec(content)) !== null) {
        // Filter out XML tags and keep only placeholder-like patterns
        const placeholder = match[1].trim();
        if (this.isValidPlaceholder(placeholder)) {
          placeholders.add(placeholder);
        }
      }

      return Array.from(placeholders);
    } catch (error) {
      logger.error('DocxTemplateProcessor', 'Failed to extract placeholders:', error);
      return [];
    }
  }

  /**
   * Validate that a template can be processed with the given data
   * @param templateBlob The template file
   * @param placeholders Expected placeholders
   * @param dataContext The data context to validate
   */
  public async validateTemplate(
    templateBlob: Blob,
    placeholders: ITemplatePlaceholder[],
    dataContext: Partial<ITemplateDataContext>
  ): Promise<{ isValid: boolean; missingFields: string[]; warnings: string[] }> {
    const missingFields: string[] = [];
    const warnings: string[] = [];

    try {
      // Extract actual placeholders from template
      const templatePlaceholders = await this.extractPlaceholders(templateBlob);
      const flattenedData = this.flattenDataContext(dataContext);

      // Check required placeholders have values
      for (let i = 0; i < placeholders.length; i++) {
        const placeholder = placeholders[i];
        if (placeholder.Required) {
          const value = flattenedData[placeholder.Key];
          if (value === undefined || value === null || value === '') {
            missingFields.push(placeholder.Label || placeholder.Key);
          }
        }
      }

      // Warn about template placeholders that may not have matching data
      for (let i = 0; i < templatePlaceholders.length; i++) {
        const templatePlaceholder = templatePlaceholders[i];
        if (flattenedData[templatePlaceholder] === undefined) {
          // Check if it's a known nested path
          if (!this.isKnownNestedPath(templatePlaceholder)) {
            warnings.push(`Template placeholder '${templatePlaceholder}' may not have a value`);
          }
        }
      }

      return {
        isValid: missingFields.length === 0,
        missingFields,
        warnings
      };
    } catch (error) {
      logger.error('DocxTemplateProcessor', 'Template validation failed:', error);
      return {
        isValid: false,
        missingFields: ['Template validation failed'],
        warnings: []
      };
    }
  }

  /**
   * Flatten the data context for template processing
   * Converts nested objects to dot-notation keys
   */
  private flattenDataContext(dataContext: Partial<ITemplateDataContext>): { [key: string]: unknown } {
    const result: { [key: string]: unknown } = {};

    // Flatten employee data
    if (dataContext.employee) {
      result['EmployeeName'] = dataContext.employee.name;
      result['EmployeeEmail'] = dataContext.employee.email;
      result['EmployeeId'] = dataContext.employee.id;
      result['Department'] = dataContext.employee.department;
      result['JobTitle'] = dataContext.employee.jobTitle;
      result['StartDate'] = dataContext.employee.startDate;
      result['EndDate'] = dataContext.employee.endDate;

      // Also add nested version
      result['employee'] = dataContext.employee;
    }

    // Flatten manager data
    if (dataContext.manager) {
      result['ManagerName'] = dataContext.manager.name;
      result['ManagerEmail'] = dataContext.manager.email;
      result['ManagerTitle'] = dataContext.manager.title;

      result['manager'] = dataContext.manager;
    }

    // Flatten company data
    if (dataContext.company) {
      result['CompanyName'] = dataContext.company.name;
      result['CompanyAddress'] = dataContext.company.address;
      result['CompanyPhone'] = dataContext.company.phone;

      result['company'] = dataContext.company;
    }

    // Flatten process data
    if (dataContext.process) {
      result['ProcessType'] = dataContext.process.type;
      result['ProcessId'] = dataContext.process.id;
      result['ProcessStatus'] = dataContext.process.status;
      result['ProcessCreatedDate'] = dataContext.process.createdDate;

      result['process'] = dataContext.process;
    }

    // Flatten document data
    if (dataContext.document) {
      result['GeneratedDate'] = dataContext.document.generatedDate;
      result['Today'] = dataContext.document.generatedDate;
      result['GeneratedBy'] = dataContext.document.generatedBy;
      result['DocumentVersion'] = dataContext.document.version;

      result['document'] = dataContext.document;
    }

    // Add custom values directly
    if (dataContext.custom) {
      const customKeys = Object.keys(dataContext.custom);
      for (let i = 0; i < customKeys.length; i++) {
        const key = customKeys[i];
        result[key] = dataContext.custom[key];
      }
    }

    return result;
  }

  /**
   * Check if a placeholder string is valid (not XML content)
   */
  private isValidPlaceholder(placeholder: string): boolean {
    // Exclude XML-like content
    if (placeholder.startsWith('w:') || placeholder.startsWith('/w:')) {
      return false;
    }
    if (placeholder.includes('<') || placeholder.includes('>')) {
      return false;
    }
    if (placeholder.includes('=')) {
      return false;
    }
    // Should contain only alphanumeric, dots, and underscores
    const validPattern = /^[a-zA-Z_][a-zA-Z0-9_.]*$/;
    return validPattern.test(placeholder);
  }

  /**
   * Check if a placeholder is a known nested path
   */
  private isKnownNestedPath(placeholder: string): boolean {
    const knownPaths = [
      'employee.name', 'employee.email', 'employee.id', 'employee.department',
      'manager.name', 'manager.email', 'manager.title',
      'company.name', 'company.address', 'company.phone',
      'process.type', 'process.id', 'process.status',
      'document.generatedDate', 'document.generatedBy'
    ];
    return knownPaths.includes(placeholder);
  }

  /**
   * Check if the file is a DOCX file
   */
  private isDocxFile(mimeType: string, url: string): boolean {
    const validMimeTypes = [
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/octet-stream'
    ];

    if (validMimeTypes.includes(mimeType)) {
      return true;
    }

    // Check by extension
    return url.toLowerCase().endsWith('.docx');
  }

  /**
   * Generate a file name based on the data context
   */
  private generateFileName(dataContext: Partial<ITemplateDataContext>): string {
    const timestamp = new Date().toISOString().slice(0, 10);
    const employeeName = dataContext.employee?.name?.replace(/\s+/g, '_') || 'Document';
    const processType = dataContext.process?.type || 'Generated';

    return `${processType}_${employeeName}_${timestamp}.docx`;
  }

  /**
   * Format a date for display in documents
   */
  private formatDate(date: Date): string {
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    const monthStr = month < 10 ? '0' + String(month) : String(month);
    const dayStr = day < 10 ? '0' + String(day) : String(day);
    return `${monthStr}/${dayStr}/${year}`;
  }
}

// Export singleton instance
export const docxTemplateProcessor = new DocxTemplateProcessor();
