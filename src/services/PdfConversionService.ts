// @ts-nocheck
/**
 * PDF Conversion Service
 * Provides PDF generation and conversion capabilities for document templates
 */

import * as pdfMakeModule from 'pdfmake/build/pdfmake';
import * as pdfFontsModule from 'pdfmake/build/vfs_fonts';
import { TDocumentDefinitions, Content, StyleDictionary } from 'pdfmake/interfaces';
import { ITemplateDataContext } from './DocxTemplateProcessor';
import { logger } from './LoggingService';

// Get pdfMake instance (handle both module formats)
const pdfMake = (pdfMakeModule as unknown as { default: typeof pdfMakeModule }).default || pdfMakeModule;
const pdfFonts = (pdfFontsModule as unknown as { default: { pdfMake: { vfs: Record<string, string> } } }).default || pdfFontsModule;

// Initialize pdfMake with fonts
if (pdfFonts && (pdfFonts as { pdfMake?: { vfs?: Record<string, string> } }).pdfMake?.vfs) {
  (pdfMake as unknown as { vfs: Record<string, string> }).vfs = (pdfFonts as { pdfMake: { vfs: Record<string, string> } }).pdfMake.vfs;
}

/**
 * PDF document configuration options
 */
export interface IPdfDocumentOptions {
  /** Document title */
  title?: string;
  /** Document author */
  author?: string;
  /** Document subject */
  subject?: string;
  /** Page size: A4, Letter, Legal, etc. */
  pageSize?: 'A4' | 'LETTER' | 'LEGAL' | 'A3' | 'A5';
  /** Page orientation */
  pageOrientation?: 'portrait' | 'landscape';
  /** Page margins [left, top, right, bottom] */
  pageMargins?: [number, number, number, number];
  /** Include header */
  includeHeader?: boolean;
  /** Include footer with page numbers */
  includeFooter?: boolean;
  /** Company logo URL (base64 or data URL) */
  logoDataUrl?: string;
}

/**
 * PDF generation result
 */
export interface IPdfGenerationResult {
  success: boolean;
  blob?: Blob;
  fileName?: string;
  error?: string;
}

/**
 * PDF section definition
 */
export interface IPdfSection {
  title: string;
  content: string | string[];
  type?: 'text' | 'list' | 'table';
  tableData?: string[][];
}

/**
 * PDF Conversion Service
 * Generates PDF documents from templates and data
 */
export class PdfConversionService {
  private readonly defaultStyles: StyleDictionary = {
    header: {
      fontSize: 24,
      bold: true,
      margin: [0, 0, 0, 20],
      color: '#0078d4'
    },
    subheader: {
      fontSize: 16,
      bold: true,
      margin: [0, 15, 0, 10],
      color: '#323130'
    },
    body: {
      fontSize: 11,
      margin: [0, 0, 0, 10],
      lineHeight: 1.4
    },
    label: {
      fontSize: 10,
      bold: true,
      color: '#605e5c'
    },
    value: {
      fontSize: 11,
      color: '#323130'
    },
    footer: {
      fontSize: 9,
      color: '#a19f9d',
      italics: true
    },
    tableHeader: {
      fontSize: 11,
      bold: true,
      fillColor: '#f3f2f1',
      color: '#323130'
    },
    tableCell: {
      fontSize: 10,
      color: '#323130'
    }
  };

  /**
   * Generate a PDF from template data context
   * @param dataContext The template data to include in the PDF
   * @param sections Custom sections to include
   * @param options PDF document options
   */
  public async generatePdf(
    dataContext: Partial<ITemplateDataContext>,
    sections: IPdfSection[],
    options?: IPdfDocumentOptions
  ): Promise<IPdfGenerationResult> {
    try {
      const docDefinition = this.buildDocumentDefinition(dataContext, sections, options);

      return new Promise((resolve) => {
        const pdfDocGenerator = pdfMake.createPdf(docDefinition);

        pdfDocGenerator.getBlob((blob: Blob) => {
          const fileName = this.generateFileName(dataContext, options);

          logger.info('PdfConversionService', `PDF generated successfully: ${fileName}`);

          resolve({
            success: true,
            blob,
            fileName
          });
        });
      });
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error during PDF generation';
      logger.error('PdfConversionService', 'Failed to generate PDF:', error);

      return {
        success: false,
        error: errorMessage
      };
    }
  }

  /**
   * Generate a simple PDF from text content
   * @param title Document title
   * @param content Text content or array of paragraphs
   * @param options PDF document options
   */
  public async generateSimplePdf(
    title: string,
    content: string | string[],
    options?: IPdfDocumentOptions
  ): Promise<IPdfGenerationResult> {
    try {
      const contentArray = Array.isArray(content) ? content : [content];

      const docDefinition: TDocumentDefinitions = {
        info: {
          title: options?.title || title,
          author: options?.author || 'JML Document Builder',
          subject: options?.subject || 'Generated Document'
        },
        pageSize: options?.pageSize || 'A4',
        pageOrientation: options?.pageOrientation || 'portrait',
        pageMargins: options?.pageMargins || [40, 60, 40, 60],
        content: [
          { text: title, style: 'header' },
          ...contentArray.map(para => ({ text: para, style: 'body' }))
        ],
        styles: this.defaultStyles,
        footer: options?.includeFooter !== false ? this.createFooter() : undefined
      };

      return new Promise((resolve) => {
        const pdfDocGenerator = pdfMake.createPdf(docDefinition);

        pdfDocGenerator.getBlob((blob: Blob) => {
          const fileName = `${title.replace(/\s+/g, '_')}_${new Date().toISOString().slice(0, 10)}.pdf`;

          resolve({
            success: true,
            blob,
            fileName
          });
        });
      });
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      logger.error('PdfConversionService', 'Failed to generate simple PDF:', error);

      return {
        success: false,
        error: errorMessage
      };
    }
  }

  /**
   * Generate an employee document PDF
   * @param dataContext Employee and process data
   * @param documentTitle Document title
   * @param bodyContent Main document body
   * @param options PDF options
   */
  public async generateEmployeeDocument(
    dataContext: Partial<ITemplateDataContext>,
    documentTitle: string,
    bodyContent: string,
    options?: IPdfDocumentOptions
  ): Promise<IPdfGenerationResult> {
    const sections: IPdfSection[] = [
      {
        title: 'Employee Information',
        type: 'table',
        content: '',
        tableData: [
          ['Name', dataContext.employee?.name || 'N/A'],
          ['Email', dataContext.employee?.email || 'N/A'],
          ['Department', dataContext.employee?.department || 'N/A'],
          ['Job Title', dataContext.employee?.jobTitle || 'N/A'],
          ['Start Date', dataContext.employee?.startDate || 'N/A']
        ]
      },
      {
        title: 'Document Content',
        type: 'text',
        content: bodyContent
      }
    ];

    if (dataContext.manager?.name) {
      sections.splice(1, 0, {
        title: 'Manager Information',
        type: 'table',
        content: '',
        tableData: [
          ['Manager Name', dataContext.manager.name],
          ['Manager Email', dataContext.manager.email || 'N/A']
        ]
      });
    }

    return this.generatePdf(dataContext, sections, {
      ...options,
      title: documentTitle
    });
  }

  /**
   * Build the PDF document definition
   */
  private buildDocumentDefinition(
    dataContext: Partial<ITemplateDataContext>,
    sections: IPdfSection[],
    options?: IPdfDocumentOptions
  ): TDocumentDefinitions {
    const content: Content[] = [];

    // Add header with logo if provided
    if (options?.includeHeader !== false) {
      content.push(this.createDocumentHeader(dataContext, options));
    }

    // Add sections
    for (let i = 0; i < sections.length; i++) {
      const section = sections[i];
      content.push(this.createSection(section));
    }

    // Add signature line
    content.push(this.createSignatureLine());

    return {
      info: {
        title: options?.title || 'Generated Document',
        author: options?.author || 'JML Document Builder',
        subject: options?.subject || 'Employee Document',
        creationDate: new Date()
      },
      pageSize: options?.pageSize || 'A4',
      pageOrientation: options?.pageOrientation || 'portrait',
      pageMargins: options?.pageMargins || [40, 60, 40, 60],
      content,
      styles: this.defaultStyles,
      footer: options?.includeFooter !== false ? this.createFooter() : undefined
    };
  }

  /**
   * Create document header content
   */
  private createDocumentHeader(
    dataContext: Partial<ITemplateDataContext>,
    options?: IPdfDocumentOptions
  ): Content {
    const headerContent: Content[] = [];

    // Add company name
    if (dataContext.company?.name) {
      headerContent.push({
        text: dataContext.company.name,
        style: 'header'
      });
    }

    // Add document title
    if (options?.title) {
      headerContent.push({
        text: options.title,
        style: 'subheader'
      });
    }

    // Add generation date
    headerContent.push({
      text: `Generated: ${dataContext.document?.generatedDate || new Date().toLocaleDateString()}`,
      style: 'footer',
      margin: [0, 0, 0, 20]
    });

    return {
      stack: headerContent
    };
  }

  /**
   * Create a section in the PDF
   */
  private createSection(section: IPdfSection): Content {
    const sectionContent: Content[] = [
      { text: section.title, style: 'subheader' }
    ];

    switch (section.type) {
      case 'table':
        if (section.tableData) {
          sectionContent.push(this.createTable(section.tableData));
        }
        break;

      case 'list':
        const listItems = Array.isArray(section.content) ? section.content : [section.content];
        sectionContent.push({
          ul: listItems,
          style: 'body'
        });
        break;

      case 'text':
      default:
        const textItems = Array.isArray(section.content) ? section.content : [section.content];
        for (let i = 0; i < textItems.length; i++) {
          sectionContent.push({
            text: textItems[i],
            style: 'body'
          });
        }
        break;
    }

    return {
      stack: sectionContent,
      margin: [0, 0, 0, 15]
    };
  }

  /**
   * Create a table from data
   */
  private createTable(data: string[][]): Content {
    return {
      table: {
        widths: ['30%', '70%'],
        body: data.map((row, index) => {
          return row.map((cell, cellIndex) => ({
            text: cell,
            style: cellIndex === 0 ? 'label' : 'value',
            fillColor: index % 2 === 0 ? '#faf9f8' : undefined,
            margin: [5, 5, 5, 5]
          }));
        })
      },
      layout: {
        hLineWidth: () => 0.5,
        vLineWidth: () => 0.5,
        hLineColor: () => '#e1dfdd',
        vLineColor: () => '#e1dfdd'
      },
      margin: [0, 0, 0, 10]
    };
  }

  /**
   * Create signature line
   */
  private createSignatureLine(): Content {
    return {
      stack: [
        { text: '', margin: [0, 30, 0, 0] },
        {
          columns: [
            {
              stack: [
                { canvas: [{ type: 'line', x1: 0, y1: 0, x2: 200, y2: 0, lineWidth: 1, lineColor: '#323130' }] },
                { text: 'Employee Signature', style: 'label', margin: [0, 5, 0, 0] },
                { text: 'Date: _______________', style: 'label', margin: [0, 10, 0, 0] }
              ],
              width: '45%'
            },
            { width: '10%', text: '' },
            {
              stack: [
                { canvas: [{ type: 'line', x1: 0, y1: 0, x2: 200, y2: 0, lineWidth: 1, lineColor: '#323130' }] },
                { text: 'Manager Signature', style: 'label', margin: [0, 5, 0, 0] },
                { text: 'Date: _______________', style: 'label', margin: [0, 10, 0, 0] }
              ],
              width: '45%'
            }
          ]
        }
      ],
      margin: [0, 40, 0, 0]
    };
  }

  /**
   * Create page footer
   */
  private createFooter(): (currentPage: number, pageCount: number) => Content {
    return (currentPage: number, pageCount: number): Content => ({
      columns: [
        {
          text: 'Generated by JML Document Builder',
          style: 'footer',
          alignment: 'left',
          margin: [40, 0, 0, 0]
        },
        {
          text: `Page ${currentPage} of ${pageCount}`,
          style: 'footer',
          alignment: 'right',
          margin: [0, 0, 40, 0]
        }
      ],
      margin: [0, 20, 0, 0]
    });
  }

  /**
   * Generate a file name for the PDF
   */
  private generateFileName(
    dataContext: Partial<ITemplateDataContext>,
    options?: IPdfDocumentOptions
  ): string {
    const timestamp = new Date().toISOString().slice(0, 10);
    const baseName = options?.title?.replace(/\s+/g, '_') || 'Document';
    const employeeName = dataContext.employee?.name?.replace(/\s+/g, '_') || '';

    if (employeeName) {
      return `${baseName}_${employeeName}_${timestamp}.pdf`;
    }

    return `${baseName}_${timestamp}.pdf`;
  }

  /**
   * Download PDF directly in browser
   * @param result PDF generation result
   */
  public downloadPdf(result: IPdfGenerationResult): void {
    if (!result.success || !result.blob) {
      logger.error('PdfConversionService', 'Cannot download: PDF generation failed');
      return;
    }

    const url = URL.createObjectURL(result.blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = result.fileName || 'document.pdf';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    logger.info('PdfConversionService', `PDF downloaded: ${result.fileName}`);
  }

  /**
   * Open PDF in new browser tab
   * @param result PDF generation result
   */
  public openPdfInNewTab(result: IPdfGenerationResult): void {
    if (!result.success || !result.blob) {
      logger.error('PdfConversionService', 'Cannot open: PDF generation failed');
      return;
    }

    const url = URL.createObjectURL(result.blob);
    window.open(url, '_blank');

    // Clean up URL after a delay
    setTimeout(() => {
      URL.revokeObjectURL(url);
    }, 10000);
  }
}

// Export singleton instance
export const pdfConversionService = new PdfConversionService();
