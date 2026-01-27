// @ts-nocheck
/**
 * Excel Generation Service
 * Creates Excel spreadsheets with data and formatting
 */

import ExcelJS from 'exceljs';
import { ITemplateDataContext } from './DocxTemplateProcessor';
import { logger } from './LoggingService';

/**
 * Excel generation result
 */
export interface IExcelGenerationResult {
  success: boolean;
  blob?: Blob;
  fileName?: string;
  error?: string;
}

/**
 * Excel document options
 */
export interface IExcelDocumentOptions {
  /** Workbook title */
  title?: string;
  /** Author name */
  author?: string;
  /** Company name */
  company?: string;
  /** Sheet name */
  sheetName?: string;
  /** Include header row styling */
  styleHeaders?: boolean;
  /** Auto-fit column widths */
  autoFitColumns?: boolean;
  /** Freeze first row */
  freezeFirstRow?: boolean;
}

/**
 * Excel table data
 */
export interface IExcelTableData {
  /** Column headers */
  headers: string[];
  /** Data rows */
  rows: (string | number | boolean | Date | null)[][];
  /** Column widths (optional) */
  columnWidths?: number[];
}

/**
 * Excel sheet configuration
 */
export interface IExcelSheetConfig {
  /** Sheet name */
  name: string;
  /** Table data for this sheet */
  data: IExcelTableData;
  /** Optional title above the table */
  title?: string;
  /** Optional subtitle */
  subtitle?: string;
}

/**
 * Excel Generation Service
 */
export class ExcelGenerationService {
  private readonly defaultHeaderStyle: Partial<ExcelJS.Style> = {
    font: { bold: true, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0078D4' } },
    alignment: { horizontal: 'center', vertical: 'middle' },
    border: {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } }
    }
  };

  private readonly defaultCellStyle: Partial<ExcelJS.Style> = {
    alignment: { vertical: 'middle' },
    border: {
      top: { style: 'thin', color: { argb: 'FFE1DFDD' } },
      left: { style: 'thin', color: { argb: 'FFE1DFDD' } },
      bottom: { style: 'thin', color: { argb: 'FFE1DFDD' } },
      right: { style: 'thin', color: { argb: 'FFE1DFDD' } }
    }
  };

  /**
   * Generate a simple Excel file from table data
   */
  public async generateExcel(
    tableData: IExcelTableData,
    options?: IExcelDocumentOptions
  ): Promise<IExcelGenerationResult> {
    try {
      const workbook = new ExcelJS.Workbook();

      // Set workbook properties
      workbook.creator = options?.author || 'JML Document Builder';
      workbook.created = new Date();
      workbook.modified = new Date();
      if (options?.company) {
        workbook.company = options.company;
      }

      // Create worksheet
      const sheetName = options?.sheetName || 'Data';
      const worksheet = workbook.addWorksheet(sheetName);

      // Add headers
      const headerRow = worksheet.addRow(tableData.headers);
      if (options?.styleHeaders !== false) {
        headerRow.eachCell((cell) => {
          Object.assign(cell, { style: this.defaultHeaderStyle });
        });
        headerRow.height = 25;
      }

      // Add data rows
      for (let i = 0; i < tableData.rows.length; i++) {
        const dataRow = worksheet.addRow(tableData.rows[i]);
        dataRow.eachCell((cell) => {
          Object.assign(cell, { style: this.defaultCellStyle });
        });

        // Alternate row colors
        if (i % 2 === 0) {
          dataRow.eachCell((cell) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFAF9F8' }
            };
          });
        }
      }

      // Set column widths
      if (tableData.columnWidths) {
        for (let i = 0; i < tableData.columnWidths.length; i++) {
          worksheet.getColumn(i + 1).width = tableData.columnWidths[i];
        }
      } else if (options?.autoFitColumns !== false) {
        this.autoFitColumns(worksheet, tableData);
      }

      // Freeze first row
      if (options?.freezeFirstRow !== false) {
        worksheet.views = [{ state: 'frozen', ySplit: 1 }];
      }

      // Generate buffer
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });

      const fileName = this.generateFileName(options);

      logger.info('ExcelGenerationService', `Excel generated: ${fileName}`);

      return {
        success: true,
        blob,
        fileName
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      logger.error('ExcelGenerationService', 'Failed to generate Excel:', error);

      return {
        success: false,
        error: errorMessage
      };
    }
  }

  /**
   * Generate Excel with multiple sheets
   */
  public async generateMultiSheetExcel(
    sheets: IExcelSheetConfig[],
    options?: IExcelDocumentOptions
  ): Promise<IExcelGenerationResult> {
    try {
      const workbook = new ExcelJS.Workbook();

      workbook.creator = options?.author || 'JML Document Builder';
      workbook.created = new Date();
      workbook.modified = new Date();

      for (let s = 0; s < sheets.length; s++) {
        const sheetConfig = sheets[s];
        const worksheet = workbook.addWorksheet(sheetConfig.name);

        let startRow = 1;

        // Add title if provided
        if (sheetConfig.title) {
          const titleRow = worksheet.addRow([sheetConfig.title]);
          titleRow.font = { bold: true, size: 16 };
          titleRow.height = 30;
          worksheet.mergeCells(startRow, 1, startRow, sheetConfig.data.headers.length);
          startRow++;
        }

        // Add subtitle if provided
        if (sheetConfig.subtitle) {
          const subtitleRow = worksheet.addRow([sheetConfig.subtitle]);
          subtitleRow.font = { italic: true, size: 12, color: { argb: 'FF605E5C' } };
          worksheet.mergeCells(startRow, 1, startRow, sheetConfig.data.headers.length);
          startRow++;

          // Add empty row
          worksheet.addRow([]);
          startRow++;
        }

        // Add headers
        const headerRow = worksheet.addRow(sheetConfig.data.headers);
        headerRow.eachCell((cell) => {
          Object.assign(cell, { style: this.defaultHeaderStyle });
        });
        headerRow.height = 25;

        // Add data
        for (let i = 0; i < sheetConfig.data.rows.length; i++) {
          const dataRow = worksheet.addRow(sheetConfig.data.rows[i]);
          dataRow.eachCell((cell) => {
            Object.assign(cell, { style: this.defaultCellStyle });
          });

          if (i % 2 === 0) {
            dataRow.eachCell((cell) => {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFAF9F8' }
              };
            });
          }
        }

        // Auto-fit columns
        this.autoFitColumns(worksheet, sheetConfig.data);

        // Freeze header row
        worksheet.views = [{ state: 'frozen', ySplit: startRow }];
      }

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });

      const fileName = this.generateFileName(options);

      logger.info('ExcelGenerationService', `Multi-sheet Excel generated: ${fileName}`);

      return {
        success: true,
        blob,
        fileName
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      logger.error('ExcelGenerationService', 'Failed to generate multi-sheet Excel:', error);

      return {
        success: false,
        error: errorMessage
      };
    }
  }

  /**
   * Generate employee data Excel from template context
   */
  public async generateEmployeeDataExcel(
    employees: Partial<ITemplateDataContext>[],
    options?: IExcelDocumentOptions
  ): Promise<IExcelGenerationResult> {
    const headers = [
      'Name',
      'Email',
      'Department',
      'Job Title',
      'Start Date',
      'Manager',
      'Status'
    ];

    const rows = employees.map(emp => [
      emp.employee?.name || '',
      emp.employee?.email || '',
      emp.employee?.department || '',
      emp.employee?.jobTitle || '',
      emp.employee?.startDate || '',
      emp.manager?.name || '',
      emp.process?.status || ''
    ]);

    return this.generateExcel(
      { headers, rows },
      {
        ...options,
        title: options?.title || 'Employee Data Report',
        sheetName: 'Employees'
      }
    );
  }

  /**
   * Generate process report Excel
   */
  public async generateProcessReportExcel(
    processes: Array<{
      id: number;
      type: string;
      employeeName: string;
      department: string;
      status: string;
      startDate: string;
      completionDate?: string;
      progress: number;
    }>,
    options?: IExcelDocumentOptions
  ): Promise<IExcelGenerationResult> {
    const headers = [
      'Process ID',
      'Type',
      'Employee',
      'Department',
      'Status',
      'Start Date',
      'Completion Date',
      'Progress %'
    ];

    const rows = processes.map(p => [
      p.id,
      p.type,
      p.employeeName,
      p.department,
      p.status,
      p.startDate,
      p.completionDate || 'N/A',
      p.progress
    ]);

    return this.generateExcel(
      { headers, rows },
      {
        ...options,
        title: options?.title || 'Process Report',
        sheetName: 'Processes'
      }
    );
  }

  /**
   * Auto-fit column widths based on content
   */
  private autoFitColumns(worksheet: ExcelJS.Worksheet, data: IExcelTableData): void {
    for (let i = 0; i < data.headers.length; i++) {
      let maxLength = data.headers[i].length;

      for (let j = 0; j < data.rows.length; j++) {
        const cellValue = data.rows[j][i];
        const cellLength = cellValue !== null && cellValue !== undefined
          ? String(cellValue).length
          : 0;
        maxLength = Math.max(maxLength, cellLength);
      }

      // Set column width with padding
      worksheet.getColumn(i + 1).width = Math.min(maxLength + 4, 50);
    }
  }

  /**
   * Generate file name
   */
  private generateFileName(options?: IExcelDocumentOptions): string {
    const timestamp = new Date().toISOString().slice(0, 10);
    const baseName = options?.title?.replace(/\s+/g, '_') || 'Export';
    return `${baseName}_${timestamp}.xlsx`;
  }

  /**
   * Download Excel file in browser
   */
  public downloadExcel(result: IExcelGenerationResult): void {
    if (!result.success || !result.blob) {
      logger.error('ExcelGenerationService', 'Cannot download: Excel generation failed');
      return;
    }

    const url = URL.createObjectURL(result.blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = result.fileName || 'export.xlsx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    logger.info('ExcelGenerationService', `Excel downloaded: ${result.fileName}`);
  }
}

export const excelGenerationService = new ExcelGenerationService();
