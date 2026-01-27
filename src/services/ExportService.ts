// @ts-nocheck
// Export Service
// Handles exporting analytics data to Excel, PDF, and CSV formats

import { ExportFormat, IExportOptions, IDashboardMetrics } from '../models';
import { ICV } from '../models/ICVManagement';
import { logger } from './LoggingService';

export class ExportService {
  /**
   * Export CVs to Excel
   */
  public static async exportCVsToExcel(cvs: ICV[]): Promise<void> {
    try {
      if (!cvs || cvs.length === 0) {
        throw new Error('No CVs to export');
      }

      // Transform CV data for export
      const exportData = cvs.map(cv => ({
        'Name': cv.CandidateName,
        'Email': cv.Email,
        'Phone': cv.Phone || '',
        'Status': cv.Status,
        'Position': cv.PositionAppliedFor || '',
        'Experience': `${cv.YearsOfExperience || 0} years`,
        'Experience Level': cv.ExperienceLevel || '',
        'Highest Education': cv.HighestEducation || '',
        'Skills': cv.Skills?.join(', ') || '',
        'Qualification Score': cv.QualificationScore || 0,
        'Source': cv.Source,
        'Submission Date': cv.SubmissionDate ? new Date(cv.SubmissionDate).toLocaleDateString() : '',
        'Shortlisted': cv.IsShortlisted ? 'Yes' : 'No',
        'Notes': cv.Notes || '',
        'Location': cv.Location || ''
      }));

      // Create CSV content
      const headers = Object.keys(exportData[0]);
      const rows: string[][] = [headers];

      for (const item of exportData) {
        const row: string[] = [];
        for (const header of headers) {
          let value = item[header as keyof typeof item];
          if (value === null || value === undefined) {
            value = '';
          }
          const stringValue = String(value);
          const escaped = stringValue.indexOf(',') !== -1 || stringValue.indexOf('"') !== -1
            ? `"${stringValue.replace(/"/g, '""')}"`
            : stringValue;
          row.push(escaped);
        }
        rows.push(row);
      }

      const csvContent = rows.map(row => row.join(',')).join('\n');
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });

      // Download file
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `CVs_Export_${new Date().toISOString().split('T')[0]}.csv`;
      link.style.display = 'none';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      logger.info('ExportService', `Successfully exported ${cvs.length} CVs to Excel`);
    } catch (error) {
      logger.error('ExportService', 'Failed to export CVs:', error);
      throw error;
    }
  }

  /**
   * Export data to specified format
   */
  public async exportData(
    data: any,
    filename: string,
    options: IExportOptions
  ): Promise<void> {
    switch (options.format) {
      case ExportFormat.Excel:
        await this.exportToExcel(data, filename, options);
        break;
      case ExportFormat.PDF:
        await this.exportToPDF(data, filename, options);
        break;
      case ExportFormat.CSV:
        await this.exportToCSV(data, filename);
        break;
      default:
        throw new Error(`Unsupported export format: ${options.format}`);
    }
  }

  /**
   * Export to Excel with formatting
   */
  private async exportToExcel(
    data: any,
    filename: string,
    options: IExportOptions
  ): Promise<void> {
    try {
      const workbook = this.createWorkbook(data, options);
      const blob = await this.workbookToBlob(workbook);
      this.downloadBlob(blob, `${filename}.xlsx`);
    } catch (error) {
      logger.error('ExportService', 'Failed to export to Excel:', error);
      throw error;
    }
  }

  /**
   * Export to PDF with charts
   */
  private async exportToPDF(
    data: any,
    filename: string,
    options: IExportOptions
  ): Promise<void> {
    try {
      const pdfContent = this.generatePDFContent(data, options);
      const blob = new Blob([pdfContent], { type: 'application/pdf' });
      this.downloadBlob(blob, `${filename}.pdf`);
    } catch (error) {
      logger.error('ExportService', 'Failed to export to PDF:', error);
      throw error;
    }
  }

  /**
   * Export to CSV
   */
  private async exportToCSV(data: any[], filename: string): Promise<void> {
    try {
      if (!data || data.length === 0) {
        throw new Error('No data to export');
      }

      const headers = Object.keys(data[0]);
      const rows: string[][] = [];

      rows.push(headers);

      for (let i = 0; i < data.length; i++) {
        const row: string[] = [];
        for (let j = 0; j < headers.length; j++) {
          const header = headers[j];
          let value = data[i][header];

          if (value instanceof Date) {
            value = value.toISOString().split('T')[0];
          } else if (typeof value === 'object' && value !== null) {
            value = JSON.stringify(value);
          } else if (value === null || value === undefined) {
            value = '';
          }

          const stringValue = String(value);
          const escaped = stringValue.indexOf(',') !== -1 || stringValue.indexOf('"') !== -1
            ? `"${stringValue.replace(/"/g, '""')}"`
            : stringValue;

          row.push(escaped);
        }
        rows.push(row);
      }

      const csvContent = rows.map(row => row.join(',')).join('\n');
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      this.downloadBlob(blob, `${filename}.csv`);
    } catch (error) {
      logger.error('ExportService', 'Failed to export to CSV:', error);
      throw error;
    }
  }

  /**
   * Create Excel workbook
   */
  private createWorkbook(data: any, options: IExportOptions): any {
    const workbook: any = {
      sheets: []
    };

    if (options.includeSummary && data.summary) {
      workbook.sheets.push(this.createSummarySheet(data.summary));
    }

    if (data.trends) {
      workbook.sheets.push(this.createDataSheet('Completion Trends', data.trends));
    }

    if (data.costs) {
      workbook.sheets.push(this.createDataSheet('Cost Analysis', data.costs));
    }

    if (data.bottlenecks) {
      workbook.sheets.push(this.createDataSheet('Task Bottlenecks', data.bottlenecks));
    }

    if (data.workload) {
      workbook.sheets.push(this.createDataSheet('Manager Workload', data.workload));
    }

    if (data.compliance) {
      workbook.sheets.push(this.createDataSheet('Compliance', data.compliance));
    }

    if (data.sla) {
      workbook.sheets.push(this.createDataSheet('SLA Metrics', data.sla));
    }

    return workbook;
  }

  /**
   * Create summary sheet
   */
  private createSummarySheet(summary: IDashboardMetrics): any {
    const rows = [
      ['Metric', 'Value'],
      ['Total Processes', summary.totalProcesses],
      ['Completed Processes', summary.completedProcesses],
      ['Active Processes', summary.activeProcesses],
      ['Overdue Processes', summary.overdueProcesses],
      ['Average Completion Time (days)', summary.averageCompletionTime.toFixed(1)],
      ['Total Cost', `$${summary.totalCost.toFixed(2)}`],
      ['Compliance Rate (%)', summary.complianceRate.toFixed(1)],
      ['NPS Score', summary.npsScore.toFixed(1)],
      ['SLA Adherence (%)', summary.slaAdherence.toFixed(1)],
      ['First-Day Readiness (%)', summary.firstDayReadiness.toFixed(1)]
    ];

    return {
      name: 'Summary',
      data: rows
    };
  }

  /**
   * Create data sheet
   */
  private createDataSheet(sheetName: string, data: any[]): any {
    if (!data || data.length === 0) {
      return {
        name: sheetName,
        data: [[`No ${sheetName} data available`]]
      };
    }

    const headers = Object.keys(data[0]);
    const rows: any[][] = [headers];

    for (let i = 0; i < data.length; i++) {
      const row: any[] = [];
      for (let j = 0; j < headers.length; j++) {
        const header = headers[j];
        let value = data[i][header];

        if (value instanceof Date) {
          value = value.toISOString().split('T')[0];
        } else if (typeof value === 'object' && value !== null) {
          value = JSON.stringify(value);
        }

        row.push(value);
      }
      rows.push(row);
    }

    return {
      name: sheetName,
      data: rows
    };
  }

  /**
   * Convert workbook to blob
   */
  private async workbookToBlob(workbook: any): Promise<Blob> {
    const content = this.workbookToCSVFormat(workbook);
    return new Blob([content], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  }

  /**
   * Convert workbook to CSV format (simplified Excel export)
   */
  private workbookToCSVFormat(workbook: any): string {
    let content = '';

    for (let i = 0; i < workbook.sheets.length; i++) {
      const sheet = workbook.sheets[i];
      content += `\n\n=== ${sheet.name} ===\n\n`;

      for (let j = 0; j < sheet.data.length; j++) {
        const row = sheet.data[j];
        content += row.join(',') + '\n';
      }
    }

    return content;
  }

  /**
   * Generate PDF content
   */
  private generatePDFContent(data: any, options: IExportOptions): string {
    let content = 'JML Analytics Report\n\n';
    content += `Generated: ${new Date().toLocaleString()}\n\n`;

    if (options.includeSummary && data.summary) {
      content += this.formatSummaryForPDF(data.summary);
    }

    if (data.trends) {
      content += '\n\nCompletion Trends:\n';
      content += this.formatDataForPDF(data.trends);
    }

    if (data.costs) {
      content += '\n\nCost Analysis:\n';
      content += this.formatDataForPDF(data.costs);
    }

    if (data.bottlenecks) {
      content += '\n\nTask Bottlenecks:\n';
      content += this.formatDataForPDF(data.bottlenecks);
    }

    return content;
  }

  /**
   * Format summary for PDF
   */
  private formatSummaryForPDF(summary: IDashboardMetrics): string {
    let content = 'Summary Metrics:\n';
    content += `Total Processes: ${summary.totalProcesses}\n`;
    content += `Completed Processes: ${summary.completedProcesses}\n`;
    content += `Active Processes: ${summary.activeProcesses}\n`;
    content += `Overdue Processes: ${summary.overdueProcesses}\n`;
    content += `Average Completion Time: ${summary.averageCompletionTime.toFixed(1)} days\n`;
    content += `Total Cost: $${summary.totalCost.toFixed(2)}\n`;
    content += `Compliance Rate: ${summary.complianceRate.toFixed(1)}%\n`;
    content += `NPS Score: ${summary.npsScore.toFixed(1)}\n`;
    content += `SLA Adherence: ${summary.slaAdherence.toFixed(1)}%\n`;
    content += `First-Day Readiness: ${summary.firstDayReadiness.toFixed(1)}%\n`;
    return content;
  }

  /**
   * Format data for PDF
   */
  private formatDataForPDF(data: any[]): string {
    if (!data || data.length === 0) {
      return 'No data available\n';
    }

    let content = '';
    const headers = Object.keys(data[0]);

    content += headers.join(' | ') + '\n';
    content += headers.map(() => '---').join(' | ') + '\n';

    for (let i = 0; i < Math.min(data.length, 50); i++) {
      const row: string[] = [];
      for (let j = 0; j < headers.length; j++) {
        const header = headers[j];
        let value = data[i][header];

        if (value instanceof Date) {
          value = value.toISOString().split('T')[0];
        } else if (typeof value === 'object' && value !== null) {
          value = JSON.stringify(value);
        } else if (value === null || value === undefined) {
          value = '';
        }

        row.push(String(value));
      }
      content += row.join(' | ') + '\n';
    }

    return content;
  }

  /**
   * Download blob as file
   */
  private downloadBlob(blob: Blob, filename: string): void {
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
  }

  /**
   * Generate PowerBI embed URL
   */
  public generatePowerBIEmbedUrl(
    workspaceId: string,
    reportId: string,
    filters?: any
  ): string {
    let url = `https://app.powerbi.com/reportEmbed?reportId=${reportId}&groupId=${workspaceId}`;

    if (filters) {
      const filterString = encodeURIComponent(JSON.stringify(filters));
      url += `&$filter=${filterString}`;
    }

    return url;
  }
}
