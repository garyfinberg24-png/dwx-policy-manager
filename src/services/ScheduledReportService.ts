// @ts-nocheck
// Scheduled Report Service
// Manages scheduled email reports

import { SPFI } from '@pnp/sp';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { IScheduledReport, ReportFrequency, ExportFormat } from '../models';
import { GraphService } from './GraphService';
import { AnalyticsService } from './AnalyticsService';
import { ExportService } from './ExportService';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class ScheduledReportService {
  private sp: SPFI;
  private graphService: GraphService;
  private analyticsService: AnalyticsService;
  private exportService: ExportService;

  constructor(
    sp: SPFI,
    graphService: GraphService,
    analyticsService: AnalyticsService
  ) {
    this.sp = sp;
    this.graphService = graphService;
    this.analyticsService = analyticsService;
    this.exportService = new ExportService();
  }

  /**
   * Get all scheduled reports
   */
  public async getScheduledReports(userId?: number): Promise<IScheduledReport[]> {
    try {
      // Validate user ID if provided
      if (userId !== undefined) {
        ValidationUtils.validateInteger(userId, 'userId', 1);
      }

      // Build secure filter
      let filterQuery: string | undefined = undefined;
      if (userId) {
        filterQuery = ValidationUtils.buildFilter('CreatedBy', 'eq', userId);
      }

      let query = this.sp.web.lists.getByTitle('JML_ScheduledReports').items
        .select('*')
        .orderBy('NextRun', true);

      if (filterQuery) {
        query = query.filter(filterQuery);
      }

      const items = await query();
      const reports: IScheduledReport[] = [];

      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        reports.push({
          id: item.ReportId || item.Id.toString(),
          reportName: item.Title,
          reportType: item.ReportType,
          frequency: item.Frequency as ReportFrequency,
          format: item.Format as ExportFormat,
          recipients: JSON.parse(item.Recipients || '[]'),
          filters: JSON.parse(item.Filters || '{}'),
          enabled: item.Enabled,
          lastRun: item.LastRun ? new Date(item.LastRun) : undefined,
          nextRun: item.NextRun ? new Date(item.NextRun) : undefined,
          createdBy: item.CreatedById,
          createdDate: new Date(item.Created)
        });
      }

      return reports;
    } catch (error) {
      logger.error('ScheduledReportService', 'Failed to get scheduled reports:', error);
      return [];
    }
  }

  /**
   * Create scheduled report
   */
  public async createScheduledReport(report: IScheduledReport): Promise<void> {
    try {
      const nextRun = this.calculateNextRun(report.frequency);

      await this.sp.web.lists.getByTitle('JML_ScheduledReports').items.add({
        Title: report.reportName,
        ReportId: report.id,
        ReportType: report.reportType,
        Frequency: report.frequency,
        Format: report.format,
        Recipients: JSON.stringify(report.recipients),
        Filters: JSON.stringify(report.filters || {}),
        Enabled: report.enabled,
        NextRun: nextRun.toISOString()
      });
    } catch (error) {
      logger.error('ScheduledReportService', 'Failed to create scheduled report:', error);
      throw error;
    }
  }

  /**
   * Update scheduled report
   */
  public async updateScheduledReport(report: IScheduledReport): Promise<void> {
    try {
      // Validate and sanitize report ID
      if (!report.id || typeof report.id !== 'string') {
        throw new Error('Invalid report ID');
      }

      // Build secure filter
      const filter = ValidationUtils.buildFilter('ReportId', 'eq', report.id.substring(0, 100));

      const items = await this.sp.web.lists.getByTitle('JML_ScheduledReports').items
        .filter(filter)
        .top(1)();

      if (items.length === 0) {
        throw new Error('Scheduled report not found');
      }

      await this.sp.web.lists.getByTitle('JML_ScheduledReports').items.getById(items[0].Id).update({
        Title: report.reportName,
        ReportType: report.reportType,
        Frequency: report.frequency,
        Format: report.format,
        Recipients: JSON.stringify(report.recipients),
        Filters: JSON.stringify(report.filters || {}),
        Enabled: report.enabled
      });
    } catch (error) {
      logger.error('ScheduledReportService', 'Failed to update scheduled report:', error);
      throw error;
    }
  }

  /**
   * Delete scheduled report
   */
  public async deleteScheduledReport(reportId: string): Promise<void> {
    try {
      // Validate and sanitize report ID
      if (!reportId || typeof reportId !== 'string') {
        throw new Error('Invalid report ID');
      }

      // Build secure filter
      const filter = ValidationUtils.buildFilter('ReportId', 'eq', reportId.substring(0, 100));

      const items = await this.sp.web.lists.getByTitle('JML_ScheduledReports').items
        .filter(filter)
        .top(1)();

      if (items.length > 0) {
        await this.sp.web.lists.getByTitle('JML_ScheduledReports').items.getById(items[0].Id).delete();
      }
    } catch (error) {
      logger.error('ScheduledReportService', 'Failed to delete scheduled report:', error);
      throw error;
    }
  }

  /**
   * Execute scheduled report
   */
  public async executeScheduledReport(reportId: string): Promise<void> {
    try {
      const reports = await this.getScheduledReports();
      let report: IScheduledReport | undefined;

      for (let i = 0; i < reports.length; i++) {
        if (reports[i].id === reportId) {
          report = reports[i];
          break;
        }
      }

      if (!report) {
        throw new Error('Scheduled report not found');
      }

      const data = await this.generateReportData(report);

      const blob = await this.generateReportFile(data, report);

      await this.sendReportEmail(report, blob);

      await this.updateLastRun(report);
    } catch (error) {
      logger.error('ScheduledReportService', 'Failed to execute scheduled report:', error);
      throw error;
    }
  }

  /**
   * Check for due reports and execute them
   */
  public async processDueReports(): Promise<void> {
    try {
      const reports = await this.getScheduledReports();
      const now = new Date();

      for (let i = 0; i < reports.length; i++) {
        const report = reports[i];
        if (!report.enabled) {
          continue;
        }

        if (report.nextRun && report.nextRun <= now) {
          await this.executeScheduledReport(report.id);
        }
      }
    } catch (error) {
      logger.error('ScheduledReportService', 'Failed to process due reports:', error);
    }
  }

  // Private helper methods

  private calculateNextRun(frequency: ReportFrequency): Date {
    const now = new Date();
    const nextRun = new Date(now.getTime());

    switch (frequency) {
      case ReportFrequency.Daily:
        nextRun.setDate(nextRun.getDate() + 1);
        nextRun.setHours(8, 0, 0, 0);
        break;
      case ReportFrequency.Weekly:
        nextRun.setDate(nextRun.getDate() + 7);
        nextRun.setHours(8, 0, 0, 0);
        break;
      case ReportFrequency.Monthly:
        nextRun.setMonth(nextRun.getMonth() + 1);
        nextRun.setDate(1);
        nextRun.setHours(8, 0, 0, 0);
        break;
      case ReportFrequency.Quarterly:
        nextRun.setMonth(nextRun.getMonth() + 3);
        nextRun.setDate(1);
        nextRun.setHours(8, 0, 0, 0);
        break;
    }

    return nextRun;
  }

  private async generateReportData(report: IScheduledReport): Promise<any> {
    const filters = report.filters;

    const summary = await this.analyticsService.getDashboardMetrics(filters);
    const trends = await this.analyticsService.getCompletionTrends(filters);
    const costs = await this.analyticsService.getCostAnalysis(filters);
    const bottlenecks = await this.analyticsService.getTaskBottlenecks(filters);
    const workload = await this.analyticsService.getManagerWorkload(filters);
    const compliance = await this.analyticsService.getComplianceScores(filters);
    const sla = await this.analyticsService.getSLAMetrics(filters);

    return {
      summary,
      trends,
      costs,
      bottlenecks,
      workload,
      compliance,
      sla
    };
  }

  private async generateReportFile(data: any, report: IScheduledReport): Promise<Blob> {
    const filename = `${report.reportName}_${new Date().toISOString().split('T')[0]}`;

    switch (report.format) {
      case ExportFormat.Excel:
        return await this.exportService['exportToExcel'](data, filename, {
          format: ExportFormat.Excel,
          includeCharts: true,
          includeSummary: true
        }) as any;
      case ExportFormat.PDF:
        return await this.exportService['exportToPDF'](data, filename, {
          format: ExportFormat.PDF,
          includeCharts: true,
          includeSummary: true
        }) as any;
      case ExportFormat.CSV:
        return await this.exportService['exportToCSV'](data.summary, filename) as any;
      default:
        throw new Error(`Unsupported format: ${report.format}`);
    }
  }

  private async sendReportEmail(report: IScheduledReport, attachment: Blob): Promise<void> {
    try {
      const subject = `${report.reportName} - ${new Date().toLocaleDateString()}`;
      const body = this.generateEmailBody(report);

      const attachmentData = await this.blobToBase64(attachment);
      const filename = `${report.reportName}.${this.getFileExtension(report.format)}`;

      for (let i = 0; i < report.recipients.length; i++) {
        const recipient = report.recipients[i];
        await this.graphService.sendEmail(
          recipient,
          subject,
          body,
          [{
            name: filename,
            contentBytes: attachmentData
          }]
        );
      }
    } catch (error) {
      logger.error('ScheduledReportService', 'Failed to send report email:', error);
      throw error;
    }
  }

  private generateEmailBody(report: IScheduledReport): string {
    let body = '<html><body>';
    body += `<h2>${report.reportName}</h2>`;
    body += `<p>Your scheduled ${report.frequency.toLowerCase()} report is attached.</p>`;
    body += '<p>This report includes:</p>';
    body += '<ul>';
    body += '<li>Dashboard Summary Metrics</li>';
    body += '<li>Completion Trends</li>';
    body += '<li>Cost Analysis</li>';
    body += '<li>Task Bottlenecks</li>';
    body += '<li>Manager Workload Distribution</li>';
    body += '<li>Compliance Scores</li>';
    body += '<li>SLA Adherence Metrics</li>';
    body += '</ul>';
    body += `<p><small>Generated: ${new Date().toLocaleString()}</small></p>`;
    body += '</body></html>';
    return body;
  }

  private async updateLastRun(report: IScheduledReport): Promise<void> {
    try {
      // Validate and sanitize report ID
      if (!report.id || typeof report.id !== 'string') {
        throw new Error('Invalid report ID');
      }

      // Build secure filter
      const filter = ValidationUtils.buildFilter('ReportId', 'eq', report.id.substring(0, 100));

      const items = await this.sp.web.lists.getByTitle('JML_ScheduledReports').items
        .filter(filter)
        .top(1)();

      if (items.length === 0) {
        return;
      }

      const now = new Date();
      const nextRun = this.calculateNextRun(report.frequency);

      await this.sp.web.lists.getByTitle('JML_ScheduledReports').items.getById(items[0].Id).update({
        LastRun: now.toISOString(),
        NextRun: nextRun.toISOString()
      });
    } catch (error) {
      logger.error('ScheduledReportService', 'Failed to update last run:', error);
    }
  }

  private async blobToBase64(blob: Blob): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        const base64 = reader.result as string;
        const base64Data = base64.split(',')[1];
        resolve(base64Data);
      };
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  private getFileExtension(format: ExportFormat): string {
    switch (format) {
      case ExportFormat.Excel:
        return 'xlsx';
      case ExportFormat.PDF:
        return 'pdf';
      case ExportFormat.CSV:
        return 'csv';
      default:
        return 'txt';
    }
  }
}
