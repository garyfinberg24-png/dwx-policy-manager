// @ts-nocheck
// ROI Export Service
// Exports ROI analytics data to Excel and PDF formats

import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  IROISummary,
  IBeforeAfterMetrics,
  IEmployeeLookupMetrics,
  ITaskAutomationMetrics,
  INotificationMetrics,
  IApprovalWorkflowMetrics
} from './ROIAnalyticsService';
import { logger } from './LoggingService';

export interface IExportOptions {
  includeDetailedMetrics?: boolean;
  includeCharts?: boolean;
  companyName?: string;
  reportTitle?: string;
  logoUrl?: string;
}

export class ROIExportService {
  /**
   * Export ROI summary to Excel workbook
   */
  public exportToExcel(
    roiSummary: IROISummary,
    beforeAfter: IBeforeAfterMetrics[],
    detailedMetrics?: {
      employeeLookup?: IEmployeeLookupMetrics;
      taskAutomation?: ITaskAutomationMetrics;
      notification?: INotificationMetrics;
      approval?: IApprovalWorkflowMetrics;
    },
    options: IExportOptions = {}
  ): Blob {
    try {
      const workbook = XLSX.utils.book_new();

      // Sheet 1: Executive Summary
      const summarySheet = this.createSummarySheet(roiSummary, options);
      XLSX.utils.book_append_sheet(workbook, summarySheet, 'Executive Summary');

      // Sheet 2: Feature Breakdown
      const breakdownSheet = this.createFeatureBreakdownSheet(roiSummary);
      XLSX.utils.book_append_sheet(workbook, breakdownSheet, 'Feature Breakdown');

      // Sheet 3: Before vs After
      const comparisonSheet = this.createComparisonSheet(beforeAfter);
      XLSX.utils.book_append_sheet(workbook, comparisonSheet, 'Before vs After');

      // Sheet 4: Detailed Metrics (if included)
      if (options.includeDetailedMetrics && detailedMetrics) {
        if (detailedMetrics.employeeLookup) {
          const lookupSheet = this.createEmployeeLookupSheet(detailedMetrics.employeeLookup);
          XLSX.utils.book_append_sheet(workbook, lookupSheet, 'Employee Lookup Metrics');
        }

        if (detailedMetrics.taskAutomation) {
          const taskSheet = this.createTaskAutomationSheet(detailedMetrics.taskAutomation);
          XLSX.utils.book_append_sheet(workbook, taskSheet, 'Task Automation Metrics');
        }

        if (detailedMetrics.notification) {
          const notificationSheet = this.createNotificationSheet(detailedMetrics.notification);
          XLSX.utils.book_append_sheet(workbook, notificationSheet, 'Notification Metrics');
        }

        if (detailedMetrics.approval) {
          const approvalSheet = this.createApprovalSheet(detailedMetrics.approval);
          XLSX.utils.book_append_sheet(workbook, approvalSheet, 'Approval Metrics');
        }
      }

      // Generate Excel file
      const excelBuffer = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

      logger.info('ROIExportService', 'Excel export generated successfully');
      return blob;
    } catch (error) {
      logger.error('ROIExportService', 'Failed to export to Excel:', error);
      throw error;
    }
  }

  /**
   * Export ROI summary to PDF
   */
  public exportToPDF(
    roiSummary: IROISummary,
    beforeAfter: IBeforeAfterMetrics[],
    options: IExportOptions = {}
  ): Blob {
    try {
      const doc = new jsPDF('p', 'mm', 'a4');
      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      let yPosition = 20;

      // Header
      doc.setFontSize(20);
      doc.setTextColor(0, 120, 212); // Microsoft blue
      const title = options.reportTitle || 'JML Solution - ROI Analysis Report';
      doc.text(title, pageWidth / 2, yPosition, { align: 'center' });
      yPosition += 10;

      doc.setFontSize(10);
      doc.setTextColor(100, 100, 100);
      const companyName = options.companyName || 'Your Organization';
      doc.text(companyName, pageWidth / 2, yPosition, { align: 'center' });
      yPosition += 5;

      const dateRange = `Period: ${this.formatDate(roiSummary.startDate)} - ${this.formatDate(roiSummary.endDate)}`;
      doc.text(dateRange, pageWidth / 2, yPosition, { align: 'center' });
      yPosition += 15;

      // Key Metrics Section
      doc.setFontSize(14);
      doc.setTextColor(0, 0, 0);
      doc.text('Key Metrics', 14, yPosition);
      yPosition += 5;

      const keyMetrics = [
        ['Metric', 'Value'],
        ['Total Cost Savings', this.formatCurrency(roiSummary.totalCostSavings)],
        ['Annualized Savings', this.formatCurrency(roiSummary.annualizedSavings)],
        ['Return on Investment', `${roiSummary.roi.toFixed(1)}%`],
        ['Payback Period', `${roiSummary.paybackMonths.toFixed(1)} months`],
        ['Total Hours Saved', `${roiSummary.totalHoursSaved.toFixed(1)} hours`],
        ['FTE Equivalent', `${roiSummary.fteEquivalent.toFixed(2)} FTE`],
        ['Automation Adoption', `${roiSummary.automationAdoptionRate.toFixed(1)}%`],
      ];

      autoTable(doc, {
        startY: yPosition,
        head: [keyMetrics[0]],
        body: keyMetrics.slice(1),
        theme: 'striped',
        headStyles: { fillColor: [0, 120, 212] },
        margin: { left: 14, right: 14 },
      });

      yPosition = (doc as any).lastAutoTable.finalY + 15;

      // Feature Breakdown Section
      if (yPosition > pageHeight - 60) {
        doc.addPage();
        yPosition = 20;
      }

      doc.setFontSize(14);
      doc.text('Savings by Feature', 14, yPosition);
      yPosition += 5;

      const featureBreakdown = [
        ['Feature', 'Hours Saved', 'Cost Savings', '% of Total'],
        [
          'Employee Master Data & Lookup',
          `${roiSummary.employeeLookupSavings.hours.toFixed(1)}`,
          this.formatCurrency(roiSummary.employeeLookupSavings.cost),
          `${roiSummary.employeeLookupSavings.percentage.toFixed(1)}%`
        ],
        [
          'Automatic Task Generation',
          `${roiSummary.taskAutomationSavings.hours.toFixed(1)}`,
          this.formatCurrency(roiSummary.taskAutomationSavings.cost),
          `${roiSummary.taskAutomationSavings.percentage.toFixed(1)}%`
        ],
        [
          'Smart Notifications & Reminders',
          `${roiSummary.notificationSavings.hours.toFixed(1)}`,
          this.formatCurrency(roiSummary.notificationSavings.cost),
          `${roiSummary.notificationSavings.percentage.toFixed(1)}%`
        ],
        [
          'Approval Workflows',
          `${roiSummary.approvalSavings.hours.toFixed(1)}`,
          this.formatCurrency(roiSummary.approvalSavings.cost),
          `${roiSummary.approvalSavings.percentage.toFixed(1)}%`
        ],
      ];

      autoTable(doc, {
        startY: yPosition,
        head: [featureBreakdown[0]],
        body: featureBreakdown.slice(1),
        theme: 'striped',
        headStyles: { fillColor: [0, 120, 212] },
        margin: { left: 14, right: 14 },
      });

      yPosition = (doc as any).lastAutoTable.finalY + 15;

      // Before vs After Section
      if (yPosition > pageHeight - 60) {
        doc.addPage();
        yPosition = 20;
      }

      doc.setFontSize(14);
      doc.text('Before vs After Automation', 14, yPosition);
      yPosition += 5;

      const comparisonData = [
        ['Metric', 'Before', 'After', 'Improvement'],
        ...beforeAfter.map(metric => [
          metric.metricName,
          `${metric.beforeValue.toFixed(1)} ${metric.unit}`,
          `${metric.afterValue.toFixed(1)} ${metric.unit}`,
          `${metric.improvementPercentage.toFixed(1)}%`
        ])
      ];

      autoTable(doc, {
        startY: yPosition,
        head: [comparisonData[0]],
        body: comparisonData.slice(1),
        theme: 'striped',
        headStyles: { fillColor: [0, 120, 212] },
        margin: { left: 14, right: 14 },
      });

      yPosition = (doc as any).lastAutoTable.finalY + 15;

      // Quality Improvements Section
      if (yPosition > pageHeight - 60) {
        doc.addPage();
        yPosition = 20;
      }

      doc.setFontSize(14);
      doc.text('Quality Improvements', 14, yPosition);
      yPosition += 5;

      const qualityMetrics = [
        ['Quality Metric', 'Score'],
        ['Error Reduction Rate', `${roiSummary.errorReductionRate.toFixed(1)}%`],
        ['Compliance Improvement', `${roiSummary.complianceImprovementRate.toFixed(1)}%`],
        ['User Satisfaction', `${roiSummary.userSatisfactionScore.toFixed(1)}/10`],
      ];

      autoTable(doc, {
        startY: yPosition,
        head: [qualityMetrics[0]],
        body: qualityMetrics.slice(1),
        theme: 'striped',
        headStyles: { fillColor: [0, 120, 212] },
        margin: { left: 14, right: 14 },
      });

      // Footer
      const pageCount = (doc as any).internal.getNumberOfPages();
      for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.setTextColor(150, 150, 150);
        doc.text(
          `Generated: ${new Date().toLocaleDateString()} | Page ${i} of ${pageCount}`,
          pageWidth / 2,
          pageHeight - 10,
          { align: 'center' }
        );
      }

      const pdfBlob = doc.output('blob');
      logger.info('ROIExportService', 'PDF export generated successfully');
      return pdfBlob;
    } catch (error) {
      logger.error('ROIExportService', 'Failed to export to PDF:', error);
      throw error;
    }
  }

  /**
   * Create Executive Summary sheet for Excel
   */
  private createSummarySheet(roiSummary: IROISummary, options: IExportOptions): XLSX.WorkSheet {
    const data = [
      [options.reportTitle || 'JML Solution - ROI Analysis Report'],
      [options.companyName || 'Your Organization'],
      [`Period: ${this.formatDate(roiSummary.startDate)} - ${this.formatDate(roiSummary.endDate)}`],
      [],
      ['KEY METRICS'],
      ['Metric', 'Value'],
      ['Total Cost Savings', this.formatCurrency(roiSummary.totalCostSavings)],
      ['Annualized Savings', this.formatCurrency(roiSummary.annualizedSavings)],
      ['Return on Investment', `${roiSummary.roi.toFixed(1)}%`],
      ['Payback Period', `${roiSummary.paybackMonths.toFixed(1)} months`],
      ['Total Hours Saved', `${roiSummary.totalHoursSaved.toFixed(1)} hours`],
      ['FTE Equivalent', `${roiSummary.fteEquivalent.toFixed(2)} FTE`],
      ['Automation Adoption Rate', `${roiSummary.automationAdoptionRate.toFixed(1)}%`],
      [],
      ['QUALITY IMPROVEMENTS'],
      ['Metric', 'Score'],
      ['Error Reduction Rate', `${roiSummary.errorReductionRate.toFixed(1)}%`],
      ['Compliance Improvement', `${roiSummary.complianceImprovementRate.toFixed(1)}%`],
      ['User Satisfaction', `${roiSummary.userSatisfactionScore.toFixed(1)}/10`],
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(data);

    // Set column widths
    worksheet['!cols'] = [
      { wch: 30 },
      { wch: 20 }
    ];

    return worksheet;
  }

  /**
   * Create Feature Breakdown sheet for Excel
   */
  private createFeatureBreakdownSheet(roiSummary: IROISummary): XLSX.WorkSheet {
    const data = [
      ['SAVINGS BY FEATURE'],
      [],
      ['Feature', 'Hours Saved', 'Cost Savings', '% of Total'],
      [
        'Employee Master Data & Lookup',
        roiSummary.employeeLookupSavings.hours.toFixed(1),
        this.formatCurrency(roiSummary.employeeLookupSavings.cost),
        `${roiSummary.employeeLookupSavings.percentage.toFixed(1)}%`
      ],
      [
        'Automatic Task Generation',
        roiSummary.taskAutomationSavings.hours.toFixed(1),
        this.formatCurrency(roiSummary.taskAutomationSavings.cost),
        `${roiSummary.taskAutomationSavings.percentage.toFixed(1)}%`
      ],
      [
        'Smart Notifications & Reminders',
        roiSummary.notificationSavings.hours.toFixed(1),
        this.formatCurrency(roiSummary.notificationSavings.cost),
        `${roiSummary.notificationSavings.percentage.toFixed(1)}%`
      ],
      [
        'Approval Workflows',
        roiSummary.approvalSavings.hours.toFixed(1),
        this.formatCurrency(roiSummary.approvalSavings.cost),
        `${roiSummary.approvalSavings.percentage.toFixed(1)}%`
      ],
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(data);

    worksheet['!cols'] = [
      { wch: 35 },
      { wch: 15 },
      { wch: 15 },
      { wch: 12 }
    ];

    return worksheet;
  }

  /**
   * Create Before vs After Comparison sheet for Excel
   */
  private createComparisonSheet(beforeAfter: IBeforeAfterMetrics[]): XLSX.WorkSheet {
    const data = [
      ['BEFORE VS AFTER AUTOMATION'],
      [],
      ['Metric', 'Before', 'After', 'Improvement', 'Improvement %'],
      ...beforeAfter.map(metric => [
        metric.metricName,
        `${metric.beforeValue.toFixed(1)} ${metric.unit}`,
        `${metric.afterValue.toFixed(1)} ${metric.unit}`,
        `${metric.improvement.toFixed(1)} ${metric.unit}`,
        `${metric.improvementPercentage.toFixed(1)}%`
      ])
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(data);

    worksheet['!cols'] = [
      { wch: 30 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 },
      { wch: 15 }
    ];

    return worksheet;
  }

  /**
   * Create Employee Lookup Metrics sheet for Excel
   */
  private createEmployeeLookupSheet(metrics: IEmployeeLookupMetrics): XLSX.WorkSheet {
    const data = [
      ['EMPLOYEE MASTER DATA & LOOKUP METRICS'],
      [],
      ['Usage Statistics'],
      ['Total Processes Created', metrics.totalProcessesCreated],
      ['Processes Using Employee Picker', metrics.processesUsingEmployeePicker],
      ['Employee Picker Usage Rate', `${metrics.employeePickerUsageRate.toFixed(1)}%`],
      [],
      ['Time Savings'],
      ['Average Manual Entry Time', `${metrics.averageManualEntryTimeMinutes} minutes`],
      ['Average Lookup Time', `${metrics.averageLookupTimeSeconds} seconds`],
      ['Time Saved Per Process', `${metrics.timeSavedPerProcess.toFixed(1)} minutes`],
      ['Total Time Saved', `${metrics.totalTimeSavedHours.toFixed(1)} hours`],
      [],
      ['Data Quality'],
      ['Data Accuracy Rate', `${metrics.dataAccuracyRate}%`],
      ['Duplicate Employees Prevented', metrics.duplicateEmployeesPrevented],
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(data);
    worksheet['!cols'] = [{ wch: 35 }, { wch: 20 }];
    return worksheet;
  }

  /**
   * Create Task Automation Metrics sheet for Excel
   */
  private createTaskAutomationSheet(metrics: ITaskAutomationMetrics): XLSX.WorkSheet {
    const data = [
      ['TASK AUTOMATION METRICS'],
      [],
      ['Usage Statistics'],
      ['Processes with Templates', metrics.processesWithTemplates],
      ['Total Tasks Generated', metrics.totalTasksGenerated],
      ['Average Tasks Per Process', metrics.averageTasksPerProcess.toFixed(1)],
      [],
      ['Time Savings'],
      ['Manual Task Creation Time', `${metrics.manualTaskCreationTimeMinutes} minutes per task`],
      ['Automated Creation Time', `${metrics.automatedTaskCreationSeconds} seconds`],
      ['Time Saved Per Process', `${metrics.timeSavedPerProcess.toFixed(1)} minutes`],
      ['Total Time Saved', `${metrics.totalTimeSavedHours.toFixed(1)} hours`],
      [],
      ['Automation Accuracy'],
      ['Task Dependencies Automated', metrics.taskDependenciesAutomated],
      ['SLA Calculations Automated', metrics.slaCalculationsAutomated],
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(data);
    worksheet['!cols'] = [{ wch: 35 }, { wch: 20 }];
    return worksheet;
  }

  /**
   * Create Notification Metrics sheet for Excel
   */
  private createNotificationSheet(metrics: INotificationMetrics): XLSX.WorkSheet {
    const data = [
      ['SMART NOTIFICATIONS & REMINDERS METRICS'],
      [],
      ['Notification Activity'],
      ['Reminders Generated', metrics.remindersGenerated],
      ['Escalations Sent', metrics.escalationsSent],
      ['Digests Sent', metrics.digestsSent],
      [],
      ['Effectiveness'],
      ['Tasks Completed After Reminder', metrics.tasksCompletedAfterReminder],
      ['Average Response Time', `${metrics.averageResponseTimeHours.toFixed(1)} hours`],
      ['Overdue Task Reduction', `${metrics.overdueTaskReduction.toFixed(1)}%`],
      [],
      ['User Engagement'],
      ['User Preferences Configured', metrics.userPreferencesConfigured],
      ['Notification Open Rate', `${metrics.notificationOpenRate.toFixed(1)}%`],
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(data);
    worksheet['!cols'] = [{ wch: 35 }, { wch: 20 }];
    return worksheet;
  }

  /**
   * Create Approval Workflow Metrics sheet for Excel
   */
  private createApprovalSheet(metrics: IApprovalWorkflowMetrics): XLSX.WorkSheet {
    const data = [
      ['APPROVAL WORKFLOW METRICS'],
      [],
      ['Usage Statistics'],
      ['Processes Requiring Approval', metrics.processesRequiringApproval],
      ['Total Approval Requests', metrics.totalApprovalRequests],
      ['Average Approval Chain Levels', metrics.approvalChainLevels.toFixed(1)],
      [],
      ['Performance'],
      ['Average Approval Time', `${metrics.averageApprovalTimeHours.toFixed(1)} hours`],
      ['SLA Compliance Rate', `${metrics.approvalSLAComplianceRate.toFixed(1)}%`],
      ['Overdue Approvals', metrics.overdueApprovalsCount],
      [],
      ['Actions'],
      ['Approved', metrics.approvedCount],
      ['Rejected', metrics.rejectedCount],
      ['Delegated', metrics.delegatedCount],
      ['Approval Rate', `${metrics.approvalRate.toFixed(1)}%`],
      ['Delegation Rate', `${metrics.delegationRate.toFixed(1)}%`],
      [],
      ['Time Savings'],
      ['Automated Notifications', metrics.automatedNotifications],
      ['Manual Follow-ups Saved', metrics.manualFollowUpsSaved],
      ['Time Saved on Follow-ups', `${metrics.timeSavedOnFollowUpsHours.toFixed(1)} hours`],
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(data);
    worksheet['!cols'] = [{ wch: 35 }, { wch: 20 }];
    return worksheet;
  }

  /**
   * Helper: Format currency
   */
  private formatCurrency(value: number): string {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: 0,
      maximumFractionDigits: 0
    }).format(value);
  }

  /**
   * Helper: Format date
   */
  private formatDate(date: Date): string {
    return new Intl.DateTimeFormat('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    }).format(date);
  }
}
