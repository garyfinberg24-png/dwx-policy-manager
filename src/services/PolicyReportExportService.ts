// @ts-nocheck
// Policy Report Export Service
// Exports policy data to Excel/CSV formats for reporting and analysis

import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../utils/pnpConfig';
import { logger } from './LoggingService';
import { PolicyLists, QuizLists, ApprovalLists } from '../constants/SharePointListNames';
import { ReportHtmlGenerator, IReportConfig, IReportSection, ITableData, IKpiItem } from '../utils/reportHtmlGenerator';
import {
  IPolicy,
  IPolicyAcknowledgement,
  IPolicyComplianceSummary,
  PolicyStatus,
  PolicyCategory,
  AcknowledgementStatus,
  ComplianceRisk,
  IPolicyQuizResult
} from '../models/IPolicy';

// ============================================================================
// EXPORT INTERFACES
// ============================================================================

export interface IPolicyReportOptions {
  includeArchived?: boolean;
  dateRangeStart?: Date;
  dateRangeEnd?: Date;
  departments?: string[];
  categories?: PolicyCategory[];
  statuses?: PolicyStatus[];
  complianceRisks?: ComplianceRisk[];
}

export interface IAcknowledgementReportOptions {
  includeCompleted?: boolean;
  includeExempted?: boolean;
  dateRangeStart?: Date;
  dateRangeEnd?: Date;
  policyIds?: number[];
  departments?: string[];
  statuses?: AcknowledgementStatus[];
}

export interface IComplianceReportOptions {
  groupBy: 'department' | 'policy' | 'location' | 'role';
  dateRangeStart?: Date;
  dateRangeEnd?: Date;
  includeDetails?: boolean;
}

export interface IExportResult {
  success: boolean;
  filename: string;
  recordCount: number;
  exportedAt: Date;
  errors?: string[];
}

// ============================================================================
// EXPORT DATA INTERFACES
// ============================================================================

interface IPolicyExportRow {
  'Policy Number': string;
  'Policy Name': string;
  'Category': string;
  'Type': string;
  'Status': string;
  'Version': string;
  'Effective Date': string;
  'Expiry Date': string;
  'Next Review Date': string;
  'Compliance Risk': string;
  'Is Mandatory': string;
  'Requires Acknowledgement': string;
  'Acknowledgement Type': string;
  'Read Timeframe': string;
  'Requires Quiz': string;
  'Quiz Passing Score': string;
  'Total Distributed': number;
  'Total Acknowledged': number;
  'Compliance %': string;
  'Average Rating': string;
  'Policy Owner': string;
  'Department Owner': string;
  'Published Date': string;
  'Tags': string;
}

interface IAcknowledgementExportRow {
  'Policy Number': string;
  'Policy Name': string;
  'Policy Category': string;
  'User Name': string;
  'User Email': string;
  'Department': string;
  'Role': string;
  'Location': string;
  'Status': string;
  'Assigned Date': string;
  'Due Date': string;
  'First Opened': string;
  'Acknowledged Date': string;
  'Days to Acknowledge': number | string;
  'Days Overdue': number | string;
  'Read Time (mins)': number | string;
  'Quiz Required': string;
  'Quiz Status': string;
  'Quiz Score': number | string;
  'Quiz Attempts': number | string;
  'Is Compliant': string;
  'Reminders Sent': number;
  'Is Exempted': string;
}

interface IComplianceExportRow {
  'Group': string;
  'Total Policies': number;
  'Total Users': number;
  'Total Acknowledged': number;
  'Total Overdue': number;
  'Total Exempted': number;
  'Compliance Rate %': string;
  'Average Time to Acknowledge (days)': string;
  'Critical Risk Count': number;
  'High Risk Count': number;
}

interface IQuizResultsExportRow {
  'Policy Number': string;
  'Policy Name': string;
  'User Name': string;
  'User Email': string;
  'Department': string;
  'Attempt Number': number;
  'Score': number;
  'Percentage': string;
  'Passed': string;
  'Time Spent (mins)': number;
  'Correct Answers': number;
  'Incorrect Answers': number;
  'Skipped Questions': number;
  'Started Date': string;
  'Completed Date': string;
}

interface IOverdueReportRow {
  'Policy Number': string;
  'Policy Name': string;
  'Policy Category': string;
  'Compliance Risk': string;
  'User Name': string;
  'User Email': string;
  'Department': string;
  'Manager': string;
  'Assigned Date': string;
  'Due Date': string;
  'Days Overdue': number;
  'Reminders Sent': number;
  'Last Reminder Date': string;
  'Manager Notified': string;
  'Escalation Level': number;
}

// ============================================================================
// POLICY REPORT EXPORT SERVICE
// ============================================================================

export class PolicyReportExportService {
  private sp: SPFI;
  private context: WebPartContext;

  // SharePoint List Names
  private readonly POLICY_LIST = PolicyLists.POLICIES;
  private readonly ACKNOWLEDGEMENT_LIST = PolicyLists.POLICY_ACKNOWLEDGEMENTS;
  private readonly QUIZ_RESULTS_LIST = QuizLists.POLICY_QUIZ_RESULTS;

  constructor(context: WebPartContext) {
    this.context = context;
    this.sp = getSP(context);
  }

  // ============================================================================
  // POLICY INVENTORY REPORT
  // ============================================================================

  /**
   * Export all policies to Excel/CSV
   */
  public async exportPolicyInventory(
    options: IPolicyReportOptions = {}
  ): Promise<IExportResult> {
    try {
      logger.info('PolicyReportExportService', 'Starting policy inventory export');

      // Build filter
      const filters: string[] = [];

      if (!options.includeArchived) {
        filters.push("Status ne 'Archived'");
        filters.push("Status ne 'Retired'");
      }

      if (options.categories && options.categories.length > 0) {
        const categoryFilters = options.categories.map(c => `PolicyCategory eq '${c}'`);
        filters.push(`(${categoryFilters.join(' or ')})`);
      }

      if (options.statuses && options.statuses.length > 0) {
        const statusFilters = options.statuses.map(s => `Status eq '${s}'`);
        filters.push(`(${statusFilters.join(' or ')})`);
      }

      // Fetch policies
      let query = this.sp.web.lists
        .getByTitle(this.POLICY_LIST)
        .items
        .select(
          'Id', 'PolicyNumber', 'PolicyName', 'PolicyCategory', 'PolicyType',
          'Status', 'VersionNumber', 'EffectiveDate', 'ExpiryDate', 'NextReviewDate',
          'ComplianceRisk', 'IsMandatory', 'RequiresAcknowledgement', 'AcknowledgementType',
          'ReadTimeframe', 'RequiresQuiz', 'QuizPassingScore', 'TotalDistributed',
          'TotalAcknowledged', 'CompliancePercentage', 'AverageRating', 'DepartmentOwner',
          'PublishedDate', 'Tags', 'PolicyOwner/Title', 'PolicyOwner/EMail'
        )
        .expand('PolicyOwner')
        .top(500);

      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      const policies = await query();

      // Transform to export format
      const exportData: IPolicyExportRow[] = policies.map((policy: any) => ({
        'Policy Number': policy.PolicyNumber || '',
        'Policy Name': policy.PolicyName || '',
        'Category': policy.PolicyCategory || '',
        'Type': policy.PolicyType || '',
        'Status': policy.Status || '',
        'Version': policy.VersionNumber || '',
        'Effective Date': this.formatDate(policy.EffectiveDate),
        'Expiry Date': this.formatDate(policy.ExpiryDate),
        'Next Review Date': this.formatDate(policy.NextReviewDate),
        'Compliance Risk': policy.ComplianceRisk || '',
        'Is Mandatory': policy.IsMandatory ? 'Yes' : 'No',
        'Requires Acknowledgement': policy.RequiresAcknowledgement ? 'Yes' : 'No',
        'Acknowledgement Type': policy.AcknowledgementType || 'N/A',
        'Read Timeframe': policy.ReadTimeframe || 'N/A',
        'Requires Quiz': policy.RequiresQuiz ? 'Yes' : 'No',
        'Quiz Passing Score': policy.QuizPassingScore ? `${policy.QuizPassingScore}%` : 'N/A',
        'Total Distributed': policy.TotalDistributed || 0,
        'Total Acknowledged': policy.TotalAcknowledged || 0,
        'Compliance %': policy.CompliancePercentage ? `${policy.CompliancePercentage.toFixed(1)}%` : '0%',
        'Average Rating': policy.AverageRating ? policy.AverageRating.toFixed(1) : 'N/A',
        'Policy Owner': policy.PolicyOwner?.Title || '',
        'Department Owner': policy.DepartmentOwner || '',
        'Published Date': this.formatDate(policy.PublishedDate),
        'Tags': Array.isArray(policy.Tags) ? policy.Tags.join(', ') : (policy.Tags || '')
      }));

      // Generate CSV and download
      const filename = `Policy_Inventory_${this.getDateStamp()}.csv`;
      this.downloadCSV(exportData, filename);

      logger.info('PolicyReportExportService', `Successfully exported ${exportData.length} policies`);

      return {
        success: true,
        filename,
        recordCount: exportData.length,
        exportedAt: new Date()
      };
    } catch (error) {
      logger.error('PolicyReportExportService', 'Failed to export policy inventory:', error);
      return {
        success: false,
        filename: '',
        recordCount: 0,
        exportedAt: new Date(),
        errors: [error instanceof Error ? error.message : 'Unknown error']
      };
    }
  }

  // ============================================================================
  // ACKNOWLEDGEMENT STATUS REPORT
  // ============================================================================

  /**
   * Export acknowledgement status to Excel/CSV
   */
  public async exportAcknowledgementStatus(
    options: IAcknowledgementReportOptions = {}
  ): Promise<IExportResult> {
    try {
      logger.info('PolicyReportExportService', 'Starting acknowledgement status export');

      // Build filter
      const filters: string[] = [];

      if (!options.includeCompleted) {
        filters.push("Status ne 'Acknowledged'");
      }

      if (!options.includeExempted) {
        filters.push("IsExempted eq false");
      }

      if (options.statuses && options.statuses.length > 0) {
        const statusFilters = options.statuses.map(s => `Status eq '${s}'`);
        filters.push(`(${statusFilters.join(' or ')})`);
      }

      if (options.policyIds && options.policyIds.length > 0) {
        const policyFilters = options.policyIds.map(id => `PolicyId eq ${id}`);
        filters.push(`(${policyFilters.join(' or ')})`);
      }

      if (options.dateRangeStart) {
        filters.push(`AssignedDate ge datetime'${options.dateRangeStart.toISOString()}'`);
      }

      if (options.dateRangeEnd) {
        filters.push(`AssignedDate le datetime'${options.dateRangeEnd.toISOString()}'`);
      }

      // Fetch acknowledgements
      let query = this.sp.web.lists
        .getByTitle(this.ACKNOWLEDGEMENT_LIST)
        .items
        .select(
          'Id', 'PolicyId', 'PolicyNumber', 'PolicyName', 'PolicyCategory',
          'UserId', 'UserEmail', 'UserDepartment', 'UserRole', 'UserLocation',
          'Status', 'AssignedDate', 'DueDate', 'FirstOpenedDate', 'AcknowledgedDate',
          'TotalReadTimeSeconds', 'QuizRequired', 'QuizStatus', 'QuizScore', 'QuizAttempts',
          'IsCompliant', 'RemindersSent', 'IsExempted', 'OverdueDays',
          'User/Title', 'User/EMail'
        )
        .expand('User')
        .top(500);

      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      const acknowledgements = await query();

      // Transform to export format
      const exportData: IAcknowledgementExportRow[] = acknowledgements.map((ack: any) => {
        const assignedDate = ack.AssignedDate ? new Date(ack.AssignedDate) : null;
        const acknowledgedDate = ack.AcknowledgedDate ? new Date(ack.AcknowledgedDate) : null;
        const dueDate = ack.DueDate ? new Date(ack.DueDate) : null;

        let daysToAcknowledge: number | string = 'N/A';
        if (assignedDate && acknowledgedDate) {
          daysToAcknowledge = Math.ceil((acknowledgedDate.getTime() - assignedDate.getTime()) / (1000 * 60 * 60 * 24));
        }

        let daysOverdue: number | string = 'N/A';
        if (dueDate && !acknowledgedDate) {
          const now = new Date();
          if (now > dueDate) {
            daysOverdue = Math.ceil((now.getTime() - dueDate.getTime()) / (1000 * 60 * 60 * 24));
          }
        } else if (ack.OverdueDays) {
          daysOverdue = ack.OverdueDays;
        }

        return {
          'Policy Number': ack.PolicyNumber || '',
          'Policy Name': ack.PolicyName || '',
          'Policy Category': ack.PolicyCategory || '',
          'User Name': ack.User?.Title || '',
          'User Email': ack.UserEmail || ack.User?.EMail || '',
          'Department': ack.UserDepartment || '',
          'Role': ack.UserRole || '',
          'Location': ack.UserLocation || '',
          'Status': ack.Status || '',
          'Assigned Date': this.formatDate(ack.AssignedDate),
          'Due Date': this.formatDate(ack.DueDate),
          'First Opened': this.formatDate(ack.FirstOpenedDate),
          'Acknowledged Date': this.formatDate(ack.AcknowledgedDate),
          'Days to Acknowledge': daysToAcknowledge,
          'Days Overdue': daysOverdue,
          'Read Time (mins)': ack.TotalReadTimeSeconds ? Math.round(ack.TotalReadTimeSeconds / 60) : 'N/A',
          'Quiz Required': ack.QuizRequired ? 'Yes' : 'No',
          'Quiz Status': ack.QuizStatus || 'N/A',
          'Quiz Score': ack.QuizScore ?? 'N/A',
          'Quiz Attempts': ack.QuizAttempts ?? 'N/A',
          'Is Compliant': ack.IsCompliant ? 'Yes' : 'No',
          'Reminders Sent': ack.RemindersSent || 0,
          'Is Exempted': ack.IsExempted ? 'Yes' : 'No'
        };
      });

      // Generate CSV and download
      const filename = `Acknowledgement_Status_${this.getDateStamp()}.csv`;
      this.downloadCSV(exportData, filename);

      logger.info('PolicyReportExportService', `Successfully exported ${exportData.length} acknowledgements`);

      return {
        success: true,
        filename,
        recordCount: exportData.length,
        exportedAt: new Date()
      };
    } catch (error) {
      logger.error('PolicyReportExportService', 'Failed to export acknowledgement status:', error);
      return {
        success: false,
        filename: '',
        recordCount: 0,
        exportedAt: new Date(),
        errors: [error instanceof Error ? error.message : 'Unknown error']
      };
    }
  }

  // ============================================================================
  // COMPLIANCE SUMMARY REPORT
  // ============================================================================

  /**
   * Export compliance summary grouped by department, policy, location, or role
   */
  public async exportComplianceSummary(
    options: IComplianceReportOptions
  ): Promise<IExportResult> {
    try {
      logger.info('PolicyReportExportService', `Starting compliance summary export (grouped by ${options.groupBy})`);

      // Fetch all acknowledgements for aggregation
      const acknowledgements = await this.sp.web.lists
        .getByTitle(this.ACKNOWLEDGEMENT_LIST)
        .items
        .select(
          'Id', 'PolicyId', 'PolicyNumber', 'PolicyName', 'PolicyCategory',
          'UserDepartment', 'UserRole', 'UserLocation', 'Status',
          'AssignedDate', 'AcknowledgedDate', 'IsCompliant', 'IsExempted',
          'OverdueDays'
        )
        .top(500)();

      // Fetch policies for risk levels
      const policies = await this.sp.web.lists
        .getByTitle(this.POLICY_LIST)
        .items
        .select('Id', 'PolicyNumber', 'PolicyName', 'ComplianceRisk')
        .top(500)();

      const policyRiskMap = new Map<number, string>();
      policies.forEach((p: any) => policyRiskMap.set(p.Id, p.ComplianceRisk));

      // Group data based on option
      const groupedData = new Map<string, {
        totalPolicies: Set<number>;
        totalUsers: Set<string>;
        acknowledged: number;
        overdue: number;
        exempted: number;
        totalDaysToAcknowledge: number;
        acknowledgedCount: number;
        criticalRisk: number;
        highRisk: number;
      }>();

      acknowledgements.forEach((ack: any) => {
        let groupKey: string;
        switch (options.groupBy) {
          case 'department':
            groupKey = ack.UserDepartment || 'Unknown';
            break;
          case 'policy':
            groupKey = `${ack.PolicyNumber} - ${ack.PolicyName}`;
            break;
          case 'location':
            groupKey = ack.UserLocation || 'Unknown';
            break;
          case 'role':
            groupKey = ack.UserRole || 'Unknown';
            break;
          default:
            groupKey = 'All';
        }

        if (!groupedData.has(groupKey)) {
          groupedData.set(groupKey, {
            totalPolicies: new Set(),
            totalUsers: new Set(),
            acknowledged: 0,
            overdue: 0,
            exempted: 0,
            totalDaysToAcknowledge: 0,
            acknowledgedCount: 0,
            criticalRisk: 0,
            highRisk: 0
          });
        }

        const group = groupedData.get(groupKey)!;
        group.totalPolicies.add(ack.PolicyId);
        group.totalUsers.add(ack.UserEmail || ack.AckUserId?.toString());

        if (ack.AckStatus === 'Acknowledged') {
          group.acknowledged++;
          if (ack.AssignedDate && ack.AcknowledgedDate) {
            const assigned = new Date(ack.AssignedDate);
            const acknowledged = new Date(ack.AcknowledgedDate);
            const days = Math.ceil((acknowledged.getTime() - assigned.getTime()) / (1000 * 60 * 60 * 24));
            group.totalDaysToAcknowledge += days;
            group.acknowledgedCount++;
          }
        } else if (ack.AckStatus === 'Overdue' || (ack.OverdueDays && ack.OverdueDays > 0)) {
          group.overdue++;
        }

        if (ack.IsExempted) {
          group.exempted++;
        }

        const risk = policyRiskMap.get(ack.PolicyId);
        if (risk === 'Critical') {
          group.criticalRisk++;
        } else if (risk === 'High') {
          group.highRisk++;
        }
      });

      // Transform to export format
      const exportData: IComplianceExportRow[] = [];
      groupedData.forEach((data, group) => {
        const totalRecords = data.acknowledged + data.overdue + (data.exempted ? 0 : 0);
        const complianceRate = totalRecords > 0
          ? ((data.acknowledged / (totalRecords - data.exempted)) * 100)
          : 0;
        const avgDays = data.acknowledgedCount > 0
          ? data.totalDaysToAcknowledge / data.acknowledgedCount
          : 0;

        exportData.push({
          'Group': group,
          'Total Policies': data.totalPolicies.size,
          'Total Users': data.totalUsers.size,
          'Total Acknowledged': data.acknowledged,
          'Total Overdue': data.overdue,
          'Total Exempted': data.exempted,
          'Compliance Rate %': `${complianceRate.toFixed(1)}%`,
          'Average Time to Acknowledge (days)': avgDays.toFixed(1),
          'Critical Risk Count': data.criticalRisk,
          'High Risk Count': data.highRisk
        });
      });

      // Sort by group name
      exportData.sort((a, b) => a.Group.localeCompare(b.Group));

      // Generate CSV and download
      const filename = `Compliance_Summary_By_${options.groupBy}_${this.getDateStamp()}.csv`;
      this.downloadCSV(exportData, filename);

      logger.info('PolicyReportExportService', `Successfully exported ${exportData.length} compliance groups`);

      return {
        success: true,
        filename,
        recordCount: exportData.length,
        exportedAt: new Date()
      };
    } catch (error) {
      logger.error('PolicyReportExportService', 'Failed to export compliance summary:', error);
      return {
        success: false,
        filename: '',
        recordCount: 0,
        exportedAt: new Date(),
        errors: [error instanceof Error ? error.message : 'Unknown error']
      };
    }
  }

  // ============================================================================
  // OVERDUE ACKNOWLEDGEMENTS REPORT
  // ============================================================================

  /**
   * Export overdue acknowledgements report
   */
  public async exportOverdueReport(): Promise<IExportResult> {
    try {
      logger.info('PolicyReportExportService', 'Starting overdue acknowledgements export');

      // Fetch overdue acknowledgements
      const acknowledgements = await this.sp.web.lists
        .getByTitle(this.ACKNOWLEDGEMENT_LIST)
        .items
        .select(
          'Id', 'PolicyId', 'PolicyNumber', 'PolicyName', 'PolicyCategory',
          'UserEmail', 'UserDepartment', 'AssignedDate', 'DueDate',
          'RemindersSent', 'LastReminderDate', 'ManagerNotified', 'EscalationLevel',
          'OverdueDays', 'User/Title', 'User/EMail'
        )
        .expand('User')
        .filter("AckStatus eq 'Overdue' or AckStatus eq 'Sent' or AckStatus eq 'InProgress'")
        .top(500)();

      // Filter to only include truly overdue items
      const now = new Date();
      const overdueAcks = acknowledgements.filter((ack: any) => {
        if (ack.DueDate) {
          const dueDate = new Date(ack.DueDate);
          return dueDate < now;
        }
        return ack.AckStatus === 'Overdue';
      });

      // Fetch policies for compliance risk
      const policyIds = Array.from(new Set(overdueAcks.map((a: any) => a.PolicyId)));
      const policies = await this.sp.web.lists
        .getByTitle(this.POLICY_LIST)
        .items
        .select('Id', 'ComplianceRisk')
        .filter(policyIds.map(id => `Id eq ${id}`).join(' or '))
        .top(500)();

      const policyRiskMap = new Map<number, string>();
      policies.forEach((p: any) => policyRiskMap.set(p.Id, p.ComplianceRisk));

      // Transform to export format
      const exportData: IOverdueReportRow[] = overdueAcks.map((ack: any) => {
        const dueDate = ack.DueDate ? new Date(ack.DueDate) : null;
        const daysOverdue = dueDate
          ? Math.ceil((now.getTime() - dueDate.getTime()) / (1000 * 60 * 60 * 24))
          : (ack.OverdueDays || 0);

        return {
          'Policy Number': ack.PolicyNumber || '',
          'Policy Name': ack.PolicyName || '',
          'Policy Category': ack.PolicyCategory || '',
          'Compliance Risk': policyRiskMap.get(ack.PolicyId) || 'Unknown',
          'User Name': ack.User?.Title || '',
          'User Email': ack.UserEmail || ack.User?.EMail || '',
          'Department': ack.UserDepartment || '',
          'Manager': '', // Would need to fetch from user profile
          'Assigned Date': this.formatDate(ack.AssignedDate),
          'Due Date': this.formatDate(ack.DueDate),
          'Days Overdue': daysOverdue,
          'Reminders Sent': ack.RemindersSent || 0,
          'Last Reminder Date': this.formatDate(ack.LastReminderDate),
          'Manager Notified': ack.ManagerNotified ? 'Yes' : 'No',
          'Escalation Level': ack.EscalationLevel || 0
        };
      });

      // Sort by days overdue (descending) and compliance risk
      const riskOrder: Record<string, number> = { 'Critical': 0, 'High': 1, 'Medium': 2, 'Low': 3, 'Informational': 4 };
      exportData.sort((a, b) => {
        const riskDiff = (riskOrder[a['Compliance Risk']] || 5) - (riskOrder[b['Compliance Risk']] || 5);
        if (riskDiff !== 0) return riskDiff;
        return b['Days Overdue'] - a['Days Overdue'];
      });

      // Generate CSV and download
      const filename = `Overdue_Acknowledgements_${this.getDateStamp()}.csv`;
      this.downloadCSV(exportData, filename);

      logger.info('PolicyReportExportService', `Successfully exported ${exportData.length} overdue acknowledgements`);

      return {
        success: true,
        filename,
        recordCount: exportData.length,
        exportedAt: new Date()
      };
    } catch (error) {
      logger.error('PolicyReportExportService', 'Failed to export overdue report:', error);
      return {
        success: false,
        filename: '',
        recordCount: 0,
        exportedAt: new Date(),
        errors: [error instanceof Error ? error.message : 'Unknown error']
      };
    }
  }

  // ============================================================================
  // QUIZ RESULTS REPORT
  // ============================================================================

  /**
   * Export quiz results report
   */
  public async exportQuizResults(
    policyIds?: number[],
    dateRangeStart?: Date,
    dateRangeEnd?: Date
  ): Promise<IExportResult> {
    try {
      logger.info('PolicyReportExportService', 'Starting quiz results export');

      // Build filter
      const filters: string[] = [];

      if (policyIds && policyIds.length > 0) {
        // Need to join through Quiz to get PolicyId
        // For simplicity, fetch all and filter client-side
      }

      if (dateRangeStart) {
        filters.push(`CompletedDate ge datetime'${dateRangeStart.toISOString()}'`);
      }

      if (dateRangeEnd) {
        filters.push(`CompletedDate le datetime'${dateRangeEnd.toISOString()}'`);
      }

      // Fetch quiz results
      let query = this.sp.web.lists
        .getByTitle(this.QUIZ_RESULTS_LIST)
        .items
        .select(
          'Id', 'QuizId', 'UserId', 'AttemptNumber', 'Score', 'Percentage',
          'Passed', 'StartedDate', 'CompletedDate', 'TimeSpentSeconds',
          'CorrectAnswers', 'IncorrectAnswers', 'SkippedQuestions',
          'User/Title', 'User/EMail', 'User/Department'
        )
        .expand('User')
        .top(500);

      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      const results = await query();

      // Fetch quiz details to get policy info
      const quizIds = Array.from(new Set(results.map((r: any) => r.QuizId)));
      let quizPolicyMap = new Map<number, { policyNumber: string; policyName: string }>();

      if (quizIds.length > 0) {
        const quizzes = await this.sp.web.lists
          .getByTitle(QuizLists.POLICY_QUIZZES)
          .items
          .select('Id', 'PolicyId', 'Policy/PolicyNumber', 'Policy/PolicyName')
          .expand('Policy')
          .filter(quizIds.map(id => `Id eq ${id}`).join(' or '))
          .top(500)();

        quizzes.forEach((q: any) => {
          quizPolicyMap.set(q.Id, {
            policyNumber: q.Policy?.PolicyNumber || '',
            policyName: q.Policy?.PolicyName || ''
          });
        });
      }

      // Transform to export format
      const exportData: IQuizResultsExportRow[] = results.map((result: any) => {
        const policyInfo = quizPolicyMap.get(result.QuizId) || { policyNumber: '', policyName: '' };

        return {
          'Policy Number': policyInfo.policyNumber,
          'Policy Name': policyInfo.policyName,
          'User Name': result.User?.Title || '',
          'User Email': result.User?.EMail || '',
          'Department': result.User?.Department || '',
          'Attempt Number': result.AttemptNumber || 1,
          'Score': result.Score || 0,
          'Percentage': result.Percentage ? `${result.Percentage.toFixed(1)}%` : '0%',
          'Passed': result.Passed ? 'Yes' : 'No',
          'Time Spent (mins)': result.TimeSpentSeconds ? Math.round(result.TimeSpentSeconds / 60) : 0,
          'Correct Answers': result.CorrectAnswers || 0,
          'Incorrect Answers': result.IncorrectAnswers || 0,
          'Skipped Questions': result.SkippedQuestions || 0,
          'Started Date': this.formatDateTime(result.StartedDate),
          'Completed Date': this.formatDateTime(result.CompletedDate)
        };
      });

      // Filter by policy IDs if specified
      let filteredData = exportData;
      if (policyIds && policyIds.length > 0) {
        // This would need the policy ID from quizPolicyMap - for now export all
      }

      // Sort by completed date descending
      filteredData.sort((a, b) => {
        const dateA = a['Completed Date'] ? new Date(a['Completed Date']).getTime() : 0;
        const dateB = b['Completed Date'] ? new Date(b['Completed Date']).getTime() : 0;
        return dateB - dateA;
      });

      // Generate CSV and download
      const filename = `Quiz_Results_${this.getDateStamp()}.csv`;
      this.downloadCSV(filteredData, filename);

      logger.info('PolicyReportExportService', `Successfully exported ${filteredData.length} quiz results`);

      return {
        success: true,
        filename,
        recordCount: filteredData.length,
        exportedAt: new Date()
      };
    } catch (error) {
      logger.error('PolicyReportExportService', 'Failed to export quiz results:', error);
      return {
        success: false,
        filename: '',
        recordCount: 0,
        exportedAt: new Date(),
        errors: [error instanceof Error ? error.message : 'Unknown error']
      };
    }
  }

  // ============================================================================
  // EXECUTIVE SUMMARY REPORT
  // ============================================================================

  /**
   * Export executive summary with key metrics
   */
  public async exportExecutiveSummary(): Promise<IExportResult> {
    try {
      logger.info('PolicyReportExportService', 'Starting executive summary export');

      // Fetch all policies
      const policies = await this.sp.web.lists
        .getByTitle(this.POLICY_LIST)
        .items
        .select(
          'Id', 'Status', 'PolicyCategory', 'ComplianceRisk', 'IsMandatory',
          'TotalDistributed', 'TotalAcknowledged', 'CompliancePercentage',
          'ExpiryDate', 'NextReviewDate'
        )
        .top(500)();

      // Fetch all acknowledgements
      const acknowledgements = await this.sp.web.lists
        .getByTitle(this.ACKNOWLEDGEMENT_LIST)
        .items
        .select('Id', 'Status', 'IsCompliant', 'OverdueDays', 'UserDepartment')
        .top(500)();

      // Calculate metrics
      const now = new Date();
      const thirtyDaysFromNow = new Date();
      thirtyDaysFromNow.setDate(thirtyDaysFromNow.getDate() + 30);

      const totalPolicies = policies.length;
      const activePolicies = policies.filter((p: any) => p.Status === 'Published').length;
      const draftPolicies = policies.filter((p: any) => p.Status === 'Draft').length;
      const expiringSoon = policies.filter((p: any) => {
        if (!p.ExpiryDate) return false;
        const expiry = new Date(p.ExpiryDate);
        return expiry <= thirtyDaysFromNow && expiry > now;
      }).length;
      const reviewDue = policies.filter((p: any) => {
        if (!p.NextReviewDate) return false;
        const review = new Date(p.NextReviewDate);
        return review <= thirtyDaysFromNow;
      }).length;

      const totalAcknowledgements = acknowledgements.length;
      const acknowledgedCount = acknowledgements.filter((a: any) => a.AckStatus === 'Acknowledged').length;
      const overdueCount = acknowledgements.filter((a: any) => a.AckStatus === 'Overdue' || a.OverdueDays > 0).length;
      const overallCompliance = totalAcknowledgements > 0
        ? ((acknowledgedCount / totalAcknowledgements) * 100)
        : 0;

      // By category
      const byCategory = new Map<string, { total: number; acknowledged: number }>();
      policies.forEach((p: any) => {
        const cat = p.PolicyCategory || 'Uncategorized';
        if (!byCategory.has(cat)) {
          byCategory.set(cat, { total: 0, acknowledged: 0 });
        }
        byCategory.get(cat)!.total++;
      });

      // By risk level
      const byRisk = new Map<string, number>();
      policies.forEach((p: any) => {
        const risk = p.ComplianceRisk || 'Unknown';
        byRisk.set(risk, (byRisk.get(risk) || 0) + 1);
      });

      // Build summary export
      const summaryData = [
        { 'Metric': 'Report Generated', 'Value': this.formatDateTime(new Date()) },
        { 'Metric': '', 'Value': '' },
        { 'Metric': '=== POLICY OVERVIEW ===', 'Value': '' },
        { 'Metric': 'Total Policies', 'Value': totalPolicies.toString() },
        { 'Metric': 'Active (Published)', 'Value': activePolicies.toString() },
        { 'Metric': 'Draft', 'Value': draftPolicies.toString() },
        { 'Metric': 'Expiring in 30 Days', 'Value': expiringSoon.toString() },
        { 'Metric': 'Review Due in 30 Days', 'Value': reviewDue.toString() },
        { 'Metric': '', 'Value': '' },
        { 'Metric': '=== COMPLIANCE OVERVIEW ===', 'Value': '' },
        { 'Metric': 'Total Acknowledgements', 'Value': totalAcknowledgements.toString() },
        { 'Metric': 'Acknowledged', 'Value': acknowledgedCount.toString() },
        { 'Metric': 'Overdue', 'Value': overdueCount.toString() },
        { 'Metric': 'Overall Compliance Rate', 'Value': `${overallCompliance.toFixed(1)}%` },
        { 'Metric': '', 'Value': '' },
        { 'Metric': '=== BY CATEGORY ===', 'Value': '' }
      ];

      byCategory.forEach((data, category) => {
        summaryData.push({ 'Metric': category, 'Value': data.total.toString() });
      });

      summaryData.push({ 'Metric': '', 'Value': '' });
      summaryData.push({ 'Metric': '=== BY COMPLIANCE RISK ===', 'Value': '' });

      ['Critical', 'High', 'Medium', 'Low', 'Informational'].forEach(risk => {
        const count = byRisk.get(risk) || 0;
        summaryData.push({ 'Metric': risk, 'Value': count.toString() });
      });

      // Generate CSV and download
      const filename = `Policy_Executive_Summary_${this.getDateStamp()}.csv`;
      this.downloadCSV(summaryData, filename);

      logger.info('PolicyReportExportService', 'Successfully exported executive summary');

      return {
        success: true,
        filename,
        recordCount: summaryData.length,
        exportedAt: new Date()
      };
    } catch (error) {
      logger.error('PolicyReportExportService', 'Failed to export executive summary:', error);
      return {
        success: false,
        filename: '',
        recordCount: 0,
        exportedAt: new Date(),
        errors: [error instanceof Error ? error.message : 'Unknown error']
      };
    }
  }

  // ============================================================================
  // AUDIT TRAIL REPORT
  // ============================================================================

  /**
   * Export audit trail from PM_PolicyAuditLog
   */
  public async exportAuditTrail(
    options: { dateRangeStart?: Date; dateRangeEnd?: Date; departments?: string[] } = {}
  ): Promise<IExportResult> {
    try {
      logger.info('PolicyReportExportService', 'Starting audit trail export');

      const filters: string[] = [];
      if (options.dateRangeStart) {
        filters.push(`ActionDate ge datetime'${options.dateRangeStart.toISOString()}'`);
      }
      if (options.dateRangeEnd) {
        filters.push(`ActionDate le datetime'${options.dateRangeEnd.toISOString()}'`);
      }

      let query = this.sp.web.lists
        .getByTitle(PolicyLists.POLICY_AUDIT_LOG)
        .items
        .select(
          'Id', 'AuditAction', 'ActionDescription', 'ActionDate',
          'PolicyId', 'PolicyTitle', 'PerformedByName', 'PerformedByEmail',
          'EntityType', 'ComplianceRelevant', 'IPAddress'
        )
        .orderBy('ActionDate', false)
        .top(500);

      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      const auditItems = await query();

      const exportData = auditItems.map((item: any) => ({
        'Date': this.formatDateTime(item.ActionDate),
        'Action': item.AuditAction || '',
        'Description': item.ActionDescription || '',
        'Policy': item.PolicyTitle || '',
        'Policy ID': item.PolicyId || '',
        'Performed By': item.PerformedByName || '',
        'Email': item.PerformedByEmail || '',
        'Entity Type': item.EntityType || '',
        'Compliance Relevant': item.ComplianceRelevant ? 'Yes' : 'No'
      }));

      const filename = `Audit_Trail_${this.getDateStamp()}.csv`;
      this.downloadCSV(exportData, filename);

      logger.info('PolicyReportExportService', `Successfully exported ${exportData.length} audit entries`);

      return {
        success: true,
        filename,
        recordCount: exportData.length,
        exportedAt: new Date()
      };
    } catch (error) {
      logger.error('PolicyReportExportService', 'Failed to export audit trail:', error);
      return {
        success: false,
        filename: '',
        recordCount: 0,
        exportedAt: new Date(),
        errors: [error instanceof Error ? error.message : 'Unknown error']
      };
    }
  }

  // ============================================================================
  // DELEGATION SUMMARY REPORT
  // ============================================================================

  /**
   * Export delegation summary from PM_ApprovalDelegations
   */
  public async exportDelegationSummary(
    options: { dateRangeStart?: Date; dateRangeEnd?: Date } = {}
  ): Promise<IExportResult> {
    try {
      logger.info('PolicyReportExportService', 'Starting delegation summary export');

      const filters: string[] = [];
      if (options.dateRangeStart) {
        filters.push(`Created ge datetime'${options.dateRangeStart.toISOString()}'`);
      }
      if (options.dateRangeEnd) {
        filters.push(`Created le datetime'${options.dateRangeEnd.toISOString()}'`);
      }

      let query = this.sp.web.lists
        .getByTitle(ApprovalLists.APPROVAL_DELEGATIONS)
        .items
        .select(
          'Id', 'Title', 'DelegatedById', 'DelegatedByName', 'DelegatedToId', 'DelegatedToName',
          'DelegationType', 'Reason', 'StartDate', 'EndDate', 'Status',
          'Created', 'Modified'
        )
        .orderBy('Created', false)
        .top(500);

      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      const delegations = await query();

      const exportData = delegations.map((item: any) => ({
        'Delegation': item.Title || '',
        'Delegated By': item.DelegatedByName || '',
        'Delegated To': item.DelegatedToName || '',
        'Type': item.DelegationType || '',
        'Reason': item.Reason || '',
        'Start Date': this.formatDate(item.StartDate),
        'End Date': this.formatDate(item.EndDate),
        'Status': item.Status || '',
        'Created': this.formatDateTime(item.Created)
      }));

      const filename = `Delegation_Summary_${this.getDateStamp()}.csv`;
      this.downloadCSV(exportData, filename);

      logger.info('PolicyReportExportService', `Successfully exported ${exportData.length} delegations`);

      return {
        success: true,
        filename,
        recordCount: exportData.length,
        exportedAt: new Date()
      };
    } catch (error) {
      logger.error('PolicyReportExportService', 'Failed to export delegation summary:', error);
      return {
        success: false,
        filename: '',
        recordCount: 0,
        exportedAt: new Date(),
        errors: [error instanceof Error ? error.message : 'Unknown error']
      };
    }
  }

  // ============================================================================
  // REVIEW SCHEDULE REPORT
  // ============================================================================

  /**
   * Export policy review schedule — upcoming, due, and overdue reviews
   */
  public async exportReviewSchedule(
    options: { dateRangeStart?: Date; dateRangeEnd?: Date; departments?: string[] } = {}
  ): Promise<IExportResult> {
    try {
      logger.info('PolicyReportExportService', 'Starting review schedule export');

      const filters: string[] = [
        "Status eq 'Published'",
        "NextReviewDate ne null"
      ];

      if (options.departments && options.departments.length > 0) {
        const deptFilters = options.departments.map(d => `DepartmentOwner eq '${d}'`);
        filters.push(`(${deptFilters.join(' or ')})`);
      }

      const policies = await this.sp.web.lists
        .getByTitle(this.POLICY_LIST)
        .items
        .select(
          'Id', 'PolicyNumber', 'PolicyName', 'PolicyCategory', 'ComplianceRisk',
          'NextReviewDate', 'ReviewFrequency', 'DepartmentOwner', 'LastReviewDate',
          'PolicyOwner/Title', 'PolicyOwner/EMail'
        )
        .expand('PolicyOwner')
        .filter(filters.join(' and '))
        .orderBy('NextReviewDate', true)
        .top(500)();

      const now = new Date();
      const exportData = policies.map((p: any) => {
        const nextReview = p.NextReviewDate ? new Date(p.NextReviewDate) : null;
        const daysUntil = nextReview ? Math.ceil((nextReview.getTime() - now.getTime()) / 86400000) : 0;
        const urgency = daysUntil < 0 ? 'Overdue' : daysUntil <= 14 ? 'Due Soon' : daysUntil <= 30 ? 'Upcoming' : 'On Track';

        return {
          'Policy Number': p.PolicyNumber || '',
          'Policy Name': p.PolicyName || '',
          'Category': p.PolicyCategory || '',
          'Risk Level': p.ComplianceRisk || '',
          'Department': p.DepartmentOwner || '',
          'Owner': p.PolicyOwner?.Title || '',
          'Review Frequency': p.ReviewFrequency || 'Annual',
          'Last Review': this.formatDate(p.LastReviewDate),
          'Next Review': this.formatDate(p.NextReviewDate),
          'Days Until Due': daysUntil,
          'Urgency': urgency
        };
      });

      // Sort: overdue first, then by days until
      exportData.sort((a: any, b: any) => a['Days Until Due'] - b['Days Until Due']);

      const filename = `Review_Schedule_${this.getDateStamp()}.csv`;
      this.downloadCSV(exportData, filename);

      logger.info('PolicyReportExportService', `Successfully exported ${exportData.length} review schedules`);

      return {
        success: true,
        filename,
        recordCount: exportData.length,
        exportedAt: new Date()
      };
    } catch (error) {
      logger.error('PolicyReportExportService', 'Failed to export review schedule:', error);
      return {
        success: false,
        filename: '',
        recordCount: 0,
        exportedAt: new Date(),
        errors: [error instanceof Error ? error.message : 'Unknown error']
      };
    }
  }

  // ============================================================================
  // PREVIEW DATA (limited query for Report Builder preview)
  // ============================================================================

  /**
   * Fetch a limited preview of compliance data by department (top 5)
   */
  public async getCompliancePreview(
    options: { dateRangeStart?: Date; dateRangeEnd?: Date; departments?: string[] } = {}
  ): Promise<{ departments: any[]; totals: { assigned: number; acknowledged: number; pending: number; overdue: number; rate: number } }> {
    try {
      const filters: string[] = [];
      if (options.dateRangeStart) {
        filters.push(`AssignedDate ge datetime'${options.dateRangeStart.toISOString()}'`);
      }
      if (options.dateRangeEnd) {
        filters.push(`AssignedDate le datetime'${options.dateRangeEnd.toISOString()}'`);
      }

      let query = this.sp.web.lists
        .getByTitle(this.ACKNOWLEDGEMENT_LIST)
        .items
        .select('Id', 'AckStatus', 'UserDepartment', 'DueDate')
        .top(500);

      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      const acks = await query();

      // Group by department
      const deptMap = new Map<string, { assigned: number; acknowledged: number; pending: number; overdue: number }>();
      const now = new Date();
      acks.forEach((ack: any) => {
        const dept = ack.UserDepartment || 'Unknown';
        if (options.departments && options.departments.length > 0 && !options.departments.includes(dept)) return;
        if (!deptMap.has(dept)) deptMap.set(dept, { assigned: 0, acknowledged: 0, pending: 0, overdue: 0 });
        const d = deptMap.get(dept)!;
        d.assigned++;
        const isAcked = ack.AckStatus === 'Acknowledged' || ack.AckStatus === 'completed';
        if (isAcked) { d.acknowledged++; }
        else if (ack.DueDate && new Date(ack.DueDate) < now) { d.overdue++; }
        else { d.pending++; }
      });

      const departments = Array.from(deptMap.entries())
        .map(([dept, data]) => ({
          department: dept,
          ...data,
          rate: data.assigned > 0 ? Math.round((data.acknowledged / data.assigned) * 100) : 0
        }))
        .sort((a, b) => b.assigned - a.assigned)
        .slice(0, 8);

      const totals = departments.reduce((acc, d) => ({
        assigned: acc.assigned + d.assigned,
        acknowledged: acc.acknowledged + d.acknowledged,
        pending: acc.pending + d.pending,
        overdue: acc.overdue + d.overdue,
        rate: 0
      }), { assigned: 0, acknowledged: 0, pending: 0, overdue: 0, rate: 0 });
      totals.rate = totals.assigned > 0 ? Math.round((totals.acknowledged / totals.assigned) * 100) : 0;

      return { departments, totals };
    } catch (error) {
      logger.error('PolicyReportExportService', 'Failed to fetch compliance preview:', error);
      return { departments: [], totals: { assigned: 0, acknowledged: 0, pending: 0, overdue: 0, rate: 0 } };
    }
  }

  /**
   * Fetch distinct departments from PM_Policies for the builder dropdown
   */
  public async getDistinctDepartments(): Promise<string[]> {
    try {
      const policies = await this.sp.web.lists
        .getByTitle(this.POLICY_LIST)
        .items
        .select('DepartmentOwner')
        .filter("DepartmentOwner ne null")
        .top(500)();

      const depts = new Set<string>();
      policies.forEach((p: any) => { if (p.DepartmentOwner) depts.add(p.DepartmentOwner); });
      return Array.from(depts).sort();
    } catch (error) {
      logger.error('PolicyReportExportService', 'Failed to fetch departments:', error);
      return [];
    }
  }

  // ============================================================================
  // PDF REPORT GENERATION (via browser print-to-PDF)
  // ============================================================================

  /**
   * Generate a branded PDF report via ReportHtmlGenerator + browser print dialog.
   * Returns the same IExportResult so callers don't need to know the format.
   */
  public async generatePdfReport(
    reportKey: string,
    reportTitle: string,
    options: { dateRangeStart?: Date; dateRangeEnd?: Date; departments?: string[] } = {}
  ): Promise<IExportResult> {
    try {
      logger.info('PolicyReportExportService', `Generating PDF for ${reportKey}`);

      const sections: IReportSection[] = [];
      let recordCount = 0;

      switch (reportKey) {
        case 'dept-compliance': {
          const preview = await this.getCompliancePreview(options);
          const t = preview.totals;
          sections.push({
            type: 'kpi-row',
            data: [
              { label: 'Compliance Rate', value: `${t.rate}%`, color: t.rate >= 80 ? '#059669' : '#d97706' },
              { label: 'Total Assigned', value: String(t.assigned) },
              { label: 'Acknowledged', value: String(t.acknowledged), color: '#059669' },
              { label: 'Overdue', value: String(t.overdue), color: t.overdue > 0 ? '#dc2626' : '#059669' }
            ] as IKpiItem[]
          });
          sections.push({ type: 'divider' });
          sections.push({
            type: 'table',
            title: 'Department Breakdown',
            data: {
              headers: ['Department', 'Assigned', 'Acknowledged', 'Pending', 'Overdue', 'Rate'],
              rows: preview.departments.map(d => [d.department, String(d.assigned), String(d.acknowledged), String(d.pending), String(d.overdue), `${d.rate}%`])
            } as ITableData
          });
          recordCount = preview.departments.length;
          break;
        }
        case 'ack-status': {
          const acks = await this.sp.web.lists.getByTitle(this.ACKNOWLEDGEMENT_LIST).items
            .select('Id', 'PolicyName', 'PolicyNumber', 'AckStatus', 'DueDate', 'UserDepartment', 'User/Title')
            .expand('User').top(100)();
          const pending = acks.filter((a: any) => a.AckStatus !== 'Acknowledged' && a.AckStatus !== 'completed');
          sections.push({
            type: 'kpi-row',
            data: [
              { label: 'Total', value: String(acks.length) },
              { label: 'Pending', value: String(pending.length), color: '#d97706' }
            ] as IKpiItem[]
          });
          sections.push({
            type: 'table',
            title: 'Pending & Overdue Acknowledgements',
            data: {
              headers: ['Policy', 'User', 'Department', 'Status', 'Due Date'],
              rows: pending.slice(0, 50).map((a: any) => [a.PolicyName || '', a.User?.Title || '', a.UserDepartment || '', a.AckStatus || '', a.DueDate ? new Date(a.DueDate).toLocaleDateString('en-GB') : ''])
            } as ITableData
          });
          recordCount = pending.length;
          break;
        }
        case 'review-schedule': {
          const policies = await this.sp.web.lists.getByTitle(this.POLICY_LIST).items
            .select('Id', 'PolicyName', 'PolicyNumber', 'PolicyCategory', 'NextReviewDate', 'ReviewFrequency', 'PolicyOwner/Title')
            .expand('PolicyOwner')
            .filter("Status eq 'Published' and NextReviewDate ne null")
            .orderBy('NextReviewDate', true).top(100)();
          const now = new Date();
          sections.push({
            type: 'table',
            title: 'Policy Review Schedule',
            data: {
              headers: ['Policy', 'Category', 'Owner', 'Frequency', 'Next Review', 'Urgency'],
              rows: policies.map((p: any) => {
                const days = p.NextReviewDate ? Math.ceil((new Date(p.NextReviewDate).getTime() - now.getTime()) / 86400000) : 0;
                return [p.PolicyName || '', p.PolicyCategory || '', p.PolicyOwner?.Title || '', p.ReviewFrequency || 'Annual', p.NextReviewDate ? new Date(p.NextReviewDate).toLocaleDateString('en-GB') : '', days < 0 ? 'OVERDUE' : days <= 30 ? 'Due Soon' : 'On Track'];
              })
            } as ITableData
          });
          recordCount = policies.length;
          break;
        }
        case 'sla-performance':
        case 'executive-summary': {
          const result = await this.getCompliancePreview(options);
          sections.push({
            type: 'kpi-row',
            data: [
              { label: 'Overall Compliance', value: `${result.totals.rate}%`, color: result.totals.rate >= 80 ? '#059669' : '#d97706' },
              { label: 'Assigned', value: String(result.totals.assigned) },
              { label: 'Overdue', value: String(result.totals.overdue), color: result.totals.overdue > 0 ? '#dc2626' : '#059669' }
            ] as IKpiItem[]
          });
          sections.push({
            type: 'table',
            title: 'Summary by Department',
            data: {
              headers: ['Department', 'Assigned', 'Acknowledged', 'Rate'],
              rows: result.departments.map(d => [d.department, String(d.assigned), String(d.acknowledged), `${d.rate}%`])
            } as ITableData
          });
          recordCount = result.departments.length;
          break;
        }
        default: {
          // Generic: use compliance preview
          const fallback = await this.getCompliancePreview(options);
          sections.push({
            type: 'summary-card',
            title: 'Report Summary',
            content: `This report covers ${fallback.totals.assigned} acknowledgements across ${fallback.departments.length} departments with an overall compliance rate of ${fallback.totals.rate}%.`
          });
          recordCount = fallback.departments.length;
          break;
        }
      }

      const dateRange = options.dateRangeStart || options.dateRangeEnd
        ? `${options.dateRangeStart ? options.dateRangeStart.toLocaleDateString('en-GB') : 'Start'} — ${options.dateRangeEnd ? options.dateRangeEnd.toLocaleDateString('en-GB') : 'Now'}`
        : undefined;

      const config: IReportConfig = {
        title: reportTitle,
        subtitle: dateRange,
        reportType: reportKey.replace(/-/g, ' ').replace(/\b\w/g, c => c.toUpperCase()),
        sections
      };

      ReportHtmlGenerator.printReport(config);

      return {
        success: true,
        filename: `${reportTitle.replace(/\s+/g, '_')}.pdf`,
        recordCount,
        exportedAt: new Date()
      };
    } catch (error) {
      logger.error('PolicyReportExportService', 'Failed to generate PDF report:', error);
      return { success: false, filename: '', recordCount: 0, exportedAt: new Date(), errors: [error instanceof Error ? error.message : 'Unknown error'] };
    }
  }

  // ============================================================================
  // REPORT STORAGE — save CSV to SharePoint document library
  // ============================================================================

  /**
   * Save generated CSV content to PM_PolicySourceDocuments/Reports/ folder.
   * Returns the server-relative URL of the uploaded file.
   */
  public async saveReportToLibrary(csvData: Record<string, any>[], filename: string): Promise<string | null> {
    try {
      const folderPath = 'PM_PolicySourceDocuments/Reports';

      // Ensure the Reports folder exists
      try {
        await this.sp.web.folders.addUsingPath(folderPath);
      } catch { /* folder may already exist — non-blocking */ }

      // Build CSV content
      if (!csvData || csvData.length === 0) return null;
      const headers = Object.keys(csvData[0]);
      const rows: string[][] = [headers];
      for (const item of csvData) {
        const row: string[] = [];
        for (const header of headers) {
          let value = item[header];
          if (value === null || value === undefined) value = '';
          else if (typeof value === 'object') value = JSON.stringify(value);
          const str = String(value);
          row.push(str.includes(',') || str.includes('"') || str.includes('\n') ? `"${str.replace(/"/g, '""')}"` : str);
        }
        rows.push(row);
      }
      const BOM = '\uFEFF';
      const csvContent = BOM + rows.map(r => r.join(',')).join('\n');

      // Upload
      const fileBlob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const buffer = await fileBlob.arrayBuffer();
      const result = await this.sp.web
        .getFolderByServerRelativePath(folderPath)
        .files.addUsingPath(filename, new Uint8Array(buffer), { Overwrite: true });

      const fileUrl = result.data?.ServerRelativeUrl || '';
      logger.info('PolicyReportExportService', `Report saved to ${fileUrl}`);
      return fileUrl;
    } catch (error) {
      logger.error('PolicyReportExportService', 'Failed to save report to library:', error);
      return null;
    }
  }

  // ============================================================================
  // UTILITY METHODS
  // ============================================================================

  /**
   * Convert data to CSV and trigger download
   */
  private downloadCSV(data: Record<string, any>[], filename: string): void {
    if (!data || data.length === 0) {
      throw new Error('No data to export');
    }

    const headers = Object.keys(data[0]);
    const rows: string[][] = [headers];

    for (const item of data) {
      const row: string[] = [];
      for (const header of headers) {
        let value = item[header];

        if (value === null || value === undefined) {
          value = '';
        } else if (value instanceof Date) {
          value = value.toISOString().split('T')[0];
        } else if (typeof value === 'object') {
          value = JSON.stringify(value);
        }

        const stringValue = String(value);
        // Escape CSV special characters
        const escaped = stringValue.includes(',') || stringValue.includes('"') || stringValue.includes('\n')
          ? `"${stringValue.replace(/"/g, '""')}"`
          : stringValue;

        row.push(escaped);
      }
      rows.push(row);
    }

    const csvContent = rows.map(row => row.join(',')).join('\n');

    // Add BOM for Excel compatibility with UTF-8
    const BOM = '\uFEFF';
    const blob = new Blob([BOM + csvContent], { type: 'text/csv;charset=utf-8;' });

    // Download
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
   * Format date for display
   */
  private formatDate(date: Date | string | null | undefined): string {
    if (!date) return '';
    const d = typeof date === 'string' ? new Date(date) : date;
    if (isNaN(d.getTime())) return '';
    return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  }

  /**
   * Format date and time for display
   */
  private formatDateTime(date: Date | string | null | undefined): string {
    if (!date) return '';
    const d = typeof date === 'string' ? new Date(date) : date;
    if (isNaN(d.getTime())) return '';
    return d.toLocaleString('en-GB', {
      day: '2-digit',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  }

  /**
   * Get date stamp for filename
   */
  private getDateStamp(): string {
    const now = new Date();
    return `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
  }
}
