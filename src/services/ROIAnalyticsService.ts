// @ts-nocheck
// ROI Analytics Service
// Measures the business impact and ROI of Phase 1 & 2 automation features

import { SPFI } from '@pnp/sp';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { logger } from './LoggingService';

/**
 * Phase 1 Impact Metrics - Employee Master Data & Lookup
 */
export interface IEmployeeLookupMetrics {
  // Usage Statistics
  totalProcessesCreated: number;
  processesUsingEmployeePicker: number;
  employeePickerUsageRate: number; // Percentage

  // Time Savings
  averageManualEntryTimeMinutes: number; // Baseline: 5 minutes
  averageLookupTimeSeconds: number; // Actual: 15 seconds
  timeSavedPerProcess: number; // Minutes
  totalTimeSavedHours: number;

  // Data Quality
  dataAccuracyRate: number; // % fewer errors with lookup vs manual
  duplicateEmployeesPrevented: number;
}

/**
 * Phase 1 Impact Metrics - Task Automation
 */
export interface ITaskAutomationMetrics {
  // Usage Statistics
  processesWithTemplates: number;
  totalTasksGenerated: number;
  averageTasksPerProcess: number;

  // Time Savings
  manualTaskCreationTimeMinutes: number; // Baseline: 2 minutes per task
  automatedTaskCreationSeconds: number; // Actual: instant
  timeSavedPerProcess: number; // Minutes
  totalTimeSavedHours: number;

  // Accuracy
  taskDependenciesAutomated: number;
  slaCalculationsAutomated: number;
}

/**
 * Phase 1 Impact Metrics - Smart Notifications
 */
export interface INotificationMetrics {
  // Notification Activity
  remindersGenerated: number;
  escalationsSent: number;
  digestsSent: number;

  // Effectiveness
  tasksCompletedAfterReminder: number;
  averageResponseTimeHours: number;
  overdueTaskReduction: number; // Percentage reduction

  // User Engagement
  userPreferencesConfigured: number;
  notificationOpenRate: number; // Percentage
}

/**
 * Phase 2 Impact Metrics - Approval Workflows
 */
export interface IApprovalWorkflowMetrics {
  // Usage Statistics
  processesRequiringApproval: number;
  totalApprovalRequests: number;
  approvalChainLevels: number; // Average

  // Performance
  averageApprovalTimeHours: number;
  approvalSLAComplianceRate: number;
  overdueApprovalsCount: number;

  // Actions
  approvedCount: number;
  rejectedCount: number;
  delegatedCount: number;
  approvalRate: number; // Percentage
  delegationRate: number; // Percentage

  // Time Savings
  automatedNotifications: number;
  manualFollowUpsSaved: number;
  timeSavedOnFollowUpsHours: number;
}

/**
 * Overall ROI Summary
 */
export interface IROISummary {
  // Time Period
  startDate: Date;
  endDate: Date;
  daysInPeriod: number;

  // Adoption Metrics
  totalProcesses: number;
  automationAdoptionRate: number; // % of processes using automation features

  // Time Savings (converted to FTE)
  totalHoursSaved: number;
  fteEquivalent: number; // Hours / 2080 (annual hours)

  // Cost Savings
  hourlyRate: number; // Average employee hourly rate
  totalCostSavings: number;
  annualizedSavings: number;

  // ROI Calculation
  implementationCost: number; // Estimated dev + training cost
  roi: number; // (Savings - Cost) / Cost * 100
  paybackMonths: number; // Months to recover implementation cost

  // Feature Breakdown
  employeeLookupSavings: {
    hours: number;
    cost: number;
    percentage: number;
  };
  taskAutomationSavings: {
    hours: number;
    cost: number;
    percentage: number;
  };
  notificationSavings: {
    hours: number;
    cost: number;
    percentage: number;
  };
  approvalSavings: {
    hours: number;
    cost: number;
    percentage: number;
  };

  // Quality Improvements
  errorReductionRate: number;
  complianceImprovementRate: number;
  userSatisfactionScore: number; // 1-10 scale
}

/**
 * Trend comparison: Before vs After automation
 */
export interface IBeforeAfterMetrics {
  metricName: string;
  beforeValue: number;
  afterValue: number;
  improvement: number; // Absolute change
  improvementPercentage: number;
  unit: string; // 'hours', 'days', 'percentage', etc.
}

export class ROIAnalyticsService {
  private sp: SPFI;

  // Baseline constants (configurable)
  private readonly MANUAL_ENTRY_TIME_MINUTES = 5;
  private readonly MANUAL_TASK_CREATION_MINUTES = 2;
  private readonly MANUAL_FOLLOW_UP_MINUTES = 10;
  private readonly AVERAGE_HOURLY_RATE = 50; // Default, can be configured

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Calculate Employee Lookup ROI
   */
  public async getEmployeeLookupMetrics(startDate: Date, endDate: Date): Promise<IEmployeeLookupMetrics> {
    try {
      // Get all processes in period
      const processes = await this.sp.web.lists
        .getByTitle('JML_Processes')
        .items.filter(`Created ge datetime'${startDate.toISOString()}' and Created le datetime'${endDate.toISOString()}'`)
        .select('Id', 'ProcessType', 'Created', 'EmployeeID')();

      // Count processes that used employee picker (have employee ID pre-filled)
      const processesWithEmployeeId = processes.filter(p => p.EmployeeID && p.ProcessType !== 'Joiner');

      const totalProcesses = processes.length;
      const pickerUsage = processesWithEmployeeId.length;
      const usageRate = totalProcesses > 0 ? (pickerUsage / totalProcesses) * 100 : 0;

      // Calculate time savings
      const timeSavedPerProcess = this.MANUAL_ENTRY_TIME_MINUTES - 0.25; // 5 min manual vs 15 sec lookup
      const totalTimeSaved = (pickerUsage * timeSavedPerProcess) / 60; // Convert to hours

      return {
        totalProcessesCreated: totalProcesses,
        processesUsingEmployeePicker: pickerUsage,
        employeePickerUsageRate: usageRate,
        averageManualEntryTimeMinutes: this.MANUAL_ENTRY_TIME_MINUTES,
        averageLookupTimeSeconds: 15,
        timeSavedPerProcess: timeSavedPerProcess,
        totalTimeSavedHours: totalTimeSaved,
        dataAccuracyRate: 98, // Estimated - lookup has fewer errors
        duplicateEmployeesPrevented: Math.floor(pickerUsage * 0.1) // 10% would have created duplicates
      };
    } catch (error) {
      logger.error('ROIAnalyticsService', 'Failed to calculate employee lookup metrics:', error);
      return this.getDefaultEmployeeLookupMetrics();
    }
  }

  /**
   * Calculate Task Automation ROI
   */
  public async getTaskAutomationMetrics(startDate: Date, endDate: Date): Promise<ITaskAutomationMetrics> {
    try {
      // Get processes with templates
      const processes = await this.sp.web.lists
        .getByTitle('JML_Processes')
        .items.filter(`Created ge datetime'${startDate.toISOString()}' and Created le datetime'${endDate.toISOString()}' and ChecklistTemplateID ne null`)
        .select('Id', 'ChecklistTemplateID')();

      // Get task assignments created via automation
      const tasks = await this.sp.web.lists
        .getByTitle('JML_TaskAssignments')
        .items.filter(`Created ge datetime'${startDate.toISOString()}' and Created le datetime'${endDate.toISOString()}'`)
        .select('Id', 'ProcessID')();

      const processesWithTemplates = processes.length;
      const totalTasksGenerated = tasks.length;
      const avgTasksPerProcess = processesWithTemplates > 0 ? totalTasksGenerated / processesWithTemplates : 0;

      // Calculate time savings
      const manualTimePerTask = this.MANUAL_TASK_CREATION_MINUTES;
      const totalManualTime = (totalTasksGenerated * manualTimePerTask) / 60; // Hours
      const timeSavedPerProcess = (avgTasksPerProcess * manualTimePerTask);

      return {
        processesWithTemplates,
        totalTasksGenerated,
        averageTasksPerProcess: avgTasksPerProcess,
        manualTaskCreationTimeMinutes: manualTimePerTask,
        automatedTaskCreationSeconds: 1, // Nearly instant
        timeSavedPerProcess,
        totalTimeSavedHours: totalManualTime,
        taskDependenciesAutomated: Math.floor(totalTasksGenerated * 0.3), // 30% have dependencies
        slaCalculationsAutomated: totalTasksGenerated
      };
    } catch (error) {
      logger.error('ROIAnalyticsService', 'Failed to calculate task automation metrics:', error);
      return this.getDefaultTaskAutomationMetrics();
    }
  }

  /**
   * Calculate Smart Notification ROI
   */
  public async getNotificationMetrics(startDate: Date, endDate: Date): Promise<INotificationMetrics> {
    try {
      // Note: This would require a notification log table in production
      // For now, we'll estimate based on task completion patterns

      const tasks = await this.sp.web.lists
        .getByTitle('JML_TaskAssignments')
        .items.filter(`Created ge datetime'${startDate.toISOString()}' and Created le datetime'${endDate.toISOString()}'`)
        .select('Id', 'Status', 'DueDate', 'CompletedDate')();

      const completedTasks = tasks.filter(t => t.Status === 'Completed');
      const overdueTasks = tasks.filter(t => {
        if (t.Status !== 'Completed' && t.DueDate) {
          const dueDate = new Date(t.DueDate);
          return dueDate < new Date();
        }
        return false;
      });

      // Estimate reminders (3 days before + 1 day before for each task)
      const estimatedReminders = tasks.length * 2;

      // Estimate effectiveness (tasks completed within SLA after reminder)
      const tasksCompletedOnTime = completedTasks.filter(t => {
        if (t.DueDate && t.CompletedDate) {
          return new Date(t.CompletedDate) <= new Date(t.DueDate);
        }
        return false;
      }).length;

      return {
        remindersGenerated: estimatedReminders,
        escalationsSent: overdueTasks.length,
        digestsSent: Math.ceil((endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24)) * 10, // Daily digests to ~10 users
        tasksCompletedAfterReminder: tasksCompletedOnTime,
        averageResponseTimeHours: 24,
        overdueTaskReduction: 80, // Target from Phase 1
        userPreferencesConfigured: 15, // Estimated
        notificationOpenRate: 75 // Estimated
      };
    } catch (error) {
      logger.error('ROIAnalyticsService', 'Failed to calculate notification metrics:', error);
      return this.getDefaultNotificationMetrics();
    }
  }

  /**
   * Calculate Approval Workflow ROI
   */
  public async getApprovalWorkflowMetrics(startDate: Date, endDate: Date): Promise<IApprovalWorkflowMetrics> {
    try {
      const approvals = await this.sp.web.lists
        .getByTitle('JML_Approvals')
        .items.filter(`RequestedDate ge datetime'${startDate.toISOString()}' and RequestedDate le datetime'${endDate.toISOString()}'`)
        .select('Id', 'Status', 'RequestedDate', 'CompletedDate', 'ProcessID')();

      const approved = approvals.filter(a => a.Status === 'Approved').length;
      const rejected = approvals.filter(a => a.Status === 'Rejected').length;
      const delegated = approvals.filter(a => a.Status === 'Delegated').length;
      const overdue = approvals.filter(a => a.IsOverdue === true).length;

      // Calculate average approval time
      const completedApprovals = approvals.filter(a => a.CompletedDate && a.RequestedDate);
      let totalApprovalHours = 0;
      for (let i = 0; i < completedApprovals.length; i++) {
        const requested = new Date(completedApprovals[i].RequestedDate);
        const completed = new Date(completedApprovals[i].CompletedDate);
        const hours = (completed.getTime() - requested.getTime()) / (1000 * 60 * 60);
        totalApprovalHours += hours;
      }
      const avgApprovalTime = completedApprovals.length > 0 ? totalApprovalHours / completedApprovals.length : 0;

      // Get unique processes requiring approval
      const uniqueProcesses = new Set(approvals.map(a => a.ProcessID));

      return {
        processesRequiringApproval: uniqueProcesses.size,
        totalApprovalRequests: approvals.length,
        approvalChainLevels: 2, // Average estimated
        averageApprovalTimeHours: avgApprovalTime,
        approvalSLAComplianceRate: overdue > 0 ? ((approvals.length - overdue) / approvals.length) * 100 : 100,
        overdueApprovalsCount: overdue,
        approvedCount: approved,
        rejectedCount: rejected,
        delegatedCount: delegated,
        approvalRate: approvals.length > 0 ? (approved / approvals.length) * 100 : 0,
        delegationRate: approvals.length > 0 ? (delegated / approvals.length) * 100 : 0,
        automatedNotifications: approvals.length, // Each approval sends notification
        manualFollowUpsSaved: approvals.length * 2, // Estimate 2 follow-ups per approval saved
        timeSavedOnFollowUpsHours: (approvals.length * 2 * this.MANUAL_FOLLOW_UP_MINUTES) / 60
      };
    } catch (error) {
      logger.error('ROIAnalyticsService', 'Failed to calculate approval workflow metrics:', error);
      return this.getDefaultApprovalMetrics();
    }
  }

  /**
   * Calculate Overall ROI
   */
  public async calculateROI(startDate: Date, endDate: Date, implementationCost: number = 50000): Promise<IROISummary> {
    try {
      const [employeeLookup, taskAutomation, notification, approval] = await Promise.all([
        this.getEmployeeLookupMetrics(startDate, endDate),
        this.getTaskAutomationMetrics(startDate, endDate),
        this.getNotificationMetrics(startDate, endDate),
        this.getApprovalWorkflowMetrics(startDate, endDate)
      ]);

      // Calculate total hours saved
      const totalHoursSaved =
        employeeLookup.totalTimeSavedHours +
        taskAutomation.totalTimeSavedHours +
        (notification.remindersGenerated * 0.5) + // 30 min saved per reminder cycle
        approval.timeSavedOnFollowUpsHours;

      // Calculate cost savings
      const hourlyRate = this.AVERAGE_HOURLY_RATE;
      const totalCostSavings = totalHoursSaved * hourlyRate;

      // Annualize savings
      const daysInPeriod = (endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24);
      const annualizedSavings = (totalCostSavings / daysInPeriod) * 365;

      // Calculate ROI
      const roi = ((annualizedSavings - implementationCost) / implementationCost) * 100;
      const paybackMonths = (implementationCost / annualizedSavings) * 12;

      // FTE equivalent
      const fteEquivalent = totalHoursSaved / 2080; // Annual work hours

      // Feature breakdown percentages
      const employeeCost = employeeLookup.totalTimeSavedHours * hourlyRate;
      const taskCost = taskAutomation.totalTimeSavedHours * hourlyRate;
      const notificationCost = (notification.remindersGenerated * 0.5) * hourlyRate;
      const approvalCost = approval.timeSavedOnFollowUpsHours * hourlyRate;

      const totalProcesses = employeeLookup.totalProcessesCreated;
      const automationUsage = employeeLookup.processesUsingEmployeePicker + taskAutomation.processesWithTemplates;
      const automationAdoptionRate = totalProcesses > 0 ? (automationUsage / (totalProcesses * 2)) * 100 : 0;

      return {
        startDate,
        endDate,
        daysInPeriod,
        totalProcesses,
        automationAdoptionRate,
        totalHoursSaved,
        fteEquivalent,
        hourlyRate,
        totalCostSavings,
        annualizedSavings,
        implementationCost,
        roi,
        paybackMonths,
        employeeLookupSavings: {
          hours: employeeLookup.totalTimeSavedHours,
          cost: employeeCost,
          percentage: totalCostSavings > 0 ? (employeeCost / totalCostSavings) * 100 : 0
        },
        taskAutomationSavings: {
          hours: taskAutomation.totalTimeSavedHours,
          cost: taskCost,
          percentage: totalCostSavings > 0 ? (taskCost / totalCostSavings) * 100 : 0
        },
        notificationSavings: {
          hours: notification.remindersGenerated * 0.5,
          cost: notificationCost,
          percentage: totalCostSavings > 0 ? (notificationCost / totalCostSavings) * 100 : 0
        },
        approvalSavings: {
          hours: approval.timeSavedOnFollowUpsHours,
          cost: approvalCost,
          percentage: totalCostSavings > 0 ? (approvalCost / totalCostSavings) * 100 : 0
        },
        errorReductionRate: 75, // Estimated
        complianceImprovementRate: 90, // Estimated with approval workflows
        userSatisfactionScore: 8.5 // Estimated
      };
    } catch (error) {
      logger.error('ROIAnalyticsService', 'Failed to calculate ROI:', error);
      throw error;
    }
  }

  /**
   * Generate before/after comparison
   */
  public async getBeforeAfterComparison(beforeStart: Date, beforeEnd: Date, afterStart: Date, afterEnd: Date): Promise<IBeforeAfterMetrics[]> {
    // This would compare metrics before automation vs after
    // For now, return estimated improvements
    return [
      {
        metricName: 'Average Process Completion Time',
        beforeValue: 14,
        afterValue: 10,
        improvement: -4,
        improvementPercentage: -28.6,
        unit: 'days'
      },
      {
        metricName: 'Data Entry Time per Process',
        beforeValue: 25,
        afterValue: 5,
        improvement: -20,
        improvementPercentage: -80,
        unit: 'minutes'
      },
      {
        metricName: 'Overdue Task Rate',
        beforeValue: 35,
        afterValue: 7,
        improvement: -28,
        improvementPercentage: -80,
        unit: 'percentage'
      },
      {
        metricName: 'Approval Response Time',
        beforeValue: 72,
        afterValue: 24,
        improvement: -48,
        improvementPercentage: -66.7,
        unit: 'hours'
      }
    ];
  }

  // Default/fallback methods
  private getDefaultEmployeeLookupMetrics(): IEmployeeLookupMetrics {
    return {
      totalProcessesCreated: 0,
      processesUsingEmployeePicker: 0,
      employeePickerUsageRate: 0,
      averageManualEntryTimeMinutes: this.MANUAL_ENTRY_TIME_MINUTES,
      averageLookupTimeSeconds: 15,
      timeSavedPerProcess: 0,
      totalTimeSavedHours: 0,
      dataAccuracyRate: 0,
      duplicateEmployeesPrevented: 0
    };
  }

  private getDefaultTaskAutomationMetrics(): ITaskAutomationMetrics {
    return {
      processesWithTemplates: 0,
      totalTasksGenerated: 0,
      averageTasksPerProcess: 0,
      manualTaskCreationTimeMinutes: this.MANUAL_TASK_CREATION_MINUTES,
      automatedTaskCreationSeconds: 1,
      timeSavedPerProcess: 0,
      totalTimeSavedHours: 0,
      taskDependenciesAutomated: 0,
      slaCalculationsAutomated: 0
    };
  }

  private getDefaultNotificationMetrics(): INotificationMetrics {
    return {
      remindersGenerated: 0,
      escalationsSent: 0,
      digestsSent: 0,
      tasksCompletedAfterReminder: 0,
      averageResponseTimeHours: 0,
      overdueTaskReduction: 0,
      userPreferencesConfigured: 0,
      notificationOpenRate: 0
    };
  }

  private getDefaultApprovalMetrics(): IApprovalWorkflowMetrics {
    return {
      processesRequiringApproval: 0,
      totalApprovalRequests: 0,
      approvalChainLevels: 0,
      averageApprovalTimeHours: 0,
      approvalSLAComplianceRate: 0,
      overdueApprovalsCount: 0,
      approvedCount: 0,
      rejectedCount: 0,
      delegatedCount: 0,
      approvalRate: 0,
      delegationRate: 0,
      automatedNotifications: 0,
      manualFollowUpsSaved: 0,
      timeSavedOnFollowUpsHours: 0
    };
  }
}
