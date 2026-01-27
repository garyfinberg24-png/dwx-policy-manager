// @ts-nocheck
// Success Metrics Service
// Tracks and calculates key performance indicators for JML processes

import { SPFI } from '@pnp/sp';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import {
  ITimeToOnboard,
  ITaskCompletionRate,
  IUserAdoption,
  IProcessCycleTime,
  IErrorRate,
  IUserSatisfaction,
  ICostSavings,
  IComplianceMetric,
  ISuccessMetricsSummary,
  IAnalyticsFilters
} from '../models';
import { IJmlProcess, IJmlTaskAssignment, IJmlAuditLog } from '../models';
import { ProcessStatus, ProcessType } from '../models/ICommon';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class SuccessMetricsService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Calculate Time to Onboard metrics
   * Days from hire date to first-day ready status
   */
  public async getTimeToOnboard(filters?: IAnalyticsFilters): Promise<ITimeToOnboard[]> {
    try {
      const processes = await this.getFilteredProcesses(filters);
      const joinerProcesses = processes.filter(p => p.ProcessType === ProcessType.Joiner);
      const metrics: ITimeToOnboard[] = [];

      for (let i = 0; i < joinerProcesses.length; i++) {
        const process = joinerProcesses[i];
        if (!process.StartDate) {
          continue;
        }

        const hireDate = typeof process.StartDate === 'string'
          ? new Date(process.StartDate)
          : process.StartDate;

        let firstDayReadyDate: Date | undefined;
        let daysToOnboard = 0;
        const targetDays = 5;

        if (process.ActualCompletionDate) {
          firstDayReadyDate = typeof process.ActualCompletionDate === 'string'
            ? new Date(process.ActualCompletionDate)
            : process.ActualCompletionDate;
          daysToOnboard = Math.floor((firstDayReadyDate.getTime() - hireDate.getTime()) / (1000 * 60 * 60 * 24));
        } else if (process.ProcessStatus === ProcessStatus.InProgress) {
          const now = new Date();
          daysToOnboard = Math.floor((now.getTime() - hireDate.getTime()) / (1000 * 60 * 60 * 24));
        }

        metrics.push({
          employeeId: typeof process.EmployeeID === 'number' ? process.EmployeeID : (typeof process.EmployeeID === 'string' ? parseInt(process.EmployeeID, 10) : 0),
          employeeName: process.EmployeeName,
          hireDate,
          firstDayReadyDate,
          daysToOnboard,
          processType: process.ProcessType,
          department: process.Department,
          isCompliant: daysToOnboard <= targetDays,
          targetDays
        });
      }

      return metrics;
    } catch (error) {
      logger.error('SuccessMetricsService', 'Failed to calculate time to onboard:', error);
      return [];
    }
  }

  /**
   * Calculate Task Completion Rate
   * Percentage of tasks completed on time
   */
  public async getTaskCompletionRate(filters?: IAnalyticsFilters): Promise<ITaskCompletionRate[]> {
    try {
      const tasks = await this.getAllTasks(filters);
      const monthlyMetrics: { [key: string]: ITaskCompletionRate } = {};

      for (let i = 0; i < tasks.length; i++) {
        const task = tasks[i];
        if (!task.DueDate) {
          continue;
        }

        const dueDate = typeof task.DueDate === 'string' ? new Date(task.DueDate) : task.DueDate;
        const monthKey = `${dueDate.getFullYear()}-${dueDate.getMonth() + 1}`;

        if (!monthlyMetrics[monthKey]) {
          monthlyMetrics[monthKey] = {
            period: new Date(dueDate.getFullYear(), dueDate.getMonth(), 1),
            totalTasks: 0,
            completedOnTime: 0,
            completedLate: 0,
            notCompleted: 0,
            onTimeRate: 0
          };
        }

        monthlyMetrics[monthKey].totalTasks++;

        if (task.ActualCompletionDate) {
          const completedDate = typeof task.ActualCompletionDate === 'string'
            ? new Date(task.ActualCompletionDate)
            : task.ActualCompletionDate;

          if (completedDate <= dueDate) {
            monthlyMetrics[monthKey].completedOnTime++;
          } else {
            monthlyMetrics[monthKey].completedLate++;
          }
        } else {
          const now = new Date();
          if (now > dueDate) {
            monthlyMetrics[monthKey].notCompleted++;
          }
        }
      }

      const metrics: ITaskCompletionRate[] = [];
      const keys = Object.keys(monthlyMetrics);
      for (let i = 0; i < keys.length; i++) {
        const key = keys[i];
        const metric = monthlyMetrics[key];
        metric.onTimeRate = metric.totalTasks > 0
          ? (metric.completedOnTime / metric.totalTasks) * 100
          : 0;
        metrics.push(metric);
      }

      return metrics;
    } catch (error) {
      logger.error('SuccessMetricsService', 'Failed to calculate task completion rate:', error);
      return [];
    }
  }

  /**
   * Calculate User Adoption metrics
   * Active users per month
   */
  public async getUserAdoption(filters?: IAnalyticsFilters): Promise<IUserAdoption[]> {
    try {
      const auditLogs = await this.getAuditLogs(filters);
      const monthlyMetrics: { [key: string]: IUserAdoption } = {};
      const userActivity: { [month: string]: Set<number> } = {};

      for (let i = 0; i < auditLogs.length; i++) {
        const log = auditLogs[i];
        const logDate = typeof log.Timestamp === 'string'
          ? new Date(log.Timestamp)
          : log.Timestamp;
        const monthKey = `${logDate.getFullYear()}-${logDate.getMonth() + 1}`;

        if (!userActivity[monthKey]) {
          userActivity[monthKey] = new Set<number>();
        }

        if (log.UserId) {
          userActivity[monthKey].add(log.UserId);
        }
      }

      const totalUsers = await this.getTotalUsers();
      const keys = Object.keys(userActivity);

      for (let i = 0; i < keys.length; i++) {
        const key = keys[i];
        const parts = key.split('-');
        const year = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1;

        const activeUsers = userActivity[key].size;
        const adoptionRate = totalUsers > 0 ? (activeUsers / totalUsers) * 100 : 0;
        const engagementScore = Math.min(100, (activeUsers / Math.max(1, totalUsers / 10)) * 100);

        monthlyMetrics[key] = {
          month: new Date(year, month, 1),
          activeUsers,
          totalUsers,
          newUsers: 0,
          returningUsers: activeUsers,
          adoptionRate,
          engagementScore
        };
      }

      const metrics: IUserAdoption[] = [];
      const metricKeys = Object.keys(monthlyMetrics);
      for (let i = 0; i < metricKeys.length; i++) {
        metrics.push(monthlyMetrics[metricKeys[i]]);
      }

      return metrics;
    } catch (error) {
      logger.error('SuccessMetricsService', 'Failed to calculate user adoption:', error);
      return [];
    }
  }

  /**
   * Calculate Process Cycle Time
   * Average days to complete a process
   */
  public async getProcessCycleTime(filters?: IAnalyticsFilters): Promise<IProcessCycleTime[]> {
    try {
      const processes = await this.getFilteredProcesses(filters);
      const completedProcesses = processes.filter(p =>
        p.ProcessStatus === ProcessStatus.Completed &&
        p.StartDate &&
        p.ActualCompletionDate
      );

      const typeMetrics: { [key: string]: { days: number[]; department?: string } } = {};

      for (let i = 0; i < completedProcesses.length; i++) {
        const process = completedProcesses[i];
        const startDate = typeof process.StartDate === 'string'
          ? new Date(process.StartDate)
          : process.StartDate;
        const completedDate = typeof process.ActualCompletionDate === 'string'
          ? new Date(process.ActualCompletionDate)
          : process.ActualCompletionDate;

        if (!completedDate) {
          continue;
        }

        const days = Math.floor((completedDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24));

        if (!typeMetrics[process.ProcessType]) {
          typeMetrics[process.ProcessType] = {
            days: [],
            department: process.Department
          };
        }

        typeMetrics[process.ProcessType].days.push(days);
      }

      const slaTargets: { [key: string]: number } = {
        [ProcessType.Joiner]: 5,
        [ProcessType.Mover]: 3,
        [ProcessType.Leaver]: 7
      };

      const metrics: IProcessCycleTime[] = [];
      const keys = Object.keys(typeMetrics);

      for (let i = 0; i < keys.length; i++) {
        const processType = keys[i];
        const data = typeMetrics[processType];
        const days = data.days;

        if (days.length === 0) {
          continue;
        }

        days.sort((a, b) => a - b);

        const sum = days.reduce((acc, val) => acc + val, 0);
        const averageDays = sum / days.length;
        const medianDays = days.length % 2 === 0
          ? (days[days.length / 2 - 1] + days[days.length / 2]) / 2
          : days[Math.floor(days.length / 2)];
        const minDays = days[0];
        const maxDays = days[days.length - 1];
        const targetDays = slaTargets[processType] || 5;

        metrics.push({
          processType,
          averageDays,
          medianDays,
          minDays,
          maxDays,
          totalProcesses: days.length,
          targetDays,
          varianceFromTarget: averageDays - targetDays,
          department: data.department
        });
      }

      return metrics;
    } catch (error) {
      logger.error('SuccessMetricsService', 'Failed to calculate process cycle time:', error);
      return [];
    }
  }

  /**
   * Calculate Error Rate
   * Failed processes and tasks
   */
  public async getErrorRate(filters?: IAnalyticsFilters): Promise<IErrorRate[]> {
    try {
      const processes = await this.getFilteredProcesses(filters);
      const tasks = await this.getAllTasks(filters);
      const monthlyMetrics: { [key: string]: IErrorRate } = {};

      for (let i = 0; i < processes.length; i++) {
        const process = processes[i];
        const startDate = process.StartDate
          ? (typeof process.StartDate === 'string' ? new Date(process.StartDate) : process.StartDate)
          : new Date();
        const monthKey = `${startDate.getFullYear()}-${startDate.getMonth() + 1}`;

        if (!monthlyMetrics[monthKey]) {
          monthlyMetrics[monthKey] = {
            period: new Date(startDate.getFullYear(), startDate.getMonth(), 1),
            totalProcesses: 0,
            failedProcesses: 0,
            totalTasks: 0,
            failedTasks: 0,
            processErrorRate: 0,
            taskErrorRate: 0,
            topErrorReasons: []
          };
        }

        monthlyMetrics[monthKey].totalProcesses++;
        if (process.ProcessStatus === ProcessStatus.Cancelled) {
          monthlyMetrics[monthKey].failedProcesses++;
        }
      }

      for (let i = 0; i < tasks.length; i++) {
        const task = tasks[i];
        const dueDate = task.DueDate
          ? (typeof task.DueDate === 'string' ? new Date(task.DueDate) : task.DueDate)
          : new Date();
        const monthKey = `${dueDate.getFullYear()}-${dueDate.getMonth() + 1}`;

        if (monthlyMetrics[monthKey]) {
          monthlyMetrics[monthKey].totalTasks++;
          if (task.Status !== 'Completed' && task.IsOverdue) {
            monthlyMetrics[monthKey].failedTasks++;
          }
        }
      }

      const metrics: IErrorRate[] = [];
      const keys = Object.keys(monthlyMetrics);

      for (let i = 0; i < keys.length; i++) {
        const key = keys[i];
        const metric = monthlyMetrics[key];
        metric.processErrorRate = metric.totalProcesses > 0
          ? (metric.failedProcesses / metric.totalProcesses) * 100
          : 0;
        metric.taskErrorRate = metric.totalTasks > 0
          ? (metric.failedTasks / metric.totalTasks) * 100
          : 0;
        metrics.push(metric);
      }

      return metrics;
    } catch (error) {
      logger.error('SuccessMetricsService', 'Failed to calculate error rate:', error);
      return [];
    }
  }

  /**
   * Calculate User Satisfaction
   * NPS and CSAT scores
   */
  public async getUserSatisfaction(filters?: IAnalyticsFilters): Promise<IUserSatisfaction[]> {
    try {
      const satisfactionData = await this.getSatisfactionSurveys(filters);
      const monthlyMetrics: { [key: string]: IUserSatisfaction } = {};

      for (let i = 0; i < satisfactionData.length; i++) {
        const survey = satisfactionData[i];
        const surveyDate = typeof survey.Date === 'string'
          ? new Date(survey.Date)
          : survey.Date;
        const monthKey = `${surveyDate.getFullYear()}-${surveyDate.getMonth() + 1}`;

        if (!monthlyMetrics[monthKey]) {
          monthlyMetrics[monthKey] = {
            period: new Date(surveyDate.getFullYear(), surveyDate.getMonth(), 1),
            totalResponses: 0,
            npsScore: 0,
            csatScore: 0,
            promoters: 0,
            passives: 0,
            detractors: 0,
            averageRating: 0,
            topFeedback: []
          };
        }

        monthlyMetrics[monthKey].totalResponses++;

        if (survey.NPSScore !== undefined) {
          if (survey.NPSScore >= 9) {
            monthlyMetrics[monthKey].promoters++;
          } else if (survey.NPSScore >= 7) {
            monthlyMetrics[monthKey].passives++;
          } else {
            monthlyMetrics[monthKey].detractors++;
          }
        }
      }

      const metrics: IUserSatisfaction[] = [];
      const keys = Object.keys(monthlyMetrics);

      for (let i = 0; i < keys.length; i++) {
        const key = keys[i];
        const metric = monthlyMetrics[key];
        const total = metric.promoters + metric.passives + metric.detractors;

        if (total > 0) {
          metric.npsScore = ((metric.promoters - metric.detractors) / total) * 100;
          metric.csatScore = ((metric.promoters + metric.passives) / total) * 100;
          metric.averageRating = ((metric.promoters * 10 + metric.passives * 7.5 + metric.detractors * 5) / total);
        }

        metrics.push(metric);
      }

      return metrics;
    } catch (error) {
      logger.error('SuccessMetricsService', 'Failed to calculate user satisfaction:', error);
      return [];
    }
  }

  /**
   * Calculate Cost Savings
   * Estimated FTE hours saved through automation
   */
  public async getCostSavings(filters?: IAnalyticsFilters): Promise<ICostSavings[]> {
    try {
      const processes = await this.getFilteredProcesses(filters);
      const monthlyMetrics: { [key: string]: ICostSavings } = {};
      const avgCostPerHour = 50;
      const hoursPerFTE = 2080;

      for (let i = 0; i < processes.length; i++) {
        const process = processes[i];
        const startDate = process.StartDate
          ? (typeof process.StartDate === 'string' ? new Date(process.StartDate) : process.StartDate)
          : new Date();
        const monthKey = `${startDate.getFullYear()}-${startDate.getMonth() + 1}`;

        if (!monthlyMetrics[monthKey]) {
          monthlyMetrics[monthKey] = {
            period: new Date(startDate.getFullYear(), startDate.getMonth(), 1),
            automatedProcesses: 0,
            manualHoursSaved: 0,
            fteEquivalent: 0,
            costSavings: 0,
            avgCostPerHour,
            roi: 0
          };
        }

        monthlyMetrics[monthKey].automatedProcesses++;

        const estimatedManualHours = process.ProcessType === ProcessType.Joiner ? 16
          : process.ProcessType === ProcessType.Mover ? 8
          : 12;

        monthlyMetrics[monthKey].manualHoursSaved += estimatedManualHours;
      }

      const metrics: ICostSavings[] = [];
      const keys = Object.keys(monthlyMetrics);

      for (let i = 0; i < keys.length; i++) {
        const key = keys[i];
        const metric = monthlyMetrics[key];
        metric.fteEquivalent = metric.manualHoursSaved / hoursPerFTE;
        metric.costSavings = metric.manualHoursSaved * avgCostPerHour;
        metric.roi = metric.costSavings;
        metrics.push(metric);
      }

      return metrics;
    } catch (error) {
      logger.error('SuccessMetricsService', 'Failed to calculate cost savings:', error);
      return [];
    }
  }

  /**
   * Calculate Compliance Score
   * Percentage of processes following policy
   */
  public async getComplianceScore(filters?: IAnalyticsFilters): Promise<IComplianceMetric[]> {
    try {
      const processes = await this.getFilteredProcesses(filters);
      const monthlyMetrics: { [key: string]: IComplianceMetric } = {};

      for (let i = 0; i < processes.length; i++) {
        const process = processes[i];
        const startDate = process.StartDate
          ? (typeof process.StartDate === 'string' ? new Date(process.StartDate) : process.StartDate)
          : new Date();
        const monthKey = `${startDate.getFullYear()}-${startDate.getMonth() + 1}`;

        if (!monthlyMetrics[monthKey]) {
          monthlyMetrics[monthKey] = {
            period: new Date(startDate.getFullYear(), startDate.getMonth(), 1),
            totalProcesses: 0,
            compliantProcesses: 0,
            complianceRate: 0,
            criticalViolations: 0,
            minorViolations: 0,
            complianceCategories: []
          };
        }

        monthlyMetrics[monthKey].totalProcesses++;

        const completionRate = process.ProgressPercentage || 0;
        const isOnTime = !process.IsOverdue;
        const hasAllTasks = (process.CompletedTasks || 0) >= (process.TotalTasks || 0);

        const isCompliant = completionRate >= 90 && isOnTime && hasAllTasks;

        if (isCompliant) {
          monthlyMetrics[monthKey].compliantProcesses++;
        } else {
          if (completionRate < 50 || process.IsOverdue) {
            monthlyMetrics[monthKey].criticalViolations++;
          } else {
            monthlyMetrics[monthKey].minorViolations++;
          }
        }
      }

      const metrics: IComplianceMetric[] = [];
      const keys = Object.keys(monthlyMetrics);

      for (let i = 0; i < keys.length; i++) {
        const key = keys[i];
        const metric = monthlyMetrics[key];
        metric.complianceRate = metric.totalProcesses > 0
          ? (metric.compliantProcesses / metric.totalProcesses) * 100
          : 0;
        metrics.push(metric);
      }

      return metrics;
    } catch (error) {
      logger.error('SuccessMetricsService', 'Failed to calculate compliance score:', error);
      return [];
    }
  }

  /**
   * Get overall success metrics summary
   */
  public async getSuccessMetricsSummary(filters?: IAnalyticsFilters): Promise<ISuccessMetricsSummary> {
    try {
      const timeToOnboard = await this.getTimeToOnboard(filters);
      const taskCompletion = await this.getTaskCompletionRate(filters);
      const userAdoption = await this.getUserAdoption(filters);
      const cycleTime = await this.getProcessCycleTime(filters);
      const errorRate = await this.getErrorRate(filters);
      const satisfaction = await this.getUserSatisfaction(filters);
      const costSavings = await this.getCostSavings(filters);
      const compliance = await this.getComplianceScore(filters);

      const avgTimeToOnboard = timeToOnboard.length > 0
        ? timeToOnboard.reduce((sum, m) => sum + m.daysToOnboard, 0) / timeToOnboard.length
        : 0;

      const avgTaskCompletionRate = taskCompletion.length > 0
        ? taskCompletion.reduce((sum, m) => sum + m.onTimeRate, 0) / taskCompletion.length
        : 0;

      const latestUserAdoption = userAdoption.length > 0 ? userAdoption[userAdoption.length - 1] : null;

      const avgCycleTime = cycleTime.length > 0
        ? cycleTime.reduce((sum, m) => sum + m.averageDays, 0) / cycleTime.length
        : 0;

      const avgErrorRate = errorRate.length > 0
        ? errorRate.reduce((sum, m) => sum + m.processErrorRate, 0) / errorRate.length
        : 0;

      const latestSatisfaction = satisfaction.length > 0 ? satisfaction[satisfaction.length - 1] : null;

      const totalCostSavings = costSavings.reduce((sum, m) => sum + m.costSavings, 0);
      const totalFTE = costSavings.reduce((sum, m) => sum + m.fteEquivalent, 0);

      const avgComplianceRate = compliance.length > 0
        ? compliance.reduce((sum, m) => sum + m.complianceRate, 0) / compliance.length
        : 0;

      return {
        period: new Date(),
        timeToOnboard: {
          average: avgTimeToOnboard,
          target: 5,
          variance: avgTimeToOnboard - 5,
          trend: this.calculateTrend(timeToOnboard.map(m => m.daysToOnboard))
        },
        taskCompletionRate: {
          rate: avgTaskCompletionRate,
          target: 90,
          trend: this.calculateTrend(taskCompletion.map(m => m.onTimeRate))
        },
        userAdoption: {
          activeUsers: latestUserAdoption?.activeUsers || 0,
          adoptionRate: latestUserAdoption?.adoptionRate || 0,
          trend: this.calculateTrend(userAdoption.map(m => m.adoptionRate))
        },
        processCycleTime: {
          average: avgCycleTime,
          target: 5,
          variance: avgCycleTime - 5,
          trend: this.calculateTrend(cycleTime.map(m => m.averageDays))
        },
        errorRate: {
          rate: avgErrorRate,
          target: 5,
          trend: this.calculateTrend(errorRate.map(m => m.processErrorRate))
        },
        userSatisfaction: {
          npsScore: latestSatisfaction?.npsScore || 0,
          csatScore: latestSatisfaction?.csatScore || 0,
          trend: this.calculateTrend(satisfaction.map(m => m.npsScore))
        },
        costSavings: {
          total: totalCostSavings,
          fteEquivalent: totalFTE,
          roi: totalCostSavings
        },
        complianceScore: {
          rate: avgComplianceRate,
          target: 95,
          trend: this.calculateTrend(compliance.map(m => m.complianceRate))
        }
      };
    } catch (error) {
      logger.error('SuccessMetricsService', 'Failed to get success metrics summary:', error);
      throw error;
    }
  }

  // Private helper methods

  private calculateTrend(values: number[]): 'improving' | 'declining' | 'stable' {
    if (values.length < 2) {
      return 'stable';
    }

    const recentValues = values.slice(-3);
    const older = recentValues[0];
    const newer = recentValues[recentValues.length - 1];

    const change = newer - older;
    const percentChange = Math.abs((change / Math.max(older, 1)) * 100);

    if (percentChange < 5) {
      return 'stable';
    }

    return change > 0 ? 'improving' : 'declining';
  }

  private async getFilteredProcesses(filters?: IAnalyticsFilters): Promise<IJmlProcess[]> {
    try {
      // Build secure filters
      const filterParts: string[] = [];

      if (filters?.startDate) {
        ValidationUtils.validateDateRange(filters.startDate, filters.endDate || new Date());
        const startFilter = ValidationUtils.buildFilter('StartDate', 'ge', filters.startDate);
        filterParts.push(startFilter);
      }

      if (filters?.endDate) {
        const endFilter = ValidationUtils.buildFilter('StartDate', 'le', filters.endDate);
        filterParts.push(endFilter);
      }

      if (filters?.departments && filters.departments.length > 0) {
        const deptFilters: string[] = [];
        for (let i = 0; i < filters.departments.length; i++) {
          // Sanitize department names
          const sanitizedDept = ValidationUtils.sanitizeForOData(filters.departments[i]);
          deptFilters.push(`Department eq '${sanitizedDept}'`);
        }
        filterParts.push(`(${deptFilters.join(' or ')})`);
      }

      if (filters?.processTypes && filters.processTypes.length > 0) {
        const typeFilters: string[] = [];
        for (let i = 0; i < filters.processTypes.length; i++) {
          ValidationUtils.validateEnum(filters.processTypes[i], ProcessType, 'ProcessType');
          typeFilters.push(ValidationUtils.buildFilter('ProcessType', 'eq', filters.processTypes[i]));
        }
        filterParts.push(`(${typeFilters.join(' or ')})`);
      }

      const filterQuery = filterParts.length > 0 ? filterParts.join(' and ') : undefined;

      let query = this.sp.web.lists.getByTitle('JML_Processes').items
        .select('*', 'Manager/Title', 'Manager/EMail')
        .expand('Manager')
        .top(5000);

      if (filterQuery) {
        query = query.filter(filterQuery);
      }

      const items = await query();
      return items as IJmlProcess[];
    } catch (error) {
      logger.error('SuccessMetricsService', 'Failed to get filtered processes:', error);
      return [];
    }
  }

  private async getAllTasks(filters?: IAnalyticsFilters): Promise<IJmlTaskAssignment[]> {
    try {
      const items = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
        .select('*', 'Task/Title')
        .expand('Task')
        .top(5000)();
      return items as IJmlTaskAssignment[];
    } catch (error) {
      logger.error('SuccessMetricsService', 'Failed to get tasks:', error);
      return [];
    }
  }

  private async getAuditLogs(filters?: IAnalyticsFilters): Promise<IJmlAuditLog[]> {
    try {
      // Build secure filters
      const filterParts: string[] = [];

      if (filters?.startDate) {
        if (filters.endDate) {
          ValidationUtils.validateDateRange(filters.startDate, filters.endDate);
        }
        const startFilter = ValidationUtils.buildFilter('Timestamp', 'ge', filters.startDate);
        filterParts.push(startFilter);
      }

      if (filters?.endDate) {
        const endFilter = ValidationUtils.buildFilter('Timestamp', 'le', filters.endDate);
        filterParts.push(endFilter);
      }

      const filterQuery = filterParts.length > 0 ? filterParts.join(' and ') : undefined;

      let query = this.sp.web.lists.getByTitle('JML_AuditLog').items
        .select('*')
        .top(5000);

      if (filterQuery) {
        query = query.filter(filterQuery);
      }

      const items = await query();
      return items as IJmlAuditLog[];
    } catch (error) {
      logger.error('SuccessMetricsService', 'Failed to get audit logs:', error);
      return [];
    }
  }

  private async getTotalUsers(): Promise<number> {
    try {
      const users = await this.sp.web.siteUsers();
      return users.length;
    } catch (error) {
      logger.error('SuccessMetricsService', 'Failed to get total users:', error);
      return 0;
    }
  }

  private async getSatisfactionSurveys(filters?: IAnalyticsFilters): Promise<any[]> {
    try {
      const items = await this.sp.web.lists.getByTitle('JML_Satisfaction').items
        .select('*')
        .top(5000)();
      return items;
    } catch (error) {
      logger.error('SuccessMetricsService', 'Failed to get satisfaction surveys:', error);
      return [];
    }
  }
}
