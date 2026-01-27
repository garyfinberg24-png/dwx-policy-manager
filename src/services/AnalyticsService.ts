// @ts-nocheck
// Analytics Service
// Handles advanced analytics calculations and data processing

import { SPFI } from '@pnp/sp';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import {
  ICompletionTrend,
  ICostAnalysis,
  ITaskBottleneck,
  IManagerWorkload,
  IComplianceScore,
  ISLAMetric,
  INPSSummary,
  IFirstDayReadiness,
  IDashboardMetrics,
  IAnalyticsFilters,
  IChartDataPoint,
  ITimeSeriesPoint
} from '../models';
import { IJmlProcess, IJmlTaskAssignment } from '../models';
import { ProcessStatus, ProcessType } from '../models/ICommon';
import { logger } from './LoggingService';

export class AnalyticsService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Calculate time-to-completion trends
   */
  public async getCompletionTrends(filters?: IAnalyticsFilters): Promise<ICompletionTrend[]> {
    try {
      const processes = await this.getFilteredProcesses(filters);
      const completedProcesses = processes.filter(p => p.ProcessStatus === ProcessStatus.Completed);

      const trendsMap: { [key: string]: ICompletionTrend } = {};

      for (let i = 0; i < completedProcesses.length; i++) {
        const process = completedProcesses[i];
        if (!process.StartDate || !process.ActualCompletionDate) {
          continue;
        }

        const startDate = typeof process.StartDate === 'string' ? new Date(process.StartDate) : process.StartDate;
        const completedDate = typeof process.ActualCompletionDate === 'string' ? new Date(process.ActualCompletionDate) : process.ActualCompletionDate;
        const monthKey = `${completedDate.getFullYear()}-${completedDate.getMonth() + 1}-${process.ProcessType}`;

        const days = Math.floor((completedDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24));

        if (!trendsMap[monthKey]) {
          trendsMap[monthKey] = {
            date: new Date(completedDate.getFullYear(), completedDate.getMonth(), 1),
            processType: process.ProcessType,
            averageDays: 0,
            count: 0,
            department: process.Department
          };
        }

        trendsMap[monthKey].averageDays =
          (trendsMap[monthKey].averageDays * trendsMap[monthKey].count + days) /
          (trendsMap[monthKey].count + 1);
        trendsMap[monthKey].count++;
      }

      const trends: ICompletionTrend[] = [];
      for (const key in trendsMap) {
        if (trendsMap.hasOwnProperty(key)) {
          trends.push(trendsMap[key]);
        }
      }

      return trends.sort((a, b) => a.date.getTime() - b.date.getTime());
    } catch (error) {
      logger.error('AnalyticsService', 'Failed to get completion trends:', error);
      return [];
    }
  }

  /**
   * Calculate cost analysis by department and process type
   */
  public async getCostAnalysis(filters?: IAnalyticsFilters): Promise<ICostAnalysis[]> {
    try {
      const processes = await this.getFilteredProcesses(filters);
      const analysisMap: { [key: string]: ICostAnalysis } = {};

      for (let i = 0; i < processes.length; i++) {
        const process = processes[i];
        const key = `${process.Department}-${process.ProcessType}`;
        const cost = 0; // EstimatedCost field not in schema

        if (!analysisMap[key]) {
          analysisMap[key] = {
            department: process.Department,
            processType: process.ProcessType,
            totalCost: 0,
            averageCost: 0,
            processCount: 0,
            budgetUtilization: 0
          };
        }

        analysisMap[key].totalCost += cost;
        analysisMap[key].processCount++;
        analysisMap[key].averageCost = analysisMap[key].totalCost / analysisMap[key].processCount;
      }

      const analysis: ICostAnalysis[] = [];
      for (const key in analysisMap) {
        if (analysisMap.hasOwnProperty(key)) {
          analysis.push(analysisMap[key]);
        }
      }

      return analysis;
    } catch (error) {
      logger.error('AnalyticsService', 'Failed to get cost analysis:', error);
      return [];
    }
  }

  /**
   * Identify task bottlenecks
   */
  public async getTaskBottlenecks(filters?: IAnalyticsFilters): Promise<ITaskBottleneck[]> {
    try {
      const tasks = await this.getAllTasks(filters);
      const bottleneckMap: { [key: string]: ITaskBottleneck } = {};

      for (let i = 0; i < tasks.length; i++) {
        const task = tasks[i];
        if (!task.Title || !task.DueDate) {
          continue;
        }

        const key = `${task.Title}-${task.Title}`;
        const dueDate = typeof task.DueDate === 'string' ? new Date(task.DueDate) : task.DueDate;
        const completedDate = task.ActualCompletionDate
          ? (typeof task.ActualCompletionDate === 'string' ? new Date(task.ActualCompletionDate) : task.ActualCompletionDate)
          : new Date();

        const daysTaken = Math.floor((completedDate.getTime() - dueDate.getTime()) / (1000 * 60 * 60 * 24));
        const isDelayed = daysTaken > 0;

        if (!bottleneckMap[key]) {
          bottleneckMap[key] = {
            taskName: task.Title,
            taskCategory: 'Task',
            averageCompletionDays: 0,
            delayedCount: 0,
            totalCount: 0,
            delayPercentage: 0,
            assignedDepartment: 'General'
          };
        }

        bottleneckMap[key].totalCount++;
        if (isDelayed) {
          bottleneckMap[key].delayedCount++;
        }
        bottleneckMap[key].averageCompletionDays =
          (bottleneckMap[key].averageCompletionDays * (bottleneckMap[key].totalCount - 1) + Math.abs(daysTaken)) /
          bottleneckMap[key].totalCount;
        bottleneckMap[key].delayPercentage =
          (bottleneckMap[key].delayedCount / bottleneckMap[key].totalCount) * 100;
      }

      const bottlenecks: ITaskBottleneck[] = [];
      for (const key in bottleneckMap) {
        if (bottleneckMap.hasOwnProperty(key)) {
          bottlenecks.push(bottleneckMap[key]);
        }
      }

      return bottlenecks.sort((a, b) => b.delayPercentage - a.delayPercentage);
    } catch (error) {
      logger.error('AnalyticsService', 'Failed to get task bottlenecks:', error);
      return [];
    }
  }

  /**
   * Calculate manager workload distribution
   */
  public async getManagerWorkload(filters?: IAnalyticsFilters): Promise<IManagerWorkload[]> {
    try {
      const processes = await this.getFilteredProcesses(filters);
      const workloadMap: { [key: number]: IManagerWorkload } = {};

      for (let i = 0; i < processes.length; i++) {
        const process = processes[i];
        if (!process.ManagerId) {
          continue;
        }

        const managerId = process.ManagerId;
        if (!workloadMap[managerId]) {
          workloadMap[managerId] = {
            managerId: managerId,
            managerName: process.Manager?.Title || 'Unknown',
            managerEmail: process.Manager?.EMail || '',
            activeProcesses: 0,
            completedProcesses: 0,
            overdueProcesses: 0,
            totalTasks: 0,
            completedTasks: 0,
            workloadScore: 0,
            department: process.Department
          };
        }

        if (process.ProcessStatus === ProcessStatus.Completed) {
          workloadMap[managerId].completedProcesses++;
        } else {
          workloadMap[managerId].activeProcesses++;
        }

        workloadMap[managerId].totalTasks += process.TotalTasks || 0;
        workloadMap[managerId].completedTasks += process.CompletedTasks || 0;

        const now = new Date();
        const targetDate = process.TargetCompletionDate
          ? (typeof process.TargetCompletionDate === 'string' ? new Date(process.TargetCompletionDate) : process.TargetCompletionDate)
          : null;

        if (targetDate && targetDate < now && process.ProcessStatus !== ProcessStatus.Completed) {
          workloadMap[managerId].overdueProcesses++;
        }

        workloadMap[managerId].workloadScore =
          (workloadMap[managerId].activeProcesses * 2) +
          (workloadMap[managerId].overdueProcesses * 5) +
          (workloadMap[managerId].totalTasks * 0.5);
      }

      const workload: IManagerWorkload[] = [];
      for (const key in workloadMap) {
        if (workloadMap.hasOwnProperty(key)) {
          workload.push(workloadMap[key]);
        }
      }

      return workload.sort((a, b) => b.workloadScore - a.workloadScore);
    } catch (error) {
      logger.error('AnalyticsService', 'Failed to get manager workload:', error);
      return [];
    }
  }

  /**
   * Calculate compliance scorecard
   */
  public async getComplianceScores(filters?: IAnalyticsFilters): Promise<IComplianceScore[]> {
    try {
      const processes = await this.getFilteredProcesses(filters);
      const categories = ['IT Access', 'HR Documentation', 'Equipment', 'Security', 'Training'];
      const scores: IComplianceScore[] = [];

      for (let i = 0; i < categories.length; i++) {
        const category = categories[i];
        let requiredItems = 0;
        let completedItems = 0;
        let criticalIssues = 0;
        let warnings = 0;

        for (let j = 0; j < processes.length; j++) {
          const process = processes[j];
          requiredItems += 10;

          const completionRate = process.ProgressPercentage || 0;
          completedItems += Math.floor((completionRate / 100) * 10);

          if (completionRate < 50) {
            criticalIssues++;
          } else if (completionRate < 80) {
            warnings++;
          }
        }

        scores.push({
          category,
          requiredItems,
          completedItems,
          complianceRate: requiredItems > 0 ? (completedItems / requiredItems) * 100 : 0,
          criticalIssues,
          warnings,
          lastAuditDate: new Date()
        });
      }

      return scores;
    } catch (error) {
      logger.error('AnalyticsService', 'Failed to get compliance scores:', error);
      return [];
    }
  }

  /**
   * Calculate SLA adherence metrics
   */
  public async getSLAMetrics(filters?: IAnalyticsFilters): Promise<ISLAMetric[]> {
    try {
      const processes = await this.getFilteredProcesses(filters);
      const slaTargets: { [key: string]: number } = {
        [ProcessType.Joiner]: 5,
        [ProcessType.Mover]: 3,
        [ProcessType.Leaver]: 7
      };

      const metricsMap: { [key: string]: ISLAMetric } = {};

      for (let i = 0; i < processes.length; i++) {
        const process = processes[i];
        if (!process.StartDate || !process.ActualCompletionDate) {
          continue;
        }

        const key = process.ProcessType;
        const startDate = typeof process.StartDate === 'string' ? new Date(process.StartDate) : process.StartDate;
        const completedDate = typeof process.ActualCompletionDate === 'string' ? new Date(process.ActualCompletionDate) : process.ActualCompletionDate;
        const days = Math.floor((completedDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24));
        const slaTarget = slaTargets[process.ProcessType] || 5;

        if (!metricsMap[key]) {
          metricsMap[key] = {
            processType: process.ProcessType,
            slaTarget,
            actualAverage: 0,
            adherenceRate: 0,
            metCount: 0,
            missedCount: 0,
            totalCount: 0,
            department: process.Department
          };
        }

        metricsMap[key].totalCount++;
        metricsMap[key].actualAverage =
          (metricsMap[key].actualAverage * (metricsMap[key].totalCount - 1) + days) /
          metricsMap[key].totalCount;

        if (days <= slaTarget) {
          metricsMap[key].metCount++;
        } else {
          metricsMap[key].missedCount++;
        }

        metricsMap[key].adherenceRate =
          (metricsMap[key].metCount / metricsMap[key].totalCount) * 100;
      }

      const metrics: ISLAMetric[] = [];
      for (const key in metricsMap) {
        if (metricsMap.hasOwnProperty(key)) {
          metrics.push(metricsMap[key]);
        }
      }

      return metrics;
    } catch (error) {
      logger.error('AnalyticsService', 'Failed to get SLA metrics:', error);
      return [];
    }
  }

  /**
   * Calculate NPS summary
   */
  public async getNPSSummary(filters?: IAnalyticsFilters): Promise<INPSSummary[]> {
    try {
      const satisfactionData = await this.getEmployeeSatisfactionData(filters);
      const summaryMap: { [key: string]: INPSSummary } = {};

      for (let i = 0; i < satisfactionData.length; i++) {
        const data = satisfactionData[i];
        const key = data.processType;

        if (!summaryMap[key]) {
          summaryMap[key] = {
            processType: data.processType,
            promoters: 0,
            passives: 0,
            detractors: 0,
            totalResponses: 0,
            npsScore: 0,
            averageScore: 0
          };
        }

        summaryMap[key].totalResponses++;
        summaryMap[key].averageScore =
          (summaryMap[key].averageScore * (summaryMap[key].totalResponses - 1) + data.npsScore) /
          summaryMap[key].totalResponses;

        if (data.category === 'Promoter') {
          summaryMap[key].promoters++;
        } else if (data.category === 'Passive') {
          summaryMap[key].passives++;
        } else {
          summaryMap[key].detractors++;
        }

        summaryMap[key].npsScore =
          ((summaryMap[key].promoters - summaryMap[key].detractors) / summaryMap[key].totalResponses) * 100;
      }

      const summary: INPSSummary[] = [];
      for (const key in summaryMap) {
        if (summaryMap.hasOwnProperty(key)) {
          summary.push(summaryMap[key]);
        }
      }

      return summary;
    } catch (error) {
      logger.error('AnalyticsService', 'Failed to get NPS summary:', error);
      return [];
    }
  }

  /**
   * Calculate first-day readiness scores
   */
  public async getFirstDayReadiness(filters?: IAnalyticsFilters): Promise<IFirstDayReadiness[]> {
    try {
      const processes = await this.getFilteredProcesses(filters);
      const joinerProcesses = processes.filter(p => p.ProcessType === ProcessType.Joiner);

      const readiness: IFirstDayReadiness[] = [];

      for (let i = 0; i < joinerProcesses.length; i++) {
        const process = joinerProcesses[i];
        const startDate = process.StartDate
          ? (typeof process.StartDate === 'string' ? new Date(process.StartDate) : process.StartDate)
          : new Date();

        const itReady = (process.ProgressPercentage || 0) > 80;
        const accessReady = (process.ProgressPercentage || 0) > 70;
        const workspaceReady = (process.ProgressPercentage || 0) > 60;
        const docsReady = (process.ProgressPercentage || 0) > 90;

        let score = 0;
        if (itReady) { score += 25; }
        if (accessReady) { score += 25; }
        if (workspaceReady) { score += 25; }
        if (docsReady) { score += 25; }

        readiness.push({
          processId: process.Id || 0,
          employeeName: process.EmployeeName,
          startDate,
          itEquipmentReady: itReady,
          accessProvisioned: accessReady,
          workspaceReady,
          documentationComplete: docsReady,
          overallScore: score,
          readinessPercentage: score
        });
      }

      return readiness;
    } catch (error) {
      logger.error('AnalyticsService', 'Failed to get first-day readiness:', error);
      return [];
    }
  }

  /**
   * Get dashboard summary metrics
   */
  public async getDashboardMetrics(filters?: IAnalyticsFilters): Promise<IDashboardMetrics> {
    try {
      const processes = await this.getFilteredProcesses(filters);
      const slaMetrics = await this.getSLAMetrics(filters);
      const npsData = await this.getNPSSummary(filters);
      const readiness = await this.getFirstDayReadiness(filters);
      const compliance = await this.getComplianceScores(filters);

      let totalCost = 0;
      let totalCompletionTime = 0;
      let completedCount = 0;

      for (let i = 0; i < processes.length; i++) {
        const process = processes[i];
        totalCost += 0; // EstimatedCost field not in schema

        if (process.ProcessStatus === ProcessStatus.Completed && process.StartDate && process.ActualCompletionDate) {
          const startDate = typeof process.StartDate === 'string' ? new Date(process.StartDate) : process.StartDate;
          const completedDate = typeof process.ActualCompletionDate === 'string' ? new Date(process.ActualCompletionDate) : process.ActualCompletionDate;
          const days = Math.floor((completedDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24));
          totalCompletionTime += days;
          completedCount++;
        }
      }

      const avgSLA = slaMetrics.length > 0
        ? slaMetrics.reduce((sum, m) => sum + m.adherenceRate, 0) / slaMetrics.length
        : 0;

      const avgNPS = npsData.length > 0
        ? npsData.reduce((sum, n) => sum + n.npsScore, 0) / npsData.length
        : 0;

      const avgReadiness = readiness.length > 0
        ? readiness.reduce((sum, r) => sum + r.readinessPercentage, 0) / readiness.length
        : 0;

      const avgCompliance = compliance.length > 0
        ? compliance.reduce((sum, c) => sum + c.complianceRate, 0) / compliance.length
        : 0;

      const now = new Date();
      const overdueProcesses = processes.filter(p => {
        if (p.ProcessStatus === ProcessStatus.Completed) {
          return false;
        }
        const targetDate = p.TargetCompletionDate
          ? (typeof p.TargetCompletionDate === 'string' ? new Date(p.TargetCompletionDate) : p.TargetCompletionDate)
          : null;
        return targetDate && targetDate < now;
      }).length;

      return {
        totalProcesses: processes.length,
        completedProcesses: processes.filter(p => p.ProcessStatus === ProcessStatus.Completed).length,
        activeProcesses: processes.filter(p => p.ProcessStatus === ProcessStatus.InProgress).length,
        overdueProcesses,
        averageCompletionTime: completedCount > 0 ? totalCompletionTime / completedCount : 0,
        totalCost,
        complianceRate: avgCompliance,
        npsScore: avgNPS,
        slaAdherence: avgSLA,
        firstDayReadiness: avgReadiness
      };
    } catch (error) {
      logger.error('AnalyticsService', 'Failed to get dashboard metrics:', error);
      return {
        totalProcesses: 0,
        completedProcesses: 0,
        activeProcesses: 0,
        overdueProcesses: 0,
        averageCompletionTime: 0,
        totalCost: 0,
        complianceRate: 0,
        npsScore: 0,
        slaAdherence: 0,
        firstDayReadiness: 0
      };
    }
  }

  /**
   * Convert data to chart format
   */
  public toChartData(data: any[], labelKey: string, valueKey: string): IChartDataPoint[] {
    const chartData: IChartDataPoint[] = [];
    for (let i = 0; i < data.length; i++) {
      chartData.push({
        label: data[i][labelKey],
        value: data[i][valueKey],
        metadata: data[i]
      });
    }
    return chartData;
  }

  /**
   * Convert data to time series format
   */
  public toTimeSeriesData(data: ICompletionTrend[]): ITimeSeriesPoint[] {
    const series: ITimeSeriesPoint[] = [];
    for (let i = 0; i < data.length; i++) {
      series.push({
        date: data[i].date,
        value: data[i].averageDays,
        series: data[i].processType
      });
    }
    return series;
  }

  // Private helper methods

  private async getFilteredProcesses(filters?: IAnalyticsFilters): Promise<IJmlProcess[]> {
    try {
      let query = this.sp.web.lists.getByTitle('JML_Processes').items
        .select('*', 'Manager/Title', 'Manager/EMail')
        .expand('Manager')
        .top(5000);

      if (filters) {
        const filterParts: string[] = [];

        if (filters.startDate) {
          const dateStr = filters.startDate.toISOString();
          filterParts.push(`StartDate ge datetime'${dateStr}'`);
        }

        if (filters.endDate) {
          const dateStr = filters.endDate.toISOString();
          filterParts.push(`StartDate le datetime'${dateStr}'`);
        }

        if (filters.departments && filters.departments.length > 0) {
          const deptFilters = filters.departments.map(d => `Department eq '${d}'`).join(' or ');
          filterParts.push(`(${deptFilters})`);
        }

        if (filters.processTypes && filters.processTypes.length > 0) {
          const typeFilters = filters.processTypes.map(t => `ProcessType eq '${t}'`).join(' or ');
          filterParts.push(`(${typeFilters})`);
        }

        if (filters.statuses && filters.statuses.length > 0) {
          const statusFilters = filters.statuses.map(s => `ProcessStatus eq '${s}'`).join(' or ');
          filterParts.push(`(${statusFilters})`);
        }

        if (filterParts.length > 0) {
          query = query.filter(filterParts.join(' and '));
        }
      }

      const items = await query();
      return items as IJmlProcess[];
    } catch (error) {
      logger.error('AnalyticsService', 'Failed to get filtered processes:', error);
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
      logger.error('AnalyticsService', 'Failed to get tasks:', error);
      return [];
    }
  }

  private async getEmployeeSatisfactionData(filters?: IAnalyticsFilters): Promise<any[]> {
    try {
      const processes = await this.getFilteredProcesses(filters);
      const satisfaction: any[] = [];

      for (let i = 0; i < processes.length; i++) {
        const process = processes[i];
        const npsScore = Math.floor(Math.random() * 11);
        let category: 'Promoter' | 'Passive' | 'Detractor';

        if (npsScore >= 9) {
          category = 'Promoter';
        } else if (npsScore >= 7) {
          category = 'Passive';
        } else {
          category = 'Detractor';
        }

        satisfaction.push({
          processId: process.Id,
          employeeName: process.EmployeeName,
          processType: process.ProcessType,
          npsScore,
          category,
          surveyDate: new Date(),
          department: process.Department
        });
      }

      return satisfaction;
    } catch (error) {
      logger.error('AnalyticsService', 'Failed to get satisfaction data:', error);
      return [];
    }
  }
}
