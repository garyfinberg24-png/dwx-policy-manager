// @ts-nocheck
// Process Analytics Service
// Provides analytics and metrics for JML process creation and execution

import { SPFI } from '@pnp/sp';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/views';
import { IJmlProcess } from '../models';
import { ProcessStatus, ProcessType } from '../models/ICommon';
import { logger } from './LoggingService';

/**
 * Process creation metrics by time period
 */
export interface IProcessCreationMetrics {
  totalProcesses: number;
  joiners: number;
  movers: number;
  leavers: number;
  averageCompletionTime: number; // days
  onTimeCompletion: number; // percentage
}

/**
 * Process statistics by process type
 */
export interface IProcessTypeStats {
  processType: ProcessType;
  total: number;
  active: number;
  completed: number;
  overdue: number;
  averageDuration: number; // days
  completionRate: number; // percentage
}

/**
 * Process trends over time
 */
export interface IProcessTrend {
  date: Date;
  joiners: number;
  movers: number;
  leavers: number;
  total: number;
}

/**
 * Department-specific process metrics
 */
export interface IDepartmentMetrics {
  department: string;
  totalProcesses: number;
  activeProcesses: number;
  completedProcesses: number;
  averageCompletionDays: number;
  onTimeRate: number; // percentage
}

/**
 * Time-based analytics filters
 */
export interface IAnalyticsTimeFilter {
  startDate?: Date;
  endDate?: Date;
  period?: 'week' | 'month' | 'quarter' | 'year' | 'all';
}

/**
 * Process status breakdown
 */
export interface IStatusBreakdown {
  status: ProcessStatus;
  count: number;
  percentage: number;
}

export class ProcessAnalyticsService {
  private sp: SPFI;
  private cacheTimeout: number = 5 * 60 * 1000; // 5 minutes
  private cache: Map<string, { data: any; timestamp: number }> = new Map();

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Get overall process creation metrics
   */
  public async getProcessCreationMetrics(filter?: IAnalyticsTimeFilter): Promise<IProcessCreationMetrics> {
    try {
      const cacheKey = `creation-metrics-${JSON.stringify(filter)}`;
      const cached = this.getFromCache<IProcessCreationMetrics>(cacheKey);
      if (cached) return cached;

      const processes = await this.getProcessesForPeriod(filter);

      const metrics: IProcessCreationMetrics = {
        totalProcesses: processes.length,
        joiners: processes.filter(p => p.ProcessType === ProcessType.Joiner).length,
        movers: processes.filter(p => p.ProcessType === ProcessType.Mover).length,
        leavers: processes.filter(p => p.ProcessType === ProcessType.Leaver).length,
        averageCompletionTime: this.calculateAverageCompletionTime(processes),
        onTimeCompletion: this.calculateOnTimeRate(processes)
      };

      this.setCache(cacheKey, metrics);
      return metrics;
    } catch (error) {
      logger.error('ProcessAnalyticsService', 'Failed to get creation metrics:', error);
      return {
        totalProcesses: 0,
        joiners: 0,
        movers: 0,
        leavers: 0,
        averageCompletionTime: 0,
        onTimeCompletion: 0
      };
    }
  }

  /**
   * Get statistics by process type
   */
  public async getProcessTypeStatistics(filter?: IAnalyticsTimeFilter): Promise<IProcessTypeStats[]> {
    try {
      const cacheKey = `type-stats-${JSON.stringify(filter)}`;
      const cached = this.getFromCache<IProcessTypeStats[]>(cacheKey);
      if (cached) return cached;

      const processes = await this.getProcessesForPeriod(filter);
      const types = [ProcessType.Joiner, ProcessType.Mover, ProcessType.Leaver];

      const stats = types.map(type => {
        const typeProcesses = processes.filter(p => p.ProcessType === type);
        const completed = typeProcesses.filter(p => p.ProcessStatus === ProcessStatus.Completed);
        const active = typeProcesses.filter(p =>
          p.ProcessStatus === ProcessStatus.InProgress ||
          p.ProcessStatus === ProcessStatus.Draft ||
          p.ProcessStatus === ProcessStatus.Pending
        );
        const overdue = typeProcesses.filter(p =>
          p.TargetCompletionDate &&
          new Date(p.TargetCompletionDate) < new Date() &&
          p.ProcessStatus !== ProcessStatus.Completed
        );

        return {
          processType: type,
          total: typeProcesses.length,
          active: active.length,
          completed: completed.length,
          overdue: overdue.length,
          averageDuration: this.calculateAverageCompletionTime(typeProcesses),
          completionRate: typeProcesses.length > 0 ? (completed.length / typeProcesses.length) * 100 : 0
        };
      });

      this.setCache(cacheKey, stats);
      return stats;
    } catch (error) {
      logger.error('ProcessAnalyticsService', 'Failed to get type statistics:', error);
      return [];
    }
  }

  /**
   * Get process creation trends over time
   */
  public async getProcessTrends(filter?: IAnalyticsTimeFilter): Promise<IProcessTrend[]> {
    try {
      const cacheKey = `trends-${JSON.stringify(filter)}`;
      const cached = this.getFromCache<IProcessTrend[]>(cacheKey);
      if (cached) return cached;

      const processes = await this.getProcessesForPeriod(filter);
      const trendsMap: Map<string, IProcessTrend> = new Map();

      processes.forEach(process => {
        if (!process.Created) return;

        const date = new Date(process.Created);
        const month = date.getMonth() + 1;
        const monthKey = `${date.getFullYear()}-${month < 10 ? '0' + month : month}`;

        if (!trendsMap.has(monthKey)) {
          trendsMap.set(monthKey, {
            date: new Date(date.getFullYear(), date.getMonth(), 1),
            joiners: 0,
            movers: 0,
            leavers: 0,
            total: 0
          });
        }

        const trend = trendsMap.get(monthKey)!;
        trend.total++;
        if (process.ProcessType === ProcessType.Joiner) trend.joiners++;
        if (process.ProcessType === ProcessType.Mover) trend.movers++;
        if (process.ProcessType === ProcessType.Leaver) trend.leavers++;
      });

      const trends = Array.from(trendsMap.values()).sort((a, b) => a.date.getTime() - b.date.getTime());

      this.setCache(cacheKey, trends);
      return trends;
    } catch (error) {
      logger.error('ProcessAnalyticsService', 'Failed to get process trends:', error);
      return [];
    }
  }

  /**
   * Get metrics by department
   */
  public async getDepartmentMetrics(filter?: IAnalyticsTimeFilter): Promise<IDepartmentMetrics[]> {
    try {
      const cacheKey = `dept-metrics-${JSON.stringify(filter)}`;
      const cached = this.getFromCache<IDepartmentMetrics[]>(cacheKey);
      if (cached) return cached;

      const processes = await this.getProcessesForPeriod(filter);
      const deptMap: Map<string, IJmlProcess[]> = new Map();

      processes.forEach(process => {
        if (!process.Department) return;
        if (!deptMap.has(process.Department)) {
          deptMap.set(process.Department, []);
        }
        deptMap.get(process.Department)!.push(process);
      });

      const metrics = Array.from(deptMap.entries()).map(([department, deptProcesses]) => {
        const completed = deptProcesses.filter(p => p.ProcessStatus === ProcessStatus.Completed);
        const active = deptProcesses.filter(p =>
          p.ProcessStatus === ProcessStatus.InProgress ||
          p.ProcessStatus === ProcessStatus.Draft ||
          p.ProcessStatus === ProcessStatus.Pending
        );

        return {
          department,
          totalProcesses: deptProcesses.length,
          activeProcesses: active.length,
          completedProcesses: completed.length,
          averageCompletionDays: this.calculateAverageCompletionTime(completed),
          onTimeRate: this.calculateOnTimeRate(completed)
        };
      });

      metrics.sort((a, b) => b.totalProcesses - a.totalProcesses);

      this.setCache(cacheKey, metrics);
      return metrics;
    } catch (error) {
      logger.error('ProcessAnalyticsService', 'Failed to get department metrics:', error);
      return [];
    }
  }

  /**
   * Get status breakdown for all processes
   */
  public async getStatusBreakdown(filter?: IAnalyticsTimeFilter): Promise<IStatusBreakdown[]> {
    try {
      const processes = await this.getProcessesForPeriod(filter);
      const statusMap: Map<ProcessStatus, number> = new Map();

      processes.forEach(process => {
        const status = process.ProcessStatus;
        statusMap.set(status, (statusMap.get(status) || 0) + 1);
      });

      const total = processes.length;
      const breakdown = Array.from(statusMap.entries()).map(([status, count]) => ({
        status,
        count,
        percentage: total > 0 ? (count / total) * 100 : 0
      }));

      return breakdown.sort((a, b) => b.count - a.count);
    } catch (error) {
      logger.error('ProcessAnalyticsService', 'Failed to get status breakdown:', error);
      return [];
    }
  }

  /**
   * Get processes for a specific time period
   */
  private async getProcessesForPeriod(filter?: IAnalyticsTimeFilter): Promise<IJmlProcess[]> {
    try {
      let query = this.sp.web.lists.getByTitle('JML_Processes_Test').items
        .select(
          'Id', 'Title', 'ProcessType', 'ProcessStatus', 'Department', 'Priority',
          'StartDate', 'TargetCompletionDate', 'ActualCompletionDate',
          'ProgressPercentage', 'TotalTasks', 'CompletedTasks', 'OverdueTasks',
          'Created', 'Modified'
        )
        .top(5000);

      if (filter?.startDate) {
        query = query.filter(`Created ge datetime'${filter.startDate.toISOString()}'`);
      }

      if (filter?.endDate) {
        query = query.filter(`Created le datetime'${filter.endDate.toISOString()}'`);
      }

      const items = await query();
      return items as IJmlProcess[];
    } catch (error) {
      logger.error('ProcessAnalyticsService', 'Failed to fetch processes:', error);
      return [];
    }
  }

  /**
   * Calculate average completion time in days
   */
  private calculateAverageCompletionTime(processes: IJmlProcess[]): number {
    const completedWithDates = processes.filter(p =>
      p.StartDate &&
      p.ActualCompletionDate &&
      p.ProcessStatus === ProcessStatus.Completed
    );

    if (completedWithDates.length === 0) return 0;

    const totalDays = completedWithDates.reduce((sum, process) => {
      const start = new Date(process.StartDate!);
      const end = new Date(process.ActualCompletionDate!);
      const days = Math.floor((end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24));
      return sum + days;
    }, 0);

    return Math.round(totalDays / completedWithDates.length);
  }

  /**
   * Calculate on-time completion rate
   */
  private calculateOnTimeRate(processes: IJmlProcess[]): number {
    const completed = processes.filter(p =>
      p.ProcessStatus === ProcessStatus.Completed &&
      p.TargetCompletionDate &&
      p.ActualCompletionDate
    );

    if (completed.length === 0) return 0;

    const onTime = completed.filter(p => {
      const target = new Date(p.TargetCompletionDate!);
      const actual = new Date(p.ActualCompletionDate!);
      return actual <= target;
    });

    return Math.round((onTime.length / completed.length) * 100);
  }

  /**
   * Get time filter boundaries
   */
  public getFilterBoundaries(period: 'week' | 'month' | 'quarter' | 'year' | 'all'): IAnalyticsTimeFilter {
    const now = new Date();
    const filter: IAnalyticsTimeFilter = { period };

    switch (period) {
      case 'week':
        filter.startDate = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
        filter.endDate = now;
        break;
      case 'month':
        filter.startDate = new Date(now.getFullYear(), now.getMonth(), 1);
        filter.endDate = now;
        break;
      case 'quarter':
        const quarter = Math.floor(now.getMonth() / 3);
        filter.startDate = new Date(now.getFullYear(), quarter * 3, 1);
        filter.endDate = now;
        break;
      case 'year':
        filter.startDate = new Date(now.getFullYear(), 0, 1);
        filter.endDate = now;
        break;
      case 'all':
        // No date filters
        break;
    }

    return filter;
  }

  /**
   * Cache management
   */
  private getFromCache<T>(key: string): T | null {
    const cached = this.cache.get(key);
    if (!cached) return null;

    const isExpired = Date.now() - cached.timestamp > this.cacheTimeout;
    if (isExpired) {
      this.cache.delete(key);
      return null;
    }

    return cached.data as T;
  }

  private setCache(key: string, data: any): void {
    this.cache.set(key, {
      data,
      timestamp: Date.now()
    });
  }

  /**
   * Clear all cached data
   */
  public clearCache(): void {
    this.cache.clear();
  }
}
