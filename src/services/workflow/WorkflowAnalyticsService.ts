// @ts-nocheck
/* eslint-disable */
/**
 * WorkflowAnalyticsService
 * Phase 4: Automation & Intelligence
 *
 * Provides predictive analytics capabilities for workflow management:
 * - Process completion time prediction
 * - Bottleneck identification
 * - Resource allocation suggestions
 *
 * @packageDocumentation
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { IWorkflowInstance, WorkflowInstanceStatus } from '../../models/IWorkflow';
import { ProcessType } from '../../models/ICommon';

/**
 * Configuration for the analytics service
 */
export interface IAnalyticsConfig {
  /** Minimum sample size for reliable predictions */
  minSampleSize: number;
  /** Number of days to look back for historical data */
  historicalDaysLookback: number;
  /** Confidence threshold for predictions (0-1) */
  confidenceThreshold: number;
  /** Enable caching of analytics results */
  enableCaching: boolean;
  /** Cache duration in minutes */
  cacheDurationMinutes: number;
}

/**
 * Represents a time prediction with confidence interval
 */
export interface ITimePrediction {
  /** Predicted completion time in hours */
  predictedHours: number;
  /** Predicted completion date */
  predictedCompletionDate: Date;
  /** Lower bound of confidence interval (hours) */
  lowerBoundHours: number;
  /** Upper bound of confidence interval (hours) */
  upperBoundHours: number;
  /** Confidence level (0-1) */
  confidence: number;
  /** Number of historical samples used */
  sampleSize: number;
  /** Factors affecting the prediction */
  factors: IPredictionFactor[];
}

/**
 * Factor affecting a prediction
 */
export interface IPredictionFactor {
  name: string;
  impact: 'positive' | 'negative' | 'neutral';
  weight: number;
  description: string;
}

/**
 * Represents a workflow bottleneck
 */
export interface IBottleneck {
  /** Unique identifier */
  id: string;
  /** Step or stage where bottleneck occurs */
  stepId: string;
  /** Step name */
  stepName: string;
  /** Type of bottleneck */
  type: BottleneckType;
  /** Severity level */
  severity: BottleneckSeverity;
  /** Average delay caused (hours) */
  averageDelayHours: number;
  /** Number of instances affected */
  affectedInstances: number;
  /** Percentage of workflows affected */
  affectedPercentage: number;
  /** Root cause analysis */
  rootCauses: string[];
  /** Suggested resolutions */
  suggestions: string[];
  /** Trend direction */
  trend: 'improving' | 'stable' | 'worsening';
}

/**
 * Types of bottlenecks
 */
export enum BottleneckType {
  ApprovalDelay = 'ApprovalDelay',
  ResourceConstraint = 'ResourceConstraint',
  DependencyBlock = 'DependencyBlock',
  SystemPerformance = 'SystemPerformance',
  ProcessComplexity = 'ProcessComplexity',
  ExternalDependency = 'ExternalDependency'
}

/**
 * Bottleneck severity levels
 */
export enum BottleneckSeverity {
  Low = 'Low',
  Medium = 'Medium',
  High = 'High',
  Critical = 'Critical'
}

/**
 * Resource allocation suggestion
 */
export interface IResourceSuggestion {
  /** Unique identifier */
  id: string;
  /** Type of resource */
  resourceType: ResourceType;
  /** Current allocation */
  currentAllocation: number;
  /** Suggested allocation */
  suggestedAllocation: number;
  /** Change direction */
  direction: 'increase' | 'decrease' | 'redistribute';
  /** Expected impact description */
  expectedImpact: string;
  /** Estimated time savings (hours per week) */
  estimatedTimeSavingsHours: number;
  /** Priority of suggestion */
  priority: SuggestionPriority;
  /** Affected process types */
  affectedProcessTypes: ProcessType[];
  /** Implementation steps */
  implementationSteps: string[];
}

/**
 * Types of resources
 */
export enum ResourceType {
  HRStaff = 'HRStaff',
  ITSupport = 'ITSupport',
  Manager = 'Manager',
  SystemCapacity = 'SystemCapacity',
  ExternalVendor = 'ExternalVendor',
  TrainingResource = 'TrainingResource'
}

/**
 * Priority levels for suggestions
 */
export enum SuggestionPriority {
  Low = 'Low',
  Medium = 'Medium',
  High = 'High',
  Critical = 'Critical'
}

/**
 * Comprehensive analytics report
 */
export interface IAnalyticsReport {
  /** Report generation timestamp */
  generatedAt: Date;
  /** Report period start */
  periodStart: Date;
  /** Report period end */
  periodEnd: Date;
  /** Overall workflow health score (0-100) */
  healthScore: number;
  /** Key performance metrics */
  metrics: IWorkflowMetrics;
  /** Identified bottlenecks */
  bottlenecks: IBottleneck[];
  /** Resource suggestions */
  resourceSuggestions: IResourceSuggestion[];
  /** Process type breakdown */
  processTypeAnalysis: IProcessTypeAnalysis[];
  /** Trends over time */
  trends: ITrendData[];
}

/**
 * Key workflow metrics
 */
export interface IWorkflowMetrics {
  /** Total workflows in period */
  totalWorkflows: number;
  /** Completed workflows */
  completedWorkflows: number;
  /** Active workflows */
  activeWorkflows: number;
  /** Average completion time (hours) */
  averageCompletionTimeHours: number;
  /** Median completion time (hours) */
  medianCompletionTimeHours: number;
  /** On-time completion rate (0-1) */
  onTimeCompletionRate: number;
  /** Average steps per workflow */
  averageStepsPerWorkflow: number;
  /** Approval rate */
  approvalRate: number;
  /** Rejection rate */
  rejectionRate: number;
}

/**
 * Analysis by process type
 */
export interface IProcessTypeAnalysis {
  processType: ProcessType;
  totalInstances: number;
  completedInstances: number;
  averageCompletionHours: number;
  onTimeRate: number;
  topBottleneck: string | null;
  trend: 'improving' | 'stable' | 'declining';
}

/**
 * Trend data point
 */
export interface ITrendData {
  date: Date;
  metric: string;
  value: number;
  trend: 'up' | 'down' | 'stable';
}

/**
 * Historical workflow data for analysis
 */
interface IHistoricalWorkflowData {
  instanceId: string;
  processType: ProcessType;
  startDate: Date;
  completedDate: Date | null;
  status: WorkflowInstanceStatus;
  totalSteps: number;
  completedSteps: number;
  stepDurations: IStepDuration[];
  department?: string;
  complexity: 'low' | 'medium' | 'high';
}

/**
 * Step duration data
 */
interface IStepDuration {
  stepId: string;
  stepName: string;
  startTime: Date;
  endTime: Date | null;
  durationHours: number;
  assignee?: string;
  wasDelayed: boolean;
}

/**
 * Cache entry for analytics data
 */
interface ICacheEntry<T> {
  data: T;
  timestamp: Date;
  expiresAt: Date;
}

/**
 * WorkflowAnalyticsService
 * Provides predictive analytics and insights for workflow management
 */
export class WorkflowAnalyticsService {
  private readonly sp: SPFI;
  private readonly config: IAnalyticsConfig;
  private readonly cache: Map<string, ICacheEntry<unknown>>;

  /** List name for workflow instances */
  private readonly instanceListName = 'JML_WorkflowInstances';
  /** List name for workflow step status */
  private readonly stepStatusListName = 'JML_WorkflowStepStatus';
  /** List name for workflow logs */
  private readonly logsListName = 'JML_WorkflowLogs';

  /**
   * Creates an instance of WorkflowAnalyticsService
   * @param sp - PnPjs SPFI instance
   * @param config - Optional configuration overrides
   */
  constructor(sp: SPFI, config?: Partial<IAnalyticsConfig>) {
    this.sp = sp;
    this.config = {
      minSampleSize: 10,
      historicalDaysLookback: 90,
      confidenceThreshold: 0.7,
      enableCaching: true,
      cacheDurationMinutes: 15,
      ...config
    };
    this.cache = new Map();
  }

  /**
   * Predicts completion time for a workflow instance
   * @param instance - The workflow instance to predict
   * @returns Time prediction with confidence interval
   */
  public async predictCompletionTime(instance: IWorkflowInstance): Promise<ITimePrediction> {
    const cacheKey = `prediction_${instance.Id}`;
    const cached = this.getFromCache<ITimePrediction>(cacheKey);
    if (cached) return cached;

    // Get historical data for similar workflows
    const historicalData = await this.getHistoricalData(instance.ProcessType);

    if (historicalData.length < this.config.minSampleSize) {
      return this.createLowConfidencePrediction(instance, historicalData.length);
    }

    // Calculate base prediction from historical averages
    const completedWorkflows = historicalData.filter(h => h.completedDate !== null);
    const completionTimes = completedWorkflows.map(h =>
      this.calculateDurationHours(h.startDate, h.completedDate!)
    );

    const mean = this.calculateMean(completionTimes);
    const stdDev = this.calculateStandardDeviation(completionTimes, mean);

    // Adjust based on current progress
    const progressFactor = this.calculateProgressFactor(instance);
    const adjustedMean = mean * progressFactor;

    // Calculate factors affecting the prediction
    const factors = await this.analyzePredictionFactors(instance, historicalData);

    // Apply factor adjustments
    const factorAdjustment = factors.reduce((acc, f) => {
      return acc + (f.impact === 'positive' ? -f.weight * 0.1 :
                    f.impact === 'negative' ? f.weight * 0.1 : 0);
    }, 0);

    const finalPrediction = adjustedMean * (1 + factorAdjustment);

    // Calculate confidence based on sample size and variance
    const confidence = this.calculateConfidence(completionTimes.length, stdDev, mean);

    const prediction: ITimePrediction = {
      predictedHours: Math.round(finalPrediction * 10) / 10,
      predictedCompletionDate: new Date(Date.now() + finalPrediction * 60 * 60 * 1000),
      lowerBoundHours: Math.round((finalPrediction - stdDev) * 10) / 10,
      upperBoundHours: Math.round((finalPrediction + stdDev) * 10) / 10,
      confidence: Math.round(confidence * 100) / 100,
      sampleSize: completionTimes.length,
      factors
    };

    this.setCache(cacheKey, prediction);
    return prediction;
  }

  /**
   * Identifies bottlenecks in workflow execution
   * @param processType - Optional filter by process type
   * @returns Array of identified bottlenecks
   */
  public async identifyBottlenecks(processType?: ProcessType): Promise<IBottleneck[]> {
    const cacheKey = `bottlenecks_${processType || 'all'}`;
    const cached = this.getFromCache<IBottleneck[]>(cacheKey);
    if (cached) return cached;

    const historicalData = await this.getHistoricalData(processType);
    const bottlenecks: IBottleneck[] = [];

    // Analyze step durations to find delays
    const stepAnalysis = this.analyzeStepDurations(historicalData);

    for (const [stepId, analysis] of Object.entries(stepAnalysis)) {
      if (analysis.averageDelay > 2) { // More than 2 hours average delay
        const severity = this.calculateBottleneckSeverity(analysis.averageDelay, analysis.frequency);
        const rootCauses = this.identifyRootCauses(analysis);

        bottlenecks.push({
          id: `bottleneck_${stepId}`,
          stepId,
          stepName: analysis.stepName,
          type: this.categorizeBottleneckType(analysis),
          severity,
          averageDelayHours: Math.round(analysis.averageDelay * 10) / 10,
          affectedInstances: analysis.affectedCount,
          affectedPercentage: Math.round((analysis.affectedCount / historicalData.length) * 100),
          rootCauses,
          suggestions: this.generateBottleneckSuggestions(analysis, rootCauses),
          trend: this.calculateBottleneckTrend(analysis)
        });
      }
    }

    // Sort by severity and impact
    bottlenecks.sort((a, b) => {
      const severityOrder = { Critical: 0, High: 1, Medium: 2, Low: 3 };
      const severityDiff = severityOrder[a.severity] - severityOrder[b.severity];
      if (severityDiff !== 0) return severityDiff;
      return b.affectedPercentage - a.affectedPercentage;
    });

    this.setCache(cacheKey, bottlenecks);
    return bottlenecks;
  }

  /**
   * Generates resource allocation suggestions
   * @returns Array of resource suggestions
   */
  public async generateResourceSuggestions(): Promise<IResourceSuggestion[]> {
    const cacheKey = 'resource_suggestions';
    const cached = this.getFromCache<IResourceSuggestion[]>(cacheKey);
    if (cached) return cached;

    const suggestions: IResourceSuggestion[] = [];
    const bottlenecks = await this.identifyBottlenecks();
    const metrics = await this.calculateMetrics();

    // Analyze approval bottlenecks
    const approvalBottlenecks = bottlenecks.filter(b => b.type === BottleneckType.ApprovalDelay);
    if (approvalBottlenecks.length > 0) {
      const totalDelay = approvalBottlenecks.reduce((sum, b) => sum + b.averageDelayHours, 0);
      suggestions.push({
        id: 'suggestion_approval_resources',
        resourceType: ResourceType.Manager,
        currentAllocation: await this.estimateCurrentAllocation(ResourceType.Manager),
        suggestedAllocation: await this.estimateCurrentAllocation(ResourceType.Manager) +
          Math.ceil(totalDelay / 8), // Additional managers based on delay
        direction: 'increase',
        expectedImpact: `Reduce approval delays by an estimated ${Math.round(totalDelay * 0.4)} hours per workflow`,
        estimatedTimeSavingsHours: Math.round(totalDelay * 0.4 * metrics.activeWorkflows / 4),
        priority: totalDelay > 24 ? SuggestionPriority.High : SuggestionPriority.Medium,
        affectedProcessTypes: [ProcessType.Joiner, ProcessType.Mover, ProcessType.Leaver],
        implementationSteps: [
          'Identify high-volume approval queues',
          'Delegate approval authority to additional managers',
          'Implement approval SLAs with escalation',
          'Consider parallel approval paths where appropriate'
        ]
      });
    }

    // Analyze IT support bottlenecks
    const itBottlenecks = bottlenecks.filter(b =>
      b.type === BottleneckType.ResourceConstraint &&
      b.rootCauses.some(c => c.toLowerCase().includes('it') || c.toLowerCase().includes('system'))
    );
    if (itBottlenecks.length > 0) {
      suggestions.push({
        id: 'suggestion_it_resources',
        resourceType: ResourceType.ITSupport,
        currentAllocation: await this.estimateCurrentAllocation(ResourceType.ITSupport),
        suggestedAllocation: await this.estimateCurrentAllocation(ResourceType.ITSupport) + 1,
        direction: 'increase',
        expectedImpact: 'Faster equipment provisioning and system access setup',
        estimatedTimeSavingsHours: Math.round(
          itBottlenecks.reduce((sum, b) => sum + b.averageDelayHours, 0) * 0.5 *
          metrics.activeWorkflows / 4
        ),
        priority: SuggestionPriority.Medium,
        affectedProcessTypes: [ProcessType.Joiner, ProcessType.Mover],
        implementationSteps: [
          'Automate common IT provisioning tasks',
          'Pre-configure equipment templates',
          'Implement self-service portal for standard requests',
          'Consider dedicated onboarding IT support during peak periods'
        ]
      });
    }

    // Analyze workload distribution
    if (metrics.activeWorkflows > 50 && metrics.onTimeCompletionRate < 0.8) {
      suggestions.push({
        id: 'suggestion_workload_redistribution',
        resourceType: ResourceType.HRStaff,
        currentAllocation: await this.estimateCurrentAllocation(ResourceType.HRStaff),
        suggestedAllocation: await this.estimateCurrentAllocation(ResourceType.HRStaff),
        direction: 'redistribute',
        expectedImpact: 'Better workload balance and improved on-time completion',
        estimatedTimeSavingsHours: Math.round(metrics.activeWorkflows * 0.5),
        priority: SuggestionPriority.Medium,
        affectedProcessTypes: [ProcessType.Joiner, ProcessType.Mover, ProcessType.Leaver],
        implementationSteps: [
          'Analyze current task distribution across HR team',
          'Identify overloaded team members',
          'Implement workload balancing in task assignment',
          'Consider cross-training for flexibility'
        ]
      });
    }

    // Sort by priority and impact
    suggestions.sort((a, b) => {
      const priorityOrder = { Critical: 0, High: 1, Medium: 2, Low: 3 };
      const priorityDiff = priorityOrder[a.priority] - priorityOrder[b.priority];
      if (priorityDiff !== 0) return priorityDiff;
      return b.estimatedTimeSavingsHours - a.estimatedTimeSavingsHours;
    });

    this.setCache(cacheKey, suggestions);
    return suggestions;
  }

  /**
   * Generates a comprehensive analytics report
   * @param periodDays - Number of days to include in report
   * @returns Complete analytics report
   */
  public async generateReport(periodDays: number = 30): Promise<IAnalyticsReport> {
    const periodEnd = new Date();
    const periodStart = new Date(periodEnd.getTime() - periodDays * 24 * 60 * 60 * 1000);

    const [metrics, bottlenecks, resourceSuggestions] = await Promise.all([
      this.calculateMetrics(periodStart, periodEnd),
      this.identifyBottlenecks(),
      this.generateResourceSuggestions()
    ]);

    const processTypeAnalysis = await this.analyzeByProcessType(periodStart, periodEnd);
    const trends = await this.calculateTrends(periodStart, periodEnd);
    const healthScore = this.calculateHealthScore(metrics, bottlenecks);

    return {
      generatedAt: new Date(),
      periodStart,
      periodEnd,
      healthScore,
      metrics,
      bottlenecks,
      resourceSuggestions,
      processTypeAnalysis,
      trends
    };
  }

  /**
   * Calculates workflow metrics for a period
   */
  public async calculateMetrics(
    startDate?: Date,
    endDate?: Date
  ): Promise<IWorkflowMetrics> {
    const historicalData = await this.getHistoricalData(
      undefined,
      startDate,
      endDate
    );

    const completed = historicalData.filter(h => h.status === WorkflowInstanceStatus.Completed);
    const active = historicalData.filter(h =>
      h.status === WorkflowInstanceStatus.Running || h.status === WorkflowInstanceStatus.Pending
    );

    const completionTimes = completed
      .filter(h => h.completedDate)
      .map(h => this.calculateDurationHours(h.startDate, h.completedDate!));

    const sortedTimes = [...completionTimes].sort((a, b) => a - b);
    const median = sortedTimes.length > 0
      ? sortedTimes[Math.floor(sortedTimes.length / 2)]
      : 0;

    // Calculate on-time rate (assuming 48 hours SLA for demonstration)
    const slaHours = 48;
    const onTimeCount = completionTimes.filter(t => t <= slaHours).length;

    return {
      totalWorkflows: historicalData.length,
      completedWorkflows: completed.length,
      activeWorkflows: active.length,
      averageCompletionTimeHours: completionTimes.length > 0
        ? Math.round(this.calculateMean(completionTimes) * 10) / 10
        : 0,
      medianCompletionTimeHours: Math.round(median * 10) / 10,
      onTimeCompletionRate: completed.length > 0
        ? Math.round((onTimeCount / completed.length) * 100) / 100
        : 1,
      averageStepsPerWorkflow: historicalData.length > 0
        ? Math.round(historicalData.reduce((sum, h) => sum + h.totalSteps, 0) / historicalData.length)
        : 0,
      approvalRate: 0.85, // Placeholder - would calculate from actual data
      rejectionRate: 0.05 // Placeholder - would calculate from actual data
    };
  }

  // ==================== Private Helper Methods ====================

  /**
   * Gets historical workflow data from SharePoint lists
   */
  private async getHistoricalData(
    processType?: ProcessType,
    startDate?: Date,
    endDate?: Date
  ): Promise<IHistoricalWorkflowData[]> {
    try {
      const lookbackDate = startDate ||
        new Date(Date.now() - this.config.historicalDaysLookback * 24 * 60 * 60 * 1000);
      const endDateTime = endDate || new Date();

      // Filter by StartedDate (the actual column name in JML_WorkflowInstances)
      const filter = `StartedDate ge datetime'${lookbackDate.toISOString()}' and StartedDate le datetime'${endDateTime.toISOString()}'`;

      const items = await this.sp.web.lists
        .getByTitle(this.instanceListName)
        .items
        .filter(filter)
        .select(
          'Id', 'Title', 'Status', 'StartedDate', 'CompletedDate',
          'TotalSteps', 'CompletedSteps', 'Context'
        )
        .top(1000)();

      // Map items and filter by processType if specified (processType is in Context JSON)
      const mappedItems = items.map((item: Record<string, unknown>) => {
        // Parse context to get processType and department
        let context: { processType?: ProcessType; department?: string } = {};
        try {
          if (item.Context) {
            context = JSON.parse(item.Context as string);
          }
        } catch {
          // Ignore parse errors
        }

        return {
          instanceId: String(item.Id),
          processType: context.processType || ProcessType.Joiner,
          startDate: new Date(item.StartedDate as string),
          completedDate: item.CompletedDate ? new Date(item.CompletedDate as string) : null,
          status: item.Status as WorkflowInstanceStatus,
          totalSteps: (item.TotalSteps as number) || 0,
          completedSteps: (item.CompletedSteps as number) || 0,
          stepDurations: [], // Would be populated from step status list
          department: context.department,
          complexity: this.determineComplexity((item.TotalSteps as number) || 0)
        };
      });

      // Filter by processType if specified
      if (processType) {
        return mappedItems.filter(item => item.processType === processType);
      }

      return mappedItems;
    } catch (error) {
      console.warn('Error fetching historical data, returning mock data:', error);
      return this.generateMockHistoricalData(processType);
    }
  }

  /**
   * Generates mock historical data for development/testing
   */
  private generateMockHistoricalData(processType?: ProcessType): IHistoricalWorkflowData[] {
    const processTypes = processType
      ? [processType]
      : [ProcessType.Joiner, ProcessType.Mover, ProcessType.Leaver];

    const mockData: IHistoricalWorkflowData[] = [];
    const now = Date.now();

    for (let i = 0; i < 50; i++) {
      const type = processTypes[i % processTypes.length];
      const startDate = new Date(now - Math.random() * 90 * 24 * 60 * 60 * 1000);
      const isCompleted = Math.random() > 0.3;
      const baseHours = type === ProcessType.Joiner ? 72 :
                        type === ProcessType.Mover ? 48 : 36;
      const completedDate = isCompleted
        ? new Date(startDate.getTime() + (baseHours + (Math.random() - 0.5) * 24) * 60 * 60 * 1000)
        : null;

      mockData.push({
        instanceId: `mock_${i}`,
        processType: type,
        startDate,
        completedDate,
        status: isCompleted ? WorkflowInstanceStatus.Completed : WorkflowInstanceStatus.Running,
        totalSteps: type === ProcessType.Joiner ? 12 : type === ProcessType.Mover ? 8 : 10,
        completedSteps: isCompleted ? (type === ProcessType.Joiner ? 12 : type === ProcessType.Mover ? 8 : 10) : Math.floor(Math.random() * 6),
        stepDurations: [],
        department: ['HR', 'IT', 'Finance', 'Operations'][i % 4],
        complexity: ['low', 'medium', 'high'][i % 3] as 'low' | 'medium' | 'high'
      });
    }

    return mockData;
  }

  /**
   * Determines workflow complexity based on step count
   */
  private determineComplexity(stepCount: number): 'low' | 'medium' | 'high' {
    if (stepCount <= 5) return 'low';
    if (stepCount <= 10) return 'high';
    return 'high';
  }

  /**
   * Calculates duration in hours between two dates
   */
  private calculateDurationHours(start: Date, end: Date): number {
    return (end.getTime() - start.getTime()) / (1000 * 60 * 60);
  }

  /**
   * Calculates mean of an array of numbers
   */
  private calculateMean(values: number[]): number {
    if (values.length === 0) return 0;
    return values.reduce((sum, v) => sum + v, 0) / values.length;
  }

  /**
   * Calculates standard deviation
   */
  private calculateStandardDeviation(values: number[], mean: number): number {
    if (values.length === 0) return 0;
    const squaredDiffs = values.map(v => Math.pow(v - mean, 2));
    return Math.sqrt(this.calculateMean(squaredDiffs));
  }

  /**
   * Calculates progress factor for prediction adjustment
   */
  private calculateProgressFactor(instance: IWorkflowInstance): number {
    const progress = instance.ProgressPercentage || 0;
    // If 50% done, we estimate remaining time is 50% of average total
    return Math.max(0.1, 1 - (progress / 100));
  }

  /**
   * Creates a low confidence prediction when sample size is insufficient
   */
  private createLowConfidencePrediction(
    instance: IWorkflowInstance,
    sampleSize: number
  ): ITimePrediction {
    // Use default estimates based on process type
    const defaultHours = instance.ProcessType === ProcessType.Joiner ? 72 :
                         instance.ProcessType === ProcessType.Mover ? 48 : 36;

    return {
      predictedHours: defaultHours,
      predictedCompletionDate: new Date(Date.now() + defaultHours * 60 * 60 * 1000),
      lowerBoundHours: defaultHours * 0.5,
      upperBoundHours: defaultHours * 2,
      confidence: 0.3,
      sampleSize,
      factors: [{
        name: 'Insufficient Data',
        impact: 'neutral',
        weight: 1,
        description: `Only ${sampleSize} historical samples available. Need at least ${this.config.minSampleSize} for reliable predictions.`
      }]
    };
  }

  /**
   * Analyzes factors affecting the prediction
   */
  private async analyzePredictionFactors(
    instance: IWorkflowInstance,
    historicalData: IHistoricalWorkflowData[]
  ): Promise<IPredictionFactor[]> {
    const factors: IPredictionFactor[] = [];

    // Parse variables from JSON string if present
    const variables = instance.Variables ? JSON.parse(instance.Variables) as Record<string, unknown> : {};

    // Department factor
    if (variables.department) {
      const deptData = historicalData.filter(h => h.department === variables.department);
      if (deptData.length >= 5) {
        const deptAvg = this.calculateMean(
          deptData.filter(h => h.completedDate).map(h =>
            this.calculateDurationHours(h.startDate, h.completedDate!)
          )
        );
        const overallAvg = this.calculateMean(
          historicalData.filter(h => h.completedDate).map(h =>
            this.calculateDurationHours(h.startDate, h.completedDate!)
          )
        );

        if (deptAvg < overallAvg * 0.9) {
          factors.push({
            name: 'Department Performance',
            impact: 'positive',
            weight: 0.15,
            description: `${variables.department} department typically completes faster than average`
          });
        } else if (deptAvg > overallAvg * 1.1) {
          factors.push({
            name: 'Department Performance',
            impact: 'negative',
            weight: 0.15,
            description: `${variables.department} department typically takes longer than average`
          });
        }
      }
    }

    // Complexity factor - use TotalSteps from the instance
    const complexity = this.determineComplexity(instance.TotalSteps || 10);
    if (complexity === 'high') {
      factors.push({
        name: 'Workflow Complexity',
        impact: 'negative',
        weight: 0.2,
        description: 'High complexity workflow with many steps'
      });
    }

    // Day of week factor
    const dayOfWeek = new Date().getDay();
    if (dayOfWeek === 5 || dayOfWeek === 6) { // Friday or Saturday
      factors.push({
        name: 'Weekend Impact',
        impact: 'negative',
        weight: 0.1,
        description: 'Workflows started near weekend may experience delays'
      });
    }

    return factors;
  }

  /**
   * Calculates prediction confidence
   */
  private calculateConfidence(sampleSize: number, stdDev: number, mean: number): number {
    // Base confidence from sample size (more samples = higher confidence)
    const sampleConfidence = Math.min(1, sampleSize / 50);

    // Variance penalty (high variance = lower confidence)
    const cv = mean > 0 ? stdDev / mean : 1; // Coefficient of variation
    const varianceConfidence = Math.max(0.3, 1 - cv);

    return sampleConfidence * varianceConfidence;
  }

  /**
   * Analyzes step durations to identify delays
   */
  private analyzeStepDurations(historicalData: IHistoricalWorkflowData[]): Record<string, {
    stepName: string;
    averageDelay: number;
    frequency: number;
    affectedCount: number;
    recentTrend: number[];
  }> {
    // Mock analysis for demonstration
    return {
      'manager_approval': {
        stepName: 'Manager Approval',
        averageDelay: 8.5,
        frequency: 0.4,
        affectedCount: Math.floor(historicalData.length * 0.4),
        recentTrend: [7, 8, 9, 8.5]
      },
      'it_provisioning': {
        stepName: 'IT Equipment Provisioning',
        averageDelay: 4.2,
        frequency: 0.25,
        affectedCount: Math.floor(historicalData.length * 0.25),
        recentTrend: [5, 4.5, 4, 4.2]
      },
      'hr_documentation': {
        stepName: 'HR Documentation Review',
        averageDelay: 2.8,
        frequency: 0.15,
        affectedCount: Math.floor(historicalData.length * 0.15),
        recentTrend: [3, 2.9, 2.8, 2.8]
      }
    };
  }

  /**
   * Calculates bottleneck severity
   */
  private calculateBottleneckSeverity(
    averageDelay: number,
    frequency: number
  ): BottleneckSeverity {
    const impact = averageDelay * frequency;
    if (impact > 5) return BottleneckSeverity.Critical;
    if (impact > 3) return BottleneckSeverity.High;
    if (impact > 1.5) return BottleneckSeverity.Medium;
    return BottleneckSeverity.Low;
  }

  /**
   * Categorizes the type of bottleneck
   */
  private categorizeBottleneckType(analysis: {
    stepName: string;
    averageDelay: number;
    frequency: number;
  }): BottleneckType {
    const name = analysis.stepName.toLowerCase();
    if (name.includes('approval')) return BottleneckType.ApprovalDelay;
    if (name.includes('it') || name.includes('system') || name.includes('provisioning')) {
      return BottleneckType.ResourceConstraint;
    }
    if (name.includes('external') || name.includes('vendor')) {
      return BottleneckType.ExternalDependency;
    }
    return BottleneckType.ProcessComplexity;
  }

  /**
   * Identifies root causes for a bottleneck
   */
  private identifyRootCauses(analysis: {
    stepName: string;
    averageDelay: number;
  }): string[] {
    const causes: string[] = [];
    const name = analysis.stepName.toLowerCase();

    if (name.includes('approval')) {
      causes.push('Approvers have high workload');
      causes.push('Approval requests not prioritized');
      if (analysis.averageDelay > 8) {
        causes.push('No escalation mechanism in place');
      }
    }

    if (name.includes('it') || name.includes('provisioning')) {
      causes.push('Manual provisioning processes');
      causes.push('Equipment inventory constraints');
    }

    if (name.includes('documentation') || name.includes('review')) {
      causes.push('Document review backlog');
      causes.push('Complex compliance requirements');
    }

    return causes;
  }

  /**
   * Generates suggestions to resolve a bottleneck
   */
  private generateBottleneckSuggestions(
    analysis: { stepName: string; averageDelay: number },
    rootCauses: string[]
  ): string[] {
    const suggestions: string[] = [];
    const name = analysis.stepName.toLowerCase();

    if (name.includes('approval')) {
      suggestions.push('Implement auto-escalation after 24 hours');
      suggestions.push('Enable parallel approvals where possible');
      suggestions.push('Set up mobile approval notifications');
    }

    if (name.includes('it') || name.includes('provisioning')) {
      suggestions.push('Automate standard equipment requests');
      suggestions.push('Pre-configure user accounts before start date');
      suggestions.push('Maintain buffer inventory of common equipment');
    }

    if (name.includes('documentation')) {
      suggestions.push('Use document templates to reduce review time');
      suggestions.push('Implement automated document validation');
    }

    // Add general suggestions based on delay severity
    if (analysis.averageDelay > 8) {
      suggestions.push('Consider splitting this step into smaller tasks');
      suggestions.push('Add dedicated resources for this step');
    }

    return suggestions;
  }

  /**
   * Calculates bottleneck trend
   */
  private calculateBottleneckTrend(analysis: {
    recentTrend: number[];
  }): 'improving' | 'stable' | 'worsening' {
    const trend = analysis.recentTrend;
    if (trend.length < 2) return 'stable';

    const recent = trend.slice(-2);
    const older = trend.slice(0, -2);
    const recentAvg = this.calculateMean(recent);
    const olderAvg = older.length > 0 ? this.calculateMean(older) : recentAvg;

    if (recentAvg < olderAvg * 0.9) return 'improving';
    if (recentAvg > olderAvg * 1.1) return 'worsening';
    return 'stable';
  }

  /**
   * Estimates current resource allocation
   */
  private async estimateCurrentAllocation(resourceType: ResourceType): Promise<number> {
    // Would query actual resource data - using estimates
    const estimates: Record<ResourceType, number> = {
      [ResourceType.HRStaff]: 5,
      [ResourceType.ITSupport]: 3,
      [ResourceType.Manager]: 15,
      [ResourceType.SystemCapacity]: 100,
      [ResourceType.ExternalVendor]: 2,
      [ResourceType.TrainingResource]: 4
    };
    return estimates[resourceType] || 5;
  }

  /**
   * Analyzes metrics by process type
   */
  private async analyzeByProcessType(
    startDate: Date,
    endDate: Date
  ): Promise<IProcessTypeAnalysis[]> {
    const analysis: IProcessTypeAnalysis[] = [];

    for (const processType of [ProcessType.Joiner, ProcessType.Mover, ProcessType.Leaver]) {
      const data = await this.getHistoricalData(processType, startDate, endDate);
      const completed = data.filter(d => d.completedDate);
      const completionTimes = completed.map(d =>
        this.calculateDurationHours(d.startDate, d.completedDate!)
      );

      const slaHours = processType === ProcessType.Joiner ? 72 : 48;
      const onTimeCount = completionTimes.filter(t => t <= slaHours).length;

      const bottlenecks = await this.identifyBottlenecks(processType);

      analysis.push({
        processType,
        totalInstances: data.length,
        completedInstances: completed.length,
        averageCompletionHours: completionTimes.length > 0
          ? Math.round(this.calculateMean(completionTimes) * 10) / 10
          : 0,
        onTimeRate: completed.length > 0
          ? Math.round((onTimeCount / completed.length) * 100) / 100
          : 1,
        topBottleneck: bottlenecks.length > 0 ? bottlenecks[0].stepName : null,
        trend: 'stable' // Would calculate from historical comparison
      });
    }

    return analysis;
  }

  /**
   * Calculates trend data over time
   */
  private async calculateTrends(
    startDate: Date,
    endDate: Date
  ): Promise<ITrendData[]> {
    const trends: ITrendData[] = [];
    const dayMs = 24 * 60 * 60 * 1000;
    const totalDays = Math.ceil((endDate.getTime() - startDate.getTime()) / dayMs);
    const interval = Math.max(1, Math.floor(totalDays / 10)); // ~10 data points

    for (let i = 0; i < totalDays; i += interval) {
      const date = new Date(startDate.getTime() + i * dayMs);
      const nextDate = new Date(date.getTime() + interval * dayMs);

      const data = await this.getHistoricalData(undefined, date, nextDate);
      const completed = data.filter(d => d.completedDate);

      if (completed.length > 0) {
        const avgTime = this.calculateMean(
          completed.map(d => this.calculateDurationHours(d.startDate, d.completedDate!))
        );

        trends.push({
          date,
          metric: 'averageCompletionTime',
          value: Math.round(avgTime * 10) / 10,
          trend: 'stable'
        });
      }
    }

    // Calculate trend directions
    for (let i = 1; i < trends.length; i++) {
      const prev = trends[i - 1].value;
      const curr = trends[i].value;
      trends[i].trend = curr < prev * 0.95 ? 'down' :
                        curr > prev * 1.05 ? 'up' : 'stable';
    }

    return trends;
  }

  /**
   * Calculates overall health score
   */
  private calculateHealthScore(
    metrics: IWorkflowMetrics,
    bottlenecks: IBottleneck[]
  ): number {
    let score = 100;

    // Deduct for low on-time rate
    if (metrics.onTimeCompletionRate < 0.9) {
      score -= (0.9 - metrics.onTimeCompletionRate) * 50;
    }

    // Deduct for critical bottlenecks
    const criticalCount = bottlenecks.filter(b =>
      b.severity === BottleneckSeverity.Critical
    ).length;
    score -= criticalCount * 10;

    // Deduct for high bottlenecks
    const highCount = bottlenecks.filter(b =>
      b.severity === BottleneckSeverity.High
    ).length;
    score -= highCount * 5;

    // Deduct for high rejection rate
    if (metrics.rejectionRate > 0.1) {
      score -= (metrics.rejectionRate - 0.1) * 100;
    }

    return Math.max(0, Math.min(100, Math.round(score)));
  }

  /**
   * Gets item from cache if valid
   */
  private getFromCache<T>(key: string): T | null {
    if (!this.config.enableCaching) return null;

    const entry = this.cache.get(key) as ICacheEntry<T> | undefined;
    if (!entry) return null;

    if (new Date() > entry.expiresAt) {
      this.cache.delete(key);
      return null;
    }

    return entry.data;
  }

  /**
   * Sets item in cache
   */
  private setCache<T>(key: string, data: T): void {
    if (!this.config.enableCaching) return;

    const now = new Date();
    const expiresAt = new Date(now.getTime() + this.config.cacheDurationMinutes * 60 * 1000);

    this.cache.set(key, {
      data,
      timestamp: now,
      expiresAt
    });
  }

  /**
   * Clears all cached data
   */
  public clearCache(): void {
    this.cache.clear();
  }
}
