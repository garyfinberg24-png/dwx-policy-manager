// Advanced Analytics Models
// Interfaces for enhanced reporting and analytics

/**
 * Time-to-completion trend data point
 */
export interface ICompletionTrend {
  date: Date;
  processType: string;
  averageDays: number;
  count: number;
  department?: string;
}

/**
 * Cost analysis data
 */
export interface ICostAnalysis {
  department: string;
  processType: string;
  totalCost: number;
  averageCost: number;
  processCount: number;
  budgetUtilization: number;
}

/**
 * Task bottleneck identification
 */
export interface ITaskBottleneck {
  taskName: string;
  taskCategory: string;
  averageCompletionDays: number;
  delayedCount: number;
  totalCount: number;
  delayPercentage: number;
  assignedDepartment: string;
}

/**
 * Manager workload distribution
 */
export interface IManagerWorkload {
  managerId: number;
  managerName: string;
  managerEmail: string;
  activeProcesses: number;
  completedProcesses: number;
  overdueProcesses: number;
  totalTasks: number;
  completedTasks: number;
  workloadScore: number;
  department: string;
}

/**
 * Compliance scorecard entry
 */
export interface IComplianceScore {
  category: string;
  requiredItems: number;
  completedItems: number;
  complianceRate: number;
  criticalIssues: number;
  warnings: number;
  lastAuditDate?: Date;
}

/**
 * SLA adherence metrics
 */
export interface ISLAMetric {
  processType: string;
  slaTarget: number;
  actualAverage: number;
  adherenceRate: number;
  metCount: number;
  missedCount: number;
  totalCount: number;
  department?: string;
}

/**
 * Employee satisfaction score (NPS)
 */
export interface IEmployeeSatisfaction {
  processId: number;
  employeeName: string;
  processType: string;
  npsScore: number;
  category: 'Promoter' | 'Passive' | 'Detractor';
  feedback?: string;
  surveyDate: Date;
  department: string;
}

/**
 * NPS Summary
 */
export interface INPSSummary {
  processType: string;
  promoters: number;
  passives: number;
  detractors: number;
  totalResponses: number;
  npsScore: number;
  averageScore: number;
}

/**
 * First-day readiness score
 */
export interface IFirstDayReadiness {
  processId: number;
  employeeName: string;
  startDate: Date;
  itEquipmentReady: boolean;
  accessProvisioned: boolean;
  workspaceReady: boolean;
  documentationComplete: boolean;
  overallScore: number;
  readinessPercentage: number;
}

/**
 * Equipment allocation efficiency
 */
export interface IEquipmentEfficiency {
  equipmentType: string;
  totalAllocated: number;
  onTimeDelivery: number;
  delayedDelivery: number;
  averageLeadTime: number;
  efficiencyRate: number;
  costPerUnit: number;
}

/**
 * Time-in-role data for movers
 */
export interface ITimeInRole {
  employeeId: number;
  employeeName: string;
  previousRole: string;
  newRole: string;
  timeInPreviousRole: number;
  department: string;
  transferDate: Date;
  transferReason?: string;
}

/**
 * Export format options
 */
export enum ExportFormat {
  Excel = 'Excel',
  PDF = 'PDF',
  CSV = 'CSV',
  PowerBI = 'PowerBI'
}

/**
 * Report schedule frequency
 */
export enum ReportFrequency {
  Daily = 'Daily',
  Weekly = 'Weekly',
  Monthly = 'Monthly',
  Quarterly = 'Quarterly'
}

/**
 * Scheduled report configuration
 */
export interface IScheduledReport {
  id: string;
  reportName: string;
  reportType: string;
  frequency: ReportFrequency;
  format: ExportFormat;
  recipients: string[];
  filters?: any;
  enabled: boolean;
  lastRun?: Date;
  nextRun?: Date;
  createdBy: number;
  createdDate: Date;
}

/**
 * Chart data point
 */
export interface IChartDataPoint {
  label: string;
  value: number;
  category?: string;
  color?: string;
  metadata?: any;
}

/**
 * Time series data point
 */
export interface ITimeSeriesPoint {
  date: Date;
  value: number;
  series: string;
}

/**
 * Analytics filter options
 */
export interface IAnalyticsFilters {
  startDate?: Date;
  endDate?: Date;
  departments?: string[];
  processTypes?: string[];
  managers?: number[];
  statuses?: string[];
}

/**
 * Dashboard summary metrics
 */
export interface IDashboardMetrics {
  totalProcesses: number;
  completedProcesses: number;
  activeProcesses: number;
  overdueProcesses: number;
  averageCompletionTime: number;
  totalCost: number;
  complianceRate: number;
  npsScore: number;
  slaAdherence: number;
  firstDayReadiness: number;
}

/**
 * Export options
 */
export interface IExportOptions {
  format: ExportFormat;
  includeCharts: boolean;
  includeSummary: boolean;
  dateRange?: {
    start: Date;
    end: Date;
  };
  filters?: IAnalyticsFilters;
  sections?: string[];
}

/**
 * PowerBI integration config
 */
export interface IPowerBIConfig {
  workspaceId: string;
  reportId: string;
  datasetId: string;
  embedUrl: string;
  accessToken?: string;
}

/**
 * Success Metrics - Key Performance Indicators
 */

/**
 * Time to Onboard metric
 * Days from hire date to first-day ready status
 */
export interface ITimeToOnboard {
  employeeId: number;
  employeeName: string;
  hireDate: Date;
  firstDayReadyDate?: Date;
  daysToOnboard: number;
  processType: string;
  department: string;
  isCompliant: boolean;
  targetDays: number;
}

/**
 * Task Completion Rate metric
 * Percentage of tasks completed on time
 */
export interface ITaskCompletionRate {
  period: Date;
  totalTasks: number;
  completedOnTime: number;
  completedLate: number;
  notCompleted: number;
  onTimeRate: number;
  department?: string;
  processType?: string;
}

/**
 * User Adoption metric
 * Active users per month
 */
export interface IUserAdoption {
  month: Date;
  activeUsers: number;
  totalUsers: number;
  newUsers: number;
  returningUsers: number;
  adoptionRate: number;
  engagementScore: number;
  department?: string;
}

/**
 * Process Cycle Time metric
 * Average days to complete a process
 */
export interface IProcessCycleTime {
  processType: string;
  averageDays: number;
  medianDays: number;
  minDays: number;
  maxDays: number;
  totalProcesses: number;
  targetDays: number;
  varianceFromTarget: number;
  department?: string;
}

/**
 * Error Rate metric
 * Failed processes and tasks
 */
export interface IErrorRate {
  period: Date;
  totalProcesses: number;
  failedProcesses: number;
  totalTasks: number;
  failedTasks: number;
  processErrorRate: number;
  taskErrorRate: number;
  topErrorReasons: Array<{
    reason: string;
    count: number;
    percentage: number;
  }>;
  department?: string;
}

/**
 * User Satisfaction metric
 * NPS or CSAT scores
 */
export interface IUserSatisfaction {
  period: Date;
  totalResponses: number;
  npsScore: number;
  csatScore: number;
  promoters: number;
  passives: number;
  detractors: number;
  averageRating: number;
  topFeedback: Array<{
    category: string;
    sentiment: 'positive' | 'neutral' | 'negative';
    count: number;
  }>;
  department?: string;
}

/**
 * Cost Savings metric
 * Estimated FTE hours saved through automation
 */
export interface ICostSavings {
  period: Date;
  automatedProcesses: number;
  manualHoursSaved: number;
  fteEquivalent: number;
  costSavings: number;
  avgCostPerHour: number;
  roi: number;
  department?: string;
  processType?: string;
}

/**
 * Compliance Score metric
 * Percentage of processes following policy
 */
export interface IComplianceMetric {
  period: Date;
  totalProcesses: number;
  compliantProcesses: number;
  complianceRate: number;
  criticalViolations: number;
  minorViolations: number;
  complianceCategories: Array<{
    category: string;
    compliant: number;
    total: number;
    rate: number;
  }>;
  department?: string;
  processType?: string;
}

/**
 * Overall Success Metrics Dashboard
 */
export interface ISuccessMetricsSummary {
  period: Date;
  timeToOnboard: {
    average: number;
    target: number;
    variance: number;
    trend: 'improving' | 'declining' | 'stable';
  };
  taskCompletionRate: {
    rate: number;
    target: number;
    trend: 'improving' | 'declining' | 'stable';
  };
  userAdoption: {
    activeUsers: number;
    adoptionRate: number;
    trend: 'improving' | 'declining' | 'stable';
  };
  processCycleTime: {
    average: number;
    target: number;
    variance: number;
    trend: 'improving' | 'declining' | 'stable';
  };
  errorRate: {
    rate: number;
    target: number;
    trend: 'improving' | 'declining' | 'stable';
  };
  userSatisfaction: {
    npsScore: number;
    csatScore: number;
    trend: 'improving' | 'declining' | 'stable';
  };
  costSavings: {
    total: number;
    fteEquivalent: number;
    roi: number;
  };
  complianceScore: {
    rate: number;
    target: number;
    trend: 'improving' | 'declining' | 'stable';
  };
}
