// Report Builder Data Models
// Complete type definitions for JML Reports Builder feature

import { IAnalyticsFilters } from './IJmlAnalytics';

/**
 * Report Category
 */
export enum ReportCategory {
  Executive = 'Executive',
  Operational = 'Operational',
  Compliance = 'Compliance',
  Financial = 'Financial',
  HR = 'HR',
  Custom = 'Custom'
}

/**
 * Widget Types
 */
export enum WidgetType {
  KPICard = 'KPICard',
  BarChart = 'BarChart',
  LineChart = 'LineChart',
  PieChart = 'PieChart',
  DonutChart = 'DonutChart',
  AreaChart = 'AreaChart',
  Table = 'Table',
  Text = 'Text',
  Image = 'Image',
  Sparkline = 'Sparkline',
  Gauge = 'Gauge',
  Heatmap = 'Heatmap',
  AINarrative = 'AINarrative'
}

/**
 * Metric Types (maps to AnalyticsService methods)
 */
export enum MetricType {
  CompletionTrends = 'CompletionTrends',
  CostAnalysis = 'CostAnalysis',
  TaskBottlenecks = 'TaskBottlenecks',
  ManagerWorkload = 'ManagerWorkload',
  ComplianceScores = 'ComplianceScores',
  SLAMetrics = 'SLAMetrics',
  NPSSummary = 'NPSSummary',
  FirstDayReadiness = 'FirstDayReadiness',
  DashboardMetrics = 'DashboardMetrics',
  EmployeeLookupMetrics = 'EmployeeLookupMetrics',
  TaskAutomationMetrics = 'TaskAutomationMetrics',
  NotificationMetrics = 'NotificationMetrics',
  ApprovalWorkflowMetrics = 'ApprovalWorkflowMetrics',
  ROISummary = 'ROISummary',
  CustomQuery = 'CustomQuery'
}

/**
 * Color Schemes
 */
export enum ColorScheme {
  Fluent = 'Fluent',
  Corporate = 'Corporate',
  Vibrant = 'Vibrant',
  Monochrome = 'Monochrome',
  Custom = 'Custom'
}

/**
 * Report Frequency for Scheduling
 */
export enum ReportFrequency {
  Daily = 'Daily',
  Weekly = 'Weekly',
  Monthly = 'Monthly',
  Quarterly = 'Quarterly'
}

/**
 * Narrative Styles for AI-generated text
 */
export enum NarrativeStyle {
  ExecutiveSummary = 'ExecutiveSummary',
  DetailedAnalysis = 'DetailedAnalysis',
  ActionItems = 'ActionItems',
  TrendAnalysis = 'TrendAnalysis',
  Comparison = 'Comparison',
  Storytelling = 'Storytelling'
}

/**
 * Narrative Mode (AI vs Manual vs Template)
 */
export enum NarrativeMode {
  Manual = 'Manual',           // Rich text editor only
  Template = 'Template',       // Pre-written template
  AIGenerated = 'AIGenerated', // AI auto-generate
  Hybrid = 'Hybrid'            // Template + AI assist
}

/**
 * Focus Areas for AI narratives
 */
export type FocusArea = 'trends' | 'insights' | 'risks' | 'opportunities' | 'recommendations' | 'comparisons';

/**
 * Report Definition - Main Report Configuration
 */
export interface IReportDefinition {
  Id?: number;
  Title: string;
  Description?: string;
  Category: ReportCategory;

  // Layout Configuration
  layout: IReportLayout;

  // Widgets on the report
  widgets: IReportWidget[];

  // Global Report Settings
  settings: IReportSettings;

  // Filters applied to all widgets
  globalFilters?: IAnalyticsFilters;

  // Scheduling
  schedule?: IReportSchedule;

  // Sharing & Permissions
  sharedWith?: number[];
  isPublic: boolean;

  // Metadata
  createdBy?: number;
  createdDate?: Date;
  modifiedBy?: number;
  modifiedDate?: Date;
  tags?: string[];
}

/**
 * Report Layout Configuration
 */
export interface IReportLayout {
  columns: number;
  rows: number;
  pageSize: 'A4' | 'Letter' | 'Legal' | 'A3';
  orientation: 'portrait' | 'landscape';
  margins: {
    top: number;
    right: number;
    bottom: number;
    left: number;
  };
}

/**
 * Report Widget Configuration
 */
export interface IReportWidget {
  id: string;
  type: WidgetType;

  // Position on grid
  position: {
    col: number;
    row: number;
    width: number;
    height: number;
  };

  // Data Configuration
  dataSource: IWidgetDataSource;

  // Display Configuration
  config: IWidgetConfig;

  // Styling
  style?: IWidgetStyle;
}

/**
 * Widget Data Source Configuration
 */
export interface IWidgetDataSource {
  metric?: MetricType;
  filters?: IAnalyticsFilters;
  aggregation?: 'sum' | 'avg' | 'count' | 'min' | 'max';
  groupBy?: string;
  sortBy?: string;
  sortOrder?: 'asc' | 'desc';
  limit?: number;
  customQuery?: string;

  // For AI Narrative widgets
  linkedWidgetId?: string;
}

/**
 * Widget Display Configuration
 */
export interface IWidgetConfig {
  title: string;
  subtitle?: string;
  showLegend?: boolean;
  showLabels?: boolean;
  showValues?: boolean;
  valueFormat?: 'number' | 'percentage' | 'currency' | 'duration';
  colorScheme?: ColorScheme;
  chartOptions?: any;

  // For AI Narrative widgets
  narrativeMode?: NarrativeMode;
  narrativeStyle?: NarrativeStyle;
  narrativeLength?: 'brief' | 'standard' | 'detailed';
  tone?: 'professional' | 'friendly' | 'technical' | 'marketing';
  focusAreas?: FocusArea[];
  templateId?: string;

  // For Text widgets
  richTextContent?: string;
}

/**
 * Widget Styling
 */
export interface IWidgetStyle {
  backgroundColor?: string;
  borderColor?: string;
  borderWidth?: number;
  borderRadius?: number;
  padding?: number;
  fontSize?: number;
  fontWeight?: 'normal' | 'bold' | 'semibold';
  textAlign?: 'left' | 'center' | 'right';
}

/**
 * Report Settings
 */
export interface IReportSettings {
  // Branding
  companyName?: string;
  companyLogo?: string;
  reportHeader?: string;
  reportFooter?: string;

  // Colors
  primaryColor?: string;
  secondaryColor?: string;
  accentColor?: string;

  // Typography
  fontFamily?: string;
  headerFontSize?: number;
  bodyFontSize?: number;

  // Display Options
  showPageNumbers?: boolean;
  showGeneratedDate?: boolean;
  showFilters?: boolean;
  showWatermark?: boolean;
  watermarkText?: string;
}

/**
 * Report Schedule Configuration
 */
export interface IReportSchedule {
  enabled: boolean;
  frequency: ReportFrequency;
  dayOfWeek?: number;
  dayOfMonth?: number;
  time?: string;
  timezone?: string;

  // Distribution
  recipients: string[];
  emailSubject?: string;
  emailBody?: string;
  attachmentFormat?: 'PDF' | 'Excel' | 'Both';

  // SharePoint Integration
  saveToLibrary?: boolean;
  libraryUrl?: string;
  folderPath?: string;

  // Schedule Metadata
  nextRun?: Date;
  lastRun?: Date;
  lastRunStatus?: 'Success' | 'Failed';
  lastRunError?: string;
}

/**
 * Generated Report Result
 */
export interface IGeneratedReport {
  definition: IReportDefinition;
  data: Map<string, any>;
  generatedAt: Date;
  generatedBy: number;
  generationTimeMs: number;
  filters: IAnalyticsFilters;
  dataSourceTimestamp: Date;
}

/**
 * Report Export Options
 */
export interface IReportExportOptions {
  format: 'PDF' | 'Excel' | 'PowerPoint';
  includeCharts: boolean;
  includeSummary: boolean;
  showFilters: boolean;
  watermark?: string;
}

/**
 * AI Narrative Context
 */
export interface IReportNarrativeContext {
  reportType: ReportCategory;
  reportTitle: string;
  dataType: MetricType;
  data: any;
  filters?: IAnalyticsFilters;
  narrativeStyle: NarrativeStyle;
  tone: 'professional' | 'friendly' | 'technical' | 'marketing';
  length: 'brief' | 'standard' | 'detailed';
  focusAreas: FocusArea[];
}

/**
 * Generated Narrative Result
 */
export interface IGeneratedNarrative {
  content: string;
  keyMetrics: string[];
  insights: string[];
  recommendations: string[];
  sentiment: 'positive' | 'neutral' | 'cautionary' | 'negative';
  metadata: {
    wordCount: number;
    generatedAt: Date;
    tokensUsed?: number;
  };
}

/**
 * Narrative Template
 */
export interface INarrativeTemplate {
  id: string;
  name: string;
  description: string;
  category: string;
  content: string;
  variables: string[];
  suggestedFor?: ReportCategory[];
}

/**
 * Report Template (Pre-built report configuration)
 */
export interface IReportTemplate {
  id: string;
  name: string;
  description: string;
  category: ReportCategory;
  thumbnail?: string;
  definition: IReportDefinition;
  previewUrl?: string;
}
