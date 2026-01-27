// AI Models and Interfaces
// TypeScript interfaces for AI-powered features

import { IJmlTask } from './IJmlTask';
import { IBaseListItem } from './ICommon';

/**
 * AI Feature Types
 */
export enum AIFeatureType {
  TaskRecommendation = 'Task Recommendation',
  PredictiveAnalytics = 'Predictive Analytics',
  Chatbot = 'Chatbot',
  AutoCategorization = 'Auto Categorization',
  SentimentAnalysis = 'Sentiment Analysis',
  DocumentGeneration = 'Document Generation',
  RiskAssessment = 'Risk Assessment',
  ReportNarrative = 'Report Narrative'
}

/**
 * Sentiment Analysis Result
 */
export interface ISentimentAnalysis {
  sentiment: 'positive' | 'neutral' | 'negative';
  score: number; // -1 to 1
  confidence: number; // 0 to 1
  keyPhrases: string[];
  emotions?: {
    joy?: number;
    sadness?: number;
    anger?: number;
    fear?: number;
    surprise?: number;
  };
  summary: string;
  recommendations?: string[];
}

/**
 * Task Recommendation
 */
export interface ITaskRecommendation {
  task: IJmlTask;
  relevanceScore: number; // 0 to 1
  reasoning: string;
  estimatedHours?: number;
  dependencies?: string[];
  suggestedAssignee?: string;
  priority?: 'High' | 'Medium' | 'Low';
}

/**
 * Completion Prediction
 */
export interface ICompletionPrediction {
  predictedDate: Date;
  confidence: number; // 0 to 1
  factors: {
    factor: string;
    impact: 'positive' | 'negative' | 'neutral';
    weight: number;
  }[];
  risks: string[];
  recommendations: string[];
  historicalAccuracy?: number;
}

/**
 * Process Risk Assessment
 */
export interface IRiskAssessment {
  riskLevel: 'Low' | 'Medium' | 'High' | 'Critical';
  riskScore: number; // 0 to 100
  complexity: 'Simple' | 'Moderate' | 'Complex' | 'Very Complex';
  factors: {
    category: string;
    description: string;
    severity: 'Low' | 'Medium' | 'High';
    mitigation?: string;
  }[];
  recommendations: string[];
  requiresApproval: boolean;
  suggestedReviewers?: string[];
}

/**
 * Document Generation Request
 */
export interface IDocumentGenerationRequest {
  templateType: DocumentTemplateType;
  templateId?: number;
  data: any;
  customInstructions?: string;
  format?: 'markdown' | 'html' | 'plain';
  includeSignature?: boolean;
  tone?: 'formal' | 'casual' | 'friendly' | 'professional';
}

/**
 * Document Template Types
 */
export enum DocumentTemplateType {
  OfferLetter = 'Offer Letter',
  WelcomeEmail = 'Welcome Email',
  ExitSummary = 'Exit Summary',
  TransferNotification = 'Transfer Notification',
  TaskInstructions = 'Task Instructions',
  ApprovalRequest = 'Approval Request',
  CompletionReport = 'Completion Report',
  FeedbackSurvey = 'Feedback Survey'
}

/**
 * Generated Document
 */
export interface IGeneratedDocument {
  content: string;
  format: 'markdown' | 'html' | 'plain';
  title: string;
  metadata: {
    generatedAt: Date;
    templateType: DocumentTemplateType;
    wordCount: number;
    estimatedReadTime: number; // in minutes
  };
  suggestions?: string[];
}

/**
 * Chatbot Message
 */
export interface IChatMessage {
  id: string;
  role: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
  context?: any;
  citations?: {
    title: string;
    url?: string;
    content: string;
  }[];
  actions?: IChatAction[];
}

/**
 * Chatbot Action
 */
export interface IChatAction {
  type: 'navigation' | 'create' | 'search' | 'link';
  label: string;
  payload: any;
  icon?: string;
}

/**
 * Chatbot Context
 */
export interface IChatContext {
  processId?: number;
  userId?: string;
  department?: string;
  role?: string;
  conversationHistory: IChatMessage[];
  sessionId: string;
}

/**
 * AI Analysis Request
 */
export interface IAIAnalysisRequest {
  type: AIFeatureType;
  data: any;
  options?: {
    includeExplanation?: boolean;
    maxResults?: number;
    threshold?: number;
    model?: string;
  };
}

/**
 * AI Analysis Response
 */
export interface IAIAnalysisResponse<T = any> {
  success: boolean;
  data?: T;
  error?: string;
  metadata: {
    timestamp: Date;
    processingTime: number; // milliseconds
    model: string;
    tokensUsed?: number;
    confidence?: number;
  };
  explanation?: string;
}

/**
 * AI Configuration
 */
export interface IAIConfig extends IBaseListItem {
  FeatureType: AIFeatureType;
  IsEnabled: boolean;
  Endpoint: string;
  ApiKey?: string;
  DeploymentName?: string;
  ModelVersion?: string;
  Temperature?: number; // 0 to 1
  MaxTokens?: number;
  TopP?: number; // 0 to 1
  FrequencyPenalty?: number; // -2 to 2
  PresencePenalty?: number; // -2 to 2
  SystemPrompt?: string;
  Configuration?: string; // JSON configuration
}

/**
 * AI Usage Log
 */
export interface IAIUsageLog extends IBaseListItem {
  FeatureType: AIFeatureType;
  ProcessID?: number;
  UserId?: string;
  UserEmail?: string;
  RequestData?: string;
  ResponseData?: string;
  TokensUsed?: number;
  ProcessingTime?: number;
  Success: boolean;
  ErrorMessage?: string;
  ModelVersion?: string;
}

/**
 * Training Data
 */
export interface ITrainingData extends IBaseListItem {
  FeatureType: AIFeatureType;
  Category: string;
  InputData: string;
  ExpectedOutput: string;
  ActualOutput?: string;
  Feedback?: 'positive' | 'negative' | 'neutral';
  IsValidated: boolean;
  ValidatedBy?: string;
  ValidatedDate?: Date;
}

/**
 * Smart Suggestion
 */
export interface ISmartSuggestion {
  type: 'task' | 'assignee' | 'timeline' | 'resource' | 'approval';
  title: string;
  description: string;
  confidence: number;
  reasoning: string;
  action?: {
    label: string;
    handler: () => void;
  };
  metadata?: any;
}

/**
 * Process Insights
 */
export interface IProcessInsights {
  processId: number;
  summary: string;
  keyFindings: string[];
  bottlenecks?: {
    taskCode: string;
    taskTitle: string;
    averageDelay: number; // hours
    recommendation: string;
  }[];
  suggestions: ISmartSuggestion[];
  comparisonToSimilar?: {
    metric: string;
    yourValue: number;
    averageValue: number;
    percentile: number;
  }[];
  riskAssessment?: IRiskAssessment;
}

/**
 * Anomaly Detection
 */
export interface IAnomalyDetection {
  isAnomaly: boolean;
  anomalyScore: number; // 0 to 1
  anomalyType?: 'duration' | 'cost' | 'task_count' | 'approval_time' | 'resource_usage';
  description: string;
  expectedValue: number;
  actualValue: number;
  deviation: number; // percentage
  recommendation: string;
}

/**
 * Predictive Model Metrics
 */
export interface IPredictiveModelMetrics {
  modelName: string;
  accuracy: number; // 0 to 1
  precision: number;
  recall: number;
  f1Score: number;
  lastTrainedDate: Date;
  sampleSize: number;
  features: string[];
}

/**
 * Auto-categorization Result
 */
export interface IAutoCategorization {
  category: string;
  subcategory?: string;
  confidence: number;
  reasoning: string;
  tags: string[];
  priority: 'High' | 'Medium' | 'Low';
  complexity: 'Simple' | 'Moderate' | 'Complex' | 'Very Complex';
  estimatedDuration?: number; // in days
}

/**
 * Feedback Analysis
 */
export interface IFeedbackAnalysis {
  overallSentiment: ISentimentAnalysis;
  themes: {
    theme: string;
    mentions: number;
    sentiment: 'positive' | 'neutral' | 'negative';
    examples: string[];
  }[];
  actionItems: {
    priority: 'High' | 'Medium' | 'Low';
    item: string;
    department?: string;
  }[];
  summary: string;
  retentionRisk?: 'Low' | 'Medium' | 'High';
}

/**
 * Report Narrative Generation Request
 */
export interface IReportNarrativeRequest {
  reportType: string; // 'Executive', 'Operational', 'Compliance', etc.
  reportTitle: string;
  dataType: string; // MetricType from IReportBuilder
  data: any; // The actual report data (metrics, charts data, etc.)
  filters?: any; // IAnalyticsFilters from IReportBuilder
  narrativeStyle: NarrativeStyle;
  tone: 'professional' | 'friendly' | 'technical' | 'marketing';
  length: 'brief' | 'standard' | 'detailed';
  focusAreas: FocusArea[];
  templateId?: string; // Optional template to use as starting point
  customInstructions?: string;
}

/**
 * Narrative Styles for AI-generated report content
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
 * Focus Areas for report narratives
 */
export type FocusArea = 'trends' | 'insights' | 'risks' | 'opportunities' | 'recommendations' | 'comparisons';

/**
 * Generated Report Narrative Result
 */
export interface IGeneratedReportNarrative {
  content: string; // The narrative text in markdown
  keyMetrics: string[]; // Highlighted key metrics
  insights: string[]; // Key insights discovered
  recommendations: string[]; // Action recommendations
  sentiment: 'positive' | 'neutral' | 'cautionary' | 'negative';
  metadata: {
    wordCount: number;
    generatedAt: Date;
    tokensUsed?: number;
    narrativeStyle: NarrativeStyle;
    focusAreas: FocusArea[];
  };
  citations?: {
    dataSource: string;
    metric: string;
    value: string | number;
  }[];
}

/**
 * Narrative Template Variable
 */
export interface INarrativeTemplateVariable {
  name: string;
  type: 'string' | 'number' | 'date' | 'percentage' | 'currency';
  description: string;
  required: boolean;
  defaultValue?: any;
  format?: string; // e.g., 'MM/DD/YYYY' for dates, '$#,##0.00' for currency
}
