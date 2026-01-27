// Survey Management System Models

export interface ISurveyTemplate {
  id: number;
  title: string;
  description: string;
  category: SurveyCategory;
  triggerType: SurveyTriggerType;
  triggerOffset: number; // Days from trigger event
  isActive: boolean;
  isAnonymous: boolean;
  allowMultipleResponses: boolean;
  questions: ISurveyQuestion[];
  estimatedMinutes: number;
  createdBy: string;
  createdDate: Date;
  modifiedDate: Date;
  targetAudience?: string; // All, Department, Role, Custom
  reminderFrequency?: number; // Days between reminders
  expiryDays?: number; // Days until survey expires
}

export enum SurveyCategory {
  Onboarding = 'Onboarding',
  Offboarding = 'Offboarding',
  Pulse = 'Pulse Survey',
  Performance = 'Performance',
  Training = 'Training Feedback',
  Manager = 'Manager Effectiveness',
  Wellbeing = 'Wellbeing',
  Engagement = 'Engagement',
  Custom = 'Custom'
}

export enum SurveyTriggerType {
  OnHireDate = 'On Hire Date',
  OnTerminationDate = 'On Termination Date',
  AfterTraining = 'After Training',
  Scheduled = 'Scheduled',
  Manual = 'Manual',
  Recurring = 'Recurring'
}

export interface ISurveyQuestion {
  id: string;
  questionText: string;
  questionType: QuestionType;
  isRequired: boolean;
  order: number;
  options?: string[]; // For multiple choice, dropdown, etc.
  allowOther?: boolean;
  minValue?: number; // For rating/scale questions
  maxValue?: number;
  scaleLabels?: { min: string; max: string };
  conditionalLogic?: IConditionalLogic;
  placeholder?: string;
  helpText?: string;
}

export enum QuestionType {
  ShortText = 'Short Text',
  LongText = 'Long Text',
  MultipleChoice = 'Multiple Choice',
  Checkboxes = 'Checkboxes',
  Dropdown = 'Dropdown',
  LinearScale = 'Linear Scale',
  Rating = 'Rating (Stars)',
  YesNo = 'Yes/No',
  NetPromoterScore = 'NPS (0-10)',
  Date = 'Date',
  Email = 'Email',
  FileUpload = 'File Upload'
}

export interface IConditionalLogic {
  showIf: {
    questionId: string;
    operator: 'equals' | 'contains' | 'greaterThan' | 'lessThan';
    value: any;
  };
}

export interface ISurveyInstance {
  id: number;
  templateId: number;
  templateTitle: string;
  employeeId: number;
  employeeName: string;
  employeeEmail: string;
  department: string;
  scheduledDate: Date;
  sentDate?: Date;
  completedDate?: Date;
  status: SurveyStatus;
  dueDate: Date;
  remindersSent: number;
  lastReminderDate?: Date;
  responseId?: number;
  isAnonymous: boolean;
}

export enum SurveyStatus {
  Scheduled = 'Scheduled',
  Sent = 'Sent',
  InProgress = 'In Progress',
  Completed = 'Completed',
  Expired = 'Expired',
  Cancelled = 'Cancelled'
}

export interface ISurveyResponse {
  id: number;
  instanceId: number;
  templateId: number;
  templateTitle?: string;
  employeeId?: number; // Null for anonymous
  employeeName?: string;
  employeeEmail?: string;
  respondentId?: string;
  respondentName?: string;
  respondentEmail?: string;
  department?: string;
  startedDate: Date;
  completedDate?: Date;
  timeSpentMinutes?: number;
  answers: ISurveyAnswer[];
  isComplete: boolean;
  overallScore?: number;
  npsScore?: number;
  submittedFromDevice?: string;
}

export interface ISurveyAnswer {
  questionId: string;
  questionText: string;
  questionType: QuestionType;
  answer: any; // String, number, array, etc.
  score?: number; // For rating questions
}

export interface ISurveyAnalytics {
  templateId: number;
  templateTitle: string;
  totalSent: number;
  totalResponses: number;
  completionRate: number;
  averageTimeMinutes: number;
  lastResponseDate?: Date;
  departmentBreakdown: IDepartmentStats[];
  questionStats: IQuestionStats[];
  trendData: ITrendData[];
  npsScore?: number; // Net Promoter Score
  sentimentScore?: number; // Calculated from text analysis
}

export interface IDepartmentStats {
  department: string;
  sent: number;
  responses: number;
  completionRate: number;
  averageScore?: number;
}

export interface IQuestionStats {
  questionId: string;
  questionText: string;
  questionType: QuestionType;
  responseCount: number;
  averageScore?: number;
  distribution?: { [key: string]: number }; // For multiple choice
  topAnswers?: string[]; // For text questions
  sentimentScore?: number;
}

export interface ITrendData {
  period: string; // Week/Month/Quarter
  sent: number;
  responses: number;
  completionRate: number;
  averageScore?: number;
}

export interface ISurveyScheduleRule {
  id: number;
  templateId: number;
  isActive: boolean;
  triggerType: SurveyTriggerType;
  offsetDays: number;
  targetDepartments?: string[];
  targetRoles?: string[];
  recurringPattern?: RecurringPattern;
  nextRunDate?: Date;
}

export enum RecurringPattern {
  Weekly = 'Weekly',
  Monthly = 'Monthly',
  Quarterly = 'Quarterly',
  Annually = 'Annually'
}

export interface ISurveyNotification {
  id: number;
  instanceId: number;
  employeeEmail: string;
  notificationType: NotificationType;
  sentDate: Date;
  subject: string;
  body: string;
  status: 'Sent' | 'Failed' | 'Pending';
}

export enum NotificationType {
  Initial = 'Initial Survey Invitation',
  Reminder = 'Reminder',
  FinalReminder = 'Final Reminder',
  ThankYou = 'Thank You',
  Expiring = 'Expiring Soon'
}
