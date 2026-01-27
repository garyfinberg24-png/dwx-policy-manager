/**
 * License Management Models
 *
 * Feature flag licensing system for JML Hub premium modules.
 *
 * SECURITY NOTE:
 * - Client-side feature flags provide UI-level protection only
 * - For high-security scenarios, implement server-side validation via Azure Functions
 * - License keys can be cryptographically signed for tamper detection
 */

import { IBaseListItem } from './ICommon';

/**
 * Available premium modules that can be licensed
 */
export enum PremiumModule {
  // Analytics & Reporting
  Analytics = 'analytics',
  ROIDashboard = 'roi',
  ROIAnalytics = 'roianalytics',
  ReportsBuilder = 'reports',
  Dashboard = 'dashboard',

  // Employee Management
  SurveyManager = 'survey',
  CVManager = 'cv',
  TalentManager = 'talent',
  SkillsBuilder = 'skills',

  // Policy Management
  PolicyManagement = 'policy',

  // Operations
  DocumentGeneration = 'document',
  SigningService = 'signing',
  AssetManager = 'asset',
  LicenseManagement = 'license',
  ProcurementManager = 'procurement',
  ContractManager = 'contract',
  SLAManager = 'sla',
  FinancialManagement = 'finance',

  // Integration & Compliance
  IntegrationHub = 'integration',
  ComplianceManager = 'compliance',
  EmailManager = 'email',

  // Learning & Engagement
  Gamification = 'gamification',
  QuizBuilder = 'quiz',

  // Advanced Features
  AIAssistant = 'ai',
  ThemeManager = 'theme',

  // Core Features (included in base package, shown in Admin Panel for visibility)
  DocumentHub = 'documenthub',
  ExternalSharingHub = 'externalsharinghub',
  WorkflowMonitor = 'workflowmonitor'
}

/**
 * License tier determines which modules are included
 */
export enum LicenseTier {
  /** Free tier - core JML functionality only */
  Free = 'Free',
  /** Starter - basic premium modules */
  Starter = 'Starter',
  /** Professional - most modules */
  Professional = 'Professional',
  /** Enterprise - all modules + priority support */
  Enterprise = 'Enterprise',
  /** Custom - individually selected modules */
  Custom = 'Custom'
}

/**
 * License status
 */
export enum LicenseStatus {
  Active = 'Active',
  Expired = 'Expired',
  Suspended = 'Suspended',
  Trial = 'Trial',
  PendingActivation = 'Pending Activation'
}

/**
 * License stored in SharePoint list (JML_Licenses)
 */
export interface ILicense extends IBaseListItem {
  /** Unique license key (e.g., "JML-ENT-2024-XXXX-XXXX") */
  LicenseKey: string;

  /** Customer/tenant identifier */
  TenantId: string;

  /** Customer organization name */
  OrganizationName: string;

  /** License tier */
  Tier: LicenseTier;

  /** Current status */
  Status: LicenseStatus;

  /** JSON array of enabled module IDs (for Custom tier) */
  EnabledModules: string;

  /** License activation date */
  ActivatedDate?: Date;

  /** License expiration date */
  ExpirationDate?: Date;

  /** Maximum number of users (0 = unlimited) */
  MaxUsers: number;

  /** Contact email for license holder */
  ContactEmail: string;

  /** Internal notes (not visible to customer) */
  Notes?: string;

  /** Digital signature for tamper detection (optional) */
  Signature?: string;

  /** Last validation timestamp */
  LastValidated?: Date;
}

/**
 * Parsed license data used in the application
 */
export interface ILicenseData {
  isValid: boolean;
  tier: LicenseTier;
  status: LicenseStatus;
  enabledModules: PremiumModule[];
  expirationDate?: Date;
  daysUntilExpiration?: number;
  maxUsers: number;
  organizationName: string;
  isTrial: boolean;
  isExpiringSoon: boolean; // Within 30 days
}

/**
 * Module license check result
 */
export interface IModuleLicenseCheck {
  moduleId: PremiumModule;
  isLicensed: boolean;
  reason?: 'not_in_tier' | 'license_expired' | 'license_suspended' | 'no_license' | 'trial_ended';
}

/**
 * License activation request
 */
export interface ILicenseActivationRequest {
  licenseKey: string;
  tenantId: string;
  contactEmail: string;
  organizationName?: string;
}

/**
 * License activation response
 */
export interface ILicenseActivationResponse {
  success: boolean;
  message: string;
  license?: ILicenseData;
  errorCode?: 'INVALID_KEY' | 'ALREADY_ACTIVATED' | 'EXPIRED' | 'TENANT_MISMATCH' | 'SERVER_ERROR';
}

/**
 * Default modules included in each tier
 */
export const TierModules: Record<LicenseTier, PremiumModule[]> = {
  [LicenseTier.Free]: [],

  [LicenseTier.Starter]: [
    PremiumModule.ThemeManager,
    PremiumModule.EmailManager,
    PremiumModule.DocumentGeneration,
    PremiumModule.SigningService
  ],

  [LicenseTier.Professional]: [
    PremiumModule.ThemeManager,
    PremiumModule.EmailManager,
    PremiumModule.DocumentGeneration,
    PremiumModule.SigningService,
    PremiumModule.Analytics,
    PremiumModule.Dashboard,
    PremiumModule.SurveyManager,
    PremiumModule.AssetManager,
    PremiumModule.LicenseManagement,
    PremiumModule.SLAManager,
    PremiumModule.ReportsBuilder,
    PremiumModule.PolicyManagement,
    PremiumModule.Gamification,
    PremiumModule.QuizBuilder
  ],

  [LicenseTier.Enterprise]: [
    // All modules
    PremiumModule.Analytics,
    PremiumModule.Dashboard,
    PremiumModule.ROIDashboard,
    PremiumModule.ROIAnalytics,
    PremiumModule.ReportsBuilder,
    PremiumModule.SurveyManager,
    PremiumModule.CVManager,
    PremiumModule.TalentManager,
    PremiumModule.SkillsBuilder,
    PremiumModule.PolicyManagement,
    PremiumModule.DocumentGeneration,
    PremiumModule.SigningService,
    PremiumModule.AssetManager,
    PremiumModule.LicenseManagement,
    PremiumModule.ProcurementManager,
    PremiumModule.ContractManager,
    PremiumModule.SLAManager,
    PremiumModule.FinancialManagement,
    PremiumModule.IntegrationHub,
    PremiumModule.ComplianceManager,
    PremiumModule.EmailManager,
    PremiumModule.Gamification,
    PremiumModule.QuizBuilder,
    PremiumModule.AIAssistant,
    PremiumModule.ThemeManager
  ],

  [LicenseTier.Custom]: [] // Defined per-license
};

/**
 * Module metadata for display
 */
export interface IPremiumModuleInfo {
  id: PremiumModule;
  name: string;
  description: string;
  icon: string; // Fluent icon name
  color: string;
  tier: LicenseTier; // Minimum tier required
}

/**
 * Premium module catalog
 */
export const PremiumModuleCatalog: IPremiumModuleInfo[] = [
  {
    id: PremiumModule.SurveyManager,
    name: 'Survey Manager',
    description: 'Employee feedback and engagement surveys',
    icon: 'ChartMultiple24Regular',
    color: '#0078d4',
    tier: LicenseTier.Professional
  },
  {
    id: PremiumModule.CVManager,
    name: 'CV Manager',
    description: 'Skills inventory and certification tracking',
    icon: 'PersonCircle24Regular',
    color: '#8764b8',
    tier: LicenseTier.Enterprise
  },
  {
    id: PremiumModule.TalentManager,
    name: 'Talent Manager',
    description: 'Recruitment pipeline and applicant tracking',
    icon: 'DocumentBulletList24Regular',
    color: '#00ad56',
    tier: LicenseTier.Enterprise
  },
  {
    id: PremiumModule.ThemeManager,
    name: 'Theme Manager',
    description: 'Custom branding and theming',
    icon: 'PaintBrush24Regular',
    color: '#d13438',
    tier: LicenseTier.Starter
  },
  {
    id: PremiumModule.DocumentGeneration,
    name: 'Document Generation',
    description: 'Automated document creation and templates',
    icon: 'DocumentPdf24Regular',
    color: '#ff8c00',
    tier: LicenseTier.Starter
  },
  {
    id: PremiumModule.IntegrationHub,
    name: 'Integration Hub',
    description: 'Connect to HRIS, ITSM, and third-party systems',
    icon: 'PlugConnected24Regular',
    color: '#107c10',
    tier: LicenseTier.Enterprise
  },
  {
    id: PremiumModule.Analytics,
    name: 'Analytics',
    description: 'Custom dashboards and data insights',
    icon: 'DataTrending24Regular',
    color: '#004e8c',
    tier: LicenseTier.Professional
  },
  {
    id: PremiumModule.ROIDashboard,
    name: 'ROI Dashboard',
    description: 'Business value and cost analysis',
    icon: 'MoneyCalculator24Regular',
    color: '#018574',
    tier: LicenseTier.Enterprise
  },
  {
    id: PremiumModule.AssetManager,
    name: 'Asset Manager',
    description: 'IT asset and equipment lifecycle management',
    icon: 'Cube24Regular',
    color: '#5c2d91',
    tier: LicenseTier.Professional
  },
  {
    id: PremiumModule.SLAManager,
    name: 'SLA Manager',
    description: 'SLA tracking and automated escalations',
    icon: 'ClockAlarm24Regular',
    color: '#ea4300',
    tier: LicenseTier.Professional
  },
  {
    id: PremiumModule.EmailManager,
    name: 'Email Manager',
    description: 'Email templates and automated communications',
    icon: 'Mail24Regular',
    color: '#0099bc',
    tier: LicenseTier.Starter
  },
  {
    id: PremiumModule.AIAssistant,
    name: 'AI Assistant',
    description: 'AI-powered recommendations and chatbot',
    icon: 'Sparkle24Regular',
    color: '#e3008c',
    tier: LicenseTier.Enterprise
  },
  {
    id: PremiumModule.ComplianceManager,
    name: 'Compliance Manager',
    description: 'Regulatory compliance and audit trails',
    icon: 'ShieldTask24Regular',
    color: '#0b6a0b',
    tier: LicenseTier.Enterprise
  },
  {
    id: PremiumModule.ROIAnalytics,
    name: 'ROI Analytics',
    description: 'Advanced ROI calculations and forecasting',
    icon: 'ChartPerson24Regular',
    color: '#005a9e',
    tier: LicenseTier.Enterprise
  },
  {
    id: PremiumModule.ProcurementManager,
    name: 'Procurement Manager',
    description: 'Purchase workflows and budget tracking',
    icon: 'Cart24Regular',
    color: '#744da9',
    tier: LicenseTier.Enterprise
  },
  {
    id: PremiumModule.SkillsBuilder,
    name: 'Skills Builder',
    description: 'Learning paths and skill assessments',
    icon: 'Hat24Regular',
    color: '#c239b3',
    tier: LicenseTier.Enterprise
  },
  {
    id: PremiumModule.ReportsBuilder,
    name: 'Reports Builder',
    description: 'Custom report creation and scheduling',
    icon: 'DocumentBulletList24Regular',
    color: '#0078d4',
    tier: LicenseTier.Professional
  },
  {
    id: PremiumModule.ContractManager,
    name: 'Contract Manager',
    description: 'Full lifecycle contract management with clause library, approvals, and obligation tracking',
    icon: 'DocumentContract24Regular',
    color: '#744da9',
    tier: LicenseTier.Enterprise
  },
  {
    id: PremiumModule.Dashboard,
    name: 'Executive Dashboard',
    description: 'Comprehensive executive dashboard with KPIs, trends, and organizational insights',
    icon: 'Board24Regular',
    color: '#0078d4',
    tier: LicenseTier.Professional
  },
  {
    id: PremiumModule.PolicyManagement,
    name: 'Policy Management',
    description: 'Complete policy lifecycle management with authoring, approval workflows, and compliance tracking',
    icon: 'DocumentText24Regular',
    color: '#0b6a0b',
    tier: LicenseTier.Professional
  },
  {
    id: PremiumModule.SigningService,
    name: 'Signing Service',
    description: 'Digital document signing with multi-party signatures, audit trails, and compliance',
    icon: 'Signature24Regular',
    color: '#744da9',
    tier: LicenseTier.Starter
  },
  {
    id: PremiumModule.LicenseManagement,
    name: 'License Management',
    description: 'Track software licenses, renewals, and compliance across your organization',
    icon: 'Key24Regular',
    color: '#5c2d91',
    tier: LicenseTier.Professional
  },
  {
    id: PremiumModule.Gamification,
    name: 'Gamification',
    description: 'Boost engagement with points, badges, leaderboards, and achievement systems',
    icon: 'Trophy24Regular',
    color: '#ffaa00',
    tier: LicenseTier.Professional
  },
  {
    id: PremiumModule.QuizBuilder,
    name: 'Quiz Builder',
    description: 'Create interactive quizzes for training, assessments, and knowledge checks',
    icon: 'Question24Regular',
    color: '#e3008c',
    tier: LicenseTier.Professional
  },
  {
    id: PremiumModule.FinancialManagement,
    name: 'Financial Management',
    description: 'Budget tracking, expense management, and payroll integration for JML processes',
    icon: 'Money24Regular',
    color: '#107c10',
    tier: LicenseTier.Enterprise
  },
  // Core Features (included in base package, shown for visibility)
  {
    id: PremiumModule.DocumentHub,
    name: 'Document Hub',
    description: 'Enterprise document management with versioning, retention policies, and compliance controls',
    icon: 'Folder24Regular',
    color: '#038387',
    tier: LicenseTier.Free
  },
  {
    id: PremiumModule.ExternalSharingHub,
    name: 'External Sharing Hub',
    description: 'Manage external sharing, guest access, and cross-tenant collaboration with governance controls',
    icon: 'Share24Regular',
    color: '#d13438',
    tier: LicenseTier.Free
  },
  {
    id: PremiumModule.WorkflowMonitor,
    name: 'Workflow Monitor',
    description: 'Monitor active workflows, track SLA compliance, and intervene on stuck processes',
    icon: 'Flow24Regular',
    color: '#5c2d91',
    tier: LicenseTier.Free
  }
];
