// @ts-nocheck
/**
 * Module Registry - Central registry for all premium modules
 *
 * This file defines:
 * - What lists each module requires
 * - Module metadata for UI display
 * - Dependencies between modules
 * - Provisioning scripts for each module
 *
 * When adding a new premium module:
 * 1. Add module ID to PremiumModule enum in ILicense.ts
 * 2. Add module definition to ModuleRegistry below
 * 3. Create provisioning script in /scripts folder
 * 4. Build components/services for the module
 */

import { PremiumModule, LicenseTier } from './ILicense';

/**
 * SharePoint list field definition for provisioning
 */
export interface IListFieldDefinition {
  name: string;
  displayName: string;
  type: 'Text' | 'Note' | 'Number' | 'DateTime' | 'Boolean' | 'Choice' | 'User' | 'UserMulti' | 'Lookup' | 'URL';
  required?: boolean;
  choices?: string[];
  defaultValue?: string | number | boolean;
  addToDefaultView?: boolean;
}

/**
 * SharePoint list definition for a module
 */
export interface IListDefinition {
  /** Internal list name (e.g., "PM_ReportDefinitions") */
  name: string;
  /** Display title */
  title: string;
  /** Description of the list's purpose */
  description: string;
  /** PowerShell provisioning script filename */
  provisioningScript?: string;
  /** List fields to create */
  fields?: IListFieldDefinition[];
  /** Whether this is a document library */
  isDocumentLibrary?: boolean;
  /** Custom views to create */
  views?: string[];
}

/**
 * Module dependency definition
 */
export interface IModuleDependency {
  moduleId: PremiumModule;
  reason: string;
}

/**
 * Complete module definition
 */
export interface IModuleDefinition {
  /** Unique module identifier */
  id: PremiumModule;
  /** Display name */
  name: string;
  /** Short tagline */
  tagline: string;
  /** Full description */
  description: string;
  /** Minimum license tier required */
  tier: LicenseTier;
  /** Fluent UI icon name */
  icon: string;
  /** Brand color (hex) */
  color: string;
  /** SharePoint lists required by this module */
  requiredLists: IListDefinition[];
  /** Other modules this depends on */
  dependencies?: IModuleDependency[];
  /** Whether this module is new (for UI badges) */
  isNew?: boolean;
  /** Whether this module is popular (for UI badges) */
  isPopular?: boolean;
  /** Category for grouping in UI */
  category: 'analytics' | 'hr' | 'operations' | 'integration' | 'compliance' | 'advanced';
}

/**
 * Module activation status
 */
export interface IModuleActivationStatus {
  moduleId: PremiumModule;
  isLicensed: boolean;
  isProvisioned: boolean;
  missingLists: string[];
  canActivate: boolean;
  activationBlockers: string[];
}

/**
 * Central Module Registry
 *
 * Contains all premium module definitions with their list dependencies.
 * This is the single source of truth for module configuration.
 */
export const ModuleRegistry: Record<PremiumModule, IModuleDefinition> = {
  // ============================================
  // ANALYTICS & REPORTING MODULES
  // ============================================

  [PremiumModule.Analytics]: {
    id: PremiumModule.Analytics,
    name: 'Advanced Analytics',
    tagline: 'Data-driven decisions',
    description: 'Transform JML data into actionable insights with customizable dashboards and predictive analytics.',
    tier: LicenseTier.Professional,
    icon: 'DataTrending24Regular',
    color: '#004e8c',
    category: 'analytics',
    isPopular: true,
    requiredLists: [
      // Uses core lists only - no additional lists needed
    ],
    dependencies: []
  },

  [PremiumModule.ROIDashboard]: {
    id: PremiumModule.ROIDashboard,
    name: 'ROI Dashboard',
    tagline: 'Prove your impact',
    description: 'Measure and demonstrate the business value of your JML processes with comprehensive ROI tracking.',
    tier: LicenseTier.Enterprise,
    icon: 'MoneyCalculator24Regular',
    color: '#018574',
    category: 'analytics',
    requiredLists: [
      // Uses core process data - no additional lists
    ]
  },

  [PremiumModule.ROIAnalytics]: {
    id: PremiumModule.ROIAnalytics,
    name: 'ROI Analytics',
    tagline: 'Maximize returns',
    description: 'Advanced ROI calculations with predictive modeling and forecasting capabilities.',
    tier: LicenseTier.Enterprise,
    icon: 'ChartPerson24Regular',
    color: '#005a9e',
    category: 'analytics',
    requiredLists: [],
    dependencies: [
      { moduleId: PremiumModule.ROIDashboard, reason: 'Extends ROI Dashboard functionality' }
    ]
  },

  [PremiumModule.ReportsBuilder]: {
    id: PremiumModule.ReportsBuilder,
    name: 'Reports Builder',
    tagline: 'Custom reports made easy',
    description: 'Create, schedule, and distribute custom reports with drag-and-drop widgets and AI-powered narratives.',
    tier: LicenseTier.Professional,
    icon: 'DocumentBulletList24Regular',
    color: '#0078d4',
    category: 'analytics',
    isNew: true,
    requiredLists: [
      {
        name: 'PM_ReportDefinitions',
        title: 'Report Definitions',
        description: 'Stores custom report definitions and layouts',
        provisioningScript: 'Create-PM_ReportDefinitions-List.ps1',
        views: ['All Items', 'My Reports', 'Public Reports', 'Report Templates']
      },
      {
        name: 'PM_ReportSchedules',
        title: 'Report Schedules',
        description: 'Scheduled report execution configuration',
        provisioningScript: 'Create-PM_ReportSchedules-List.ps1',
        views: ['All Items', 'Active Schedules', 'My Schedules']
      },
      {
        name: 'PM_ReportExecutionLog',
        title: 'Report Execution Log',
        description: 'Audit log of report generation runs',
        provisioningScript: 'Create-PM_ReportExecutionLog-List.ps1'
      },
      {
        name: 'PM_NarrativeTemplates',
        title: 'Narrative Templates',
        description: 'AI narrative templates for reports',
        provisioningScript: 'Create-PM_NarrativeTemplates-Library.ps1',
        isDocumentLibrary: true
      }
    ]
  },

  // ============================================
  // HR & TALENT MODULES
  // ============================================

  [PremiumModule.SurveyManager]: {
    id: PremiumModule.SurveyManager,
    name: 'Survey Manager',
    tagline: 'Listen. Learn. Act.',
    description: 'Capture employee feedback, measure engagement, and drive continuous improvement with intelligent survey analytics.',
    tier: LicenseTier.Professional,
    icon: 'ChartMultiple24Regular',
    color: '#0078d4',
    category: 'hr',
    requiredLists: [
      {
        name: 'PM_Surveys',
        title: 'Surveys',
        description: 'Survey definitions and questions',
        provisioningScript: 'Create-PM_Surveys-List.ps1'
      },
      {
        name: 'PM_SurveyResponses',
        title: 'Survey Responses',
        description: 'Individual survey responses',
        provisioningScript: 'Create-PM_SurveyResponses-List.ps1'
      },
      {
        name: 'PM_SurveyTemplates',
        title: 'Survey Templates',
        description: 'Reusable survey templates',
        provisioningScript: 'Create-PM_SurveyTemplates-List.ps1'
      }
    ]
  },

  [PremiumModule.CVManager]: {
    id: PremiumModule.CVManager,
    name: 'CV Manager',
    tagline: 'Skills that scale',
    description: 'Build a dynamic skills inventory, track certifications, and identify talent gaps.',
    tier: LicenseTier.Enterprise,
    icon: 'PersonCircle24Regular',
    color: '#8764b8',
    category: 'hr',
    requiredLists: [
      {
        name: 'PM_CVDatabase',
        title: 'CV Database',
        description: 'Employee CV and profile information',
        provisioningScript: 'Create-PM_CVDatabase-List.ps1'
      },
      {
        name: 'PM_Skills',
        title: 'Skills Catalog',
        description: 'Master list of skills',
        provisioningScript: 'Create-PM_Skills-List.ps1'
      },
      {
        name: 'PM_UserSkills',
        title: 'User Skills',
        description: 'Employee skill assignments and proficiency levels',
        provisioningScript: 'Create-PM_UserSkills-List.ps1'
      }
    ]
  },

  [PremiumModule.TalentManager]: {
    id: PremiumModule.TalentManager,
    name: 'Talent Manager',
    tagline: 'Recruit smarter. Onboard faster.',
    description: 'Streamline your entire recruitment pipeline from candidate sourcing to onboarding.',
    tier: LicenseTier.Enterprise,
    icon: 'DocumentBulletList24Regular',
    color: '#00ad56',
    category: 'hr',
    requiredLists: [
      {
        name: 'PM_Candidates',
        title: 'Candidates',
        description: 'Candidate profiles and applications',
        provisioningScript: 'Create-PM_Candidates-List.ps1'
      },
      {
        name: 'PM_JobRequisitions',
        title: 'Job Requisitions',
        description: 'Open positions and job postings',
        provisioningScript: 'Create-PM_JobRequisitions-List.ps1'
      },
      {
        name: 'PM_Interviews',
        title: 'Interviews',
        description: 'Interview scheduling and tracking',
        provisioningScript: 'Create-PM_Interviews-List.ps1'
      },
      {
        name: 'PM_JobOffers',
        title: 'Job Offers',
        description: 'Offer letters and acceptance tracking',
        provisioningScript: 'Create-PM_JobOffers-List.ps1'
      }
    ]
  },

  [PremiumModule.SkillsBuilder]: {
    id: PremiumModule.SkillsBuilder,
    name: 'Skills Builder',
    tagline: 'Grow your talent',
    description: 'Comprehensive learning management with personalized paths, skill assessments, and certification tracking.',
    tier: LicenseTier.Enterprise,
    icon: 'Hat24Regular',
    color: '#c239b3',
    category: 'hr',
    isNew: true,
    requiredLists: [
      {
        name: 'PM_TrainingCatalog',
        title: 'Training Catalog',
        description: 'Available training courses and programs',
        provisioningScript: 'Create-PM_TrainingCatalog-List.ps1'
      },
      {
        name: 'PM_LearningPaths',
        title: 'Learning Paths',
        description: 'Structured learning journeys',
        provisioningScript: 'Create-PM_LearningPaths-List.ps1'
      },
      {
        name: 'PM_TrainingEnrollments',
        title: 'Training Enrollments',
        description: 'Employee course enrollments and progress',
        provisioningScript: 'Create-PM_TrainingEnrollments-List.ps1'
      },
      {
        name: 'PM_Certifications',
        title: 'Certifications',
        description: 'Certification definitions',
        provisioningScript: 'Create-PM_Certifications-List.ps1'
      },
      {
        name: 'PM_UserCertifications',
        title: 'User Certifications',
        description: 'Employee certification records',
        provisioningScript: 'Create-PM_UserCertifications-List.ps1'
      }
    ],
    dependencies: [
      { moduleId: PremiumModule.CVManager, reason: 'Integrates with CV skill profiles' }
    ]
  },

  // ============================================
  // OPERATIONS MODULES
  // ============================================

  [PremiumModule.DocumentGeneration]: {
    id: PremiumModule.DocumentGeneration,
    name: 'Document Generation',
    tagline: 'Documents made easy',
    description: 'Generate professional documents automatically with template-driven workflows and digital signatures.',
    tier: LicenseTier.Starter,
    icon: 'DocumentPdf24Regular',
    color: '#ff8c00',
    category: 'operations',
    requiredLists: [
      {
        name: 'PM_DocumentTemplates',
        title: 'Document Templates',
        description: 'Document generation templates',
        provisioningScript: 'Create-PM_DocumentTemplates-Library.ps1',
        isDocumentLibrary: true
      },
      {
        name: 'PM_SignatureRequests',
        title: 'Signature Requests',
        description: 'Digital signature workflow tracking',
        provisioningScript: 'Create-PM_SignatureRequests-List.ps1'
      }
    ]
  },

  [PremiumModule.AssetManager]: {
    id: PremiumModule.AssetManager,
    name: 'Asset Manager',
    tagline: 'Track it all',
    description: 'Manage IT assets, equipment, and licenses throughout the employee lifecycle.',
    tier: LicenseTier.Professional,
    icon: 'Cube24Regular',
    color: '#5c2d91',
    category: 'operations',
    requiredLists: [
      // Uses core PM_Assets and PM_AssetCheckouts lists
      {
        name: 'PM_Asset_Configuration',
        title: 'Asset Configuration',
        description: 'Asset module settings and categories',
        provisioningScript: 'Create-PM_Asset_Configuration-List.ps1'
      }
    ]
  },

  [PremiumModule.ProcurementManager]: {
    id: PremiumModule.ProcurementManager,
    name: 'Procurement Manager',
    tagline: 'Streamline purchasing',
    description: 'End-to-end procurement workflows for equipment, software, and services with approval chains.',
    tier: LicenseTier.Enterprise,
    icon: 'Cart24Regular',
    color: '#744da9',
    category: 'operations',
    isNew: true,
    requiredLists: [
      {
        name: 'PM_Vendors',
        title: 'Vendors',
        description: 'Vendor/supplier directory',
        provisioningScript: 'Create-PM_Vendors-List.ps1'
      },
      {
        name: 'PM_Requisitions',
        title: 'Requisitions',
        description: 'Purchase requisitions',
        provisioningScript: 'Create-PM_Requisitions-List.ps1'
      },
      {
        name: 'PM_PurchaseOrders',
        title: 'Purchase Orders',
        description: 'Approved purchase orders',
        provisioningScript: 'Create-PM_PurchaseOrders-List.ps1'
      }
    ],
    dependencies: [
      { moduleId: PremiumModule.AssetManager, reason: 'Links purchases to asset inventory' }
    ]
  },

  [PremiumModule.SLAManager]: {
    id: PremiumModule.SLAManager,
    name: 'SLA Manager',
    tagline: 'Deliver on time. Every time.',
    description: 'Track SLA compliance, automate escalations, and ensure timely completion of critical tasks.',
    tier: LicenseTier.Professional,
    icon: 'ClockAlarm24Regular',
    color: '#ea4300',
    category: 'operations',
    requiredLists: [
      {
        name: 'PM_SLADefinitions',
        title: 'SLA Definitions',
        description: 'SLA rules and thresholds',
        provisioningScript: 'Create-PM_SLADefinitions-List.ps1'
      },
      {
        name: 'PM_EscalationRules',
        title: 'Escalation Rules',
        description: 'Automated escalation configuration',
        provisioningScript: 'Create-PM_EscalationRules-List.ps1'
      }
    ]
  },

  [PremiumModule.ContractManager]: {
    id: PremiumModule.ContractManager,
    name: 'Contract Manager',
    tagline: 'Contracts made simple',
    description: 'Full lifecycle contract management with clause library, approvals, obligation tracking, e-signatures, and complete audit trail.',
    tier: LicenseTier.Enterprise,
    icon: 'DocumentContract24Regular',
    color: '#744da9',
    category: 'operations',
    isNew: true,
    isPopular: true,
    requiredLists: [
      {
        name: 'PM_Contracts',
        title: 'Contracts',
        description: 'Contract records and metadata',
        provisioningScript: 'Deploy-ContractManager-Lists.ps1'
      },
      {
        name: 'PM_ContractParties',
        title: 'Contract Parties',
        description: 'Parties involved in contracts',
        provisioningScript: 'Deploy-ContractManager-Lists.ps1'
      },
      {
        name: 'PM_ContractClauses',
        title: 'Contract Clauses',
        description: 'Clauses assigned to contracts',
        provisioningScript: 'Deploy-ContractManager-Lists.ps1'
      },
      {
        name: 'PM_ClauseLibrary',
        title: 'Clause Library',
        description: 'Master clause library (200+ standard clauses)',
        provisioningScript: 'Deploy-ContractManager-Lists.ps1'
      },
      {
        name: 'PM_ContractObligations',
        title: 'Contract Obligations',
        description: 'Trackable obligations and deadlines',
        provisioningScript: 'Deploy-ContractManager-Lists.ps1'
      },
      {
        name: 'PM_ContractApprovals',
        title: 'Contract Approvals',
        description: 'Approval workflow records',
        provisioningScript: 'Deploy-ContractManager-Lists.ps1'
      },
      {
        name: 'PM_ContractVersions',
        title: 'Contract Versions',
        description: 'Contract version history',
        provisioningScript: 'Deploy-ContractManager-Lists.ps1'
      },
      {
        name: 'PM_ContractDocuments',
        title: 'Contract Documents',
        description: 'Document library for contract attachments',
        provisioningScript: 'Deploy-ContractManager-Lists.ps1',
        isDocumentLibrary: true
      },
      {
        name: 'PM_ContractAuditLog',
        title: 'Contract Audit Log',
        description: 'Complete audit trail of all contract activities',
        provisioningScript: 'Deploy-ContractManager-Lists.ps1'
      },
      {
        name: 'PM_ContractTemplates',
        title: 'Contract Templates',
        description: '70+ pre-built contract templates',
        provisioningScript: 'Deploy-ContractManager-Lists.ps1'
      }
    ],
    dependencies: [
      { moduleId: PremiumModule.ProcurementManager, reason: 'Links contracts to procurement workflows' }
    ]
  },

  [PremiumModule.EmailManager]: {
    id: PremiumModule.EmailManager,
    name: 'Email Manager',
    tagline: 'Communicate with confidence',
    description: 'Design beautiful email templates, automate communications, and maintain consistent messaging.',
    tier: LicenseTier.Starter,
    icon: 'Mail24Regular',
    color: '#0099bc',
    category: 'operations',
    requiredLists: [
      {
        name: 'PM_EmailTemplates',
        title: 'Email Templates',
        description: 'Email template definitions',
        provisioningScript: 'Create-PM_EmailTemplates-List.ps1'
      },
      {
        name: 'PM_EmailQueue',
        title: 'Email Queue',
        description: 'Outbound email queue',
        provisioningScript: 'Create-PM_EmailQueue-List.ps1'
      }
    ]
  },

  // ============================================
  // INTEGRATION MODULES
  // ============================================

  [PremiumModule.IntegrationHub]: {
    id: PremiumModule.IntegrationHub,
    name: 'Integration Hub',
    tagline: 'Connect everything',
    description: 'Unify your HR ecosystem with seamless integrations to HRIS, ITSM, payroll, and third-party systems.',
    tier: LicenseTier.Enterprise,
    icon: 'PlugConnected24Regular',
    color: '#107c10',
    category: 'integration',
    requiredLists: [
      {
        name: 'PM_IntegrationConfigs',
        title: 'Integration Configs',
        description: 'Integration connection settings',
        provisioningScript: 'Create-PM_IntegrationConfigs-List.ps1'
      },
      {
        name: 'PM_IntegrationLogs',
        title: 'Integration Logs',
        description: 'Integration operation audit trail',
        provisioningScript: 'Create-PM_IntegrationLogs-List.ps1'
      },
      {
        name: 'PM_IntegrationMappings',
        title: 'Integration Mappings',
        description: 'Field and data mappings between systems',
        provisioningScript: 'Create-PM_IntegrationMappings-List.ps1'
      },
      {
        name: 'PM_WebhookConfigs',
        title: 'Webhook Configs',
        description: 'Webhook endpoint configuration',
        provisioningScript: 'Create-PM_WebhookConfigs-List.ps1'
      }
    ]
  },

  // ============================================
  // COMPLIANCE MODULES
  // ============================================

  [PremiumModule.ComplianceManager]: {
    id: PremiumModule.ComplianceManager,
    name: 'Compliance Manager',
    tagline: 'Stay compliant. Stay secure.',
    description: 'Ensure regulatory compliance with automated policy enforcement, audit trails, and governance controls.',
    tier: LicenseTier.Enterprise,
    icon: 'ShieldTask24Regular',
    color: '#0b6a0b',
    category: 'compliance',
    requiredLists: [
      {
        name: 'PM_DataRetentionPolicies',
        title: 'Data Retention Policies',
        description: 'Data retention rules and schedules',
        provisioningScript: 'Create-PM_DataRetentionPolicies-List.ps1'
      },
      {
        name: 'PM_ConsentRecords',
        title: 'Consent Records',
        description: 'GDPR consent tracking',
        provisioningScript: 'Create-PM_ConsentRecords-List.ps1'
      },
      {
        name: 'PM_DataSubjectRequests',
        title: 'Data Subject Requests',
        description: 'DSAR tracking (access, deletion, portability)',
        provisioningScript: 'Create-PM_DataSubjectRequests-List.ps1'
      }
    ]
  },

  // ============================================
  // ADVANCED MODULES
  // ============================================

  [PremiumModule.AIAssistant]: {
    id: PremiumModule.AIAssistant,
    name: 'AI Assistant',
    tagline: 'Intelligence amplified',
    description: 'Leverage AI-powered recommendations, chatbots, and automation to accelerate JML processes.',
    tier: LicenseTier.Enterprise,
    icon: 'Sparkle24Regular',
    color: '#e3008c',
    category: 'advanced',
    isNew: true,
    isPopular: true,
    requiredLists: [
      {
        name: 'PM_AIUsageLogs',
        title: 'AI Usage Logs',
        description: 'AI feature usage and token tracking',
        provisioningScript: 'Create-PM_AIUsageLogs-List.ps1'
      },
      {
        name: 'PM_AI_Configs',
        title: 'AI Configurations',
        description: 'AI model settings and prompts',
        provisioningScript: 'Create-PM_AI_Configs-List.ps1'
      }
    ]
  },

  [PremiumModule.ThemeManager]: {
    id: PremiumModule.ThemeManager,
    name: 'Theme Manager',
    tagline: 'Your brand. Your way.',
    description: 'Create consistent, on-brand experiences with powerful theming tools.',
    tier: LicenseTier.Starter,
    icon: 'PaintBrush24Regular',
    color: '#d13438',
    category: 'advanced',
    requiredLists: [
      {
        name: 'PM_DepartmentBranding',
        title: 'Department Branding',
        description: 'Department-specific theme overrides',
        provisioningScript: 'Create-PM_DepartmentBranding-List.ps1'
      }
    ]
  },

  // ============================================
  // DASHBOARD MODULE
  // ============================================

  [PremiumModule.Dashboard]: {
    id: PremiumModule.Dashboard,
    name: 'Executive Dashboard',
    tagline: 'See the big picture',
    description: 'Comprehensive executive dashboard with KPIs, trends, and organizational insights.',
    tier: LicenseTier.Professional,
    icon: 'Board24Regular',
    color: '#0078d4',
    category: 'analytics',
    isPopular: true,
    requiredLists: []
  },

  // ============================================
  // POLICY MANAGEMENT MODULE
  // ============================================

  [PremiumModule.PolicyManagement]: {
    id: PremiumModule.PolicyManagement,
    name: 'Policy Management',
    tagline: 'Policies made simple',
    description: 'Complete policy lifecycle management with authoring, approval workflows, version control, and compliance tracking.',
    tier: LicenseTier.Professional,
    icon: 'DocumentText24Regular',
    color: '#0b6a0b',
    category: 'compliance',
    isPopular: true,
    requiredLists: [
      {
        name: 'PM_Policies',
        title: 'Policies',
        description: 'Policy definitions and metadata',
        provisioningScript: 'Deploy-PolicyManagement-Lists.ps1'
      },
      {
        name: 'PM_PolicyVersions',
        title: 'Policy Versions',
        description: 'Version history for policies',
        provisioningScript: 'Deploy-PolicyManagement-Lists.ps1'
      },
      {
        name: 'PM_PolicyAcknowledgements',
        title: 'Policy Acknowledgements',
        description: 'Employee policy acknowledgement records',
        provisioningScript: 'Deploy-PolicyManagement-Lists.ps1'
      },
      {
        name: 'PM_PolicyPacks',
        title: 'Policy Packs',
        description: 'Grouped policy bundles',
        provisioningScript: 'Deploy-PolicyManagement-Lists.ps1'
      },
      {
        name: 'PM_PolicyCategories',
        title: 'Policy Categories',
        description: 'Policy category taxonomy',
        provisioningScript: 'Deploy-PolicyManagement-Lists.ps1'
      }
    ]
  },

  // ============================================
  // SIGNING SERVICE MODULE
  // ============================================

  [PremiumModule.SigningService]: {
    id: PremiumModule.SigningService,
    name: 'Signing Service',
    tagline: 'Sign with confidence',
    description: 'Digital document signing with multi-party signatures, audit trails, and compliance.',
    tier: LicenseTier.Starter,
    icon: 'Signature24Regular',
    color: '#744da9',
    category: 'operations',
    requiredLists: [
      {
        name: 'PM_SignatureRequests',
        title: 'Signature Requests',
        description: 'Digital signature request tracking',
        provisioningScript: 'Create-PM_SignatureRequests-List.ps1'
      },
      {
        name: 'PM_SignatureAuditLog',
        title: 'Signature Audit Log',
        description: 'Complete audit trail for all signing activities',
        provisioningScript: 'Create-PM_SignatureAuditLog-List.ps1'
      }
    ]
  },

  // ============================================
  // LICENSE MANAGEMENT MODULE
  // ============================================

  [PremiumModule.LicenseManagement]: {
    id: PremiumModule.LicenseManagement,
    name: 'License Management',
    tagline: 'Track every license',
    description: 'Track software licenses, renewals, and compliance across your organization.',
    tier: LicenseTier.Professional,
    icon: 'Key24Regular',
    color: '#5c2d91',
    category: 'operations',
    requiredLists: [
      {
        name: 'PM_SoftwareLicenses',
        title: 'Software Licenses',
        description: 'Software license inventory',
        provisioningScript: 'Create-PM_SoftwareLicenses-List.ps1'
      },
      {
        name: 'PM_LicenseAssignments',
        title: 'License Assignments',
        description: 'User license assignments',
        provisioningScript: 'Create-PM_LicenseAssignments-List.ps1'
      }
    ]
  },

  // ============================================
  // GAMIFICATION MODULE
  // ============================================

  [PremiumModule.Gamification]: {
    id: PremiumModule.Gamification,
    name: 'Gamification',
    tagline: 'Engage and motivate',
    description: 'Boost engagement with points, badges, leaderboards, and achievement systems.',
    tier: LicenseTier.Professional,
    icon: 'Trophy24Regular',
    color: '#ffaa00',
    category: 'hr',
    isNew: true,
    isPopular: true,
    requiredLists: [
      {
        name: 'PM_GamificationPoints',
        title: 'Gamification Points',
        description: 'User points and scores',
        provisioningScript: 'Create-PM_Gamification-Lists.ps1'
      },
      {
        name: 'PM_GamificationBadges',
        title: 'Gamification Badges',
        description: 'Badge definitions',
        provisioningScript: 'Create-PM_Gamification-Lists.ps1'
      },
      {
        name: 'PM_GamificationAchievements',
        title: 'Gamification Achievements',
        description: 'User badge awards and achievements',
        provisioningScript: 'Create-PM_Gamification-Lists.ps1'
      },
      {
        name: 'PM_GamificationLeaderboards',
        title: 'Gamification Leaderboards',
        description: 'Leaderboard configurations',
        provisioningScript: 'Create-PM_Gamification-Lists.ps1'
      }
    ]
  },

  // ============================================
  // QUIZ BUILDER MODULE
  // ============================================

  [PremiumModule.QuizBuilder]: {
    id: PremiumModule.QuizBuilder,
    name: 'Quiz Builder',
    tagline: 'Test and learn',
    description: 'Create interactive quizzes for training, assessments, and knowledge checks.',
    tier: LicenseTier.Professional,
    icon: 'Question24Regular',
    color: '#e3008c',
    category: 'hr',
    isNew: true,
    requiredLists: [
      {
        name: 'PM_Quizzes',
        title: 'Quizzes',
        description: 'Quiz definitions',
        provisioningScript: 'Create-PM_Quizzes-List.ps1'
      },
      {
        name: 'PM_QuizQuestions',
        title: 'Quiz Questions',
        description: 'Quiz question bank',
        provisioningScript: 'Create-PM_QuizQuestions-List.ps1'
      },
      {
        name: 'PM_QuizAttempts',
        title: 'Quiz Attempts',
        description: 'User quiz attempts and scores',
        provisioningScript: 'Create-PM_QuizAttempts-List.ps1'
      }
    ],
    dependencies: [
      { moduleId: PremiumModule.SkillsBuilder, reason: 'Integrates with training programs' }
    ]
  },

  // ============================================
  // FINANCIAL MANAGEMENT MODULE
  // ============================================

  [PremiumModule.FinancialManagement]: {
    id: PremiumModule.FinancialManagement,
    name: 'Financial Management',
    tagline: 'Track every dollar',
    description: 'Budget tracking, expense management, and payroll integration for JML processes.',
    tier: LicenseTier.Enterprise,
    icon: 'Money24Regular',
    color: '#107c10',
    category: 'operations',
    isNew: true,
    requiredLists: [
      {
        name: 'PM_Budgets',
        title: 'Budgets',
        description: 'Department and project budgets',
        provisioningScript: 'Create-FinanceLists.ps1'
      },
      {
        name: 'PM_Expenses',
        title: 'Expenses',
        description: 'Expense tracking',
        provisioningScript: 'Create-FinanceLists-Expenses.ps1'
      },
      {
        name: 'PM_PayrollSummary',
        title: 'Payroll Summary',
        description: 'Payroll integration data',
        provisioningScript: 'Create-FinanceLists-PayrollSummary.ps1'
      }
    ]
  },

  // ============================================
  // CORE FEATURES (Included in Base Package)
  // ============================================

  [PremiumModule.DocumentHub]: {
    id: PremiumModule.DocumentHub,
    name: 'Document Hub',
    tagline: 'Centralized document management',
    description: 'Enterprise document management with versioning, retention policies, search, and compliance controls. Available to all users.',
    tier: LicenseTier.Free,
    icon: 'DocumentSet24Regular',
    color: '#038387',
    category: 'operations',
    isNew: true,
    requiredLists: [
      {
        name: 'PM_Documents',
        title: 'Documents',
        description: 'Central document library',
        provisioningScript: 'Create-PM_Documents-Library.ps1',
        isDocumentLibrary: true
      },
      {
        name: 'PM_DocumentCategories',
        title: 'Document Categories',
        description: 'Document category taxonomy',
        provisioningScript: 'Create-PM_DocumentCategories-List.ps1'
      }
    ]
  },

  [PremiumModule.ExternalSharingHub]: {
    id: PremiumModule.ExternalSharingHub,
    name: 'External Sharing Hub',
    tagline: 'Secure external collaboration',
    description: 'Manage external sharing, guest access, and cross-tenant collaboration with full governance controls. IT Admin tool.',
    tier: LicenseTier.Free,
    icon: 'Share24Regular',
    color: '#d13438',
    category: 'compliance',
    isNew: true,
    requiredLists: [
      {
        name: 'PM_ExternalShares',
        title: 'External Shares',
        description: 'Tracking of external sharing activities',
        provisioningScript: 'Create-PM_ExternalShares-List.ps1'
      },
      {
        name: 'PM_GuestUsers',
        title: 'Guest Users',
        description: 'Guest user directory',
        provisioningScript: 'Create-PM_GuestUsers-List.ps1'
      },
      {
        name: 'PM_TrustedOrganizations',
        title: 'Trusted Organizations',
        description: 'Cross-tenant trust relationships',
        provisioningScript: 'Create-PM_TrustedOrganizations-List.ps1'
      }
    ]
  },

  [PremiumModule.WorkflowMonitor]: {
    id: PremiumModule.WorkflowMonitor,
    name: 'Workflow Monitor',
    tagline: 'Real-time workflow visibility',
    description: 'Monitor active workflows, track SLA compliance, and intervene on stuck processes. IT Admin tool for workflow operations.',
    tier: LicenseTier.Free,
    icon: 'Flow24Regular',
    color: '#5c2d91',
    category: 'operations',
    isNew: true,
    requiredLists: [
      {
        name: 'PM_WorkflowInstances',
        title: 'Workflow Instances',
        description: 'Active workflow instance tracking',
        provisioningScript: 'Create-PM_WorkflowInstances-List.ps1'
      },
      {
        name: 'PM_WorkflowSteps',
        title: 'Workflow Steps',
        description: 'Workflow step execution log',
        provisioningScript: 'Create-PM_WorkflowSteps-List.ps1'
      }
    ]
  }
};

/**
 * Get all modules for a specific tier
 */
export function getModulesForTier(tier: LicenseTier): IModuleDefinition[] {
  return Object.values(ModuleRegistry).filter(m => {
    const tierOrder = [LicenseTier.Free, LicenseTier.Starter, LicenseTier.Professional, LicenseTier.Enterprise];
    return tierOrder.indexOf(m.tier) <= tierOrder.indexOf(tier);
  });
}

/**
 * Get all required lists for a set of modules
 */
export function getRequiredListsForModules(moduleIds: PremiumModule[]): IListDefinition[] {
  const lists: IListDefinition[] = [];
  const seen = new Set<string>();

  for (const moduleId of moduleIds) {
    const module = ModuleRegistry[moduleId];
    if (module) {
      for (const list of module.requiredLists) {
        if (!seen.has(list.name)) {
          seen.add(list.name);
          lists.push(list);
        }
      }
    }
  }

  return lists;
}

/**
 * Get modules grouped by category
 */
export function getModulesByCategory(): Record<string, IModuleDefinition[]> {
  const categories: Record<string, IModuleDefinition[]> = {};

  for (const module of Object.values(ModuleRegistry)) {
    if (!categories[module.category]) {
      categories[module.category] = [];
    }
    categories[module.category].push(module);
  }

  return categories;
}

/**
 * Category display names
 */
export const CategoryDisplayNames: Record<string, string> = {
  analytics: 'Analytics & Reporting',
  hr: 'HR & Talent Management',
  operations: 'Operations & Workflow',
  integration: 'Integration & APIs',
  compliance: 'Compliance & Security',
  advanced: 'Advanced Features'
};
