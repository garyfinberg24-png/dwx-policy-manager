// @ts-nocheck
/**
 * Policy Template Library Service
 * Provides comprehensive template management for policy creation
 * Includes pre-built templates, categories, customization, and sharing
 */

import { SPFI, spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { logger } from './LoggingService';
import { TemplateLibraryLists } from '../constants/SharePointListNames';

/**
 * Template category enumeration
 */
export enum TemplateCategory {
  HumanResources = 'Human Resources',
  InformationSecurity = 'Information Security',
  DataPrivacy = 'Data Privacy',
  HealthAndSafety = 'Health & Safety',
  FinanceCompliance = 'Finance & Compliance',
  OperationalProcedures = 'Operational Procedures',
  ITGovernance = 'IT Governance',
  LegalCompliance = 'Legal & Compliance',
  CustomerService = 'Customer Service',
  EnvironmentalSocial = 'Environmental & Social',
  Custom = 'Custom'
}

/**
 * Template complexity level
 */
export enum TemplateComplexity {
  Basic = 'Basic',
  Standard = 'Standard',
  Advanced = 'Advanced',
  Enterprise = 'Enterprise'
}

/**
 * Template industry focus
 */
export enum TemplateIndustry {
  General = 'General',
  Healthcare = 'Healthcare',
  FinancialServices = 'Financial Services',
  Technology = 'Technology',
  Manufacturing = 'Manufacturing',
  Retail = 'Retail',
  Education = 'Education',
  Government = 'Government',
  NonProfit = 'Non-Profit',
  Legal = 'Legal'
}

/**
 * Template section type
 */
export enum TemplateSectionType {
  Header = 'Header',
  Purpose = 'Purpose',
  Scope = 'Scope',
  Definitions = 'Definitions',
  Responsibilities = 'Responsibilities',
  Policy = 'Policy',
  Procedure = 'Procedure',
  Compliance = 'Compliance',
  Enforcement = 'Enforcement',
  References = 'References',
  Appendix = 'Appendix',
  Custom = 'Custom'
}

/**
 * Template section interface
 */
export interface ITemplateSection {
  id: string;
  type: TemplateSectionType;
  title: string;
  content: string;
  placeholder?: string;
  required: boolean;
  order: number;
  guidance?: string;
  examples?: string[];
}

/**
 * Policy template interface
 */
export interface IPolicyTemplate {
  id: number;
  title: string;
  description: string;
  category: TemplateCategory;
  industry: TemplateIndustry;
  complexity: TemplateComplexity;
  sections: ITemplateSection[];
  tags: string[];
  version: string;
  author: string;
  createdDate: Date;
  modifiedDate: Date;
  usageCount: number;
  rating: number;
  ratingCount: number;
  isPublic: boolean;
  isFeatured: boolean;
  previewImage?: string;
  estimatedTime?: string;
  regulatoryFrameworks?: string[];
}

/**
 * Template search criteria
 */
export interface ITemplateSearchCriteria {
  searchText?: string;
  categories?: TemplateCategory[];
  industries?: TemplateIndustry[];
  complexity?: TemplateComplexity[];
  tags?: string[];
  isFeatured?: boolean;
  minRating?: number;
  regulatoryFramework?: string;
  sortBy?: 'title' | 'usageCount' | 'rating' | 'modifiedDate';
  sortDirection?: 'asc' | 'desc';
  pageSize?: number;
  pageIndex?: number;
}

/**
 * Template search results
 */
export interface ITemplateSearchResult {
  templates: IPolicyTemplate[];
  totalCount: number;
  pageIndex: number;
  pageSize: number;
  hasMore: boolean;
  facets: ITemplateFacets;
}

/**
 * Template facets for filtering
 */
export interface ITemplateFacets {
  categories: { category: TemplateCategory; count: number }[];
  industries: { industry: TemplateIndustry; count: number }[];
  complexity: { complexity: TemplateComplexity; count: number }[];
  tags: { tag: string; count: number }[];
}

/**
 * Template customization options
 */
export interface ITemplateCustomization {
  templateId: number;
  newTitle: string;
  newDescription?: string;
  selectedSections: string[];
  sectionOverrides?: Record<string, Partial<ITemplateSection>>;
  customSections?: ITemplateSection[];
  metadata?: Record<string, string>;
}

/**
 * User template preferences
 */
export interface IUserTemplatePreferences {
  favoriteTemplates: number[];
  recentTemplates: number[];
  preferredCategories: TemplateCategory[];
  preferredIndustry: TemplateIndustry;
}

/**
 * Template usage record
 */
export interface ITemplateUsageRecord {
  templateId: number;
  userId: string;
  usedDate: Date;
  policyId?: number;
  customizations?: string[];
}

/**
 * Built-in template definitions
 */
const BUILT_IN_TEMPLATES: Partial<IPolicyTemplate>[] = [
  // Human Resources Templates
  {
    title: 'Employee Code of Conduct',
    description: 'Comprehensive code of conduct policy covering professional behavior, ethics, and workplace standards.',
    category: TemplateCategory.HumanResources,
    industry: TemplateIndustry.General,
    complexity: TemplateComplexity.Standard,
    tags: ['conduct', 'ethics', 'behavior', 'workplace'],
    estimatedTime: '2-3 hours',
    sections: [
      {
        id: 'coc-purpose',
        type: TemplateSectionType.Purpose,
        title: 'Purpose',
        content: 'This Code of Conduct establishes the principles and standards of behavior expected from all employees of [Company Name]. It provides guidance on ethical decision-making and professional conduct.',
        placeholder: '[Company Name]',
        required: true,
        order: 1,
        guidance: 'Clearly state why this policy exists and its importance to the organization.'
      },
      {
        id: 'coc-scope',
        type: TemplateSectionType.Scope,
        title: 'Scope',
        content: 'This policy applies to all employees, contractors, consultants, and temporary staff, regardless of position or tenure. It covers behavior during work hours, at company events, and when representing the company externally.',
        required: true,
        order: 2,
        guidance: 'Define who is covered by this policy and in what circumstances.'
      },
      {
        id: 'coc-professional-conduct',
        type: TemplateSectionType.Policy,
        title: 'Professional Conduct Standards',
        content: `All employees are expected to:\n\n1. **Integrity**: Act honestly and ethically in all business dealings\n2. **Respect**: Treat colleagues, clients, and partners with dignity and respect\n3. **Professionalism**: Maintain professional appearance and communication\n4. **Confidentiality**: Protect confidential information and trade secrets\n5. **Compliance**: Follow all applicable laws, regulations, and company policies`,
        required: true,
        order: 3,
        examples: [
          'Arriving on time for meetings',
          'Responding to communications promptly',
          'Dressing appropriately for the work environment'
        ]
      },
      {
        id: 'coc-conflicts',
        type: TemplateSectionType.Policy,
        title: 'Conflicts of Interest',
        content: 'Employees must avoid situations where personal interests conflict with company interests. All potential conflicts must be disclosed to management immediately.',
        required: true,
        order: 4
      },
      {
        id: 'coc-enforcement',
        type: TemplateSectionType.Enforcement,
        title: 'Enforcement & Consequences',
        content: 'Violations of this Code of Conduct may result in disciplinary action, up to and including termination of employment. Severity of consequences will be proportionate to the violation.',
        required: true,
        order: 5
      }
    ]
  },
  {
    title: 'Remote Work Policy',
    description: 'Policy governing remote and hybrid work arrangements, including eligibility, expectations, and technology requirements.',
    category: TemplateCategory.HumanResources,
    industry: TemplateIndustry.General,
    complexity: TemplateComplexity.Standard,
    tags: ['remote work', 'hybrid', 'work from home', 'flexible work'],
    estimatedTime: '2-3 hours',
    sections: [
      {
        id: 'rw-purpose',
        type: TemplateSectionType.Purpose,
        title: 'Purpose',
        content: 'This policy outlines the guidelines and expectations for employees who work remotely or in a hybrid arrangement. It aims to balance flexibility with productivity and collaboration.',
        required: true,
        order: 1
      },
      {
        id: 'rw-eligibility',
        type: TemplateSectionType.Policy,
        title: 'Eligibility Criteria',
        content: `Remote work eligibility is determined by:\n\n1. **Role Suitability**: The job functions can be performed remotely\n2. **Performance History**: Employee demonstrates consistent performance\n3. **Manager Approval**: Direct supervisor approves the arrangement\n4. **Technology Access**: Employee has reliable internet and workspace`,
        required: true,
        order: 2
      },
      {
        id: 'rw-expectations',
        type: TemplateSectionType.Policy,
        title: 'Work Expectations',
        content: `Remote employees must:\n\n- Maintain regular working hours as agreed with their manager\n- Be available and responsive during core hours ([Core Hours])\n- Attend required meetings via video conference\n- Maintain a safe and ergonomic workspace\n- Ensure data security and confidentiality`,
        placeholder: '[Core Hours]',
        required: true,
        order: 3
      },
      {
        id: 'rw-technology',
        type: TemplateSectionType.Policy,
        title: 'Technology Requirements',
        content: 'Employees must have:\n\n- Reliable high-speed internet connection\n- Company-approved devices or secure personal devices\n- VPN access for secure connections\n- Collaboration tools (Teams, Slack, etc.)',
        required: true,
        order: 4
      },
      {
        id: 'rw-review',
        type: TemplateSectionType.Policy,
        title: 'Review & Modification',
        content: 'Remote work arrangements will be reviewed quarterly. The company reserves the right to modify or revoke remote work privileges based on business needs or performance concerns.',
        required: true,
        order: 5
      }
    ]
  },
  {
    title: 'Anti-Harassment Policy',
    description: 'Comprehensive policy prohibiting harassment and discrimination in the workplace, with reporting procedures and investigation protocols.',
    category: TemplateCategory.HumanResources,
    industry: TemplateIndustry.General,
    complexity: TemplateComplexity.Standard,
    tags: ['harassment', 'discrimination', 'workplace safety', 'HR'],
    estimatedTime: '3-4 hours',
    regulatoryFrameworks: ['EEOC', 'Title VII'],
    sections: [
      {
        id: 'ah-purpose',
        type: TemplateSectionType.Purpose,
        title: 'Purpose & Commitment',
        content: '[Company Name] is committed to providing a work environment free from harassment, discrimination, and retaliation. This policy prohibits all forms of harassment and outlines procedures for reporting and addressing complaints.',
        placeholder: '[Company Name]',
        required: true,
        order: 1
      },
      {
        id: 'ah-definitions',
        type: TemplateSectionType.Definitions,
        title: 'Definitions',
        content: `**Harassment**: Unwelcome conduct based on protected characteristics that creates a hostile work environment.\n\n**Sexual Harassment**: Unwelcome sexual advances, requests for sexual favors, or verbal/physical conduct of a sexual nature.\n\n**Discrimination**: Unfair treatment based on protected characteristics.\n\n**Retaliation**: Adverse action against someone who reports harassment or participates in an investigation.`,
        required: true,
        order: 2
      },
      {
        id: 'ah-prohibited',
        type: TemplateSectionType.Policy,
        title: 'Prohibited Conduct',
        content: `The following conduct is strictly prohibited:\n\n1. Verbal harassment (slurs, jokes, comments)\n2. Physical harassment (unwanted touching, blocking movement)\n3. Visual harassment (offensive images, gestures)\n4. Sexual harassment of any kind\n5. Bullying or intimidation\n6. Retaliation against reporters or witnesses`,
        required: true,
        order: 3
      },
      {
        id: 'ah-reporting',
        type: TemplateSectionType.Procedure,
        title: 'Reporting Procedures',
        content: `Employees should report harassment to:\n\n1. **Immediate Supervisor** (unless they are involved)\n2. **Human Resources**: [HR Contact]\n3. **Anonymous Hotline**: [Hotline Number]\n\nAll reports will be treated confidentially to the extent possible.`,
        placeholder: '[HR Contact], [Hotline Number]',
        required: true,
        order: 4
      },
      {
        id: 'ah-investigation',
        type: TemplateSectionType.Procedure,
        title: 'Investigation Process',
        content: 'All complaints will be investigated promptly and thoroughly. Investigations will include:\n\n1. Interview with complainant\n2. Interview with accused\n3. Interview with witnesses\n4. Review of relevant documentation\n5. Determination and appropriate action',
        required: true,
        order: 5
      },
      {
        id: 'ah-consequences',
        type: TemplateSectionType.Enforcement,
        title: 'Consequences',
        content: 'Violations may result in disciplinary action up to and including termination. False accusations made in bad faith may also result in disciplinary action.',
        required: true,
        order: 6
      }
    ]
  },

  // Information Security Templates
  {
    title: 'Information Security Policy',
    description: 'Comprehensive information security policy covering data protection, access controls, and security requirements.',
    category: TemplateCategory.InformationSecurity,
    industry: TemplateIndustry.General,
    complexity: TemplateComplexity.Advanced,
    tags: ['security', 'data protection', 'access control', 'IT'],
    estimatedTime: '4-6 hours',
    regulatoryFrameworks: ['ISO 27001', 'NIST'],
    sections: [
      {
        id: 'is-purpose',
        type: TemplateSectionType.Purpose,
        title: 'Purpose',
        content: 'This Information Security Policy establishes the framework for protecting [Company Name]\'s information assets from threats, whether internal or external, deliberate or accidental.',
        placeholder: '[Company Name]',
        required: true,
        order: 1
      },
      {
        id: 'is-scope',
        type: TemplateSectionType.Scope,
        title: 'Scope',
        content: 'This policy applies to all employees, contractors, and third parties who have access to company information systems and data. It covers all forms of information: digital, paper, and verbal.',
        required: true,
        order: 2
      },
      {
        id: 'is-classification',
        type: TemplateSectionType.Policy,
        title: 'Data Classification',
        content: `Information is classified into four categories:\n\n1. **Public**: Information approved for public release\n2. **Internal**: General business information for internal use\n3. **Confidential**: Sensitive business information with restricted access\n4. **Restricted**: Highly sensitive information requiring strict controls`,
        required: true,
        order: 3
      },
      {
        id: 'is-access',
        type: TemplateSectionType.Policy,
        title: 'Access Control',
        content: `Access to information systems must be:\n\n1. **Authorized**: Approved by appropriate management\n2. **Authenticated**: Using strong credentials and MFA\n3. **Logged**: All access is recorded and monitored\n4. **Reviewed**: Regularly audited for appropriateness`,
        required: true,
        order: 4
      },
      {
        id: 'is-passwords',
        type: TemplateSectionType.Policy,
        title: 'Password & Authentication',
        content: 'Passwords must:\n\n- Be at least 12 characters long\n- Contain uppercase, lowercase, numbers, and symbols\n- Be changed every 90 days\n- Not be shared or written down\n\nMulti-factor authentication is required for all critical systems.',
        required: true,
        order: 5
      },
      {
        id: 'is-incident',
        type: TemplateSectionType.Procedure,
        title: 'Incident Response',
        content: 'Security incidents must be reported immediately to the IT Security team at [Security Email]. Employees must not attempt to investigate or remediate incidents themselves.',
        placeholder: '[Security Email]',
        required: true,
        order: 6
      },
      {
        id: 'is-compliance',
        type: TemplateSectionType.Compliance,
        title: 'Compliance & Auditing',
        content: 'Compliance with this policy will be audited annually. Violations may result in disciplinary action and/or legal consequences.',
        required: true,
        order: 7
      }
    ]
  },
  {
    title: 'Acceptable Use Policy',
    description: 'Policy defining acceptable use of company IT resources including computers, email, internet, and mobile devices.',
    category: TemplateCategory.InformationSecurity,
    industry: TemplateIndustry.General,
    complexity: TemplateComplexity.Basic,
    tags: ['acceptable use', 'IT', 'internet', 'email'],
    estimatedTime: '1-2 hours',
    sections: [
      {
        id: 'aup-purpose',
        type: TemplateSectionType.Purpose,
        title: 'Purpose',
        content: 'This policy defines acceptable use of company IT resources and establishes guidelines for responsible use of technology in the workplace.',
        required: true,
        order: 1
      },
      {
        id: 'aup-acceptable',
        type: TemplateSectionType.Policy,
        title: 'Acceptable Use',
        content: 'Company IT resources may be used for:\n\n1. Business-related activities\n2. Limited personal use that does not interfere with work\n3. Professional development and training\n4. Authorized communication',
        required: true,
        order: 2
      },
      {
        id: 'aup-prohibited',
        type: TemplateSectionType.Policy,
        title: 'Prohibited Activities',
        content: `The following activities are prohibited:\n\n1. **Illegal Activities**: Using resources for any unlawful purpose\n2. **Harassment**: Sending offensive or threatening content\n3. **Unauthorized Access**: Attempting to access restricted systems\n4. **Personal Business**: Running personal businesses using company resources\n5. **Malicious Software**: Downloading or installing unauthorized software\n6. **Data Theft**: Copying confidential information without authorization`,
        required: true,
        order: 3
      },
      {
        id: 'aup-email',
        type: TemplateSectionType.Policy,
        title: 'Email & Communication',
        content: 'Email should be:\n\n- Professional and appropriate\n- Used for business purposes primarily\n- Not used for chain letters or spam\n- Treated as company property subject to monitoring',
        required: true,
        order: 4
      },
      {
        id: 'aup-monitoring',
        type: TemplateSectionType.Policy,
        title: 'Monitoring & Privacy',
        content: 'The company reserves the right to monitor all use of IT resources. Users should have no expectation of privacy when using company systems.',
        required: true,
        order: 5
      }
    ]
  },
  {
    title: 'Password Management Policy',
    description: 'Detailed policy on password creation, storage, and management requirements.',
    category: TemplateCategory.InformationSecurity,
    industry: TemplateIndustry.General,
    complexity: TemplateComplexity.Basic,
    tags: ['password', 'authentication', 'security', 'access'],
    estimatedTime: '1 hour',
    regulatoryFrameworks: ['NIST 800-63'],
    sections: [
      {
        id: 'pw-purpose',
        type: TemplateSectionType.Purpose,
        title: 'Purpose',
        content: 'This policy establishes requirements for creating and managing passwords to protect company systems and data from unauthorized access.',
        required: true,
        order: 1
      },
      {
        id: 'pw-requirements',
        type: TemplateSectionType.Policy,
        title: 'Password Requirements',
        content: `All passwords must meet the following criteria:\n\n**Length**: Minimum 12 characters (16+ recommended)\n**Complexity**: Include at least 3 of the following:\n- Uppercase letters (A-Z)\n- Lowercase letters (a-z)\n- Numbers (0-9)\n- Special characters (!@#$%^&*)\n\n**Prohibited**: Passwords must not contain:\n- Username or email\n- Common words or phrases\n- Sequential patterns (123456, qwerty)\n- Personal information (birthdays, names)`,
        required: true,
        order: 2
      },
      {
        id: 'pw-expiration',
        type: TemplateSectionType.Policy,
        title: 'Password Expiration',
        content: 'Passwords must be changed:\n\n- Every 90 days for standard accounts\n- Every 60 days for privileged accounts\n- Immediately if compromise is suspected\n\nPrevious 12 passwords cannot be reused.',
        required: true,
        order: 3
      },
      {
        id: 'pw-storage',
        type: TemplateSectionType.Policy,
        title: 'Password Storage',
        content: 'Passwords must not be:\n\n- Written on paper or sticky notes\n- Stored in unencrypted files\n- Shared via email or chat\n\nApproved password managers: [List Approved Tools]',
        placeholder: '[List Approved Tools]',
        required: true,
        order: 4
      },
      {
        id: 'pw-mfa',
        type: TemplateSectionType.Policy,
        title: 'Multi-Factor Authentication',
        content: 'MFA is required for:\n\n- All remote access\n- Privileged accounts\n- Financial systems\n- Email access\n- VPN connections',
        required: true,
        order: 5
      }
    ]
  },

  // Data Privacy Templates
  {
    title: 'Data Privacy Policy',
    description: 'Comprehensive data privacy policy covering personal data handling, subject rights, and compliance requirements.',
    category: TemplateCategory.DataPrivacy,
    industry: TemplateIndustry.General,
    complexity: TemplateComplexity.Advanced,
    tags: ['privacy', 'GDPR', 'personal data', 'compliance'],
    estimatedTime: '4-6 hours',
    regulatoryFrameworks: ['GDPR', 'CCPA', 'HIPAA'],
    sections: [
      {
        id: 'dp-purpose',
        type: TemplateSectionType.Purpose,
        title: 'Purpose',
        content: 'This policy outlines how [Company Name] collects, uses, stores, and protects personal data in compliance with applicable privacy regulations.',
        placeholder: '[Company Name]',
        required: true,
        order: 1
      },
      {
        id: 'dp-principles',
        type: TemplateSectionType.Policy,
        title: 'Data Protection Principles',
        content: `We adhere to the following principles:\n\n1. **Lawfulness**: Data is processed lawfully and fairly\n2. **Purpose Limitation**: Data is collected for specific, legitimate purposes\n3. **Data Minimization**: Only necessary data is collected\n4. **Accuracy**: Data is kept accurate and up to date\n5. **Storage Limitation**: Data is not kept longer than necessary\n6. **Security**: Data is protected against unauthorized access`,
        required: true,
        order: 2
      },
      {
        id: 'dp-rights',
        type: TemplateSectionType.Policy,
        title: 'Data Subject Rights',
        content: `Individuals have the right to:\n\n1. **Access**: Request copies of their personal data\n2. **Rectification**: Request correction of inaccurate data\n3. **Erasure**: Request deletion of their data\n4. **Portability**: Receive data in a portable format\n5. **Objection**: Object to certain processing activities\n6. **Restriction**: Request limitation of processing`,
        required: true,
        order: 3
      },
      {
        id: 'dp-lawful-basis',
        type: TemplateSectionType.Policy,
        title: 'Lawful Basis for Processing',
        content: 'We process personal data only when we have a lawful basis:\n\n- Consent of the data subject\n- Performance of a contract\n- Legal obligation\n- Vital interests\n- Public interest\n- Legitimate interests',
        required: true,
        order: 4
      },
      {
        id: 'dp-breach',
        type: TemplateSectionType.Procedure,
        title: 'Data Breach Response',
        content: 'In the event of a data breach:\n\n1. Incident is reported to DPO within 24 hours\n2. Assessment of impact is conducted\n3. Supervisory authority notified within 72 hours if required\n4. Affected individuals notified if high risk\n5. Remediation actions implemented',
        required: true,
        order: 5
      },
      {
        id: 'dp-contact',
        type: TemplateSectionType.References,
        title: 'Contact Information',
        content: 'Data Protection Officer: [DPO Name]\nEmail: [DPO Email]\nPhone: [DPO Phone]\n\nFor data subject requests, contact: [Privacy Email]',
        placeholder: '[DPO Name], [DPO Email], [DPO Phone], [Privacy Email]',
        required: true,
        order: 6
      }
    ]
  },
  {
    title: 'Data Retention Policy',
    description: 'Policy defining data retention periods, archival, and secure destruction requirements.',
    category: TemplateCategory.DataPrivacy,
    industry: TemplateIndustry.General,
    complexity: TemplateComplexity.Standard,
    tags: ['retention', 'archival', 'destruction', 'records management'],
    estimatedTime: '2-3 hours',
    regulatoryFrameworks: ['GDPR', 'SOX'],
    sections: [
      {
        id: 'dr-purpose',
        type: TemplateSectionType.Purpose,
        title: 'Purpose',
        content: 'This policy defines how long different types of data and records are retained, and the procedures for secure archival and destruction.',
        required: true,
        order: 1
      },
      {
        id: 'dr-schedule',
        type: TemplateSectionType.Policy,
        title: 'Retention Schedule',
        content: `**HR Records**:\n- Employee files: 7 years after termination\n- Payroll records: 7 years\n- Recruitment records: 3 years\n\n**Financial Records**:\n- Tax records: 7 years\n- Invoices: 7 years\n- Contracts: Duration + 7 years\n\n**Operational Records**:\n- Emails: 3 years\n- Project files: 5 years after completion`,
        required: true,
        order: 2
      },
      {
        id: 'dr-archival',
        type: TemplateSectionType.Procedure,
        title: 'Archival Procedures',
        content: 'Records past their active period but within retention period must be:\n\n1. Moved to secure archive storage\n2. Indexed for retrieval\n3. Access-restricted to authorized personnel\n4. Reviewed annually for destruction eligibility',
        required: true,
        order: 3
      },
      {
        id: 'dr-destruction',
        type: TemplateSectionType.Procedure,
        title: 'Secure Destruction',
        content: 'When retention periods expire:\n\n**Digital Records**: Secure deletion with certified wiping software\n**Paper Records**: Cross-cut shredding or certified destruction\n**Media**: Physical destruction or degaussing\n\nAll destruction must be documented and certified.',
        required: true,
        order: 4
      }
    ]
  },

  // Health & Safety Templates
  {
    title: 'Workplace Health & Safety Policy',
    description: 'Comprehensive health and safety policy covering hazard identification, incident reporting, and safety protocols.',
    category: TemplateCategory.HealthAndSafety,
    industry: TemplateIndustry.General,
    complexity: TemplateComplexity.Standard,
    tags: ['health', 'safety', 'workplace', 'hazards'],
    estimatedTime: '3-4 hours',
    regulatoryFrameworks: ['OSHA'],
    sections: [
      {
        id: 'hs-commitment',
        type: TemplateSectionType.Purpose,
        title: 'Safety Commitment',
        content: '[Company Name] is committed to providing a safe and healthy workplace for all employees, contractors, and visitors. Safety is everyone\'s responsibility.',
        placeholder: '[Company Name]',
        required: true,
        order: 1
      },
      {
        id: 'hs-responsibilities',
        type: TemplateSectionType.Responsibilities,
        title: 'Responsibilities',
        content: `**Management**:\n- Provide safe working conditions\n- Ensure adequate training\n- Investigate incidents promptly\n\n**Supervisors**:\n- Enforce safety rules\n- Conduct safety inspections\n- Address hazards immediately\n\n**Employees**:\n- Follow safety procedures\n- Report hazards and incidents\n- Use required PPE`,
        required: true,
        order: 2
      },
      {
        id: 'hs-hazards',
        type: TemplateSectionType.Policy,
        title: 'Hazard Identification',
        content: 'Regular workplace inspections will identify hazards. Employees should report hazards immediately using the safety reporting system. Common hazards include:\n\n- Slip, trip, and fall hazards\n- Electrical hazards\n- Ergonomic risks\n- Chemical exposure\n- Fire risks',
        required: true,
        order: 3
      },
      {
        id: 'hs-incident',
        type: TemplateSectionType.Procedure,
        title: 'Incident Reporting',
        content: 'All incidents, injuries, and near-misses must be reported within 24 hours using the incident report form. Serious incidents must be reported immediately to [Safety Officer].',
        placeholder: '[Safety Officer]',
        required: true,
        order: 4
      },
      {
        id: 'hs-emergency',
        type: TemplateSectionType.Procedure,
        title: 'Emergency Procedures',
        content: 'Emergency procedures are posted throughout the facility. All employees must:\n\n- Know evacuation routes\n- Attend emergency drills\n- Know location of first aid kits\n- Know how to contact emergency services',
        required: true,
        order: 5
      }
    ]
  },

  // IT Governance Templates
  {
    title: 'Change Management Policy',
    description: 'IT change management policy covering change requests, approvals, and implementation procedures.',
    category: TemplateCategory.ITGovernance,
    industry: TemplateIndustry.Technology,
    complexity: TemplateComplexity.Advanced,
    tags: ['change management', 'IT', 'ITIL', 'governance'],
    estimatedTime: '3-4 hours',
    regulatoryFrameworks: ['ITIL', 'COBIT'],
    sections: [
      {
        id: 'cm-purpose',
        type: TemplateSectionType.Purpose,
        title: 'Purpose',
        content: 'This policy ensures that changes to IT systems are implemented in a controlled manner, minimizing risk and ensuring service continuity.',
        required: true,
        order: 1
      },
      {
        id: 'cm-scope',
        type: TemplateSectionType.Scope,
        title: 'Scope',
        content: 'This policy applies to all changes to:\n\n- Production systems and infrastructure\n- Network configurations\n- Security settings\n- Application deployments\n- Database modifications',
        required: true,
        order: 2
      },
      {
        id: 'cm-classification',
        type: TemplateSectionType.Policy,
        title: 'Change Classification',
        content: `**Standard Changes**: Pre-approved, low-risk changes (e.g., routine patches)\n\n**Normal Changes**: Require CAB review and approval\n\n**Emergency Changes**: Require expedited approval for critical issues\n\nAll changes must be documented in the change management system.`,
        required: true,
        order: 3
      },
      {
        id: 'cm-process',
        type: TemplateSectionType.Procedure,
        title: 'Change Process',
        content: '1. **Request**: Submit change request with details\n2. **Assessment**: Risk and impact assessment\n3. **Approval**: CAB or delegated approval\n4. **Implementation**: Execute during approved window\n5. **Verification**: Test and validate\n6. **Documentation**: Update records and close',
        required: true,
        order: 4
      },
      {
        id: 'cm-rollback',
        type: TemplateSectionType.Procedure,
        title: 'Rollback Procedures',
        content: 'All changes must have a documented rollback plan. If issues occur during implementation, the rollback plan must be executed immediately.',
        required: true,
        order: 5
      }
    ]
  },
  {
    title: 'Business Continuity Policy',
    description: 'Policy for maintaining critical operations during disruptions, including disaster recovery and crisis management.',
    category: TemplateCategory.ITGovernance,
    industry: TemplateIndustry.General,
    complexity: TemplateComplexity.Enterprise,
    tags: ['business continuity', 'disaster recovery', 'BCP', 'crisis'],
    estimatedTime: '6-8 hours',
    regulatoryFrameworks: ['ISO 22301'],
    sections: [
      {
        id: 'bc-purpose',
        type: TemplateSectionType.Purpose,
        title: 'Purpose',
        content: 'This policy ensures [Company Name] can continue critical operations during and after disruptions, protecting employees, customers, and stakeholders.',
        placeholder: '[Company Name]',
        required: true,
        order: 1
      },
      {
        id: 'bc-objectives',
        type: TemplateSectionType.Policy,
        title: 'Recovery Objectives',
        content: `**Recovery Time Objective (RTO)**: Maximum acceptable downtime\n- Critical systems: 4 hours\n- Essential systems: 24 hours\n- Non-essential: 72 hours\n\n**Recovery Point Objective (RPO)**: Maximum acceptable data loss\n- Critical data: 1 hour\n- Business data: 24 hours`,
        required: true,
        order: 2
      },
      {
        id: 'bc-critical',
        type: TemplateSectionType.Policy,
        title: 'Critical Functions',
        content: 'The following functions are critical:\n\n1. [Critical Function 1]\n2. [Critical Function 2]\n3. [Critical Function 3]\n\nDetailed recovery procedures are in the Business Continuity Plan.',
        placeholder: '[Critical Function 1], [Critical Function 2], [Critical Function 3]',
        required: true,
        order: 3
      },
      {
        id: 'bc-testing',
        type: TemplateSectionType.Procedure,
        title: 'Testing & Exercises',
        content: 'BC plans must be tested:\n\n- Tabletop exercises: Quarterly\n- Walkthrough tests: Semi-annually\n- Full simulation: Annually\n\nTest results must be documented and plans updated accordingly.',
        required: true,
        order: 4
      },
      {
        id: 'bc-maintenance',
        type: TemplateSectionType.Procedure,
        title: 'Plan Maintenance',
        content: 'BC plans must be reviewed and updated:\n\n- Annually at minimum\n- After significant changes\n- After incidents or tests\n\nAll updates require senior management approval.',
        required: true,
        order: 5
      }
    ]
  },

  // Finance & Compliance Templates
  {
    title: 'Expense Reimbursement Policy',
    description: 'Policy governing employee expense claims, approval process, and reimbursement procedures.',
    category: TemplateCategory.FinanceCompliance,
    industry: TemplateIndustry.General,
    complexity: TemplateComplexity.Basic,
    tags: ['expenses', 'reimbursement', 'travel', 'finance'],
    estimatedTime: '2 hours',
    sections: [
      {
        id: 'er-purpose',
        type: TemplateSectionType.Purpose,
        title: 'Purpose',
        content: 'This policy provides guidelines for claiming and processing business-related expenses incurred by employees.',
        required: true,
        order: 1
      },
      {
        id: 'er-eligible',
        type: TemplateSectionType.Policy,
        title: 'Eligible Expenses',
        content: `Reimbursable expenses include:\n\n- **Travel**: Flights, trains, taxis (economy class)\n- **Accommodation**: Up to [nightly limit] per night\n- **Meals**: Reasonable costs during business travel\n- **Supplies**: Pre-approved business supplies\n- **Training**: Pre-approved courses and materials`,
        placeholder: '[nightly limit]',
        required: true,
        order: 2
      },
      {
        id: 'er-submission',
        type: TemplateSectionType.Procedure,
        title: 'Submission Requirements',
        content: 'Expense claims must:\n\n1. Be submitted within 30 days of expense\n2. Include itemized receipts\n3. Have manager approval\n4. Use approved expense system\n\nClaims without receipts may be rejected.',
        required: true,
        order: 3
      },
      {
        id: 'er-approval',
        type: TemplateSectionType.Policy,
        title: 'Approval Limits',
        content: '| Amount | Approver |\n|--------|----------|\n| Up to $500 | Line Manager |\n| $500-$2,000 | Department Head |\n| Over $2,000 | Finance Director |',
        required: true,
        order: 4
      }
    ]
  },
  {
    title: 'Anti-Money Laundering Policy',
    description: 'Policy for preventing and detecting money laundering activities in compliance with AML regulations.',
    category: TemplateCategory.FinanceCompliance,
    industry: TemplateIndustry.FinancialServices,
    complexity: TemplateComplexity.Enterprise,
    tags: ['AML', 'compliance', 'financial crime', 'KYC'],
    estimatedTime: '5-6 hours',
    regulatoryFrameworks: ['AML', 'BSA', 'FATF'],
    sections: [
      {
        id: 'aml-purpose',
        type: TemplateSectionType.Purpose,
        title: 'Purpose',
        content: 'This policy establishes controls to prevent [Company Name] from being used for money laundering or terrorist financing activities.',
        placeholder: '[Company Name]',
        required: true,
        order: 1
      },
      {
        id: 'aml-kyc',
        type: TemplateSectionType.Policy,
        title: 'Know Your Customer (KYC)',
        content: `All customers must be verified before establishing a business relationship:\n\n**Individual Customers**:\n- Government-issued ID\n- Proof of address\n- Source of funds (if applicable)\n\n**Business Customers**:\n- Business registration documents\n- Beneficial ownership information\n- Due diligence on directors`,
        required: true,
        order: 2
      },
      {
        id: 'aml-monitoring',
        type: TemplateSectionType.Policy,
        title: 'Transaction Monitoring',
        content: 'Transactions are monitored for suspicious patterns:\n\n- Unusual transaction patterns\n- Transactions inconsistent with customer profile\n- Structuring to avoid reporting thresholds\n- Transactions with high-risk jurisdictions',
        required: true,
        order: 3
      },
      {
        id: 'aml-reporting',
        type: TemplateSectionType.Procedure,
        title: 'Suspicious Activity Reporting',
        content: 'Suspicious activities must be reported:\n\n1. Report to MLRO immediately\n2. MLRO assesses the activity\n3. SAR filed with authorities if warranted\n\nTipping off customers about reports is prohibited.',
        required: true,
        order: 4
      }
    ]
  }
];

/**
 * Policy Template Library Service
 * Manages template browsing, search, customization, and usage tracking
 */
export class PolicyTemplateLibraryService {
  private sp: SPFI;
  private context: WebPartContext;

  private readonly templateListName = TemplateLibraryLists.TEMPLATES;
  private readonly usageListName = TemplateLibraryLists.TEMPLATE_USAGE;
  private readonly preferencesListName = TemplateLibraryLists.USER_PREFERENCES;

  constructor(context: WebPartContext) {
    this.context = context;
    this.sp = spfi().using(SPFx(context));
  }

  /**
   * Get all template categories with counts
   */
  public async getCategories(): Promise<{ category: TemplateCategory; count: number; description: string }[]> {
    try {
      const templates = await this.getAllTemplates();

      const categoryCounts = new Map<TemplateCategory, number>();
      templates.forEach(t => {
        const count = categoryCounts.get(t.category) || 0;
        categoryCounts.set(t.category, count + 1);
      });

      const categoryDescriptions: Record<TemplateCategory, string> = {
        [TemplateCategory.HumanResources]: 'Employee policies, workplace conduct, and HR procedures',
        [TemplateCategory.InformationSecurity]: 'Data protection, access controls, and security protocols',
        [TemplateCategory.DataPrivacy]: 'Privacy policies, data handling, and subject rights',
        [TemplateCategory.HealthAndSafety]: 'Workplace safety, hazard prevention, and incident reporting',
        [TemplateCategory.FinanceCompliance]: 'Financial controls, expense policies, and regulatory compliance',
        [TemplateCategory.OperationalProcedures]: 'Standard operating procedures and process documentation',
        [TemplateCategory.ITGovernance]: 'IT policies, change management, and technology standards',
        [TemplateCategory.LegalCompliance]: 'Legal requirements, contracts, and regulatory obligations',
        [TemplateCategory.CustomerService]: 'Customer interaction standards and service policies',
        [TemplateCategory.EnvironmentalSocial]: 'Environmental, social, and governance (ESG) policies',
        [TemplateCategory.Custom]: 'User-created custom templates'
      };

      return Object.values(TemplateCategory).map(category => ({
        category,
        count: categoryCounts.get(category) || 0,
        description: categoryDescriptions[category]
      }));
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', 'Failed to get categories', error);
      throw error;
    }
  }

  /**
   * Get all templates (built-in + custom from SharePoint)
   */
  public async getAllTemplates(): Promise<IPolicyTemplate[]> {
    try {
      // Get built-in templates
      const builtInTemplates = this.getBuiltInTemplates();

      // Try to get custom templates from SharePoint
      let customTemplates: IPolicyTemplate[] = [];
      try {
        const items = await this.sp.web.lists.getByTitle(this.templateListName)
          .items
          .select('*')
          .top(500)();

        customTemplates = items.map(item => this.mapSharePointItemToTemplate(item));
      } catch {
        // List may not exist yet, continue with built-in only
        logger.info('PolicyTemplateLibraryService', 'Custom templates list not found, using built-in templates only');
      }

      return [...builtInTemplates, ...customTemplates];
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', 'Failed to get all templates', error);
      throw error;
    }
  }

  /**
   * Get built-in templates
   */
  public getBuiltInTemplates(): IPolicyTemplate[] {
    return BUILT_IN_TEMPLATES.map((template, index) => ({
      id: -(index + 1), // Negative IDs for built-in templates
      title: template.title || '',
      description: template.description || '',
      category: template.category || TemplateCategory.Custom,
      industry: template.industry || TemplateIndustry.General,
      complexity: template.complexity || TemplateComplexity.Standard,
      sections: template.sections as ITemplateSection[] || [],
      tags: template.tags || [],
      version: '1.0',
      author: 'JML System',
      createdDate: new Date('2024-01-01'),
      modifiedDate: new Date('2024-01-01'),
      usageCount: 0,
      rating: 4.5,
      ratingCount: 100,
      isPublic: true,
      isFeatured: true,
      estimatedTime: template.estimatedTime,
      regulatoryFrameworks: template.regulatoryFrameworks
    }));
  }

  /**
   * Search templates with filters
   */
  public async searchTemplates(criteria: ITemplateSearchCriteria): Promise<ITemplateSearchResult> {
    try {
      let templates = await this.getAllTemplates();

      // Apply filters
      if (criteria.searchText) {
        const searchLower = criteria.searchText.toLowerCase();
        templates = templates.filter(t =>
          t.title.toLowerCase().includes(searchLower) ||
          t.description.toLowerCase().includes(searchLower) ||
          t.tags.some(tag => tag.toLowerCase().includes(searchLower))
        );
      }

      if (criteria.categories && criteria.categories.length > 0) {
        templates = templates.filter(t => criteria.categories!.includes(t.category));
      }

      if (criteria.industries && criteria.industries.length > 0) {
        templates = templates.filter(t => criteria.industries!.includes(t.industry));
      }

      if (criteria.complexity && criteria.complexity.length > 0) {
        templates = templates.filter(t => criteria.complexity!.includes(t.complexity));
      }

      if (criteria.tags && criteria.tags.length > 0) {
        templates = templates.filter(t =>
          criteria.tags!.some(tag => t.tags.includes(tag))
        );
      }

      if (criteria.isFeatured !== undefined) {
        templates = templates.filter(t => t.isFeatured === criteria.isFeatured);
      }

      if (criteria.minRating !== undefined) {
        templates = templates.filter(t => t.rating >= criteria.minRating!);
      }

      if (criteria.regulatoryFramework) {
        templates = templates.filter(t =>
          t.regulatoryFrameworks?.includes(criteria.regulatoryFramework!)
        );
      }

      // Calculate facets before pagination
      const facets = this.calculateFacets(templates);

      // Apply sorting
      const sortBy = criteria.sortBy || 'title';
      const sortDir = criteria.sortDirection || 'asc';

      templates.sort((a, b) => {
        let comparison = 0;
        switch (sortBy) {
          case 'usageCount':
            comparison = a.usageCount - b.usageCount;
            break;
          case 'rating':
            comparison = a.rating - b.rating;
            break;
          case 'modifiedDate':
            comparison = a.modifiedDate.getTime() - b.modifiedDate.getTime();
            break;
          default:
            comparison = a.title.localeCompare(b.title);
        }
        return sortDir === 'desc' ? -comparison : comparison;
      });

      // Apply pagination
      const pageSize = criteria.pageSize || 20;
      const pageIndex = criteria.pageIndex || 0;
      const startIndex = pageIndex * pageSize;
      const paginatedTemplates = templates.slice(startIndex, startIndex + pageSize);

      return {
        templates: paginatedTemplates,
        totalCount: templates.length,
        pageIndex,
        pageSize,
        hasMore: startIndex + pageSize < templates.length,
        facets
      };
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', 'Failed to search templates', error);
      throw error;
    }
  }

  /**
   * Get a specific template by ID
   */
  public async getTemplate(templateId: number): Promise<IPolicyTemplate | null> {
    try {
      // Check built-in templates first (negative IDs)
      if (templateId < 0) {
        const builtInTemplates = this.getBuiltInTemplates();
        return builtInTemplates.find(t => t.id === templateId) || null;
      }

      // Get from SharePoint
      const item = await this.sp.web.lists.getByTitle(this.templateListName)
        .items.getById(templateId)
        .select('*')();

      return this.mapSharePointItemToTemplate(item);
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', `Failed to get template ${templateId}`, error);
      return null;
    }
  }

  /**
   * Get featured templates
   */
  public async getFeaturedTemplates(limit: number = 6): Promise<IPolicyTemplate[]> {
    const result = await this.searchTemplates({
      isFeatured: true,
      sortBy: 'rating',
      sortDirection: 'desc',
      pageSize: limit
    });
    return result.templates;
  }

  /**
   * Get templates by category
   */
  public async getTemplatesByCategory(category: TemplateCategory): Promise<IPolicyTemplate[]> {
    const result = await this.searchTemplates({
      categories: [category],
      sortBy: 'usageCount',
      sortDirection: 'desc'
    });
    return result.templates;
  }

  /**
   * Get templates for a regulatory framework
   */
  public async getTemplatesForFramework(framework: string): Promise<IPolicyTemplate[]> {
    const result = await this.searchTemplates({
      regulatoryFramework: framework,
      sortBy: 'rating',
      sortDirection: 'desc'
    });
    return result.templates;
  }

  /**
   * Clone a template with customizations
   */
  public async cloneTemplate(customization: ITemplateCustomization): Promise<IPolicyTemplate> {
    try {
      const sourceTemplate = await this.getTemplate(customization.templateId);
      if (!sourceTemplate) {
        throw new Error(`Template ${customization.templateId} not found`);
      }

      // Build customized sections
      let sections = sourceTemplate.sections.filter(s =>
        customization.selectedSections.includes(s.id)
      );

      // Apply section overrides
      if (customization.sectionOverrides) {
        sections = sections.map(section => ({
          ...section,
          ...customization.sectionOverrides![section.id]
        }));
      }

      // Add custom sections
      if (customization.customSections) {
        sections = [...sections, ...customization.customSections];
      }

      // Re-order sections
      sections.sort((a, b) => a.order - b.order);

      // Create cloned template
      const clonedTemplate: Partial<IPolicyTemplate> = {
        title: customization.newTitle,
        description: customization.newDescription || sourceTemplate.description,
        category: sourceTemplate.category,
        industry: sourceTemplate.industry,
        complexity: sourceTemplate.complexity,
        sections,
        tags: sourceTemplate.tags,
        version: '1.0',
        author: this.context.pageContext.user.displayName,
        createdDate: new Date(),
        modifiedDate: new Date(),
        usageCount: 0,
        rating: 0,
        ratingCount: 0,
        isPublic: false,
        isFeatured: false,
        estimatedTime: sourceTemplate.estimatedTime,
        regulatoryFrameworks: sourceTemplate.regulatoryFrameworks
      };

      // Save to SharePoint
      const savedItem = await this.sp.web.lists.getByTitle(this.templateListName)
        .items.add({
          Title: clonedTemplate.title,
          Description: clonedTemplate.description,
          Category: clonedTemplate.category,
          Industry: clonedTemplate.industry,
          Complexity: clonedTemplate.complexity,
          Sections: JSON.stringify(clonedTemplate.sections),
          Tags: clonedTemplate.tags?.join(';'),
          Version: clonedTemplate.version,
          EstimatedTime: clonedTemplate.estimatedTime,
          RegulatoryFrameworks: clonedTemplate.regulatoryFrameworks?.join(';'),
          IsPublic: clonedTemplate.isPublic,
          IsFeatured: clonedTemplate.isFeatured
        });

      // Record usage
      await this.recordTemplateUsage(customization.templateId, savedItem.data.Id);

      logger.info('PolicyTemplateLibraryService', `Template cloned: ${clonedTemplate.title}`);

      return {
        ...clonedTemplate,
        id: savedItem.data.Id
      } as IPolicyTemplate;
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', 'Failed to clone template', error);
      throw error;
    }
  }

  /**
   * Save a new custom template
   */
  public async saveTemplate(template: Partial<IPolicyTemplate>): Promise<number> {
    try {
      const result = await this.sp.web.lists.getByTitle(this.templateListName)
        .items.add({
          Title: template.title,
          Description: template.description,
          Category: template.category,
          Industry: template.industry || TemplateIndustry.General,
          Complexity: template.complexity || TemplateComplexity.Standard,
          Sections: JSON.stringify(template.sections),
          Tags: template.tags?.join(';'),
          Version: template.version || '1.0',
          EstimatedTime: template.estimatedTime,
          RegulatoryFrameworks: template.regulatoryFrameworks?.join(';'),
          IsPublic: template.isPublic || false,
          IsFeatured: false
        });

      logger.info('PolicyTemplateLibraryService', `Template saved: ${template.title}`);
      return result.data.Id;
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', 'Failed to save template', error);
      throw error;
    }
  }

  /**
   * Update an existing custom template
   */
  public async updateTemplate(templateId: number, updates: Partial<IPolicyTemplate>): Promise<void> {
    try {
      if (templateId < 0) {
        throw new Error('Cannot update built-in templates');
      }

      const updateData: Record<string, unknown> = {};

      if (updates.title !== undefined) updateData.Title = updates.title;
      if (updates.description !== undefined) updateData.Description = updates.description;
      if (updates.category !== undefined) updateData.Category = updates.category;
      if (updates.industry !== undefined) updateData.Industry = updates.industry;
      if (updates.complexity !== undefined) updateData.Complexity = updates.complexity;
      if (updates.sections !== undefined) updateData.Sections = JSON.stringify(updates.sections);
      if (updates.tags !== undefined) updateData.Tags = updates.tags.join(';');
      if (updates.version !== undefined) updateData.Version = updates.version;
      if (updates.estimatedTime !== undefined) updateData.EstimatedTime = updates.estimatedTime;
      if (updates.regulatoryFrameworks !== undefined) {
        updateData.RegulatoryFrameworks = updates.regulatoryFrameworks.join(';');
      }
      if (updates.isPublic !== undefined) updateData.IsPublic = updates.isPublic;

      await this.sp.web.lists.getByTitle(this.templateListName)
        .items.getById(templateId)
        .update(updateData);

      logger.info('PolicyTemplateLibraryService', `Template updated: ${templateId}`);
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', `Failed to update template ${templateId}`, error);
      throw error;
    }
  }

  /**
   * Delete a custom template
   */
  public async deleteTemplate(templateId: number): Promise<void> {
    try {
      if (templateId < 0) {
        throw new Error('Cannot delete built-in templates');
      }

      await this.sp.web.lists.getByTitle(this.templateListName)
        .items.getById(templateId)
        .delete();

      logger.info('PolicyTemplateLibraryService', `Template deleted: ${templateId}`);
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', `Failed to delete template ${templateId}`, error);
      throw error;
    }
  }

  /**
   * Rate a template
   */
  public async rateTemplate(templateId: number, rating: number): Promise<void> {
    try {
      if (rating < 1 || rating > 5) {
        throw new Error('Rating must be between 1 and 5');
      }

      const template = await this.getTemplate(templateId);
      if (!template) {
        throw new Error(`Template ${templateId} not found`);
      }

      // Calculate new average rating
      const newRatingCount = template.ratingCount + 1;
      const newRating = ((template.rating * template.ratingCount) + rating) / newRatingCount;

      if (templateId > 0) {
        await this.sp.web.lists.getByTitle(this.templateListName)
          .items.getById(templateId)
          .update({
            Rating: newRating,
            RatingCount: newRatingCount
          });
      }

      logger.info('PolicyTemplateLibraryService', `Template ${templateId} rated: ${rating}`);
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', `Failed to rate template ${templateId}`, error);
      throw error;
    }
  }

  /**
   * Record template usage
   */
  public async recordTemplateUsage(templateId: number, policyId?: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.usageListName)
        .items.add({
          Title: `Usage: Template ${templateId}`,
          TemplateId: templateId,
          PolicyId: policyId,
          UsedDate: new Date().toISOString()
        });

      // Increment usage count if custom template
      if (templateId > 0) {
        const template = await this.getTemplate(templateId);
        if (template) {
          await this.sp.web.lists.getByTitle(this.templateListName)
            .items.getById(templateId)
            .update({
              UsageCount: template.usageCount + 1
            });
        }
      }
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', `Failed to record template usage: ${templateId}`, error);
      // Don't throw - usage tracking is non-critical
    }
  }

  /**
   * Get user's favorite templates
   */
  public async getUserFavorites(): Promise<number[]> {
    try {
      const userId = this.context.pageContext.user.loginName;
      const items = await this.sp.web.lists.getByTitle(this.preferencesListName)
        .items
        .filter(`UserId eq '${userId}'`)
        .select('FavoriteTemplates')
        .top(1)();

      if (items.length > 0 && items[0].FavoriteTemplates) {
        return items[0].FavoriteTemplates.split(',').map((id: string) => parseInt(id, 10));
      }
      return [];
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', 'Failed to get user favorites', error);
      return [];
    }
  }

  /**
   * Add template to favorites
   */
  public async addToFavorites(templateId: number): Promise<void> {
    try {
      const favorites = await this.getUserFavorites();
      if (!favorites.includes(templateId)) {
        favorites.push(templateId);
        await this.saveUserFavorites(favorites);
      }
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', `Failed to add favorite: ${templateId}`, error);
      throw error;
    }
  }

  /**
   * Remove template from favorites
   */
  public async removeFromFavorites(templateId: number): Promise<void> {
    try {
      const favorites = await this.getUserFavorites();
      const index = favorites.indexOf(templateId);
      if (index > -1) {
        favorites.splice(index, 1);
        await this.saveUserFavorites(favorites);
      }
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', `Failed to remove favorite: ${templateId}`, error);
      throw error;
    }
  }

  /**
   * Get user's recent templates
   */
  public async getRecentTemplates(limit: number = 5): Promise<IPolicyTemplate[]> {
    try {
      const userId = this.context.pageContext.user.loginName;
      const usageItems = await this.sp.web.lists.getByTitle(this.usageListName)
        .items
        .filter(`Author/EMail eq '${userId}'`)
        .orderBy('UsedDate', false)
        .select('TemplateId', 'UsedDate')
        .top(limit * 2)(); // Get extra to handle duplicates

      // Get unique template IDs
      const seenIds = new Set<number>();
      const uniqueTemplateIds: number[] = [];
      for (const item of usageItems) {
        if (!seenIds.has(item.TemplateId) && uniqueTemplateIds.length < limit) {
          seenIds.add(item.TemplateId);
          uniqueTemplateIds.push(item.TemplateId);
        }
      }

      // Fetch templates
      const templates: IPolicyTemplate[] = [];
      for (const id of uniqueTemplateIds) {
        const template = await this.getTemplate(id);
        if (template) {
          templates.push(template);
        }
      }

      return templates;
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', 'Failed to get recent templates', error);
      return [];
    }
  }

  /**
   * Get all available tags
   */
  public async getAllTags(): Promise<{ tag: string; count: number }[]> {
    try {
      const templates = await this.getAllTemplates();
      const tagCounts = new Map<string, number>();

      templates.forEach(t => {
        t.tags.forEach(tag => {
          const count = tagCounts.get(tag) || 0;
          tagCounts.set(tag, count + 1);
        });
      });

      return Array.from(tagCounts.entries())
        .map(([tag, count]) => ({ tag, count }))
        .sort((a, b) => b.count - a.count);
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', 'Failed to get tags', error);
      return [];
    }
  }

  /**
   * Get all regulatory frameworks
   */
  public async getAllFrameworks(): Promise<{ framework: string; count: number }[]> {
    try {
      const templates = await this.getAllTemplates();
      const frameworkCounts = new Map<string, number>();

      templates.forEach(t => {
        t.regulatoryFrameworks?.forEach(framework => {
          const count = frameworkCounts.get(framework) || 0;
          frameworkCounts.set(framework, count + 1);
        });
      });

      return Array.from(frameworkCounts.entries())
        .map(([framework, count]) => ({ framework, count }))
        .sort((a, b) => b.count - a.count);
    } catch (error) {
      logger.error('PolicyTemplateLibraryService', 'Failed to get frameworks', error);
      return [];
    }
  }

  /**
   * Generate policy content from template
   */
  public generatePolicyContent(template: IPolicyTemplate, replacements?: Record<string, string>): string {
    let content = '';

    // Sort sections by order
    const sortedSections = [...template.sections].sort((a, b) => a.order - b.order);

    for (const section of sortedSections) {
      content += `## ${section.title}\n\n`;

      let sectionContent = section.content;

      // Apply placeholder replacements
      if (replacements) {
        for (const [placeholder, value] of Object.entries(replacements)) {
          sectionContent = sectionContent.replace(new RegExp(placeholder.replace(/[[\]]/g, '\\$&'), 'g'), value);
        }
      }

      content += `${sectionContent}\n\n`;
    }

    return content;
  }

  /**
   * Get template placeholders
   */
  public getTemplatePlaceholders(template: IPolicyTemplate): string[] {
    const placeholders = new Set<string>();

    for (const section of template.sections) {
      if (section.placeholder) {
        // Split multiple placeholders
        const matches = section.placeholder.match(/\[[^\]]+\]/g);
        if (matches) {
          matches.forEach(p => placeholders.add(p));
        }
      }

      // Also search content for placeholders
      const contentMatches = section.content.match(/\[[^\]]+\]/g);
      if (contentMatches) {
        contentMatches.forEach(p => placeholders.add(p));
      }
    }

    return Array.from(placeholders);
  }

  /**
   * Helper: Calculate facets for search results
   */
  private calculateFacets(templates: IPolicyTemplate[]): ITemplateFacets {
    const categoryMap = new Map<TemplateCategory, number>();
    const industryMap = new Map<TemplateIndustry, number>();
    const complexityMap = new Map<TemplateComplexity, number>();
    const tagMap = new Map<string, number>();

    templates.forEach(t => {
      categoryMap.set(t.category, (categoryMap.get(t.category) || 0) + 1);
      industryMap.set(t.industry, (industryMap.get(t.industry) || 0) + 1);
      complexityMap.set(t.complexity, (complexityMap.get(t.complexity) || 0) + 1);
      t.tags.forEach(tag => {
        tagMap.set(tag, (tagMap.get(tag) || 0) + 1);
      });
    });

    return {
      categories: Array.from(categoryMap.entries())
        .map(([category, count]) => ({ category, count }))
        .sort((a, b) => b.count - a.count),
      industries: Array.from(industryMap.entries())
        .map(([industry, count]) => ({ industry, count }))
        .sort((a, b) => b.count - a.count),
      complexity: Array.from(complexityMap.entries())
        .map(([complexity, count]) => ({ complexity, count }))
        .sort((a, b) => b.count - a.count),
      tags: Array.from(tagMap.entries())
        .map(([tag, count]) => ({ tag, count }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 20) // Top 20 tags
    };
  }

  /**
   * Helper: Save user favorites
   */
  private async saveUserFavorites(favorites: number[]): Promise<void> {
    const userId = this.context.pageContext.user.loginName;

    const items = await this.sp.web.lists.getByTitle(this.preferencesListName)
      .items
      .filter(`UserId eq '${userId}'`)
      .select('Id')
      .top(1)();

    const favoritesString = favorites.join(',');

    if (items.length > 0) {
      await this.sp.web.lists.getByTitle(this.preferencesListName)
        .items.getById(items[0].Id)
        .update({ FavoriteTemplates: favoritesString });
    } else {
      await this.sp.web.lists.getByTitle(this.preferencesListName)
        .items.add({
          Title: userId,
          UserId: userId,
          FavoriteTemplates: favoritesString
        });
    }
  }

  /**
   * Helper: Map SharePoint item to template
   */
  private mapSharePointItemToTemplate(item: Record<string, unknown>): IPolicyTemplate {
    return {
      id: item.Id as number,
      title: item.Title as string,
      description: item.Description as string || '',
      category: item.Category as TemplateCategory || TemplateCategory.Custom,
      industry: item.Industry as TemplateIndustry || TemplateIndustry.General,
      complexity: item.Complexity as TemplateComplexity || TemplateComplexity.Standard,
      sections: item.Sections ? JSON.parse(item.Sections as string) : [],
      tags: item.Tags ? (item.Tags as string).split(';') : [],
      version: item.Version as string || '1.0',
      author: (item.Author as { Title?: string })?.Title || '',
      createdDate: new Date(item.Created as string),
      modifiedDate: new Date(item.Modified as string),
      usageCount: item.UsageCount as number || 0,
      rating: item.Rating as number || 0,
      ratingCount: item.RatingCount as number || 0,
      isPublic: item.IsPublic as boolean || false,
      isFeatured: item.IsFeatured as boolean || false,
      previewImage: item.PreviewImage as string,
      estimatedTime: item.EstimatedTime as string,
      regulatoryFrameworks: item.RegulatoryFrameworks
        ? (item.RegulatoryFrameworks as string).split(';')
        : []
    };
  }
}

// Export singleton factory
let serviceInstance: PolicyTemplateLibraryService | null = null;

export const getPolicyTemplateLibraryService = (context: WebPartContext): PolicyTemplateLibraryService => {
  if (!serviceInstance) {
    serviceInstance = new PolicyTemplateLibraryService(context);
  }
  return serviceInstance;
};
