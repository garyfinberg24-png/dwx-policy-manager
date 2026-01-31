// PM_Configuration List Model

import { IBaseListItem } from './ICommon';

export interface IJmlConfiguration extends IBaseListItem {
  // Configuration Key-Value
  ConfigKey: string; // Unique key (e.g., "EmailNotifications.Enabled")
  ConfigValue: string;
  Description?: string;

  // Metadata
  Category?: string; // e.g., "Email", "SLA", "Integration"
  DataType?: string; // e.g., "string", "number", "boolean", "json"
  IsActive: boolean;
  IsSystemConfig?: boolean; // If true, only admins can modify

  // Validation
  AllowedValues?: string; // Comma-separated for dropdown configs
  ValidationRegex?: string;
  MinValue?: number;
  MaxValue?: number;

  // Audit
  LastModifiedBy?: string;
  LastModifiedDate?: Date;
}

// Common configuration keys (constants)
export class ConfigKeys {
  // Email Settings
  static readonly EMAIL_ENABLED = 'EmailNotifications.Enabled';
  static readonly EMAIL_FROM = 'EmailNotifications.FromAddress';
  static readonly EMAIL_TEMPLATE = 'EmailNotifications.DefaultTemplate';

  // SLA Settings
  static readonly SLA_DEFAULT_HOURS = 'SLA.DefaultHours';
  static readonly SLA_REMINDER_DAYS = 'SLA.ReminderBeforeDueDays';
  static readonly SLA_ESCALATION_DAYS = 'SLA.EscalationAfterDueDays';

  // Integration Settings
  static readonly GRAPH_API_ENABLED = 'Integration.GraphAPI.Enabled';
  static readonly TEAMS_NOTIFICATIONS = 'Integration.Teams.NotificationsEnabled';

  // AI / Integration Settings
  static readonly AI_FUNCTION_URL = 'Integration.AI.FunctionUrl';

  // Business Rules
  static readonly APPROVAL_REQUIRED = 'BusinessRules.ApprovalRequired';
  static readonly AUTO_ASSIGN_TASKS = 'BusinessRules.AutoAssignTasks';

  // Policy Authoring Settings
  static readonly POLICY_TEMPLATES_ONLY = 'Policy.TemplatesOnly';
  static readonly POLICY_REQUIRE_APPROVAL = 'Policy.RequireApproval';
  static readonly POLICY_DEFAULT_CATEGORY = 'Policy.DefaultCategory';
  static readonly POLICY_ALLOWED_DOCUMENT_TYPES = 'Policy.AllowedDocumentTypes';

  // Onboarding Experience Settings
  static readonly ONBOARDING_ENABLED = 'Onboarding.Enabled';
  static readonly ONBOARDING_DEFAULT_THEME = 'Onboarding.DefaultTheme';
  static readonly ONBOARDING_GAMIFICATION_ENABLED = 'Onboarding.GamificationEnabled';

  // Currency Settings
  static readonly CURRENCY_DEFAULT = 'Currency.Default';
  static readonly CURRENCY_DECIMAL_SEPARATOR = 'Currency.DecimalSeparator';
  static readonly CURRENCY_THOUSANDS_SEPARATOR = 'Currency.ThousandsSeparator';
  static readonly CURRENCY_DECIMAL_PLACES = 'Currency.DecimalPlaces';
  static readonly CURRENCY_SYMBOL_POSITION = 'Currency.SymbolPosition';
  static readonly CURRENCY_SHOW_CODE = 'Currency.ShowCode';
  static readonly CURRENCY_NEGATIVE_FORMAT = 'Currency.NegativeFormat';
  static readonly CURRENCY_ENABLED_LIST = 'Currency.EnabledList';
}

// Typed configuration helper
export interface ITypedConfig {
  emailEnabled: boolean;
  emailFrom: string;
  slaDefaultHours: number;
  slaReminderDays: number;
  slaEscalationDays: number;
  graphApiEnabled: boolean;
  teamsNotifications: boolean;
  approvalRequired: boolean;
  autoAssignTasks: boolean;
  // Currency settings
  currencyDefault: string;
  currencyDecimalSeparator: '.' | ',';
  currencyThousandsSeparator: ',' | '.' | ' ' | '';
  currencyDecimalPlaces: number | null;
  currencySymbolPosition: 'before' | 'after' | null;
  currencyShowCode: boolean;
  currencyNegativeFormat: 'minus' | 'parentheses' | 'minusAfter';
  currencyEnabledList: string[];
}
