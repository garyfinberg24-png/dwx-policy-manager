// Integration Models
// Interfaces for external system integrations

import { IBaseListItem, IUser } from './ICommon';

/**
 * Integration Types
 */
export enum IntegrationType {
  EntraID = 'Entra ID',
  MicrosoftTeams = 'Microsoft Teams',
  Planner = 'Planner',
  Exchange = 'Exchange',
  PowerAutomate = 'Power Automate',
  SAP = 'SAP',
  Workday = 'Workday',
  ServiceNow = 'ServiceNow',
  Custom = 'Custom'
}

/**
 * Integration Status
 */
export enum IntegrationStatus {
  NotConfigured = 'Not Configured',
  Active = 'Active',
  Inactive = 'Inactive',
  Error = 'Error',
  Syncing = 'Syncing'
}

/**
 * Sync Direction
 */
export enum SyncDirection {
  Import = 'Import',
  Export = 'Export',
  Bidirectional = 'Bidirectional'
}

/**
 * Integration Configuration
 */
export interface IIntegrationConfig extends IBaseListItem {
  IntegrationType: IntegrationType;
  Status: IntegrationStatus;
  EndpointUrl?: string;
  ApiKey?: string;
  ClientId?: string;
  TenantId?: string;
  IsEnabled: boolean;
  SyncDirection: SyncDirection;
  SyncFrequency?: number; // in minutes
  LastSyncDate?: Date;
  LastSyncStatus?: string;
  ErrorMessage?: string;
  Configuration?: string; // JSON configuration
  ProcessTypes?: string[]; // Which process types use this integration
}

/**
 * Integration Log
 */
export interface IIntegrationLog extends IBaseListItem {
  IntegrationConfigId: number;
  IntegrationType: IntegrationType;
  ProcessID?: number;
  Action: string;
  Status: 'Success' | 'Failed' | 'Warning';
  RequestData?: string; // JSON
  ResponseData?: string; // JSON
  ErrorMessage?: string;
  ExecutionTime?: number; // in milliseconds
  CreatedBy?: IUser;
}

/**
 * Employee from Entra ID
 */
export interface IEntraIDEmployee {
  id: string;
  userPrincipalName: string;
  displayName: string;
  givenName: string;
  surname: string;
  mail: string;
  mobilePhone?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  manager?: {
    id: string;
    displayName: string;
    mail: string;
  };
  employeeId?: string;
  companyName?: string;
  employeeHireDate?: Date;
}

/**
 * Teams Channel
 */
export interface ITeamsChannel {
  id: string;
  displayName: string;
  description?: string;
  webUrl?: string;
  membershipType?: 'standard' | 'private' | 'shared';
}

/**
 * Teams Channel Creation Request
 */
export interface ITeamsChannelRequest {
  teamId: string;
  displayName: string;
  description?: string;
  membershipType?: 'standard' | 'private';
  members?: string[]; // User IDs
  owners?: string[]; // User IDs
}

/**
 * Planner Plan
 */
export interface IPlannerPlan {
  id: string;
  title: string;
  owner?: string; // Group ID
  createdDateTime?: Date;
}

/**
 * Planner Task
 */
export interface IPlannerTask {
  id?: string;
  planId: string;
  bucketId?: string;
  title: string;
  percentComplete: number;
  startDateTime?: Date;
  dueDateTime?: Date;
  assignments?: { [userId: string]: { orderHint: string } };
  priority?: number; // 0-10
  description?: string;
  checklist?: { [key: string]: { title: string; isChecked: boolean } };
  references?: { [url: string]: { alias: string; type: string } };
}

/**
 * Planner Task Sync Request
 */
export interface IPlannerSyncRequest {
  processId: number;
  planId: string;
  bucketId?: string;
  tasks: {
    taskId: number;
    title: string;
    assignedTo: string;
    dueDate: Date;
    description?: string;
  }[];
}

/**
 * Exchange Mailbox
 */
export interface IExchangeMailbox {
  id: string;
  displayName: string;
  emailAddress: string;
  mailboxType: 'User' | 'Shared' | 'Room' | 'Equipment';
  isEnabled: boolean;
}

/**
 * Exchange Mailbox Creation Request
 */
export interface IExchangeMailboxRequest {
  displayName: string;
  alias: string;
  members?: string[]; // User IDs
  owners?: string[]; // User IDs
  emailAddresses?: string[];
}

/**
 * Power Automate Flow
 */
export interface IPowerAutomateFlow {
  id: string;
  name: string;
  displayName: string;
  state: 'Started' | 'Stopped' | 'Suspended';
  createdTime: Date;
  lastModifiedTime: Date;
}

/**
 * Power Automate Trigger Request
 */
export interface IPowerAutomateTrigger {
  flowId: string;
  triggerName?: string;
  inputs: {
    [key: string]: any;
  };
}

/**
 * SAP/Workday Employee Data
 */
export interface IHRSystemEmployee {
  employeeId: string;
  firstName: string;
  lastName: string;
  email: string;
  jobTitle: string;
  department: string;
  managerId?: string;
  managerName?: string;
  startDate: Date;
  location?: string;
  costCenter?: string;
  division?: string;
  employeeType?: 'FTE' | 'Contractor' | 'Intern';
  workPhone?: string;
  mobilePhone?: string;
  customFields?: { [key: string]: any };
}

/**
 * Integration Mapping
 */
export interface IIntegrationMapping extends IBaseListItem {
  IntegrationType: IntegrationType;
  SourceField: string;
  TargetField: string;
  TransformationRule?: string; // JavaScript expression
  IsRequired: boolean;
  DefaultValue?: string;
}

/**
 * Integration Event
 */
export interface IIntegrationEvent {
  eventType: 'sync' | 'create' | 'update' | 'delete';
  integrationType: IntegrationType;
  timestamp: Date;
  processId?: number;
  entityId?: string;
  entityType?: string;
  success: boolean;
  message?: string;
  data?: any;
}

/**
 * Integration Service Response
 */
export interface IIntegrationResponse<T = any> {
  success: boolean;
  data?: T;
  error?: string;
  statusCode?: number;
  timestamp: Date;
}

/**
 * Bulk Integration Request
 */
export interface IBulkIntegrationRequest {
  integrationType: IntegrationType;
  action: 'create' | 'update' | 'sync';
  items: any[];
  processId?: number;
}

/**
 * Integration Health Status
 */
export interface IIntegrationHealth {
  integrationType: IntegrationType;
  isHealthy: boolean;
  lastCheckTime: Date;
  responseTime?: number; // in ms
  errorCount: number;
  successRate: number; // percentage
  message?: string;
}

/**
 * Webhook Configuration
 */
export interface IWebhookConfig extends IBaseListItem {
  IntegrationType: IntegrationType;
  WebhookUrl: string;
  Secret?: string;
  Events: string[]; // Array of event types to listen for
  IsActive: boolean;
  Headers?: string; // JSON object of custom headers
  RetryPolicy?: {
    maxRetries: number;
    retryDelay: number; // in seconds
  };
}
