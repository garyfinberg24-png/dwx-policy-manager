// @ts-nocheck
/**
 * Policy Audit Service
 * Comprehensive audit logging for all policy operations
 * Tracks approvals, rejections, acknowledgements, and all changes
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import { IPolicy, IPolicyAcknowledgement, PolicyStatus, AcknowledgementStatus } from '../models/IPolicy';
import { logger } from './LoggingService';
import { PolicyLists, SystemLists } from '../constants/SharePointListNames';

// ============================================================================
// ENUMS & INTERFACES
// ============================================================================

/**
 * Audit event types
 */
export enum AuditEventType {
  // Policy Lifecycle
  PolicyCreated = 'Policy.Created',
  PolicyUpdated = 'Policy.Updated',
  PolicyDeleted = 'Policy.Deleted',
  PolicyViewed = 'Policy.Viewed',
  PolicyDownloaded = 'Policy.Downloaded',
  PolicyPrinted = 'Policy.Printed',

  // Workflow Events
  PolicySubmitted = 'Policy.Submitted',
  PolicyApproved = 'Policy.Approved',
  PolicyRejected = 'Policy.Rejected',
  PolicyPublished = 'Policy.Published',
  PolicyRepublished = 'Policy.Republished',
  PolicyArchived = 'Policy.Archived',
  PolicyRetired = 'Policy.Retired',
  PolicyRestored = 'Policy.Restored',

  // Version Events
  VersionCreated = 'Version.Created',
  VersionActivated = 'Version.Activated',
  VersionRolledBack = 'Version.RolledBack',

  // Distribution Events
  PolicyDistributed = 'Distribution.Created',
  DistributionCancelled = 'Distribution.Cancelled',
  RecipientAdded = 'Distribution.RecipientAdded',
  RecipientRemoved = 'Distribution.RecipientRemoved',

  // Acknowledgement Events
  AcknowledgementAssigned = 'Acknowledgement.Assigned',
  AcknowledgementOpened = 'Acknowledgement.Opened',
  AcknowledgementCompleted = 'Acknowledgement.Completed',
  AcknowledgementFailed = 'Acknowledgement.Failed',
  AcknowledgementOverdue = 'Acknowledgement.Overdue',
  AcknowledgementReset = 'Acknowledgement.Reset',

  // Delegation Events
  DelegationRequested = 'Delegation.Requested',
  DelegationApproved = 'Delegation.Approved',
  DelegationRejected = 'Delegation.Rejected',
  DelegationRevoked = 'Delegation.Revoked',

  // Exemption Events
  ExemptionRequested = 'Exemption.Requested',
  ExemptionApproved = 'Exemption.Approved',
  ExemptionRejected = 'Exemption.Rejected',
  ExemptionExpired = 'Exemption.Expired',
  ExemptionRevoked = 'Exemption.Revoked',

  // Quiz Events
  QuizStarted = 'Quiz.Started',
  QuizCompleted = 'Quiz.Completed',
  QuizPassed = 'Quiz.Passed',
  QuizFailed = 'Quiz.Failed',

  // Administrative Events
  BulkOperation = 'Admin.BulkOperation',
  SettingsChanged = 'Admin.SettingsChanged',
  PermissionsChanged = 'Admin.PermissionsChanged',
  RetentionApplied = 'Admin.RetentionApplied',
  DataExported = 'Admin.DataExported',
  DataPurged = 'Admin.DataPurged',

  // Security Events
  UnauthorizedAccess = 'Security.UnauthorizedAccess',
  ValidationFailed = 'Security.ValidationFailed',
  SuspiciousActivity = 'Security.SuspiciousActivity'
}

/**
 * Audit severity levels
 */
export enum AuditSeverity {
  Info = 'Info',
  Warning = 'Warning',
  Error = 'Error',
  Critical = 'Critical',
  Security = 'Security'
}

/**
 * Audit log entry
 */
export interface IAuditEntry {
  Id?: number;

  // Event identification
  EventId: string; // GUID for unique identification
  EventType: AuditEventType;
  EventCategory: string; // Policy, Acknowledgement, Admin, Security
  Severity: AuditSeverity;

  // Entity information
  EntityType: 'Policy' | 'Acknowledgement' | 'Exemption' | 'Distribution' | 'Quiz' | 'System' | 'AuditLog';
  EntityId: number;
  EntityName?: string;
  PolicyId?: number;
  PolicyNumber?: string;
  PolicyName?: string;

  // Action details
  Action: string;
  ActionDescription: string;
  ActionResult: 'Success' | 'Failure' | 'Partial';

  // User information
  PerformedById: number;
  PerformedByEmail: string;
  PerformedByName?: string;
  PerformedByRoles?: string; // JSON array of roles

  // Target user (if different from performer)
  TargetUserId?: number;
  TargetUserEmail?: string;
  TargetUserName?: string;

  // Technical details
  IPAddress?: string;
  UserAgent?: string;
  DeviceType?: string;
  SessionId?: string;
  RequestId?: string;

  // Change tracking
  PreviousValue?: string; // JSON of previous state
  NewValue?: string; // JSON of new state
  ChangedFields?: string; // JSON array of field names
  ChangeDetails?: string; // Human-readable change summary

  // Compliance
  ComplianceRelevant: boolean;
  RegulatoryFramework?: string; // GDPR, SOX, HIPAA, etc.
  DataClassification?: string;
  RetentionCategory?: string;

  // Timestamps
  EventTimestamp: Date;
  RecordedTimestamp?: Date;

  // Additional context
  Metadata?: string; // JSON for additional data
  Notes?: string;
  ParentEventId?: string; // For related events
}

/**
 * Audit query filters
 */
export interface IAuditQueryFilters {
  eventTypes?: AuditEventType[];
  entityTypes?: string[];
  policyIds?: number[];
  userIds?: number[];
  severities?: AuditSeverity[];
  dateFrom?: Date;
  dateTo?: Date;
  searchText?: string;
  complianceRelevantOnly?: boolean;
}

/**
 * Audit report summary
 */
export interface IAuditSummary {
  totalEvents: number;
  eventsByType: { type: string; count: number }[];
  eventsBySeverity: { severity: string; count: number }[];
  eventsByUser: { userId: number; userName: string; count: number }[];
  eventsByPolicy: { policyId: number; policyName: string; count: number }[];
  securityEvents: number;
  failedOperations: number;
  dateRange: { from: Date; to: Date };
}

// ============================================================================
// SERVICE CLASS
// ============================================================================

export class PolicyAuditService {
  private sp: SPFI;
  private currentUserId: number = 0;
  private currentUserEmail: string = '';
  private currentUserName: string = '';
  private sessionId: string = '';
  private initialized: boolean = false;

  private readonly AUDIT_LOG_LIST = PolicyLists.POLICY_AUDIT_LOG;
  private readonly AUDIT_ARCHIVE_LIST = SystemLists.AUDIT_ARCHIVE;

  constructor(sp: SPFI) {
    this.sp = sp;
    this.sessionId = this.generateSessionId();
  }

  // ============================================================================
  // INITIALIZATION
  // ============================================================================

  /**
   * Initialize the audit service
   */
  public async initialize(): Promise<void> {
    if (this.initialized) return;

    try {
      const user = await this.sp.web.currentUser();
      this.currentUserId = user.Id;
      this.currentUserEmail = user.Email;
      this.currentUserName = user.Title;
      this.initialized = true;
    } catch (error) {
      logger.error('PolicyAuditService', 'Failed to initialize:', error);
      throw error;
    }
  }

  /**
   * Generate unique session ID
   */
  private generateSessionId(): string {
    return `${Date.now()}-${Math.random().toString(36).substring(2, 15)}`;
  }

  /**
   * Generate unique event ID
   */
  private generateEventId(): string {
    return `EVT-${Date.now()}-${Math.random().toString(36).substring(2, 9)}`;
  }

  // ============================================================================
  // AUDIT LOGGING METHODS
  // ============================================================================

  /**
   * Log an audit event
   */
  public async logEvent(entry: Partial<IAuditEntry>): Promise<string> {
    await this.initialize();

    const eventId = this.generateEventId();
    const fullEntry: IAuditEntry = {
      EventId: eventId,
      EventType: entry.EventType || AuditEventType.PolicyViewed,
      EventCategory: this.getEventCategory(entry.EventType || AuditEventType.PolicyViewed),
      Severity: entry.Severity || AuditSeverity.Info,
      EntityType: entry.EntityType || 'Policy',
      EntityId: entry.EntityId || 0,
      EntityName: entry.EntityName,
      PolicyId: entry.PolicyId,
      PolicyNumber: entry.PolicyNumber,
      PolicyName: entry.PolicyName,
      Action: entry.Action || entry.EventType || 'Unknown',
      ActionDescription: entry.ActionDescription || '',
      ActionResult: entry.ActionResult || 'Success',
      PerformedById: entry.PerformedById || this.currentUserId,
      PerformedByEmail: entry.PerformedByEmail || this.currentUserEmail,
      PerformedByName: entry.PerformedByName || this.currentUserName,
      PerformedByRoles: entry.PerformedByRoles,
      TargetUserId: entry.TargetUserId,
      TargetUserEmail: entry.TargetUserEmail,
      TargetUserName: entry.TargetUserName,
      IPAddress: entry.IPAddress,
      UserAgent: entry.UserAgent || (typeof navigator !== 'undefined' ? navigator.userAgent : undefined),
      DeviceType: entry.DeviceType || this.detectDeviceType(),
      SessionId: this.sessionId,
      RequestId: entry.RequestId,
      PreviousValue: entry.PreviousValue,
      NewValue: entry.NewValue,
      ChangedFields: entry.ChangedFields,
      ChangeDetails: entry.ChangeDetails,
      ComplianceRelevant: entry.ComplianceRelevant ?? this.isComplianceRelevant(entry.EventType || AuditEventType.PolicyViewed),
      RegulatoryFramework: entry.RegulatoryFramework,
      DataClassification: entry.DataClassification,
      RetentionCategory: entry.RetentionCategory,
      EventTimestamp: entry.EventTimestamp || new Date(),
      RecordedTimestamp: new Date(),
      Metadata: entry.Metadata,
      Notes: entry.Notes,
      ParentEventId: entry.ParentEventId
    };

    try {
      await this.sp.web.lists
        .getByTitle(this.AUDIT_LOG_LIST)
        .items.add({
          Title: eventId,
          EventId: fullEntry.EventId,
          EventType: fullEntry.EventType,
          EventCategory: fullEntry.EventCategory,
          Severity: fullEntry.Severity,
          EntityType: fullEntry.EntityType,
          EntityId: fullEntry.EntityId,
          EntityName: fullEntry.EntityName,
          PolicyId: fullEntry.PolicyId,
          PolicyNumber: fullEntry.PolicyNumber,
          PolicyName: fullEntry.PolicyName,
          Action: fullEntry.Action,
          ActionDescription: fullEntry.ActionDescription,
          ActionResult: fullEntry.ActionResult,
          PerformedById: fullEntry.PerformedById,
          PerformedByEmail: fullEntry.PerformedByEmail,
          PerformedByName: fullEntry.PerformedByName,
          PerformedByRoles: fullEntry.PerformedByRoles,
          TargetUserId: fullEntry.TargetUserId,
          TargetUserEmail: fullEntry.TargetUserEmail,
          TargetUserName: fullEntry.TargetUserName,
          IPAddress: fullEntry.IPAddress,
          UserAgent: fullEntry.UserAgent,
          DeviceType: fullEntry.DeviceType,
          SessionId: fullEntry.SessionId,
          RequestId: fullEntry.RequestId,
          PreviousValue: fullEntry.PreviousValue,
          NewValue: fullEntry.NewValue,
          ChangedFields: fullEntry.ChangedFields,
          ChangeDetails: fullEntry.ChangeDetails,
          ComplianceRelevant: fullEntry.ComplianceRelevant,
          RegulatoryFramework: fullEntry.RegulatoryFramework,
          DataClassification: fullEntry.DataClassification,
          RetentionCategory: fullEntry.RetentionCategory,
          EventTimestamp: fullEntry.EventTimestamp.toISOString(),
          RecordedTimestamp: fullEntry.RecordedTimestamp?.toISOString(),
          Metadata: fullEntry.Metadata,
          Notes: fullEntry.Notes,
          ParentEventId: fullEntry.ParentEventId
        });

      logger.info('PolicyAuditService', `Audit event logged: ${eventId} - ${fullEntry.EventType}`);
      return eventId;

    } catch (error) {
      logger.error('PolicyAuditService', 'Failed to log audit event:', error);
      // Don't throw - audit logging should not break main operations
      return eventId;
    }
  }

  // ============================================================================
  // SPECIFIC EVENT LOGGING METHODS
  // ============================================================================

  /**
   * Log policy approval
   */
  public async logPolicyApproval(
    policy: IPolicy,
    comments?: string
  ): Promise<string> {
    return this.logEvent({
      EventType: AuditEventType.PolicyApproved,
      Severity: AuditSeverity.Info,
      EntityType: 'Policy',
      EntityId: policy.Id!,
      EntityName: policy.PolicyName,
      PolicyId: policy.Id,
      PolicyNumber: policy.PolicyNumber,
      PolicyName: policy.PolicyName,
      ActionDescription: `Policy "${policy.PolicyName}" (${policy.PolicyNumber}) was approved`,
      ComplianceRelevant: true,
      Notes: comments,
      NewValue: JSON.stringify({ status: PolicyStatus.Approved })
    });
  }

  /**
   * Log policy rejection
   */
  public async logPolicyRejection(
    policy: IPolicy,
    reason: string
  ): Promise<string> {
    return this.logEvent({
      EventType: AuditEventType.PolicyRejected,
      Severity: AuditSeverity.Warning,
      EntityType: 'Policy',
      EntityId: policy.Id!,
      EntityName: policy.PolicyName,
      PolicyId: policy.Id,
      PolicyNumber: policy.PolicyNumber,
      PolicyName: policy.PolicyName,
      ActionDescription: `Policy "${policy.PolicyName}" (${policy.PolicyNumber}) was rejected`,
      ComplianceRelevant: true,
      Notes: reason,
      NewValue: JSON.stringify({ status: PolicyStatus.Draft, rejectionReason: reason })
    });
  }

  /**
   * Log policy acknowledgement
   */
  public async logAcknowledgement(
    acknowledgement: IPolicyAcknowledgement,
    policy: IPolicy,
    method: string
  ): Promise<string> {
    return this.logEvent({
      EventType: AuditEventType.AcknowledgementCompleted,
      Severity: AuditSeverity.Info,
      EntityType: 'Acknowledgement',
      EntityId: acknowledgement.Id!,
      PolicyId: policy.Id,
      PolicyNumber: policy.PolicyNumber,
      PolicyName: policy.PolicyName,
      TargetUserId: acknowledgement.AckUserId,
      TargetUserEmail: acknowledgement.UserEmail,
      ActionDescription: `User acknowledged policy "${policy.PolicyName}" (${policy.PolicyNumber})`,
      ComplianceRelevant: true,
      Metadata: JSON.stringify({
        method: method,
        quizRequired: acknowledgement.QuizRequired,
        quizScore: acknowledgement.QuizScore,
        readTime: acknowledgement.TotalReadTimeSeconds
      })
    });
  }

  /**
   * Log policy publish
   */
  public async logPolicyPublish(
    policy: IPolicy,
    distributionDetails: { totalRecipients: number; scope: string }
  ): Promise<string> {
    return this.logEvent({
      EventType: AuditEventType.PolicyPublished,
      Severity: AuditSeverity.Info,
      EntityType: 'Policy',
      EntityId: policy.Id!,
      EntityName: policy.PolicyName,
      PolicyId: policy.Id,
      PolicyNumber: policy.PolicyNumber,
      PolicyName: policy.PolicyName,
      ActionDescription: `Policy "${policy.PolicyName}" was published to ${distributionDetails.totalRecipients} recipients`,
      ComplianceRelevant: true,
      Metadata: JSON.stringify(distributionDetails)
    });
  }

  /**
   * Log policy archive
   */
  public async logPolicyArchive(
    policy: IPolicy,
    reason: string
  ): Promise<string> {
    return this.logEvent({
      EventType: AuditEventType.PolicyArchived,
      Severity: AuditSeverity.Info,
      EntityType: 'Policy',
      EntityId: policy.Id!,
      EntityName: policy.PolicyName,
      PolicyId: policy.Id,
      PolicyNumber: policy.PolicyNumber,
      PolicyName: policy.PolicyName,
      ActionDescription: `Policy "${policy.PolicyName}" was archived`,
      ComplianceRelevant: true,
      Notes: reason
    });
  }

  /**
   * Log policy update with change tracking
   */
  public async logPolicyUpdate(
    policy: IPolicy,
    previousState: Partial<IPolicy>,
    changes: string[]
  ): Promise<string> {
    return this.logEvent({
      EventType: AuditEventType.PolicyUpdated,
      Severity: AuditSeverity.Info,
      EntityType: 'Policy',
      EntityId: policy.Id!,
      EntityName: policy.PolicyName,
      PolicyId: policy.Id,
      PolicyNumber: policy.PolicyNumber,
      PolicyName: policy.PolicyName,
      ActionDescription: `Policy "${policy.PolicyName}" was updated: ${changes.join(', ')}`,
      ComplianceRelevant: changes.some(c =>
        ['PolicyContent', 'EffectiveDate', 'ComplianceRisk', 'RequiresAcknowledgement'].includes(c)
      ),
      PreviousValue: JSON.stringify(previousState),
      ChangedFields: JSON.stringify(changes),
      ChangeDetails: this.generateChangeDetails(previousState, policy, changes)
    });
  }

  /**
   * Log exemption approval
   */
  public async logExemptionApproval(
    policyId: number,
    policyName: string,
    userId: number,
    userEmail: string,
    reason: string,
    expiryDate?: Date
  ): Promise<string> {
    return this.logEvent({
      EventType: AuditEventType.ExemptionApproved,
      Severity: AuditSeverity.Warning,
      EntityType: 'Exemption',
      EntityId: policyId,
      PolicyId: policyId,
      PolicyName: policyName,
      TargetUserId: userId,
      TargetUserEmail: userEmail,
      ActionDescription: `Exemption approved for user ${userEmail} on policy "${policyName}"`,
      ComplianceRelevant: true,
      Notes: reason,
      Metadata: JSON.stringify({ expiryDate: expiryDate?.toISOString() })
    });
  }

  /**
   * Log delegation
   */
  public async logDelegation(
    policyId: number,
    policyName: string,
    fromUserId: number,
    fromUserEmail: string,
    toUserId: number,
    toUserEmail: string,
    reason: string
  ): Promise<string> {
    return this.logEvent({
      EventType: AuditEventType.DelegationApproved,
      Severity: AuditSeverity.Info,
      EntityType: 'Acknowledgement',
      EntityId: policyId,
      PolicyId: policyId,
      PolicyName: policyName,
      TargetUserId: fromUserId,
      TargetUserEmail: fromUserEmail,
      ActionDescription: `Policy acknowledgement delegated from ${fromUserEmail} to ${toUserEmail}`,
      ComplianceRelevant: true,
      Notes: reason,
      Metadata: JSON.stringify({
        delegatedToId: toUserId,
        delegatedToEmail: toUserEmail
      })
    });
  }

  /**
   * Log security event
   */
  public async logSecurityEvent(
    eventType: AuditEventType,
    details: string,
    entityId?: number,
    metadata?: object
  ): Promise<string> {
    return this.logEvent({
      EventType: eventType,
      Severity: AuditSeverity.Security,
      EntityType: 'System',
      EntityId: entityId || 0,
      ActionDescription: details,
      ActionResult: 'Failure',
      ComplianceRelevant: true,
      Metadata: metadata ? JSON.stringify(metadata) : undefined
    });
  }

  /**
   * Log bulk operation
   */
  public async logBulkOperation(
    operation: string,
    policyIds: number[],
    successCount: number,
    failCount: number
  ): Promise<string> {
    return this.logEvent({
      EventType: AuditEventType.BulkOperation,
      Severity: failCount > 0 ? AuditSeverity.Warning : AuditSeverity.Info,
      EntityType: 'System',
      EntityId: 0,
      ActionDescription: `Bulk ${operation}: ${successCount} succeeded, ${failCount} failed`,
      ActionResult: failCount === 0 ? 'Success' : (successCount > 0 ? 'Partial' : 'Failure'),
      ComplianceRelevant: true,
      Metadata: JSON.stringify({
        policyIds,
        successCount,
        failCount,
        totalCount: policyIds.length
      })
    });
  }

  // ============================================================================
  // AUDIT QUERY METHODS
  // ============================================================================

  /**
   * Query audit logs
   */
  public async queryAuditLogs(
    filters: IAuditQueryFilters,
    page: number = 1,
    pageSize: number = 50
  ): Promise<{ entries: IAuditEntry[]; totalCount: number }> {
    try {
      let filterConditions: string[] = [];

      if (filters.eventTypes?.length) {
        const typeFilter = filters.eventTypes.map(t => `EventType eq '${t}'`).join(' or ');
        filterConditions.push(`(${typeFilter})`);
      }

      if (filters.entityTypes?.length) {
        const entityFilter = filters.entityTypes.map(t => `EntityType eq '${t}'`).join(' or ');
        filterConditions.push(`(${entityFilter})`);
      }

      if (filters.policyIds?.length) {
        const policyFilter = filters.policyIds.map(id => `PolicyId eq ${id}`).join(' or ');
        filterConditions.push(`(${policyFilter})`);
      }

      if (filters.userIds?.length) {
        const userFilter = filters.userIds.map(id => `PerformedById eq ${id}`).join(' or ');
        filterConditions.push(`(${userFilter})`);
      }

      if (filters.severities?.length) {
        const sevFilter = filters.severities.map(s => `Severity eq '${s}'`).join(' or ');
        filterConditions.push(`(${sevFilter})`);
      }

      if (filters.dateFrom) {
        filterConditions.push(`EventTimestamp ge datetime'${filters.dateFrom.toISOString()}'`);
      }

      if (filters.dateTo) {
        filterConditions.push(`EventTimestamp le datetime'${filters.dateTo.toISOString()}'`);
      }

      if (filters.complianceRelevantOnly) {
        filterConditions.push('ComplianceRelevant eq true');
      }

      let query = this.sp.web.lists
        .getByTitle(this.AUDIT_LOG_LIST)
        .items.orderBy('EventTimestamp', false);

      if (filterConditions.length > 0) {
        query = query.filter(filterConditions.join(' and '));
      }

      // Get total count
      const allItems = await query.top(500)();
      const totalCount = allItems.length;

      // Apply pagination
      const startIndex = (page - 1) * pageSize;
      const entries = allItems.slice(startIndex, startIndex + pageSize) as IAuditEntry[];

      return { entries, totalCount };

    } catch (error) {
      logger.error('PolicyAuditService', 'Failed to query audit logs:', error);
      return { entries: [], totalCount: 0 };
    }
  }

  /**
   * Get audit logs for a specific policy
   */
  public async getPolicyAuditTrail(policyId: number): Promise<IAuditEntry[]> {
    const result = await this.queryAuditLogs({ policyIds: [policyId] }, 1, 1000);
    return result.entries;
  }

  /**
   * Get audit logs for a specific user
   */
  public async getUserAuditTrail(userId: number): Promise<IAuditEntry[]> {
    const result = await this.queryAuditLogs({ userIds: [userId] }, 1, 1000);
    return result.entries;
  }

  /**
   * Get compliance-relevant audit events
   */
  public async getComplianceAuditTrail(
    dateFrom: Date,
    dateTo: Date
  ): Promise<IAuditEntry[]> {
    const result = await this.queryAuditLogs({
      complianceRelevantOnly: true,
      dateFrom,
      dateTo
    }, 1, 5000);
    return result.entries;
  }

  /**
   * Get security events
   */
  public async getSecurityEvents(dateFrom?: Date): Promise<IAuditEntry[]> {
    const result = await this.queryAuditLogs({
      severities: [AuditSeverity.Security],
      dateFrom: dateFrom || new Date(Date.now() - 30 * 24 * 60 * 60 * 1000) // Last 30 days
    }, 1, 1000);
    return result.entries;
  }

  /**
   * Generate audit summary
   */
  public async generateAuditSummary(
    dateFrom: Date,
    dateTo: Date
  ): Promise<IAuditSummary> {
    try {
      const result = await this.queryAuditLogs({ dateFrom, dateTo }, 1, 10000);
      const entries = result.entries;

      // Events by type
      const typeMap = new Map<string, number>();
      entries.forEach(e => {
        typeMap.set(e.EventType, (typeMap.get(e.EventType) || 0) + 1);
      });
      const eventsByType = Array.from(typeMap.entries())
        .map(([type, count]) => ({ type, count }))
        .sort((a, b) => b.count - a.count);

      // Events by severity
      const sevMap = new Map<string, number>();
      entries.forEach(e => {
        sevMap.set(e.Severity, (sevMap.get(e.Severity) || 0) + 1);
      });
      const eventsBySeverity = Array.from(sevMap.entries())
        .map(([severity, count]) => ({ severity, count }));

      // Events by user
      const userMap = new Map<number, { name: string; count: number }>();
      entries.forEach(e => {
        const existing = userMap.get(e.PerformedById) || { name: e.PerformedByName || e.PerformedByEmail, count: 0 };
        existing.count++;
        userMap.set(e.PerformedById, existing);
      });
      const eventsByUser = Array.from(userMap.entries())
        .map(([userId, data]) => ({ userId, userName: data.name, count: data.count }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 20);

      // Events by policy
      const policyMap = new Map<number, { name: string; count: number }>();
      entries.filter(e => e.PolicyId).forEach(e => {
        const existing = policyMap.get(e.PolicyId!) || { name: e.PolicyName || 'Unknown', count: 0 };
        existing.count++;
        policyMap.set(e.PolicyId!, existing);
      });
      const eventsByPolicy = Array.from(policyMap.entries())
        .map(([policyId, data]) => ({ policyId, policyName: data.name, count: data.count }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 20);

      // Security and failure counts
      const securityEvents = entries.filter(e => e.Severity === AuditSeverity.Security).length;
      const failedOperations = entries.filter(e => e.ActionResult === 'Failure').length;

      return {
        totalEvents: entries.length,
        eventsByType,
        eventsBySeverity,
        eventsByUser,
        eventsByPolicy,
        securityEvents,
        failedOperations,
        dateRange: { from: dateFrom, to: dateTo }
      };

    } catch (error) {
      logger.error('PolicyAuditService', 'Failed to generate audit summary:', error);
      throw error;
    }
  }

  // ============================================================================
  // UTILITY METHODS
  // ============================================================================

  /**
   * Get event category from event type
   */
  private getEventCategory(eventType: AuditEventType): string {
    if (eventType.startsWith('Policy.')) return 'Policy';
    if (eventType.startsWith('Version.')) return 'Version';
    if (eventType.startsWith('Distribution.')) return 'Distribution';
    if (eventType.startsWith('Acknowledgement.')) return 'Acknowledgement';
    if (eventType.startsWith('Delegation.')) return 'Delegation';
    if (eventType.startsWith('Exemption.')) return 'Exemption';
    if (eventType.startsWith('Quiz.')) return 'Quiz';
    if (eventType.startsWith('Admin.')) return 'Administration';
    if (eventType.startsWith('Security.')) return 'Security';
    return 'Other';
  }

  /**
   * Determine if event type is compliance relevant
   */
  private isComplianceRelevant(eventType: AuditEventType): boolean {
    const complianceEvents = [
      AuditEventType.PolicyApproved,
      AuditEventType.PolicyRejected,
      AuditEventType.PolicyPublished,
      AuditEventType.PolicyArchived,
      AuditEventType.PolicyRetired,
      AuditEventType.AcknowledgementCompleted,
      AuditEventType.AcknowledgementOverdue,
      AuditEventType.ExemptionApproved,
      AuditEventType.ExemptionRejected,
      AuditEventType.DelegationApproved,
      AuditEventType.QuizPassed,
      AuditEventType.QuizFailed,
      AuditEventType.UnauthorizedAccess,
      AuditEventType.DataPurged
    ];
    return complianceEvents.includes(eventType);
  }

  /**
   * Detect device type from user agent
   */
  private detectDeviceType(): string {
    if (typeof navigator === 'undefined') return 'Unknown';

    const ua = navigator.userAgent.toLowerCase();
    if (/mobile|android|iphone|ipad|tablet/i.test(ua)) {
      if (/tablet|ipad/i.test(ua)) return 'Tablet';
      return 'Mobile';
    }
    return 'Desktop';
  }

  /**
   * Generate human-readable change details
   */
  private generateChangeDetails(
    previous: Partial<IPolicy>,
    current: IPolicy,
    changedFields: string[]
  ): string {
    const changes: string[] = [];

    for (const field of changedFields) {
      const prevValue = (previous as any)[field];
      const currValue = (current as any)[field];

      if (field === 'Status') {
        changes.push(`Status changed from "${prevValue}" to "${currValue}"`);
      } else if (field === 'ComplianceRisk') {
        changes.push(`Compliance risk changed from "${prevValue}" to "${currValue}"`);
      } else if (field === 'VersionNumber') {
        changes.push(`Version updated from "${prevValue}" to "${currValue}"`);
      } else {
        changes.push(`${field} was modified`);
      }
    }

    return changes.join('; ');
  }

  /**
   * Export audit logs to JSON
   */
  public async exportAuditLogs(
    filters: IAuditQueryFilters
  ): Promise<string> {
    const result = await this.queryAuditLogs(filters, 1, 10000);

    // Log the export
    await this.logEvent({
      EventType: AuditEventType.DataExported,
      Severity: AuditSeverity.Info,
      EntityType: 'System',
      EntityId: 0,
      ActionDescription: `Exported ${result.entries.length} audit log entries`,
      ComplianceRelevant: true,
      Metadata: JSON.stringify(filters)
    });

    return JSON.stringify(result.entries, null, 2);
  }

  /**
   * Archive old audit logs
   */
  public async archiveOldLogs(olderThanDays: number = 365): Promise<number> {
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - olderThanDays);

    try {
      const oldEntries = await this.sp.web.lists
        .getByTitle(this.AUDIT_LOG_LIST)
        .items.filter(`EventTimestamp lt datetime'${cutoffDate.toISOString()}'`)
        .top(1000)();

      let archivedCount = 0;
      for (const entry of oldEntries) {
        // Copy to archive
        await this.sp.web.lists
          .getByTitle(this.AUDIT_ARCHIVE_LIST)
          .items.add(entry);

        // Delete from main log
        await this.sp.web.lists
          .getByTitle(this.AUDIT_LOG_LIST)
          .items.getById(entry.Id)
          .delete();

        archivedCount++;
      }

      // Log the archival
      await this.logEvent({
        EventType: AuditEventType.DataPurged,
        Severity: AuditSeverity.Info,
        EntityType: 'System',
        EntityId: 0,
        ActionDescription: `Archived ${archivedCount} audit log entries older than ${olderThanDays} days`,
        ComplianceRelevant: true
      });

      return archivedCount;

    } catch (error) {
      logger.error('PolicyAuditService', 'Failed to archive old logs:', error);
      throw error;
    }
  }
}
