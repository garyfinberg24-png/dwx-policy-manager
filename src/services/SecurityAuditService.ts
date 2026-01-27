// @ts-nocheck
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { UserRole } from './RoleDetectionService';

/**
 * Security audit event types
 */
export enum SecurityAuditEventType {
  UnauthorizedRoleAccess = 'Unauthorized Role Access',
  UnauthorizedAppAccess = 'Unauthorized App Access',
  RoleSwitched = 'Role Switched',
  AppAccessDenied = 'App Access Denied',
  SuspiciousActivity = 'Suspicious Activity'
}

/**
 * Security audit event severity levels
 */
export enum SecurityAuditSeverity {
  Low = 'Low',
  Medium = 'Medium',
  High = 'High',
  Critical = 'Critical'
}

/**
 * Security audit log entry interface
 */
export interface ISecurityAuditEntry {
  EventType: string;
  Severity: string;
  UserEmail: string;
  UserDisplayName: string;
  AttemptedRole?: string;
  ActualRoles?: string;
  AttemptedApp?: string;
  IPAddress?: string;
  UserAgent?: string;
  Details: string;
  Timestamp: Date;
}

/**
 * Service for logging security audit events to SharePoint
 */
export class SecurityAuditService {
  private sp: SPFI;
  private listName: string = 'JML Security Audit Log';
  private enabled: boolean = true; // Can be configured via Admin Panel

  constructor(sp: SPFI, enabled: boolean = true) {
    this.sp = sp;
    this.enabled = enabled;
  }

  /**
   * Log unauthorized role access attempt
   */
  public async logUnauthorizedRoleAccess(
    userEmail: string,
    userDisplayName: string,
    attemptedRole: UserRole,
    actualRoles: UserRole[]
  ): Promise<void> {
    if (!this.enabled) return;

    const entry: ISecurityAuditEntry = {
      EventType: SecurityAuditEventType.UnauthorizedRoleAccess,
      Severity: SecurityAuditSeverity.High,
      UserEmail: userEmail,
      UserDisplayName: userDisplayName,
      AttemptedRole: attemptedRole,
      ActualRoles: actualRoles.join(', '),
      Details: `User attempted to access role "${attemptedRole}" but only has roles: ${actualRoles.join(', ')}`,
      Timestamp: new Date(),
      IPAddress: await this.getClientIP(),
      UserAgent: navigator.userAgent
    };

    await this.logEvent(entry);
    console.warn('[SecurityAudit] Unauthorized role access attempt logged:', entry);
  }

  /**
   * Log unauthorized app access attempt
   */
  public async logUnauthorizedAppAccess(
    userEmail: string,
    userDisplayName: string,
    appName: string,
    userRole: UserRole
  ): Promise<void> {
    if (!this.enabled) return;

    const entry: ISecurityAuditEntry = {
      EventType: SecurityAuditEventType.UnauthorizedAppAccess,
      Severity: SecurityAuditSeverity.Medium,
      UserEmail: userEmail,
      UserDisplayName: userDisplayName,
      AttemptedApp: appName,
      ActualRoles: userRole,
      Details: `User with role "${userRole}" attempted to access app "${appName}" without permission`,
      Timestamp: new Date(),
      IPAddress: await this.getClientIP(),
      UserAgent: navigator.userAgent
    };

    await this.logEvent(entry);
    console.warn('[SecurityAudit] Unauthorized app access attempt logged:', entry);
  }

  /**
   * Log successful role switch
   */
  public async logRoleSwitch(
    userEmail: string,
    userDisplayName: string,
    fromRole: UserRole,
    toRole: UserRole
  ): Promise<void> {
    if (!this.enabled) return;

    const entry: ISecurityAuditEntry = {
      EventType: SecurityAuditEventType.RoleSwitched,
      Severity: SecurityAuditSeverity.Low,
      UserEmail: userEmail,
      UserDisplayName: userDisplayName,
      Details: `User switched from role "${fromRole}" to "${toRole}"`,
      Timestamp: new Date()
    };

    await this.logEvent(entry);
    console.log('[SecurityAudit] Role switch logged:', entry);
  }

  /**
   * Log app access denial
   */
  public async logAppAccessDenied(
    userEmail: string,
    userDisplayName: string,
    appName: string,
    userRole: UserRole,
    reason: string
  ): Promise<void> {
    if (!this.enabled) return;

    const entry: ISecurityAuditEntry = {
      EventType: SecurityAuditEventType.AppAccessDenied,
      Severity: SecurityAuditSeverity.Low,
      UserEmail: userEmail,
      UserDisplayName: userDisplayName,
      AttemptedApp: appName,
      ActualRoles: userRole,
      Details: `Access denied to app "${appName}". Reason: ${reason}`,
      Timestamp: new Date()
    };

    await this.logEvent(entry);
    console.info('[SecurityAudit] App access denied logged:', entry);
  }

  /**
   * Log suspicious activity
   */
  public async logSuspiciousActivity(
    userEmail: string,
    userDisplayName: string,
    details: string,
    severity: SecurityAuditSeverity = SecurityAuditSeverity.High
  ): Promise<void> {
    if (!this.enabled) return;

    const entry: ISecurityAuditEntry = {
      EventType: SecurityAuditEventType.SuspiciousActivity,
      Severity: severity,
      UserEmail: userEmail,
      UserDisplayName: userDisplayName,
      Details: details,
      Timestamp: new Date(),
      IPAddress: await this.getClientIP(),
      UserAgent: navigator.userAgent
    };

    await this.logEvent(entry);
    console.error('[SecurityAudit] Suspicious activity logged:', entry);
  }

  /**
   * Write audit entry to SharePoint list
   */
  private async logEvent(entry: ISecurityAuditEntry): Promise<void> {
    try {
      // Ensure list exists
      await this.ensureAuditList();

      // Add item to list
      await this.sp.web.lists.getByTitle(this.listName).items.add({
        Title: `${entry.EventType} - ${entry.UserDisplayName}`,
        EventType: entry.EventType,
        Severity: entry.Severity,
        UserEmail: entry.UserEmail,
        UserDisplayName: entry.UserDisplayName,
        AttemptedRole: entry.AttemptedRole || '',
        ActualRoles: entry.ActualRoles || '',
        AttemptedApp: entry.AttemptedApp || '',
        IPAddress: entry.IPAddress || '',
        UserAgent: entry.UserAgent || '',
        Details: entry.Details,
        AuditTimestamp: entry.Timestamp.toISOString()
      });

      console.log('[SecurityAudit] Event successfully logged to SharePoint');
    } catch (error) {
      // Don't let audit failures break the application
      console.error('[SecurityAudit] Failed to log event to SharePoint:', error);
      console.error('[SecurityAudit] Event details:', entry);
    }
  }

  /**
   * Ensure the Security Audit Log list exists
   */
  private async ensureAuditList(): Promise<void> {
    try {
      // Check if list exists
      await this.sp.web.lists.getByTitle(this.listName)();
    } catch {
      // List doesn't exist, create it
      console.log('[SecurityAudit] Creating Security Audit Log list...');
      try {
        await this.sp.web.lists.add(this.listName, 'Security audit log for JML solution', 100, false);

        // Add custom fields
        const list = this.sp.web.lists.getByTitle(this.listName);
        await list.fields.addText('EventType', { MaxLength: 100 });
        await list.fields.addChoice('Severity', { Choices: ['Low', 'Medium', 'High', 'Critical'] });
        await list.fields.addText('UserEmail', { MaxLength: 255 });
        await list.fields.addText('UserDisplayName', { MaxLength: 255 });
        await list.fields.addText('AttemptedRole', { MaxLength: 100 });
        await list.fields.addText('ActualRoles', { MaxLength: 500 });
        await list.fields.addText('AttemptedApp', { MaxLength: 255 });
        await list.fields.addText('IPAddress', { MaxLength: 50 });
        await list.fields.addText('UserAgent', { MaxLength: 500 });
        await list.fields.addMultilineText('Details', { NumberOfLines: 6, RichText: true, RestrictedMode: false });
        await list.fields.addDateTime('AuditTimestamp');

        console.log('[SecurityAudit] Security Audit Log list created successfully');
      } catch (createError) {
        console.error('[SecurityAudit] Failed to create audit list:', createError);
      }
    }
  }

  /**
   * Get client IP address (best effort)
   */
  private async getClientIP(): Promise<string> {
    try {
      // This is client-side, so we can only get approximate IP
      // In production, you'd want to log this server-side
      return 'Client-side (not available)';
    } catch {
      return 'Unknown';
    }
  }

  /**
   * Enable or disable audit logging
   */
  public setEnabled(enabled: boolean): void {
    this.enabled = enabled;
    console.log(`[SecurityAudit] Audit logging ${enabled ? 'enabled' : 'disabled'}`);
  }

  /**
   * Get recent audit events (for admin viewing)
   */
  public async getRecentEvents(maxResults: number = 50): Promise<ISecurityAuditEntry[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.listName)
        .items
        .select('Title', 'EventType', 'Severity', 'UserEmail', 'UserDisplayName',
                'AttemptedRole', 'ActualRoles', 'AttemptedApp', 'IPAddress',
                'UserAgent', 'Details', 'AuditTimestamp', 'Created')
        .orderBy('Created', false)
        .top(maxResults)();

      return items.map(item => ({
        EventType: item.EventType,
        Severity: item.Severity,
        UserEmail: item.UserEmail,
        UserDisplayName: item.UserDisplayName,
        AttemptedRole: item.AttemptedRole,
        ActualRoles: item.ActualRoles,
        AttemptedApp: item.AttemptedApp,
        IPAddress: item.IPAddress,
        UserAgent: item.UserAgent,
        Details: item.Details,
        Timestamp: new Date(item.AuditTimestamp || item.Created)
      }));
    } catch (error) {
      console.error('[SecurityAudit] Failed to retrieve audit events:', error);
      return [];
    }
  }
}
