// @ts-nocheck
/**
 * SecurityAuditService — Enhanced security audit logging for Policy Manager.
 *
 * Tracks unauthorized access attempts, role switches, suspicious activity,
 * and app access denials. All events written to PM_SecurityAuditLog with
 * severity levels, risk scoring, IP tracking, and user agent capture.
 *
 * Adapted from JML SecurityAuditService + ExternalSharingAuditService
 * for the DWx Policy Manager suite.
 */
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { logger } from './LoggingService';

export enum SecurityEventType {
  UnauthorizedRoleAccess = 'Unauthorized Role Access',
  UnauthorizedAppAccess = 'Unauthorized App Access',
  RoleSwitched = 'Role Switched',
  AppAccessDenied = 'App Access Denied',
  SuspiciousActivity = 'Suspicious Activity',
  BulkOperationDetected = 'Bulk Operation Detected',
  ConfigurationChanged = 'Configuration Changed',
  DataExported = 'Data Exported',
  FailedAuthentication = 'Failed Authentication',
  PolicyViolation = 'Policy Violation'
}

export enum SecuritySeverity {
  Low = 'Low',
  Medium = 'Medium',
  High = 'High',
  Critical = 'Critical'
}

export interface ISecurityAuditEntry {
  Id?: number;
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
  RiskScore?: number;
  SessionId?: string;
  Timestamp: Date;
}

export interface ISecurityAlert {
  id: number;
  title: string;
  description: string;
  severity: SecuritySeverity;
  category: 'access' | 'data' | 'identity' | 'compliance' | 'threat' | 'policy';
  timestamp: Date;
  status: 'new' | 'investigating' | 'resolved' | 'dismissed';
  affectedUsers?: number;
  riskScore: number;
}

export interface ISecuritySummary {
  totalEvents: number;
  criticalCount: number;
  highCount: number;
  mediumCount: number;
  lowCount: number;
  avgRiskScore: number;
  topEventTypes: { type: string; count: number }[];
  recentAlerts: ISecurityAlert[];
}

const LIST_NAME = 'PM_SecurityAuditLog';
const ALERTS_LIST = 'PM_SecurityAlerts';

export class SecurityAuditService {
  private sp: SPFI;
  private enabled: boolean = true;
  private sessionId: string;

  constructor(sp: SPFI, enabled: boolean = true) {
    this.sp = sp;
    this.enabled = enabled;
    this.sessionId = `session_${Date.now()}_${Math.random().toString(36).substring(2, 8)}`;
  }

  public setEnabled(enabled: boolean): void {
    this.enabled = enabled;
  }

  // ─── Risk Scoring ──────────────────────────────────────────────────

  public calculateRiskScore(eventType: SecurityEventType, severity: SecuritySeverity, context?: {
    isOffHours?: boolean;
    recentFailedAttempts?: number;
    isBulkOperation?: boolean;
  }): number {
    const baseScores: Record<string, number> = {
      [SecurityEventType.UnauthorizedRoleAccess]: 60,
      [SecurityEventType.UnauthorizedAppAccess]: 50,
      [SecurityEventType.SuspiciousActivity]: 70,
      [SecurityEventType.BulkOperationDetected]: 40,
      [SecurityEventType.FailedAuthentication]: 55,
      [SecurityEventType.PolicyViolation]: 65,
      [SecurityEventType.DataExported]: 30,
      [SecurityEventType.ConfigurationChanged]: 25,
      [SecurityEventType.RoleSwitched]: 10,
      [SecurityEventType.AppAccessDenied]: 20,
    };
    let score = baseScores[eventType] || 30;
    if (severity === SecuritySeverity.Critical) score = Math.min(100, score + 20);
    else if (severity === SecuritySeverity.High) score = Math.min(100, score + 10);
    if (context?.isOffHours) score = Math.min(100, score + 10);
    if (context?.isBulkOperation) score = Math.min(100, score + 15);
    if (context?.recentFailedAttempts) score = Math.min(100, score + Math.min(30, context.recentFailedAttempts * 10));
    return score;
  }

  // ─── Event Logging ─────────────────────────────────────────────────

  public async logUnauthorizedRoleAccess(userEmail: string, userDisplayName: string, attemptedRole: string, actualRoles: string[]): Promise<void> {
    const riskScore = this.calculateRiskScore(SecurityEventType.UnauthorizedRoleAccess, SecuritySeverity.High);
    await this.logEvent({
      EventType: SecurityEventType.UnauthorizedRoleAccess, Severity: SecuritySeverity.High,
      UserEmail: userEmail, UserDisplayName: userDisplayName,
      AttemptedRole: attemptedRole, ActualRoles: actualRoles.join(', '),
      Details: `User attempted to access role "${attemptedRole}" but only has roles: ${actualRoles.join(', ')}`,
      RiskScore: riskScore, Timestamp: new Date(),
      IPAddress: this.getClientInfo(), UserAgent: typeof navigator !== 'undefined' ? navigator.userAgent : ''
    });
  }

  public async logUnauthorizedAppAccess(userEmail: string, userDisplayName: string, appName: string, userRole: string): Promise<void> {
    const riskScore = this.calculateRiskScore(SecurityEventType.UnauthorizedAppAccess, SecuritySeverity.Medium);
    await this.logEvent({
      EventType: SecurityEventType.UnauthorizedAppAccess, Severity: SecuritySeverity.Medium,
      UserEmail: userEmail, UserDisplayName: userDisplayName,
      AttemptedApp: appName, ActualRoles: userRole,
      Details: `User with role "${userRole}" attempted to access "${appName}" without permission`,
      RiskScore: riskScore, Timestamp: new Date(),
      IPAddress: this.getClientInfo(), UserAgent: typeof navigator !== 'undefined' ? navigator.userAgent : ''
    });
  }

  public async logRoleSwitch(userEmail: string, userDisplayName: string, fromRole: string, toRole: string): Promise<void> {
    await this.logEvent({
      EventType: SecurityEventType.RoleSwitched, Severity: SecuritySeverity.Low,
      UserEmail: userEmail, UserDisplayName: userDisplayName,
      Details: `User switched from role "${fromRole}" to "${toRole}"`,
      RiskScore: 10, Timestamp: new Date()
    });
  }

  public async logSuspiciousActivity(userEmail: string, userDisplayName: string, details: string, severity: SecuritySeverity = SecuritySeverity.High): Promise<void> {
    const riskScore = this.calculateRiskScore(SecurityEventType.SuspiciousActivity, severity);
    await this.logEvent({
      EventType: SecurityEventType.SuspiciousActivity, Severity: severity,
      UserEmail: userEmail, UserDisplayName: userDisplayName,
      Details: details, RiskScore: riskScore, Timestamp: new Date(),
      IPAddress: this.getClientInfo(), UserAgent: typeof navigator !== 'undefined' ? navigator.userAgent : ''
    });
    if (riskScore >= 70) {
      await this.createAlert({ title: `Suspicious Activity: ${userDisplayName}`, description: details, severity, category: 'threat', riskScore, affectedUsers: 1 });
    }
  }

  public async logBulkOperation(userEmail: string, userDisplayName: string, operation: string, itemCount: number): Promise<void> {
    const riskScore = this.calculateRiskScore(SecurityEventType.BulkOperationDetected, SecuritySeverity.Medium, { isBulkOperation: true });
    await this.logEvent({
      EventType: SecurityEventType.BulkOperationDetected,
      Severity: itemCount > 50 ? SecuritySeverity.High : SecuritySeverity.Medium,
      UserEmail: userEmail, UserDisplayName: userDisplayName,
      Details: `Bulk operation: ${operation} affecting ${itemCount} items`,
      RiskScore: riskScore, Timestamp: new Date()
    });
  }

  public async logConfigurationChange(userEmail: string, userDisplayName: string, section: string, details: string): Promise<void> {
    await this.logEvent({
      EventType: SecurityEventType.ConfigurationChanged, Severity: SecuritySeverity.Medium,
      UserEmail: userEmail, UserDisplayName: userDisplayName,
      Details: `Configuration changed: ${section} — ${details}`,
      RiskScore: 25, Timestamp: new Date()
    });
  }

  public async logDataExport(userEmail: string, userDisplayName: string, exportType: string, recordCount: number): Promise<void> {
    const severity = recordCount > 100 ? SecuritySeverity.High : SecuritySeverity.Medium;
    const riskScore = this.calculateRiskScore(SecurityEventType.DataExported, severity);
    await this.logEvent({
      EventType: SecurityEventType.DataExported, Severity: severity,
      UserEmail: userEmail, UserDisplayName: userDisplayName,
      Details: `Data exported: ${exportType} (${recordCount} records)`,
      RiskScore: riskScore, Timestamp: new Date()
    });
  }

  // ─── Event Persistence ─────────────────────────────────────────────

  private async logEvent(entry: ISecurityAuditEntry): Promise<void> {
    if (!this.enabled) return;
    try {
      await this.sp.web.lists.getByTitle(LIST_NAME).items.add({
        Title: `${entry.EventType} — ${entry.UserDisplayName || 'System'}`,
        EventType: entry.EventType,
        Severity: entry.Severity,
        UserEmail: entry.UserEmail || '',
        UserDisplayName: entry.UserDisplayName || '',
        AttemptedRole: entry.AttemptedRole || '',
        ActualRoles: entry.ActualRoles || '',
        AttemptedApp: entry.AttemptedApp || '',
        IPAddress: entry.IPAddress || '',
        UserAgent: (entry.UserAgent || '').substring(0, 500),
        Details: entry.Details,
        RiskScore: entry.RiskScore || 0,
        SessionId: this.sessionId,
        AuditTimestamp: entry.Timestamp.toISOString()
      });
    } catch (error) {
      logger.warn('SecurityAuditService', `Failed to log security event: ${entry.EventType}`, error);
    }
  }

  // ─── Alerts ────────────────────────────────────────────────────────

  private async createAlert(alert: Omit<ISecurityAlert, 'id' | 'timestamp' | 'status'>): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(ALERTS_LIST).items.add({
        Title: alert.title,
        Description: alert.description,
        Severity: alert.severity,
        Category: alert.category,
        RiskScore: alert.riskScore,
        AffectedUsers: alert.affectedUsers || 0,
        AlertStatus: 'new',
        AlertTimestamp: new Date().toISOString()
      });
    } catch { /* Alerts list may not exist yet */ }
  }

  public async getAlerts(maxResults: number = 20): Promise<ISecurityAlert[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(ALERTS_LIST)
        .items.filter("AlertStatus ne 'Resolved' and AlertStatus ne 'Dismissed'")
        .select('Id', 'Title', 'Description', 'Severity', 'Category', 'RiskScore', 'AffectedUsers', 'AlertStatus', 'AlertTimestamp', 'Created')
        .orderBy('Created', false).top(maxResults)();
      return items.map(item => ({
        id: item.Id, title: item.Title, description: item.Description || '',
        severity: item.Severity as SecuritySeverity,
        category: (item.Category || 'threat') as ISecurityAlert['category'],
        timestamp: new Date(item.AlertTimestamp || item.Created),
        status: (item.AlertStatus || 'new') as ISecurityAlert['status'],
        affectedUsers: item.AffectedUsers || 0, riskScore: item.RiskScore || 0
      }));
    } catch { return []; }
  }

  public async updateAlertStatus(alertId: number, status: 'investigating' | 'resolved' | 'dismissed'): Promise<void> {
    await this.sp.web.lists.getByTitle(ALERTS_LIST).items.getById(alertId).update({ AlertStatus: status });
  }

  // ─── Queries ───────────────────────────────────────────────────────

  public async getRecentEvents(maxResults: number = 100): Promise<ISecurityAuditEntry[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(LIST_NAME)
        .items.select('Id', 'Title', 'EventType', 'Severity', 'UserEmail', 'UserDisplayName', 'AttemptedRole', 'ActualRoles', 'AttemptedApp', 'IPAddress', 'UserAgent', 'Details', 'RiskScore', 'SessionId', 'AuditTimestamp', 'Created')
        .orderBy('Created', false).top(maxResults)();
      return items.map(item => ({
        Id: item.Id, EventType: item.EventType || '', Severity: item.Severity || 'Low',
        UserEmail: item.UserEmail || '', UserDisplayName: item.UserDisplayName || '',
        AttemptedRole: item.AttemptedRole, ActualRoles: item.ActualRoles,
        AttemptedApp: item.AttemptedApp, IPAddress: item.IPAddress, UserAgent: item.UserAgent,
        Details: item.Details || '', RiskScore: item.RiskScore || 0, SessionId: item.SessionId,
        Timestamp: new Date(item.AuditTimestamp || item.Created)
      }));
    } catch (error) {
      logger.error('SecurityAuditService', 'Failed to retrieve audit events:', error);
      return [];
    }
  }

  public async getSecuritySummary(days: number = 30): Promise<ISecuritySummary> {
    try {
      const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - days);
      const events = await this.getRecentEvents(500);
      const recent = events.filter(e => e.Timestamp >= cutoff);
      const critical = recent.filter(e => e.Severity === SecuritySeverity.Critical).length;
      const high = recent.filter(e => e.Severity === SecuritySeverity.High).length;
      const medium = recent.filter(e => e.Severity === SecuritySeverity.Medium).length;
      const low = recent.filter(e => e.Severity === SecuritySeverity.Low).length;
      const avgRisk = recent.length > 0 ? recent.reduce((sum, e) => sum + (e.RiskScore || 0), 0) / recent.length : 0;
      const typeCounts: Record<string, number> = {};
      recent.forEach(e => { typeCounts[e.EventType] = (typeCounts[e.EventType] || 0) + 1; });
      const topEventTypes = Object.entries(typeCounts).map(([type, count]) => ({ type, count })).sort((a, b) => b.count - a.count).slice(0, 5);
      const alerts = await this.getAlerts(10);
      return { totalEvents: recent.length, criticalCount: critical, highCount: high, mediumCount: medium, lowCount: low, avgRiskScore: Math.round(avgRisk), topEventTypes, recentAlerts: alerts };
    } catch {
      return { totalEvents: 0, criticalCount: 0, highCount: 0, mediumCount: 0, lowCount: 0, avgRiskScore: 0, topEventTypes: [], recentAlerts: [] };
    }
  }

  public async archiveOldEvents(olderThanDays: number = 365): Promise<number> {
    try {
      const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - olderThanDays);
      const items = await this.sp.web.lists.getByTitle(LIST_NAME)
        .items.filter(`Created lt '${cutoff.toISOString()}'`).select('Id').top(5000)();
      let archived = 0;
      for (const item of items) {
        try { await this.sp.web.lists.getByTitle(LIST_NAME).items.getById(item.Id).delete(); archived++; } catch { /* skip */ }
      }
      return archived;
    } catch { return 0; }
  }

  private getClientInfo(): string {
    try { return typeof window !== 'undefined' ? window.location.hostname : 'Server-side'; } catch { return 'Unknown'; }
  }
}
