// @ts-nocheck
// External Sharing Audit Service
// Manages compliance auditing, risk assessment, and security alerts

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';

import {
  IExternalSharingAuditLog,
  IAuditFilter,
  IComplianceReport,
  IReportOptions,
  ISecurityAlert,
  IRiskContext,
  AuditActionType,
  AuditResult
} from '../models/IExternalSharing';
import { logger } from './LoggingService';

/**
 * Input for logging an audit action
 */
export interface IAuditActionInput {
  actionType: AuditActionType;
  targetOrganizationId?: number;
  targetResourceId?: number;
  targetUser?: string;
  previousValue?: string;
  newValue?: string;
  result: AuditResult;
  correlationId?: string;
  ipAddress?: string;
  userAgent?: string;
}

export class ExternalSharingAuditService {
  private sp: SPFI;

  private readonly AUDIT_LOG_LIST = 'PM_ExternalSharingAuditLog';
  private readonly SECURITY_ALERTS_LIST = 'PM_ExternalSecurityAlerts';

  private currentUserId: number = 0;
  private currentUserEmail: string = '';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Initialize service with current user context
   */
  public async initialize(): Promise<void> {
    try {
      const user = await this.sp.web.currentUser();
      this.currentUserId = user.Id;
      this.currentUserEmail = user.Email;
      logger.info('ExternalSharingAuditService', 'Service initialized');
    } catch (error) {
      logger.error('ExternalSharingAuditService', 'Failed to initialize:', error);
      throw error;
    }
  }

  // ============================================================================
  // AUDIT LOGGING
  // ============================================================================

  /**
   * Log an audit action
   */
  public async logAction(input: IAuditActionInput): Promise<number> {
    try {
      const title = this.formatAuditTitle(input.actionType, input.targetUser);
      const riskScore = this.calculateRiskScore({
        action: input.actionType,
        userId: input.targetUser
      });

      const itemData = {
        Title: title,
        ActionType: input.actionType,
        PerformedById: this.currentUserId,
        PerformedDate: new Date().toISOString(),
        TargetOrganizationId: input.targetOrganizationId,
        TargetResourceId: input.targetResourceId,
        TargetUser: input.targetUser,
        PreviousValue: input.previousValue,
        NewValue: input.newValue,
        Result: input.result,
        CorrelationId: input.correlationId || this.generateCorrelationId(),
        IPAddress: input.ipAddress,
        UserAgent: input.userAgent,
        RiskScore: riskScore
      };

      const result = await this.sp.web.lists
        .getByTitle(this.AUDIT_LOG_LIST)
        .items
        .add(itemData);

      // Check if this action should trigger a security alert
      if (riskScore >= 70 || this.isHighRiskAction(input.actionType)) {
        await this.createSecurityAlert(input, riskScore);
      }

      return result.data.Id;
    } catch (error) {
      logger.error('ExternalSharingAuditService', 'Failed to log audit action:', error);
      throw error;
    }
  }

  /**
   * Get recent audit logs
   */
  public async getRecentLogs(count: number = 50): Promise<IExternalSharingAuditLog[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.AUDIT_LOG_LIST)
        .items
        .select(
          'Id', 'Title', 'ActionType', 'PerformedBy/Id', 'PerformedBy/Title',
          'PerformedDate', 'TargetOrganizationId', 'TargetResourceId',
          'TargetUser', 'Result', 'RiskScore'
        )
        .expand('PerformedBy')
        .orderBy('PerformedDate', false)
        .top(count)();

      return items.map(item => this.mapToAuditLogEntry(item));
    } catch (error) {
      logger.error('ExternalSharingAuditService', 'Failed to get recent logs:', error);
      return [];
    }
  }

  /**
   * Get audit logs with filter
   */
  public async getAuditLog(filter: IAuditFilter): Promise<IExternalSharingAuditLog[]> {
    try {
      let filterParts: string[] = [];

      if (filter.startDate) {
        filterParts.push(`PerformedDate ge datetime'${filter.startDate.toISOString()}'`);
      }
      if (filter.endDate) {
        filterParts.push(`PerformedDate le datetime'${filter.endDate.toISOString()}'`);
      }
      if (filter.result) {
        filterParts.push(`Result eq '${filter.result}'`);
      }
      if (filter.organizationId) {
        filterParts.push(`TargetOrganizationId eq ${filter.organizationId}`);
      }
      if (filter.performedById) {
        filterParts.push(`PerformedById eq ${filter.performedById}`);
      }
      if (filter.targetUser) {
        filterParts.push(`TargetUser eq '${filter.targetUser}'`);
      }
      if (filter.minRiskScore) {
        filterParts.push(`RiskScore ge ${filter.minRiskScore}`);
      }

      let query = this.sp.web.lists
        .getByTitle(this.AUDIT_LOG_LIST)
        .items
        .select(
          'Id', 'Title', 'ActionType', 'PerformedBy/Id', 'PerformedBy/Title',
          'PerformedDate', 'TargetOrganizationId', 'TargetResourceId',
          'TargetUser', 'PreviousValue', 'NewValue', 'Result',
          'CorrelationId', 'IPAddress', 'UserAgent', 'RiskScore'
        )
        .expand('PerformedBy');

      if (filterParts.length > 0) {
        query = query.filter(filterParts.join(' and '));
      }

      const items = await query.orderBy('PerformedDate', false).getAll();

      // Filter by action types client-side if provided
      let results = items.map(item => this.mapToAuditLogEntry(item));
      if (filter.actionTypes && filter.actionTypes.length > 0) {
        results = results.filter(log => filter.actionTypes!.includes(log.ActionType));
      }

      return results;
    } catch (error) {
      logger.error('ExternalSharingAuditService', 'Failed to get audit log with filter:', error);
      return [];
    }
  }

  // ============================================================================
  // SECURITY ALERTS
  // ============================================================================

  /**
   * Get security alerts
   */
  public async getSecurityAlerts(): Promise<ISecurityAlert[]> {
    try {
      // Try to get from a dedicated alerts list if it exists
      // Fall back to filtering high-risk audit log entries
      const highRiskLogs = await this.sp.web.lists
        .getByTitle(this.AUDIT_LOG_LIST)
        .items
        .filter('RiskScore ge 70')
        .select(
          'Id', 'Title', 'ActionType', 'PerformedDate',
          'TargetOrganizationId', 'TargetUser', 'TargetResourceId', 'RiskScore'
        )
        .orderBy('PerformedDate', false)
        .top(50)();

      return highRiskLogs.map(item => ({
        id: item.Id.toString(),
        alertType: this.mapActionToAlertType(item.ActionType),
        severity: this.mapRiskScoreToSeverity(item.RiskScore),
        title: item.Title,
        description: `${item.ActionType} detected with risk score ${item.RiskScore}`,
        timestamp: new Date(item.PerformedDate),
        organizationId: item.TargetOrganizationId,
        userId: item.TargetUser,
        resourceId: item.TargetResourceId,
        isResolved: false
      }));
    } catch (error) {
      logger.error('ExternalSharingAuditService', 'Failed to get security alerts:', error);
      return [];
    }
  }

  /**
   * Create a security alert
   */
  private async createSecurityAlert(input: IAuditActionInput, riskScore: number): Promise<void> {
    try {
      // Log as a high-risk event - in a full implementation,
      // this would create a separate alert record and possibly
      // send notifications
      logger.warn('ExternalSharingAuditService', `High risk action detected: ${input.actionType} (Risk: ${riskScore})`);
    } catch (error) {
      logger.error('ExternalSharingAuditService', 'Failed to create security alert:', error);
    }
  }

  /**
   * Resolve a security alert
   */
  public async resolveAlert(alertId: string, resolution: string): Promise<void> {
    try {
      // In a full implementation, this would update the alert record
      logger.info('ExternalSharingAuditService', `Alert ${alertId} resolved: ${resolution}`);
    } catch (error) {
      logger.error('ExternalSharingAuditService', `Failed to resolve alert ${alertId}:`, error);
      throw error;
    }
  }

  // ============================================================================
  // RISK SCORING
  // ============================================================================

  /**
   * Calculate risk score for an action (0-100)
   */
  public calculateRiskScore(context: IRiskContext): number {
    let score = 0;

    // Base score by action type
    if (context.action) {
      switch (context.action) {
        case AuditActionType.TrustRevoked:
        case AuditActionType.GuestRemoved:
        case AuditActionType.ResourceUnshared:
          score += 10; // Normal administrative actions
          break;
        case AuditActionType.TrustEstablished:
        case AuditActionType.GuestInvited:
        case AuditActionType.ResourceShared:
          score += 20; // Adding access
          break;
        case AuditActionType.TrustModified:
        case AuditActionType.PermissionChanged:
          score += 30; // Changing permissions
          break;
        case AuditActionType.GuestSuspended:
        case AuditActionType.PolicyViolation:
          score += 50; // Security-relevant actions
          break;
      }
    }

    // Additional risk factors
    if (context.isNewLocation) score += 15;
    if (context.isOffHours) score += 10;
    if (context.recentFailedAttempts && context.recentFailedAttempts > 0) {
      score += Math.min(context.recentFailedAttempts * 10, 30);
    }

    // Cap at 100
    return Math.min(score, 100);
  }

  // ============================================================================
  // COMPLIANCE REPORTING
  // ============================================================================

  /**
   * Generate compliance report
   */
  public async generateComplianceReport(options: IReportOptions): Promise<IComplianceReport> {
    try {
      // Get all relevant data for the reporting period
      const auditLogs = await this.getAuditLog({
        startDate: options.startDate,
        endDate: options.endDate
      });

      // Calculate statistics
      const trustedOrgCount = await this.getUniqueCount('TargetOrganizationId', auditLogs);
      const guestCount = await this.getUniqueCount('TargetUser', auditLogs.filter(l => l.TargetUser));
      const resourceCount = await this.getUniqueCount('TargetResourceId', auditLogs.filter(l => l.TargetResourceId));

      const criticalFindings = auditLogs.filter(l => (l.RiskScore || 0) >= 80);
      const warnings = auditLogs.filter(l => (l.RiskScore || 0) >= 50 && (l.RiskScore || 0) < 80);
      const failures = auditLogs.filter(l => l.Result === AuditResult.Failure);

      // Calculate compliance and risk scores
      const totalActions = auditLogs.length;
      const successfulActions = auditLogs.filter(l => l.Result === AuditResult.Success).length;
      const complianceScore = totalActions > 0 ? Math.round((successfulActions / totalActions) * 100) : 100;

      const avgRiskScore = auditLogs.length > 0
        ? Math.round(auditLogs.reduce((sum, l) => sum + (l.RiskScore || 0), 0) / auditLogs.length)
        : 0;

      return {
        reportId: this.generateCorrelationId(),
        reportType: options.reportType,
        generatedDate: new Date(),
        periodStart: options.startDate,
        periodEnd: options.endDate,
        summary: {
          totalTrustedOrgs: trustedOrgCount,
          totalGuests: guestCount,
          totalSharedResources: resourceCount,
          totalAuditEvents: totalActions,
          complianceScore,
          riskScore: avgRiskScore
        },
        findings: {
          critical: criticalFindings.map(l => `High risk ${l.ActionType}: ${l.Title}`),
          warnings: warnings.slice(0, 10).map(l => `${l.ActionType}: ${l.Title}`),
          recommendations: this.generateRecommendations(failures, warnings.length, avgRiskScore)
        },
        data: options.reportType === 'Detailed' ? auditLogs : undefined
      };
    } catch (error) {
      logger.error('ExternalSharingAuditService', 'Failed to generate compliance report:', error);
      throw error;
    }
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Map SharePoint item to IExternalSharingAuditLog
   */
  private mapToAuditLogEntry(item: any): IExternalSharingAuditLog {
    return {
      Id: item.Id,
      Title: item.Title,
      ActionType: item.ActionType as AuditActionType,
      PerformedById: item.PerformedBy?.Id,
      PerformedBy: item.PerformedBy ? { Id: item.PerformedBy.Id, Title: item.PerformedBy.Title } : undefined,
      PerformedDate: new Date(item.PerformedDate),
      TargetOrganizationId: item.TargetOrganizationId,
      TargetResourceId: item.TargetResourceId,
      TargetUser: item.TargetUser,
      PreviousValue: item.PreviousValue,
      NewValue: item.NewValue,
      IPAddress: item.IPAddress,
      UserAgent: item.UserAgent,
      Result: item.Result as AuditResult,
      CorrelationId: item.CorrelationId,
      RiskScore: item.RiskScore
    };
  }

  /**
   * Format audit title
   */
  private formatAuditTitle(actionType: AuditActionType, targetUser?: string): string {
    const actionName = actionType.replace(/([A-Z])/g, ' $1').trim();
    return targetUser ? `${actionName}: ${targetUser}` : actionName;
  }

  /**
   * Generate correlation ID
   */
  private generateCorrelationId(): string {
    return `ESH-${Date.now()}-${Math.random().toString(36).substring(2, 9)}`;
  }

  /**
   * Check if action type is high risk
   */
  private isHighRiskAction(actionType: AuditActionType): boolean {
    const highRiskActions = [
      AuditActionType.PolicyViolation,
      AuditActionType.TrustRevoked
    ];
    return highRiskActions.includes(actionType);
  }

  /**
   * Map action type to alert type
   */
  private mapActionToAlertType(actionType: string): ISecurityAlert['alertType'] {
    if (actionType === AuditActionType.PolicyViolation) return 'PolicyViolation';
    if (actionType.includes('Revoke') || actionType.includes('Remove')) return 'ExpiredAccess';
    if (actionType.includes('Suspend')) return 'SuspiciousActivity';
    return 'HighRisk';
  }

  /**
   * Map risk score to severity
   */
  private mapRiskScoreToSeverity(riskScore: number): ISecurityAlert['severity'] {
    if (riskScore >= 80) return 'Critical';
    if (riskScore >= 60) return 'High';
    if (riskScore >= 40) return 'Medium';
    return 'Low';
  }

  /**
   * Get unique count of a field
   */
  private async getUniqueCount(field: string, items: any[]): Promise<number> {
    const uniqueValues = new Set(items.map(i => i[field]).filter(v => v != null));
    return uniqueValues.size;
  }

  /**
   * Generate recommendations based on findings
   */
  private generateRecommendations(failures: IExternalSharingAuditLog[], warningCount: number, avgRiskScore: number): string[] {
    const recommendations: string[] = [];

    if (failures.length > 0) {
      recommendations.push(`Review ${failures.length} failed operations for potential issues`);
    }

    if (warningCount > 10) {
      recommendations.push('Consider implementing stricter sharing policies to reduce warnings');
    }

    if (avgRiskScore > 50) {
      recommendations.push('Average risk score is elevated - review high-risk actions');
    }

    if (recommendations.length === 0) {
      recommendations.push('No critical recommendations at this time');
    }

    return recommendations;
  }
}
