// @ts-nocheck
/**
 * SLAComplianceService
 *
 * Reads admin-configured SLA targets from PM_SLAConfigs and PM_Configuration,
 * then measures ACTUAL performance against those targets from live SP data.
 *
 * This service powers:
 *   - PolicyAnalytics SLA tab (real data instead of mocks)
 *   - Admin dashboard SLA summary
 *   - SLA breach notifications
 *   - SLA compliance reporting
 *
 * SLA Types:
 *   - Acknowledgement: Are users acknowledging policies within the target days?
 *   - Approval: Are approvers actioning requests within the target days?
 *   - Review: Are policy owners completing reviews within the target days?
 *   - Authoring: Are authors completing drafts within the target days?
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { logger } from './LoggingService';

// ─── Interfaces ────────────────────────────────────────────────

export interface ISLATarget {
  processType: string;
  targetDays: number;
  warningThresholdDays: number;
  isActive: boolean;
}

export interface ISLAMetricResult {
  processType: string;
  targetDays: number;
  warningDays: number;
  totalItems: number;
  completedWithinSLA: number;
  completedOutsideSLA: number;
  currentlyAtRisk: number;     // within warning threshold
  currentlyBreached: number;   // past target
  slaCompliancePercent: number; // 0-100
  avgCompletionDays: number;
  status: 'Met' | 'At Risk' | 'Breached';
}

export interface ISLABreachItem {
  id: number;
  title: string;
  entityType: string;
  assignedTo: string;
  assignedDate: Date;
  targetDate: Date;
  daysOverdue: number;
  policyId?: number;
  policyName?: string;
}

export interface ISLADashboard {
  metrics: ISLAMetricResult[];
  breaches: ISLABreachItem[];
  overallCompliancePercent: number;
  totalProcessed: number;
  totalBreaches: number;
  lastCalculated: Date;
}

// ─── Service ───────────────────────────────────────────────────

export class SLAComplianceService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Get all active SLA targets from admin config
   */
  public async getSLATargets(): Promise<ISLATarget[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle('PM_Configuration')
        .items.filter("substringof('Admin.SLA', ConfigKey)")
        .select('ConfigKey', 'ConfigValue')
        .top(20)();

      const targets: ISLATarget[] = [];
      for (const item of items) {
        try {
          const config = JSON.parse(item.ConfigValue);
          if (config.IsActive !== false) {
            targets.push({
              processType: config.ProcessType || item.ConfigKey.replace('Admin.SLA.', ''),
              targetDays: Number(config.TargetDays) || 7,
              warningThresholdDays: Number(config.WarningThresholdDays) || 2,
              isActive: config.IsActive !== false
            });
          }
        } catch { /* skip malformed entries */ }
      }

      // Also try the dedicated SLA list
      try {
        const slaItems = await this.sp.web.lists
          .getByTitle('PM_SLAConfigs')
          .items.filter('IsActive eq true')
          .select('Id', 'Title', 'ProcessType', 'TargetDays', 'WarningThresholdDays', 'IsActive')
          .top(20)();

        for (const sla of slaItems) {
          // Don't duplicate if already loaded from PM_Configuration
          if (!targets.find(t => t.processType === sla.ProcessType)) {
            targets.push({
              processType: sla.ProcessType,
              targetDays: sla.TargetDays || 7,
              warningThresholdDays: sla.WarningThresholdDays || 2,
              isActive: true
            });
          }
        }
      } catch { /* PM_SLAConfigs may not exist */ }

      return targets;
    } catch (error) {
      logger.error('SLAComplianceService', 'Failed to load SLA targets:', error);
      return [];
    }
  }

  /**
   * Calculate full SLA dashboard with real data
   */
  public async calculateDashboard(): Promise<ISLADashboard> {
    const targets = await this.getSLATargets();
    const metrics: ISLAMetricResult[] = [];
    const allBreaches: ISLABreachItem[] = [];
    let totalProcessed = 0;
    let totalMet = 0;

    for (const target of targets) {
      try {
        let result: ISLAMetricResult;

        switch (target.processType) {
          case 'Acknowledgement':
            result = await this.measureAcknowledgementSLA(target);
            break;
          case 'Approval':
            result = await this.measureApprovalSLA(target);
            break;
          case 'Review':
            result = await this.measureReviewSLA(target);
            break;
          case 'Authoring':
            result = await this.measureAuthoringSLA(target);
            break;
          default:
            result = await this.measureAcknowledgementSLA(target);
            break;
        }

        metrics.push(result);
        totalProcessed += result.totalItems;
        totalMet += result.completedWithinSLA;

        // Collect breaches
        if (result.currentlyBreached > 0) {
          const breaches = await this.getBreachItems(target);
          allBreaches.push(...breaches);
        }
      } catch (err) {
        logger.warn('SLAComplianceService', `Failed to measure ${target.processType} SLA:`, err);
      }
    }

    const overallCompliancePercent = totalProcessed > 0
      ? Math.round((totalMet / totalProcessed) * 100)
      : 100;

    return {
      metrics,
      breaches: allBreaches,
      overallCompliancePercent,
      totalProcessed,
      totalBreaches: allBreaches.length,
      lastCalculated: new Date()
    };
  }

  // ─── Acknowledgement SLA ──────────────────────────────────────

  private async measureAcknowledgementSLA(target: ISLATarget): Promise<ISLAMetricResult> {
    const now = new Date();
    let items: any[] = [];

    try {
      items = await this.sp.web.lists
        .getByTitle('PM_PolicyAcknowledgements')
        .items.select('Id', 'AckStatus', 'AssignedDate', 'AcknowledgedDate', 'DueDate', 'PolicyId')
        .top(500)();
    } catch {
      return this.emptyMetric(target);
    }

    let completedWithinSLA = 0;
    let completedOutsideSLA = 0;
    let currentlyAtRisk = 0;
    let currentlyBreached = 0;
    let totalCompletionDays = 0;
    let completedCount = 0;

    for (const item of items) {
      const assigned = new Date(item.AssignedDate);
      const targetDate = new Date(assigned.getTime() + target.targetDays * 86400000);
      const warningDate = new Date(targetDate.getTime() - target.warningThresholdDays * 86400000);

      if (item.AckStatus === 'Acknowledged' && item.AcknowledgedDate) {
        const ackDate = new Date(item.AcknowledgedDate);
        const daysTaken = Math.ceil((ackDate.getTime() - assigned.getTime()) / 86400000);
        totalCompletionDays += daysTaken;
        completedCount++;

        if (daysTaken <= target.targetDays) {
          completedWithinSLA++;
        } else {
          completedOutsideSLA++;
        }
      } else {
        // Still pending
        if (now > targetDate) {
          currentlyBreached++;
        } else if (now > warningDate) {
          currentlyAtRisk++;
        }
      }
    }

    const totalItems = items.length;
    const totalCompleted = completedWithinSLA + completedOutsideSLA;
    const slaCompliancePercent = totalCompleted > 0
      ? Math.round((completedWithinSLA / totalCompleted) * 100)
      : (currentlyBreached > 0 ? 0 : 100);
    const avgCompletionDays = completedCount > 0
      ? Math.round((totalCompletionDays / completedCount) * 10) / 10
      : 0;

    return {
      processType: 'Acknowledgement',
      targetDays: target.targetDays,
      warningDays: target.warningThresholdDays,
      totalItems,
      completedWithinSLA,
      completedOutsideSLA,
      currentlyAtRisk,
      currentlyBreached,
      slaCompliancePercent,
      avgCompletionDays,
      status: currentlyBreached > 0 ? 'Breached' : currentlyAtRisk > 0 ? 'At Risk' : 'Met'
    };
  }

  // ─── Approval SLA ─────────────────────────────────────────────

  private async measureApprovalSLA(target: ISLATarget): Promise<ISLAMetricResult> {
    const now = new Date();
    let items: any[] = [];

    try {
      items = await this.sp.web.lists
        .getByTitle('PM_Approvals')
        .items.select('Id', 'Status', 'Created', 'Modified', 'PolicyId')
        .top(500)();
    } catch {
      return this.emptyMetric(target);
    }

    let completedWithinSLA = 0;
    let completedOutsideSLA = 0;
    let currentlyAtRisk = 0;
    let currentlyBreached = 0;
    let totalCompletionDays = 0;
    let completedCount = 0;

    for (const item of items) {
      const created = new Date(item.Created);
      const targetDate = new Date(created.getTime() + target.targetDays * 86400000);
      const warningDate = new Date(targetDate.getTime() - target.warningThresholdDays * 86400000);

      if (item.Status === 'Approved' || item.Status === 'Rejected') {
        const completedDate = new Date(item.Modified);
        const daysTaken = Math.ceil((completedDate.getTime() - created.getTime()) / 86400000);
        totalCompletionDays += daysTaken;
        completedCount++;

        if (daysTaken <= target.targetDays) {
          completedWithinSLA++;
        } else {
          completedOutsideSLA++;
        }
      } else if (item.Status === 'Pending') {
        if (now > targetDate) {
          currentlyBreached++;
        } else if (now > warningDate) {
          currentlyAtRisk++;
        }
      }
    }

    const totalItems = items.length;
    const totalCompleted = completedWithinSLA + completedOutsideSLA;
    const slaCompliancePercent = totalCompleted > 0
      ? Math.round((completedWithinSLA / totalCompleted) * 100)
      : (currentlyBreached > 0 ? 0 : 100);
    const avgCompletionDays = completedCount > 0
      ? Math.round((totalCompletionDays / completedCount) * 10) / 10
      : 0;

    return {
      processType: 'Approval',
      targetDays: target.targetDays,
      warningDays: target.warningThresholdDays,
      totalItems,
      completedWithinSLA,
      completedOutsideSLA,
      currentlyAtRisk,
      currentlyBreached,
      slaCompliancePercent,
      avgCompletionDays,
      status: currentlyBreached > 0 ? 'Breached' : currentlyAtRisk > 0 ? 'At Risk' : 'Met'
    };
  }

  // ─── Review SLA ───────────────────────────────────────────────

  private async measureReviewSLA(target: ISLATarget): Promise<ISLAMetricResult> {
    const now = new Date();
    let items: any[] = [];

    try {
      items = await this.sp.web.lists
        .getByTitle('PM_Policies')
        .items.filter("PolicyStatus eq 'Published'")
        .select('Id', 'Title', 'PolicyName', 'NextReviewDate', 'LastReviewDate', 'PolicyStatus')
        .top(500)();
    } catch {
      return this.emptyMetric(target);
    }

    let completedWithinSLA = 0;
    let completedOutsideSLA = 0;
    let currentlyAtRisk = 0;
    let currentlyBreached = 0;

    for (const policy of items) {
      if (!policy.NextReviewDate) continue;

      const reviewDue = new Date(policy.NextReviewDate);
      const warningDate = new Date(reviewDue.getTime() - target.warningThresholdDays * 86400000);

      if (policy.LastReviewDate) {
        const reviewed = new Date(policy.LastReviewDate);
        if (reviewed <= reviewDue) {
          completedWithinSLA++;
        } else {
          completedOutsideSLA++;
        }
      } else {
        // Not yet reviewed
        if (now > reviewDue) {
          currentlyBreached++;
        } else if (now > warningDate) {
          currentlyAtRisk++;
        } else {
          completedWithinSLA++; // Not due yet — counts as meeting SLA
        }
      }
    }

    const totalItems = items.filter(p => p.NextReviewDate).length;
    const totalCompleted = completedWithinSLA + completedOutsideSLA;
    const slaCompliancePercent = totalCompleted > 0
      ? Math.round((completedWithinSLA / totalCompleted) * 100)
      : 100;

    return {
      processType: 'Review',
      targetDays: target.targetDays,
      warningDays: target.warningThresholdDays,
      totalItems,
      completedWithinSLA,
      completedOutsideSLA,
      currentlyAtRisk,
      currentlyBreached,
      slaCompliancePercent,
      avgCompletionDays: 0, // Not applicable for reviews (they're date-based)
      status: currentlyBreached > 0 ? 'Breached' : currentlyAtRisk > 0 ? 'At Risk' : 'Met'
    };
  }

  // ─── Authoring SLA ────────────────────────────────────────────

  private async measureAuthoringSLA(target: ISLATarget): Promise<ISLAMetricResult> {
    const now = new Date();
    let items: any[] = [];

    try {
      items = await this.sp.web.lists
        .getByTitle('PM_Policies')
        .items.filter("PolicyStatus eq 'Draft' or PolicyStatus eq 'In Review' or PolicyStatus eq 'Published'")
        .select('Id', 'PolicyStatus', 'Created', 'PublishedDate')
        .top(500)();
    } catch {
      return this.emptyMetric(target);
    }

    let completedWithinSLA = 0;
    let completedOutsideSLA = 0;
    let currentlyAtRisk = 0;
    let currentlyBreached = 0;
    let totalCompletionDays = 0;
    let completedCount = 0;

    for (const policy of items) {
      const created = new Date(policy.Created);
      const targetDate = new Date(created.getTime() + target.targetDays * 86400000);
      const warningDate = new Date(targetDate.getTime() - target.warningThresholdDays * 86400000);

      if (policy.PolicyStatus === 'Published' && policy.PublishedDate) {
        const published = new Date(policy.PublishedDate);
        const daysTaken = Math.ceil((published.getTime() - created.getTime()) / 86400000);
        totalCompletionDays += daysTaken;
        completedCount++;

        if (daysTaken <= target.targetDays) {
          completedWithinSLA++;
        } else {
          completedOutsideSLA++;
        }
      } else if (policy.PolicyStatus === 'Draft' || policy.PolicyStatus === 'In Review') {
        if (now > targetDate) {
          currentlyBreached++;
        } else if (now > warningDate) {
          currentlyAtRisk++;
        }
      }
    }

    const totalItems = items.length;
    const totalCompleted = completedWithinSLA + completedOutsideSLA;
    const slaCompliancePercent = totalCompleted > 0
      ? Math.round((completedWithinSLA / totalCompleted) * 100)
      : (currentlyBreached > 0 ? 0 : 100);
    const avgCompletionDays = completedCount > 0
      ? Math.round((totalCompletionDays / completedCount) * 10) / 10
      : 0;

    return {
      processType: 'Authoring',
      targetDays: target.targetDays,
      warningDays: target.warningThresholdDays,
      totalItems,
      completedWithinSLA,
      completedOutsideSLA,
      currentlyAtRisk,
      currentlyBreached,
      slaCompliancePercent,
      avgCompletionDays,
      status: currentlyBreached > 0 ? 'Breached' : currentlyAtRisk > 0 ? 'At Risk' : 'Met'
    };
  }

  // ─── Breach Items ─────────────────────────────────────────────

  private async getBreachItems(target: ISLATarget): Promise<ISLABreachItem[]> {
    const now = new Date();
    const breaches: ISLABreachItem[] = [];

    try {
      if (target.processType === 'Acknowledgement') {
        const items = await this.sp.web.lists
          .getByTitle('PM_PolicyAcknowledgements')
          .items.filter("AckStatus ne 'Acknowledged' and AckStatus ne 'Exempted'")
          .select('Id', 'Title', 'AckUserId', 'AssignedDate', 'DueDate', 'PolicyId', 'UserEmail')
          .top(100)();

        for (const item of items) {
          const assigned = new Date(item.AssignedDate);
          const targetDate = new Date(assigned.getTime() + target.targetDays * 86400000);
          if (now > targetDate) {
            const daysOverdue = Math.ceil((now.getTime() - targetDate.getTime()) / 86400000);
            breaches.push({
              id: item.Id,
              title: item.Title || 'Acknowledgement',
              entityType: 'Acknowledgement',
              assignedTo: item.UserEmail || `User ${item.AckUserId}`,
              assignedDate: assigned,
              targetDate,
              daysOverdue,
              policyId: item.PolicyId
            });
          }
        }
      } else if (target.processType === 'Approval') {
        const items = await this.sp.web.lists
          .getByTitle('PM_Approvals')
          .items.filter("Status eq 'Pending'")
          .select('Id', 'Title', 'Created', 'PolicyId')
          .top(100)();

        for (const item of items) {
          const created = new Date(item.Created);
          const targetDate = new Date(created.getTime() + target.targetDays * 86400000);
          if (now > targetDate) {
            const daysOverdue = Math.ceil((now.getTime() - targetDate.getTime()) / 86400000);
            breaches.push({
              id: item.Id,
              title: item.Title || 'Approval',
              entityType: 'Approval',
              assignedTo: 'Approver',
              assignedDate: created,
              targetDate,
              daysOverdue,
              policyId: item.PolicyId
            });
          }
        }
      }
    } catch (error) {
      logger.warn('SLAComplianceService', `Failed to get breach items for ${target.processType}:`, error);
    }

    return breaches;
  }

  // ─── Helpers ──────────────────────────────────────────────────

  private emptyMetric(target: ISLATarget): ISLAMetricResult {
    return {
      processType: target.processType,
      targetDays: target.targetDays,
      warningDays: target.warningThresholdDays,
      totalItems: 0,
      completedWithinSLA: 0,
      completedOutsideSLA: 0,
      currentlyAtRisk: 0,
      currentlyBreached: 0,
      slaCompliancePercent: 100,
      avgCompletionDays: 0,
      status: 'Met'
    };
  }
}
