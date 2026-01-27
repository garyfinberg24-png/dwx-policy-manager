// @ts-nocheck
// M365 License Service
// Comprehensive Microsoft 365 license management and optimization

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  IM365License,
  IM365LicenseAssignment,
  IM365LicenseUsageReport,
  IM365LicenseStatistics,
  IM365LicenseOptimization,
  IM365LicenseFilterCriteria,
  IM365LicenseRenewal,
  M365LicenseType,
  M365SubscriptionType
} from '../models/IAsset';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class M365LicenseService {
  private sp: SPFI;
  private readonly M365_LICENSES_LIST = 'M365 Licenses';
  private readonly M365_LICENSE_ASSIGNMENTS_LIST = 'M365 License Assignments';
  private readonly M365_LICENSE_RENEWALS_LIST = 'M365 License Renewals';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==================== License CRUD Operations ====================

  public async getLicenses(filter?: IM365LicenseFilterCriteria): Promise<IM365License[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.M365_LICENSES_LIST).items
        .select(
          'Id', 'Title', 'LicenseType', 'SubscriptionType',
          'TotalLicenses', 'AssignedLicenses', 'AvailableLicenses',
          'SubscriptionId', 'SkuId', 'SkuPartNumber',
          'PurchaseDate', 'StartDate', 'ExpirationDate', 'RenewalDate', 'AutoRenew',
          'CostPerLicense', 'BillingPeriod', 'TotalCost', 'NextBillingDate',
          'Vendor', 'ResellerContact', 'ContractNumber', 'PurchaseOrderNumber',
          'TenantId', 'AdminContactId', 'AdminContact/Title', 'AdminContact/EMail',
          'Department', 'CostCenter',
          'IsActive', 'IsExpiringSoon', 'HasUnusedLicenses',
          'ServicesIncluded', 'AddOns', 'ComplianceNotes', 'AuditDate',
          'Comments', 'Created', 'CreatedById', 'Modified', 'ModifiedById'
        )
        .expand('AdminContact');

      // Apply filters
      if (filter) {
        const filters: string[] = [];

        if (filter.licenseType && filter.licenseType.length > 0) {
          const typeFilters = filter.licenseType.map(t =>
            ValidationUtils.buildFilter('LicenseType', 'eq', t)
          );
          filters.push(`(${typeFilters.join(' or ')})`);
        }

        if (filter.subscriptionType && filter.subscriptionType.length > 0) {
          const subTypeFilters = filter.subscriptionType.map(st =>
            ValidationUtils.buildFilter('SubscriptionType', 'eq', st)
          );
          filters.push(`(${subTypeFilters.join(' or ')})`);
        }

        if (filter.department) {
          const validDept = ValidationUtils.sanitizeForOData(filter.department);
          filters.push(ValidationUtils.buildFilter('Department', 'eq', validDept));
        }

        if (filter.isActive !== undefined) {
          filters.push(`IsActive eq ${filter.isActive ? 1 : 0}`);
        }

        if (filter.isExpiringSoon) {
          filters.push('IsExpiringSoon eq 1');
        }

        if (filter.hasUnusedLicenses) {
          filters.push('HasUnusedLicenses eq 1');
        }

        if (filter.searchTerm) {
          const validTerm = ValidationUtils.sanitizeForOData(filter.searchTerm);
          filters.push(`(substringof('${validTerm}', Title) or substringof('${validTerm}', LicenseType))`);
        }

        if (filters.length > 0) {
          query = query.filter(filters.join(' and '));
        }
      }

      const items = await query.top(1000)();
      return items.map(this.mapLicenseFromSP);
    } catch (error) {
      logger.error('M365LicenseService', 'Error getting M365 licenses:', error);
      throw error;
    }
  }

  public async getLicenseById(id: number): Promise<IM365License> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const item = await this.sp.web.lists.getByTitle(this.M365_LICENSES_LIST).items
        .getById(validId)
        .select(
          'Id', 'Title', 'LicenseType', 'SubscriptionType',
          'TotalLicenses', 'AssignedLicenses', 'AvailableLicenses',
          'SubscriptionId', 'SkuId', 'SkuPartNumber',
          'PurchaseDate', 'StartDate', 'ExpirationDate', 'RenewalDate', 'AutoRenew',
          'CostPerLicense', 'BillingPeriod', 'TotalCost', 'NextBillingDate',
          'Vendor', 'ResellerContact', 'ContractNumber', 'PurchaseOrderNumber',
          'TenantId', 'AdminContactId', 'AdminContact/Title', 'AdminContact/EMail',
          'Department', 'CostCenter',
          'IsActive', 'IsExpiringSoon', 'HasUnusedLicenses',
          'ServicesIncluded', 'AddOns', 'ComplianceNotes', 'AuditDate',
          'Comments', 'Created', 'CreatedById', 'Modified', 'ModifiedById'
        )
        .expand('AdminContact')();

      return this.mapLicenseFromSP(item);
    } catch (error) {
      logger.error('M365LicenseService', 'Error getting M365 license by ID:', error);
      throw error;
    }
  }

  public async createLicense(license: Partial<IM365License>): Promise<number> {
    try {
      if (!license.Title || !license.LicenseType || !license.TotalLicenses) {
        throw new Error('Title, LicenseType, and TotalLicenses are required');
      }

      ValidationUtils.validateEnum(license.LicenseType, M365LicenseType, 'LicenseType');
      ValidationUtils.validateEnum(license.SubscriptionType, M365SubscriptionType, 'SubscriptionType');

      const itemData: any = {
        Title: ValidationUtils.sanitizeHtml(license.Title),
        LicenseType: license.LicenseType,
        SubscriptionType: license.SubscriptionType,
        TotalLicenses: ValidationUtils.validateInteger(license.TotalLicenses, 'TotalLicenses', 0),
        AssignedLicenses: license.AssignedLicenses || 0,
        AvailableLicenses: (license.TotalLicenses || 0) - (license.AssignedLicenses || 0),
        IsActive: license.IsActive !== undefined ? license.IsActive : true
      };

      if (license.SubscriptionId) itemData.SubscriptionId = ValidationUtils.sanitizeHtml(license.SubscriptionId);
      if (license.SkuId) itemData.SkuId = ValidationUtils.sanitizeHtml(license.SkuId);
      if (license.SkuPartNumber) itemData.SkuPartNumber = ValidationUtils.sanitizeHtml(license.SkuPartNumber);
      if (license.PurchaseDate) itemData.PurchaseDate = ValidationUtils.validateDate(license.PurchaseDate, 'PurchaseDate');
      if (license.StartDate) itemData.StartDate = ValidationUtils.validateDate(license.StartDate, 'StartDate');
      if (license.ExpirationDate) itemData.ExpirationDate = ValidationUtils.validateDate(license.ExpirationDate, 'ExpirationDate');
      if (license.RenewalDate) itemData.RenewalDate = ValidationUtils.validateDate(license.RenewalDate, 'RenewalDate');
      if (license.AutoRenew !== undefined) itemData.AutoRenew = license.AutoRenew;
      if (license.CostPerLicense !== undefined) itemData.CostPerLicense = license.CostPerLicense;
      if (license.BillingPeriod) itemData.BillingPeriod = license.BillingPeriod;
      if (license.Vendor) itemData.Vendor = ValidationUtils.sanitizeHtml(license.Vendor);
      if (license.Department) itemData.Department = ValidationUtils.sanitizeHtml(license.Department);
      if (license.CostCenter) itemData.CostCenter = ValidationUtils.sanitizeHtml(license.CostCenter);
      if (license.AdminContactId) itemData.AdminContactId = ValidationUtils.validateInteger(license.AdminContactId, 'AdminContactId', 1);
      if (license.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(license.Notes);

      // Calculate total cost
      if (license.CostPerLicense && license.TotalLicenses) {
        itemData.TotalCost = license.CostPerLicense * license.TotalLicenses;
      }

      const result = await this.sp.web.lists.getByTitle(this.M365_LICENSES_LIST).items.add(itemData);
      return result.data.Id;
    } catch (error) {
      logger.error('M365LicenseService', 'Error creating M365 license:', error);
      throw error;
    }
  }

  public async updateLicense(id: number, updates: Partial<IM365License>): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const itemData: any = {};

      if (updates.Title) itemData.Title = ValidationUtils.sanitizeHtml(updates.Title);
      if (updates.TotalLicenses !== undefined) {
        itemData.TotalLicenses = ValidationUtils.validateInteger(updates.TotalLicenses, 'TotalLicenses', 0);
      }
      if (updates.AssignedLicenses !== undefined) {
        itemData.AssignedLicenses = ValidationUtils.validateInteger(updates.AssignedLicenses, 'AssignedLicenses', 0);
      }

      // Recalculate available licenses if either total or assigned changed
      if (updates.TotalLicenses !== undefined || updates.AssignedLicenses !== undefined) {
        const current = await this.getLicenseById(validId);
        const total = updates.TotalLicenses !== undefined ? updates.TotalLicenses : current.TotalLicenses;
        const assigned = updates.AssignedLicenses !== undefined ? updates.AssignedLicenses : current.AssignedLicenses;
        itemData.AvailableLicenses = total - assigned;
      }

      if (updates.ExpirationDate) itemData.ExpirationDate = ValidationUtils.validateDate(updates.ExpirationDate, 'ExpirationDate');
      if (updates.RenewalDate) itemData.RenewalDate = ValidationUtils.validateDate(updates.RenewalDate, 'RenewalDate');
      if (updates.AutoRenew !== undefined) itemData.AutoRenew = updates.AutoRenew;
      if (updates.CostPerLicense !== undefined) itemData.CostPerLicense = updates.CostPerLicense;
      if (updates.IsActive !== undefined) itemData.IsActive = updates.IsActive;
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes ? ValidationUtils.sanitizeHtml(updates.Notes) : null;

      await this.sp.web.lists.getByTitle(this.M365_LICENSES_LIST).items.getById(validId).update(itemData);
    } catch (error) {
      logger.error('M365LicenseService', 'Error updating M365 license:', error);
      throw error;
    }
  }

  // ==================== License Assignment Operations ====================

  public async assignLicense(
    licenseId: number,
    userId: number,
    assignedById: number,
    processId?: number,
    reason?: string
  ): Promise<number> {
    try {
      const validLicenseId = ValidationUtils.validateInteger(licenseId, 'licenseId', 1);
      const validUserId = ValidationUtils.validateInteger(userId, 'userId', 1);
      const validAssignedById = ValidationUtils.validateInteger(assignedById, 'assignedById', 1);

      // Check if license has available capacity
      const license = await this.getLicenseById(validLicenseId);
      if (license.AvailableLicenses <= 0) {
        throw new Error(`No available licenses for ${license.Title}`);
      }

      // Check if user already has this license
      const existingAssignments = await this.getLicenseAssignments(validLicenseId, validUserId);
      if (existingAssignments.length > 0) {
        throw new Error('User already has this license assigned');
      }

      // Create assignment record
      const assignmentData: any = {
        LicenseId: validLicenseId,
        AssignedToId: validUserId,
        AssignedById: validAssignedById,
        AssignedDate: new Date().toISOString(),
        IsActive: true,
        Status: 'Active'
      };

      if (processId) {
        assignmentData.ProcessId = ValidationUtils.validateInteger(processId, 'processId', 1);
      }

      if (reason) {
        assignmentData.AssignmentReason = ValidationUtils.sanitizeHtml(reason);
      }

      const result = await this.sp.web.lists.getByTitle(this.M365_LICENSE_ASSIGNMENTS_LIST).items.add(assignmentData);

      // Update license assigned count
      await this.updateLicense(validLicenseId, {
        AssignedLicenses: license.AssignedLicenses + 1
      });

      return result.data.Id;
    } catch (error) {
      logger.error('M365LicenseService', 'Error assigning M365 license:', error);
      throw error;
    }
  }

  public async unassignLicense(assignmentId: number, reason?: string): Promise<void> {
    try {
      const validAssignmentId = ValidationUtils.validateInteger(assignmentId, 'assignmentId', 1);

      // Get assignment
      const assignment = await this.sp.web.lists.getByTitle(this.M365_LICENSE_ASSIGNMENTS_LIST).items
        .getById(validAssignmentId)
        .select('Id', 'LicenseId', 'IsActive')();

      if (!assignment.IsActive) {
        throw new Error('License assignment is already inactive');
      }

      // Update assignment
      const updateData: any = {
        UnassignedDate: new Date().toISOString(),
        IsActive: false,
        Status: 'Removed'
      };

      if (reason) {
        updateData.Notes = ValidationUtils.sanitizeHtml(reason);
      }

      await this.sp.web.lists.getByTitle(this.M365_LICENSE_ASSIGNMENTS_LIST).items
        .getById(validAssignmentId)
        .update(updateData);

      // Update license assigned count
      const license = await this.getLicenseById(assignment.LicenseId);
      await this.updateLicense(assignment.LicenseId, {
        AssignedLicenses: Math.max(0, license.AssignedLicenses - 1)
      });
    } catch (error) {
      logger.error('M365LicenseService', 'Error unassigning M365 license:', error);
      throw error;
    }
  }

  public async getLicenseAssignments(licenseId?: number, userId?: number): Promise<IM365LicenseAssignment[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.M365_LICENSE_ASSIGNMENTS_LIST).items
        .select(
          'Id', 'LicenseId', 'AssignedToId', 'AssignedTo/Title', 'AssignedTo/EMail',
          'AssignedById', 'AssignedBy/Title', 'AssignedDate', 'UnassignedDate',
          'ProcessId', 'AssignmentReason', 'Department', 'IsActive', 'Status',
          'LastUsedDate', 'DaysUnused', 'UsagePercentage', 'EnabledServices',
          'Comments', 'Created', 'Modified'
        )
        .expand('AssignedTo', 'AssignedBy')
        .orderBy('AssignedDate', false);

      const filters: string[] = [];

      if (licenseId !== undefined) {
        const validLicenseId = ValidationUtils.validateInteger(licenseId, 'licenseId', 1);
        filters.push(ValidationUtils.buildFilter('LicenseId', 'eq', validLicenseId));
      }

      if (userId !== undefined) {
        const validUserId = ValidationUtils.validateInteger(userId, 'userId', 1);
        filters.push(ValidationUtils.buildFilter('AssignedToId', 'eq', validUserId));
      }

      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      const items = await query.top(5000)();
      return items.map(this.mapAssignmentFromSP);
    } catch (error) {
      logger.error('M365LicenseService', 'Error getting M365 license assignments:', error);
      throw error;
    }
  }

  // ==================== Statistics and Reporting ====================

  public async getLicenseStatistics(): Promise<IM365LicenseStatistics> {
    try {
      const allLicenses = await this.getLicenses({ isActive: true });

      const stats: IM365LicenseStatistics = {
        totalLicenses: allLicenses.reduce((sum, l) => sum + l.TotalLicenses, 0),
        totalAssigned: allLicenses.reduce((sum, l) => sum + l.AssignedLicenses, 0),
        totalAvailable: allLicenses.reduce((sum, l) => sum + l.AvailableLicenses, 0),
        totalCost: allLicenses.reduce((sum, l) => sum + (l.TotalCost || 0), 0),
        monthlyRecurringCost: 0,
        annualRecurringCost: 0,
        utilizationRate: 0,
        byLicenseType: {},
        byDepartment: {},
        expiringSoon: 0,
        upForRenewal: 0,
        inactiveLicenses: 0,
        recentActivity: {
          newAssignments: 0,
          removedAssignments: 0,
          renewals: 0
        }
      };

      // Calculate utilization rate
      if (stats.totalLicenses > 0) {
        stats.utilizationRate = (stats.totalAssigned / stats.totalLicenses) * 100;
      }

      // Calculate recurring costs
      for (const license of allLicenses) {
        if (license.CostPerLicense && license.TotalLicenses) {
          if (license.BillingPeriod === 'Monthly') {
            stats.monthlyRecurringCost += license.CostPerLicense * license.TotalLicenses;
            stats.annualRecurringCost += license.CostPerLicense * license.TotalLicenses * 12;
          } else if (license.BillingPeriod === 'Annual') {
            stats.annualRecurringCost += license.CostPerLicense * license.TotalLicenses;
            stats.monthlyRecurringCost += (license.CostPerLicense * license.TotalLicenses) / 12;
          }
        }
      }

      // Group by license type
      for (const license of allLicenses) {
        if (!stats.byLicenseType[license.LicenseType]) {
          stats.byLicenseType[license.LicenseType] = {
            total: 0,
            assigned: 0,
            available: 0,
            cost: 0
          };
        }

        stats.byLicenseType[license.LicenseType]!.total += license.TotalLicenses;
        stats.byLicenseType[license.LicenseType]!.assigned += license.AssignedLicenses;
        stats.byLicenseType[license.LicenseType]!.available += license.AvailableLicenses;
        stats.byLicenseType[license.LicenseType]!.cost += license.TotalCost || 0;
      }

      // Group by department
      for (const license of allLicenses) {
        if (license.Department) {
          if (!stats.byDepartment[license.Department]) {
            stats.byDepartment[license.Department] = {
              total: 0,
              cost: 0
            };
          }

          stats.byDepartment[license.Department].total += license.TotalLicenses;
          stats.byDepartment[license.Department].cost += license.TotalCost || 0;
        }
      }

      // Expiring soon (90 days)
      const now = new Date();
      const ninetyDaysFromNow = new Date(now.getTime() + 90 * 24 * 60 * 60 * 1000);
      stats.expiringSoon = allLicenses.filter(l =>
        l.ExpirationDate &&
        new Date(l.ExpirationDate) >= now &&
        new Date(l.ExpirationDate) <= ninetyDaysFromNow
      ).length;

      // Up for renewal (30 days)
      const thirtyDaysFromNow = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
      stats.upForRenewal = allLicenses.filter(l =>
        l.RenewalDate &&
        new Date(l.RenewalDate) >= now &&
        new Date(l.RenewalDate) <= thirtyDaysFromNow
      ).length;

      return stats;
    } catch (error) {
      logger.error('M365LicenseService', 'Error getting M365 license statistics:', error);
      throw error;
    }
  }

  public async getUsageReports(): Promise<IM365LicenseUsageReport[]> {
    try {
      const licenses = await this.getLicenses({ isActive: true });
      const reports: IM365LicenseUsageReport[] = [];

      for (const license of licenses) {
        const utilizationRate = license.TotalLicenses > 0
          ? (license.AssignedLicenses / license.TotalLicenses) * 100
          : 0;

        const costPerMonth = license.BillingPeriod === 'Monthly'
          ? (license.CostPerLicense || 0) * license.TotalLicenses
          : ((license.CostPerLicense || 0) * license.TotalLicenses) / 12;

        const report: IM365LicenseUsageReport = {
          LicenseType: license.LicenseType,
          TotalLicenses: license.TotalLicenses,
          AssignedLicenses: license.AssignedLicenses,
          UnassignedLicenses: license.AvailableLicenses,
          UtilizationRate: utilizationRate,
          InactiveLicenses: 0, // Would need usage data from Microsoft Graph API
          CostPerMonth: costPerMonth,
          WastedCost: 0, // Would calculate based on inactive licenses
          RecommendedAction: this.getRecommendedAction(utilizationRate, license.AvailableLicenses)
        };

        reports.push(report);
      }

      return reports;
    } catch (error) {
      logger.error('M365LicenseService', 'Error getting M365 usage reports:', error);
      throw error;
    }
  }

  private getRecommendedAction(utilizationRate: number, availableLicenses: number): string {
    if (utilizationRate < 70) {
      return `Consider reducing licenses. ${availableLicenses} licenses unused.`;
    } else if (utilizationRate > 95) {
      return 'Consider purchasing additional licenses to accommodate growth.';
    } else {
      return 'License allocation is optimal.';
    }
  }

  public async getOptimizationRecommendations(): Promise<IM365LicenseOptimization[]> {
    try {
      const licenses = await this.getLicenses({ isActive: true });
      const recommendations: IM365LicenseOptimization[] = [];

      for (const license of licenses) {
        const utilizationRate = license.TotalLicenses > 0
          ? (license.AssignedLicenses / license.TotalLicenses) * 100
          : 0;

        // Only recommend optimization if underutilized
        if (utilizationRate < 80 && license.AvailableLicenses > 5) {
          const recommendedCount = Math.ceil(license.AssignedLicenses * 1.1); // 10% buffer
          const licensesToRemove = license.TotalLicenses - recommendedCount;
          const potentialSavings = licensesToRemove * (license.CostPerLicense || 0) *
            (license.BillingPeriod === 'Annual' ? 1 : 12);

          const recommendation: IM365LicenseOptimization = {
            licenseType: license.LicenseType,
            currentCount: license.TotalLicenses,
            recommendedCount: recommendedCount,
            potentialSavings: potentialSavings,
            reason: `${license.AvailableLicenses} unused licenses (${utilizationRate.toFixed(1)}% utilization)`,
            inactiveUsers: [], // Would need usage data
            priority: potentialSavings > 10000 ? 'High' : potentialSavings > 5000 ? 'Medium' : 'Low'
          };

          recommendations.push(recommendation);
        }
      }

      return recommendations.sort((a, b) => b.potentialSavings - a.potentialSavings);
    } catch (error) {
      logger.error('M365LicenseService', 'Error getting optimization recommendations:', error);
      throw error;
    }
  }

  // ==================== Mapping Functions ====================

  private mapLicenseFromSP(item: any): IM365License {
    return {
      Id: item.Id,
      Title: item.Title,
      LicenseType: item.LicenseType as M365LicenseType,
      SubscriptionType: item.SubscriptionType as M365SubscriptionType,
      TotalLicenses: item.TotalLicenses,
      AssignedLicenses: item.AssignedLicenses,
      AvailableLicenses: item.AvailableLicenses,
      SubscriptionId: item.SubscriptionId,
      SkuId: item.SkuId,
      SkuPartNumber: item.SkuPartNumber,
      PurchaseDate: item.PurchaseDate ? new Date(item.PurchaseDate) : undefined,
      StartDate: item.StartDate ? new Date(item.StartDate) : undefined,
      ExpirationDate: item.ExpirationDate ? new Date(item.ExpirationDate) : undefined,
      RenewalDate: item.RenewalDate ? new Date(item.RenewalDate) : undefined,
      AutoRenew: item.AutoRenew,
      CostPerLicense: item.CostPerLicense,
      BillingPeriod: item.BillingPeriod,
      TotalCost: item.TotalCost,
      NextBillingDate: item.NextBillingDate ? new Date(item.NextBillingDate) : undefined,
      Vendor: item.Vendor,
      ResellerContact: item.ResellerContact,
      ContractNumber: item.ContractNumber,
      PurchaseOrderNumber: item.PurchaseOrderNumber,
      TenantId: item.TenantId,
      AdminContactId: item.AdminContactId,
      AdminContact: item.AdminContact,
      Department: item.Department,
      CostCenter: item.CostCenter,
      IsActive: item.IsActive,
      IsExpiringSoon: item.IsExpiringSoon,
      HasUnusedLicenses: item.HasUnusedLicenses,
      ServicesIncluded: item.ServicesIncluded,
      AddOns: item.AddOns,
      ComplianceNotes: item.ComplianceNotes,
      AuditDate: item.AuditDate ? new Date(item.AuditDate) : undefined,
      Notes: item.Notes,
      Created: item.Created ? new Date(item.Created) : undefined,
      CreatedById: item.CreatedById,
      Modified: item.Modified ? new Date(item.Modified) : undefined,
      ModifiedById: item.ModifiedById
    };
  }

  private mapAssignmentFromSP(item: any): IM365LicenseAssignment {
    return {
      Id: item.Id,
      LicenseId: item.LicenseId,
      AssignedToId: item.AssignedToId,
      AssignedTo: item.AssignedTo,
      AssignedById: item.AssignedById,
      AssignedBy: item.AssignedBy,
      AssignedDate: new Date(item.AssignedDate),
      UnassignedDate: item.UnassignedDate ? new Date(item.UnassignedDate) : undefined,
      ProcessId: item.ProcessId,
      AssignmentReason: item.AssignmentReason,
      Department: item.Department,
      IsActive: item.IsActive,
      Status: item.Status,
      LastUsedDate: item.LastUsedDate ? new Date(item.LastUsedDate) : undefined,
      DaysUnused: item.DaysUnused,
      UsagePercentage: item.UsagePercentage,
      EnabledServices: item.EnabledServices,
      Notes: item.Notes,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }
}
