// @ts-nocheck
/* eslint-disable @typescript-eslint/no-explicit-any */
// Budget Service
// Budget management, allocation tracking, and spend monitoring
// Note: Some fields may not exist in the SharePoint list - mapping handles this gracefully

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  IBudget,
  IBudgetAllocation,
  IBudgetAlert,
  BudgetStatus,
  VendorCategory,
  Currency
} from '../models/IProcurement';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export interface IBudgetFilter {
  searchTerm?: string;
  status?: BudgetStatus[];
  department?: string;
  fiscalYear?: number;
  category?: VendorCategory;
  belowThreshold?: boolean;
}

export class BudgetService {
  private sp: SPFI;
  private readonly BUDGETS_LIST = 'JML_Budgets';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==================== Budget CRUD Operations ====================

  public async getBudgets(filter?: IBudgetFilter): Promise<IBudget[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.BUDGETS_LIST).items
        .select(
          'Id', 'Title', 'BudgetCode', 'FiscalYear', 'Department', 'CostCenter',
          'Category', 'Status', 'BudgetAmount', 'AllocatedAmount', 'SpentAmount',
          'RemainingAmount', 'Currency', 'WarningThreshold', 'CriticalThreshold',
          'StartDate', 'EndDate', 'Notes',
          'Created', 'Modified', 'Author/Title', 'Editor/Title'
        )
        .expand('Author', 'Editor');

      // Apply filters
      if (filter) {
        const filters: string[] = [];

        if (filter.searchTerm) {
          const term = ValidationUtils.sanitizeForOData(filter.searchTerm);
          filters.push(`(substringof('${term}', Title) or substringof('${term}', BudgetCode) or substringof('${term}', Department))`);
        }

        if (filter.status && filter.status.length > 0) {
          const statusFilters = filter.status.map(s =>
            ValidationUtils.buildFilter('Status', 'eq', s)
          );
          filters.push(`(${statusFilters.join(' or ')})`);
        }

        if (filter.department) {
          const dept = ValidationUtils.sanitizeForOData(filter.department);
          filters.push(ValidationUtils.buildFilter('Department', 'eq', dept));
        }

        if (filter.fiscalYear !== undefined) {
          filters.push(`FiscalYear eq ${filter.fiscalYear}`);
        }

        if (filter.category) {
          filters.push(ValidationUtils.buildFilter('Category', 'eq', filter.category));
        }

        if (filters.length > 0) {
          query = query.filter(filters.join(' and '));
        }
      }

      const items = await query.orderBy('Department', true).top(5000)();
      return items.map(this.mapBudgetFromSP);
    } catch (error) {
      logger.error('BudgetService', 'Error getting budgets:', error);
      throw error;
    }
  }

  public async getBudgetById(id: number): Promise<IBudget> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const item = await this.sp.web.lists.getByTitle(this.BUDGETS_LIST).items
        .getById(validId)
        .select(
          'Id', 'Title', 'BudgetCode', 'FiscalYear', 'Department', 'CostCenter',
          'Category', 'Status', 'BudgetAmount', 'AllocatedAmount', 'SpentAmount',
          'RemainingAmount', 'Currency', 'WarningThreshold', 'CriticalThreshold',
          'StartDate', 'EndDate', 'Notes',
          'Created', 'Modified', 'Author/Title', 'Editor/Title'
        )
        .expand('Author', 'Editor')();

      return this.mapBudgetFromSP(item);
    } catch (error) {
      logger.error('BudgetService', 'Error getting budget by ID:', error);
      throw error;
    }
  }

  public async getBudgetByCode(budgetCode: string): Promise<IBudget | null> {
    try {
      if (!budgetCode || typeof budgetCode !== 'string') {
        throw new Error('Invalid budget code');
      }

      const validCode = ValidationUtils.sanitizeForOData(budgetCode.substring(0, 50));
      const filter = ValidationUtils.buildFilter('BudgetCode', 'eq', validCode);

      const items = await this.sp.web.lists.getByTitle(this.BUDGETS_LIST).items
        .select('Id', 'BudgetCode')
        .filter(filter)
        .top(1)();

      if (items.length === 0) {
        return null;
      }

      return this.getBudgetById(items[0].Id);
    } catch (error) {
      logger.error('BudgetService', 'Error getting budget by code:', error);
      throw error;
    }
  }

  public async createBudget(budget: Partial<IBudget>): Promise<number> {
    try {
      // Validate required fields
      if (!budget.Title || !budget.Department || !budget.FiscalYear || budget.BudgetAmount === undefined) {
        throw new Error('Title, Department, FiscalYear, and BudgetAmount are required');
      }

      // Generate budget code
      const budgetCode = budget.BudgetCode || await this.generateBudgetCode(budget.Department, budget.FiscalYear);

      // Check if budget code already exists
      const existing = await this.getBudgetByCode(budgetCode);
      if (existing) {
        throw new Error(`Budget code ${budgetCode} already exists`);
      }

      const itemData: Record<string, unknown> = {
        Title: ValidationUtils.sanitizeHtml(budget.Title),
        BudgetCode: budgetCode,
        FiscalYear: budget.FiscalYear,
        Department: ValidationUtils.sanitizeHtml(budget.Department),
        Status: budget.Status || BudgetStatus.Active,
        BudgetAmount: budget.BudgetAmount,
        AllocatedAmount: budget.AllocatedAmount || 0,
        SpentAmount: budget.SpentAmount || 0,
        RemainingAmount: budget.BudgetAmount - (budget.SpentAmount || 0),
        Currency: budget.Currency || Currency.GBP,
        WarningThreshold: budget.WarningThreshold || 80,
        CriticalThreshold: budget.CriticalThreshold || 95,
        StartDate: budget.StartDate ? ValidationUtils.validateDate(budget.StartDate, 'StartDate') : new Date(`${budget.FiscalYear}-01-01`).toISOString(),
        EndDate: budget.EndDate ? ValidationUtils.validateDate(budget.EndDate, 'EndDate') : new Date(`${budget.FiscalYear}-12-31`).toISOString()
      };

      // Optional fields
      if (budget.CostCenter) itemData.CostCenter = ValidationUtils.sanitizeHtml(budget.CostCenter);
      if (budget.Category) itemData.Category = budget.Category;
      if (budget.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(budget.Notes);

      const result = await this.sp.web.lists.getByTitle(this.BUDGETS_LIST).items.add(itemData);
      return result.data.Id;
    } catch (error) {
      logger.error('BudgetService', 'Error creating budget:', error);
      throw error;
    }
  }

  public async updateBudget(id: number, updates: Partial<IBudget>): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const itemData: Record<string, unknown> = {};

      if (updates.Title) itemData.Title = ValidationUtils.sanitizeHtml(updates.Title);
      if (updates.Status) {
        ValidationUtils.validateEnum(updates.Status, BudgetStatus, 'Status');
        itemData.Status = updates.Status;
      }
      if (updates.BudgetAmount !== undefined) itemData.BudgetAmount = updates.BudgetAmount;
      if (updates.AllocatedAmount !== undefined) itemData.AllocatedAmount = updates.AllocatedAmount;
      if (updates.SpentAmount !== undefined) itemData.SpentAmount = updates.SpentAmount;
      if (updates.RemainingAmount !== undefined) itemData.RemainingAmount = updates.RemainingAmount;
      if (updates.Currency) itemData.Currency = updates.Currency;
      if (updates.WarningThreshold !== undefined) itemData.WarningThreshold = updates.WarningThreshold;
      if (updates.CriticalThreshold !== undefined) itemData.CriticalThreshold = updates.CriticalThreshold;
      if (updates.StartDate) itemData.StartDate = ValidationUtils.validateDate(updates.StartDate, 'StartDate');
      if (updates.EndDate) itemData.EndDate = ValidationUtils.validateDate(updates.EndDate, 'EndDate');
      if (updates.CostCenter !== undefined) itemData.CostCenter = updates.CostCenter ? ValidationUtils.sanitizeHtml(updates.CostCenter) : null;
      if (updates.Category !== undefined) itemData.Category = updates.Category;
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes ? ValidationUtils.sanitizeHtml(updates.Notes) : null;

      await this.sp.web.lists.getByTitle(this.BUDGETS_LIST).items.getById(validId).update(itemData);
    } catch (error) {
      logger.error('BudgetService', 'Error updating budget:', error);
      throw error;
    }
  }

  public async deleteBudget(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Check if budget is in Pending status
      const budget = await this.getBudgetById(validId);
      if (budget.Status !== BudgetStatus.Pending) {
        throw new Error('Only pending budgets can be deleted');
      }

      await this.sp.web.lists.getByTitle(this.BUDGETS_LIST).items.getById(validId).delete();
    } catch (error) {
      logger.error('BudgetService', 'Error deleting budget:', error);
      throw error;
    }
  }

  // ==================== Budget Operations ====================

  public async allocateBudget(id: number, amount: number, documentType: string, documentId: number, documentNumber: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const budget = await this.getBudgetById(validId);

      if (budget.Status !== BudgetStatus.Active) {
        throw new Error('Can only allocate from active budgets');
      }

      const newAllocated = (budget.AllocatedAmount || 0) + amount;
      const newRemaining = budget.BudgetAmount - (budget.SpentAmount || 0) - amount;

      if (newRemaining < 0) {
        throw new Error(`Insufficient budget. Requested: ${amount}, Available: ${budget.RemainingAmount}`);
      }

      await this.updateBudget(validId, {
        AllocatedAmount: newAllocated,
        RemainingAmount: newRemaining,
        Notes: budget.Notes
          ? `${budget.Notes}\n\nAllocated ${amount} for ${documentType} ${documentNumber}`
          : `Allocated ${amount} for ${documentType} ${documentNumber}`
      });
    } catch (error) {
      logger.error('BudgetService', 'Error allocating budget:', error);
      throw error;
    }
  }

  public async releaseAllocation(id: number, amount: number, documentType: string, documentNumber: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const budget = await this.getBudgetById(validId);

      const newAllocated = Math.max(0, (budget.AllocatedAmount || 0) - amount);
      const newRemaining = budget.BudgetAmount - (budget.SpentAmount || 0);

      await this.updateBudget(validId, {
        AllocatedAmount: newAllocated,
        RemainingAmount: newRemaining,
        Notes: budget.Notes
          ? `${budget.Notes}\n\nReleased ${amount} from ${documentType} ${documentNumber}`
          : `Released ${amount} from ${documentType} ${documentNumber}`
      });
    } catch (error) {
      logger.error('BudgetService', 'Error releasing budget allocation:', error);
      throw error;
    }
  }

  public async recordSpend(id: number, amount: number, documentType: string, documentNumber: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const budget = await this.getBudgetById(validId);

      const newSpent = (budget.SpentAmount || 0) + amount;
      const newAllocated = Math.max(0, (budget.AllocatedAmount || 0) - amount);
      const newRemaining = budget.BudgetAmount - newSpent;

      await this.updateBudget(validId, {
        SpentAmount: newSpent,
        AllocatedAmount: newAllocated,
        RemainingAmount: newRemaining,
        Notes: budget.Notes
          ? `${budget.Notes}\n\nSpent ${amount} on ${documentType} ${documentNumber}`
          : `Spent ${amount} on ${documentType} ${documentNumber}`
      });
    } catch (error) {
      logger.error('BudgetService', 'Error recording spend:', error);
      throw error;
    }
  }

  public async freezeBudget(id: number, reason?: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const budget = await this.getBudgetById(validId);

      const updateData: Partial<IBudget> = {
        Status: BudgetStatus.Frozen
      };

      if (reason) {
        updateData.Notes = budget.Notes
          ? `${budget.Notes}\n\nFrozen: ${reason}`
          : `Frozen: ${reason}`;
      }

      await this.updateBudget(validId, updateData);
    } catch (error) {
      logger.error('BudgetService', 'Error freezing budget:', error);
      throw error;
    }
  }

  public async unfreezeBudget(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const budget = await this.getBudgetById(validId);

      if (budget.Status !== BudgetStatus.Frozen) {
        throw new Error('Budget is not frozen');
      }

      await this.updateBudget(validId, {
        Status: BudgetStatus.Active,
        Notes: budget.Notes
          ? `${budget.Notes}\n\nUnfrozen on ${new Date().toISOString()}`
          : `Unfrozen on ${new Date().toISOString()}`
      });
    } catch (error) {
      logger.error('BudgetService', 'Error unfreezing budget:', error);
      throw error;
    }
  }

  public async closeBudget(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      await this.updateBudget(validId, {
        Status: BudgetStatus.Closed
      });
    } catch (error) {
      logger.error('BudgetService', 'Error closing budget:', error);
      throw error;
    }
  }

  // ==================== Query Methods ====================

  public async getActiveBudgets(): Promise<IBudget[]> {
    try {
      return this.getBudgets({ status: [BudgetStatus.Active] });
    } catch (error) {
      logger.error('BudgetService', 'Error getting active budgets:', error);
      throw error;
    }
  }

  public async getDepartmentBudgets(department: string): Promise<IBudget[]> {
    try {
      return this.getBudgets({ department, status: [BudgetStatus.Active] });
    } catch (error) {
      logger.error('BudgetService', 'Error getting department budgets:', error);
      throw error;
    }
  }

  public async getCurrentYearBudgets(): Promise<IBudget[]> {
    try {
      const currentYear = new Date().getFullYear();
      return this.getBudgets({ fiscalYear: currentYear });
    } catch (error) {
      logger.error('BudgetService', 'Error getting current year budgets:', error);
      throw error;
    }
  }

  public async getBudgetAlerts(): Promise<IBudgetAlert[]> {
    try {
      const budgets = await this.getActiveBudgets();
      const alerts: IBudgetAlert[] = [];

      for (const budget of budgets) {
        const utilization = this.calculateUtilization(budget);

        if (utilization >= budget.CriticalThreshold) {
          alerts.push({
            budgetId: budget.Id!,
            budgetName: budget.Title,
            department: budget.Department,
            alertType: utilization >= 100 ? 'Exceeded' : 'Critical',
            utilized: utilization,
            threshold: budget.CriticalThreshold,
            message: utilization >= 100
              ? `Budget ${budget.BudgetCode} has exceeded the limit (${utilization.toFixed(1)}%)`
              : `Budget ${budget.BudgetCode} is at critical level (${utilization.toFixed(1)}%)`
          });
        } else if (utilization >= budget.WarningThreshold) {
          alerts.push({
            budgetId: budget.Id!,
            budgetName: budget.Title,
            department: budget.Department,
            alertType: 'Warning',
            utilized: utilization,
            threshold: budget.WarningThreshold,
            message: `Budget ${budget.BudgetCode} is approaching limit (${utilization.toFixed(1)}%)`
          });
        }
      }

      return alerts.sort((a, b) => b.utilized - a.utilized);
    } catch (error) {
      logger.error('BudgetService', 'Error getting budget alerts:', error);
      throw error;
    }
  }

  // ==================== Statistics ====================

  public async getBudgetStatistics(): Promise<{
    total: number;
    active: number;
    frozen: number;
    closed: number;
    totalBudget: number;
    totalAllocated: number;
    totalSpent: number;
    totalRemaining: number;
    avgUtilization: number;
    overBudget: number;
    nearLimit: number;
    byDepartment: { department: string; budget: number; spent: number; utilization: number }[];
  }> {
    try {
      const budgets = await this.getBudgets();

      const stats = {
        total: budgets.length,
        active: 0,
        frozen: 0,
        closed: 0,
        totalBudget: 0,
        totalAllocated: 0,
        totalSpent: 0,
        totalRemaining: 0,
        avgUtilization: 0,
        overBudget: 0,
        nearLimit: 0,
        byDepartment: [] as { department: string; budget: number; spent: number; utilization: number }[]
      };

      const departmentMap = new Map<string, { budget: number; spent: number }>();
      let totalUtilization = 0;
      let utilizationCount = 0;

      for (const budget of budgets) {
        // Count by status
        switch (budget.Status) {
          case BudgetStatus.Active:
            stats.active++;
            break;
          case BudgetStatus.Frozen:
            stats.frozen++;
            break;
          case BudgetStatus.Closed:
            stats.closed++;
            break;
        }

        // Calculate totals for active budgets
        if (budget.Status === BudgetStatus.Active) {
          stats.totalBudget += budget.BudgetAmount || 0;
          stats.totalAllocated += budget.AllocatedAmount || 0;
          stats.totalSpent += budget.SpentAmount || 0;
          stats.totalRemaining += budget.RemainingAmount || 0;

          const utilization = this.calculateUtilization(budget);
          totalUtilization += utilization;
          utilizationCount++;

          if (utilization >= 100) {
            stats.overBudget++;
          } else if (utilization >= budget.WarningThreshold) {
            stats.nearLimit++;
          }

          // By department
          const existing = departmentMap.get(budget.Department);
          if (existing) {
            existing.budget += budget.BudgetAmount || 0;
            existing.spent += budget.SpentAmount || 0;
          } else {
            departmentMap.set(budget.Department, {
              budget: budget.BudgetAmount || 0,
              spent: budget.SpentAmount || 0
            });
          }
        }
      }

      stats.avgUtilization = utilizationCount > 0 ? totalUtilization / utilizationCount : 0;

      // Convert department map to array
      stats.byDepartment = Array.from(departmentMap.entries())
        .map(([department, data]) => ({
          department,
          budget: data.budget,
          spent: data.spent,
          utilization: data.budget > 0 ? (data.spent / data.budget) * 100 : 0
        }))
        .sort((a, b) => b.utilization - a.utilization);

      return stats;
    } catch (error) {
      logger.error('BudgetService', 'Error getting budget statistics:', error);
      throw error;
    }
  }

  // ==================== Helper Functions ====================

  private calculateUtilization(budget: IBudget): number {
    if (!budget.BudgetAmount || budget.BudgetAmount === 0) {
      return 0;
    }
    return ((budget.SpentAmount || 0) / budget.BudgetAmount) * 100;
  }

  private async generateBudgetCode(department: string, fiscalYear: number): Promise<string> {
    try {
      const deptCode = department.substring(0, 3).toUpperCase();
      const prefix = `BUD-${fiscalYear}-${deptCode}-`;

      const items = await this.sp.web.lists.getByTitle(this.BUDGETS_LIST).items
        .select('BudgetCode')
        .filter(`substringof('${prefix}', BudgetCode)`)
        .orderBy('Id', false)
        .top(1)();

      let nextNumber = 1;
      if (items.length > 0 && items[0].BudgetCode) {
        const match = items[0].BudgetCode.match(new RegExp(`BUD-${fiscalYear}-${deptCode}-(\\d+)`));
        if (match) {
          nextNumber = parseInt(match[1], 10) + 1;
        }
      }

      return `${prefix}${nextNumber.toString().padStart(2, '0')}`;
    } catch (error) {
      logger.error('BudgetService', 'Error generating budget code:', error);
      return `BUD-${Date.now()}`;
    }
  }

  // ==================== Mapping Functions ====================

  private mapBudgetFromSP(item: any): IBudget {
    // Map SharePoint fields to interface, handling missing fields gracefully
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      BudgetCode: item.BudgetCode as string,
      FiscalYear: item.FiscalYear as number,
      Department: item.Department as string,
      CostCenter: item.CostCenter as string,
      Category: item.Category as VendorCategory,
      Status: item.Status as BudgetStatus || BudgetStatus.Active,
      BudgetAmount: item.BudgetAmount as number || 0,
      AllocatedAmount: item.AllocatedAmount as number || 0,
      SpentAmount: item.SpentAmount as number || 0,
      RemainingAmount: item.RemainingAmount as number || 0,
      Currency: item.Currency as Currency || Currency.GBP,
      WarningThreshold: item.WarningThreshold as number || 80,
      CriticalThreshold: item.CriticalThreshold as number || 95,
      StartDate: item.StartDate ? new Date(item.StartDate as string) : new Date(),
      EndDate: item.EndDate ? new Date(item.EndDate as string) : new Date(),
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }
}
