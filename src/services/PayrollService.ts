// @ts-nocheck
/* eslint-disable @typescript-eslint/no-explicit-any */
// Payroll Service
// Payroll summary tracking and department payroll management

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  IPayrollSummary,
  IPayrollFilter,
  PayrollStatus
} from '../models/IFinancialManagement';
import { Currency } from '../models/IProcurement';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class PayrollService {
  private sp: SPFI;
  private readonly PAYROLL_LIST = 'PM_PayrollSummary';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==================== Payroll CRUD Operations ====================

  public async getPayrollSummaries(filter?: IPayrollFilter): Promise<IPayrollSummary[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.PAYROLL_LIST).items
        .select(
          'Id', 'Title', 'PayrollCode', 'PeriodStart', 'PeriodEnd', 'FiscalYear', 'FiscalMonth',
          'CostCenter', 'Department', 'Status', 'Currency',
          'HeadCount', 'GrossPay', 'Deductions', 'NetPay', 'EmployerContributions', 'TotalCost',
          'BudgetAmount', 'Variance', 'VariancePercent',
          'OvertimeHours', 'OvertimeCost', 'BonusPayments',
          'TaxWithheld', 'PensionContributions', 'NIContributions',
          'ProcessedDate', 'Notes',
          'Created', 'Modified', 'Author/Title', 'Editor/Title',
          'ProcessedBy/Id', 'ProcessedBy/Title'
        )
        .expand('Author', 'Editor', 'ProcessedBy');

      // Apply filters
      if (filter) {
        const filters: string[] = [];

        if (filter.department) {
          const dept = ValidationUtils.sanitizeForOData(filter.department);
          filters.push(ValidationUtils.buildFilter('Department', 'eq', dept));
        }

        if (filter.costCenter) {
          const cc = ValidationUtils.sanitizeForOData(filter.costCenter);
          filters.push(ValidationUtils.buildFilter('CostCenter', 'eq', cc));
        }

        if (filter.fiscalYear !== undefined) {
          filters.push(`FiscalYear eq ${filter.fiscalYear}`);
        }

        if (filter.fiscalMonth !== undefined) {
          filters.push(`FiscalMonth eq ${filter.fiscalMonth}`);
        }

        if (filter.status && filter.status.length > 0) {
          const statusFilters = filter.status.map(s =>
            ValidationUtils.buildFilter('Status', 'eq', s)
          );
          filters.push(`(${statusFilters.join(' or ')})`);
        }

        if (filter.fromDate) {
          filters.push(`PeriodStart ge datetime'${filter.fromDate.toISOString()}'`);
        }

        if (filter.toDate) {
          filters.push(`PeriodEnd le datetime'${filter.toDate.toISOString()}'`);
        }

        if (filters.length > 0) {
          query = query.filter(filters.join(' and '));
        }
      }

      const items = await query.orderBy('PeriodStart', false).top(5000)();
      return items.map(this.mapPayrollFromSP);
    } catch (error) {
      logger.error('PayrollService', 'Error getting payroll summaries:', error);
      throw error;
    }
  }

  public async getPayrollById(id: number): Promise<IPayrollSummary> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const item = await this.sp.web.lists.getByTitle(this.PAYROLL_LIST).items
        .getById(validId)
        .select(
          'Id', 'Title', 'PayrollCode', 'PeriodStart', 'PeriodEnd', 'FiscalYear', 'FiscalMonth',
          'CostCenter', 'Department', 'Status', 'Currency',
          'HeadCount', 'GrossPay', 'Deductions', 'NetPay', 'EmployerContributions', 'TotalCost',
          'BudgetAmount', 'Variance', 'VariancePercent',
          'OvertimeHours', 'OvertimeCost', 'BonusPayments',
          'TaxWithheld', 'PensionContributions', 'NIContributions',
          'ProcessedDate', 'Notes',
          'Created', 'Modified', 'Author/Title', 'Editor/Title',
          'ProcessedBy/Id', 'ProcessedBy/Title'
        )
        .expand('Author', 'Editor', 'ProcessedBy')();

      return this.mapPayrollFromSP(item);
    } catch (error) {
      logger.error('PayrollService', 'Error getting payroll by ID:', error);
      throw error;
    }
  }

  public async getPayrollByCode(payrollCode: string): Promise<IPayrollSummary | null> {
    try {
      if (!payrollCode || typeof payrollCode !== 'string') {
        throw new Error('Invalid payroll code');
      }

      const validCode = ValidationUtils.sanitizeForOData(payrollCode.substring(0, 50));
      const filter = ValidationUtils.buildFilter('PayrollCode', 'eq', validCode);

      const items = await this.sp.web.lists.getByTitle(this.PAYROLL_LIST).items
        .select('Id', 'PayrollCode')
        .filter(filter)
        .top(1)();

      if (items.length === 0) {
        return null;
      }

      return this.getPayrollById(items[0].Id);
    } catch (error) {
      logger.error('PayrollService', 'Error getting payroll by code:', error);
      throw error;
    }
  }

  public async createPayrollSummary(payroll: Partial<IPayrollSummary>): Promise<number> {
    try {
      // Validate required fields
      if (!payroll.Title || !payroll.Department || !payroll.PeriodStart || !payroll.PeriodEnd) {
        throw new Error('Title, Department, PeriodStart, and PeriodEnd are required');
      }

      // Generate payroll code
      const payrollCode = payroll.PayrollCode || await this.generatePayrollCode(payroll.Department, payroll.FiscalYear, payroll.FiscalMonth);

      const itemData: Record<string, unknown> = {
        Title: ValidationUtils.sanitizeHtml(payroll.Title),
        PayrollCode: payrollCode,
        PeriodStart: ValidationUtils.validateDate(payroll.PeriodStart, 'PeriodStart'),
        PeriodEnd: ValidationUtils.validateDate(payroll.PeriodEnd, 'PeriodEnd'),
        FiscalYear: payroll.FiscalYear || new Date(payroll.PeriodStart).getFullYear(),
        FiscalMonth: payroll.FiscalMonth || new Date(payroll.PeriodStart).getMonth() + 1,
        Department: ValidationUtils.sanitizeHtml(payroll.Department),
        Status: payroll.Status || PayrollStatus.Draft,
        Currency: payroll.Currency || Currency.GBP,
        HeadCount: payroll.HeadCount || 0,
        GrossPay: payroll.GrossPay || 0,
        Deductions: payroll.Deductions || 0,
        NetPay: payroll.NetPay || 0,
        EmployerContributions: payroll.EmployerContributions || 0,
        TotalCost: payroll.TotalCost || 0,
        BudgetAmount: payroll.BudgetAmount || 0,
        Variance: payroll.Variance || 0,
        VariancePercent: payroll.VariancePercent || 0
      };

      // Optional fields
      if (payroll.CostCenter) itemData.CostCenter = ValidationUtils.sanitizeHtml(payroll.CostCenter);
      if (payroll.OvertimeHours !== undefined) itemData.OvertimeHours = payroll.OvertimeHours;
      if (payroll.OvertimeCost !== undefined) itemData.OvertimeCost = payroll.OvertimeCost;
      if (payroll.BonusPayments !== undefined) itemData.BonusPayments = payroll.BonusPayments;
      if (payroll.TaxWithheld !== undefined) itemData.TaxWithheld = payroll.TaxWithheld;
      if (payroll.PensionContributions !== undefined) itemData.PensionContributions = payroll.PensionContributions;
      if (payroll.NIContributions !== undefined) itemData.NIContributions = payroll.NIContributions;
      if (payroll.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(payroll.Notes);

      const result = await this.sp.web.lists.getByTitle(this.PAYROLL_LIST).items.add(itemData);
      return result.data.Id;
    } catch (error) {
      logger.error('PayrollService', 'Error creating payroll summary:', error);
      throw error;
    }
  }

  public async updatePayrollSummary(id: number, updates: Partial<IPayrollSummary>): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const itemData: Record<string, unknown> = {};

      if (updates.Title) itemData.Title = ValidationUtils.sanitizeHtml(updates.Title);
      if (updates.Status) {
        ValidationUtils.validateEnum(updates.Status, PayrollStatus, 'Status');
        itemData.Status = updates.Status;
      }
      if (updates.HeadCount !== undefined) itemData.HeadCount = updates.HeadCount;
      if (updates.GrossPay !== undefined) itemData.GrossPay = updates.GrossPay;
      if (updates.Deductions !== undefined) itemData.Deductions = updates.Deductions;
      if (updates.NetPay !== undefined) itemData.NetPay = updates.NetPay;
      if (updates.EmployerContributions !== undefined) itemData.EmployerContributions = updates.EmployerContributions;
      if (updates.TotalCost !== undefined) itemData.TotalCost = updates.TotalCost;
      if (updates.BudgetAmount !== undefined) itemData.BudgetAmount = updates.BudgetAmount;
      if (updates.Variance !== undefined) itemData.Variance = updates.Variance;
      if (updates.VariancePercent !== undefined) itemData.VariancePercent = updates.VariancePercent;
      if (updates.OvertimeHours !== undefined) itemData.OvertimeHours = updates.OvertimeHours;
      if (updates.OvertimeCost !== undefined) itemData.OvertimeCost = updates.OvertimeCost;
      if (updates.BonusPayments !== undefined) itemData.BonusPayments = updates.BonusPayments;
      if (updates.TaxWithheld !== undefined) itemData.TaxWithheld = updates.TaxWithheld;
      if (updates.PensionContributions !== undefined) itemData.PensionContributions = updates.PensionContributions;
      if (updates.NIContributions !== undefined) itemData.NIContributions = updates.NIContributions;
      if (updates.ProcessedDate) itemData.ProcessedDate = ValidationUtils.validateDate(updates.ProcessedDate, 'ProcessedDate');
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes ? ValidationUtils.sanitizeHtml(updates.Notes) : null;

      await this.sp.web.lists.getByTitle(this.PAYROLL_LIST).items.getById(validId).update(itemData);
    } catch (error) {
      logger.error('PayrollService', 'Error updating payroll summary:', error);
      throw error;
    }
  }

  public async deletePayrollSummary(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Check if payroll is in Draft status
      const payroll = await this.getPayrollById(validId);
      if (payroll.Status !== PayrollStatus.Draft) {
        throw new Error('Only draft payroll summaries can be deleted');
      }

      await this.sp.web.lists.getByTitle(this.PAYROLL_LIST).items.getById(validId).delete();
    } catch (error) {
      logger.error('PayrollService', 'Error deleting payroll summary:', error);
      throw error;
    }
  }

  // ==================== Payroll Workflow Operations ====================

  public async processPayroll(id: number, processedById: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const payroll = await this.getPayrollById(validId);

      if (payroll.Status !== PayrollStatus.Draft && payroll.Status !== PayrollStatus.Processing) {
        throw new Error('Only draft or processing payroll summaries can be processed');
      }

      await this.updatePayrollSummary(validId, {
        Status: PayrollStatus.Processed,
        ProcessedDate: new Date(),
        ProcessedById: processedById
      });
    } catch (error) {
      logger.error('PayrollService', 'Error processing payroll:', error);
      throw error;
    }
  }

  public async finalizePayroll(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const payroll = await this.getPayrollById(validId);

      if (payroll.Status !== PayrollStatus.Processed) {
        throw new Error('Only processed payroll summaries can be finalized');
      }

      await this.updatePayrollSummary(validId, {
        Status: PayrollStatus.Finalized
      });
    } catch (error) {
      logger.error('PayrollService', 'Error finalizing payroll:', error);
      throw error;
    }
  }

  // ==================== Query Methods ====================

  public async getCurrentMonthPayroll(): Promise<IPayrollSummary[]> {
    try {
      const now = new Date();
      return this.getPayrollSummaries({
        fiscalYear: now.getFullYear(),
        fiscalMonth: now.getMonth() + 1
      });
    } catch (error) {
      logger.error('PayrollService', 'Error getting current month payroll:', error);
      throw error;
    }
  }

  public async getDepartmentPayroll(department: string): Promise<IPayrollSummary[]> {
    try {
      return this.getPayrollSummaries({ department });
    } catch (error) {
      logger.error('PayrollService', 'Error getting department payroll:', error);
      throw error;
    }
  }

  public async getYearPayroll(fiscalYear: number): Promise<IPayrollSummary[]> {
    try {
      return this.getPayrollSummaries({ fiscalYear });
    } catch (error) {
      logger.error('PayrollService', 'Error getting year payroll:', error);
      throw error;
    }
  }

  // ==================== Statistics ====================

  public async getPayrollStatistics(): Promise<{
    totalPayrollCost: number;
    totalHeadCount: number;
    avgCostPerEmployee: number;
    totalGrossPay: number;
    totalDeductions: number;
    totalEmployerContributions: number;
    totalVariance: number;
    byDepartment: { department: string; headCount: number; totalCost: number; variance: number }[];
    byMonth: { month: string; totalCost: number; headCount: number }[];
    overtimeTotal: number;
    bonusTotal: number;
  }> {
    try {
      const currentYear = new Date().getFullYear();
      const payrolls = await this.getPayrollSummaries({ fiscalYear: currentYear });

      const stats = {
        totalPayrollCost: 0,
        totalHeadCount: 0,
        avgCostPerEmployee: 0,
        totalGrossPay: 0,
        totalDeductions: 0,
        totalEmployerContributions: 0,
        totalVariance: 0,
        byDepartment: [] as { department: string; headCount: number; totalCost: number; variance: number }[],
        byMonth: [] as { month: string; totalCost: number; headCount: number }[],
        overtimeTotal: 0,
        bonusTotal: 0
      };

      const departmentMap = new Map<string, { headCount: number; totalCost: number; variance: number }>();
      const monthMap = new Map<number, { totalCost: number; headCount: number }>();

      // Get most recent month per department to avoid counting headcount multiple times
      const latestByDept = new Map<string, IPayrollSummary>();

      for (const payroll of payrolls) {
        // Track latest payroll per department for accurate headcount
        const existing = latestByDept.get(payroll.Department);
        if (!existing || payroll.FiscalMonth > existing.FiscalMonth) {
          latestByDept.set(payroll.Department, payroll);
        }

        stats.totalPayrollCost += payroll.TotalCost || 0;
        stats.totalGrossPay += payroll.GrossPay || 0;
        stats.totalDeductions += payroll.Deductions || 0;
        stats.totalEmployerContributions += payroll.EmployerContributions || 0;
        stats.totalVariance += payroll.Variance || 0;
        stats.overtimeTotal += payroll.OvertimeCost || 0;
        stats.bonusTotal += payroll.BonusPayments || 0;

        // By department (aggregate)
        const deptData = departmentMap.get(payroll.Department);
        if (deptData) {
          deptData.totalCost += payroll.TotalCost || 0;
          deptData.variance += payroll.Variance || 0;
        } else {
          departmentMap.set(payroll.Department, {
            headCount: payroll.HeadCount || 0,
            totalCost: payroll.TotalCost || 0,
            variance: payroll.Variance || 0
          });
        }

        // By month
        const monthKey = payroll.FiscalMonth;
        const monthData = monthMap.get(monthKey);
        if (monthData) {
          monthData.totalCost += payroll.TotalCost || 0;
          monthData.headCount += payroll.HeadCount || 0;
        } else {
          monthMap.set(monthKey, {
            totalCost: payroll.TotalCost || 0,
            headCount: payroll.HeadCount || 0
          });
        }
      }

      // Calculate total headcount from latest payroll entries
      Array.from(latestByDept.values()).forEach(payroll => {
        stats.totalHeadCount += payroll.HeadCount || 0;
      });

      stats.avgCostPerEmployee = stats.totalHeadCount > 0
        ? stats.totalPayrollCost / stats.totalHeadCount
        : 0;

      // Convert department map - use headcount from latest payroll
      stats.byDepartment = Array.from(departmentMap.entries())
        .map(([department, data]) => {
          const latestPayroll = latestByDept.get(department);
          return {
            department,
            headCount: latestPayroll?.HeadCount || data.headCount,
            totalCost: data.totalCost,
            variance: data.variance
          };
        })
        .sort((a, b) => b.totalCost - a.totalCost);

      // Convert month map
      const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      stats.byMonth = Array.from(monthMap.entries())
        .map(([month, data]) => ({
          month: `${monthNames[month - 1]} ${currentYear}`,
          totalCost: data.totalCost,
          headCount: data.headCount
        }))
        .sort((a, b) => {
          const aMonth = monthNames.indexOf(a.month.split(' ')[0]);
          const bMonth = monthNames.indexOf(b.month.split(' ')[0]);
          return aMonth - bMonth;
        });

      return stats;
    } catch (error) {
      logger.error('PayrollService', 'Error getting payroll statistics:', error);
      throw error;
    }
  }

  // ==================== Helper Functions ====================

  private async generatePayrollCode(department: string, fiscalYear?: number, fiscalMonth?: number): Promise<string> {
    try {
      const year = fiscalYear || new Date().getFullYear();
      const month = fiscalMonth || new Date().getMonth() + 1;
      const deptCode = department.substring(0, 3).toUpperCase();

      return `PAY-${year}-${month.toString().padStart(2, '0')}-${deptCode}`;
    } catch (error) {
      logger.error('PayrollService', 'Error generating payroll code:', error);
      return `PAY-${Date.now()}`;
    }
  }

  // ==================== Mapping Functions ====================

  private mapPayrollFromSP(item: any): IPayrollSummary {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      PayrollCode: item.PayrollCode as string,
      PeriodStart: item.PeriodStart ? new Date(item.PeriodStart as string) : new Date(),
      PeriodEnd: item.PeriodEnd ? new Date(item.PeriodEnd as string) : new Date(),
      FiscalYear: item.FiscalYear as number,
      FiscalMonth: item.FiscalMonth as number,
      CostCenter: item.CostCenter as string,
      Department: item.Department as string,
      Status: item.Status as PayrollStatus || PayrollStatus.Draft,
      Currency: item.Currency as Currency || Currency.GBP,
      HeadCount: item.HeadCount as number || 0,
      GrossPay: item.GrossPay as number || 0,
      Deductions: item.Deductions as number || 0,
      NetPay: item.NetPay as number || 0,
      EmployerContributions: item.EmployerContributions as number || 0,
      TotalCost: item.TotalCost as number || 0,
      BudgetAmount: item.BudgetAmount as number || 0,
      Variance: item.Variance as number || 0,
      VariancePercent: item.VariancePercent as number || 0,
      OvertimeHours: item.OvertimeHours as number,
      OvertimeCost: item.OvertimeCost as number,
      BonusPayments: item.BonusPayments as number,
      TaxWithheld: item.TaxWithheld as number,
      PensionContributions: item.PensionContributions as number,
      NIContributions: item.NIContributions as number,
      ProcessedDate: item.ProcessedDate ? new Date(item.ProcessedDate as string) : undefined,
      ProcessedById: item.ProcessedById as number,
      ProcessedBy: item.ProcessedBy ? {
        Id: item.ProcessedBy.Id,
        Title: item.ProcessedBy.Title
      } : undefined,
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }
}
