// @ts-nocheck
/* eslint-disable @typescript-eslint/no-explicit-any */
// Expense Service
// Expense tracking, submission, and approval management

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  IExpense,
  IExpenseFilter,
  ExpenseStatus,
  ExpenseCategory
} from '../models/IFinancialManagement';
import { Currency } from '../models/IProcurement';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class ExpenseService {
  private sp: SPFI;
  private readonly EXPENSES_LIST = 'JML_Expenses';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==================== Expense CRUD Operations ====================

  public async getExpenses(filter?: IExpenseFilter): Promise<IExpense[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.EXPENSES_LIST).items
        .select(
          'Id', 'Title', 'ExpenseCode', 'ExpenseDate', 'Amount', 'Category',
          'CostCenter', 'Department', 'Status', 'Currency', 'Vendor', 'Notes',
          'SubmittedDate', 'ApprovalDate', 'PaymentDate', 'PaymentReference',
          'RejectionReason', 'ReceiptUrl',
          'Created', 'Modified', 'Author/Id', 'Author/Title', 'Editor/Title',
          'SubmittedBy/Id', 'SubmittedBy/Title', 'ApprovedBy/Id', 'ApprovedBy/Title'
        )
        .expand('Author', 'Editor', 'SubmittedBy', 'ApprovedBy');

      // Apply filters
      if (filter) {
        const filters: string[] = [];

        if (filter.searchTerm) {
          const term = ValidationUtils.sanitizeForOData(filter.searchTerm);
          filters.push(`(substringof('${term}', Title) or substringof('${term}', ExpenseCode) or substringof('${term}', Vendor))`);
        }

        if (filter.status && filter.status.length > 0) {
          const statusFilters = filter.status.map(s =>
            ValidationUtils.buildFilter('Status', 'eq', s)
          );
          filters.push(`(${statusFilters.join(' or ')})`);
        }

        if (filter.category && filter.category.length > 0) {
          const categoryFilters = filter.category.map(c =>
            ValidationUtils.buildFilter('Category', 'eq', c)
          );
          filters.push(`(${categoryFilters.join(' or ')})`);
        }

        if (filter.department) {
          const dept = ValidationUtils.sanitizeForOData(filter.department);
          filters.push(ValidationUtils.buildFilter('Department', 'eq', dept));
        }

        if (filter.costCenter) {
          const cc = ValidationUtils.sanitizeForOData(filter.costCenter);
          filters.push(ValidationUtils.buildFilter('CostCenter', 'eq', cc));
        }

        if (filter.submittedById) {
          filters.push(`SubmittedById eq ${filter.submittedById}`);
        }

        if (filter.fromDate) {
          filters.push(`ExpenseDate ge datetime'${filter.fromDate.toISOString()}'`);
        }

        if (filter.toDate) {
          filters.push(`ExpenseDate le datetime'${filter.toDate.toISOString()}'`);
        }

        if (filter.minAmount !== undefined) {
          filters.push(`Amount ge ${filter.minAmount}`);
        }

        if (filter.maxAmount !== undefined) {
          filters.push(`Amount le ${filter.maxAmount}`);
        }

        if (filters.length > 0) {
          query = query.filter(filters.join(' and '));
        }
      }

      const items = await query.orderBy('ExpenseDate', false).top(5000)();
      return items.map(this.mapExpenseFromSP);
    } catch (error) {
      logger.error('ExpenseService', 'Error getting expenses:', error);
      throw error;
    }
  }

  public async getExpenseById(id: number): Promise<IExpense> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const item = await this.sp.web.lists.getByTitle(this.EXPENSES_LIST).items
        .getById(validId)
        .select(
          'Id', 'Title', 'ExpenseCode', 'ExpenseDate', 'Amount', 'Category',
          'CostCenter', 'Department', 'Status', 'Currency', 'Vendor', 'Notes',
          'SubmittedDate', 'ApprovalDate', 'PaymentDate', 'PaymentReference',
          'RejectionReason', 'ReceiptUrl',
          'Created', 'Modified', 'Author/Id', 'Author/Title', 'Editor/Title',
          'SubmittedBy/Id', 'SubmittedBy/Title', 'ApprovedBy/Id', 'ApprovedBy/Title'
        )
        .expand('Author', 'Editor', 'SubmittedBy', 'ApprovedBy')();

      return this.mapExpenseFromSP(item);
    } catch (error) {
      logger.error('ExpenseService', 'Error getting expense by ID:', error);
      throw error;
    }
  }

  public async getExpenseByCode(expenseCode: string): Promise<IExpense | null> {
    try {
      if (!expenseCode || typeof expenseCode !== 'string') {
        throw new Error('Invalid expense code');
      }

      const validCode = ValidationUtils.sanitizeForOData(expenseCode.substring(0, 50));
      const filter = ValidationUtils.buildFilter('ExpenseCode', 'eq', validCode);

      const items = await this.sp.web.lists.getByTitle(this.EXPENSES_LIST).items
        .select('Id', 'ExpenseCode')
        .filter(filter)
        .top(1)();

      if (items.length === 0) {
        return null;
      }

      return this.getExpenseById(items[0].Id);
    } catch (error) {
      logger.error('ExpenseService', 'Error getting expense by code:', error);
      throw error;
    }
  }

  public async createExpense(expense: Partial<IExpense>): Promise<number> {
    try {
      // Validate required fields
      if (!expense.Title || !expense.Amount || !expense.Category || !expense.Department || !expense.ExpenseDate) {
        throw new Error('Title, Amount, Category, Department, and ExpenseDate are required');
      }

      // Generate expense code
      const expenseCode = expense.ExpenseCode || await this.generateExpenseCode();

      const itemData: Record<string, unknown> = {
        Title: ValidationUtils.sanitizeHtml(expense.Title),
        ExpenseCode: expenseCode,
        ExpenseDate: ValidationUtils.validateDate(expense.ExpenseDate, 'ExpenseDate'),
        Amount: expense.Amount,
        Category: expense.Category,
        Department: ValidationUtils.sanitizeHtml(expense.Department),
        Status: expense.Status || ExpenseStatus.Draft,
        Currency: expense.Currency || Currency.GBP
      };

      // Optional fields
      if (expense.CostCenter) itemData.CostCenter = ValidationUtils.sanitizeHtml(expense.CostCenter);
      if (expense.Vendor) itemData.Vendor = ValidationUtils.sanitizeHtml(expense.Vendor);
      if (expense.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(expense.Notes);
      if (expense.ReceiptUrl) itemData.ReceiptUrl = expense.ReceiptUrl;

      const result = await this.sp.web.lists.getByTitle(this.EXPENSES_LIST).items.add(itemData);
      return result.data.Id;
    } catch (error) {
      logger.error('ExpenseService', 'Error creating expense:', error);
      throw error;
    }
  }

  public async updateExpense(id: number, updates: Partial<IExpense>): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const itemData: Record<string, unknown> = {};

      if (updates.Title) itemData.Title = ValidationUtils.sanitizeHtml(updates.Title);
      if (updates.ExpenseDate) itemData.ExpenseDate = ValidationUtils.validateDate(updates.ExpenseDate, 'ExpenseDate');
      if (updates.Amount !== undefined) itemData.Amount = updates.Amount;
      if (updates.Category) {
        ValidationUtils.validateEnum(updates.Category, ExpenseCategory, 'Category');
        itemData.Category = updates.Category;
      }
      if (updates.Status) {
        ValidationUtils.validateEnum(updates.Status, ExpenseStatus, 'Status');
        itemData.Status = updates.Status;
      }
      if (updates.CostCenter !== undefined) itemData.CostCenter = updates.CostCenter ? ValidationUtils.sanitizeHtml(updates.CostCenter) : null;
      if (updates.Department) itemData.Department = ValidationUtils.sanitizeHtml(updates.Department);
      if (updates.Currency) itemData.Currency = updates.Currency;
      if (updates.Vendor !== undefined) itemData.Vendor = updates.Vendor ? ValidationUtils.sanitizeHtml(updates.Vendor) : null;
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes ? ValidationUtils.sanitizeHtml(updates.Notes) : null;
      if (updates.ReceiptUrl !== undefined) itemData.ReceiptUrl = updates.ReceiptUrl;
      if (updates.SubmittedDate) itemData.SubmittedDate = ValidationUtils.validateDate(updates.SubmittedDate, 'SubmittedDate');
      if (updates.ApprovalDate) itemData.ApprovalDate = ValidationUtils.validateDate(updates.ApprovalDate, 'ApprovalDate');
      if (updates.PaymentDate) itemData.PaymentDate = ValidationUtils.validateDate(updates.PaymentDate, 'PaymentDate');
      if (updates.PaymentReference !== undefined) itemData.PaymentReference = updates.PaymentReference;
      if (updates.RejectionReason !== undefined) itemData.RejectionReason = updates.RejectionReason ? ValidationUtils.sanitizeHtml(updates.RejectionReason) : null;

      await this.sp.web.lists.getByTitle(this.EXPENSES_LIST).items.getById(validId).update(itemData);
    } catch (error) {
      logger.error('ExpenseService', 'Error updating expense:', error);
      throw error;
    }
  }

  public async deleteExpense(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Check if expense is in Draft status
      const expense = await this.getExpenseById(validId);
      if (expense.Status !== ExpenseStatus.Draft) {
        throw new Error('Only draft expenses can be deleted');
      }

      await this.sp.web.lists.getByTitle(this.EXPENSES_LIST).items.getById(validId).delete();
    } catch (error) {
      logger.error('ExpenseService', 'Error deleting expense:', error);
      throw error;
    }
  }

  // ==================== Expense Workflow Operations ====================

  public async submitExpense(id: number, submitterId: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const expense = await this.getExpenseById(validId);

      if (expense.Status !== ExpenseStatus.Draft) {
        throw new Error('Only draft expenses can be submitted');
      }

      await this.updateExpense(validId, {
        Status: ExpenseStatus.Submitted,
        SubmittedDate: new Date(),
        SubmittedById: submitterId
      });
    } catch (error) {
      logger.error('ExpenseService', 'Error submitting expense:', error);
      throw error;
    }
  }

  public async approveExpense(id: number, approverId: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const expense = await this.getExpenseById(validId);

      if (expense.Status !== ExpenseStatus.Submitted && expense.Status !== ExpenseStatus.PendingApproval) {
        throw new Error('Expense must be submitted or pending approval to approve');
      }

      await this.updateExpense(validId, {
        Status: ExpenseStatus.Approved,
        ApprovalDate: new Date(),
        ApprovedById: approverId
      });
    } catch (error) {
      logger.error('ExpenseService', 'Error approving expense:', error);
      throw error;
    }
  }

  public async rejectExpense(id: number, approverId: number, reason: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      if (!reason) {
        throw new Error('Rejection reason is required');
      }

      const expense = await this.getExpenseById(validId);

      if (expense.Status !== ExpenseStatus.Submitted && expense.Status !== ExpenseStatus.PendingApproval) {
        throw new Error('Expense must be submitted or pending approval to reject');
      }

      await this.updateExpense(validId, {
        Status: ExpenseStatus.Rejected,
        ApprovalDate: new Date(),
        ApprovedById: approverId,
        RejectionReason: reason
      });
    } catch (error) {
      logger.error('ExpenseService', 'Error rejecting expense:', error);
      throw error;
    }
  }

  public async markExpenseAsPaid(id: number, paymentReference: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const expense = await this.getExpenseById(validId);

      if (expense.Status !== ExpenseStatus.Approved) {
        throw new Error('Only approved expenses can be marked as paid');
      }

      await this.updateExpense(validId, {
        Status: ExpenseStatus.Paid,
        PaymentDate: new Date(),
        PaymentReference: paymentReference
      });
    } catch (error) {
      logger.error('ExpenseService', 'Error marking expense as paid:', error);
      throw error;
    }
  }

  // ==================== Query Methods ====================

  public async getPendingExpenses(): Promise<IExpense[]> {
    try {
      return this.getExpenses({ status: [ExpenseStatus.Submitted, ExpenseStatus.PendingApproval] });
    } catch (error) {
      logger.error('ExpenseService', 'Error getting pending expenses:', error);
      throw error;
    }
  }

  public async getMyExpenses(userId: number): Promise<IExpense[]> {
    try {
      return this.getExpenses({ submittedById: userId });
    } catch (error) {
      logger.error('ExpenseService', 'Error getting user expenses:', error);
      throw error;
    }
  }

  public async getDepartmentExpenses(department: string): Promise<IExpense[]> {
    try {
      return this.getExpenses({ department });
    } catch (error) {
      logger.error('ExpenseService', 'Error getting department expenses:', error);
      throw error;
    }
  }

  public async getExpensesByDateRange(fromDate: Date, toDate: Date): Promise<IExpense[]> {
    try {
      return this.getExpenses({ fromDate, toDate });
    } catch (error) {
      logger.error('ExpenseService', 'Error getting expenses by date range:', error);
      throw error;
    }
  }

  // ==================== Statistics ====================

  public async getExpenseStatistics(): Promise<{
    total: number;
    totalAmount: number;
    byStatus: { [key in ExpenseStatus]?: number };
    byCategory: { [key in ExpenseCategory]?: number };
    byDepartment: { department: string; count: number; amount: number }[];
    pendingApprovalCount: number;
    pendingApprovalAmount: number;
    avgExpenseAmount: number;
  }> {
    try {
      const expenses = await this.getExpenses();

      const stats = {
        total: expenses.length,
        totalAmount: 0,
        byStatus: {} as { [key in ExpenseStatus]?: number },
        byCategory: {} as { [key in ExpenseCategory]?: number },
        byDepartment: [] as { department: string; count: number; amount: number }[],
        pendingApprovalCount: 0,
        pendingApprovalAmount: 0,
        avgExpenseAmount: 0
      };

      const departmentMap = new Map<string, { count: number; amount: number }>();

      for (const expense of expenses) {
        stats.totalAmount += expense.Amount || 0;

        // By status
        if (expense.Status) {
          stats.byStatus[expense.Status] = (stats.byStatus[expense.Status] || 0) + 1;
        }

        // By category
        if (expense.Category) {
          stats.byCategory[expense.Category] = (stats.byCategory[expense.Category] || 0) + 1;
        }

        // Pending approval
        if (expense.Status === ExpenseStatus.Submitted || expense.Status === ExpenseStatus.PendingApproval) {
          stats.pendingApprovalCount++;
          stats.pendingApprovalAmount += expense.Amount || 0;
        }

        // By department
        const existing = departmentMap.get(expense.Department);
        if (existing) {
          existing.count++;
          existing.amount += expense.Amount || 0;
        } else {
          departmentMap.set(expense.Department, {
            count: 1,
            amount: expense.Amount || 0
          });
        }
      }

      stats.avgExpenseAmount = expenses.length > 0 ? stats.totalAmount / expenses.length : 0;

      // Convert department map to array
      stats.byDepartment = Array.from(departmentMap.entries())
        .map(([department, data]) => ({
          department,
          count: data.count,
          amount: data.amount
        }))
        .sort((a, b) => b.amount - a.amount);

      return stats;
    } catch (error) {
      logger.error('ExpenseService', 'Error getting expense statistics:', error);
      throw error;
    }
  }

  // ==================== Helper Functions ====================

  private async generateExpenseCode(): Promise<string> {
    try {
      const currentYear = new Date().getFullYear();
      const prefix = `EXP-${currentYear}-`;

      const items = await this.sp.web.lists.getByTitle(this.EXPENSES_LIST).items
        .select('ExpenseCode')
        .filter(`substringof('${prefix}', ExpenseCode)`)
        .orderBy('Id', false)
        .top(1)();

      let nextNumber = 1;
      if (items.length > 0 && items[0].ExpenseCode) {
        const match = items[0].ExpenseCode.match(/EXP-\d{4}-(\d+)/);
        if (match) {
          nextNumber = parseInt(match[1], 10) + 1;
        }
      }

      return `${prefix}${nextNumber.toString().padStart(3, '0')}`;
    } catch (error) {
      logger.error('ExpenseService', 'Error generating expense code:', error);
      return `EXP-${Date.now()}`;
    }
  }

  // ==================== Mapping Functions ====================

  private mapExpenseFromSP(item: any): IExpense {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      ExpenseCode: item.ExpenseCode as string,
      ExpenseDate: item.ExpenseDate ? new Date(item.ExpenseDate as string) : new Date(),
      Amount: item.Amount as number || 0,
      Category: item.Category as ExpenseCategory,
      CostCenter: item.CostCenter as string,
      Department: item.Department as string,
      Status: item.Status as ExpenseStatus || ExpenseStatus.Draft,
      Currency: item.Currency as Currency || Currency.GBP,
      Vendor: item.Vendor as string,
      Notes: item.Notes as string,
      SubmittedDate: item.SubmittedDate ? new Date(item.SubmittedDate as string) : undefined,
      ApprovalDate: item.ApprovalDate ? new Date(item.ApprovalDate as string) : undefined,
      PaymentDate: item.PaymentDate ? new Date(item.PaymentDate as string) : undefined,
      PaymentReference: item.PaymentReference as string,
      RejectionReason: item.RejectionReason as string,
      ReceiptUrl: item.ReceiptUrl as string,
      SubmittedById: item.SubmittedById as number,
      SubmittedBy: item.SubmittedBy ? {
        Id: item.SubmittedBy.Id,
        Title: item.SubmittedBy.Title
      } : undefined,
      ApprovedById: item.ApprovedById as number,
      ApprovedBy: item.ApprovedBy ? {
        Id: item.ApprovedBy.Id,
        Title: item.ApprovedBy.Title
      } : undefined,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }
}
