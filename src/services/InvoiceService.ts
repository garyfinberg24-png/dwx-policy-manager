// @ts-nocheck
/* eslint-disable @typescript-eslint/no-explicit-any */
// TODO: Fix VendorInvoiceNumber, PaidAmount, BalanceDue type mismatches with IInvoice
// Invoice Service
// Invoice processing, 3-way matching, and payment management
// Note: Some fields may not exist in the SharePoint list - mapping handles this gracefully

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  IInvoice,
  IInvoiceLineItem,
  IInvoiceFilter,
  InvoiceStatus,
  Currency
} from '../models/IProcurement';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class InvoiceService {
  private sp: SPFI;
  private readonly INVOICES_LIST = 'PM_Invoices';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==================== Invoice CRUD Operations ====================

  public async getInvoices(filter?: IInvoiceFilter): Promise<IInvoice[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.INVOICES_LIST).items
        .select(
          'Id', 'Title', 'InvoiceNumber', 'VendorInvoiceNumber', 'VendorId',
          'PurchaseOrderId', 'InvoiceDate', 'ReceivedDate', 'DueDate',
          'SubTotal', 'TaxAmount', 'TotalAmount', 'Currency', 'Status',
          'PaidAmount', 'PaymentDate', 'PaymentReference', 'BalanceDue',
          'DisputeReason', 'Notes',
          'Created', 'Modified', 'Author/Title', 'Editor/Title'
        )
        .expand('Author', 'Editor');

      // Apply filters
      if (filter) {
        const filters: string[] = [];

        if (filter.searchTerm) {
          const term = ValidationUtils.sanitizeForOData(filter.searchTerm);
          filters.push(`(substringof('${term}', Title) or substringof('${term}', InvoiceNumber))`);
        }

        if (filter.status && filter.status.length > 0) {
          const statusFilters = filter.status.map(s =>
            ValidationUtils.buildFilter('Status', 'eq', s)
          );
          filters.push(`(${statusFilters.join(' or ')})`);
        }

        if (filter.vendorId !== undefined) {
          const validVendorId = ValidationUtils.validateInteger(filter.vendorId, 'vendorId', 1);
          filters.push(`VendorId eq ${validVendorId}`);
        }

        if (filter.purchaseOrderId !== undefined) {
          const validPOId = ValidationUtils.validateInteger(filter.purchaseOrderId, 'purchaseOrderId', 1);
          filters.push(`PurchaseOrderId eq ${validPOId}`);
        }

        if (filter.department) {
          const dept = ValidationUtils.sanitizeForOData(filter.department);
          filters.push(ValidationUtils.buildFilter('Department', 'eq', dept));
        }

        if (filter.fromDate) {
          ValidationUtils.validateDate(filter.fromDate, 'fromDate');
          filters.push(`InvoiceDate ge datetime'${filter.fromDate.toISOString()}'`);
        }

        if (filter.toDate) {
          ValidationUtils.validateDate(filter.toDate, 'toDate');
          filters.push(`InvoiceDate le datetime'${filter.toDate.toISOString()}'`);
        }

        if (filter.overdue) {
          const now = new Date().toISOString();
          filters.push(`DueDate lt datetime'${now}'`);
          filters.push(`(Status ne '${InvoiceStatus.Paid}' and Status ne '${InvoiceStatus.Cancelled}')`);
        }

        if (filter.minAmount !== undefined) {
          filters.push(`TotalAmount ge ${filter.minAmount}`);
        }

        if (filter.maxAmount !== undefined) {
          filters.push(`TotalAmount le ${filter.maxAmount}`);
        }

        if (filters.length > 0) {
          query = query.filter(filters.join(' and '));
        }
      }

      const items = await query.orderBy('InvoiceDate', false).top(5000)();
      return items.map(this.mapInvoiceFromSP);
    } catch (error) {
      logger.error('InvoiceService', 'Error getting invoices:', error);
      throw error;
    }
  }

  public async getInvoiceById(id: number): Promise<IInvoice> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const item = await this.sp.web.lists.getByTitle(this.INVOICES_LIST).items
        .getById(validId)
        .select(
          'Id', 'Title', 'InvoiceNumber', 'VendorInvoiceNumber', 'VendorId',
          'PurchaseOrderId', 'InvoiceDate', 'ReceivedDate', 'DueDate',
          'SubTotal', 'TaxAmount', 'TotalAmount', 'Currency', 'Status',
          'PaidAmount', 'PaymentDate', 'PaymentReference', 'BalanceDue',
          'DisputeReason', 'Notes',
          'Created', 'Modified', 'Author/Title', 'Editor/Title'
        )
        .expand('Author', 'Editor')();

      return this.mapInvoiceFromSP(item);
    } catch (error) {
      logger.error('InvoiceService', 'Error getting invoice by ID:', error);
      throw error;
    }
  }

  public async getInvoiceByNumber(invoiceNumber: string): Promise<IInvoice | null> {
    try {
      if (!invoiceNumber || typeof invoiceNumber !== 'string') {
        throw new Error('Invalid invoice number');
      }

      const validNumber = ValidationUtils.sanitizeForOData(invoiceNumber.substring(0, 50));
      const filter = ValidationUtils.buildFilter('InvoiceNumber', 'eq', validNumber);

      const items = await this.sp.web.lists.getByTitle(this.INVOICES_LIST).items
        .select('Id', 'InvoiceNumber')
        .filter(filter)
        .top(1)();

      if (items.length === 0) {
        return null;
      }

      return this.getInvoiceById(items[0].Id);
    } catch (error) {
      logger.error('InvoiceService', 'Error getting invoice by number:', error);
      throw error;
    }
  }

  public async createInvoice(invoice: Partial<IInvoice>): Promise<number> {
    try {
      // Validate required fields
      if (!invoice.VendorId || !invoice.TotalAmount || !invoice.InvoiceDate || !invoice.DueDate) {
        throw new Error('VendorId, TotalAmount, InvoiceDate, and DueDate are required');
      }

      // Generate invoice number
      const invoiceNumber = await this.generateInvoiceNumber();

      const itemData: Record<string, unknown> = {
        Title: invoice.Title || `Invoice from Vendor ${invoice.VendorId}`,
        InvoiceNumber: invoiceNumber,
        VendorId: ValidationUtils.validateInteger(invoice.VendorId, 'VendorId', 1),
        InvoiceDate: ValidationUtils.validateDate(invoice.InvoiceDate, 'InvoiceDate'),
        ReceivedDate: invoice.ReceivedDate ? ValidationUtils.validateDate(invoice.ReceivedDate, 'ReceivedDate') : new Date().toISOString(),
        DueDate: ValidationUtils.validateDate(invoice.DueDate, 'DueDate'),
        SubTotal: invoice.Subtotal || 0,
        TaxAmount: invoice.TaxAmount || 0,
        TotalAmount: invoice.TotalAmount,
        Currency: invoice.Currency || Currency.GBP,
        Status: invoice.Status || InvoiceStatus.Received,
        PaidAmount: 0,
        BalanceDue: invoice.TotalAmount
      };

      // Optional fields
      if (invoice.VendorInvoiceNumber) itemData.VendorInvoiceNumber = ValidationUtils.sanitizeHtml(invoice.VendorInvoiceNumber);
      if (invoice.PurchaseOrderId) itemData.PurchaseOrderId = ValidationUtils.validateInteger(invoice.PurchaseOrderId, 'PurchaseOrderId', 1);
      if (invoice.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(invoice.Notes);

      const result = await this.sp.web.lists.getByTitle(this.INVOICES_LIST).items.add(itemData);
      return result.data.Id;
    } catch (error) {
      logger.error('InvoiceService', 'Error creating invoice:', error);
      throw error;
    }
  }

  public async updateInvoice(id: number, updates: Partial<IInvoice>): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const itemData: Record<string, unknown> = {};

      if (updates.Title) itemData.Title = ValidationUtils.sanitizeHtml(updates.Title);
      if (updates.Status) {
        ValidationUtils.validateEnum(updates.Status, InvoiceStatus, 'Status');
        itemData.Status = updates.Status;
      }
      if (updates.InvoiceDate) itemData.InvoiceDate = ValidationUtils.validateDate(updates.InvoiceDate, 'InvoiceDate');
      if (updates.ReceivedDate) itemData.ReceivedDate = ValidationUtils.validateDate(updates.ReceivedDate, 'ReceivedDate');
      if (updates.DueDate) itemData.DueDate = ValidationUtils.validateDate(updates.DueDate, 'DueDate');
      if (updates.Subtotal !== undefined) itemData.SubTotal = updates.Subtotal;
      if (updates.TaxAmount !== undefined) itemData.TaxAmount = updates.TaxAmount;
      if (updates.TotalAmount !== undefined) itemData.TotalAmount = updates.TotalAmount;
      if (updates.Currency) itemData.Currency = updates.Currency;
      if (updates.PaidAmount !== undefined) itemData.PaidAmount = updates.PaidAmount;
      if (updates.PaymentDate) itemData.PaymentDate = ValidationUtils.validateDate(updates.PaymentDate, 'PaymentDate');
      if (updates.PaymentReference !== undefined) itemData.PaymentReference = updates.PaymentReference ? ValidationUtils.sanitizeHtml(updates.PaymentReference) : null;
      if (updates.BalanceDue !== undefined) itemData.BalanceDue = updates.BalanceDue;
      if (updates.DisputeReason !== undefined) itemData.DisputeReason = updates.DisputeReason ? ValidationUtils.sanitizeHtml(updates.DisputeReason) : null;
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes ? ValidationUtils.sanitizeHtml(updates.Notes) : null;

      await this.sp.web.lists.getByTitle(this.INVOICES_LIST).items.getById(validId).update(itemData);
    } catch (error) {
      logger.error('InvoiceService', 'Error updating invoice:', error);
      throw error;
    }
  }

  public async deleteInvoice(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Check if invoice is in Received status
      const invoice = await this.getInvoiceById(validId);
      if (invoice.Status !== InvoiceStatus.Received) {
        throw new Error('Only newly received invoices can be deleted');
      }

      await this.sp.web.lists.getByTitle(this.INVOICES_LIST).items.getById(validId).delete();
    } catch (error) {
      logger.error('InvoiceService', 'Error deleting invoice:', error);
      throw error;
    }
  }

  // ==================== Invoice Workflow ====================

  public async approveInvoice(id: number, approvedById: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);
      const validApprovedById = ValidationUtils.validateInteger(approvedById, 'approvedById', 1);

      const invoice = await this.getInvoiceById(validId);
      if (invoice.Status !== InvoiceStatus.PendingApproval && invoice.Status !== InvoiceStatus.Matched) {
        throw new Error('Invoice must be in Pending Approval or Matched status to approve');
      }

      await this.sp.web.lists.getByTitle(this.INVOICES_LIST).items.getById(validId).update({
        Status: InvoiceStatus.Approved,
        ApprovedById: validApprovedById,
        ApprovedDate: new Date().toISOString()
      });
    } catch (error) {
      logger.error('InvoiceService', 'Error approving invoice:', error);
      throw error;
    }
  }

  public async disputeInvoice(id: number, reason: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      if (!reason) {
        throw new Error('Dispute reason is required');
      }

      await this.updateInvoice(validId, {
        Status: InvoiceStatus.Disputed,
        DisputeReason: reason
      });
    } catch (error) {
      logger.error('InvoiceService', 'Error disputing invoice:', error);
      throw error;
    }
  }

  public async resolveDispute(id: number, resolution: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const invoice = await this.getInvoiceById(validId);
      if (invoice.Status !== InvoiceStatus.Disputed) {
        throw new Error('Invoice is not in disputed status');
      }

      await this.sp.web.lists.getByTitle(this.INVOICES_LIST).items.getById(validId).update({
        Status: InvoiceStatus.PendingApproval,
        DisputeResolvedDate: new Date().toISOString(),
        DisputeResolution: ValidationUtils.sanitizeHtml(resolution),
        Notes: invoice.Notes
          ? `${invoice.Notes}\n\nDispute Resolution: ${resolution}`
          : `Dispute Resolution: ${resolution}`
      });
    } catch (error) {
      logger.error('InvoiceService', 'Error resolving dispute:', error);
      throw error;
    }
  }

  public async schedulePayment(id: number, paymentDate: Date): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const invoice = await this.getInvoiceById(validId);
      if (invoice.Status !== InvoiceStatus.Approved) {
        throw new Error('Invoice must be approved before scheduling payment');
      }

      await this.updateInvoice(validId, {
        Status: InvoiceStatus.Scheduled,
        PaymentDate: paymentDate
      });
    } catch (error) {
      logger.error('InvoiceService', 'Error scheduling payment:', error);
      throw error;
    }
  }

  public async recordPayment(id: number, paymentReference: string, paymentDate?: Date, amount?: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const invoice = await this.getInvoiceById(validId);
      const paymentAmount = amount || invoice.TotalAmount;
      const newPaidAmount = (invoice.PaidAmount || 0) + paymentAmount;
      const newBalanceDue = invoice.TotalAmount - newPaidAmount;

      await this.updateInvoice(validId, {
        Status: newBalanceDue <= 0 ? InvoiceStatus.Paid : invoice.Status,
        PaidAmount: newPaidAmount,
        PaymentDate: paymentDate || new Date(),
        PaymentReference: paymentReference,
        BalanceDue: Math.max(0, newBalanceDue)
      });
    } catch (error) {
      logger.error('InvoiceService', 'Error recording payment:', error);
      throw error;
    }
  }

  public async cancelInvoice(id: number, reason?: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const invoice = await this.getInvoiceById(validId);
      if (invoice.Status === InvoiceStatus.Paid) {
        throw new Error('Cannot cancel a paid invoice');
      }

      const updateData: Partial<IInvoice> = {
        Status: InvoiceStatus.Cancelled
      };

      if (reason) {
        updateData.Notes = invoice.Notes
          ? `${invoice.Notes}\n\nCancellation Reason: ${reason}`
          : `Cancellation Reason: ${reason}`;
      }

      await this.updateInvoice(validId, updateData);
    } catch (error) {
      logger.error('InvoiceService', 'Error cancelling invoice:', error);
      throw error;
    }
  }

  // ==================== Query Methods ====================

  public async getOverdueInvoices(): Promise<IInvoice[]> {
    try {
      return this.getInvoices({ overdue: true });
    } catch (error) {
      logger.error('InvoiceService', 'Error getting overdue invoices:', error);
      throw error;
    }
  }

  public async getPendingInvoices(): Promise<IInvoice[]> {
    try {
      return this.getInvoices({
        status: [
          InvoiceStatus.Received,
          InvoiceStatus.PendingMatch,
          InvoiceStatus.PendingApproval,
          InvoiceStatus.Matched
        ]
      });
    } catch (error) {
      logger.error('InvoiceService', 'Error getting pending invoices:', error);
      throw error;
    }
  }

  public async getVendorInvoices(vendorId: number): Promise<IInvoice[]> {
    try {
      return this.getInvoices({ vendorId });
    } catch (error) {
      logger.error('InvoiceService', 'Error getting vendor invoices:', error);
      throw error;
    }
  }

  public async getPOInvoices(purchaseOrderId: number): Promise<IInvoice[]> {
    try {
      return this.getInvoices({ purchaseOrderId });
    } catch (error) {
      logger.error('InvoiceService', 'Error getting PO invoices:', error);
      throw error;
    }
  }

  // ==================== Statistics ====================

  public async getInvoiceStatistics(): Promise<{
    total: number;
    pending: number;
    approved: number;
    paid: number;
    overdue: number;
    disputed: number;
    totalValue: number;
    totalPaid: number;
    totalOutstanding: number;
    avgPaymentDays: number;
    byStatus: { [key: string]: number };
  }> {
    try {
      const invoices = await this.getInvoices();
      const today = new Date();

      const stats = {
        total: invoices.length,
        pending: 0,
        approved: 0,
        paid: 0,
        overdue: 0,
        disputed: 0,
        totalValue: 0,
        totalPaid: 0,
        totalOutstanding: 0,
        avgPaymentDays: 0,
        byStatus: {} as { [key: string]: number }
      };

      let totalPaymentDays = 0;
      let paidCount = 0;

      for (const invoice of invoices) {
        // Count by status
        stats.byStatus[invoice.Status] = (stats.byStatus[invoice.Status] || 0) + 1;

        // Calculate totals
        if (invoice.Status !== InvoiceStatus.Cancelled) {
          stats.totalValue += invoice.TotalAmount || 0;
          stats.totalPaid += invoice.PaidAmount || 0;
          stats.totalOutstanding += invoice.BalanceDue || 0;
        }

        // Count specific statuses
        switch (invoice.Status) {
          case InvoiceStatus.Received:
          case InvoiceStatus.PendingMatch:
          case InvoiceStatus.PendingApproval:
          case InvoiceStatus.Matched:
            stats.pending++;
            break;
          case InvoiceStatus.Approved:
          case InvoiceStatus.Scheduled:
            stats.approved++;
            break;
          case InvoiceStatus.Paid:
            stats.paid++;
            // Calculate payment days
            if (invoice.PaymentDate && invoice.InvoiceDate) {
              const daysDiff = Math.floor(
                (new Date(invoice.PaymentDate).getTime() - new Date(invoice.InvoiceDate).getTime()) /
                (1000 * 60 * 60 * 24)
              );
              totalPaymentDays += daysDiff;
              paidCount++;
            }
            break;
          case InvoiceStatus.Disputed:
            stats.disputed++;
            break;
        }

        // Check overdue
        if (invoice.DueDate && new Date(invoice.DueDate) < today &&
            invoice.Status !== InvoiceStatus.Paid && invoice.Status !== InvoiceStatus.Cancelled) {
          stats.overdue++;
        }
      }

      stats.avgPaymentDays = paidCount > 0 ? Math.round(totalPaymentDays / paidCount) : 0;

      return stats;
    } catch (error) {
      logger.error('InvoiceService', 'Error getting invoice statistics:', error);
      throw error;
    }
  }

  // ==================== Helper Functions ====================

  private async generateInvoiceNumber(): Promise<string> {
    try {
      const year = new Date().getFullYear();
      const prefix = `INV-${year}-`;

      const items = await this.sp.web.lists.getByTitle(this.INVOICES_LIST).items
        .select('InvoiceNumber')
        .filter(`substringof('${prefix}', InvoiceNumber)`)
        .orderBy('Id', false)
        .top(1)();

      let nextNumber = 1;
      if (items.length > 0 && items[0].InvoiceNumber) {
        const match = items[0].InvoiceNumber.match(/INV-\d{4}-(\d+)/);
        if (match) {
          nextNumber = parseInt(match[1], 10) + 1;
        }
      }

      return `${prefix}${nextNumber.toString().padStart(3, '0')}`;
    } catch (error) {
      logger.error('InvoiceService', 'Error generating invoice number:', error);
      return `INV-${Date.now()}`;
    }
  }

  // ==================== Mapping Functions ====================

  private mapInvoiceFromSP(item: Record<string, unknown>): IInvoice {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      InvoiceNumber: item.InvoiceNumber as string,
      VendorId: item.VendorId as number,
      VendorInvoiceNumber: item.VendorInvoiceNumber as string,
      PurchaseOrderId: item.PurchaseOrderId as number,
      InvoiceDate: item.InvoiceDate ? new Date(item.InvoiceDate as string) : new Date(),
      ReceivedDate: item.ReceivedDate ? new Date(item.ReceivedDate as string) : new Date(),
      DueDate: item.DueDate ? new Date(item.DueDate as string) : new Date(),
      Subtotal: item.SubTotal as number || 0,
      TaxAmount: item.TaxAmount as number || 0,
      TotalAmount: item.TotalAmount as number || 0,
      Currency: item.Currency as Currency || Currency.GBP,
      Status: item.Status as InvoiceStatus || InvoiceStatus.Received,
      PaidAmount: item.PaidAmount as number || 0,
      PaymentDate: item.PaymentDate ? new Date(item.PaymentDate as string) : undefined,
      PaymentReference: item.PaymentReference as string,
      BalanceDue: item.BalanceDue as number || 0,
      DisputeReason: item.DisputeReason as string,
      DisputeResolvedDate: item.DisputeResolvedDate ? new Date(item.DisputeResolvedDate as string) : undefined,
      DisputeResolution: item.DisputeResolution as string,
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }
}
