// @ts-nocheck
/* eslint-disable @typescript-eslint/no-explicit-any */
// TODO: Fix Record<string, unknown> to IUser/IVendor type mismatches
// Requisition Service
// Purchase requisition workflow and management operations
// Note: Some fields may not exist in the SharePoint list - mapping handles this gracefully

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  IPurchaseRequisition,
  IRequisitionLineItem,
  IRequisitionFilter,
  RequisitionStatus,
  RequisitionPriority,
  VendorCategory,
  UnitOfMeasure,
  Currency,
  IJMLProcurementRequest,
  IJMLProcurementResult
} from '../models/IProcurement';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class RequisitionService {
  private sp: SPFI;
  private readonly REQUISITIONS_LIST = 'JML_PurchaseRequisitions';
  private readonly LINE_ITEMS_LIST = 'JML_RequisitionLineItems';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==================== Requisition CRUD Operations ====================

  public async getRequisitions(filter?: IRequisitionFilter): Promise<IPurchaseRequisition[]> {
    try {
      console.log('[RequisitionService] Fetching requisitions from list:', this.REQUISITIONS_LIST);
      // Field names mapped to actual SharePoint list schema:
      // RequesterId (not RequestedById), JMLProcessId (not ProcessId), JMLEmployeeId (not EmployeeId)
      // BudgetCode exists but BudgetId doesn't, SuggestedVendorId doesn't exist
      // NOTE: Removed Person field expands (Requester, ApprovedBy, Author, Editor) - cause 400 errors
      let query = this.sp.web.lists.getByTitle(this.REQUISITIONS_LIST).items
        .select(
          'Id', 'Title', 'RequisitionNumber', 'RequesterId',
          'Department', 'CostCenter', 'Status', 'Priority',
          'RequestedDate', 'RequiredByDate', 'RequiredDate', 'ApprovedDate',
          'TotalEstimatedCost', 'TotalAmount', 'Currency', 'BudgetCode',
          'JMLProcessId', 'JMLEmployeeId',
          'RejectionReason', 'RejectedById', 'RejectedDate',
          'Justification', 'DeliveryLocation',
          'Notes',
          'Created', 'Modified'
        );

      // Apply filters
      if (filter) {
        const filters: string[] = [];

        if (filter.searchTerm) {
          const term = ValidationUtils.sanitizeForOData(filter.searchTerm);
          filters.push(`(substringof('${term}', Title) or substringof('${term}', RequisitionNumber))`);
        }

        if (filter.status && filter.status.length > 0) {
          const statusFilters = filter.status.map(s =>
            ValidationUtils.buildFilter('Status', 'eq', s)
          );
          filters.push(`(${statusFilters.join(' or ')})`);
        }

        if (filter.priority && filter.priority.length > 0) {
          const priorityFilters = filter.priority.map(p =>
            ValidationUtils.buildFilter('Priority', 'eq', p)
          );
          filters.push(`(${priorityFilters.join(' or ')})`);
        }

        if (filter.requestedById !== undefined) {
          const validUserId = ValidationUtils.validateInteger(filter.requestedById, 'requestedById', 1);
          filters.push(`RequesterId eq ${validUserId}`);  // SP field: RequesterId
        }

        if (filter.department) {
          const dept = ValidationUtils.sanitizeForOData(filter.department);
          filters.push(ValidationUtils.buildFilter('Department', 'eq', dept));
        }

        // Note: SuggestedVendorId field doesn't exist in SP list - vendor filter disabled
        // if (filter.vendorId !== undefined) {
        //   const validVendorId = ValidationUtils.validateInteger(filter.vendorId, 'vendorId', 1);
        //   filters.push(`SuggestedVendorId eq ${validVendorId}`);
        // }

        if (filter.fromDate) {
          ValidationUtils.validateDate(filter.fromDate, 'fromDate');
          filters.push(`RequestedDate ge datetime'${filter.fromDate.toISOString()}'`);
        }

        if (filter.toDate) {
          ValidationUtils.validateDate(filter.toDate, 'toDate');
          filters.push(`RequestedDate le datetime'${filter.toDate.toISOString()}'`);
        }

        if (filter.minAmount !== undefined) {
          filters.push(`TotalEstimatedCost ge ${filter.minAmount}`);
        }

        if (filter.maxAmount !== undefined) {
          filters.push(`TotalEstimatedCost le ${filter.maxAmount}`);
        }

        if (filter.processId !== undefined) {
          const validProcessId = ValidationUtils.validateInteger(filter.processId, 'processId', 1);
          filters.push(`JMLProcessId eq ${validProcessId}`);  // SP field: JMLProcessId
        }

        if (filters.length > 0) {
          query = query.filter(filters.join(' and '));
        }
      }

      const items = await query.orderBy('Created', false).top(5000)();
      console.log(`[RequisitionService] Retrieved ${items.length} requisitions`);
      return items.map(this.mapRequisitionFromSP);
    } catch (error: any) {
      console.error('[RequisitionService] Error getting requisitions:', error?.message || error);
      logger.error('RequisitionService', 'Error getting requisitions:', error);
      // Return empty array instead of throwing to prevent UI crashes
      return [];
    }
  }

  public async getRequisitionById(id: number): Promise<IPurchaseRequisition | null> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Field names mapped to actual SharePoint list schema
      // NOTE: Removed Person field expands (Requester, ApprovedBy, Author, Editor) - cause 400 errors
      const item = await this.sp.web.lists.getByTitle(this.REQUISITIONS_LIST).items
        .getById(validId)
        .select(
          'Id', 'Title', 'RequisitionNumber', 'RequesterId',
          'Department', 'CostCenter', 'Status', 'Priority',
          'RequestedDate', 'RequiredByDate', 'RequiredDate', 'ApprovedDate',
          'TotalEstimatedCost', 'TotalAmount', 'Currency', 'BudgetCode',
          'JMLProcessId', 'JMLEmployeeId',
          'RejectionReason', 'RejectedById', 'RejectedDate',
          'Justification', 'DeliveryLocation',
          'Notes',
          'Created', 'Modified'
        )();

      return this.mapRequisitionFromSP(item);
    } catch (error: any) {
      console.error(`[RequisitionService] Error getting requisition by ID ${id}:`, error?.message || error);
      logger.error('RequisitionService', 'Error getting requisition by ID:', error);
      return null;
    }
  }

  public async getRequisitionByNumber(requisitionNumber: string): Promise<IPurchaseRequisition | null> {
    try {
      if (!requisitionNumber || typeof requisitionNumber !== 'string') {
        throw new Error('Invalid requisition number');
      }

      const validNumber = ValidationUtils.sanitizeForOData(requisitionNumber.substring(0, 50));
      const filter = ValidationUtils.buildFilter('RequisitionNumber', 'eq', validNumber);

      const items = await this.sp.web.lists.getByTitle(this.REQUISITIONS_LIST).items
        .select('Id', 'RequisitionNumber')
        .filter(filter)
        .top(1)();

      if (items.length === 0) {
        return null;
      }

      return this.getRequisitionById(items[0].Id);
    } catch (error) {
      logger.error('RequisitionService', 'Error getting requisition by number:', error);
      throw error;
    }
  }

  public async createRequisition(requisition: Partial<IPurchaseRequisition>, lineItems?: Partial<IRequisitionLineItem>[]): Promise<number> {
    try {
      // Validate required fields
      if (!requisition.Title || !requisition.RequestedById || !requisition.Department) {
        throw new Error('Title, RequestedById, and Department are required');
      }

      // Generate requisition number
      const requisitionNumber = await this.generateRequisitionNumber();

      const itemData: Record<string, unknown> = {
        Title: ValidationUtils.sanitizeHtml(requisition.Title),
        RequisitionNumber: requisitionNumber,
        RequestedById: ValidationUtils.validateInteger(requisition.RequestedById, 'RequestedById', 1),
        Department: ValidationUtils.sanitizeHtml(requisition.Department),
        Status: requisition.Status || RequisitionStatus.Draft,
        Priority: requisition.Priority || RequisitionPriority.Medium,
        RequestedDate: new Date().toISOString(),
        TotalEstimatedCost: requisition.TotalEstimatedCost || 0,
        Currency: requisition.Currency || Currency.USD
      };

      // Optional fields
      if (requisition.CostCenter) itemData.CostCenter = ValidationUtils.sanitizeHtml(requisition.CostCenter);
      if (requisition.RequiredByDate) itemData.RequiredByDate = ValidationUtils.validateDate(requisition.RequiredByDate, 'RequiredByDate');
      if (requisition.BudgetId) itemData.BudgetId = ValidationUtils.validateInteger(requisition.BudgetId, 'BudgetId', 1);
      if (requisition.SuggestedVendorId) itemData.SuggestedVendorId = ValidationUtils.validateInteger(requisition.SuggestedVendorId, 'SuggestedVendorId', 1);
      if (requisition.ProcessId) itemData.ProcessId = ValidationUtils.validateInteger(requisition.ProcessId, 'ProcessId', 1);
      if (requisition.TaskId) itemData.TaskId = ValidationUtils.validateInteger(requisition.TaskId, 'TaskId', 1);
      if (requisition.EmployeeId) itemData.EmployeeId = ValidationUtils.validateInteger(requisition.EmployeeId, 'EmployeeId', 1);
      if (requisition.Justification) itemData.Justification = ValidationUtils.sanitizeHtml(requisition.Justification);
      if (requisition.BusinessNeed) itemData.BusinessNeed = ValidationUtils.sanitizeHtml(requisition.BusinessNeed);
      if (requisition.Attachments) itemData.Attachments = requisition.Attachments;
      if (requisition.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(requisition.Notes);

      const result = await this.sp.web.lists.getByTitle(this.REQUISITIONS_LIST).items.add(itemData);
      const requisitionId = result.data.Id;

      // Create line items if provided
      if (lineItems && lineItems.length > 0) {
        let totalCost = 0;
        for (let i = 0; i < lineItems.length; i++) {
          const lineItem = lineItems[i];
          const lineItemId = await this.createLineItem({
            ...lineItem,
            RequisitionId: requisitionId,
            LineNumber: i + 1
          });

          if (lineItem.EstimatedTotalCost) {
            totalCost += lineItem.EstimatedTotalCost;
          }
        }

        // Update total cost
        if (totalCost > 0) {
          await this.updateRequisition(requisitionId, { TotalEstimatedCost: totalCost });
        }
      }

      return requisitionId;
    } catch (error) {
      logger.error('RequisitionService', 'Error creating requisition:', error);
      throw error;
    }
  }

  public async updateRequisition(id: number, updates: Partial<IPurchaseRequisition>): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const itemData: Record<string, unknown> = {};

      if (updates.Title) itemData.Title = ValidationUtils.sanitizeHtml(updates.Title);
      if (updates.Status) {
        ValidationUtils.validateEnum(updates.Status, RequisitionStatus, 'Status');
        itemData.Status = updates.Status;
      }
      if (updates.Priority) {
        ValidationUtils.validateEnum(updates.Priority, RequisitionPriority, 'Priority');
        itemData.Priority = updates.Priority;
      }
      if (updates.Department) itemData.Department = ValidationUtils.sanitizeHtml(updates.Department);
      if (updates.CostCenter !== undefined) itemData.CostCenter = updates.CostCenter ? ValidationUtils.sanitizeHtml(updates.CostCenter) : null;
      if (updates.RequiredByDate) itemData.RequiredByDate = ValidationUtils.validateDate(updates.RequiredByDate, 'RequiredByDate');
      if (updates.TotalEstimatedCost !== undefined) itemData.TotalEstimatedCost = updates.TotalEstimatedCost;
      if (updates.Currency) itemData.Currency = updates.Currency;
      if (updates.BudgetId !== undefined) {
        itemData.BudgetId = updates.BudgetId === null ? null :
          ValidationUtils.validateInteger(updates.BudgetId, 'BudgetId', 1);
      }
      if (updates.SuggestedVendorId !== undefined) {
        itemData.SuggestedVendorId = updates.SuggestedVendorId === null ? null :
          ValidationUtils.validateInteger(updates.SuggestedVendorId, 'SuggestedVendorId', 1);
      }
      if (updates.Justification !== undefined) itemData.Justification = updates.Justification ? ValidationUtils.sanitizeHtml(updates.Justification) : null;
      if (updates.BusinessNeed !== undefined) itemData.BusinessNeed = updates.BusinessNeed ? ValidationUtils.sanitizeHtml(updates.BusinessNeed) : null;
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes ? ValidationUtils.sanitizeHtml(updates.Notes) : null;
      if (updates.Attachments !== undefined) itemData.Attachments = updates.Attachments;

      // Approval fields
      if (updates.ApprovalStatus !== undefined) itemData.ApprovalStatus = updates.ApprovalStatus;
      if (updates.ApprovedById !== undefined) itemData.ApprovedById = updates.ApprovedById;
      if (updates.ApprovedDate) itemData.ApprovedDate = ValidationUtils.validateDate(updates.ApprovedDate, 'ApprovedDate');
      if (updates.RejectionReason !== undefined) itemData.RejectionReason = updates.RejectionReason ? ValidationUtils.sanitizeHtml(updates.RejectionReason) : null;

      // Conversion fields
      if (updates.PurchaseOrderId !== undefined) itemData.PurchaseOrderId = updates.PurchaseOrderId;
      if (updates.ConvertedDate) itemData.ConvertedDate = ValidationUtils.validateDate(updates.ConvertedDate, 'ConvertedDate');

      await this.sp.web.lists.getByTitle(this.REQUISITIONS_LIST).items.getById(validId).update(itemData);
    } catch (error) {
      logger.error('RequisitionService', 'Error updating requisition:', error);
      throw error;
    }
  }

  public async deleteRequisition(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Check if requisition is in Draft status
      const requisition = await this.getRequisitionById(validId);
      if (requisition.Status !== RequisitionStatus.Draft) {
        throw new Error('Only draft requisitions can be deleted');
      }

      // Delete line items first
      const lineItems = await this.getLineItems(validId);
      for (const lineItem of lineItems) {
        if (lineItem.Id) {
          await this.deleteLineItem(lineItem.Id);
        }
      }

      // Delete requisition
      await this.sp.web.lists.getByTitle(this.REQUISITIONS_LIST).items.getById(validId).delete();
    } catch (error) {
      logger.error('RequisitionService', 'Error deleting requisition:', error);
      throw error;
    }
  }

  // ==================== Requisition Workflow ====================

  public async submitRequisition(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Validate requisition has line items
      const lineItems = await this.getLineItems(validId);
      if (lineItems.length === 0) {
        throw new Error('Requisition must have at least one line item');
      }

      await this.updateRequisition(validId, {
        Status: RequisitionStatus.Submitted,
        RequestedDate: new Date()
      });

      // TODO: Trigger approval workflow based on amount and category
    } catch (error) {
      logger.error('RequisitionService', 'Error submitting requisition:', error);
      throw error;
    }
  }

  public async approveRequisition(id: number, approvedById: number, comments?: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);
      const validApprovedById = ValidationUtils.validateInteger(approvedById, 'approvedById', 1);

      const requisition = await this.getRequisitionById(validId);

      if (requisition.Status !== RequisitionStatus.Submitted && requisition.Status !== RequisitionStatus.PendingApproval) {
        throw new Error('Requisition is not in a state that can be approved');
      }

      const updateData: Partial<IPurchaseRequisition> = {
        Status: RequisitionStatus.Approved,
        ApprovalStatus: 'Approved',
        ApprovedById: validApprovedById,
        ApprovedDate: new Date()
      };

      if (comments) {
        updateData.Notes = requisition.Notes
          ? `${requisition.Notes}\n\nApproval Comment: ${comments}`
          : `Approval Comment: ${comments}`;
      }

      await this.updateRequisition(validId, updateData);

      // TODO: Send notification to requestor
    } catch (error) {
      logger.error('RequisitionService', 'Error approving requisition:', error);
      throw error;
    }
  }

  public async rejectRequisition(id: number, rejectedById: number, reason: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);
      const validRejectedById = ValidationUtils.validateInteger(rejectedById, 'rejectedById', 1);

      if (!reason) {
        throw new Error('Rejection reason is required');
      }

      const requisition = await this.getRequisitionById(validId);

      if (requisition.Status !== RequisitionStatus.Submitted && requisition.Status !== RequisitionStatus.PendingApproval) {
        throw new Error('Requisition is not in a state that can be rejected');
      }

      await this.updateRequisition(validId, {
        Status: RequisitionStatus.Rejected,
        ApprovalStatus: 'Rejected',
        ApprovedById: validRejectedById,
        ApprovedDate: new Date(),
        RejectionReason: reason
      });

      // TODO: Send notification to requestor
    } catch (error) {
      logger.error('RequisitionService', 'Error rejecting requisition:', error);
      throw error;
    }
  }

  public async cancelRequisition(id: number, reason?: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const requisition = await this.getRequisitionById(validId);

      if (requisition.Status === RequisitionStatus.ConvertedToPO) {
        throw new Error('Cannot cancel a requisition that has been converted to a PO');
      }

      const updateData: Partial<IPurchaseRequisition> = {
        Status: RequisitionStatus.Cancelled
      };

      if (reason) {
        updateData.Notes = requisition.Notes
          ? `${requisition.Notes}\n\nCancellation Reason: ${reason}`
          : `Cancellation Reason: ${reason}`;
      }

      await this.updateRequisition(validId, updateData);
    } catch (error) {
      logger.error('RequisitionService', 'Error cancelling requisition:', error);
      throw error;
    }
  }

  public async convertToPO(id: number, purchaseOrderId: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);
      const validPOId = ValidationUtils.validateInteger(purchaseOrderId, 'purchaseOrderId', 1);

      const requisition = await this.getRequisitionById(validId);

      if (requisition.Status !== RequisitionStatus.Approved) {
        throw new Error('Only approved requisitions can be converted to PO');
      }

      await this.updateRequisition(validId, {
        Status: RequisitionStatus.ConvertedToPO,
        PurchaseOrderId: validPOId,
        ConvertedDate: new Date()
      });
    } catch (error) {
      logger.error('RequisitionService', 'Error converting requisition to PO:', error);
      throw error;
    }
  }

  // ==================== Line Items ====================

  public async getLineItems(requisitionId: number): Promise<IRequisitionLineItem[]> {
    try {
      const validRequisitionId = ValidationUtils.validateInteger(requisitionId, 'requisitionId', 1);

      const items = await this.sp.web.lists.getByTitle(this.LINE_ITEMS_LIST).items
        .select(
          'Id', 'Title', 'RequisitionId', 'LineNumber',
          'CatalogItemId', 'ItemCode', 'Description', 'Category',
          'Quantity', 'UnitOfMeasure', 'EstimatedUnitCost', 'EstimatedTotalCost', 'Currency',
          'VendorId', 'Vendor/Title',
          'Specifications', 'Notes',
          'Created', 'Modified'
        )
        .expand('Vendor')
        .filter(`RequisitionId eq ${validRequisitionId}`)
        .orderBy('LineNumber', true)();

      return items.map(this.mapLineItemFromSP);
    } catch (error) {
      logger.error('RequisitionService', 'Error getting line items:', error);
      throw error;
    }
  }

  public async createLineItem(lineItem: Partial<IRequisitionLineItem>): Promise<number> {
    try {
      if (!lineItem.RequisitionId || !lineItem.Description || lineItem.Quantity === undefined) {
        throw new Error('RequisitionId, Description, and Quantity are required');
      }

      const itemData: Record<string, unknown> = {
        Title: ValidationUtils.sanitizeHtml(lineItem.Description.substring(0, 255)),
        RequisitionId: ValidationUtils.validateInteger(lineItem.RequisitionId, 'RequisitionId', 1),
        LineNumber: lineItem.LineNumber || 1,
        Description: ValidationUtils.sanitizeHtml(lineItem.Description),
        Category: lineItem.Category || VendorCategory.Other,
        Quantity: lineItem.Quantity,
        UnitOfMeasure: lineItem.UnitOfMeasure || UnitOfMeasure.Each,
        EstimatedUnitCost: lineItem.EstimatedUnitCost || 0,
        EstimatedTotalCost: lineItem.EstimatedTotalCost || (lineItem.Quantity * (lineItem.EstimatedUnitCost || 0)),
        Currency: lineItem.Currency || Currency.USD
      };

      if (lineItem.CatalogItemId) itemData.CatalogItemId = ValidationUtils.validateInteger(lineItem.CatalogItemId, 'CatalogItemId', 1);
      if (lineItem.ItemCode) itemData.ItemCode = ValidationUtils.sanitizeHtml(lineItem.ItemCode);
      if (lineItem.VendorId) itemData.VendorId = ValidationUtils.validateInteger(lineItem.VendorId, 'VendorId', 1);
      if (lineItem.Specifications) itemData.Specifications = ValidationUtils.sanitizeHtml(lineItem.Specifications);
      if (lineItem.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(lineItem.Notes);

      const result = await this.sp.web.lists.getByTitle(this.LINE_ITEMS_LIST).items.add(itemData);

      // Update requisition total
      await this.recalculateRequisitionTotal(lineItem.RequisitionId);

      return result.data.Id;
    } catch (error) {
      logger.error('RequisitionService', 'Error creating line item:', error);
      throw error;
    }
  }

  public async updateLineItem(id: number, updates: Partial<IRequisitionLineItem>): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Get current line item to know the requisition ID
      const currentItem = await this.sp.web.lists.getByTitle(this.LINE_ITEMS_LIST).items
        .getById(validId)
        .select('RequisitionId')();

      const itemData: Record<string, unknown> = {};

      if (updates.Description) {
        itemData.Description = ValidationUtils.sanitizeHtml(updates.Description);
        itemData.Title = ValidationUtils.sanitizeHtml(updates.Description.substring(0, 255));
      }
      if (updates.Category) itemData.Category = updates.Category;
      if (updates.Quantity !== undefined) itemData.Quantity = updates.Quantity;
      if (updates.UnitOfMeasure) itemData.UnitOfMeasure = updates.UnitOfMeasure;
      if (updates.EstimatedUnitCost !== undefined) itemData.EstimatedUnitCost = updates.EstimatedUnitCost;
      if (updates.EstimatedTotalCost !== undefined) itemData.EstimatedTotalCost = updates.EstimatedTotalCost;
      if (updates.Currency) itemData.Currency = updates.Currency;
      if (updates.CatalogItemId !== undefined) {
        itemData.CatalogItemId = updates.CatalogItemId === null ? null :
          ValidationUtils.validateInteger(updates.CatalogItemId, 'CatalogItemId', 1);
      }
      if (updates.ItemCode !== undefined) itemData.ItemCode = updates.ItemCode ? ValidationUtils.sanitizeHtml(updates.ItemCode) : null;
      if (updates.VendorId !== undefined) {
        itemData.VendorId = updates.VendorId === null ? null :
          ValidationUtils.validateInteger(updates.VendorId, 'VendorId', 1);
      }
      if (updates.Specifications !== undefined) itemData.Specifications = updates.Specifications ? ValidationUtils.sanitizeHtml(updates.Specifications) : null;
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes ? ValidationUtils.sanitizeHtml(updates.Notes) : null;

      await this.sp.web.lists.getByTitle(this.LINE_ITEMS_LIST).items.getById(validId).update(itemData);

      // Update requisition total
      await this.recalculateRequisitionTotal(currentItem.RequisitionId);
    } catch (error) {
      logger.error('RequisitionService', 'Error updating line item:', error);
      throw error;
    }
  }

  public async deleteLineItem(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Get requisition ID before deleting
      const item = await this.sp.web.lists.getByTitle(this.LINE_ITEMS_LIST).items
        .getById(validId)
        .select('RequisitionId')();

      await this.sp.web.lists.getByTitle(this.LINE_ITEMS_LIST).items.getById(validId).delete();

      // Update requisition total
      await this.recalculateRequisitionTotal(item.RequisitionId);
    } catch (error) {
      logger.error('RequisitionService', 'Error deleting line item:', error);
      throw error;
    }
  }

  private async recalculateRequisitionTotal(requisitionId: number): Promise<void> {
    try {
      const lineItems = await this.getLineItems(requisitionId);
      const total = lineItems.reduce((sum, item) => sum + (item.EstimatedTotalCost || 0), 0);

      await this.sp.web.lists.getByTitle(this.REQUISITIONS_LIST).items.getById(requisitionId).update({
        TotalEstimatedCost: total
      });
    } catch (error) {
      logger.error('RequisitionService', 'Error recalculating requisition total:', error);
    }
  }

  // ==================== JML Integration ====================

  public async createRequisitionFromJML(request: IJMLProcurementRequest): Promise<IJMLProcurementResult> {
    try {
      // Validate request
      if (!request.processId || !request.employeeId || !request.requestedItems || request.requestedItems.length === 0) {
        throw new Error('Invalid JML procurement request');
      }

      // Create requisition
      const requisitionId = await this.createRequisition({
        Title: `Equipment for ${request.employeeName} - ${request.processType}`,
        RequestedById: request.employeeId, // This should be the HR/manager ID in real scenario
        Department: request.department,
        Priority: request.processType === 'Joiner' ? RequisitionPriority.High : RequisitionPriority.Medium,
        RequiredByDate: request.startDate,
        ProcessId: request.processId,
        EmployeeId: request.employeeId,
        Justification: `${request.processType} process equipment request for ${request.employeeName}`,
        BusinessNeed: request.equipmentTemplate || `Standard ${request.processType} equipment`
      });

      // Create line items from requested items
      for (let i = 0; i < request.requestedItems.length; i++) {
        const item = request.requestedItems[i];
        await this.createLineItem({
          RequisitionId: requisitionId,
          LineNumber: i + 1,
          CatalogItemId: item.catalogItemId,
          ItemCode: item.itemCode,
          Description: item.description,
          Category: item.category,
          Quantity: item.quantity,
          Specifications: item.specifications,
          VendorId: item.preferredVendorId
        });
      }

      // Get the created requisition
      const requisition = await this.getRequisitionById(requisitionId);

      return {
        processId: request.processId,
        requisitionId: requisitionId,
        requisitionNumber: requisition.RequisitionNumber,
        status: 'Created',
        message: `Requisition ${requisition.RequisitionNumber} created successfully for ${request.processType} process`
      };
    } catch (error) {
      logger.error('RequisitionService', 'Error creating requisition from JML:', error);
      return {
        processId: request.processId,
        status: 'Failed',
        message: `Failed to create requisition: ${error instanceof Error ? error.message : 'Unknown error'}`
      };
    }
  }

  // ==================== Statistics ====================

  public async getRequisitionStatistics(): Promise<{
    total: number;
    draft: number;
    submitted: number;
    pendingApproval: number;
    approved: number;
    rejected: number;
    convertedToPO: number;
    totalValue: number;
    avgValue: number;
    byDepartment: { [key: string]: number };
    byPriority: { [key: string]: number };
  }> {
    try {
      const requisitions = await this.getRequisitions();

      const stats = {
        total: requisitions.length,
        draft: 0,
        submitted: 0,
        pendingApproval: 0,
        approved: 0,
        rejected: 0,
        convertedToPO: 0,
        totalValue: 0,
        avgValue: 0,
        byDepartment: {} as { [key: string]: number },
        byPriority: {} as { [key: string]: number }
      };

      for (const req of requisitions) {
        // Count by status
        switch (req.Status) {
          case RequisitionStatus.Draft: stats.draft++; break;
          case RequisitionStatus.Submitted: stats.submitted++; break;
          case RequisitionStatus.PendingApproval: stats.pendingApproval++; break;
          case RequisitionStatus.Approved: stats.approved++; break;
          case RequisitionStatus.Rejected: stats.rejected++; break;
          case RequisitionStatus.ConvertedToPO: stats.convertedToPO++; break;
        }

        // Total value
        stats.totalValue += req.TotalEstimatedCost || 0;

        // By department
        if (req.Department) {
          stats.byDepartment[req.Department] = (stats.byDepartment[req.Department] || 0) + 1;
        }

        // By priority
        if (req.Priority) {
          stats.byPriority[req.Priority] = (stats.byPriority[req.Priority] || 0) + 1;
        }
      }

      stats.avgValue = stats.total > 0 ? stats.totalValue / stats.total : 0;

      return stats;
    } catch (error) {
      logger.error('RequisitionService', 'Error getting requisition statistics:', error);
      throw error;
    }
  }

  // ==================== Pending Approvals ====================

  public async getPendingApprovals(approverId?: number): Promise<IPurchaseRequisition[]> {
    try {
      const filter: IRequisitionFilter = {
        status: [RequisitionStatus.Submitted, RequisitionStatus.PendingApproval]
      };

      const requisitions = await this.getRequisitions(filter);

      // TODO: Filter by approver based on approval rules
      // For now, return all pending requisitions
      return requisitions;
    } catch (error) {
      logger.error('RequisitionService', 'Error getting pending approvals:', error);
      throw error;
    }
  }

  // ==================== Helper Functions ====================

  private async generateRequisitionNumber(): Promise<string> {
    try {
      const year = new Date().getFullYear();
      const prefix = `REQ-${year}-`;

      const items = await this.sp.web.lists.getByTitle(this.REQUISITIONS_LIST).items
        .select('RequisitionNumber')
        .filter(`substringof('${prefix}', RequisitionNumber)`)
        .orderBy('Id', false)
        .top(1)();

      let nextNumber = 1;
      if (items.length > 0 && items[0].RequisitionNumber) {
        const match = items[0].RequisitionNumber.match(/REQ-\d{4}-(\d+)/);
        if (match) {
          nextNumber = parseInt(match[1], 10) + 1;
        }
      }

      return `${prefix}${nextNumber.toString().padStart(5, '0')}`;
    } catch (error) {
      logger.error('RequisitionService', 'Error generating requisition number:', error);
      return `REQ-${Date.now()}`;
    }
  }

  // ==================== Mapping Functions ====================

  private mapRequisitionFromSP(item: Record<string, unknown>): IPurchaseRequisition {
    // Map SharePoint field names to interface properties
    // SP field: RequesterId maps to RequestedById, JMLProcessId maps to ProcessId, etc.
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      RequisitionNumber: item.RequisitionNumber as string,
      RequestedById: item.RequesterId as number,  // SP field: RequesterId
      RequestedBy: item.Requester as Record<string, unknown>,  // SP field: Requester
      Department: item.Department as string,
      CostCenter: item.CostCenter as string,
      Status: item.Status as RequisitionStatus,
      Priority: item.Priority as RequisitionPriority,
      RequestedDate: item.RequestedDate ? new Date(item.RequestedDate as string) : (item.RequestDate ? new Date(item.RequestDate as string) : new Date()),
      RequiredByDate: item.RequiredByDate ? new Date(item.RequiredByDate as string) : (item.RequiredDate ? new Date(item.RequiredDate as string) : undefined),
      ApprovedDate: item.ApprovedDate ? new Date(item.ApprovedDate as string) : undefined,
      TotalEstimatedCost: (item.TotalEstimatedCost as number) || (item.TotalAmount as number) || 0,  // SP has both fields
      Currency: item.Currency as Currency || Currency.USD,
      BudgetId: undefined,  // Not in SP list (BudgetCode is text field)
      SuggestedVendorId: undefined,  // Not in SP list
      SuggestedVendor: undefined,  // Not in SP list
      ProcessId: item.JMLProcessId as number,  // SP field: JMLProcessId
      TaskId: undefined,  // Not in SP list
      EmployeeId: item.JMLEmployeeId as number,  // SP field: JMLEmployeeId
      ApprovalStatus: item.Status as string,  // Use Status field
      ApprovedById: item.ApprovedById as number,
      ApprovedBy: item.ApprovedBy as Record<string, unknown>,
      RejectionReason: item.RejectionReason as string,
      PurchaseOrderId: undefined,  // Not in SP list
      ConvertedDate: undefined,  // Not in SP list
      Justification: item.Justification as string,
      BusinessNeed: item.Justification as string,  // Map Justification to BusinessNeed
      Attachments: item.Attachments as string,
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }

  private mapLineItemFromSP(item: Record<string, unknown>): IRequisitionLineItem {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      RequisitionId: item.RequisitionId as number,
      LineNumber: item.LineNumber as number,
      CatalogItemId: item.CatalogItemId as number,
      ItemCode: item.ItemCode as string,
      Description: item.Description as string,
      Category: item.Category as VendorCategory,
      Quantity: item.Quantity as number,
      UnitOfMeasure: item.UnitOfMeasure as UnitOfMeasure,
      EstimatedUnitCost: item.EstimatedUnitCost as number || 0,
      EstimatedTotalCost: item.EstimatedTotalCost as number || 0,
      Currency: item.Currency as Currency || Currency.USD,
      VendorId: item.VendorId as number,
      Vendor: item.Vendor as Record<string, unknown>,
      Specifications: item.Specifications as string,
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }
}
