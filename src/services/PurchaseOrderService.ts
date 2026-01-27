// @ts-nocheck
/* eslint-disable @typescript-eslint/no-explicit-any */
// TODO: Fix Record<string, unknown> to IVendor/IUser type mismatches
// Purchase Order Service
// PO lifecycle management, goods receipt, and asset integration
// Note: Some fields may not exist in the SharePoint list - mapping handles this gracefully

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  IPurchaseOrder,
  IPOLineItem,
  IGoodsReceipt,
  IReceiptLineItem,
  IPurchaseOrderFilter,
  IPurchaseRequisition,
  IRequisitionLineItem,
  POStatus,
  ReceiptStatus,
  PaymentTerms,
  Currency,
  UnitOfMeasure
} from '../models/IProcurement';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';
import { RequisitionService } from './RequisitionService';

export class PurchaseOrderService {
  private sp: SPFI;
  private requisitionService: RequisitionService;
  private readonly PO_LIST = 'JML_PurchaseOrders';
  private readonly PO_LINE_ITEMS_LIST = 'JML_POLineItems';
  private readonly GOODS_RECEIPTS_LIST = 'JML_GoodsReceipts';
  private readonly RECEIPT_LINE_ITEMS_LIST = 'JML_ReceiptLineItems';

  constructor(sp: SPFI) {
    this.sp = sp;
    this.requisitionService = new RequisitionService(sp);
  }

  // ==================== Purchase Order CRUD ====================

  public async getPurchaseOrders(filter?: IPurchaseOrderFilter): Promise<IPurchaseOrder[]> {
    try {
      console.log('[PurchaseOrderService] Fetching purchase orders from list:', this.PO_LIST);
      // NOTE: Removed Person field expands (ApprovedBy, SentBy, Author, Editor) - cause 400 errors
      let query = this.sp.web.lists.getByTitle(this.PO_LIST).items
        .select(
          'Id', 'Title', 'PONumber', 'VendorId', 'Status',
          'OrderDate', 'ExpectedDeliveryDate', 'ActualDeliveryDate',
          'SentDate', 'AcknowledgedDate', 'ClosedDate',
          'ShipToAddress', 'ShipToAttention', 'BillToAddress',
          'TaxRate', 'TaxAmount', 'ShippingCost', 'DiscountAmount', 'TotalAmount',
          'Currency', 'PaymentTerms', 'VendorQuoteNumber', 'VendorReference',
          'ApprovedDate',
          'TermsAndConditions', 'SpecialInstructions',
          'BudgetId', 'CostCenter', 'Department', 'ProcessId',
          'Notes',
          'Created', 'Modified'
        );

      // Apply filters
      if (filter) {
        const filters: string[] = [];

        if (filter.searchTerm) {
          const term = ValidationUtils.sanitizeForOData(filter.searchTerm);
          filters.push(`(substringof('${term}', Title) or substringof('${term}', PONumber))`);
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

        if (filter.department) {
          const dept = ValidationUtils.sanitizeForOData(filter.department);
          filters.push(ValidationUtils.buildFilter('Department', 'eq', dept));
        }

        if (filter.fromDate) {
          ValidationUtils.validateDate(filter.fromDate, 'fromDate');
          filters.push(`OrderDate ge datetime'${filter.fromDate.toISOString()}'`);
        }

        if (filter.toDate) {
          ValidationUtils.validateDate(filter.toDate, 'toDate');
          filters.push(`OrderDate le datetime'${filter.toDate.toISOString()}'`);
        }

        if (filter.minAmount !== undefined) {
          filters.push(`TotalAmount ge ${filter.minAmount}`);
        }

        if (filter.maxAmount !== undefined) {
          filters.push(`TotalAmount le ${filter.maxAmount}`);
        }

        if (filter.overdue) {
          const now = new Date().toISOString();
          filters.push(`ExpectedDeliveryDate lt datetime'${now}'`);
          filters.push(`(Status eq '${POStatus.Sent}' or Status eq '${POStatus.Acknowledged}' or Status eq '${POStatus.PartiallyReceived}')`);
        }

        if (filters.length > 0) {
          query = query.filter(filters.join(' and '));
        }
      }

      const items = await query.orderBy('Created', false).top(5000)();
      console.log(`[PurchaseOrderService] Retrieved ${items.length} purchase orders`);
      return items.map(this.mapPOFromSP);
    } catch (error: any) {
      console.error('[PurchaseOrderService] Error getting purchase orders:', error?.message || error);
      logger.error('PurchaseOrderService', 'Error getting purchase orders:', error);
      // Return empty array instead of throwing to prevent UI crashes
      return [];
    }
  }

  public async getPurchaseOrderById(id: number): Promise<IPurchaseOrder | null> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // NOTE: Removed Person field expands (ApprovedBy, SentBy, Author, Editor) - cause 400 errors
      const item = await this.sp.web.lists.getByTitle(this.PO_LIST).items
        .getById(validId)
        .select(
          'Id', 'Title', 'PONumber', 'VendorId', 'Status',
          'OrderDate', 'ExpectedDeliveryDate', 'ActualDeliveryDate',
          'SentDate', 'AcknowledgedDate', 'ClosedDate',
          'ShipToAddress', 'ShipToAttention', 'BillToAddress',
          'TaxRate', 'TaxAmount', 'ShippingCost', 'DiscountAmount', 'TotalAmount',
          'Currency', 'PaymentTerms', 'VendorQuoteNumber', 'VendorReference',
          'ApprovedDate',
          'TermsAndConditions', 'SpecialInstructions',
          'BudgetId', 'CostCenter', 'Department', 'ProcessId',
          'Notes',
          'Created', 'Modified'
        )();

      return this.mapPOFromSP(item);
    } catch (error: any) {
      console.error(`[PurchaseOrderService] Error getting PO by ID ${id}:`, error?.message || error);
      logger.error('PurchaseOrderService', 'Error getting purchase order by ID:', error);
      return null;
    }
  }

  public async getPurchaseOrderByNumber(poNumber: string): Promise<IPurchaseOrder | null> {
    try {
      if (!poNumber || typeof poNumber !== 'string') {
        throw new Error('Invalid PO number');
      }

      const validNumber = ValidationUtils.sanitizeForOData(poNumber.substring(0, 50));
      const filter = ValidationUtils.buildFilter('PONumber', 'eq', validNumber);

      const items = await this.sp.web.lists.getByTitle(this.PO_LIST).items
        .select('Id', 'PONumber')
        .filter(filter)
        .top(1)();

      if (items.length === 0) {
        return null;
      }

      return this.getPurchaseOrderById(items[0].Id);
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error getting PO by number:', error);
      throw error;
    }
  }

  public async createPurchaseOrder(po: Partial<IPurchaseOrder>, lineItems?: Partial<IPOLineItem>[]): Promise<number> {
    try {
      // Validate required fields
      if (!po.Title || !po.VendorId) {
        throw new Error('Title and VendorId are required');
      }

      // Generate PO number
      const poNumber = await this.generatePONumber();

      const itemData: Record<string, unknown> = {
        Title: ValidationUtils.sanitizeHtml(po.Title),
        PONumber: poNumber,
        VendorId: ValidationUtils.validateInteger(po.VendorId, 'VendorId', 1),
        Status: po.Status || POStatus.Draft,
        OrderDate: new Date().toISOString(),
        Subtotal: po.Subtotal || 0,
        TaxAmount: po.TaxAmount || 0,
        ShippingCost: po.ShippingCost || 0,
        DiscountAmount: po.DiscountAmount || 0,
        TotalAmount: po.TotalAmount || 0,
        Currency: po.Currency || Currency.USD,
        PaymentTerms: po.PaymentTerms || PaymentTerms.Net30
      };

      // Optional fields
      if (po.RequisitionIds) itemData.RequisitionIds = po.RequisitionIds;
      if (po.ExpectedDeliveryDate) itemData.ExpectedDeliveryDate = ValidationUtils.validateDate(po.ExpectedDeliveryDate, 'ExpectedDeliveryDate');
      if (po.ShipToAddress) itemData.ShipToAddress = ValidationUtils.sanitizeHtml(po.ShipToAddress);
      if (po.ShipToAttention) itemData.ShipToAttention = ValidationUtils.sanitizeHtml(po.ShipToAttention);
      if (po.BillToAddress) itemData.BillToAddress = ValidationUtils.sanitizeHtml(po.BillToAddress);
      if (po.BillToAttention) itemData.BillToAttention = ValidationUtils.sanitizeHtml(po.BillToAttention);
      if (po.TaxRate !== undefined) itemData.TaxRate = po.TaxRate;
      if (po.VendorQuoteNumber) itemData.VendorQuoteNumber = ValidationUtils.sanitizeHtml(po.VendorQuoteNumber);
      if (po.TermsAndConditions) itemData.TermsAndConditions = ValidationUtils.sanitizeHtml(po.TermsAndConditions);
      if (po.SpecialInstructions) itemData.SpecialInstructions = ValidationUtils.sanitizeHtml(po.SpecialInstructions);
      if (po.BudgetId) itemData.BudgetId = ValidationUtils.validateInteger(po.BudgetId, 'BudgetId', 1);
      if (po.CostCenter) itemData.CostCenter = ValidationUtils.sanitizeHtml(po.CostCenter);
      if (po.Department) itemData.Department = ValidationUtils.sanitizeHtml(po.Department);
      if (po.ProcessId) itemData.ProcessId = ValidationUtils.validateInteger(po.ProcessId, 'ProcessId', 1);
      if (po.TaskId) itemData.TaskId = ValidationUtils.validateInteger(po.TaskId, 'TaskId', 1);
      if (po.Attachments) itemData.Attachments = po.Attachments;
      if (po.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(po.Notes);

      const result = await this.sp.web.lists.getByTitle(this.PO_LIST).items.add(itemData);
      const poId = result.data.Id;

      // Create line items if provided
      if (lineItems && lineItems.length > 0) {
        for (let i = 0; i < lineItems.length; i++) {
          await this.createPOLineItem({
            ...lineItems[i],
            PurchaseOrderId: poId,
            LineNumber: i + 1
          });
        }

        // Recalculate totals
        await this.recalculatePOTotal(poId);
      }

      return poId;
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error creating purchase order:', error);
      throw error;
    }
  }

  public async createPOFromRequisition(requisitionId: number): Promise<number> {
    try {
      const validRequisitionId = ValidationUtils.validateInteger(requisitionId, 'requisitionId', 1);

      // Get requisition
      const requisition = await this.requisitionService.getRequisitionById(validRequisitionId);

      if (requisition.Status !== 'Approved') {
        throw new Error('Only approved requisitions can be converted to PO');
      }

      if (!requisition.SuggestedVendorId) {
        throw new Error('Requisition must have a suggested vendor');
      }

      // Get requisition line items
      const reqLineItems = await this.requisitionService.getLineItems(validRequisitionId);

      // Create PO
      const poId = await this.createPurchaseOrder({
        Title: requisition.Title,
        VendorId: requisition.SuggestedVendorId,
        RequisitionIds: JSON.stringify([validRequisitionId]),
        Department: requisition.Department,
        CostCenter: requisition.CostCenter,
        BudgetId: requisition.BudgetId,
        ProcessId: requisition.ProcessId,
        TaskId: requisition.TaskId,
        Currency: requisition.Currency,
        Notes: `Created from Requisition ${requisition.RequisitionNumber}`
      });

      // Create PO line items from requisition line items
      for (const reqItem of reqLineItems) {
        await this.createPOLineItem({
          PurchaseOrderId: poId,
          LineNumber: reqItem.LineNumber,
          CatalogItemId: reqItem.CatalogItemId,
          ItemCode: reqItem.ItemCode,
          Description: reqItem.Description,
          Quantity: reqItem.Quantity,
          UnitOfMeasure: reqItem.UnitOfMeasure,
          UnitPrice: reqItem.EstimatedUnitCost,
          TotalPrice: reqItem.EstimatedTotalCost,
          Specifications: reqItem.Specifications,
          Notes: reqItem.Notes
        });
      }

      // Update requisition to mark as converted
      await this.requisitionService.convertToPO(validRequisitionId, poId);

      // Recalculate totals
      await this.recalculatePOTotal(poId);

      return poId;
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error creating PO from requisition:', error);
      throw error;
    }
  }

  public async updatePurchaseOrder(id: number, updates: Partial<IPurchaseOrder>): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const itemData: Record<string, unknown> = {};

      if (updates.Title) itemData.Title = ValidationUtils.sanitizeHtml(updates.Title);
      if (updates.Status) {
        ValidationUtils.validateEnum(updates.Status, POStatus, 'Status');
        itemData.Status = updates.Status;
      }
      if (updates.ExpectedDeliveryDate) itemData.ExpectedDeliveryDate = ValidationUtils.validateDate(updates.ExpectedDeliveryDate, 'ExpectedDeliveryDate');
      if (updates.ActualDeliveryDate) itemData.ActualDeliveryDate = ValidationUtils.validateDate(updates.ActualDeliveryDate, 'ActualDeliveryDate');
      if (updates.ShipToAddress !== undefined) itemData.ShipToAddress = updates.ShipToAddress ? ValidationUtils.sanitizeHtml(updates.ShipToAddress) : null;
      if (updates.ShipToAttention !== undefined) itemData.ShipToAttention = updates.ShipToAttention ? ValidationUtils.sanitizeHtml(updates.ShipToAttention) : null;
      if (updates.BillToAddress !== undefined) itemData.BillToAddress = updates.BillToAddress ? ValidationUtils.sanitizeHtml(updates.BillToAddress) : null;
      if (updates.BillToAttention !== undefined) itemData.BillToAttention = updates.BillToAttention ? ValidationUtils.sanitizeHtml(updates.BillToAttention) : null;
      if (updates.Subtotal !== undefined) itemData.Subtotal = updates.Subtotal;
      if (updates.TaxRate !== undefined) itemData.TaxRate = updates.TaxRate;
      if (updates.TaxAmount !== undefined) itemData.TaxAmount = updates.TaxAmount;
      if (updates.ShippingCost !== undefined) itemData.ShippingCost = updates.ShippingCost;
      if (updates.DiscountAmount !== undefined) itemData.DiscountAmount = updates.DiscountAmount;
      if (updates.TotalAmount !== undefined) itemData.TotalAmount = updates.TotalAmount;
      if (updates.PaymentTerms) itemData.PaymentTerms = updates.PaymentTerms;
      if (updates.VendorQuoteNumber !== undefined) itemData.VendorQuoteNumber = updates.VendorQuoteNumber ? ValidationUtils.sanitizeHtml(updates.VendorQuoteNumber) : null;
      if (updates.VendorReference !== undefined) itemData.VendorReference = updates.VendorReference ? ValidationUtils.sanitizeHtml(updates.VendorReference) : null;
      if (updates.TermsAndConditions !== undefined) itemData.TermsAndConditions = updates.TermsAndConditions ? ValidationUtils.sanitizeHtml(updates.TermsAndConditions) : null;
      if (updates.SpecialInstructions !== undefined) itemData.SpecialInstructions = updates.SpecialInstructions ? ValidationUtils.sanitizeHtml(updates.SpecialInstructions) : null;
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes ? ValidationUtils.sanitizeHtml(updates.Notes) : null;
      if (updates.Attachments !== undefined) itemData.Attachments = updates.Attachments;

      // Status change fields
      if (updates.ApprovedById) itemData.ApprovedById = updates.ApprovedById;
      if (updates.ApprovedDate) itemData.ApprovedDate = updates.ApprovedDate;
      if (updates.SentById) itemData.SentById = updates.SentById;
      if (updates.SentDate) itemData.SentDate = updates.SentDate;
      if (updates.SentMethod !== undefined) itemData.SentMethod = updates.SentMethod;
      if (updates.AcknowledgedDate) itemData.AcknowledgedDate = updates.AcknowledgedDate;
      if (updates.ClosedDate) itemData.ClosedDate = updates.ClosedDate;

      await this.sp.web.lists.getByTitle(this.PO_LIST).items.getById(validId).update(itemData);
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error updating purchase order:', error);
      throw error;
    }
  }

  public async deletePurchaseOrder(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Check if PO is in Draft status
      const po = await this.getPurchaseOrderById(validId);
      if (po.Status !== POStatus.Draft) {
        throw new Error('Only draft purchase orders can be deleted');
      }

      // Delete line items first
      const lineItems = await this.getPOLineItems(validId);
      for (const item of lineItems) {
        if (item.Id) {
          await this.deletePOLineItem(item.Id);
        }
      }

      // Delete PO
      await this.sp.web.lists.getByTitle(this.PO_LIST).items.getById(validId).delete();
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error deleting purchase order:', error);
      throw error;
    }
  }

  // ==================== PO Workflow ====================

  public async approvePO(id: number, approvedById: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);
      const validApprovedById = ValidationUtils.validateInteger(approvedById, 'approvedById', 1);

      await this.updatePurchaseOrder(validId, {
        Status: POStatus.Approved,
        ApprovedById: validApprovedById,
        ApprovedDate: new Date()
      });
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error approving PO:', error);
      throw error;
    }
  }

  public async sendPO(id: number, sentById: number, method: string = 'Email'): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);
      const validSentById = ValidationUtils.validateInteger(sentById, 'sentById', 1);

      const po = await this.getPurchaseOrderById(validId);
      if (po.Status !== POStatus.Draft && po.Status !== POStatus.Approved) {
        throw new Error('PO must be in Draft or Approved status to send');
      }

      await this.updatePurchaseOrder(validId, {
        Status: POStatus.Sent,
        SentById: validSentById,
        SentDate: new Date(),
        SentMethod: method
      });

      // TODO: Actually send email to vendor
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error sending PO:', error);
      throw error;
    }
  }

  public async acknowledgePO(id: number, vendorReference?: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      await this.updatePurchaseOrder(validId, {
        Status: POStatus.Acknowledged,
        AcknowledgedDate: new Date(),
        VendorReference: vendorReference
      });
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error acknowledging PO:', error);
      throw error;
    }
  }

  public async closePO(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Verify all items are received
      const lineItems = await this.getPOLineItems(validId);
      const allReceived = lineItems.every(item => item.ReceivedStatus === ReceiptStatus.Complete);

      if (!allReceived) {
        throw new Error('Cannot close PO - not all items have been received');
      }

      await this.updatePurchaseOrder(validId, {
        Status: POStatus.Closed,
        ClosedDate: new Date()
      });
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error closing PO:', error);
      throw error;
    }
  }

  public async cancelPO(id: number, reason?: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const po = await this.getPurchaseOrderById(validId);
      if (po.Status === POStatus.Closed) {
        throw new Error('Cannot cancel a closed PO');
      }

      const updateData: Partial<IPurchaseOrder> = {
        Status: POStatus.Cancelled
      };

      if (reason) {
        updateData.Notes = po.Notes
          ? `${po.Notes}\n\nCancellation Reason: ${reason}`
          : `Cancellation Reason: ${reason}`;
      }

      await this.updatePurchaseOrder(validId, updateData);
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error cancelling PO:', error);
      throw error;
    }
  }

  // ==================== PO Line Items ====================

  public async getPOLineItems(purchaseOrderId: number): Promise<IPOLineItem[]> {
    try {
      const validPOId = ValidationUtils.validateInteger(purchaseOrderId, 'purchaseOrderId', 1);

      const items = await this.sp.web.lists.getByTitle(this.PO_LINE_ITEMS_LIST).items
        .select(
          'Id', 'Title', 'PurchaseOrderId', 'LineNumber',
          'CatalogItemId', 'ItemCode', 'Description',
          'Quantity', 'UnitOfMeasure', 'UnitPrice', 'TotalPrice',
          'TaxRate', 'TaxAmount',
          'QuantityReceived', 'QuantityPending', 'QuantityRejected', 'ReceivedStatus',
          'AssetIds', 'Specifications', 'DeliveryDate', 'Notes',
          'Created', 'Modified'
        )
        .filter(`PurchaseOrderId eq ${validPOId}`)
        .orderBy('LineNumber', true)();

      return items.map(this.mapPOLineItemFromSP);
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error getting PO line items:', error);
      throw error;
    }
  }

  public async createPOLineItem(lineItem: Partial<IPOLineItem>): Promise<number> {
    try {
      if (!lineItem.PurchaseOrderId || !lineItem.Description || lineItem.Quantity === undefined) {
        throw new Error('PurchaseOrderId, Description, and Quantity are required');
      }

      const totalPrice = lineItem.TotalPrice || (lineItem.Quantity * (lineItem.UnitPrice || 0));

      const itemData: Record<string, unknown> = {
        Title: ValidationUtils.sanitizeHtml(lineItem.Description.substring(0, 255)),
        PurchaseOrderId: ValidationUtils.validateInteger(lineItem.PurchaseOrderId, 'PurchaseOrderId', 1),
        LineNumber: lineItem.LineNumber || 1,
        Description: ValidationUtils.sanitizeHtml(lineItem.Description),
        Quantity: lineItem.Quantity,
        UnitOfMeasure: lineItem.UnitOfMeasure || UnitOfMeasure.Each,
        UnitPrice: lineItem.UnitPrice || 0,
        TotalPrice: totalPrice,
        QuantityReceived: 0,
        QuantityPending: lineItem.Quantity,
        QuantityRejected: 0,
        ReceivedStatus: ReceiptStatus.Pending
      };

      if (lineItem.CatalogItemId) itemData.CatalogItemId = ValidationUtils.validateInteger(lineItem.CatalogItemId, 'CatalogItemId', 1);
      if (lineItem.ItemCode) itemData.ItemCode = ValidationUtils.sanitizeHtml(lineItem.ItemCode);
      if (lineItem.TaxRate !== undefined) itemData.TaxRate = lineItem.TaxRate;
      if (lineItem.TaxAmount !== undefined) itemData.TaxAmount = lineItem.TaxAmount;
      if (lineItem.Specifications) itemData.Specifications = ValidationUtils.sanitizeHtml(lineItem.Specifications);
      if (lineItem.DeliveryDate) itemData.DeliveryDate = ValidationUtils.validateDate(lineItem.DeliveryDate, 'DeliveryDate');
      if (lineItem.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(lineItem.Notes);

      const result = await this.sp.web.lists.getByTitle(this.PO_LINE_ITEMS_LIST).items.add(itemData);

      // Recalculate PO total
      await this.recalculatePOTotal(lineItem.PurchaseOrderId);

      return result.data.Id;
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error creating PO line item:', error);
      throw error;
    }
  }

  public async updatePOLineItem(id: number, updates: Partial<IPOLineItem>): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Get current item for PO ID
      const currentItem = await this.sp.web.lists.getByTitle(this.PO_LINE_ITEMS_LIST).items
        .getById(validId)
        .select('PurchaseOrderId')();

      const itemData: Record<string, unknown> = {};

      if (updates.Description) {
        itemData.Description = ValidationUtils.sanitizeHtml(updates.Description);
        itemData.Title = ValidationUtils.sanitizeHtml(updates.Description.substring(0, 255));
      }
      if (updates.Quantity !== undefined) {
        itemData.Quantity = updates.Quantity;
        itemData.QuantityPending = updates.Quantity - (updates.QuantityReceived || 0);
      }
      if (updates.UnitOfMeasure) itemData.UnitOfMeasure = updates.UnitOfMeasure;
      if (updates.UnitPrice !== undefined) itemData.UnitPrice = updates.UnitPrice;
      if (updates.TotalPrice !== undefined) itemData.TotalPrice = updates.TotalPrice;
      if (updates.TaxRate !== undefined) itemData.TaxRate = updates.TaxRate;
      if (updates.TaxAmount !== undefined) itemData.TaxAmount = updates.TaxAmount;
      if (updates.QuantityReceived !== undefined) itemData.QuantityReceived = updates.QuantityReceived;
      if (updates.QuantityPending !== undefined) itemData.QuantityPending = updates.QuantityPending;
      if (updates.QuantityRejected !== undefined) itemData.QuantityRejected = updates.QuantityRejected;
      if (updates.ReceivedStatus) itemData.ReceivedStatus = updates.ReceivedStatus;
      if (updates.AssetIds !== undefined) itemData.AssetIds = updates.AssetIds;
      if (updates.Specifications !== undefined) itemData.Specifications = updates.Specifications ? ValidationUtils.sanitizeHtml(updates.Specifications) : null;
      if (updates.DeliveryDate) itemData.DeliveryDate = ValidationUtils.validateDate(updates.DeliveryDate, 'DeliveryDate');
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes ? ValidationUtils.sanitizeHtml(updates.Notes) : null;

      await this.sp.web.lists.getByTitle(this.PO_LINE_ITEMS_LIST).items.getById(validId).update(itemData);

      // Recalculate PO total
      await this.recalculatePOTotal(currentItem.PurchaseOrderId);
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error updating PO line item:', error);
      throw error;
    }
  }

  public async deletePOLineItem(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Get PO ID before deleting
      const item = await this.sp.web.lists.getByTitle(this.PO_LINE_ITEMS_LIST).items
        .getById(validId)
        .select('PurchaseOrderId')();

      await this.sp.web.lists.getByTitle(this.PO_LINE_ITEMS_LIST).items.getById(validId).delete();

      // Recalculate PO total
      await this.recalculatePOTotal(item.PurchaseOrderId);
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error deleting PO line item:', error);
      throw error;
    }
  }

  private async recalculatePOTotal(purchaseOrderId: number): Promise<void> {
    try {
      const lineItems = await this.getPOLineItems(purchaseOrderId);
      const subtotal = lineItems.reduce((sum, item) => sum + (item.TotalPrice || 0), 0);
      const taxAmount = lineItems.reduce((sum, item) => sum + (item.TaxAmount || 0), 0);

      const po = await this.getPurchaseOrderById(purchaseOrderId);
      const totalAmount = subtotal + taxAmount + (po.ShippingCost || 0) - (po.DiscountAmount || 0);

      await this.sp.web.lists.getByTitle(this.PO_LIST).items.getById(purchaseOrderId).update({
        Subtotal: subtotal,
        TaxAmount: taxAmount,
        TotalAmount: totalAmount
      });
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error recalculating PO total:', error);
    }
  }

  // ==================== Goods Receipt ====================

  public async createGoodsReceipt(receipt: Partial<IGoodsReceipt>, lineItems: Partial<IReceiptLineItem>[]): Promise<number> {
    try {
      if (!receipt.PurchaseOrderId || !receipt.ReceivedById) {
        throw new Error('PurchaseOrderId and ReceivedById are required');
      }

      // Generate receipt number
      const receiptNumber = await this.generateReceiptNumber();

      // Get PO for vendor ID
      const po = await this.getPurchaseOrderById(receipt.PurchaseOrderId);

      const itemData: Record<string, unknown> = {
        Title: `Receipt for ${po.PONumber}`,
        ReceiptNumber: receiptNumber,
        PurchaseOrderId: ValidationUtils.validateInteger(receipt.PurchaseOrderId, 'PurchaseOrderId', 1),
        VendorId: po.VendorId,
        ReceiptDate: new Date().toISOString(),
        ReceivedById: ValidationUtils.validateInteger(receipt.ReceivedById, 'ReceivedById', 1),
        Status: ReceiptStatus.Complete
      };

      if (receipt.DeliveryNote) itemData.DeliveryNote = ValidationUtils.sanitizeHtml(receipt.DeliveryNote);
      if (receipt.PackingSlip) itemData.PackingSlip = ValidationUtils.sanitizeHtml(receipt.PackingSlip);
      if (receipt.CarrierName) itemData.CarrierName = ValidationUtils.sanitizeHtml(receipt.CarrierName);
      if (receipt.TrackingNumber) itemData.TrackingNumber = ValidationUtils.sanitizeHtml(receipt.TrackingNumber);
      if (receipt.ReceivedAtLocation) itemData.ReceivedAtLocation = ValidationUtils.sanitizeHtml(receipt.ReceivedAtLocation);
      if (receipt.StorageLocation) itemData.StorageLocation = ValidationUtils.sanitizeHtml(receipt.StorageLocation);
      if (receipt.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(receipt.Notes);
      if (receipt.Attachments) itemData.Attachments = receipt.Attachments;

      const result = await this.sp.web.lists.getByTitle(this.GOODS_RECEIPTS_LIST).items.add(itemData);
      const receiptId = result.data.Id;

      // Create receipt line items and update PO line items
      let hasPartialReceipt = false;
      for (const lineItem of lineItems) {
        await this.createReceiptLineItem({
          ...lineItem,
          ReceiptId: receiptId
        });

        // Update PO line item quantities
        if (lineItem.POLineItemId && lineItem.QuantityReceived !== undefined) {
          const poLineItem = await this.sp.web.lists.getByTitle(this.PO_LINE_ITEMS_LIST).items
            .getById(lineItem.POLineItemId)
            .select('Quantity', 'QuantityReceived', 'QuantityRejected')();

          const newReceived = (poLineItem.QuantityReceived || 0) + lineItem.QuantityReceived;
          const newRejected = (poLineItem.QuantityRejected || 0) + (lineItem.QuantityRejected || 0);
          const newPending = poLineItem.Quantity - newReceived - newRejected;

          let status = ReceiptStatus.Pending;
          if (newReceived >= poLineItem.Quantity) {
            status = ReceiptStatus.Complete;
          } else if (newReceived > 0) {
            status = ReceiptStatus.Partial;
            hasPartialReceipt = true;
          }

          await this.updatePOLineItem(lineItem.POLineItemId, {
            QuantityReceived: newReceived,
            QuantityPending: newPending,
            QuantityRejected: newRejected,
            ReceivedStatus: status
          });
        }
      }

      // Update PO status
      const poLineItems = await this.getPOLineItems(receipt.PurchaseOrderId);
      const allComplete = poLineItems.every(item => item.ReceivedStatus === ReceiptStatus.Complete);
      const anyReceived = poLineItems.some(item => (item.QuantityReceived || 0) > 0);

      let newPOStatus = po.Status;
      if (allComplete) {
        newPOStatus = POStatus.Received;
        await this.updatePurchaseOrder(receipt.PurchaseOrderId, {
          Status: newPOStatus,
          ActualDeliveryDate: new Date()
        });
      } else if (anyReceived) {
        newPOStatus = POStatus.PartiallyReceived;
        await this.updatePurchaseOrder(receipt.PurchaseOrderId, {
          Status: newPOStatus
        });
      }

      return receiptId;
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error creating goods receipt:', error);
      throw error;
    }
  }

  public async getGoodsReceipts(purchaseOrderId: number): Promise<IGoodsReceipt[]> {
    try {
      const validPOId = ValidationUtils.validateInteger(purchaseOrderId, 'purchaseOrderId', 1);

      const items = await this.sp.web.lists.getByTitle(this.GOODS_RECEIPTS_LIST).items
        .select(
          'Id', 'Title', 'ReceiptNumber', 'PurchaseOrderId', 'VendorId',
          'ReceiptDate', 'ReceivedById', 'ReceivedBy/Title',
          'DeliveryNote', 'PackingSlip', 'CarrierName', 'TrackingNumber',
          'Status', 'ReceivedAtLocation', 'StorageLocation',
          'Attachments', 'Notes',
          'Created', 'Modified'
        )
        .expand('ReceivedBy')
        .filter(`PurchaseOrderId eq ${validPOId}`)
        .orderBy('ReceiptDate', false)();

      return items.map(this.mapGoodsReceiptFromSP);
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error getting goods receipts:', error);
      throw error;
    }
  }

  private async createReceiptLineItem(lineItem: Partial<IReceiptLineItem>): Promise<number> {
    try {
      if (!lineItem.ReceiptId || !lineItem.POLineItemId) {
        throw new Error('ReceiptId and POLineItemId are required');
      }

      const itemData: Record<string, unknown> = {
        Title: `Line item for receipt`,
        ReceiptId: ValidationUtils.validateInteger(lineItem.ReceiptId, 'ReceiptId', 1),
        POLineItemId: ValidationUtils.validateInteger(lineItem.POLineItemId, 'POLineItemId', 1),
        QuantityExpected: lineItem.QuantityExpected || 0,
        QuantityReceived: lineItem.QuantityReceived || 0,
        QuantityRejected: lineItem.QuantityRejected || 0,
        Condition: lineItem.Condition || 'Good'
      };

      if (lineItem.RejectionReason) itemData.RejectionReason = ValidationUtils.sanitizeHtml(lineItem.RejectionReason);
      if (lineItem.SerialNumbers) itemData.SerialNumbers = lineItem.SerialNumbers;
      if (lineItem.BatchNumber) itemData.BatchNumber = ValidationUtils.sanitizeHtml(lineItem.BatchNumber);
      if (lineItem.AssetIdsCreated) itemData.AssetIdsCreated = lineItem.AssetIdsCreated;
      if (lineItem.InspectedById) itemData.InspectedById = ValidationUtils.validateInteger(lineItem.InspectedById, 'InspectedById', 1);
      if (lineItem.InspectionDate) itemData.InspectionDate = ValidationUtils.validateDate(lineItem.InspectionDate, 'InspectionDate');
      if (lineItem.InspectionNotes) itemData.InspectionNotes = ValidationUtils.sanitizeHtml(lineItem.InspectionNotes);
      if (lineItem.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(lineItem.Notes);

      const result = await this.sp.web.lists.getByTitle(this.RECEIPT_LINE_ITEMS_LIST).items.add(itemData);
      return result.data.Id;
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error creating receipt line item:', error);
      throw error;
    }
  }

  // ==================== Statistics ====================

  public async getPOStatistics(): Promise<{
    total: number;
    draft: number;
    sent: number;
    acknowledged: number;
    partiallyReceived: number;
    received: number;
    closed: number;
    cancelled: number;
    totalValue: number;
    avgValue: number;
    overdue: number;
    byVendor: { vendorId: number; vendorName: string; count: number; value: number }[];
  }> {
    try {
      const pos = await this.getPurchaseOrders();

      const stats = {
        total: pos.length,
        draft: 0,
        sent: 0,
        acknowledged: 0,
        partiallyReceived: 0,
        received: 0,
        closed: 0,
        cancelled: 0,
        totalValue: 0,
        avgValue: 0,
        overdue: 0,
        byVendor: [] as { vendorId: number; vendorName: string; count: number; value: number }[]
      };

      const vendorMap = new Map<number, { name: string; count: number; value: number }>();
      const now = new Date();

      for (const po of pos) {
        // Count by status
        switch (po.Status) {
          case POStatus.Draft: stats.draft++; break;
          case POStatus.Sent: stats.sent++; break;
          case POStatus.Acknowledged: stats.acknowledged++; break;
          case POStatus.PartiallyReceived: stats.partiallyReceived++; break;
          case POStatus.Received: stats.received++; break;
          case POStatus.Closed: stats.closed++; break;
          case POStatus.Cancelled: stats.cancelled++; break;
        }

        // Total value (exclude cancelled)
        if (po.Status !== POStatus.Cancelled) {
          stats.totalValue += po.TotalAmount || 0;
        }

        // Check overdue
        if (po.ExpectedDeliveryDate && new Date(po.ExpectedDeliveryDate) < now &&
            (po.Status === POStatus.Sent || po.Status === POStatus.Acknowledged || po.Status === POStatus.PartiallyReceived)) {
          stats.overdue++;
        }

        // By vendor
        if (po.VendorId) {
          const existing = vendorMap.get(po.VendorId);
          if (existing) {
            existing.count++;
            existing.value += po.TotalAmount || 0;
          } else {
            vendorMap.set(po.VendorId, {
              name: (po.Vendor as unknown as Record<string, unknown>)?.Title as string || 'Unknown',
              count: 1,
              value: po.TotalAmount || 0
            });
          }
        }
      }

      stats.avgValue = stats.total > 0 ? stats.totalValue / (stats.total - stats.cancelled) : 0;

      // Convert vendor map to array
      stats.byVendor = Array.from(vendorMap.entries())
        .map(([vendorId, data]) => ({
          vendorId,
          vendorName: data.name,
          count: data.count,
          value: data.value
        }))
        .sort((a, b) => b.value - a.value)
        .slice(0, 10);

      return stats;
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error getting PO statistics:', error);
      throw error;
    }
  }

  // ==================== Helper Functions ====================

  private async generatePONumber(): Promise<string> {
    try {
      const year = new Date().getFullYear();
      const prefix = `PO-${year}-`;

      const items = await this.sp.web.lists.getByTitle(this.PO_LIST).items
        .select('PONumber')
        .filter(`substringof('${prefix}', PONumber)`)
        .orderBy('Id', false)
        .top(1)();

      let nextNumber = 1;
      if (items.length > 0 && items[0].PONumber) {
        const match = items[0].PONumber.match(/PO-\d{4}-(\d+)/);
        if (match) {
          nextNumber = parseInt(match[1], 10) + 1;
        }
      }

      return `${prefix}${nextNumber.toString().padStart(5, '0')}`;
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error generating PO number:', error);
      return `PO-${Date.now()}`;
    }
  }

  private async generateReceiptNumber(): Promise<string> {
    try {
      const year = new Date().getFullYear();
      const prefix = `GR-${year}-`;

      const items = await this.sp.web.lists.getByTitle(this.GOODS_RECEIPTS_LIST).items
        .select('ReceiptNumber')
        .filter(`substringof('${prefix}', ReceiptNumber)`)
        .orderBy('Id', false)
        .top(1)();

      let nextNumber = 1;
      if (items.length > 0 && items[0].ReceiptNumber) {
        const match = items[0].ReceiptNumber.match(/GR-\d{4}-(\d+)/);
        if (match) {
          nextNumber = parseInt(match[1], 10) + 1;
        }
      }

      return `${prefix}${nextNumber.toString().padStart(5, '0')}`;
    } catch (error) {
      logger.error('PurchaseOrderService', 'Error generating receipt number:', error);
      return `GR-${Date.now()}`;
    }
  }

  // ==================== Mapping Functions ====================

  private mapPOFromSP(item: Record<string, unknown>): IPurchaseOrder {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      PONumber: item.PONumber as string,
      VendorId: item.VendorId as number,
      Vendor: item.Vendor as Record<string, unknown>,
      Status: item.Status as POStatus,
      RequisitionIds: item.RequisitionIds as string,
      OrderDate: item.OrderDate ? new Date(item.OrderDate as string) : new Date(),
      ExpectedDeliveryDate: item.ExpectedDeliveryDate ? new Date(item.ExpectedDeliveryDate as string) : undefined,
      ActualDeliveryDate: item.ActualDeliveryDate ? new Date(item.ActualDeliveryDate as string) : undefined,
      SentDate: item.SentDate ? new Date(item.SentDate as string) : undefined,
      AcknowledgedDate: item.AcknowledgedDate ? new Date(item.AcknowledgedDate as string) : undefined,
      ClosedDate: item.ClosedDate ? new Date(item.ClosedDate as string) : undefined,
      ShipToAddress: item.ShipToAddress as string,
      ShipToAttention: item.ShipToAttention as string,
      BillToAddress: item.BillToAddress as string,
      BillToAttention: item.BillToAttention as string,
      Subtotal: item.Subtotal as number || 0,
      TaxRate: item.TaxRate as number,
      TaxAmount: item.TaxAmount as number || 0,
      ShippingCost: item.ShippingCost as number || 0,
      DiscountAmount: item.DiscountAmount as number || 0,
      TotalAmount: item.TotalAmount as number || 0,
      Currency: item.Currency as Currency || Currency.USD,
      PaymentTerms: item.PaymentTerms as PaymentTerms || PaymentTerms.Net30,
      VendorQuoteNumber: item.VendorQuoteNumber as string,
      VendorReference: item.VendorReference as string,
      ApprovedById: item.ApprovedById as number,
      ApprovedBy: item.ApprovedBy as Record<string, unknown>,
      ApprovedDate: item.ApprovedDate ? new Date(item.ApprovedDate as string) : undefined,
      SentById: item.SentById as number,
      SentBy: item.SentBy as Record<string, unknown>,
      SentMethod: item.SentMethod as string,
      TermsAndConditions: item.TermsAndConditions as string,
      SpecialInstructions: item.SpecialInstructions as string,
      BudgetId: item.BudgetId as number,
      CostCenter: item.CostCenter as string,
      Department: item.Department as string,
      ProcessId: item.ProcessId as number,
      TaskId: item.TaskId as number,
      Attachments: item.Attachments as string,
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }

  private mapPOLineItemFromSP(item: Record<string, unknown>): IPOLineItem {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      PurchaseOrderId: item.PurchaseOrderId as number,
      LineNumber: item.LineNumber as number,
      CatalogItemId: item.CatalogItemId as number,
      ItemCode: item.ItemCode as string,
      Description: item.Description as string,
      Quantity: item.Quantity as number,
      UnitOfMeasure: item.UnitOfMeasure as UnitOfMeasure,
      UnitPrice: item.UnitPrice as number || 0,
      TotalPrice: item.TotalPrice as number || 0,
      TaxRate: item.TaxRate as number,
      TaxAmount: item.TaxAmount as number,
      QuantityReceived: item.QuantityReceived as number || 0,
      QuantityPending: item.QuantityPending as number || 0,
      QuantityRejected: item.QuantityRejected as number || 0,
      ReceivedStatus: item.ReceivedStatus as ReceiptStatus || ReceiptStatus.Pending,
      AssetIds: item.AssetIds as string,
      Specifications: item.Specifications as string,
      DeliveryDate: item.DeliveryDate ? new Date(item.DeliveryDate as string) : undefined,
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }

  private mapGoodsReceiptFromSP(item: Record<string, unknown>): IGoodsReceipt {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      ReceiptNumber: item.ReceiptNumber as string,
      PurchaseOrderId: item.PurchaseOrderId as number,
      VendorId: item.VendorId as number,
      ReceiptDate: item.ReceiptDate ? new Date(item.ReceiptDate as string) : new Date(),
      ReceivedById: item.ReceivedById as number,
      ReceivedBy: item.ReceivedBy as Record<string, unknown>,
      DeliveryNote: item.DeliveryNote as string,
      PackingSlip: item.PackingSlip as string,
      CarrierName: item.CarrierName as string,
      TrackingNumber: item.TrackingNumber as string,
      Status: item.Status as ReceiptStatus,
      ReceivedAtLocation: item.ReceivedAtLocation as string,
      StorageLocation: item.StorageLocation as string,
      Attachments: item.Attachments as string,
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }
}
