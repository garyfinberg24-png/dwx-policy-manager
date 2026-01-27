// @ts-nocheck
/* eslint-disable @typescript-eslint/no-explicit-any */
// Procurement Service
// Unified procurement management service providing dashboard, statistics, and cross-entity operations
// Note: Some fields may not exist in the SharePoint list - mapping handles this gracefully

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  IProcurementDashboard,
  IProcurementStatistics,
  IVendorAlert,
  IBudgetAlert,
  ISpendTrendItem,
  IVendor,
  IContract,
  IPurchaseRequisition,
  IPurchaseOrder,
  IInvoice,
  VendorStatus,
  ContractStatus,
  RequisitionStatus,
  POStatus,
  InvoiceStatus,
  VendorCategory,
  IJMLProcurementRequest,
  IJMLProcurementResult
} from '../models/IProcurement';
import { VendorService } from './VendorService';
import { RequisitionService } from './RequisitionService';
import { PurchaseOrderService } from './PurchaseOrderService';
import { ContractService } from './ContractService';
import { InvoiceService } from './InvoiceService';
import { BudgetService } from './BudgetService';
import { CatalogService } from './CatalogService';
import { logger } from './LoggingService';

export class ProcurementService {
  private sp: SPFI;
  private vendorService: VendorService;
  private requisitionService: RequisitionService;
  private purchaseOrderService: PurchaseOrderService;
  private contractService: ContractService;
  private invoiceService: InvoiceService;
  private budgetService: BudgetService;
  private catalogService: CatalogService;

  constructor(sp: SPFI) {
    this.sp = sp;
    this.vendorService = new VendorService(sp);
    this.requisitionService = new RequisitionService(sp);
    this.purchaseOrderService = new PurchaseOrderService(sp);
    this.contractService = new ContractService(sp);
    this.invoiceService = new InvoiceService(sp);
    this.budgetService = new BudgetService(sp);
    this.catalogService = new CatalogService(sp);
  }

  // ==================== Service Accessors ====================

  public get vendors(): VendorService {
    return this.vendorService;
  }

  public get requisitions(): RequisitionService {
    return this.requisitionService;
  }

  public get purchaseOrders(): PurchaseOrderService {
    return this.purchaseOrderService;
  }

  public get contracts(): ContractService {
    return this.contractService;
  }

  public get invoices(): InvoiceService {
    return this.invoiceService;
  }

  public get budgets(): BudgetService {
    return this.budgetService;
  }

  public get catalog(): CatalogService {
    return this.catalogService;
  }

  // ==================== Dashboard ====================

  public async getDashboard(): Promise<IProcurementDashboard> {
    try {
      // Fetch all data in parallel for better performance
      const [
        statistics,
        recentRequisitions,
        pendingApprovals,
        recentPOs,
        expiringContracts,
        overdueInvoices,
        vendorAlerts,
        budgetAlerts,
        spendTrend
      ] = await Promise.all([
        this.getStatistics(),
        this.getRecentRequisitions(10),
        this.getPendingApprovals(10),
        this.getRecentPurchaseOrders(10),
        this.getExpiringContracts(90),
        this.getOverdueInvoices(10),
        this.getVendorAlerts(),
        this.getBudgetAlerts(),
        this.getSpendTrend(12)
      ]);

      return {
        statistics,
        recentRequisitions,
        pendingApprovals,
        recentPurchaseOrders: recentPOs,
        expiringContracts,
        overdueInvoices,
        vendorAlerts,
        budgetAlerts,
        spendTrend
      };
    } catch (error) {
      logger.error('ProcurementService', 'Error getting dashboard:', error);
      throw error;
    }
  }

  // ==================== Statistics ====================

  public async getStatistics(): Promise<IProcurementStatistics> {
    try {
      // Fetch statistics from all services in parallel
      const [
        vendorStats,
        requisitionStats,
        poStats,
        invoiceStats,
        contractStats,
        budgetStats
      ] = await Promise.all([
        this.vendorService.getVendorStatistics(),
        this.requisitionService.getRequisitionStatistics(),
        this.purchaseOrderService.getPOStatistics(),
        this.invoiceService.getInvoiceStatistics(),
        this.contractService.getContractStatistics(),
        this.budgetService.getBudgetStatistics()
      ]);

      // Get spend data
      const spendData = await this.calculateSpendData();

      return {
        // Vendors
        totalVendors: vendorStats.total,
        activeVendors: vendorStats.active,
        preferredVendors: vendorStats.preferred,
        vendorsByCategory: vendorStats.byCategory as { [key in VendorCategory]?: number },
        vendorsByStatus: {} as { [key in VendorStatus]?: number },
        avgVendorRating: vendorStats.avgRating,

        // Requisitions
        totalRequisitions: requisitionStats.total,
        pendingApproval: requisitionStats.pendingApproval,
        requisitionsByStatus: requisitionStats.byDepartment as unknown as { [key in RequisitionStatus]?: number },
        avgRequisitionValue: requisitionStats.avgValue,
        avgApprovalTime: 0, // TODO: Calculate from approval history

        // Purchase Orders
        totalPOs: poStats.total,
        openPOs: poStats.sent + poStats.acknowledged + poStats.partiallyReceived,
        posByStatus: {} as { [key in POStatus]?: number },
        totalPOValue: poStats.totalValue,
        avgPOValue: poStats.avgValue,

        // Invoices
        totalInvoices: invoiceStats.total,
        pendingInvoices: invoiceStats.pending,
        overdueInvoices: invoiceStats.overdue,
        invoicesByStatus: invoiceStats.byStatus as { [key in InvoiceStatus]?: number },
        totalInvoiceValue: invoiceStats.totalValue,
        avgPaymentDays: invoiceStats.avgPaymentDays,

        // Contracts
        totalContracts: contractStats.total,
        activeContracts: contractStats.active,
        expiringContracts: contractStats.expiring90Days,
        contractsByStatus: contractStats.byStatus as { [key in ContractStatus]?: number },
        totalContractValue: contractStats.totalValue,

        // Spend
        totalSpendYTD: spendData.ytd,
        totalSpendMTD: spendData.mtd,
        spendByCategory: spendData.byCategory,
        spendByDepartment: spendData.byDepartment,
        topVendorsBySpend: poStats.byVendor.slice(0, 10).map(v => ({ vendorId: v.vendorId, vendorName: v.vendorName, spend: v.value })),

        // Budget
        totalBudget: budgetStats.totalBudget,
        totalAllocated: budgetStats.totalAllocated,
        totalSpent: budgetStats.totalSpent,
        totalRemaining: budgetStats.totalRemaining,
        budgetUtilization: budgetStats.avgUtilization
      };
    } catch (error) {
      logger.error('ProcurementService', 'Error getting statistics:', error);
      throw error;
    }
  }

  // ==================== Recent Items ====================

  private async getRecentRequisitions(limit: number): Promise<IPurchaseRequisition[]> {
    try {
      const requisitions = await this.requisitionService.getRequisitions();
      return requisitions.slice(0, limit);
    } catch (error) {
      logger.error('ProcurementService', 'Error getting recent requisitions:', error);
      return [];
    }
  }

  private async getPendingApprovals(limit: number): Promise<IPurchaseRequisition[]> {
    try {
      const pending = await this.requisitionService.getPendingApprovals();
      return pending.slice(0, limit);
    } catch (error) {
      logger.error('ProcurementService', 'Error getting pending approvals:', error);
      return [];
    }
  }

  private async getRecentPurchaseOrders(limit: number): Promise<IPurchaseOrder[]> {
    try {
      const pos = await this.purchaseOrderService.getPurchaseOrders();
      return pos.slice(0, limit);
    } catch (error) {
      logger.error('ProcurementService', 'Error getting recent POs:', error);
      return [];
    }
  }

  private async getExpiringContracts(days: number): Promise<IContract[]> {
    try {
      return await this.contractService.getExpiringContracts(days);
    } catch (error) {
      logger.error('ProcurementService', 'Error getting expiring contracts:', error);
      return [];
    }
  }

  private async getOverdueInvoices(limit: number): Promise<IInvoice[]> {
    try {
      const overdue = await this.invoiceService.getOverdueInvoices();
      return overdue.slice(0, limit);
    } catch (error) {
      logger.error('ProcurementService', 'Error getting overdue invoices:', error);
      return [];
    }
  }

  // ==================== Alerts ====================

  private async getVendorAlerts(): Promise<IVendorAlert[]> {
    try {
      const alerts: IVendorAlert[] = [];
      const today = new Date();

      // Get vendors with low ratings
      const vendors = await this.vendorService.getVendors();
      for (const vendor of vendors) {
        if (vendor.Rating && vendor.Rating < 3) {
          alerts.push({
            vendorId: vendor.Id!,
            vendorName: vendor.Title,
            alertType: 'Performance',
            message: `Vendor ${vendor.Title} has a low rating of ${vendor.Rating.toFixed(1)}`,
            severity: vendor.Rating < 2 ? 'Critical' : 'Warning',
            date: new Date()
          });
        }

        // Check insurance expiry
        if (vendor.InsuranceExpiry) {
          const expiryDate = new Date(vendor.InsuranceExpiry);
          const daysUntilExpiry = Math.floor((expiryDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));

          if (daysUntilExpiry < 0) {
            alerts.push({
              vendorId: vendor.Id!,
              vendorName: vendor.Title,
              alertType: 'Compliance',
              message: `Insurance for ${vendor.Title} has expired`,
              severity: 'Critical',
              date: expiryDate
            });
          } else if (daysUntilExpiry <= 30) {
            alerts.push({
              vendorId: vendor.Id!,
              vendorName: vendor.Title,
              alertType: 'Compliance',
              message: `Insurance for ${vendor.Title} expires in ${daysUntilExpiry} days`,
              severity: 'Warning',
              date: expiryDate
            });
          }
        }
      }

      // Get contract alerts
      const expiringContracts = await this.contractService.getExpiringContracts(30);
      for (const contract of expiringContracts) {
        alerts.push({
          vendorId: contract.VendorId,
          vendorName: contract.Title,
          alertType: 'Contract',
          message: `Contract ${contract.ContractNumber} expires on ${new Date(contract.EndDate).toLocaleDateString()}`,
          severity: 'Warning',
          date: new Date(contract.EndDate)
        });
      }

      return alerts.sort((a, b) => {
        const severityOrder = { Critical: 0, Warning: 1, Info: 2 };
        return severityOrder[a.severity] - severityOrder[b.severity];
      }).slice(0, 20);
    } catch (error) {
      logger.error('ProcurementService', 'Error getting vendor alerts:', error);
      return [];
    }
  }

  private async getBudgetAlerts(): Promise<IBudgetAlert[]> {
    try {
      return await this.budgetService.getBudgetAlerts();
    } catch (error) {
      logger.error('ProcurementService', 'Error getting budget alerts:', error);
      return [];
    }
  }

  // ==================== Spend Analysis ====================

  private async calculateSpendData(): Promise<{
    ytd: number;
    mtd: number;
    byCategory: { [key in VendorCategory]?: number };
    byDepartment: { [department: string]: number };
  }> {
    try {
      const today = new Date();
      const yearStart = new Date(today.getFullYear(), 0, 1);
      const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);

      // Get paid invoices for the year
      const invoices = await this.invoiceService.getInvoices({
        status: [InvoiceStatus.Paid],
        fromDate: yearStart
      });

      let ytd = 0;
      let mtd = 0;
      const byCategory: { [key in VendorCategory]?: number } = {};
      const byDepartment: { [department: string]: number } = {};

      for (const invoice of invoices) {
        const amount = invoice.TotalAmount || 0;
        ytd += amount;

        if (invoice.PaymentDate && new Date(invoice.PaymentDate) >= monthStart) {
          mtd += amount;
        }

        // By department
        if (invoice.Department) {
          byDepartment[invoice.Department] = (byDepartment[invoice.Department] || 0) + amount;
        }
      }

      return { ytd, mtd, byCategory, byDepartment };
    } catch (error) {
      logger.error('ProcurementService', 'Error calculating spend data:', error);
      return { ytd: 0, mtd: 0, byCategory: {}, byDepartment: {} };
    }
  }

  private async getSpendTrend(months: number): Promise<ISpendTrendItem[]> {
    try {
      const trend: ISpendTrendItem[] = [];
      const today = new Date();

      for (let i = months - 1; i >= 0; i--) {
        const monthDate = new Date(today.getFullYear(), today.getMonth() - i, 1);
        const monthEnd = new Date(today.getFullYear(), today.getMonth() - i + 1, 0);
        const period = `${monthDate.getFullYear()}-${String(monthDate.getMonth() + 1).padStart(2, '0')}`;

        // Get invoices for this month
        const invoices = await this.invoiceService.getInvoices({
          status: [InvoiceStatus.Paid],
          fromDate: monthDate,
          toDate: monthEnd
        });

        const spend = invoices.reduce((sum, inv) => sum + (inv.TotalAmount || 0), 0);

        // Get budget for this month (simplified - use monthly average of yearly budget)
        const budgets = await this.budgetService.getCurrentYearBudgets();
        const monthlyBudget = budgets.reduce((sum, b) => sum + (b.BudgetAmount || 0), 0) / 12;

        trend.push({
          period,
          spend,
          budget: monthlyBudget,
          variance: monthlyBudget - spend
        });
      }

      return trend;
    } catch (error) {
      logger.error('ProcurementService', 'Error getting spend trend:', error);
      return [];
    }
  }

  // ==================== JML Integration ====================

  public async processJMLProcurementRequest(request: IJMLProcurementRequest): Promise<IJMLProcurementResult> {
    try {
      return await this.requisitionService.createRequisitionFromJML(request);
    } catch (error) {
      logger.error('ProcurementService', 'Error processing JML procurement request:', error);
      return {
        processId: request.processId,
        status: 'Failed',
        message: `Failed to process procurement request: ${error instanceof Error ? error.message : 'Unknown error'}`
      };
    }
  }

  // ==================== Cross-Entity Operations ====================

  public async getVendorFullProfile(vendorId: number): Promise<{
    vendor: IVendor;
    contracts: IContract[];
    recentPOs: IPurchaseOrder[];
    pendingInvoices: IInvoice[];
    openIssues: { Id: number; Title: string; Status: string }[];
  }> {
    try {
      const [vendor, contracts, pos, invoices, issues] = await Promise.all([
        this.vendorService.getVendorById(vendorId),
        this.contractService.getVendorContracts(vendorId),
        this.purchaseOrderService.getPurchaseOrders({ vendorId }),
        this.invoiceService.getVendorInvoices(vendorId),
        this.vendorService.getVendorIssues(vendorId, true)
      ]);

      return {
        vendor,
        contracts,
        recentPOs: pos.slice(0, 10),
        pendingInvoices: invoices.filter(i =>
          i.Status !== InvoiceStatus.Paid && i.Status !== InvoiceStatus.Cancelled
        ),
        openIssues: issues.map(i => ({
          Id: i.Id!,
          Title: i.Title,
          Status: i.Status
        }))
      };
    } catch (error) {
      logger.error('ProcurementService', 'Error getting vendor full profile:', error);
      throw error;
    }
  }

  public async getPOFullDetails(poId: number): Promise<{
    po: IPurchaseOrder;
    vendor: IVendor;
    lineItems: { Id: number; Description: string; Quantity: number; UnitPrice: number; TotalPrice: number }[];
    invoices: IInvoice[];
    receipts: { Id: number; ReceiptNumber: string; ReceiptDate: Date }[];
  }> {
    try {
      const po = await this.purchaseOrderService.getPurchaseOrderById(poId);
      const [vendor, lineItems, invoices, receipts] = await Promise.all([
        this.vendorService.getVendorById(po.VendorId),
        this.purchaseOrderService.getPOLineItems(poId),
        this.invoiceService.getPOInvoices(poId),
        this.purchaseOrderService.getGoodsReceipts(poId)
      ]);

      return {
        po,
        vendor,
        lineItems: lineItems.map(li => ({
          Id: li.Id!,
          Description: li.Description,
          Quantity: li.Quantity,
          UnitPrice: li.UnitPrice,
          TotalPrice: li.TotalPrice
        })),
        invoices,
        receipts: receipts.map(r => ({
          Id: r.Id!,
          ReceiptNumber: r.ReceiptNumber,
          ReceiptDate: r.ReceiptDate
        }))
      };
    } catch (error) {
      logger.error('ProcurementService', 'Error getting PO full details:', error);
      throw error;
    }
  }

  // ==================== Quick Actions ====================

  public async convertRequisitionToPO(requisitionId: number): Promise<number> {
    try {
      return await this.purchaseOrderService.createPOFromRequisition(requisitionId);
    } catch (error) {
      logger.error('ProcurementService', 'Error converting requisition to PO:', error);
      throw error;
    }
  }

  public async quickApproveRequisition(requisitionId: number, approverId: number): Promise<void> {
    try {
      await this.requisitionService.approveRequisition(requisitionId, approverId);
    } catch (error) {
      logger.error('ProcurementService', 'Error quick approving requisition:', error);
      throw error;
    }
  }

  public async quickPayInvoice(invoiceId: number, paymentReference: string): Promise<void> {
    try {
      await this.invoiceService.recordPayment(invoiceId, paymentReference);
    } catch (error) {
      logger.error('ProcurementService', 'Error quick paying invoice:', error);
      throw error;
    }
  }

  // ==================== Search ====================

  public async globalSearch(searchTerm: string): Promise<{
    vendors: IVendor[];
    requisitions: IPurchaseRequisition[];
    purchaseOrders: IPurchaseOrder[];
    contracts: IContract[];
    invoices: IInvoice[];
  }> {
    try {
      const [vendors, requisitions, purchaseOrders, contracts, invoices] = await Promise.all([
        this.vendorService.getVendors({ searchTerm }),
        this.requisitionService.getRequisitions({ searchTerm }),
        this.purchaseOrderService.getPurchaseOrders({ searchTerm }),
        this.contractService.getContracts({ searchTerm }),
        this.invoiceService.getInvoices({ searchTerm })
      ]);

      return {
        vendors: vendors.slice(0, 10),
        requisitions: requisitions.slice(0, 10),
        purchaseOrders: purchaseOrders.slice(0, 10),
        contracts: contracts.slice(0, 10),
        invoices: invoices.slice(0, 10)
      };
    } catch (error) {
      logger.error('ProcurementService', 'Error in global search:', error);
      throw error;
    }
  }
}
