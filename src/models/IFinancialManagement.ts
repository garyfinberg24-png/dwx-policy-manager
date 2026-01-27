// Financial Management Models
// Interfaces for expenses, payroll summaries, and financial dashboard

import { IBaseListItem, IUser } from './ICommon';
import { Currency } from './IProcurement';

// ==================== ENUMS ====================

export enum ExpenseStatus {
  Draft = 'Draft',
  Submitted = 'Submitted',
  PendingApproval = 'Pending Approval',
  Approved = 'Approved',
  Rejected = 'Rejected',
  Paid = 'Paid',
  Cancelled = 'Cancelled'
}

export enum ExpenseCategory {
  Travel = 'Travel',
  Equipment = 'Equipment',
  Software = 'Software',
  Training = 'Training',
  Supplies = 'Supplies',
  ProfessionalServices = 'Professional Services',
  MealsEntertainment = 'Meals & Entertainment',
  Telecommunications = 'Telecommunications',
  Other = 'Other'
}

export enum PayrollStatus {
  Draft = 'Draft',
  Processing = 'Processing',
  Processed = 'Processed',
  Finalized = 'Finalized',
  Cancelled = 'Cancelled'
}

// ==================== EXPENSE INTERFACES ====================

export interface IExpense extends IBaseListItem {
  ExpenseCode: string;
  ExpenseDate: Date;
  Amount: number;
  Category: ExpenseCategory;
  CostCenter: string;
  Department: string;
  Status: ExpenseStatus;
  Currency: Currency;
  Vendor?: string;
  Notes?: string;

  // Workflow dates
  SubmittedDate?: Date;
  ApprovalDate?: Date;
  PaymentDate?: Date;
  PaymentReference?: string;

  // Submitter
  SubmittedById?: number;
  SubmittedBy?: IUser;

  // Approver
  ApprovedById?: number;
  ApprovedBy?: IUser;
  RejectionReason?: string;

  // Attachments
  ReceiptUrl?: string;
  Attachments?: string; // JSON array

  // JML Integration
  ProcessId?: number;
  EmployeeId?: number;
}

export interface IExpenseFilter {
  searchTerm?: string;
  status?: ExpenseStatus[];
  category?: ExpenseCategory[];
  department?: string;
  costCenter?: string;
  submittedById?: number;
  fromDate?: Date;
  toDate?: Date;
  minAmount?: number;
  maxAmount?: number;
}

// ==================== PAYROLL INTERFACES ====================

export interface IPayrollSummary extends IBaseListItem {
  PayrollCode: string;
  PeriodStart: Date;
  PeriodEnd: Date;
  FiscalYear: number;
  FiscalMonth: number;
  CostCenter: string;
  Department: string;
  Status: PayrollStatus;
  Currency: Currency;

  // Headcount
  HeadCount: number;

  // Pay breakdown
  GrossPay: number;
  Deductions: number;
  NetPay: number;
  EmployerContributions: number;
  TotalCost: number;

  // Budget comparison
  BudgetAmount: number;
  Variance: number;
  VariancePercent: number;

  // Additional components
  OvertimeHours?: number;
  OvertimeCost?: number;
  BonusPayments?: number;

  // Tax breakdown
  TaxWithheld?: number;
  PensionContributions?: number;
  NIContributions?: number;

  // Processing
  ProcessedDate?: Date;
  ProcessedById?: number;
  ProcessedBy?: IUser;

  Notes?: string;
}

export interface IPayrollFilter {
  department?: string;
  costCenter?: string;
  fiscalYear?: number;
  fiscalMonth?: number;
  status?: PayrollStatus[];
  fromDate?: Date;
  toDate?: Date;
}

// ==================== COST CENTER INTERFACES ====================

export interface ICostCenter extends IBaseListItem {
  CostCenterCode: string;
  Department: string;
  BudgetOwnerId?: number;
  BudgetOwner?: IUser;
  IsActive: boolean;
  Description?: string;

  // Budget allocation
  AnnualBudget?: number;
  AllocatedBudget?: number;
  SpentToDate?: number;
  Currency: Currency;
}

// ==================== FINANCIAL DASHBOARD INTERFACES ====================

export interface IFinancialDashboardData {
  // Budget Overview
  totalBudget: number;
  totalSpent: number;
  totalRemaining: number;
  budgetUtilization: number; // percentage

  // Expense Summary
  totalExpenses: number;
  pendingExpenseApprovals: number;
  expensesByCategory: { [key in ExpenseCategory]?: number };
  expensesByStatus: { [key in ExpenseStatus]?: number };
  expensesByDepartment: { [department: string]: number };

  // Payroll Summary
  totalPayrollCost: number;
  totalHeadCount: number;
  avgCostPerEmployee: number;
  payrollVariance: number;

  // Invoice Summary (from existing)
  totalOutstandingInvoices: number;
  overdueInvoices: number;
  totalInvoiceValue: number;

  // Trends
  monthlySpendTrend: IMonthlySpendItem[];
  budgetVsActual: IBudgetVsActualItem[];
}

export interface IMonthlySpendItem {
  month: string; // e.g., "2024-01"
  expenses: number;
  payroll: number;
  invoices: number;
  total: number;
}

export interface IBudgetVsActualItem {
  department: string;
  budget: number;
  actual: number;
  variance: number;
  variancePercent: number;
}

export interface ICostCenterAnalysis {
  costCenter: string;
  costCenterName: string;
  department: string;
  budget: number;
  spent: number;
  remaining: number;
  utilization: number;
  variance: number;
  variancePercent: number;
  status: 'Under' | 'OnTrack' | 'Warning' | 'Over';
}

// ==================== FINANCIAL ALERTS ====================

export interface IFinancialAlert {
  id: string;
  type: 'Budget' | 'Expense' | 'Payroll' | 'Invoice';
  severity: 'Info' | 'Warning' | 'Critical';
  title: string;
  message: string;
  department?: string;
  costCenter?: string;
  amount?: number;
  threshold?: number;
  date: Date;
  actionUrl?: string;
}

// ==================== FINANCIAL KPI INTERFACES ====================

export interface IFinancialKPIs {
  // Budget Health
  budgetUtilization: number;
  budgetVariance: number;
  daysUntilFiscalYearEnd: number;
  projectedYearEndBalance: number;

  // Expense Metrics
  avgExpenseProcessingDays: number;
  expenseApprovalRate: number;
  topExpenseCategory: ExpenseCategory;
  expenseGrowthRate: number; // month-over-month

  // Payroll Metrics
  payrollAsPercentOfRevenue?: number;
  avgCostPerHeadcount: number;
  payrollGrowthRate: number;
  overtimePercentage: number;

  // Efficiency
  invoicePaymentDays: number;
  expenseReimbursementDays: number;
  budgetForecastAccuracy: number;
}

// ==================== EXPORT INTERFACES ====================

export interface IFinancialExportOptions {
  format: 'PDF' | 'Excel' | 'CSV';
  reportType: 'Expenses' | 'Payroll' | 'Budget' | 'Combined';
  dateRange: {
    from: Date;
    to: Date;
  };
  departments?: string[];
  costCenters?: string[];
  includeCharts?: boolean;
  includeDetails?: boolean;
}
