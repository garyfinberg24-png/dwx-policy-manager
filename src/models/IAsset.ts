// Asset Management Models
// Comprehensive interfaces for IT asset tracking and management

export enum AssetStatus {
  Available = 'Available',
  Assigned = 'Assigned',
  InMaintenance = 'In Maintenance',
  Reserved = 'Reserved',
  Retired = 'Retired',
  Lost = 'Lost',
  Damaged = 'Damaged',
  Disposed = 'Disposed'
}

export enum AssetCondition {
  New = 'New',
  Excellent = 'Excellent',
  Good = 'Good',
  Fair = 'Fair',
  Poor = 'Poor',
  Broken = 'Broken'
}

export enum AssetCategory {
  Hardware = 'Hardware',
  Software = 'Software',
  Peripheral = 'Peripheral',
  Network = 'Network',
  Mobile = 'Mobile',
  Furniture = 'Furniture',
  Other = 'Other'
}

export enum MaintenanceType {
  Preventive = 'Preventive',
  Corrective = 'Corrective',
  Upgrade = 'Upgrade',
  Cleaning = 'Cleaning',
  Calibration = 'Calibration'
}

export enum CheckoutStatus {
  CheckedOut = 'Checked Out',
  CheckedIn = 'Checked In',
  Overdue = 'Overdue',
  Cancelled = 'Cancelled'
}

export enum DepreciationMethod {
  StraightLine = 'Straight Line',
  DecliningBalance = 'Declining Balance',
  UnitsOfProduction = 'Units of Production',
  None = 'None'
}

export enum M365LicenseType {
  Microsoft365E3 = 'Microsoft 365 E3',
  Microsoft365E5 = 'Microsoft 365 E5',
  Microsoft365F3 = 'Microsoft 365 F3',
  Microsoft365BusinessBasic = 'Microsoft 365 Business Basic',
  Microsoft365BusinessStandard = 'Microsoft 365 Business Standard',
  Microsoft365BusinessPremium = 'Microsoft 365 Business Premium',
  Office365E3 = 'Office 365 E3',
  Office365E5 = 'Office 365 E5',
  EnterpriseE3 = 'Enterprise E3',
  EnterpriseE5 = 'Enterprise E5',
  PowerBIPro = 'Power BI Pro',
  PowerBIPremium = 'Power BI Premium',
  ProjectPlan3 = 'Project Plan 3',
  ProjectPlan5 = 'Project Plan 5',
  VisioPlan2 = 'Visio Plan 2',
  DefenderForOffice365 = 'Defender for Office 365',
  AzureADP1 = 'Azure AD Premium P1',
  AzureADP2 = 'Azure AD Premium P2',
  ExchangeOnlinePlan1 = 'Exchange Online Plan 1',
  ExchangeOnlinePlan2 = 'Exchange Online Plan 2',
  SharePointOnlinePlan1 = 'SharePoint Online Plan 1',
  SharePointOnlinePlan2 = 'SharePoint Online Plan 2',
  TeamsPremium = 'Teams Premium',
  Intune = 'Microsoft Intune',
  Other = 'Other'
}

export enum M365SubscriptionType {
  User = 'User',
  Device = 'Device',
  SiteBased = 'Site-Based',
  Pooled = 'Pooled',
  SharedComputer = 'Shared Computer Activation'
}

// Main Asset Interface
export interface IAsset {
  Id?: number;
  Title: string;
  AssetTag: string; // Unique identifier (e.g., LAP-001, MON-045)
  Barcode?: string; // Barcode or QR code for scanning
  SerialNumber?: string;
  AssetTypeId: number;
  AssetType?: IAssetType; // Lookup expansion
  Category: AssetCategory;
  Status: AssetStatus;
  Condition: AssetCondition;

  // Assignment Information
  AssignedToId?: number;
  AssignedTo?: any; // User lookup
  AssignedDate?: Date;
  AssignedById?: number;
  AssignedBy?: any; // User lookup

  // Location Information
  Location?: string; // Building/Floor/Room
  Department?: string;
  CostCenter?: string;

  // Financial Information
  PurchaseDate?: Date;
  PurchaseCost?: number;
  CurrentValue?: number;
  DepreciationMethod?: DepreciationMethod;
  DepreciationRate?: number; // Percentage per year
  SalvageValue?: number;
  WarrantyExpiration?: Date;

  // Vendor Information
  Vendor?: string;
  PurchaseOrderNumber?: string;
  InvoiceNumber?: string;

  // Maintenance Information
  LastMaintenanceDate?: Date;
  NextMaintenanceDate?: Date;
  MaintenanceSchedule?: string; // e.g., "Every 6 months"

  // Software-Specific
  LicenseKey?: string;
  LicenseExpiration?: Date;
  LicenseType?: string; // e.g., "Per User", "Site License"
  MaxLicenses?: number;
  CurrentLicensesUsed?: number;

  // Hardware-Specific
  Manufacturer?: string;
  Model?: string;
  Specifications?: string; // JSON string with specs
  IPAddress?: string;
  MACAddress?: string;
  HostName?: string;

  // Lifecycle
  RetirementDate?: Date;
  RetirementReason?: string;
  DisposalDate?: Date;
  DisposalMethod?: string;

  // Metadata
  Notes?: string;
  Attachments?: string; // JSON array of attachment URLs
  Created?: Date;
  CreatedById?: number;
  Modified?: Date;
  ModifiedById?: number;
}

// Asset Type Interface
export interface IAssetType {
  Id?: number;
  Title: string; // e.g., "Laptop", "Monitor", "Microsoft Office 365"
  Category: AssetCategory;
  Description?: string;
  Manufacturer?: string;
  Model?: string;

  // Default specifications
  DefaultSpecs?: string; // JSON string
  DefaultDepreciationMethod?: DepreciationMethod;
  DefaultDepreciationRate?: number;
  DefaultWarrantyPeriod?: number; // Months
  DefaultMaintenanceSchedule?: string;

  // Lifecycle
  IsActive: boolean;
  Created?: Date;
  Modified?: Date;
}

// Asset Assignment Interface
export interface IAssetAssignment {
  Id?: number;
  AssetId: number;
  Asset?: IAsset;
  AssignedToId: number;
  AssignedTo?: any; // User lookup
  AssignedById: number;
  AssignedBy?: any; // User lookup
  AssignedDate: Date;
  ExpectedReturnDate?: Date;
  ActualReturnDate?: Date;

  // Assignment Context
  ProcessId?: number; // Link to JML process if assigned during onboarding
  TaskId?: number; // Link to JML task
  AssignmentReason?: string;

  // Location tracking
  AssignedLocation?: string;

  // Status
  Status: CheckoutStatus;
  IsActive: boolean;

  // Return information
  ReturnCondition?: AssetCondition;
  ReturnNotes?: string;

  // Metadata
  Notes?: string;
  Created?: Date;
  Modified?: Date;
}

// Asset Checkout Interface (for temporary assignments)
export interface IAssetCheckout {
  Id?: number;
  AssetId: number;
  Asset?: IAsset;
  CheckedOutToId: number;
  CheckedOutTo?: any; // User lookup
  CheckedOutById: number;
  CheckedOutBy?: any; // User lookup
  CheckoutDate: Date;
  ExpectedReturnDate: Date;
  ActualReturnDate?: Date;

  // Purpose
  Purpose?: string;
  Location?: string; // Where the asset will be used

  // Status
  Status: CheckoutStatus;
  IsOverdue: boolean;

  // Return information
  CheckedInDate?: Date;
  CheckedInById?: number;
  CheckedInBy?: any;
  ReturnCondition?: AssetCondition;
  ReturnNotes?: string;

  // Notifications
  ReminderSent?: boolean;
  OverdueNotificationSent?: boolean;

  // Metadata
  Notes?: string;
  Created?: Date;
  Modified?: Date;
}

// Asset Maintenance Interface
export interface IAssetMaintenance {
  Id?: number;
  AssetId: number;
  Asset?: IAsset;
  MaintenanceType: MaintenanceType;
  ScheduledDate: Date;
  CompletedDate?: Date;

  // Maintenance Details
  PerformedById?: number;
  PerformedBy?: any; // User lookup
  Vendor?: string;
  Cost?: number;
  Description: string;

  // Status
  IsCompleted: boolean;
  IsCancelled: boolean;

  // Follow-up
  NextMaintenanceDate?: Date;

  // Parts and Labor
  PartsUsed?: string; // JSON array of parts
  LaborHours?: number;

  // Metadata
  Notes?: string;
  Attachments?: string; // JSON array of attachment URLs
  Created?: Date;
  Modified?: Date;
}

// Asset Transfer Interface
export interface IAssetTransfer {
  Id?: number;
  AssetId: number;
  Asset?: IAsset;

  // From
  FromUserId?: number;
  FromUser?: any;
  FromLocation?: string;
  FromDepartment?: string;

  // To
  ToUserId?: number;
  ToUser?: any;
  ToLocation?: string;
  ToDepartment?: string;

  // Transfer Details
  TransferDate: Date;
  RequestedById: number;
  RequestedBy?: any;
  ApprovedById?: number;
  ApprovedBy?: any;
  ApprovalDate?: Date;

  // Status
  Status: 'Pending' | 'Approved' | 'Rejected' | 'Completed' | 'Cancelled';

  // Reason
  TransferReason?: string;
  Notes?: string;

  // Metadata
  Created?: Date;
  Modified?: Date;
}

// Asset Audit Interface
export interface IAssetAudit {
  Id?: number;
  AuditName: string;
  AuditDate: Date;
  AuditedById: number;
  AuditedBy?: any;

  // Scope
  Department?: string;
  Location?: string;
  Category?: AssetCategory;

  // Results
  TotalAssetsAudited: number;
  AssetsFound: number;
  AssetsMissing: number;
  AssetsDiscrepancy: number;

  // Status
  Status: 'In Progress' | 'Completed' | 'Cancelled';
  CompletedDate?: Date;

  // Details
  Notes?: string;
  AuditResults?: string; // JSON array of audit line items

  // Metadata
  Created?: Date;
  Modified?: Date;
}

// Asset Audit Line Item
export interface IAssetAuditItem {
  Id?: number;
  AuditId: number;
  AssetId: number;
  Asset?: IAsset;

  // Expected vs Actual
  ExpectedLocation?: string;
  ActualLocation?: string;
  ExpectedCondition?: AssetCondition;
  ActualCondition?: AssetCondition;
  ExpectedAssignedToId?: number;
  ActualAssignedToId?: number;

  // Status
  Found: boolean;
  HasDiscrepancy: boolean;
  DiscrepancyNotes?: string;

  // Actions
  ActionRequired?: string;
  ActionTaken?: string;

  // Metadata
  Created?: Date;
  Modified?: Date;
}

// Asset Depreciation Schedule
export interface IAssetDepreciation {
  Id?: number;
  AssetId: number;
  Asset?: IAsset;
  PeriodDate: Date; // Month-end date

  // Values
  BookValue: number;
  DepreciationAmount: number;
  AccumulatedDepreciation: number;
  NetBookValue: number;

  // Status
  IsCalculated: boolean;
  CalculatedDate?: Date;

  // Metadata
  Created?: Date;
  Modified?: Date;
}

// Asset Request Interface (for new asset requests)
export interface IAssetRequest {
  Id?: number;
  RequestedById: number;
  RequestedBy?: any;
  RequestDate: Date;

  // Request Details
  AssetTypeId?: number;
  AssetType?: IAssetType;
  Category: AssetCategory;
  Description: string;
  Justification?: string;
  Quantity: number;

  // Priority and Timeline
  Priority: 'Low' | 'Medium' | 'High' | 'Urgent';
  RequiredByDate?: Date;

  // Financial
  EstimatedCost?: number;
  BudgetCode?: string;

  // Approval
  Status: 'Pending' | 'Approved' | 'Rejected' | 'Ordered' | 'Fulfilled' | 'Cancelled';
  ApprovedById?: number;
  ApprovedBy?: any;
  ApprovalDate?: Date;
  RejectionReason?: string;

  // Fulfillment
  FulfilledDate?: Date;
  FulfilledAssetIds?: string; // JSON array of asset IDs

  // Metadata
  Notes?: string;
  Created?: Date;
  Modified?: Date;
}

// Asset Dashboard Statistics
export interface IAssetStatistics {
  totalAssets: number;
  totalValue: number;
  byStatus: {
    [key in AssetStatus]?: number;
  };
  byCategory: {
    [key in AssetCategory]?: number;
  };
  byCondition: {
    [key in AssetCondition]?: number;
  };
  assignedAssets: number;
  availableAssets: number;
  inMaintenanceAssets: number;
  overdueCheckouts: number;
  expiringSoonLicenses: number; // Within 30 days
  upcomingMaintenance: number; // Within 30 days
  recentActivity: {
    newAssignments: number;
    recentCheckouts: number;
    completedMaintenance: number;
  };
}

// Asset Filter Criteria
export interface IAssetFilterCriteria {
  status?: AssetStatus[];
  category?: AssetCategory[];
  condition?: AssetCondition[];
  assignedToId?: number;
  location?: string;
  department?: string;
  assetTypeId?: number;
  manufacturer?: string;
  searchTerm?: string;
  fromDate?: Date;
  toDate?: Date;
  isAvailable?: boolean;
  hasWarranty?: boolean;
  needsMaintenance?: boolean;
}

// Asset History Entry
export interface IAssetHistoryEntry {
  id: string;
  timestamp: Date;
  action: 'Created' | 'Assigned' | 'Unassigned' | 'CheckedOut' | 'CheckedIn' |
          'Maintenance' | 'Transfer' | 'StatusChanged' | 'ConditionChanged' | 'Updated';
  performedById: number;
  performedBy: string;
  description: string;
  details?: any; // Additional context
}

// Integration with JML Process
export interface IJMLAssetAssignment {
  processId: number;
  employeeId: number;
  assetIds: number[];
  assignmentDate: Date;
  assignedById: number;
}

// Bulk Operations
export interface IAssetBulkOperation {
  operation: 'Assign' | 'Unassign' | 'ChangeStatus' | 'ChangeLocation' | 'Retire' | 'Delete';
  assetIds: number[];
  parameters?: {
    assignToId?: number;
    status?: AssetStatus;
    location?: string;
    retirementDate?: Date;
    retirementReason?: string;
  };
}

// ==================== M365 License Management ====================

// M365 License Interface
export interface IM365License {
  Id?: number;
  Title: string;
  LicenseType: M365LicenseType;
  SubscriptionType: M365SubscriptionType;

  // License Pool Information
  TotalLicenses: number;
  AssignedLicenses: number;
  AvailableLicenses: number; // Calculated: Total - Assigned

  // Subscription Details
  SubscriptionId?: string; // Microsoft 365 Subscription ID
  SkuId?: string; // Microsoft SKU ID
  SkuPartNumber?: string; // e.g., "ENTERPRISEPACK"

  // Dates
  PurchaseDate?: Date;
  StartDate?: Date;
  ExpirationDate?: Date;
  RenewalDate?: Date;
  AutoRenew?: boolean;

  // Financial
  CostPerLicense?: number;
  BillingPeriod?: 'Monthly' | 'Annual';
  TotalCost?: number; // Calculated: TotalLicenses * CostPerLicense
  NextBillingDate?: Date;

  // Vendor/Reseller
  Vendor?: string;
  ResellerContact?: string;
  ContractNumber?: string;
  PurchaseOrderNumber?: string;

  // Management
  TenantId?: string;
  AdminContact?: any; // User lookup
  AdminContactId?: number;
  Department?: string; // Department responsible for licenses
  CostCenter?: string;

  // Status
  IsActive: boolean;
  IsExpiringSoon?: boolean; // Within 90 days
  HasUnusedLicenses?: boolean;

  // Services Included
  ServicesIncluded?: string; // JSON string of included services
  AddOns?: string; // JSON string of add-on licenses

  // Compliance
  ComplianceNotes?: string;
  AuditDate?: Date;

  // Metadata
  Notes?: string;
  Created?: Date;
  CreatedById?: number;
  Modified?: Date;
  ModifiedById?: number;
}

// M365 License Assignment
export interface IM365LicenseAssignment {
  Id?: number;
  LicenseId: number;
  License?: IM365License;
  AssignedToId: number;
  AssignedTo?: any; // User lookup
  AssignedById: number;
  AssignedBy?: any; // User lookup
  AssignedDate: Date;
  UnassignedDate?: Date;

  // Assignment Details
  ProcessId?: number; // Link to JML process
  AssignmentReason?: string;
  Department?: string;

  // Status
  IsActive: boolean;
  Status: 'Active' | 'Suspended' | 'Removed';

  // Usage Tracking
  LastUsedDate?: Date;
  DaysUnused?: number;
  UsagePercentage?: number; // 0-100

  // Services Enabled
  EnabledServices?: string; // JSON string of enabled services

  // Metadata
  Notes?: string;
  Created?: Date;
  Modified?: Date;
}

// M365 License Usage Report
export interface IM365LicenseUsageReport {
  LicenseType: M365LicenseType;
  TotalLicenses: number;
  AssignedLicenses: number;
  UnassignedLicenses: number;
  UtilizationRate: number; // Percentage
  InactiveLicenses: number; // Assigned but not used in 30+ days
  CostPerMonth: number;
  WastedCost: number; // Cost of inactive licenses
  RecommendedAction?: string;
}

// M365 License Statistics
export interface IM365LicenseStatistics {
  totalLicenses: number;
  totalAssigned: number;
  totalAvailable: number;
  totalCost: number;
  monthlyRecurringCost: number;
  annualRecurringCost: number;
  utilizationRate: number; // Percentage
  byLicenseType: {
    [key in M365LicenseType]?: {
      total: number;
      assigned: number;
      available: number;
      cost: number;
    };
  };
  byDepartment: {
    [department: string]: {
      total: number;
      cost: number;
    };
  };
  expiringSoon: number; // Within 90 days
  upForRenewal: number; // Within 30 days
  inactiveLicenses: number; // Not used in 30+ days
  recentActivity: {
    newAssignments: number;
    removedAssignments: number;
    renewals: number;
  };
}

// M365 License Optimization Recommendation
export interface IM365LicenseOptimization {
  licenseType: M365LicenseType;
  currentCount: number;
  recommendedCount: number;
  potentialSavings: number;
  reason: string;
  inactiveUsers: string[]; // User names with inactive licenses
  priority: 'High' | 'Medium' | 'Low';
}

// M365 License Filter Criteria
export interface IM365LicenseFilterCriteria {
  licenseType?: M365LicenseType[];
  subscriptionType?: M365SubscriptionType[];
  department?: string;
  isActive?: boolean;
  isExpiringSoon?: boolean;
  hasUnusedLicenses?: boolean;
  assignedToId?: number;
  searchTerm?: string;
}

// M365 License Renewal Request
export interface IM365LicenseRenewal {
  Id?: number;
  LicenseId: number;
  License?: IM365License;
  RequestedById: number;
  RequestedBy?: any;
  RequestDate: Date;
  RenewalDate: Date;
  QuantityRequested: number;
  EstimatedCost: number;
  Justification?: string;
  Status: 'Pending' | 'Approved' | 'Rejected' | 'Completed';
  ApprovedById?: number;
  ApprovedBy?: any;
  ApprovalDate?: Date;
  Notes?: string;
  Created?: Date;
  Modified?: Date;
}
