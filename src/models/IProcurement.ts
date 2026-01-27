// Procurement Management Models
// Comprehensive interfaces for vendor management, purchase requisitions, POs, contracts, and invoices

import { IBaseListItem, IUser } from './ICommon';

// ==================== ENUMS ====================

export enum VendorStatus {
  Active = 'Active',
  Inactive = 'Inactive',
  PendingApproval = 'Pending Approval',
  Blacklisted = 'Blacklisted',
  Suspended = 'Suspended'
}

export enum VendorType {
  Supplier = 'Supplier',
  ServiceProvider = 'Service Provider',
  Contractor = 'Contractor',
  Consultant = 'Consultant',
  Distributor = 'Distributor',
  Manufacturer = 'Manufacturer'
}

export enum VendorCategory {
  ITHardware = 'IT Hardware',
  ITSoftware = 'IT Software',
  ITServices = 'IT Services',
  OfficeSupplies = 'Office Supplies',
  Furniture = 'Furniture',
  ProfessionalServices = 'Professional Services',
  Utilities = 'Utilities',
  Marketing = 'Marketing',
  Travel = 'Travel',
  Facilities = 'Facilities',
  Telecommunications = 'Telecommunications',
  Security = 'Security',
  Training = 'Training',
  Catering = 'Catering',
  Other = 'Other'
}

export enum RequisitionStatus {
  Draft = 'Draft',
  Submitted = 'Submitted',
  PendingApproval = 'Pending Approval',
  Approved = 'Approved',
  Rejected = 'Rejected',
  ConvertedToPO = 'Converted to PO',
  Cancelled = 'Cancelled',
  OnHold = 'On Hold'
}

export enum RequisitionPriority {
  Low = 'Low',
  Medium = 'Medium',
  High = 'High',
  Urgent = 'Urgent',
  Critical = 'Critical'
}

export enum POStatus {
  Draft = 'Draft',
  PendingApproval = 'Pending Approval',
  Approved = 'Approved',
  Sent = 'Sent',
  Acknowledged = 'Acknowledged',
  PartiallyReceived = 'Partially Received',
  Received = 'Received',
  Closed = 'Closed',
  Cancelled = 'Cancelled',
  Disputed = 'Disputed'
}

export enum ContractStatus {
  Draft = 'Draft',
  UnderReview = 'Under Review',
  PendingSignature = 'Pending Signature',
  Active = 'Active',
  Expired = 'Expired',
  Terminated = 'Terminated',
  Renewed = 'Renewed',
  OnHold = 'On Hold'
}

export enum ContractType {
  MasterAgreement = 'Master Agreement',
  StatementOfWork = 'Statement of Work',
  NDA = 'NDA',
  SLA = 'SLA',
  MaintenanceAgreement = 'Maintenance Agreement',
  Subscription = 'Subscription',
  License = 'License Agreement',
  LeaseAgreement = 'Lease Agreement',
  ServiceContract = 'Service Contract',
  PurchaseAgreement = 'Purchase Agreement',
  Other = 'Other'
}

export enum InvoiceStatus {
  Received = 'Received',
  PendingMatch = 'Pending Match',
  Matched = 'Matched',
  PartialMatch = 'Partial Match',
  MatchException = 'Match Exception',
  PendingApproval = 'Pending Approval',
  Approved = 'Approved',
  Disputed = 'Disputed',
  Scheduled = 'Scheduled for Payment',
  Paid = 'Paid',
  Cancelled = 'Cancelled'
}

export enum PaymentTerms {
  Immediate = 'Immediate',
  Net7 = 'Net 7',
  Net15 = 'Net 15',
  Net30 = 'Net 30',
  Net45 = 'Net 45',
  Net60 = 'Net 60',
  Net90 = 'Net 90',
  DueOnReceipt = 'Due on Receipt',
  Prepaid = 'Prepaid',
  EndOfMonth = 'End of Month',
  Custom = 'Custom'
}

export enum ReceiptStatus {
  Pending = 'Pending',
  Partial = 'Partial',
  Complete = 'Complete',
  Rejected = 'Rejected'
}

export enum UnitOfMeasure {
  Each = 'Each',
  Box = 'Box',
  Pack = 'Pack',
  Case = 'Case',
  License = 'License',
  Subscription = 'Subscription',
  Hour = 'Hour',
  Day = 'Day',
  Week = 'Week',
  Month = 'Month',
  Year = 'Year',
  Project = 'Project',
  Service = 'Service',
  Kilogram = 'Kilogram',
  Liter = 'Liter',
  Meter = 'Meter',
  SquareMeter = 'Square Meter',
  Other = 'Other'
}

export enum Currency {
  USD = 'USD',
  EUR = 'EUR',
  GBP = 'GBP',
  CAD = 'CAD',
  AUD = 'AUD',
  JPY = 'JPY',
  CHF = 'CHF',
  CNY = 'CNY',
  INR = 'INR',
  ZAR = 'ZAR',
  AED = 'AED',
  SGD = 'SGD',
  NZD = 'NZD',
  MXN = 'MXN',
  BRL = 'BRL',
  // Additional major world currencies
  HKD = 'HKD',
  SEK = 'SEK',
  NOK = 'NOK',
  DKK = 'DKK',
  PLN = 'PLN',
  THB = 'THB',
  MYR = 'MYR',
  IDR = 'IDR',
  PHP = 'PHP',
  KRW = 'KRW',
  TWD = 'TWD',
  TRY = 'TRY',
  RUB = 'RUB',
  SAR = 'SAR',
  ILS = 'ILS',
  EGP = 'EGP',
  NGN = 'NGN',
  KES = 'KES',
  CLP = 'CLP',
  COP = 'COP',
  ARS = 'ARS',
  PEN = 'PEN'
}

/**
 * Currency metadata interface for comprehensive currency configuration
 */
export interface ICurrencyInfo {
  /** ISO 4217 currency code */
  code: Currency;
  /** Currency symbol (e.g., R, $, €) */
  symbol: string;
  /** Full currency name */
  name: string;
  /** Number of decimal places (0 for JPY/KRW, 2 for most) */
  decimalPlaces: number;
  /** Symbol position: 'before' or 'after' the amount */
  symbolPosition: 'before' | 'after';
  /** Whether to include space between symbol and amount */
  symbolSpace: boolean;
  /** Country/region associated with this currency */
  country: string;
  /** Flag emoji for visual display */
  flag: string;
}

/**
 * Comprehensive currency data for all supported currencies
 * Default currency for JML: South African Rand (ZAR)
 */
export const CURRENCY_DATA: Record<Currency, ICurrencyInfo> = {
  // ============================================
  // AFRICA
  // ============================================
  [Currency.ZAR]: {
    code: Currency.ZAR,
    symbol: 'R',
    name: 'South African Rand',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'South Africa',
    flag: '🇿🇦'
  },
  [Currency.NGN]: {
    code: Currency.NGN,
    symbol: '₦',
    name: 'Nigerian Naira',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Nigeria',
    flag: '🇳🇬'
  },
  [Currency.KES]: {
    code: Currency.KES,
    symbol: 'KSh',
    name: 'Kenyan Shilling',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Kenya',
    flag: '🇰🇪'
  },
  [Currency.EGP]: {
    code: Currency.EGP,
    symbol: 'E£',
    name: 'Egyptian Pound',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Egypt',
    flag: '🇪🇬'
  },

  // ============================================
  // AMERICAS
  // ============================================
  [Currency.USD]: {
    code: Currency.USD,
    symbol: '$',
    name: 'US Dollar',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'United States',
    flag: '🇺🇸'
  },
  [Currency.CAD]: {
    code: Currency.CAD,
    symbol: 'C$',
    name: 'Canadian Dollar',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Canada',
    flag: '🇨🇦'
  },
  [Currency.MXN]: {
    code: Currency.MXN,
    symbol: '$',
    name: 'Mexican Peso',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Mexico',
    flag: '🇲🇽'
  },
  [Currency.BRL]: {
    code: Currency.BRL,
    symbol: 'R$',
    name: 'Brazilian Real',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Brazil',
    flag: '🇧🇷'
  },
  [Currency.ARS]: {
    code: Currency.ARS,
    symbol: '$',
    name: 'Argentine Peso',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Argentina',
    flag: '🇦🇷'
  },
  [Currency.CLP]: {
    code: Currency.CLP,
    symbol: '$',
    name: 'Chilean Peso',
    decimalPlaces: 0,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Chile',
    flag: '🇨🇱'
  },
  [Currency.COP]: {
    code: Currency.COP,
    symbol: '$',
    name: 'Colombian Peso',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Colombia',
    flag: '🇨🇴'
  },
  [Currency.PEN]: {
    code: Currency.PEN,
    symbol: 'S/',
    name: 'Peruvian Sol',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Peru',
    flag: '🇵🇪'
  },

  // ============================================
  // EUROPE
  // ============================================
  [Currency.EUR]: {
    code: Currency.EUR,
    symbol: '€',
    name: 'Euro',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'European Union',
    flag: '🇪🇺'
  },
  [Currency.GBP]: {
    code: Currency.GBP,
    symbol: '£',
    name: 'British Pound',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'United Kingdom',
    flag: '🇬🇧'
  },
  [Currency.CHF]: {
    code: Currency.CHF,
    symbol: 'CHF',
    name: 'Swiss Franc',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: true,
    country: 'Switzerland',
    flag: '🇨🇭'
  },
  [Currency.SEK]: {
    code: Currency.SEK,
    symbol: 'kr',
    name: 'Swedish Krona',
    decimalPlaces: 2,
    symbolPosition: 'after',
    symbolSpace: true,
    country: 'Sweden',
    flag: '🇸🇪'
  },
  [Currency.NOK]: {
    code: Currency.NOK,
    symbol: 'kr',
    name: 'Norwegian Krone',
    decimalPlaces: 2,
    symbolPosition: 'after',
    symbolSpace: true,
    country: 'Norway',
    flag: '🇳🇴'
  },
  [Currency.DKK]: {
    code: Currency.DKK,
    symbol: 'kr',
    name: 'Danish Krone',
    decimalPlaces: 2,
    symbolPosition: 'after',
    symbolSpace: true,
    country: 'Denmark',
    flag: '🇩🇰'
  },
  [Currency.PLN]: {
    code: Currency.PLN,
    symbol: 'zł',
    name: 'Polish Zloty',
    decimalPlaces: 2,
    symbolPosition: 'after',
    symbolSpace: true,
    country: 'Poland',
    flag: '🇵🇱'
  },
  [Currency.TRY]: {
    code: Currency.TRY,
    symbol: '₺',
    name: 'Turkish Lira',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Turkey',
    flag: '🇹🇷'
  },
  [Currency.RUB]: {
    code: Currency.RUB,
    symbol: '₽',
    name: 'Russian Ruble',
    decimalPlaces: 2,
    symbolPosition: 'after',
    symbolSpace: true,
    country: 'Russia',
    flag: '🇷🇺'
  },

  // ============================================
  // MIDDLE EAST
  // ============================================
  [Currency.AED]: {
    code: Currency.AED,
    symbol: 'د.إ',
    name: 'UAE Dirham',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: true,
    country: 'United Arab Emirates',
    flag: '🇦🇪'
  },
  [Currency.SAR]: {
    code: Currency.SAR,
    symbol: '﷼',
    name: 'Saudi Riyal',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: true,
    country: 'Saudi Arabia',
    flag: '🇸🇦'
  },
  [Currency.ILS]: {
    code: Currency.ILS,
    symbol: '₪',
    name: 'Israeli Shekel',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Israel',
    flag: '🇮🇱'
  },

  // ============================================
  // ASIA-PACIFIC
  // ============================================
  [Currency.JPY]: {
    code: Currency.JPY,
    symbol: '¥',
    name: 'Japanese Yen',
    decimalPlaces: 0,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Japan',
    flag: '🇯🇵'
  },
  [Currency.CNY]: {
    code: Currency.CNY,
    symbol: '¥',
    name: 'Chinese Yuan',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'China',
    flag: '🇨🇳'
  },
  [Currency.HKD]: {
    code: Currency.HKD,
    symbol: 'HK$',
    name: 'Hong Kong Dollar',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Hong Kong',
    flag: '🇭🇰'
  },
  [Currency.TWD]: {
    code: Currency.TWD,
    symbol: 'NT$',
    name: 'Taiwan Dollar',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Taiwan',
    flag: '🇹🇼'
  },
  [Currency.KRW]: {
    code: Currency.KRW,
    symbol: '₩',
    name: 'South Korean Won',
    decimalPlaces: 0,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'South Korea',
    flag: '🇰🇷'
  },
  [Currency.INR]: {
    code: Currency.INR,
    symbol: '₹',
    name: 'Indian Rupee',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'India',
    flag: '🇮🇳'
  },
  [Currency.SGD]: {
    code: Currency.SGD,
    symbol: 'S$',
    name: 'Singapore Dollar',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Singapore',
    flag: '🇸🇬'
  },
  [Currency.MYR]: {
    code: Currency.MYR,
    symbol: 'RM',
    name: 'Malaysian Ringgit',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Malaysia',
    flag: '🇲🇾'
  },
  [Currency.THB]: {
    code: Currency.THB,
    symbol: '฿',
    name: 'Thai Baht',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Thailand',
    flag: '🇹🇭'
  },
  [Currency.IDR]: {
    code: Currency.IDR,
    symbol: 'Rp',
    name: 'Indonesian Rupiah',
    decimalPlaces: 0,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Indonesia',
    flag: '🇮🇩'
  },
  [Currency.PHP]: {
    code: Currency.PHP,
    symbol: '₱',
    name: 'Philippine Peso',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Philippines',
    flag: '🇵🇭'
  },

  // ============================================
  // OCEANIA
  // ============================================
  [Currency.AUD]: {
    code: Currency.AUD,
    symbol: 'A$',
    name: 'Australian Dollar',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'Australia',
    flag: '🇦🇺'
  },
  [Currency.NZD]: {
    code: Currency.NZD,
    symbol: 'NZ$',
    name: 'New Zealand Dollar',
    decimalPlaces: 2,
    symbolPosition: 'before',
    symbolSpace: false,
    country: 'New Zealand',
    flag: '🇳🇿'
  }
};

/**
 * Currency formatting settings interface
 */
export interface ICurrencySettings {
  /** Default currency code */
  defaultCurrency: Currency;
  /** Decimal separator character */
  decimalSeparator: '.' | ',';
  /** Thousands separator character */
  thousandsSeparator: ',' | '.' | ' ' | '';
  /** Override decimal places (null = use currency default) */
  decimalPlacesOverride: number | null;
  /** Symbol position override (null = use currency default) */
  symbolPositionOverride: 'before' | 'after' | null;
  /** Whether to show currency code after symbol (e.g., "R 100 ZAR") */
  showCurrencyCode: boolean;
  /** Negative number format */
  negativeFormat: 'minus' | 'parentheses' | 'minusAfter';
  /** List of enabled currencies (empty = all enabled) */
  enabledCurrencies: Currency[];
}

/**
 * Default currency settings for JML (South African focused)
 */
export const DEFAULT_CURRENCY_SETTINGS: ICurrencySettings = {
  defaultCurrency: Currency.ZAR,
  decimalSeparator: '.',
  thousandsSeparator: ',',
  decimalPlacesOverride: null,
  symbolPositionOverride: null,
  showCurrencyCode: false,
  negativeFormat: 'minus',
  enabledCurrencies: []
};

/**
 * Format a number as currency using the provided settings
 */
export function formatCurrency(
  amount: number,
  currencyCode: Currency = Currency.ZAR,
  settings: Partial<ICurrencySettings> = {}
): string {
  const currencyInfo = CURRENCY_DATA[currencyCode];
  const mergedSettings = { ...DEFAULT_CURRENCY_SETTINGS, ...settings };

  const decimalPlaces = mergedSettings.decimalPlacesOverride ?? currencyInfo.decimalPlaces;
  const symbolPosition = mergedSettings.symbolPositionOverride ?? currencyInfo.symbolPosition;

  // Handle negative numbers
  const isNegative = amount < 0;
  const absoluteAmount = Math.abs(amount);

  // Format the number
  const parts = absoluteAmount.toFixed(decimalPlaces).split('.');
  const integerPart = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, mergedSettings.thousandsSeparator);
  const decimalPart = parts[1];

  let formattedNumber = decimalPlaces > 0
    ? `${integerPart}${mergedSettings.decimalSeparator}${decimalPart}`
    : integerPart;

  // Add symbol
  const space = currencyInfo.symbolSpace ? ' ' : '';
  const currencyCodeSuffix = mergedSettings.showCurrencyCode ? ` ${currencyCode}` : '';

  let result: string;
  if (symbolPosition === 'before') {
    result = `${currencyInfo.symbol}${space}${formattedNumber}${currencyCodeSuffix}`;
  } else {
    result = `${formattedNumber}${space}${currencyInfo.symbol}${currencyCodeSuffix}`;
  }

  // Handle negative formatting
  if (isNegative) {
    switch (mergedSettings.negativeFormat) {
      case 'parentheses':
        result = `(${result})`;
        break;
      case 'minusAfter':
        result = `${result}-`;
        break;
      default: // 'minus'
        result = `-${result}`;
        break;
    }
  }

  return result;
}

/**
 * Major world currencies - the most commonly used currencies globally
 * Shown by default in currency dropdowns for a cleaner user experience
 */
export const MAJOR_CURRENCIES: Currency[] = [
  Currency.ZAR,  // South African Rand (JML default)
  Currency.USD,  // US Dollar
  Currency.EUR,  // Euro
  Currency.GBP,  // British Pound
  Currency.JPY,  // Japanese Yen
  Currency.CHF,  // Swiss Franc
  Currency.CAD,  // Canadian Dollar
  Currency.AUD,  // Australian Dollar
  Currency.CNY,  // Chinese Yuan
  Currency.INR,  // Indian Rupee
  Currency.AED,  // UAE Dirham
  Currency.SGD,  // Singapore Dollar
  Currency.HKD,  // Hong Kong Dollar
  Currency.NZD,  // New Zealand Dollar
  Currency.BRL   // Brazilian Real
];

/**
 * Get major currencies data for simplified dropdown
 */
export function getMajorCurrencies(): ICurrencyInfo[] {
  return MAJOR_CURRENCIES.map(code => CURRENCY_DATA[code]);
}

/**
 * Get currencies grouped by region for display in dropdowns
 */
export function getCurrenciesByRegion(): Record<string, ICurrencyInfo[]> {
  return {
    'Africa': [
      CURRENCY_DATA[Currency.ZAR],
      CURRENCY_DATA[Currency.NGN],
      CURRENCY_DATA[Currency.KES],
      CURRENCY_DATA[Currency.EGP]
    ],
    'Americas': [
      CURRENCY_DATA[Currency.USD],
      CURRENCY_DATA[Currency.CAD],
      CURRENCY_DATA[Currency.MXN],
      CURRENCY_DATA[Currency.BRL],
      CURRENCY_DATA[Currency.ARS],
      CURRENCY_DATA[Currency.CLP],
      CURRENCY_DATA[Currency.COP],
      CURRENCY_DATA[Currency.PEN]
    ],
    'Europe': [
      CURRENCY_DATA[Currency.EUR],
      CURRENCY_DATA[Currency.GBP],
      CURRENCY_DATA[Currency.CHF],
      CURRENCY_DATA[Currency.SEK],
      CURRENCY_DATA[Currency.NOK],
      CURRENCY_DATA[Currency.DKK],
      CURRENCY_DATA[Currency.PLN],
      CURRENCY_DATA[Currency.TRY],
      CURRENCY_DATA[Currency.RUB]
    ],
    'Middle East': [
      CURRENCY_DATA[Currency.AED],
      CURRENCY_DATA[Currency.SAR],
      CURRENCY_DATA[Currency.ILS]
    ],
    'Asia-Pacific': [
      CURRENCY_DATA[Currency.JPY],
      CURRENCY_DATA[Currency.CNY],
      CURRENCY_DATA[Currency.HKD],
      CURRENCY_DATA[Currency.TWD],
      CURRENCY_DATA[Currency.KRW],
      CURRENCY_DATA[Currency.INR],
      CURRENCY_DATA[Currency.SGD],
      CURRENCY_DATA[Currency.MYR],
      CURRENCY_DATA[Currency.THB],
      CURRENCY_DATA[Currency.IDR],
      CURRENCY_DATA[Currency.PHP]
    ],
    'Oceania': [
      CURRENCY_DATA[Currency.AUD],
      CURRENCY_DATA[Currency.NZD]
    ]
  };
}

export enum BudgetStatus {
  Active = 'Active',
  Frozen = 'Frozen',
  Closed = 'Closed',
  Pending = 'Pending'
}

export enum ApprovalAction {
  Approve = 'Approve',
  Reject = 'Reject',
  RequestInfo = 'Request Info',
  Delegate = 'Delegate',
  Escalate = 'Escalate'
}

export enum CatalogItemStatus {
  Active = 'Active',
  Inactive = 'Inactive',
  Discontinued = 'Discontinued',
  PendingApproval = 'Pending Approval'
}

export enum CatalogItemType {
  Product = 'Product',
  Service = 'Service',
  Asset = 'Asset',
  Consumable = 'Consumable',
  Software = 'Software'
}

// ==================== VENDOR INTERFACES ====================

export interface IVendor extends IBaseListItem {
  VendorCode: string;
  VendorType: VendorType;
  Category: VendorCategory;
  Status: VendorStatus;

  // Company Information
  LegalName?: string;
  TradingName?: string;
  TaxId?: string;
  DunsNumber?: string;
  CompanyRegistration?: string;
  Website?: string;

  // Address
  AddressLine1?: string;
  AddressLine2?: string;
  City?: string;
  State?: string;
  Country?: string;
  PostalCode?: string;

  // Primary Contact
  PrimaryContactId?: number;
  PrimaryContact?: IUser;
  PrimaryPhone?: string;
  PrimaryEmail?: string;

  // Financial Information
  PaymentTerms: PaymentTerms;
  Currency: Currency;
  BankName?: string;
  BankAccountNumber?: string;
  BankRoutingNumber?: string;
  BankSwiftCode?: string;

  // Vendor Management
  PreferredVendor: boolean;
  ApprovedDate?: Date;
  ApprovedById?: number;
  ApprovedBy?: IUser;
  LastReviewDate?: Date;
  NextReviewDate?: Date;

  // Performance
  Rating?: number;
  RatingCount?: number;
  TotalOrders?: number;
  TotalSpend?: number;

  // Compliance
  Certifications?: string; // JSON array
  InsuranceExpiry?: Date;
  ComplianceDocuments?: string; // JSON array

  // Metadata
  Notes?: string;
  Tags?: string; // JSON array
}

export interface IVendorContact extends IBaseListItem {
  VendorId: number;
  Vendor?: IVendor;
  FirstName: string;
  LastName: string;
  Email: string;
  Phone?: string;
  Mobile?: string;
  Role: string;
  Department?: string;
  IsPrimary: boolean;
  IsActive: boolean;
  Notes?: string;
}

export interface IVendorDocument {
  Id?: number;
  VendorId: number;
  DocumentType: string;
  DocumentName: string;
  DocumentUrl: string;
  ExpiryDate?: Date;
  UploadedById?: number;
  UploadedBy?: IUser;
  UploadedDate?: Date;
  Notes?: string;
}

// ==================== CATALOG INTERFACES ====================

export interface ICatalogItem extends IBaseListItem {
  ItemCode: string;
  Category: VendorCategory;
  SubCategory?: string;
  Description?: string;
  UnitOfMeasure: UnitOfMeasure;
  IsActive: boolean;

  // Pricing
  DefaultPrice?: number;
  Currency: Currency;
  MinOrderQuantity?: number;
  MaxOrderQuantity?: number;

  // Vendor
  PreferredVendorId?: number;
  PreferredVendor?: IVendor;

  // Lead Time
  LeadTimeDays?: number;

  // Specifications
  Specifications?: string;
  ImageUrl?: string;

  // Asset Integration
  CreateAssetOnReceipt: boolean;
  AssetCategory?: string;
  AssetTypeId?: number;

  // Metadata
  Notes?: string;
  Tags?: string; // JSON array
}

export interface IVendorPricing {
  Id?: number;
  VendorId: number;
  Vendor?: IVendor;
  CatalogItemId: number;
  CatalogItem?: ICatalogItem;
  UnitPrice: number;
  Currency: Currency;
  MinQuantity?: number;
  MaxQuantity?: number;
  EffectiveFrom: Date;
  EffectiveTo?: Date;
  ContractId?: number;
  Contract?: IContract;
  IsPreferred: boolean;
  Notes?: string;
}

// ==================== REQUISITION INTERFACES ====================

export interface IPurchaseRequisition extends IBaseListItem {
  RequisitionNumber: string;
  RequestedById: number;
  RequestedBy?: IUser;
  Department: string;
  CostCenter?: string;
  Status: RequisitionStatus;
  Priority: RequisitionPriority;

  // Dates
  RequestedDate: Date;
  RequiredByDate?: Date;
  ApprovedDate?: Date;

  // Financial
  TotalEstimatedCost: number;
  Currency: Currency;
  BudgetId?: number;
  Budget?: IBudget;

  // Vendor Suggestion
  SuggestedVendorId?: number;
  SuggestedVendor?: IVendor;

  // JML Integration
  ProcessId?: number;
  TaskId?: number;
  EmployeeId?: number;

  // Approval
  ApprovalStatus?: string;
  ApprovedById?: number;
  ApprovedBy?: IUser;
  RejectionReason?: string;

  // Conversion
  PurchaseOrderId?: number;
  PurchaseOrder?: IPurchaseOrder;
  ConvertedDate?: Date;

  // Justification
  Justification?: string;
  BusinessNeed?: string;

  // Metadata
  Attachments?: string; // JSON array
  Notes?: string;
}

export interface IRequisitionLineItem extends IBaseListItem {
  RequisitionId: number;
  Requisition?: IPurchaseRequisition;
  LineNumber: number;

  // Item Details
  CatalogItemId?: number;
  CatalogItem?: ICatalogItem;
  ItemCode?: string;
  Description: string;
  Category: VendorCategory;

  // Quantity & Pricing
  Quantity: number;
  UnitOfMeasure: UnitOfMeasure;
  EstimatedUnitCost: number;
  EstimatedTotalCost: number;
  Currency: Currency;

  // Vendor
  VendorId?: number;
  Vendor?: IVendor;

  // Specifications
  Specifications?: string;
  Notes?: string;
}

// ==================== PURCHASE ORDER INTERFACES ====================

export interface IPurchaseOrder extends IBaseListItem {
  PONumber: string;
  VendorId: number;
  Vendor?: IVendor;
  Status: POStatus;

  // Source
  RequisitionIds?: string; // JSON array

  // Dates
  OrderDate: Date;
  ExpectedDeliveryDate?: Date;
  ActualDeliveryDate?: Date;
  SentDate?: Date;
  AcknowledgedDate?: Date;
  ClosedDate?: Date;

  // Addresses
  ShipToAddress?: string;
  ShipToAttention?: string;
  BillToAddress?: string;
  BillToAttention?: string;

  // Financial
  Subtotal: number;
  TaxRate?: number;
  TaxAmount?: number;
  ShippingCost?: number;
  DiscountAmount?: number;
  TotalAmount: number;
  Currency: Currency;
  PaymentTerms: PaymentTerms;

  // Vendor Reference
  VendorQuoteNumber?: string;
  VendorReference?: string;

  // Approval
  ApprovedById?: number;
  ApprovedBy?: IUser;
  ApprovedDate?: Date;

  // Sent
  SentById?: number;
  SentBy?: IUser;
  SentMethod?: string;

  // Terms
  TermsAndConditions?: string;
  SpecialInstructions?: string;

  // Budget
  BudgetId?: number;
  Budget?: IBudget;
  CostCenter?: string;
  Department?: string;

  // JML Integration
  ProcessId?: number;
  TaskId?: number;

  // Metadata
  Attachments?: string; // JSON array
  Notes?: string;
}

export interface IPOLineItem extends IBaseListItem {
  PurchaseOrderId: number;
  PurchaseOrder?: IPurchaseOrder;
  LineNumber: number;

  // Item Details
  CatalogItemId?: number;
  CatalogItem?: ICatalogItem;
  ItemCode?: string;
  Description: string;

  // Quantity & Pricing
  Quantity: number;
  UnitOfMeasure: UnitOfMeasure;
  UnitPrice: number;
  TotalPrice: number;
  TaxRate?: number;
  TaxAmount?: number;

  // Receipt Tracking
  QuantityReceived: number;
  QuantityPending: number;
  QuantityRejected: number;
  ReceivedStatus: ReceiptStatus;

  // Asset Creation
  AssetIds?: string; // JSON array of created asset IDs

  // Specifications
  Specifications?: string;
  DeliveryDate?: Date;
  Notes?: string;
}

// ==================== GOODS RECEIPT INTERFACES ====================

export interface IGoodsReceipt extends IBaseListItem {
  ReceiptNumber: string;
  PurchaseOrderId: number;
  PurchaseOrder?: IPurchaseOrder;
  VendorId: number;
  Vendor?: IVendor;

  // Receipt Details
  ReceiptDate: Date;
  ReceivedById: number;
  ReceivedBy?: IUser;

  // Delivery Info
  DeliveryNote?: string;
  PackingSlip?: string;
  CarrierName?: string;
  TrackingNumber?: string;

  // Status
  Status: ReceiptStatus;

  // Location
  ReceivedAtLocation?: string;
  StorageLocation?: string;

  // Metadata
  Attachments?: string; // JSON array (photos, delivery docs)
  Notes?: string;
}

export interface IReceiptLineItem extends IBaseListItem {
  ReceiptId: number;
  Receipt?: IGoodsReceipt;
  POLineItemId: number;
  POLineItem?: IPOLineItem;

  // Quantities
  QuantityExpected: number;
  QuantityReceived: number;
  QuantityRejected: number;

  // Condition
  Condition: string;
  RejectionReason?: string;

  // Serial/Batch
  SerialNumbers?: string; // JSON array
  BatchNumber?: string;

  // Asset Creation
  AssetIdsCreated?: string; // JSON array

  // Inspection
  InspectedById?: number;
  InspectedBy?: IUser;
  InspectionDate?: Date;
  InspectionNotes?: string;

  Notes?: string;
}

// ==================== CONTRACT INTERFACES ====================

export interface IContract extends IBaseListItem {
  ContractNumber: string;
  VendorId: number;
  Vendor?: IVendor;
  ContractType: ContractType;
  Status: ContractStatus;

  // Dates
  StartDate: Date;
  EndDate: Date;
  SignedDate?: Date;
  RenewalDate?: Date;
  TerminationDate?: Date;

  // Renewal
  AutoRenew: boolean;
  RenewalTermMonths?: number;
  NotificationDays: number;

  // Financial
  ContractValue: number;
  AnnualValue?: number;
  Currency: Currency;
  PaymentSchedule?: string;

  // Ownership
  OwnerIds?: string; // JSON array of user IDs
  Department?: string;
  CostCenter?: string;

  // Terms
  KeyTerms?: string; // JSON object
  Milestones?: string; // JSON array
  SLATerms?: string;
  PenaltyClause?: string;
  TerminationClause?: string;

  // Documents
  DocumentLibraryUrl?: string;
  DocumentIds?: string; // JSON array

  // Parent Contract (for amendments)
  ParentContractId?: number;
  ParentContract?: IContract;
  Version: number;
  AmendmentReason?: string;

  // Metadata
  Notes?: string;
  Tags?: string; // JSON array
}

export interface IContractMilestone {
  Id?: number;
  ContractId: number;
  Contract?: IContract;
  Title: string;
  Description?: string;
  DueDate: Date;
  CompletedDate?: Date;
  IsCompleted: boolean;
  Amount?: number;
  Currency?: Currency;
  Notes?: string;
}

// ==================== INVOICE INTERFACES ====================

export interface IInvoice extends IBaseListItem {
  InvoiceNumber: string;
  VendorId: number;
  Vendor?: IVendor;

  // References
  PurchaseOrderId?: number;
  PurchaseOrder?: IPurchaseOrder;
  ContractId?: number;
  Contract?: IContract;
  ReceiptId?: number;
  Receipt?: IGoodsReceipt;

  // Dates
  InvoiceDate: Date;
  ReceivedDate: Date;
  DueDate: Date;

  // Amounts
  Subtotal: number;
  TaxAmount?: number;
  ShippingAmount?: number;
  DiscountAmount?: number;
  TotalAmount: number;
  Currency: Currency;

  // Status
  Status: InvoiceStatus;
  MatchStatus?: string;
  MatchExceptionNotes?: string;

  // Approval
  ApprovedById?: number;
  ApprovedBy?: IUser;
  ApprovedDate?: Date;

  // Payment
  PaymentDate?: Date;
  PaymentReference?: string;
  PaymentMethod?: string;

  // Dispute
  DisputeReason?: string;
  DisputeResolvedDate?: Date;
  DisputeResolution?: string;

  // Budget
  BudgetId?: number;
  Budget?: IBudget;
  CostCenter?: string;
  Department?: string;

  // Metadata
  Attachments?: string; // JSON array
  Notes?: string;
}

export interface IInvoiceLineItem extends IBaseListItem {
  InvoiceId: number;
  Invoice?: IInvoice;
  POLineItemId?: number;
  POLineItem?: IPOLineItem;
  LineNumber: number;

  // Item Details
  Description: string;
  Quantity: number;
  UnitPrice: number;
  TotalPrice: number;
  TaxRate?: number;
  TaxAmount?: number;

  // Matching
  MatchedQuantity?: number;
  MatchedAmount?: number;
  MatchException?: string;

  Notes?: string;
}

// ==================== BUDGET INTERFACES ====================

export interface IBudget extends IBaseListItem {
  BudgetCode: string;
  FiscalYear: number;
  Department: string;
  CostCenter?: string;
  Category?: VendorCategory;
  Status: BudgetStatus;

  // Amounts
  BudgetAmount: number;
  AllocatedAmount: number;
  SpentAmount: number;
  RemainingAmount: number;
  Currency: Currency;

  // Thresholds
  WarningThreshold: number; // Percentage
  CriticalThreshold: number; // Percentage

  // Ownership
  OwnerIds?: string; // JSON array of user IDs
  ApproverIds?: string; // JSON array of user IDs

  // Period
  StartDate: Date;
  EndDate: Date;

  // Metadata
  Notes?: string;
}

export interface IBudgetAllocation {
  Id?: number;
  BudgetId: number;
  Budget?: IBudget;
  DocumentType: 'Requisition' | 'PurchaseOrder' | 'Invoice';
  DocumentId: number;
  DocumentNumber: string;
  Amount: number;
  Currency: Currency;
  AllocationDate: Date;
  AllocatedById: number;
  AllocatedBy?: IUser;
  IsReleased: boolean;
  ReleasedDate?: Date;
  Notes?: string;
}

// ==================== VENDOR PERFORMANCE INTERFACES ====================

export interface IVendorPerformance extends IBaseListItem {
  VendorId: number;
  Vendor?: IVendor;
  ReviewPeriod: string; // e.g., "2024-Q1"
  ReviewYear: number;
  ReviewQuarter: number;

  // Scores (1-5)
  OnTimeDeliveryScore: number;
  QualityScore: number;
  ResponsivenessScore: number;
  PricingScore: number;
  ComplianceScore: number;
  OverallScore: number;

  // Metrics
  TotalPOsInPeriod: number;
  TotalValueInPeriod: number;
  OnTimeDeliveryRate: number;
  DefectRate: number;
  ResponseTimeAvgDays: number;

  // Issues
  IssuesCount: number;
  ResolvedIssuesCount: number;

  // Review
  ReviewedById: number;
  ReviewedBy?: IUser;
  ReviewDate: Date;

  Comments?: string;
  ActionItems?: string; // JSON array
  RecommendedAction?: string;
}

export interface IVendorIssue extends IBaseListItem {
  VendorId: number;
  Vendor?: IVendor;
  PurchaseOrderId?: number;
  PurchaseOrder?: IPurchaseOrder;
  InvoiceId?: number;
  Invoice?: IInvoice;

  IssueType: string;
  Severity: 'Low' | 'Medium' | 'High' | 'Critical';
  Status: 'Open' | 'In Progress' | 'Resolved' | 'Closed' | 'Escalated';

  Description: string;
  RootCause?: string;
  Resolution?: string;

  ReportedById: number;
  ReportedBy?: IUser;
  ReportedDate: Date;

  AssignedToId?: number;
  AssignedTo?: IUser;

  ResolvedById?: number;
  ResolvedBy?: IUser;
  ResolvedDate?: Date;

  ImpactAmount?: number;
  Currency?: Currency;

  Notes?: string;
}

// ==================== STATISTICS & DASHBOARD INTERFACES ====================

export interface IProcurementStatistics {
  // Vendors
  totalVendors: number;
  activeVendors: number;
  preferredVendors: number;
  vendorsByCategory: { [key in VendorCategory]?: number };
  vendorsByStatus: { [key in VendorStatus]?: number };
  avgVendorRating: number;

  // Requisitions
  totalRequisitions: number;
  pendingApproval: number;
  requisitionsByStatus: { [key in RequisitionStatus]?: number };
  avgRequisitionValue: number;
  avgApprovalTime: number; // in hours

  // Purchase Orders
  totalPOs: number;
  openPOs: number;
  posByStatus: { [key in POStatus]?: number };
  totalPOValue: number;
  avgPOValue: number;

  // Invoices
  totalInvoices: number;
  pendingInvoices: number;
  overdueInvoices: number;
  invoicesByStatus: { [key in InvoiceStatus]?: number };
  totalInvoiceValue: number;
  avgPaymentDays: number;

  // Contracts
  totalContracts: number;
  activeContracts: number;
  expiringContracts: number; // within 90 days
  contractsByStatus: { [key in ContractStatus]?: number };
  totalContractValue: number;

  // Spend
  totalSpendYTD: number;
  totalSpendMTD: number;
  spendByCategory: { [key in VendorCategory]?: number };
  spendByDepartment: { [department: string]: number };
  topVendorsBySpend: { vendorId: number; vendorName: string; spend: number }[];

  // Budget
  totalBudget: number;
  totalAllocated: number;
  totalSpent: number;
  totalRemaining: number;
  budgetUtilization: number; // percentage
}

export interface IProcurementDashboard {
  statistics: IProcurementStatistics;
  recentRequisitions: IPurchaseRequisition[];
  pendingApprovals: IPurchaseRequisition[];
  recentPurchaseOrders: IPurchaseOrder[];
  expiringContracts: IContract[];
  overdueInvoices: IInvoice[];
  vendorAlerts: IVendorAlert[];
  budgetAlerts: IBudgetAlert[];
  spendTrend: ISpendTrendItem[];
}

export interface IVendorAlert {
  vendorId: number;
  vendorName: string;
  alertType: 'Performance' | 'Compliance' | 'Contract' | 'Payment';
  message: string;
  severity: 'Info' | 'Warning' | 'Critical';
  date: Date;
}

export interface IBudgetAlert {
  budgetId: number;
  budgetName: string;
  department: string;
  alertType: 'Warning' | 'Critical' | 'Exceeded';
  utilized: number;
  threshold: number;
  message: string;
}

export interface ISpendTrendItem {
  period: string; // e.g., "2024-01"
  spend: number;
  budget: number;
  variance: number;
}

// ==================== FILTER INTERFACES ====================

export interface IVendorFilter {
  searchTerm?: string;
  status?: VendorStatus[];
  type?: VendorType[];
  category?: VendorCategory[];
  preferredOnly?: boolean;
  minRating?: number;
  country?: string;
  hasActiveContracts?: boolean;
}

export interface IRequisitionFilter {
  searchTerm?: string;
  status?: RequisitionStatus[];
  priority?: RequisitionPriority[];
  requestedById?: number;
  department?: string;
  vendorId?: number;
  fromDate?: Date;
  toDate?: Date;
  minAmount?: number;
  maxAmount?: number;
  processId?: number; // JML integration
}

export interface IPurchaseOrderFilter {
  searchTerm?: string;
  status?: POStatus[];
  vendorId?: number;
  department?: string;
  fromDate?: Date;
  toDate?: Date;
  minAmount?: number;
  maxAmount?: number;
  overdue?: boolean;
}

export interface IContractFilter {
  searchTerm?: string;
  status?: ContractStatus[];
  type?: ContractType[];
  vendorId?: number;
  department?: string;
  expiringWithinDays?: number;
  minValue?: number;
  maxValue?: number;
}

export interface IInvoiceFilter {
  searchTerm?: string;
  status?: InvoiceStatus[];
  vendorId?: number;
  purchaseOrderId?: number;
  department?: string;
  fromDate?: Date;
  toDate?: Date;
  overdue?: boolean;
  minAmount?: number;
  maxAmount?: number;
}

// ==================== JML INTEGRATION INTERFACES ====================

export interface IJMLProcurementRequest {
  processId: number;
  processType: 'Joiner' | 'Mover' | 'Leaver';
  employeeId: number;
  employeeName: string;
  department: string;
  startDate: Date;
  equipmentTemplate?: string;
  requestedItems: IJMLRequestedItem[];
}

export interface IJMLRequestedItem {
  catalogItemId?: number;
  itemCode?: string;
  description: string;
  category: VendorCategory;
  quantity: number;
  specifications?: string;
  preferredVendorId?: number;
}

export interface IJMLProcurementResult {
  processId: number;
  requisitionId?: number;
  requisitionNumber?: string;
  purchaseOrderIds?: number[];
  assetIds?: number[];
  status: 'Created' | 'Pending' | 'Approved' | 'Fulfilled' | 'Failed';
  message?: string;
}

// ==================== APPROVAL INTERFACES ====================

export interface IProcurementApprovalRule {
  Id?: number;
  Title: string;
  DocumentType: 'Requisition' | 'PurchaseOrder' | 'Contract' | 'Invoice';
  MinAmount?: number;
  MaxAmount?: number;
  Category?: VendorCategory;
  Department?: string;
  ApproverIds: string; // JSON array of user IDs
  ApprovalOrder: number;
  RequireAll: boolean;
  IsActive: boolean;
}

export interface IProcurementApproval {
  Id?: number;
  DocumentType: 'Requisition' | 'PurchaseOrder' | 'Contract' | 'Invoice';
  DocumentId: number;
  DocumentNumber: string;

  RequestedById: number;
  RequestedBy?: IUser;
  RequestedDate: Date;

  ApproverId: number;
  Approver?: IUser;
  ApprovalLevel: number;

  Status: 'Pending' | 'Approved' | 'Rejected' | 'Delegated';
  Action?: ApprovalAction;
  ActionDate?: Date;
  Comments?: string;

  DelegatedToId?: number;
  DelegatedTo?: IUser;
}

// ==================== EXPORT INTERFACES ====================

export interface IProcurementExportOptions {
  format: 'PDF' | 'Excel' | 'CSV';
  includeLineItems: boolean;
  includeHistory: boolean;
  includeApprovals: boolean;
  dateRange?: {
    from: Date;
    to: Date;
  };
}

export interface IPOPrintOptions {
  includeTerms: boolean;
  includeCompanyLogo: boolean;
  includeSignature: boolean;
  copies: number;
}
