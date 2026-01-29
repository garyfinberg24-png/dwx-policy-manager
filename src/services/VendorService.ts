// @ts-nocheck
/* eslint-disable @typescript-eslint/no-explicit-any */
// Vendor Service
// Comprehensive vendor management operations

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  IVendor,
  IVendorContact,
  IVendorDocument,
  IVendorPerformance,
  IVendorIssue,
  IVendorFilter,
  VendorStatus,
  VendorType,
  VendorCategory,
  PaymentTerms,
  Currency
} from '../models/IProcurement';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class VendorService {
  private sp: SPFI;
  private readonly VENDORS_LIST = 'PM_Vendors';
  private readonly VENDOR_CONTACTS_LIST = 'PM_VendorContacts';
  private readonly VENDOR_DOCUMENTS_LIST = 'PM_VendorDocuments';
  private readonly VENDOR_PERFORMANCE_LIST = 'PM_VendorPerformance';
  private readonly VENDOR_ISSUES_LIST = 'PM_VendorIssues';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==================== Vendor CRUD Operations ====================

  public async getVendors(filter?: IVendorFilter): Promise<IVendor[]> {
    try {
      console.log(`[VendorService] Fetching vendors from list: ${this.VENDORS_LIST}`);
      // Field names mapped to actual SharePoint list schema
      // NOTE: Removed ApprovedById - column may not exist or may cause 400 errors
      let query = this.sp.web.lists.getByTitle(this.VENDORS_LIST).items
        .select(
          'Id', 'Title', 'VendorCode', 'VendorType', 'VendorCategory', 'Status',
          'VendorName', 'TaxId', 'RegistrationNumber', 'Website',
          'Address', 'City', 'State', 'Country', 'PostalCode',
          'Phone', 'Email', 'PaymentTerms', 'Currency',
          'PreferredVendor', 'IsPreferred', 'IsActive',
          'ApprovedDate',
          'Rating',
          'InsuranceExpiryDate', 'CertificationExpiryDate',
          'Notes',
          'Created', 'Modified'
        );

      // Apply filters
      if (filter) {
        const filters: string[] = [];

        if (filter.searchTerm) {
          const term = ValidationUtils.sanitizeForOData(filter.searchTerm);
          filters.push(`(substringof('${term}', Title) or substringof('${term}', VendorCode) or substringof('${term}', VendorName))`);
        }

        if (filter.status && filter.status.length > 0) {
          const statusFilters = filter.status.map(s =>
            ValidationUtils.buildFilter('Status', 'eq', s)
          );
          filters.push(`(${statusFilters.join(' or ')})`);
        }

        if (filter.type && filter.type.length > 0) {
          const typeFilters = filter.type.map(t =>
            ValidationUtils.buildFilter('VendorType', 'eq', t)
          );
          filters.push(`(${typeFilters.join(' or ')})`);
        }

        if (filter.category && filter.category.length > 0) {
          const categoryFilters = filter.category.map(c =>
            ValidationUtils.buildFilter('VendorCategory', 'eq', c)
          );
          filters.push(`(${categoryFilters.join(' or ')})`);
        }

        if (filter.preferredOnly) {
          filters.push('PreferredVendor eq 1');
        }

        if (filter.minRating !== undefined) {
          filters.push(`Rating ge ${filter.minRating}`);
        }

        if (filter.country) {
          const country = ValidationUtils.sanitizeForOData(filter.country);
          filters.push(ValidationUtils.buildFilter('Country', 'eq', country));
        }

        if (filters.length > 0) {
          query = query.filter(filters.join(' and '));
        }
      }

      const items = await query.orderBy('Title', true).top(5000)();
      console.log(`[VendorService] Retrieved ${items.length} vendors from ${this.VENDORS_LIST}`);
      return items.map(this.mapVendorFromSP);
    } catch (error: any) {
      console.error(`[VendorService] Error getting vendors from ${this.VENDORS_LIST}:`, error?.message || error);
      logger.error('VendorService', 'Error getting vendors:', error);
      // Return empty array instead of throwing to prevent UI crashes
      return [];
    }
  }

  public async getVendorById(id: number): Promise<IVendor | null> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Field names mapped to actual SharePoint list schema
      // NOTE: Removed ApprovedById - column may not exist or may cause 400 errors
      const item = await this.sp.web.lists.getByTitle(this.VENDORS_LIST).items
        .getById(validId)
        .select(
          'Id', 'Title', 'VendorCode', 'VendorType', 'VendorCategory', 'Status',
          'VendorName', 'TaxId', 'RegistrationNumber', 'Website',
          'Address', 'City', 'State', 'Country', 'PostalCode',
          'Phone', 'Email', 'PaymentTerms', 'Currency',
          'BankName', 'BankAccountNumber', 'BankRoutingNumber',
          'PreferredVendor', 'IsPreferred', 'IsActive',
          'ApprovedDate',
          'Rating',
          'InsuranceExpiryDate', 'CertificationExpiryDate',
          'Notes',
          'Created', 'Modified'
        )();

      return this.mapVendorFromSP(item);
    } catch (error: any) {
      console.error(`[VendorService] Error getting vendor by ID ${id}:`, error?.message || error);
      logger.error('VendorService', 'Error getting vendor by ID:', error);
      return null;
    }
  }

  public async getVendorByCode(vendorCode: string): Promise<IVendor | null> {
    try {
      if (!vendorCode || typeof vendorCode !== 'string') {
        throw new Error('Invalid vendor code');
      }

      const validCode = ValidationUtils.sanitizeForOData(vendorCode.substring(0, 50));
      const filter = ValidationUtils.buildFilter('VendorCode', 'eq', validCode);

      const items = await this.sp.web.lists.getByTitle(this.VENDORS_LIST).items
        .select('Id', 'Title', 'VendorCode', 'Status')
        .filter(filter)
        .top(1)();

      if (items.length === 0) {
        return null;
      }

      return this.getVendorById(items[0].Id);
    } catch (error) {
      logger.error('VendorService', 'Error getting vendor by code:', error);
      throw error;
    }
  }

  public async createVendor(vendor: Partial<IVendor>): Promise<number> {
    try {
      // Validate required fields
      if (!vendor.Title) {
        throw new Error('Vendor name (Title) is required');
      }

      // Generate vendor code if not provided
      const vendorCode = vendor.VendorCode || await this.generateVendorCode();

      // Check if vendor code already exists
      const existing = await this.getVendorByCode(vendorCode);
      if (existing) {
        throw new Error(`Vendor code ${vendorCode} already exists`);
      }

      const itemData: Record<string, unknown> = {
        Title: ValidationUtils.sanitizeHtml(vendor.Title),
        VendorCode: vendorCode,
        VendorType: vendor.VendorType || VendorType.Supplier,
        Category: vendor.Category || VendorCategory.Other,
        Status: vendor.Status || VendorStatus.PendingApproval,
        PaymentTerms: vendor.PaymentTerms || PaymentTerms.Net30,
        Currency: vendor.Currency || Currency.USD,
        PreferredVendor: vendor.PreferredVendor || false
      };

      // Optional fields
      if (vendor.LegalName) itemData.LegalName = ValidationUtils.sanitizeHtml(vendor.LegalName);
      if (vendor.TradingName) itemData.TradingName = ValidationUtils.sanitizeHtml(vendor.TradingName);
      if (vendor.TaxId) itemData.TaxId = ValidationUtils.sanitizeHtml(vendor.TaxId);
      if (vendor.DunsNumber) itemData.DunsNumber = ValidationUtils.sanitizeHtml(vendor.DunsNumber);
      if (vendor.CompanyRegistration) itemData.CompanyRegistration = ValidationUtils.sanitizeHtml(vendor.CompanyRegistration);
      if (vendor.Website) itemData.Website = ValidationUtils.sanitizeHtml(vendor.Website);

      // Address
      if (vendor.AddressLine1) itemData.AddressLine1 = ValidationUtils.sanitizeHtml(vendor.AddressLine1);
      if (vendor.AddressLine2) itemData.AddressLine2 = ValidationUtils.sanitizeHtml(vendor.AddressLine2);
      if (vendor.City) itemData.City = ValidationUtils.sanitizeHtml(vendor.City);
      if (vendor.State) itemData.State = ValidationUtils.sanitizeHtml(vendor.State);
      if (vendor.Country) itemData.Country = ValidationUtils.sanitizeHtml(vendor.Country);
      if (vendor.PostalCode) itemData.PostalCode = ValidationUtils.sanitizeHtml(vendor.PostalCode);

      // Contact
      if (vendor.PrimaryContactId) itemData.PrimaryContactId = ValidationUtils.validateInteger(vendor.PrimaryContactId, 'PrimaryContactId', 1);
      if (vendor.PrimaryPhone) itemData.PrimaryPhone = ValidationUtils.sanitizeHtml(vendor.PrimaryPhone);
      if (vendor.PrimaryEmail) itemData.PrimaryEmail = ValidationUtils.sanitizeHtml(vendor.PrimaryEmail);

      // Notes
      if (vendor.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(vendor.Notes);
      if (vendor.Tags) itemData.Tags = vendor.Tags;

      const result = await this.sp.web.lists.getByTitle(this.VENDORS_LIST).items.add(itemData);
      return result.data.Id;
    } catch (error) {
      logger.error('VendorService', 'Error creating vendor:', error);
      throw error;
    }
  }

  public async updateVendor(id: number, updates: Partial<IVendor>): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const itemData: Record<string, unknown> = {};

      if (updates.Title) itemData.Title = ValidationUtils.sanitizeHtml(updates.Title);
      if (updates.VendorType) {
        ValidationUtils.validateEnum(updates.VendorType, VendorType, 'VendorType');
        itemData.VendorType = updates.VendorType;
      }
      if (updates.Category) {
        ValidationUtils.validateEnum(updates.Category, VendorCategory, 'Category');
        itemData.Category = updates.Category;
      }
      if (updates.Status) {
        ValidationUtils.validateEnum(updates.Status, VendorStatus, 'Status');
        itemData.Status = updates.Status;
      }
      if (updates.PaymentTerms) {
        ValidationUtils.validateEnum(updates.PaymentTerms, PaymentTerms, 'PaymentTerms');
        itemData.PaymentTerms = updates.PaymentTerms;
      }
      if (updates.Currency) {
        ValidationUtils.validateEnum(updates.Currency, Currency, 'Currency');
        itemData.Currency = updates.Currency;
      }

      // Company info
      if (updates.LegalName !== undefined) itemData.LegalName = updates.LegalName ? ValidationUtils.sanitizeHtml(updates.LegalName) : null;
      if (updates.TradingName !== undefined) itemData.TradingName = updates.TradingName ? ValidationUtils.sanitizeHtml(updates.TradingName) : null;
      if (updates.TaxId !== undefined) itemData.TaxId = updates.TaxId ? ValidationUtils.sanitizeHtml(updates.TaxId) : null;
      if (updates.DunsNumber !== undefined) itemData.DunsNumber = updates.DunsNumber ? ValidationUtils.sanitizeHtml(updates.DunsNumber) : null;
      if (updates.Website !== undefined) itemData.Website = updates.Website ? ValidationUtils.sanitizeHtml(updates.Website) : null;

      // Address
      if (updates.AddressLine1 !== undefined) itemData.AddressLine1 = updates.AddressLine1 ? ValidationUtils.sanitizeHtml(updates.AddressLine1) : null;
      if (updates.AddressLine2 !== undefined) itemData.AddressLine2 = updates.AddressLine2 ? ValidationUtils.sanitizeHtml(updates.AddressLine2) : null;
      if (updates.City !== undefined) itemData.City = updates.City ? ValidationUtils.sanitizeHtml(updates.City) : null;
      if (updates.State !== undefined) itemData.State = updates.State ? ValidationUtils.sanitizeHtml(updates.State) : null;
      if (updates.Country !== undefined) itemData.Country = updates.Country ? ValidationUtils.sanitizeHtml(updates.Country) : null;
      if (updates.PostalCode !== undefined) itemData.PostalCode = updates.PostalCode ? ValidationUtils.sanitizeHtml(updates.PostalCode) : null;

      // Contact
      if (updates.PrimaryContactId !== undefined) {
        itemData.PrimaryContactId = updates.PrimaryContactId === null ? null :
          ValidationUtils.validateInteger(updates.PrimaryContactId, 'PrimaryContactId', 1);
      }
      if (updates.PrimaryPhone !== undefined) itemData.PrimaryPhone = updates.PrimaryPhone ? ValidationUtils.sanitizeHtml(updates.PrimaryPhone) : null;
      if (updates.PrimaryEmail !== undefined) itemData.PrimaryEmail = updates.PrimaryEmail ? ValidationUtils.sanitizeHtml(updates.PrimaryEmail) : null;

      // Preferences
      if (updates.PreferredVendor !== undefined) itemData.PreferredVendor = updates.PreferredVendor;

      // Dates
      if (updates.ApprovedDate) itemData.ApprovedDate = ValidationUtils.validateDate(updates.ApprovedDate, 'ApprovedDate');
      if (updates.ApprovedById) itemData.ApprovedById = ValidationUtils.validateInteger(updates.ApprovedById, 'ApprovedById', 1);
      if (updates.LastReviewDate) itemData.LastReviewDate = ValidationUtils.validateDate(updates.LastReviewDate, 'LastReviewDate');
      if (updates.NextReviewDate) itemData.NextReviewDate = ValidationUtils.validateDate(updates.NextReviewDate, 'NextReviewDate');

      // Compliance
      if (updates.Certifications !== undefined) itemData.Certifications = updates.Certifications;
      if (updates.InsuranceExpiry) itemData.InsuranceExpiry = ValidationUtils.validateDate(updates.InsuranceExpiry, 'InsuranceExpiry');
      if (updates.ComplianceDocuments !== undefined) itemData.ComplianceDocuments = updates.ComplianceDocuments;

      // Notes
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes ? ValidationUtils.sanitizeHtml(updates.Notes) : null;
      if (updates.Tags !== undefined) itemData.Tags = updates.Tags;

      await this.sp.web.lists.getByTitle(this.VENDORS_LIST).items.getById(validId).update(itemData);
    } catch (error) {
      logger.error('VendorService', 'Error updating vendor:', error);
      throw error;
    }
  }

  public async deleteVendor(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Check if vendor has active POs or contracts before deleting
      // This would require checking other lists - for now, just mark as inactive
      await this.updateVendor(validId, { Status: VendorStatus.Inactive });
    } catch (error) {
      logger.error('VendorService', 'Error deleting vendor:', error);
      throw error;
    }
  }

  public async approveVendor(id: number, approvedById: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);
      const validApprovedById = ValidationUtils.validateInteger(approvedById, 'approvedById', 1);

      await this.sp.web.lists.getByTitle(this.VENDORS_LIST).items.getById(validId).update({
        Status: VendorStatus.Active,
        ApprovedDate: new Date().toISOString(),
        ApprovedById: validApprovedById
      });
    } catch (error) {
      logger.error('VendorService', 'Error approving vendor:', error);
      throw error;
    }
  }

  public async togglePreferredVendor(id: number): Promise<boolean> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);
      const vendor = await this.getVendorById(validId);
      const newStatus = !vendor.PreferredVendor;

      await this.sp.web.lists.getByTitle(this.VENDORS_LIST).items.getById(validId).update({
        PreferredVendor: newStatus
      });

      return newStatus;
    } catch (error) {
      logger.error('VendorService', 'Error toggling preferred vendor:', error);
      throw error;
    }
  }

  // ==================== Vendor Contacts ====================

  public async getVendorContacts(vendorId: number): Promise<IVendorContact[]> {
    try {
      const validVendorId = ValidationUtils.validateInteger(vendorId, 'vendorId', 1);

      const items = await this.sp.web.lists.getByTitle(this.VENDOR_CONTACTS_LIST).items
        .select(
          'Id', 'Title', 'VendorId', 'FirstName', 'LastName', 'Email', 'Phone', 'Mobile',
          'Role', 'Department', 'IsPrimary', 'IsActive', 'Notes',
          'Created', 'Modified'
        )
        .filter(`VendorId eq ${validVendorId}`)
        .orderBy('IsPrimary', false)
        .orderBy('LastName', true)();

      return items.map(this.mapVendorContactFromSP);
    } catch (error) {
      logger.error('VendorService', 'Error getting vendor contacts:', error);
      throw error;
    }
  }

  public async createVendorContact(contact: Partial<IVendorContact>): Promise<number> {
    try {
      if (!contact.VendorId || !contact.FirstName || !contact.LastName || !contact.Email) {
        throw new Error('VendorId, FirstName, LastName, and Email are required');
      }

      const itemData: Record<string, unknown> = {
        Title: `${contact.FirstName} ${contact.LastName}`,
        VendorId: ValidationUtils.validateInteger(contact.VendorId, 'VendorId', 1),
        FirstName: ValidationUtils.sanitizeHtml(contact.FirstName),
        LastName: ValidationUtils.sanitizeHtml(contact.LastName),
        Email: ValidationUtils.sanitizeHtml(contact.Email),
        IsPrimary: contact.IsPrimary || false,
        IsActive: contact.IsActive !== undefined ? contact.IsActive : true
      };

      if (contact.Phone) itemData.Phone = ValidationUtils.sanitizeHtml(contact.Phone);
      if (contact.Mobile) itemData.Mobile = ValidationUtils.sanitizeHtml(contact.Mobile);
      if (contact.Role) itemData.Role = ValidationUtils.sanitizeHtml(contact.Role);
      if (contact.Department) itemData.Department = ValidationUtils.sanitizeHtml(contact.Department);
      if (contact.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(contact.Notes);

      const result = await this.sp.web.lists.getByTitle(this.VENDOR_CONTACTS_LIST).items.add(itemData);
      return result.data.Id;
    } catch (error) {
      logger.error('VendorService', 'Error creating vendor contact:', error);
      throw error;
    }
  }

  public async updateVendorContact(id: number, updates: Partial<IVendorContact>): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const itemData: Record<string, unknown> = {};

      if (updates.FirstName) itemData.FirstName = ValidationUtils.sanitizeHtml(updates.FirstName);
      if (updates.LastName) itemData.LastName = ValidationUtils.sanitizeHtml(updates.LastName);
      if (updates.FirstName || updates.LastName) {
        itemData.Title = `${updates.FirstName || ''} ${updates.LastName || ''}`.trim();
      }
      if (updates.Email) itemData.Email = ValidationUtils.sanitizeHtml(updates.Email);
      if (updates.Phone !== undefined) itemData.Phone = updates.Phone ? ValidationUtils.sanitizeHtml(updates.Phone) : null;
      if (updates.Mobile !== undefined) itemData.Mobile = updates.Mobile ? ValidationUtils.sanitizeHtml(updates.Mobile) : null;
      if (updates.Role !== undefined) itemData.Role = updates.Role ? ValidationUtils.sanitizeHtml(updates.Role) : null;
      if (updates.Department !== undefined) itemData.Department = updates.Department ? ValidationUtils.sanitizeHtml(updates.Department) : null;
      if (updates.IsPrimary !== undefined) itemData.IsPrimary = updates.IsPrimary;
      if (updates.IsActive !== undefined) itemData.IsActive = updates.IsActive;
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes ? ValidationUtils.sanitizeHtml(updates.Notes) : null;

      await this.sp.web.lists.getByTitle(this.VENDOR_CONTACTS_LIST).items.getById(validId).update(itemData);
    } catch (error) {
      logger.error('VendorService', 'Error updating vendor contact:', error);
      throw error;
    }
  }

  public async deleteVendorContact(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);
      await this.sp.web.lists.getByTitle(this.VENDOR_CONTACTS_LIST).items.getById(validId).delete();
    } catch (error) {
      logger.error('VendorService', 'Error deleting vendor contact:', error);
      throw error;
    }
  }

  // ==================== Vendor Performance ====================

  public async getVendorPerformance(vendorId: number, limit?: number): Promise<IVendorPerformance[]> {
    try {
      const validVendorId = ValidationUtils.validateInteger(vendorId, 'vendorId', 1);

      const items = await this.sp.web.lists.getByTitle(this.VENDOR_PERFORMANCE_LIST).items
        .select(
          'Id', 'Title', 'VendorId', 'ReviewPeriod', 'ReviewYear', 'ReviewQuarter',
          'OnTimeDeliveryScore', 'QualityScore', 'ResponsivenessScore', 'PricingScore', 'ComplianceScore', 'OverallScore',
          'TotalPOsInPeriod', 'TotalValueInPeriod', 'OnTimeDeliveryRate', 'DefectRate', 'ResponseTimeAvgDays',
          'IssuesCount', 'ResolvedIssuesCount',
          'ReviewedById', 'ReviewedBy/Title', 'ReviewDate',
          'Comments', 'ActionItems', 'RecommendedAction',
          'Created', 'Modified'
        )
        .expand('ReviewedBy')
        .filter(`VendorId eq ${validVendorId}`)
        .orderBy('ReviewYear', false)
        .orderBy('ReviewQuarter', false)
        .top(limit || 12)();

      return items.map(this.mapVendorPerformanceFromSP);
    } catch (error) {
      logger.error('VendorService', 'Error getting vendor performance:', error);
      throw error;
    }
  }

  public async createPerformanceReview(review: Partial<IVendorPerformance>): Promise<number> {
    try {
      if (!review.VendorId || !review.ReviewPeriod || review.OverallScore === undefined) {
        throw new Error('VendorId, ReviewPeriod, and OverallScore are required');
      }

      const itemData: Record<string, unknown> = {
        Title: `${review.ReviewPeriod} Review`,
        VendorId: ValidationUtils.validateInteger(review.VendorId, 'VendorId', 1),
        ReviewPeriod: review.ReviewPeriod,
        ReviewYear: review.ReviewYear || new Date().getFullYear(),
        ReviewQuarter: review.ReviewQuarter || Math.ceil((new Date().getMonth() + 1) / 3),
        OnTimeDeliveryScore: review.OnTimeDeliveryScore || 0,
        QualityScore: review.QualityScore || 0,
        ResponsivenessScore: review.ResponsivenessScore || 0,
        PricingScore: review.PricingScore || 0,
        ComplianceScore: review.ComplianceScore || 0,
        OverallScore: review.OverallScore,
        TotalPOsInPeriod: review.TotalPOsInPeriod || 0,
        TotalValueInPeriod: review.TotalValueInPeriod || 0,
        OnTimeDeliveryRate: review.OnTimeDeliveryRate || 0,
        DefectRate: review.DefectRate || 0,
        ResponseTimeAvgDays: review.ResponseTimeAvgDays || 0,
        IssuesCount: review.IssuesCount || 0,
        ResolvedIssuesCount: review.ResolvedIssuesCount || 0,
        ReviewDate: new Date().toISOString()
      };

      if (review.ReviewedById) itemData.ReviewedById = ValidationUtils.validateInteger(review.ReviewedById, 'ReviewedById', 1);
      if (review.Comments) itemData.Comments = ValidationUtils.sanitizeHtml(review.Comments);
      if (review.ActionItems) itemData.ActionItems = review.ActionItems;
      if (review.RecommendedAction) itemData.RecommendedAction = ValidationUtils.sanitizeHtml(review.RecommendedAction);

      const result = await this.sp.web.lists.getByTitle(this.VENDOR_PERFORMANCE_LIST).items.add(itemData);

      // Update vendor's overall rating
      await this.updateVendorRating(review.VendorId);

      return result.data.Id;
    } catch (error) {
      logger.error('VendorService', 'Error creating performance review:', error);
      throw error;
    }
  }

  // ==================== Vendor Issues ====================

  public async getVendorIssues(vendorId: number, openOnly?: boolean): Promise<IVendorIssue[]> {
    try {
      const validVendorId = ValidationUtils.validateInteger(vendorId, 'vendorId', 1);

      let query = this.sp.web.lists.getByTitle(this.VENDOR_ISSUES_LIST).items
        .select(
          'Id', 'Title', 'VendorId', 'PurchaseOrderId', 'InvoiceId',
          'IssueType', 'Severity', 'Status', 'Description', 'RootCause', 'Resolution',
          'ReportedById', 'ReportedBy/Title', 'ReportedDate',
          'AssignedToId', 'AssignedTo/Title',
          'ResolvedById', 'ResolvedBy/Title', 'ResolvedDate',
          'ImpactAmount', 'Currency', 'Notes',
          'Created', 'Modified'
        )
        .expand('ReportedBy', 'AssignedTo', 'ResolvedBy')
        .orderBy('Created', false);

      let filter = `VendorId eq ${validVendorId}`;
      if (openOnly) {
        filter += ` and (Status eq 'Open' or Status eq 'In Progress')`;
      }

      const items = await query.filter(filter).top(100)();

      return items.map(this.mapVendorIssueFromSP);
    } catch (error) {
      logger.error('VendorService', 'Error getting vendor issues:', error);
      throw error;
    }
  }

  public async createVendorIssue(issue: Partial<IVendorIssue>): Promise<number> {
    try {
      if (!issue.VendorId || !issue.IssueType || !issue.Description || !issue.ReportedById) {
        throw new Error('VendorId, IssueType, Description, and ReportedById are required');
      }

      const itemData: Record<string, unknown> = {
        Title: `${issue.IssueType} - ${new Date().toISOString().split('T')[0]}`,
        VendorId: ValidationUtils.validateInteger(issue.VendorId, 'VendorId', 1),
        IssueType: ValidationUtils.sanitizeHtml(issue.IssueType),
        Severity: issue.Severity || 'Medium',
        Status: 'Open',
        Description: ValidationUtils.sanitizeHtml(issue.Description),
        ReportedById: ValidationUtils.validateInteger(issue.ReportedById, 'ReportedById', 1),
        ReportedDate: new Date().toISOString()
      };

      if (issue.PurchaseOrderId) itemData.PurchaseOrderId = ValidationUtils.validateInteger(issue.PurchaseOrderId, 'PurchaseOrderId', 1);
      if (issue.InvoiceId) itemData.InvoiceId = ValidationUtils.validateInteger(issue.InvoiceId, 'InvoiceId', 1);
      if (issue.RootCause) itemData.RootCause = ValidationUtils.sanitizeHtml(issue.RootCause);
      if (issue.AssignedToId) itemData.AssignedToId = ValidationUtils.validateInteger(issue.AssignedToId, 'AssignedToId', 1);
      if (issue.ImpactAmount) itemData.ImpactAmount = issue.ImpactAmount;
      if (issue.Currency) itemData.Currency = issue.Currency;
      if (issue.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(issue.Notes);

      const result = await this.sp.web.lists.getByTitle(this.VENDOR_ISSUES_LIST).items.add(itemData);
      return result.data.Id;
    } catch (error) {
      logger.error('VendorService', 'Error creating vendor issue:', error);
      throw error;
    }
  }

  public async resolveVendorIssue(id: number, resolvedById: number, resolution: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);
      const validResolvedById = ValidationUtils.validateInteger(resolvedById, 'resolvedById', 1);

      await this.sp.web.lists.getByTitle(this.VENDOR_ISSUES_LIST).items.getById(validId).update({
        Status: 'Resolved',
        Resolution: ValidationUtils.sanitizeHtml(resolution),
        ResolvedById: validResolvedById,
        ResolvedDate: new Date().toISOString()
      });
    } catch (error) {
      logger.error('VendorService', 'Error resolving vendor issue:', error);
      throw error;
    }
  }

  // ==================== Statistics ====================

  public async getVendorStatistics(): Promise<{
    total: number;
    active: number;
    preferred: number;
    pending: number;
    byCategory: { [key: string]: number };
    avgRating: number;
  }> {
    try {
      const vendors = await this.getVendors();

      const stats = {
        total: vendors.length,
        active: vendors.filter(v => v.Status === VendorStatus.Active).length,
        preferred: vendors.filter(v => v.PreferredVendor).length,
        pending: vendors.filter(v => v.Status === VendorStatus.PendingApproval).length,
        byCategory: {} as { [key: string]: number },
        avgRating: 0
      };

      // Count by category
      for (const category of Object.values(VendorCategory)) {
        stats.byCategory[category] = vendors.filter(v => v.Category === category).length;
      }

      // Calculate average rating
      const ratedVendors = vendors.filter(v => v.Rating && v.Rating > 0);
      if (ratedVendors.length > 0) {
        stats.avgRating = ratedVendors.reduce((sum, v) => sum + (v.Rating || 0), 0) / ratedVendors.length;
      }

      return stats;
    } catch (error) {
      logger.error('VendorService', 'Error getting vendor statistics:', error);
      throw error;
    }
  }

  // ==================== Helper Functions ====================

  private async generateVendorCode(): Promise<string> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.VENDORS_LIST).items
        .select('VendorCode')
        .orderBy('Id', false)
        .top(1)();

      let nextNumber = 1;
      if (items.length > 0 && items[0].VendorCode) {
        const match = items[0].VendorCode.match(/VND-(\d+)/);
        if (match) {
          nextNumber = parseInt(match[1], 10) + 1;
        }
      }

      return `VND-${nextNumber.toString().padStart(5, '0')}`;
    } catch (error) {
      logger.error('VendorService', 'Error generating vendor code:', error);
      return `VND-${Date.now()}`;
    }
  }

  private async updateVendorRating(vendorId: number): Promise<void> {
    try {
      const performances = await this.getVendorPerformance(vendorId, 4); // Last 4 reviews

      if (performances.length > 0) {
        const avgRating = performances.reduce((sum, p) => sum + p.OverallScore, 0) / performances.length;

        await this.sp.web.lists.getByTitle(this.VENDORS_LIST).items.getById(vendorId).update({
          Rating: Math.round(avgRating * 10) / 10,
          RatingCount: performances.length
        });
      }
    } catch (error) {
      logger.error('VendorService', 'Error updating vendor rating:', error);
    }
  }

  // ==================== Mapping Functions ====================

  private mapVendorFromSP(item: Record<string, unknown>): IVendor {
    // Map SharePoint field names to interface properties
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      VendorCode: item.VendorCode as string,
      VendorType: item.VendorType as VendorType,
      Category: item.VendorCategory as VendorCategory, // SP field: VendorCategory
      Status: item.Status as VendorStatus,
      LegalName: item.VendorName as string, // SP field: VendorName
      TradingName: item.VendorName as string, // SP field: VendorName
      TaxId: item.TaxId as string,
      DunsNumber: undefined, // Not in SP list
      CompanyRegistration: item.RegistrationNumber as string, // SP field: RegistrationNumber
      Website: item.Website as string,
      AddressLine1: item.Address as string, // SP field: Address
      AddressLine2: undefined, // Not in SP list
      City: item.City as string,
      State: item.State as string,
      Country: item.Country as string,
      PostalCode: item.PostalCode as string,
      PrimaryContactId: undefined,
      PrimaryContact: undefined,
      PrimaryPhone: item.Phone as string, // SP field: Phone
      PrimaryEmail: item.Email as string, // SP field: Email
      PaymentTerms: item.PaymentTerms as PaymentTerms,
      Currency: item.Currency as Currency,
      BankName: item.BankName as string,
      BankAccountNumber: item.BankAccountNumber as string,
      BankRoutingNumber: item.BankRoutingNumber as string,
      BankSwiftCode: undefined, // Not in SP list
      PreferredVendor: (item.PreferredVendor || item.IsPreferred) as boolean,
      ApprovedDate: item.ApprovedDate ? new Date(item.ApprovedDate as string) : undefined,
      ApprovedById: item.ApprovedById as number,
      ApprovedBy: undefined,
      LastReviewDate: undefined, // Not in SP list
      NextReviewDate: undefined, // Not in SP list
      Rating: item.Rating as number,
      RatingCount: 0, // Not in SP list
      TotalOrders: 0, // Not in SP list
      TotalSpend: 0, // Not in SP list
      Certifications: undefined, // Not in SP list (has CertificationExpiryDate instead)
      InsuranceExpiry: item.InsuranceExpiryDate ? new Date(item.InsuranceExpiryDate as string) : undefined, // SP field: InsuranceExpiryDate
      ComplianceDocuments: undefined,
      Notes: item.Notes as string,
      Tags: undefined, // Not in SP list
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }

  private mapVendorContactFromSP(item: Record<string, unknown>): IVendorContact {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      VendorId: item.VendorId as number,
      FirstName: item.FirstName as string,
      LastName: item.LastName as string,
      Email: item.Email as string,
      Phone: item.Phone as string,
      Mobile: item.Mobile as string,
      Role: item.Role as string,
      Department: item.Department as string,
      IsPrimary: item.IsPrimary as boolean,
      IsActive: item.IsActive as boolean,
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }

  private mapVendorPerformanceFromSP(item: Record<string, unknown>): IVendorPerformance {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      VendorId: item.VendorId as number,
      ReviewPeriod: item.ReviewPeriod as string,
      ReviewYear: item.ReviewYear as number,
      ReviewQuarter: item.ReviewQuarter as number,
      OnTimeDeliveryScore: item.OnTimeDeliveryScore as number,
      QualityScore: item.QualityScore as number,
      ResponsivenessScore: item.ResponsivenessScore as number,
      PricingScore: item.PricingScore as number,
      ComplianceScore: item.ComplianceScore as number,
      OverallScore: item.OverallScore as number,
      TotalPOsInPeriod: item.TotalPOsInPeriod as number,
      TotalValueInPeriod: item.TotalValueInPeriod as number,
      OnTimeDeliveryRate: item.OnTimeDeliveryRate as number,
      DefectRate: item.DefectRate as number,
      ResponseTimeAvgDays: item.ResponseTimeAvgDays as number,
      IssuesCount: item.IssuesCount as number,
      ResolvedIssuesCount: item.ResolvedIssuesCount as number,
      ReviewedById: item.ReviewedById as number,
      ReviewedBy: item.ReviewedBy as any,
      ReviewDate: item.ReviewDate ? new Date(item.ReviewDate as string) : new Date(),
      Comments: item.Comments as string,
      ActionItems: item.ActionItems as string,
      RecommendedAction: item.RecommendedAction as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }

  private mapVendorIssueFromSP(item: Record<string, unknown>): IVendorIssue {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      VendorId: item.VendorId as number,
      PurchaseOrderId: item.PurchaseOrderId as number,
      InvoiceId: item.InvoiceId as number,
      IssueType: item.IssueType as string,
      Severity: item.Severity as 'Low' | 'Medium' | 'High' | 'Critical',
      Status: item.Status as 'Open' | 'In Progress' | 'Resolved' | 'Closed' | 'Escalated',
      Description: item.Description as string,
      RootCause: item.RootCause as string,
      Resolution: item.Resolution as string,
      ReportedById: item.ReportedById as number,
      ReportedBy: item.ReportedBy as any,
      ReportedDate: item.ReportedDate ? new Date(item.ReportedDate as string) : new Date(),
      AssignedToId: item.AssignedToId as number,
      AssignedTo: item.AssignedTo as any,
      ResolvedById: item.ResolvedById as number,
      ResolvedBy: item.ResolvedBy as any,
      ResolvedDate: item.ResolvedDate ? new Date(item.ResolvedDate as string) : undefined,
      ImpactAmount: item.ImpactAmount as number,
      Currency: item.Currency as Currency,
      Notes: item.Notes as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }
}
