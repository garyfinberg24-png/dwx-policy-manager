// @ts-nocheck
// Asset Service
// Core service for asset management operations

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  IAsset,
  IAssetType,
  IAssetAssignment,
  IAssetCheckout,
  IAssetMaintenance,
  IAssetTransfer,
  IAssetAudit,
  IAssetAuditItem,
  IAssetRequest,
  IAssetStatistics,
  IAssetFilterCriteria,
  IAssetHistoryEntry,
  AssetStatus,
  AssetCategory,
  AssetCondition,
  CheckoutStatus,
  MaintenanceType
} from '../models/IAsset';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class AssetService {
  private sp: SPFI;
  private readonly ASSETS_LIST = 'Assets';
  private readonly ASSET_TYPES_LIST = 'Asset Types';
  private readonly ASSET_ASSIGNMENTS_LIST = 'Asset Assignments';
  private readonly ASSET_CHECKOUTS_LIST = 'Asset Checkouts';
  private readonly ASSET_MAINTENANCE_LIST = 'Asset Maintenance';
  private readonly ASSET_TRANSFERS_LIST = 'Asset Transfers';
  private readonly ASSET_AUDITS_LIST = 'Asset Audits';
  private readonly ASSET_AUDIT_ITEMS_LIST = 'Asset Audit Items';
  private readonly ASSET_REQUESTS_LIST = 'Asset Requests';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==================== Asset CRUD Operations ====================

  public async getAssets(filter?: IAssetFilterCriteria): Promise<IAsset[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.ASSETS_LIST).items
        .select(
          'Id', 'Title', 'AssetTag', 'Barcode', 'SerialNumber',
          'AssetTypeId', 'Category', 'Status', 'Condition',
          'AssignedToId', 'AssignedTo/Title', 'AssignedTo/EMail', 'AssignedDate', 'AssignedById',
          'Location', 'Department', 'CostCenter',
          'PurchaseDate', 'PurchaseCost', 'CurrentValue', 'DepreciationMethod', 'DepreciationRate',
          'SalvageValue', 'WarrantyExpiration',
          'Vendor', 'PurchaseOrderNumber', 'InvoiceNumber',
          'LastMaintenanceDate', 'NextMaintenanceDate', 'MaintenanceSchedule',
          'LicenseKey', 'LicenseExpiration', 'LicenseType', 'MaxLicenses', 'CurrentLicensesUsed',
          'Manufacturer', 'Model', 'Specifications', 'IPAddress', 'MACAddress', 'HostName',
          'RetirementDate', 'RetirementReason', 'DisposalDate', 'DisposalMethod',
          'Comments', 'Attachments',
          'Created', 'CreatedById', 'Modified', 'ModifiedById'
        )
        .expand('AssignedTo');

      // Apply filters
      if (filter) {
        const filters: string[] = [];

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

        if (filter.condition && filter.condition.length > 0) {
          const conditionFilters = filter.condition.map(c =>
            ValidationUtils.buildFilter('Condition', 'eq', c)
          );
          filters.push(`(${conditionFilters.join(' or ')})`);
        }

        if (filter.assignedToId !== undefined) {
          const validUserId = ValidationUtils.validateInteger(filter.assignedToId, 'assignedToId', 1);
          filters.push(ValidationUtils.buildFilter('AssignedToId', 'eq', validUserId));
        }

        if (filter.location) {
          const validLocation = ValidationUtils.sanitizeForOData(filter.location);
          filters.push(`substringof('${validLocation}', Location)`);
        }

        if (filter.department) {
          const validDept = ValidationUtils.sanitizeForOData(filter.department);
          filters.push(ValidationUtils.buildFilter('Department', 'eq', validDept));
        }

        if (filter.assetTypeId !== undefined) {
          const validTypeId = ValidationUtils.validateInteger(filter.assetTypeId, 'assetTypeId', 1);
          filters.push(ValidationUtils.buildFilter('AssetTypeId', 'eq', validTypeId));
        }

        if (filter.manufacturer) {
          const validManuf = ValidationUtils.sanitizeForOData(filter.manufacturer);
          filters.push(ValidationUtils.buildFilter('Manufacturer', 'eq', validManuf));
        }

        if (filter.searchTerm) {
          const validTerm = ValidationUtils.sanitizeForOData(filter.searchTerm);
          filters.push(`(substringof('${validTerm}', Title) or substringof('${validTerm}', AssetTag) or substringof('${validTerm}', SerialNumber))`);
        }

        if (filter.isAvailable) {
          filters.push(ValidationUtils.buildFilter('Status', 'eq', AssetStatus.Available));
        }

        if (filter.hasWarranty) {
          const now = new Date();
          filters.push(ValidationUtils.buildFilter('WarrantyExpiration', 'ge', now));
        }

        if (filter.needsMaintenance) {
          const now = new Date();
          filters.push(ValidationUtils.buildFilter('NextMaintenanceDate', 'le', now));
        }

        if (filters.length > 0) {
          query = query.filter(filters.join(' and '));
        }
      }

      const items = await query.top(5000)();
      return items.map(this.mapAssetFromSP);
    } catch (error) {
      logger.error('AssetService', 'Error getting assets:', error);
      throw error;
    }
  }

  public async getAssetById(id: number): Promise<IAsset> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const item = await this.sp.web.lists.getByTitle(this.ASSETS_LIST).items
        .getById(validId)
        .select(
          'Id', 'Title', 'AssetTag', 'Barcode', 'SerialNumber',
          'AssetTypeId', 'Category', 'Status', 'Condition',
          'AssignedToId', 'AssignedTo/Title', 'AssignedTo/EMail', 'AssignedDate', 'AssignedById',
          'Location', 'Department', 'CostCenter',
          'PurchaseDate', 'PurchaseCost', 'CurrentValue', 'DepreciationMethod', 'DepreciationRate',
          'SalvageValue', 'WarrantyExpiration',
          'Vendor', 'PurchaseOrderNumber', 'InvoiceNumber',
          'LastMaintenanceDate', 'NextMaintenanceDate', 'MaintenanceSchedule',
          'LicenseKey', 'LicenseExpiration', 'LicenseType', 'MaxLicenses', 'CurrentLicensesUsed',
          'Manufacturer', 'Model', 'Specifications', 'IPAddress', 'MACAddress', 'HostName',
          'RetirementDate', 'RetirementReason', 'DisposalDate', 'DisposalMethod',
          'Comments', 'Attachments',
          'Created', 'CreatedById', 'Modified', 'ModifiedById'
        )
        .expand('AssignedTo')();

      return this.mapAssetFromSP(item);
    } catch (error) {
      logger.error('AssetService', 'Error getting asset by ID:', error);
      throw error;
    }
  }

  public async getAssetByTag(assetTag: string): Promise<IAsset | null> {
    try {
      if (!assetTag || typeof assetTag !== 'string') {
        throw new Error('Invalid asset tag');
      }

      const validTag = ValidationUtils.sanitizeForOData(assetTag.substring(0, 50));
      const filter = ValidationUtils.buildFilter('AssetTag', 'eq', validTag);

      const items = await this.sp.web.lists.getByTitle(this.ASSETS_LIST).items
        .select(
          'Id', 'Title', 'AssetTag', 'Barcode', 'SerialNumber',
          'AssetTypeId', 'Category', 'Status', 'Condition',
          'AssignedToId', 'AssignedTo/Title', 'AssignedTo/EMail',
          'Location', 'Department'
        )
        .expand('AssignedTo')
        .filter(filter)
        .top(1)();

      if (items.length === 0) {
        return null;
      }

      return this.mapAssetFromSP(items[0]);
    } catch (error) {
      logger.error('AssetService', 'Error getting asset by tag:', error);
      throw error;
    }
  }

  public async createAsset(asset: Partial<IAsset>): Promise<number> {
    try {
      // Validate required fields
      if (!asset.Title || !asset.AssetTag || !asset.AssetTypeId) {
        throw new Error('Title, AssetTag, and AssetTypeId are required');
      }

      ValidationUtils.validateEnum(asset.Status, AssetStatus, 'Status');
      ValidationUtils.validateEnum(asset.Category, AssetCategory, 'Category');
      ValidationUtils.validateEnum(asset.Condition, AssetCondition, 'Condition');

      // Check if asset tag already exists
      const existing = await this.getAssetByTag(asset.AssetTag);
      if (existing) {
        throw new Error(`Asset tag ${asset.AssetTag} already exists`);
      }

      const itemData: any = {
        Title: ValidationUtils.sanitizeHtml(asset.Title),
        AssetTag: ValidationUtils.sanitizeHtml(asset.AssetTag),
        AssetTypeId: ValidationUtils.validateInteger(asset.AssetTypeId, 'AssetTypeId', 1),
        Category: asset.Category,
        Status: asset.Status || AssetStatus.Available,
        Condition: asset.Condition || AssetCondition.New
      };

      // Optional fields
      if (asset.Barcode) itemData.Barcode = ValidationUtils.sanitizeHtml(asset.Barcode);
      if (asset.SerialNumber) itemData.SerialNumber = ValidationUtils.sanitizeHtml(asset.SerialNumber);
      if (asset.Location) itemData.Location = ValidationUtils.sanitizeHtml(asset.Location);
      if (asset.Department) itemData.Department = ValidationUtils.sanitizeHtml(asset.Department);
      if (asset.CostCenter) itemData.CostCenter = ValidationUtils.sanitizeHtml(asset.CostCenter);
      if (asset.PurchaseDate) itemData.PurchaseDate = ValidationUtils.validateDate(asset.PurchaseDate, 'PurchaseDate');
      if (asset.PurchaseCost !== undefined) itemData.PurchaseCost = ValidationUtils.validateInteger(asset.PurchaseCost, 'PurchaseCost', 0);
      if (asset.Vendor) itemData.Vendor = ValidationUtils.sanitizeHtml(asset.Vendor);
      if (asset.Manufacturer) itemData.Manufacturer = ValidationUtils.sanitizeHtml(asset.Manufacturer);
      if (asset.Model) itemData.Model = ValidationUtils.sanitizeHtml(asset.Model);
      if (asset.Specifications) itemData.Specifications = ValidationUtils.sanitizeHtml(asset.Specifications);
      if (asset.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(asset.Notes);
      if (asset.WarrantyExpiration) itemData.WarrantyExpiration = ValidationUtils.validateDate(asset.WarrantyExpiration, 'WarrantyExpiration');
      if (asset.LicenseKey) itemData.LicenseKey = ValidationUtils.sanitizeHtml(asset.LicenseKey);
      if (asset.LicenseExpiration) itemData.LicenseExpiration = ValidationUtils.validateDate(asset.LicenseExpiration, 'LicenseExpiration');
      if (asset.MaxLicenses) itemData.MaxLicenses = ValidationUtils.validateInteger(asset.MaxLicenses, 'MaxLicenses', 1);

      const result = await this.sp.web.lists.getByTitle(this.ASSETS_LIST).items.add(itemData);
      return result.data.Id;
    } catch (error) {
      logger.error('AssetService', 'Error creating asset:', error);
      throw error;
    }
  }

  public async updateAsset(id: number, updates: Partial<IAsset>): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const itemData: any = {};

      if (updates.Title) itemData.Title = ValidationUtils.sanitizeHtml(updates.Title);
      if (updates.Status) {
        ValidationUtils.validateEnum(updates.Status, AssetStatus, 'Status');
        itemData.Status = updates.Status;
      }
      if (updates.Condition) {
        ValidationUtils.validateEnum(updates.Condition, AssetCondition, 'Condition');
        itemData.Condition = updates.Condition;
      }
      if (updates.AssignedToId !== undefined) {
        itemData.AssignedToId = updates.AssignedToId === null ? null :
          ValidationUtils.validateInteger(updates.AssignedToId, 'AssignedToId', 1);
      }
      if (updates.AssignedDate) itemData.AssignedDate = ValidationUtils.validateDate(updates.AssignedDate, 'AssignedDate');
      if (updates.Location !== undefined) itemData.Location = updates.Location ? ValidationUtils.sanitizeHtml(updates.Location) : null;
      if (updates.Department !== undefined) itemData.Department = updates.Department ? ValidationUtils.sanitizeHtml(updates.Department) : null;
      if (updates.CurrentValue !== undefined) itemData.CurrentValue = ValidationUtils.validateInteger(updates.CurrentValue, 'CurrentValue', 0);
      if (updates.LastMaintenanceDate) itemData.LastMaintenanceDate = ValidationUtils.validateDate(updates.LastMaintenanceDate, 'LastMaintenanceDate');
      if (updates.NextMaintenanceDate) itemData.NextMaintenanceDate = ValidationUtils.validateDate(updates.NextMaintenanceDate, 'NextMaintenanceDate');
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes ? ValidationUtils.sanitizeHtml(updates.Notes) : null;
      if (updates.IPAddress) itemData.IPAddress = ValidationUtils.sanitizeHtml(updates.IPAddress);
      if (updates.MACAddress) itemData.MACAddress = ValidationUtils.sanitizeHtml(updates.MACAddress);
      if (updates.HostName) itemData.HostName = ValidationUtils.sanitizeHtml(updates.HostName);

      await this.sp.web.lists.getByTitle(this.ASSETS_LIST).items.getById(validId).update(itemData);
    } catch (error) {
      logger.error('AssetService', 'Error updating asset:', error);
      throw error;
    }
  }

  public async deleteAsset(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Check if asset is currently assigned
      const asset = await this.getAssetById(validId);
      if (asset.Status === AssetStatus.Assigned) {
        throw new Error('Cannot delete an asset that is currently assigned');
      }

      await this.sp.web.lists.getByTitle(this.ASSETS_LIST).items.getById(validId).delete();
    } catch (error) {
      logger.error('AssetService', 'Error deleting asset:', error);
      throw error;
    }
  }

  // ==================== Asset Assignment Operations ====================

  public async assignAsset(assetId: number, userId: number, assignedById: number, notes?: string): Promise<number> {
    try {
      const validAssetId = ValidationUtils.validateInteger(assetId, 'assetId', 1);
      const validUserId = ValidationUtils.validateInteger(userId, 'userId', 1);
      const validAssignedById = ValidationUtils.validateInteger(assignedById, 'assignedById', 1);

      // Check if asset is available
      const asset = await this.getAssetById(validAssetId);
      if (asset.Status !== AssetStatus.Available && asset.Status !== AssetStatus.Reserved) {
        throw new Error(`Asset ${asset.AssetTag} is not available for assignment`);
      }

      // Create assignment record
      const assignmentData: any = {
        AssetId: validAssetId,
        AssignedToId: validUserId,
        AssignedById: validAssignedById,
        AssignedDate: new Date().toISOString(),
        Status: CheckoutStatus.CheckedOut,
        IsActive: true
      };

      if (notes) {
        assignmentData.Notes = ValidationUtils.sanitizeHtml(notes);
      }

      const assignmentResult = await this.sp.web.lists.getByTitle(this.ASSET_ASSIGNMENTS_LIST).items.add(assignmentData);

      // Update asset status
      await this.updateAsset(validAssetId, {
        Status: AssetStatus.Assigned,
        AssignedToId: validUserId,
        AssignedDate: new Date(),
        AssignedById: validAssignedById
      });

      return assignmentResult.data.Id;
    } catch (error) {
      logger.error('AssetService', 'Error assigning asset:', error);
      throw error;
    }
  }

  public async unassignAsset(assetId: number, returnCondition?: AssetCondition, returnNotes?: string): Promise<void> {
    try {
      const validAssetId = ValidationUtils.validateInteger(assetId, 'assetId', 1);

      // Get active assignment
      const filter = `${ValidationUtils.buildFilter('AssetId', 'eq', validAssetId)} and IsActive eq 1`;
      const assignments = await this.sp.web.lists.getByTitle(this.ASSET_ASSIGNMENTS_LIST).items
        .filter(filter)
        .top(1)();

      if (assignments.length === 0) {
        throw new Error('No active assignment found for this asset');
      }

      const assignment = assignments[0];

      // Update assignment record
      const updateData: any = {
        ActualReturnDate: new Date().toISOString(),
        Status: CheckoutStatus.CheckedIn,
        IsActive: false
      };

      if (returnCondition) {
        ValidationUtils.validateEnum(returnCondition, AssetCondition, 'returnCondition');
        updateData.ReturnCondition = returnCondition;
      }

      if (returnNotes) {
        updateData.ReturnNotes = ValidationUtils.sanitizeHtml(returnNotes);
      }

      await this.sp.web.lists.getByTitle(this.ASSET_ASSIGNMENTS_LIST).items.getById(assignment.Id).update(updateData);

      // Update asset status
      const assetUpdates: Partial<IAsset> = {
        Status: AssetStatus.Available,
        AssignedToId: undefined,
        AssignedDate: undefined
      };

      if (returnCondition) {
        assetUpdates.Condition = returnCondition;
      }

      await this.updateAsset(validAssetId, assetUpdates);
    } catch (error) {
      logger.error('AssetService', 'Error unassigning asset:', error);
      throw error;
    }
  }

  public async getAssetAssignments(assetId?: number, userId?: number): Promise<IAssetAssignment[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.ASSET_ASSIGNMENTS_LIST).items
        .select(
          'Id', 'AssetId', 'AssignedToId', 'AssignedTo/Title', 'AssignedTo/EMail',
          'AssignedById', 'AssignedBy/Title', 'AssignedDate', 'ExpectedReturnDate',
          'ActualReturnDate', 'ProcessId', 'TaskId', 'AssignmentReason', 'AssignedLocation',
          'Status', 'IsActive', 'ReturnCondition', 'ReturnNotes', 'Comments',
          'Created', 'Modified'
        )
        .expand('AssignedTo', 'AssignedBy')
        .orderBy('AssignedDate', false);

      const filters: string[] = [];

      if (assetId !== undefined) {
        const validAssetId = ValidationUtils.validateInteger(assetId, 'assetId', 1);
        filters.push(ValidationUtils.buildFilter('AssetId', 'eq', validAssetId));
      }

      if (userId !== undefined) {
        const validUserId = ValidationUtils.validateInteger(userId, 'userId', 1);
        filters.push(ValidationUtils.buildFilter('AssignedToId', 'eq', validUserId));
      }

      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      const items = await query.top(1000)();
      return items.map(this.mapAssetAssignmentFromSP);
    } catch (error) {
      logger.error('AssetService', 'Error getting asset assignments:', error);
      throw error;
    }
  }

  // ==================== Asset Types ====================

  public async getAssetTypes(activeOnly: boolean = true): Promise<IAssetType[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.ASSET_TYPES_LIST).items
        .select(
          'Id', 'Title', 'Category', 'Description', 'Manufacturer', 'Model',
          'DefaultSpecs', 'DefaultDepreciationMethod', 'DefaultDepreciationRate',
          'DefaultWarrantyPeriod', 'DefaultMaintenanceSchedule', 'IsActive',
          'Created', 'Modified'
        );

      if (activeOnly) {
        query = query.filter('IsActive eq 1');
      }

      const items = await query.top(500)();
      return items.map(this.mapAssetTypeFromSP);
    } catch (error) {
      logger.error('AssetService', 'Error getting asset types:', error);
      throw error;
    }
  }

  public async createAssetType(assetType: Partial<IAssetType>): Promise<number> {
    try {
      if (!assetType.Title || !assetType.Category) {
        throw new Error('Title and Category are required');
      }

      ValidationUtils.validateEnum(assetType.Category, AssetCategory, 'Category');

      const itemData: any = {
        Title: ValidationUtils.sanitizeHtml(assetType.Title),
        Category: assetType.Category,
        IsActive: assetType.IsActive !== undefined ? assetType.IsActive : true
      };

      if (assetType.Description) itemData.Description = ValidationUtils.sanitizeHtml(assetType.Description);
      if (assetType.Manufacturer) itemData.Manufacturer = ValidationUtils.sanitizeHtml(assetType.Manufacturer);
      if (assetType.Model) itemData.Model = ValidationUtils.sanitizeHtml(assetType.Model);
      if (assetType.DefaultSpecs) itemData.DefaultSpecs = ValidationUtils.sanitizeHtml(assetType.DefaultSpecs);

      const result = await this.sp.web.lists.getByTitle(this.ASSET_TYPES_LIST).items.add(itemData);
      return result.data.Id;
    } catch (error) {
      logger.error('AssetService', 'Error creating asset type:', error);
      throw error;
    }
  }

  // ==================== Statistics and Reporting ====================

  public async getAssetStatistics(): Promise<IAssetStatistics> {
    try {
      const allAssets = await this.getAssets();

      const stats: IAssetStatistics = {
        totalAssets: allAssets.length,
        totalValue: allAssets.reduce((sum, a) => sum + (a.CurrentValue || a.PurchaseCost || 0), 0),
        byStatus: {},
        byCategory: {},
        byCondition: {},
        assignedAssets: allAssets.filter(a => a.Status === AssetStatus.Assigned).length,
        availableAssets: allAssets.filter(a => a.Status === AssetStatus.Available).length,
        inMaintenanceAssets: allAssets.filter(a => a.Status === AssetStatus.InMaintenance).length,
        overdueCheckouts: 0,
        expiringSoonLicenses: 0,
        upcomingMaintenance: 0,
        recentActivity: {
          newAssignments: 0,
          recentCheckouts: 0,
          completedMaintenance: 0
        }
      };

      // Count by status
      for (const status of Object.values(AssetStatus)) {
        stats.byStatus[status] = allAssets.filter(a => a.Status === status).length;
      }

      // Count by category
      for (const category of Object.values(AssetCategory)) {
        stats.byCategory[category] = allAssets.filter(a => a.Category === category).length;
      }

      // Count by condition
      for (const condition of Object.values(AssetCondition)) {
        stats.byCondition[condition] = allAssets.filter(a => a.Condition === condition).length;
      }

      // Expiring licenses (within 30 days)
      const now = new Date();
      const thirtyDaysFromNow = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
      stats.expiringSoonLicenses = allAssets.filter(a =>
        a.LicenseExpiration &&
        new Date(a.LicenseExpiration) >= now &&
        new Date(a.LicenseExpiration) <= thirtyDaysFromNow
      ).length;

      // Upcoming maintenance (within 30 days)
      stats.upcomingMaintenance = allAssets.filter(a =>
        a.NextMaintenanceDate &&
        new Date(a.NextMaintenanceDate) >= now &&
        new Date(a.NextMaintenanceDate) <= thirtyDaysFromNow
      ).length;

      return stats;
    } catch (error) {
      logger.error('AssetService', 'Error getting asset statistics:', error);
      throw error;
    }
  }

  // ==================== Mapping Functions ====================

  private mapAssetFromSP(item: any): IAsset {
    return {
      Id: item.Id,
      Title: item.Title,
      AssetTag: item.AssetTag,
      Barcode: item.Barcode,
      SerialNumber: item.SerialNumber,
      AssetTypeId: item.AssetTypeId,
      Category: item.Category as AssetCategory,
      Status: item.Status as AssetStatus,
      Condition: item.Condition as AssetCondition,
      AssignedToId: item.AssignedToId,
      AssignedTo: item.AssignedTo,
      AssignedDate: item.AssignedDate ? new Date(item.AssignedDate) : undefined,
      AssignedById: item.AssignedById,
      Location: item.Location,
      Department: item.Department,
      CostCenter: item.CostCenter,
      PurchaseDate: item.PurchaseDate ? new Date(item.PurchaseDate) : undefined,
      PurchaseCost: item.PurchaseCost,
      CurrentValue: item.CurrentValue,
      DepreciationMethod: item.DepreciationMethod,
      DepreciationRate: item.DepreciationRate,
      SalvageValue: item.SalvageValue,
      WarrantyExpiration: item.WarrantyExpiration ? new Date(item.WarrantyExpiration) : undefined,
      Vendor: item.Vendor,
      PurchaseOrderNumber: item.PurchaseOrderNumber,
      InvoiceNumber: item.InvoiceNumber,
      LastMaintenanceDate: item.LastMaintenanceDate ? new Date(item.LastMaintenanceDate) : undefined,
      NextMaintenanceDate: item.NextMaintenanceDate ? new Date(item.NextMaintenanceDate) : undefined,
      MaintenanceSchedule: item.MaintenanceSchedule,
      LicenseKey: item.LicenseKey,
      LicenseExpiration: item.LicenseExpiration ? new Date(item.LicenseExpiration) : undefined,
      LicenseType: item.LicenseType,
      MaxLicenses: item.MaxLicenses,
      CurrentLicensesUsed: item.CurrentLicensesUsed,
      Manufacturer: item.Manufacturer,
      Model: item.Model,
      Specifications: item.Specifications,
      IPAddress: item.IPAddress,
      MACAddress: item.MACAddress,
      HostName: item.HostName,
      RetirementDate: item.RetirementDate ? new Date(item.RetirementDate) : undefined,
      RetirementReason: item.RetirementReason,
      DisposalDate: item.DisposalDate ? new Date(item.DisposalDate) : undefined,
      DisposalMethod: item.DisposalMethod,
      Notes: item.Notes,
      Attachments: item.Attachments,
      Created: item.Created ? new Date(item.Created) : undefined,
      CreatedById: item.CreatedById,
      Modified: item.Modified ? new Date(item.Modified) : undefined,
      ModifiedById: item.ModifiedById
    };
  }

  private mapAssetTypeFromSP(item: any): IAssetType {
    return {
      Id: item.Id,
      Title: item.Title,
      Category: item.Category as AssetCategory,
      Description: item.Description,
      Manufacturer: item.Manufacturer,
      Model: item.Model,
      DefaultSpecs: item.DefaultSpecs,
      DefaultDepreciationMethod: item.DefaultDepreciationMethod,
      DefaultDepreciationRate: item.DefaultDepreciationRate,
      DefaultWarrantyPeriod: item.DefaultWarrantyPeriod,
      DefaultMaintenanceSchedule: item.DefaultMaintenanceSchedule,
      IsActive: item.IsActive,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }

  private mapAssetAssignmentFromSP(item: any): IAssetAssignment {
    return {
      Id: item.Id,
      AssetId: item.AssetId,
      AssignedToId: item.AssignedToId,
      AssignedTo: item.AssignedTo,
      AssignedById: item.AssignedById,
      AssignedBy: item.AssignedBy,
      AssignedDate: new Date(item.AssignedDate),
      ExpectedReturnDate: item.ExpectedReturnDate ? new Date(item.ExpectedReturnDate) : undefined,
      ActualReturnDate: item.ActualReturnDate ? new Date(item.ActualReturnDate) : undefined,
      ProcessId: item.ProcessId,
      TaskId: item.TaskId,
      AssignmentReason: item.AssignmentReason,
      AssignedLocation: item.AssignedLocation,
      Status: item.Status as CheckoutStatus,
      IsActive: item.IsActive,
      ReturnCondition: item.ReturnCondition as AssetCondition,
      ReturnNotes: item.ReturnNotes,
      Notes: item.Notes,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }
}
