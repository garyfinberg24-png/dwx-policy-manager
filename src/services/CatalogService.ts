// @ts-nocheck
/* eslint-disable @typescript-eslint/no-explicit-any */
// Catalog Service
// Catalog item management and vendor pricing
// Note: Some fields may not exist in the SharePoint list - mapping handles this gracefully

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  ICatalogItem,
  IVendorPricing,
  VendorCategory,
  UnitOfMeasure,
  Currency
} from '../models/IProcurement';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export interface ICatalogFilter {
  searchTerm?: string;
  category?: VendorCategory[];
  vendorId?: number;
  activeOnly?: boolean;
  minPrice?: number;
  maxPrice?: number;
}

export class CatalogService {
  private sp: SPFI;
  private readonly CATALOG_LIST = 'JML_CatalogItems';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==================== Catalog CRUD Operations ====================

  public async getCatalogItems(filter?: ICatalogFilter): Promise<ICatalogItem[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.CATALOG_LIST).items
        .select(
          'Id', 'Title', 'ItemCode', 'Category', 'SubCategory', 'Description',
          'UnitOfMeasure', 'IsActive', 'DefaultPrice', 'Currency',
          'MinOrderQuantity', 'MaxOrderQuantity', 'PreferredVendorId',
          'LeadTimeDays', 'Specifications', 'ImageUrl',
          'CreateAssetOnReceipt', 'AssetCategory', 'Notes', 'Tags',
          'Created', 'Modified', 'Author/Title', 'Editor/Title'
        )
        .expand('Author', 'Editor');

      // Apply filters
      if (filter) {
        const filters: string[] = [];

        if (filter.searchTerm) {
          const term = ValidationUtils.sanitizeForOData(filter.searchTerm);
          filters.push(`(substringof('${term}', Title) or substringof('${term}', ItemCode) or substringof('${term}', Description))`);
        }

        if (filter.category && filter.category.length > 0) {
          const categoryFilters = filter.category.map(c =>
            ValidationUtils.buildFilter('Category', 'eq', c)
          );
          filters.push(`(${categoryFilters.join(' or ')})`);
        }

        if (filter.vendorId !== undefined) {
          const validVendorId = ValidationUtils.validateInteger(filter.vendorId, 'vendorId', 1);
          filters.push(`PreferredVendorId eq ${validVendorId}`);
        }

        if (filter.activeOnly) {
          filters.push('IsActive eq 1');
        }

        if (filter.minPrice !== undefined) {
          filters.push(`DefaultPrice ge ${filter.minPrice}`);
        }

        if (filter.maxPrice !== undefined) {
          filters.push(`DefaultPrice le ${filter.maxPrice}`);
        }

        if (filters.length > 0) {
          query = query.filter(filters.join(' and '));
        }
      }

      const items = await query.orderBy('Category', true).orderBy('Title', true).top(5000)();
      return items.map(this.mapCatalogItemFromSP);
    } catch (error) {
      logger.error('CatalogService', 'Error getting catalog items:', error);
      throw error;
    }
  }

  public async getCatalogItemById(id: number): Promise<ICatalogItem> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const item = await this.sp.web.lists.getByTitle(this.CATALOG_LIST).items
        .getById(validId)
        .select(
          'Id', 'Title', 'ItemCode', 'Category', 'SubCategory', 'Description',
          'UnitOfMeasure', 'IsActive', 'DefaultPrice', 'Currency',
          'MinOrderQuantity', 'MaxOrderQuantity', 'PreferredVendorId',
          'LeadTimeDays', 'Specifications', 'ImageUrl',
          'CreateAssetOnReceipt', 'AssetCategory', 'AssetTypeId', 'Notes', 'Tags',
          'Created', 'Modified', 'Author/Title', 'Editor/Title'
        )
        .expand('Author', 'Editor')();

      return this.mapCatalogItemFromSP(item);
    } catch (error) {
      logger.error('CatalogService', 'Error getting catalog item by ID:', error);
      throw error;
    }
  }

  public async getCatalogItemByCode(itemCode: string): Promise<ICatalogItem | null> {
    try {
      if (!itemCode || typeof itemCode !== 'string') {
        throw new Error('Invalid item code');
      }

      const validCode = ValidationUtils.sanitizeForOData(itemCode.substring(0, 50));
      const filter = ValidationUtils.buildFilter('ItemCode', 'eq', validCode);

      const items = await this.sp.web.lists.getByTitle(this.CATALOG_LIST).items
        .select('Id', 'ItemCode')
        .filter(filter)
        .top(1)();

      if (items.length === 0) {
        return null;
      }

      return this.getCatalogItemById(items[0].Id);
    } catch (error) {
      logger.error('CatalogService', 'Error getting catalog item by code:', error);
      throw error;
    }
  }

  public async createCatalogItem(item: Partial<ICatalogItem>): Promise<number> {
    try {
      // Validate required fields
      if (!item.Title || !item.Category) {
        throw new Error('Title and Category are required');
      }

      // Generate item code
      const itemCode = item.ItemCode || await this.generateItemCode(item.Category);

      // Check if item code already exists
      const existing = await this.getCatalogItemByCode(itemCode);
      if (existing) {
        throw new Error(`Item code ${itemCode} already exists`);
      }

      const itemData: Record<string, unknown> = {
        Title: ValidationUtils.sanitizeHtml(item.Title),
        ItemCode: itemCode,
        Category: item.Category,
        UnitOfMeasure: item.UnitOfMeasure || UnitOfMeasure.Each,
        IsActive: item.IsActive !== undefined ? item.IsActive : true,
        DefaultPrice: item.DefaultPrice || 0,
        Currency: item.Currency || Currency.GBP,
        CreateAssetOnReceipt: item.CreateAssetOnReceipt || false
      };

      // Optional fields
      if (item.SubCategory) itemData.SubCategory = ValidationUtils.sanitizeHtml(item.SubCategory);
      if (item.Description) itemData.Description = ValidationUtils.sanitizeHtml(item.Description);
      if (item.MinOrderQuantity !== undefined) itemData.MinOrderQuantity = item.MinOrderQuantity;
      if (item.MaxOrderQuantity !== undefined) itemData.MaxOrderQuantity = item.MaxOrderQuantity;
      if (item.PreferredVendorId) itemData.PreferredVendorId = ValidationUtils.validateInteger(item.PreferredVendorId, 'PreferredVendorId', 1);
      if (item.LeadTimeDays !== undefined) itemData.LeadTimeDays = item.LeadTimeDays;
      if (item.Specifications) itemData.Specifications = ValidationUtils.sanitizeHtml(item.Specifications);
      if (item.ImageUrl) itemData.ImageUrl = ValidationUtils.sanitizeHtml(item.ImageUrl);
      if (item.AssetCategory) itemData.AssetCategory = ValidationUtils.sanitizeHtml(item.AssetCategory);
      if (item.AssetTypeId) itemData.AssetTypeId = ValidationUtils.validateInteger(item.AssetTypeId, 'AssetTypeId', 1);
      if (item.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(item.Notes);
      if (item.Tags) itemData.Tags = item.Tags;

      const result = await this.sp.web.lists.getByTitle(this.CATALOG_LIST).items.add(itemData);
      return result.data.Id;
    } catch (error) {
      logger.error('CatalogService', 'Error creating catalog item:', error);
      throw error;
    }
  }

  public async updateCatalogItem(id: number, updates: Partial<ICatalogItem>): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const itemData: Record<string, unknown> = {};

      if (updates.Title) itemData.Title = ValidationUtils.sanitizeHtml(updates.Title);
      if (updates.Category) {
        ValidationUtils.validateEnum(updates.Category, VendorCategory, 'Category');
        itemData.Category = updates.Category;
      }
      if (updates.SubCategory !== undefined) itemData.SubCategory = updates.SubCategory ? ValidationUtils.sanitizeHtml(updates.SubCategory) : null;
      if (updates.Description !== undefined) itemData.Description = updates.Description ? ValidationUtils.sanitizeHtml(updates.Description) : null;
      if (updates.UnitOfMeasure) itemData.UnitOfMeasure = updates.UnitOfMeasure;
      if (updates.IsActive !== undefined) itemData.IsActive = updates.IsActive;
      if (updates.DefaultPrice !== undefined) itemData.DefaultPrice = updates.DefaultPrice;
      if (updates.Currency) itemData.Currency = updates.Currency;
      if (updates.MinOrderQuantity !== undefined) itemData.MinOrderQuantity = updates.MinOrderQuantity;
      if (updates.MaxOrderQuantity !== undefined) itemData.MaxOrderQuantity = updates.MaxOrderQuantity;
      if (updates.PreferredVendorId !== undefined) {
        itemData.PreferredVendorId = updates.PreferredVendorId === null ? null :
          ValidationUtils.validateInteger(updates.PreferredVendorId, 'PreferredVendorId', 1);
      }
      if (updates.LeadTimeDays !== undefined) itemData.LeadTimeDays = updates.LeadTimeDays;
      if (updates.Specifications !== undefined) itemData.Specifications = updates.Specifications ? ValidationUtils.sanitizeHtml(updates.Specifications) : null;
      if (updates.ImageUrl !== undefined) itemData.ImageUrl = updates.ImageUrl ? ValidationUtils.sanitizeHtml(updates.ImageUrl) : null;
      if (updates.CreateAssetOnReceipt !== undefined) itemData.CreateAssetOnReceipt = updates.CreateAssetOnReceipt;
      if (updates.AssetCategory !== undefined) itemData.AssetCategory = updates.AssetCategory ? ValidationUtils.sanitizeHtml(updates.AssetCategory) : null;
      if (updates.AssetTypeId !== undefined) {
        itemData.AssetTypeId = updates.AssetTypeId === null ? null :
          ValidationUtils.validateInteger(updates.AssetTypeId, 'AssetTypeId', 1);
      }
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes ? ValidationUtils.sanitizeHtml(updates.Notes) : null;
      if (updates.Tags !== undefined) itemData.Tags = updates.Tags;

      await this.sp.web.lists.getByTitle(this.CATALOG_LIST).items.getById(validId).update(itemData);
    } catch (error) {
      logger.error('CatalogService', 'Error updating catalog item:', error);
      throw error;
    }
  }

  public async deleteCatalogItem(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Soft delete - mark as inactive
      await this.updateCatalogItem(validId, { IsActive: false });
    } catch (error) {
      logger.error('CatalogService', 'Error deleting catalog item:', error);
      throw error;
    }
  }

  public async activateCatalogItem(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);
      await this.updateCatalogItem(validId, { IsActive: true });
    } catch (error) {
      logger.error('CatalogService', 'Error activating catalog item:', error);
      throw error;
    }
  }

  // ==================== Query Methods ====================

  public async getActiveItems(): Promise<ICatalogItem[]> {
    try {
      return this.getCatalogItems({ activeOnly: true });
    } catch (error) {
      logger.error('CatalogService', 'Error getting active items:', error);
      throw error;
    }
  }

  public async getItemsByCategory(category: VendorCategory): Promise<ICatalogItem[]> {
    try {
      return this.getCatalogItems({ category: [category], activeOnly: true });
    } catch (error) {
      logger.error('CatalogService', 'Error getting items by category:', error);
      throw error;
    }
  }

  public async getVendorItems(vendorId: number): Promise<ICatalogItem[]> {
    try {
      return this.getCatalogItems({ vendorId, activeOnly: true });
    } catch (error) {
      logger.error('CatalogService', 'Error getting vendor items:', error);
      throw error;
    }
  }

  public async searchItems(searchTerm: string): Promise<ICatalogItem[]> {
    try {
      return this.getCatalogItems({ searchTerm, activeOnly: true });
    } catch (error) {
      logger.error('CatalogService', 'Error searching items:', error);
      throw error;
    }
  }

  public async getAssetItems(): Promise<ICatalogItem[]> {
    try {
      const items = await this.getActiveItems();
      return items.filter(item => item.CreateAssetOnReceipt);
    } catch (error) {
      logger.error('CatalogService', 'Error getting asset items:', error);
      throw error;
    }
  }

  // ==================== Statistics ====================

  public async getCatalogStatistics(): Promise<{
    total: number;
    active: number;
    inactive: number;
    byCategory: { [key: string]: number };
    withVendor: number;
    withAssetCreation: number;
    avgPrice: number;
  }> {
    try {
      const items = await this.getCatalogItems();

      const stats = {
        total: items.length,
        active: 0,
        inactive: 0,
        byCategory: {} as { [key: string]: number },
        withVendor: 0,
        withAssetCreation: 0,
        avgPrice: 0
      };

      let totalPrice = 0;
      let priceCount = 0;

      for (const item of items) {
        // Count active/inactive
        if (item.IsActive) {
          stats.active++;
        } else {
          stats.inactive++;
        }

        // Count by category
        stats.byCategory[item.Category] = (stats.byCategory[item.Category] || 0) + 1;

        // Count with vendor
        if (item.PreferredVendorId) {
          stats.withVendor++;
        }

        // Count with asset creation
        if (item.CreateAssetOnReceipt) {
          stats.withAssetCreation++;
        }

        // Calculate average price
        if (item.DefaultPrice && item.DefaultPrice > 0) {
          totalPrice += item.DefaultPrice;
          priceCount++;
        }
      }

      stats.avgPrice = priceCount > 0 ? totalPrice / priceCount : 0;

      return stats;
    } catch (error) {
      logger.error('CatalogService', 'Error getting catalog statistics:', error);
      throw error;
    }
  }

  // ==================== Helper Functions ====================

  private async generateItemCode(category: VendorCategory): Promise<string> {
    try {
      // Create category prefix map
      const categoryPrefixes: { [key in VendorCategory]: string } = {
        [VendorCategory.ITHardware]: 'HW',
        [VendorCategory.ITSoftware]: 'SW',
        [VendorCategory.ITServices]: 'SV',
        [VendorCategory.OfficeSupplies]: 'OS',
        [VendorCategory.Furniture]: 'FN',
        [VendorCategory.ProfessionalServices]: 'PS',
        [VendorCategory.Utilities]: 'UT',
        [VendorCategory.Marketing]: 'MK',
        [VendorCategory.Travel]: 'TR',
        [VendorCategory.Facilities]: 'FC',
        [VendorCategory.Telecommunications]: 'TC',
        [VendorCategory.Security]: 'SC',
        [VendorCategory.Training]: 'TN',
        [VendorCategory.Catering]: 'CT',
        [VendorCategory.Other]: 'OT'
      };

      const prefix = categoryPrefixes[category] || 'OT';

      const items = await this.sp.web.lists.getByTitle(this.CATALOG_LIST).items
        .select('ItemCode')
        .filter(`substringof('${prefix}-', ItemCode)`)
        .orderBy('Id', false)
        .top(1)();

      let nextNumber = 1;
      if (items.length > 0 && items[0].ItemCode) {
        const match = items[0].ItemCode.match(new RegExp(`${prefix}-(\\d+)`));
        if (match) {
          nextNumber = parseInt(match[1], 10) + 1;
        }
      }

      return `${prefix}-${nextNumber.toString().padStart(5, '0')}`;
    } catch (error) {
      logger.error('CatalogService', 'Error generating item code:', error);
      return `ITM-${Date.now()}`;
    }
  }

  // ==================== Mapping Functions ====================

  private mapCatalogItemFromSP(item: Record<string, unknown>): ICatalogItem {
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      ItemCode: item.ItemCode as string,
      Category: item.Category as VendorCategory,
      SubCategory: item.SubCategory as string,
      Description: item.Description as string,
      UnitOfMeasure: item.UnitOfMeasure as UnitOfMeasure || UnitOfMeasure.Each,
      IsActive: item.IsActive as boolean,
      DefaultPrice: item.DefaultPrice as number || 0,
      Currency: item.Currency as Currency || Currency.GBP,
      MinOrderQuantity: item.MinOrderQuantity as number,
      MaxOrderQuantity: item.MaxOrderQuantity as number,
      PreferredVendorId: item.PreferredVendorId as number,
      LeadTimeDays: item.LeadTimeDays as number,
      Specifications: item.Specifications as string,
      ImageUrl: item.ImageUrl as string,
      CreateAssetOnReceipt: item.CreateAssetOnReceipt as boolean || false,
      AssetCategory: item.AssetCategory as string,
      AssetTypeId: item.AssetTypeId as number,
      Notes: item.Notes as string,
      Tags: item.Tags as string,
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }
}
