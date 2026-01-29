// @ts-nocheck
/* eslint-disable @typescript-eslint/no-explicit-any */
// TODO: Fix 'Description' property type mismatch with IContract interface
// Contract Service
// Contract lifecycle management, renewals, and compliance tracking
// Note: Some fields may not exist in the SharePoint list - mapping handles this gracefully

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  IContract,
  IContractFilter,
  ContractStatus,
  ContractType,
  Currency,
  PaymentTerms
} from '../models/IProcurement';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class ContractService {
  private sp: SPFI;
  private readonly CONTRACTS_LIST = 'PM_Contracts';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==================== Contract CRUD Operations ====================

  public async getContracts(filter?: IContractFilter): Promise<IContract[]> {
    try {
      console.log('[ContractService] Fetching contracts from list:', this.CONTRACTS_LIST);
      // NOTE: Removed Person field expands (Author, Editor) - cause 400 errors when columns don't exist
      let query = this.sp.web.lists.getByTitle(this.CONTRACTS_LIST).items
        .select(
          'Id', 'Title', 'ContractNumber', 'VendorId', 'ContractType', 'Status',
          'StartDate', 'EndDate', 'RenewalDate', 'AutoRenew', 'AutoRenewal', 'NoticePeriodDays',
          'TotalValue', 'AnnualValue', 'Currency', 'Description', 'Notes',
          'ApprovedDate', 'ApprovedById', 'OwnerId', 'RenewalTermMonths', 'PaymentTerms',
          'Created', 'Modified'
        );

      // Apply filters
      if (filter) {
        const filters: string[] = [];

        if (filter.searchTerm) {
          const term = ValidationUtils.sanitizeForOData(filter.searchTerm);
          filters.push(`(substringof('${term}', Title) or substringof('${term}', ContractNumber))`);
        }

        if (filter.status && filter.status.length > 0) {
          const statusFilters = filter.status.map(s =>
            ValidationUtils.buildFilter('Status', 'eq', s)
          );
          filters.push(`(${statusFilters.join(' or ')})`);
        }

        if (filter.type && filter.type.length > 0) {
          const typeFilters = filter.type.map(t =>
            ValidationUtils.buildFilter('ContractType', 'eq', t)
          );
          filters.push(`(${typeFilters.join(' or ')})`);
        }

        if (filter.vendorId !== undefined) {
          const validVendorId = ValidationUtils.validateInteger(filter.vendorId, 'vendorId', 1);
          filters.push(`VendorId eq ${validVendorId}`);
        }

        if (filter.department) {
          const dept = ValidationUtils.sanitizeForOData(filter.department);
          filters.push(ValidationUtils.buildFilter('Department', 'eq', dept));
        }

        if (filter.expiringWithinDays !== undefined) {
          const futureDate = new Date();
          futureDate.setDate(futureDate.getDate() + filter.expiringWithinDays);
          filters.push(`EndDate le datetime'${futureDate.toISOString()}'`);
          filters.push(`Status eq '${ContractStatus.Active}'`);
        }

        if (filter.minValue !== undefined) {
          filters.push(`TotalValue ge ${filter.minValue}`);
        }

        if (filter.maxValue !== undefined) {
          filters.push(`TotalValue le ${filter.maxValue}`);
        }

        if (filters.length > 0) {
          query = query.filter(filters.join(' and '));
        }
      }

      const items = await query.orderBy('EndDate', true).top(5000)();
      console.log(`[ContractService] Retrieved ${items.length} contracts`);
      return items.map(this.mapContractFromSP);
    } catch (error: any) {
      console.error('[ContractService] Error getting contracts:', error?.message || error);
      logger.error('ContractService', 'Error getting contracts:', error);
      return [];
    }
  }

  public async getContractById(id: number): Promise<IContract | null> {
    try {
      console.log('[ContractService] Fetching contract by ID:', id);
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // NOTE: Removed Person field expands (Author, Editor) - cause 400 errors when columns don't exist
      const item = await this.sp.web.lists.getByTitle(this.CONTRACTS_LIST).items
        .getById(validId)
        .select(
          'Id', 'Title', 'ContractNumber', 'VendorId', 'ContractType', 'Status',
          'StartDate', 'EndDate', 'RenewalDate', 'AutoRenew', 'AutoRenewal', 'NoticePeriodDays',
          'TotalValue', 'AnnualValue', 'Currency', 'Description', 'Notes',
          'ApprovedDate', 'ApprovedById', 'OwnerId', 'RenewalTermMonths', 'PaymentTerms',
          'Created', 'Modified'
        )();

      return this.mapContractFromSP(item);
    } catch (error: any) {
      console.error('[ContractService] Error getting contract by ID:', error?.message || error);
      logger.error('ContractService', 'Error getting contract by ID:', error);
      return null;
    }
  }

  public async getContractByNumber(contractNumber: string): Promise<IContract | null> {
    try {
      if (!contractNumber || typeof contractNumber !== 'string') {
        throw new Error('Invalid contract number');
      }

      const validNumber = ValidationUtils.sanitizeForOData(contractNumber.substring(0, 50));
      const filter = ValidationUtils.buildFilter('ContractNumber', 'eq', validNumber);

      const items = await this.sp.web.lists.getByTitle(this.CONTRACTS_LIST).items
        .select('Id', 'ContractNumber')
        .filter(filter)
        .top(1)();

      if (items.length === 0) {
        return null;
      }

      return this.getContractById(items[0].Id);
    } catch (error) {
      logger.error('ContractService', 'Error getting contract by number:', error);
      throw error;
    }
  }

  public async createContract(contract: Partial<IContract>): Promise<number> {
    try {
      // Validate required fields
      if (!contract.Title || !contract.VendorId || !contract.StartDate || !contract.EndDate) {
        throw new Error('Title, VendorId, StartDate, and EndDate are required');
      }

      // Generate contract number
      const contractNumber = contract.ContractNumber || await this.generateContractNumber();

      // Check if contract number already exists
      const existing = await this.getContractByNumber(contractNumber);
      if (existing) {
        throw new Error(`Contract number ${contractNumber} already exists`);
      }

      const itemData: Record<string, unknown> = {
        Title: ValidationUtils.sanitizeHtml(contract.Title),
        ContractNumber: contractNumber,
        VendorId: ValidationUtils.validateInteger(contract.VendorId, 'VendorId', 1),
        ContractType: contract.ContractType || ContractType.ServiceContract,
        Status: contract.Status || ContractStatus.Draft,
        StartDate: ValidationUtils.validateDate(contract.StartDate, 'StartDate'),
        EndDate: ValidationUtils.validateDate(contract.EndDate, 'EndDate'),
        AutoRenew: contract.AutoRenew || false,
        NoticePeriodDays: contract.NotificationDays || 30,
        TotalValue: contract.ContractValue || 0,
        Currency: contract.Currency || Currency.GBP
      };

      // Optional fields
      if (contract.RenewalDate) itemData.RenewalDate = ValidationUtils.validateDate(contract.RenewalDate, 'RenewalDate');
      if (contract.Description) itemData.Description = ValidationUtils.sanitizeHtml(contract.Description);
      if (contract.Notes) itemData.Notes = ValidationUtils.sanitizeHtml(contract.Notes);

      const result = await this.sp.web.lists.getByTitle(this.CONTRACTS_LIST).items.add(itemData);
      return result.data.Id;
    } catch (error) {
      logger.error('ContractService', 'Error creating contract:', error);
      throw error;
    }
  }

  public async updateContract(id: number, updates: Partial<IContract>): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const itemData: Record<string, unknown> = {};

      if (updates.Title) itemData.Title = ValidationUtils.sanitizeHtml(updates.Title);
      if (updates.ContractType) {
        ValidationUtils.validateEnum(updates.ContractType, ContractType, 'ContractType');
        itemData.ContractType = updates.ContractType;
      }
      if (updates.Status) {
        ValidationUtils.validateEnum(updates.Status, ContractStatus, 'Status');
        itemData.Status = updates.Status;
      }
      if (updates.StartDate) itemData.StartDate = ValidationUtils.validateDate(updates.StartDate, 'StartDate');
      if (updates.EndDate) itemData.EndDate = ValidationUtils.validateDate(updates.EndDate, 'EndDate');
      if (updates.RenewalDate) itemData.RenewalDate = ValidationUtils.validateDate(updates.RenewalDate, 'RenewalDate');
      if (updates.AutoRenew !== undefined) itemData.AutoRenew = updates.AutoRenew;
      if (updates.NotificationDays !== undefined) itemData.NoticePeriodDays = updates.NotificationDays;
      if (updates.ContractValue !== undefined) itemData.TotalValue = updates.ContractValue;
      if (updates.Currency) itemData.Currency = updates.Currency;
      if (updates.Description !== undefined) itemData.Description = updates.Description ? ValidationUtils.sanitizeHtml(updates.Description) : null;
      if (updates.Notes !== undefined) itemData.Notes = updates.Notes ? ValidationUtils.sanitizeHtml(updates.Notes) : null;

      await this.sp.web.lists.getByTitle(this.CONTRACTS_LIST).items.getById(validId).update(itemData);
    } catch (error) {
      logger.error('ContractService', 'Error updating contract:', error);
      throw error;
    }
  }

  public async deleteContract(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      // Check if contract is in Draft status
      const contract = await this.getContractById(validId);
      if (contract.Status !== ContractStatus.Draft) {
        throw new Error('Only draft contracts can be deleted');
      }

      await this.sp.web.lists.getByTitle(this.CONTRACTS_LIST).items.getById(validId).delete();
    } catch (error) {
      logger.error('ContractService', 'Error deleting contract:', error);
      throw error;
    }
  }

  // ==================== Contract Workflow ====================

  public async activateContract(id: number): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const contract = await this.getContractById(validId);
      if (contract.Status !== ContractStatus.Draft && contract.Status !== ContractStatus.PendingSignature) {
        throw new Error('Contract must be in Draft or Pending Signature status to activate');
      }

      await this.updateContract(validId, {
        Status: ContractStatus.Active,
        SignedDate: new Date()
      });
    } catch (error) {
      logger.error('ContractService', 'Error activating contract:', error);
      throw error;
    }
  }

  public async terminateContract(id: number, reason?: string): Promise<void> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const contract = await this.getContractById(validId);
      if (contract.Status !== ContractStatus.Active) {
        throw new Error('Only active contracts can be terminated');
      }

      const updateData: Partial<IContract> = {
        Status: ContractStatus.Terminated,
        TerminationDate: new Date()
      };

      if (reason) {
        updateData.Notes = contract.Notes
          ? `${contract.Notes}\n\nTermination Reason: ${reason}`
          : `Termination Reason: ${reason}`;
      }

      await this.updateContract(validId, updateData);
    } catch (error) {
      logger.error('ContractService', 'Error terminating contract:', error);
      throw error;
    }
  }

  public async renewContract(id: number, newEndDate: Date, newValue?: number): Promise<number> {
    try {
      const validId = ValidationUtils.validateInteger(id, 'id', 1);

      const contract = await this.getContractById(validId);

      // Create new contract based on existing
      const newContractId = await this.createContract({
        Title: `${contract.Title} (Renewed)`,
        VendorId: contract.VendorId,
        ContractType: contract.ContractType,
        StartDate: contract.EndDate, // New contract starts when old one ends
        EndDate: newEndDate,
        AutoRenew: contract.AutoRenew,
        NotificationDays: contract.NotificationDays,
        ContractValue: newValue || contract.ContractValue,
        Currency: contract.Currency,
        Description: contract.Description,
        Notes: `Renewed from contract ${contract.ContractNumber}`
      });

      // Update original contract status
      await this.updateContract(validId, {
        Status: ContractStatus.Renewed,
        Notes: contract.Notes
          ? `${contract.Notes}\n\nRenewed to contract ID: ${newContractId}`
          : `Renewed to contract ID: ${newContractId}`
      });

      return newContractId;
    } catch (error) {
      logger.error('ContractService', 'Error renewing contract:', error);
      throw error;
    }
  }

  // ==================== Expiry Management ====================

  public async getExpiringContracts(days: number = 90): Promise<IContract[]> {
    try {
      return this.getContracts({
        status: [ContractStatus.Active],
        expiringWithinDays: days
      });
    } catch (error) {
      logger.error('ContractService', 'Error getting expiring contracts:', error);
      throw error;
    }
  }

  public async getContractsRequiringRenewalNotice(): Promise<IContract[]> {
    try {
      const contracts = await this.getContracts({
        status: [ContractStatus.Active]
      });

      const today = new Date();
      return contracts.filter(contract => {
        if (contract.RenewalDate && contract.NotificationDays) {
          const notificationDate = new Date(contract.RenewalDate);
          notificationDate.setDate(notificationDate.getDate() - contract.NotificationDays);
          return notificationDate <= today && contract.RenewalDate > today;
        }
        return false;
      });
    } catch (error) {
      logger.error('ContractService', 'Error getting contracts requiring renewal notice:', error);
      throw error;
    }
  }

  public async getExpiredContracts(): Promise<IContract[]> {
    try {
      const today = new Date();
      const contracts = await this.getContracts({
        status: [ContractStatus.Active]
      });

      return contracts.filter(contract =>
        contract.EndDate && new Date(contract.EndDate) < today
      );
    } catch (error) {
      logger.error('ContractService', 'Error getting expired contracts:', error);
      throw error;
    }
  }

  // ==================== Vendor Contracts ====================

  public async getVendorContracts(vendorId: number, activeOnly: boolean = false): Promise<IContract[]> {
    try {
      const filter: IContractFilter = { vendorId };
      if (activeOnly) {
        filter.status = [ContractStatus.Active];
      }

      return this.getContracts(filter);
    } catch (error) {
      logger.error('ContractService', 'Error getting vendor contracts:', error);
      throw error;
    }
  }

  // ==================== Statistics ====================

  public async getContractStatistics(): Promise<{
    total: number;
    active: number;
    expiring30Days: number;
    expiring90Days: number;
    expired: number;
    totalValue: number;
    byType: { [key: string]: number };
    byStatus: { [key: string]: number };
  }> {
    try {
      const contracts = await this.getContracts();
      const today = new Date();
      const in30Days = new Date();
      in30Days.setDate(today.getDate() + 30);
      const in90Days = new Date();
      in90Days.setDate(today.getDate() + 90);

      const stats = {
        total: contracts.length,
        active: 0,
        expiring30Days: 0,
        expiring90Days: 0,
        expired: 0,
        totalValue: 0,
        byType: {} as { [key: string]: number },
        byStatus: {} as { [key: string]: number }
      };

      for (const contract of contracts) {
        // Count by status
        stats.byStatus[contract.Status] = (stats.byStatus[contract.Status] || 0) + 1;

        // Count by type
        stats.byType[contract.ContractType] = (stats.byType[contract.ContractType] || 0) + 1;

        // Active contracts
        if (contract.Status === ContractStatus.Active) {
          stats.active++;
          stats.totalValue += contract.ContractValue || 0;

          // Check expiry
          if (contract.EndDate) {
            const endDate = new Date(contract.EndDate);
            if (endDate < today) {
              stats.expired++;
            } else if (endDate <= in30Days) {
              stats.expiring30Days++;
              stats.expiring90Days++;
            } else if (endDate <= in90Days) {
              stats.expiring90Days++;
            }
          }
        }
      }

      return stats;
    } catch (error) {
      logger.error('ContractService', 'Error getting contract statistics:', error);
      throw error;
    }
  }

  // ==================== Helper Functions ====================

  private async generateContractNumber(): Promise<string> {
    try {
      const year = new Date().getFullYear();
      const prefix = `CON-${year}-`;

      const items = await this.sp.web.lists.getByTitle(this.CONTRACTS_LIST).items
        .select('ContractNumber')
        .filter(`substringof('${prefix}', ContractNumber)`)
        .orderBy('Id', false)
        .top(1)();

      let nextNumber = 1;
      if (items.length > 0 && items[0].ContractNumber) {
        const match = items[0].ContractNumber.match(/CON-\d{4}-(\d+)/);
        if (match) {
          nextNumber = parseInt(match[1], 10) + 1;
        }
      }

      return `${prefix}${nextNumber.toString().padStart(3, '0')}`;
    } catch (error) {
      logger.error('ContractService', 'Error generating contract number:', error);
      return `CON-${Date.now()}`;
    }
  }

  // ==================== Mapping Functions ====================

  private mapContractFromSP(item: Record<string, unknown>): IContract {
    // Map SharePoint field names to interface properties
    // Note: SignedDate/SignedById don't exist in SP list - use ApprovedDate instead
    return {
      Id: item.Id as number,
      Title: item.Title as string,
      ContractNumber: item.ContractNumber as string,
      VendorId: item.VendorId as number,
      ContractType: item.ContractType as ContractType,
      Status: item.Status as ContractStatus,
      StartDate: item.StartDate ? new Date(item.StartDate as string) : new Date(),
      EndDate: item.EndDate ? new Date(item.EndDate as string) : new Date(),
      SignedDate: item.ApprovedDate ? new Date(item.ApprovedDate as string) : undefined,  // SP field: ApprovedDate
      RenewalDate: item.RenewalDate ? new Date(item.RenewalDate as string) : undefined,
      TerminationDate: undefined,  // Not in SP list
      AutoRenew: (item.AutoRenew as boolean) || (item.AutoRenewal as boolean) || false,  // SP has both AutoRenew and AutoRenewal
      RenewalTermMonths: item.RenewalTermMonths as number,
      NotificationDays: item.NoticePeriodDays as number || 30,
      ContractValue: item.TotalValue as number || 0,
      AnnualValue: item.AnnualValue as number,
      Currency: item.Currency as Currency || Currency.GBP,
      Description: item.Description as string,
      Notes: item.Notes as string,
      Version: 1,  // Not in SP list
      Created: item.Created ? new Date(item.Created as string) : undefined,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined
    };
  }
}
