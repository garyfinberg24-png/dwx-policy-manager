// @ts-nocheck
/**
 * AssetActionHandler
 * Handles equipment, asset, and license operations within workflow execution
 * - Create equipment requests for new hires
 * - Create asset return requests for leavers
 * - Track license reclamation
 *
 * Uses SharePoint lists:
 * - JML_EquipmentRequests: Equipment provisioning requests
 * - JML_AssetReturns: Asset return tracking for leavers
 * - JML_LicenseReclamation: Software license recovery tracking
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import {
  IActionContext,
  IActionResult,
  IActionConfig
} from '../../../models/IWorkflow';
import { logger } from '../../LoggingService';

// ============================================================================
// INTERFACES
// ============================================================================

/**
 * Equipment request status
 */
export enum EquipmentRequestStatus {
  Pending = 'Pending',
  Approved = 'Approved',
  Ordered = 'Ordered',
  Delivered = 'Delivered',
  Assigned = 'Assigned',
  Rejected = 'Rejected',
  Cancelled = 'Cancelled'
}

/**
 * Asset return status
 */
export enum AssetReturnStatus {
  Pending = 'Pending',
  Scheduled = 'Scheduled',
  Collected = 'Collected',
  Verified = 'Verified',
  Complete = 'Complete',
  Missing = 'Missing',
  Damaged = 'Damaged'
}

/**
 * License reclamation status
 */
export enum LicenseReclamationStatus {
  Pending = 'Pending',
  InProgress = 'InProgress',
  Reclaimed = 'Reclaimed',
  Reassigned = 'Reassigned',
  Failed = 'Failed'
}

/**
 * Equipment request item
 */
export interface IEquipmentRequest {
  Id?: number;
  Title: string;
  EquipmentType: string;
  Description?: string;
  RequestedFor: string;         // Employee name
  RequestedForEmail: string;
  RequestedById?: number;
  ProcessId: number;
  Priority: 'Low' | 'Normal' | 'High' | 'Urgent';
  Status: EquipmentRequestStatus;
  RequestedDate: Date;
  RequiredByDate?: Date;
  ApprovedById?: number;
  ApprovedDate?: Date;
  Notes?: string;
}

/**
 * Asset return item
 */
export interface IAssetReturn {
  Id?: number;
  Title: string;
  AssetTag?: string;
  AssetCategory: string;
  AssetDescription: string;
  EmployeeName: string;
  EmployeeEmail: string;
  ProcessId: number;
  Status: AssetReturnStatus;
  ReturnDeadline: Date;
  ScheduledReturnDate?: Date;
  ActualReturnDate?: Date;
  CollectedById?: number;
  Condition?: 'Good' | 'Fair' | 'Poor' | 'Damaged' | 'Missing';
  Notes?: string;
}

/**
 * License reclamation item
 */
export interface ILicenseReclamation {
  Id?: number;
  Title: string;
  LicenseType: string;
  LicenseId?: string;
  EmployeeName: string;
  EmployeeEmail: string;
  ProcessId: number;
  Status: LicenseReclamationStatus;
  ReclaimDeadline: Date;
  ReclaimedDate?: Date;
  ReassignedTo?: string;
  ReassignedDate?: Date;
  Notes?: string;
}

// ============================================================================
// HANDLER CLASS
// ============================================================================

export class AssetActionHandler {
  private sp: SPFI;
  private siteUrl: string;

  // List names
  private readonly equipmentRequestsListName = 'JML_EquipmentRequests';
  private readonly assetReturnsListName = 'JML_AssetReturns';
  private readonly licenseReclamationListName = 'JML_LicenseReclamation';

  constructor(sp: SPFI, siteUrl: string) {
    this.sp = sp;
    this.siteUrl = siteUrl;
  }

  // ============================================================================
  // EQUIPMENT REQUESTS (JOINERS)
  // ============================================================================

  /**
   * Create equipment request(s) for a new hire
   */
  public async createEquipmentRequest(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const equipmentTypes = this.resolveEquipmentTypes(config, context);

      if (equipmentTypes.length === 0) {
        return { success: false, error: 'No equipment types specified' };
      }

      const employeeName = (context.process['EmployeeName'] || context.process['Title']) as string;
      const employeeEmail = context.process['EmployeeEmail'] as string;
      const processId = context.workflowInstance.ProcessId;
      const startDate = context.process['StartDate'] as Date;

      if (!employeeName || !employeeEmail) {
        return { success: false, error: 'Employee name and email are required' };
      }

      const createdIds: number[] = [];
      const errors: string[] = [];

      for (const equipmentType of equipmentTypes) {
        try {
          const requestData: Partial<IEquipmentRequest> = {
            Title: `${equipmentType} for ${employeeName}`,
            EquipmentType: equipmentType,
            Description: config.equipmentDescription || `Standard ${equipmentType} provisioning for new hire`,
            RequestedFor: employeeName,
            RequestedForEmail: employeeEmail,
            ProcessId: processId,
            Priority: this.mapPriorityFromDays(startDate),
            Status: EquipmentRequestStatus.Pending,
            RequestedDate: new Date(),
            RequiredByDate: startDate
          };

          const result = await this.sp.web.lists
            .getByTitle(this.equipmentRequestsListName)
            .items.add(requestData);

          createdIds.push(result.data.Id);
          logger.info('AssetActionHandler', `Created equipment request ${result.data.Id} for ${equipmentType}`);
        } catch (error) {
          errors.push(`Failed to create request for ${equipmentType}: ${error instanceof Error ? error.message : 'Unknown error'}`);
          logger.error('AssetActionHandler', `Error creating equipment request for ${equipmentType}`, error);
        }
      }

      return {
        success: errors.length === 0,
        error: errors.length > 0 ? errors.join('; ') : undefined,
        createdItemIds: createdIds,
        outputVariables: {
          equipmentRequestIds: createdIds,
          equipmentRequestCount: createdIds.length,
          equipmentTypes
        },
        nextAction: 'continue'
      };
    } catch (error) {
      logger.error('AssetActionHandler', 'Error in createEquipmentRequest', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to create equipment request'
      };
    }
  }

  /**
   * Create standard equipment package for new hire
   */
  public async createStandardEquipmentPackage(
    employeeName: string,
    employeeEmail: string,
    processId: number,
    startDate: Date,
    role?: string
  ): Promise<{ success: boolean; requestIds: number[]; error?: string }> {
    try {
      // Define standard equipment based on role
      const standardEquipment = this.getStandardEquipmentForRole(role);
      const createdIds: number[] = [];

      for (const equipment of standardEquipment) {
        const requestData: Partial<IEquipmentRequest> = {
          Title: `${equipment.type} for ${employeeName}`,
          EquipmentType: equipment.type,
          Description: equipment.description,
          RequestedFor: employeeName,
          RequestedForEmail: employeeEmail,
          ProcessId: processId,
          Priority: this.mapPriorityFromDays(startDate),
          Status: EquipmentRequestStatus.Pending,
          RequestedDate: new Date(),
          RequiredByDate: startDate
        };

        const result = await this.sp.web.lists
          .getByTitle(this.equipmentRequestsListName)
          .items.add(requestData);

        createdIds.push(result.data.Id);
      }

      logger.info('AssetActionHandler', `Created ${createdIds.length} standard equipment requests for ${employeeName}`);

      return { success: true, requestIds: createdIds };
    } catch (error) {
      logger.error('AssetActionHandler', 'Error creating standard equipment package', error);
      return {
        success: false,
        requestIds: [],
        error: error instanceof Error ? error.message : 'Failed to create equipment package'
      };
    }
  }

  // ============================================================================
  // ASSET RETURNS (LEAVERS)
  // ============================================================================

  /**
   * Create asset return request(s) for a leaver
   */
  public async createAssetReturnRequest(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const employeeName = (context.process['EmployeeName'] || context.process['Title']) as string;
      const employeeEmail = context.process['EmployeeEmail'] as string;
      const processId = context.workflowInstance.ProcessId;
      const lastWorkingDay = this.resolveDate(config, context, 'returnDeadline');

      if (!employeeName || !employeeEmail) {
        return { success: false, error: 'Employee name and email are required' };
      }

      // Get employee's assigned assets (from a hypothetical JML_AssignedAssets list)
      const assignedAssets = await this.getEmployeeAssignedAssets(employeeEmail);
      const createdIds: number[] = [];
      const errors: string[] = [];

      for (const asset of assignedAssets) {
        try {
          const returnData: Partial<IAssetReturn> = {
            Title: `Return: ${asset.assetTag || asset.category}`,
            AssetTag: asset.assetTag,
            AssetCategory: asset.category,
            AssetDescription: asset.description,
            EmployeeName: employeeName,
            EmployeeEmail: employeeEmail,
            ProcessId: processId,
            Status: AssetReturnStatus.Pending,
            ReturnDeadline: lastWorkingDay
          };

          const result = await this.sp.web.lists
            .getByTitle(this.assetReturnsListName)
            .items.add(returnData);

          createdIds.push(result.data.Id);
          logger.info('AssetActionHandler', `Created asset return request ${result.data.Id} for ${asset.assetTag || asset.category}`);
        } catch (error) {
          errors.push(`Failed to create return for ${asset.assetTag}: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
      }

      // If no assigned assets found, create a general return reminder
      if (assignedAssets.length === 0) {
        const generalReturnData: Partial<IAssetReturn> = {
          Title: `Asset Return Check: ${employeeName}`,
          AssetCategory: 'General',
          AssetDescription: 'Please verify and return all company assets',
          EmployeeName: employeeName,
          EmployeeEmail: employeeEmail,
          ProcessId: processId,
          Status: AssetReturnStatus.Pending,
          ReturnDeadline: lastWorkingDay,
          Notes: 'No specific assets found in inventory. Please verify with IT if employee has any company equipment.'
        };

        const result = await this.sp.web.lists
          .getByTitle(this.assetReturnsListName)
          .items.add(generalReturnData);

        createdIds.push(result.data.Id);
      }

      return {
        success: errors.length === 0,
        error: errors.length > 0 ? errors.join('; ') : undefined,
        createdItemIds: createdIds,
        outputVariables: {
          assetReturnIds: createdIds,
          assetReturnCount: createdIds.length,
          lastWorkingDay: lastWorkingDay.toISOString()
        },
        nextAction: 'continue'
      };
    } catch (error) {
      logger.error('AssetActionHandler', 'Error in createAssetReturnRequest', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to create asset return request'
      };
    }
  }

  // ============================================================================
  // LICENSE RECLAMATION (LEAVERS)
  // ============================================================================

  /**
   * Create license reclamation records for a leaver
   */
  public async reclaimLicenses(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const licenseTypes = this.resolveLicenseTypes(config, context);
      const employeeName = (context.process['EmployeeName'] || context.process['Title']) as string;
      const employeeEmail = context.process['EmployeeEmail'] as string;
      const processId = context.workflowInstance.ProcessId;
      const lastWorkingDay = this.resolveDate(config, context, 'returnDeadline');

      if (!employeeName || !employeeEmail) {
        return { success: false, error: 'Employee name and email are required' };
      }

      // If no specific licenses provided, try to get from employee's license assignments
      let licensesToReclaim = licenseTypes;
      if (licensesToReclaim.length === 0) {
        licensesToReclaim = await this.getEmployeeLicenses(employeeEmail);
      }

      const createdIds: number[] = [];
      const errors: string[] = [];

      for (const licenseType of licensesToReclaim) {
        try {
          const reclamationData: Partial<ILicenseReclamation> = {
            Title: `Reclaim ${licenseType} from ${employeeName}`,
            LicenseType: licenseType,
            EmployeeName: employeeName,
            EmployeeEmail: employeeEmail,
            ProcessId: processId,
            Status: LicenseReclamationStatus.Pending,
            ReclaimDeadline: lastWorkingDay
          };

          const result = await this.sp.web.lists
            .getByTitle(this.licenseReclamationListName)
            .items.add(reclamationData);

          createdIds.push(result.data.Id);
          logger.info('AssetActionHandler', `Created license reclamation ${result.data.Id} for ${licenseType}`);
        } catch (error) {
          errors.push(`Failed to create reclamation for ${licenseType}: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
      }

      return {
        success: errors.length === 0,
        error: errors.length > 0 ? errors.join('; ') : undefined,
        createdItemIds: createdIds,
        outputVariables: {
          licenseReclamationIds: createdIds,
          licenseReclamationCount: createdIds.length,
          licenseTypes: licensesToReclaim
        },
        nextAction: 'continue'
      };
    } catch (error) {
      logger.error('AssetActionHandler', 'Error in reclaimLicenses', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to create license reclamation records'
      };
    }
  }

  // ============================================================================
  // STATUS TRACKING
  // ============================================================================

  /**
   * Get equipment request status for a process
   */
  public async getEquipmentRequestStatus(processId: number): Promise<{
    total: number;
    pending: number;
    completed: number;
    items: Array<{ id: number; type: string; status: string }>;
  }> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.equipmentRequestsListName)
        .items
        .filter(`ProcessId eq ${processId}`)
        .select('Id', 'EquipmentType', 'Status')();

      return {
        total: items.length,
        pending: items.filter((i: { Status: string }) =>
          i.Status === EquipmentRequestStatus.Pending ||
          i.Status === EquipmentRequestStatus.Approved ||
          i.Status === EquipmentRequestStatus.Ordered
        ).length,
        completed: items.filter((i: { Status: string }) =>
          i.Status === EquipmentRequestStatus.Assigned ||
          i.Status === EquipmentRequestStatus.Delivered
        ).length,
        items: items.map((i: { Id: number; EquipmentType: string; Status: string }) => ({
          id: i.Id,
          type: i.EquipmentType,
          status: i.Status
        }))
      };
    } catch (error) {
      logger.error('AssetActionHandler', 'Error getting equipment request status', error);
      return { total: 0, pending: 0, completed: 0, items: [] };
    }
  }

  /**
   * Get asset return status for a process
   */
  public async getAssetReturnStatus(processId: number): Promise<{
    total: number;
    pending: number;
    completed: number;
    items: Array<{ id: number; assetTag: string; status: string }>;
  }> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.assetReturnsListName)
        .items
        .filter(`ProcessId eq ${processId}`)
        .select('Id', 'AssetTag', 'Status')();

      return {
        total: items.length,
        pending: items.filter((i: { Status: string }) =>
          i.Status === AssetReturnStatus.Pending ||
          i.Status === AssetReturnStatus.Scheduled
        ).length,
        completed: items.filter((i: { Status: string }) =>
          i.Status === AssetReturnStatus.Complete ||
          i.Status === AssetReturnStatus.Verified
        ).length,
        items: items.map((i: { Id: number; AssetTag: string; Status: string }) => ({
          id: i.Id,
          assetTag: i.AssetTag || 'N/A',
          status: i.Status
        }))
      };
    } catch (error) {
      logger.error('AssetActionHandler', 'Error getting asset return status', error);
      return { total: 0, pending: 0, completed: 0, items: [] };
    }
  }

  // ============================================================================
  // PRIVATE HELPER METHODS
  // ============================================================================

  private resolveEquipmentTypes(config: IActionConfig, context: IActionContext): string[] {
    const types: string[] = [];

    if (config.equipmentType) {
      types.push(config.equipmentType);
    }

    if (config.equipmentTypes && config.equipmentTypes.length > 0) {
      types.push(...config.equipmentTypes);
    }

    // Check process fields for equipment requirements
    const roleEquipment = context.process['RequiredEquipment'];
    if (Array.isArray(roleEquipment)) {
      types.push(...roleEquipment.filter((e): e is string => typeof e === 'string'));
    }

    return Array.from(new Set(types)); // Remove duplicates
  }

  private resolveLicenseTypes(config: IActionConfig, context: IActionContext): string[] {
    const types: string[] = [];

    if (config.licenseType) {
      types.push(config.licenseType);
    }

    if (config.licenseTypes && config.licenseTypes.length > 0) {
      types.push(...config.licenseTypes);
    }

    return Array.from(new Set(types));
  }

  private resolveDate(config: IActionConfig, context: IActionContext, type: 'returnDeadline'): Date {
    const dateField = type === 'returnDeadline' ? config.returnDeadlineField : undefined;
    const dateValue = type === 'returnDeadline' ? config.returnDeadline : undefined;

    // From field
    if (dateField) {
      const fieldValue = context.process[dateField];
      if (fieldValue instanceof Date) {
        return fieldValue;
      }
      if (typeof fieldValue === 'string') {
        return new Date(fieldValue);
      }
    }

    // Static value
    if (dateValue) {
      return new Date(dateValue);
    }

    // Try common date fields
    const lastWorkingDay = context.process['LastWorkingDay'] || context.process['EndDate'];
    if (lastWorkingDay instanceof Date) {
      return lastWorkingDay;
    }
    if (typeof lastWorkingDay === 'string') {
      return new Date(lastWorkingDay);
    }

    // Default to 2 weeks from now
    const defaultDate = new Date();
    defaultDate.setDate(defaultDate.getDate() + 14);
    return defaultDate;
  }

  private mapPriorityFromDays(targetDate: Date): 'Low' | 'Normal' | 'High' | 'Urgent' {
    const daysUntil = Math.floor((targetDate.getTime() - Date.now()) / (1000 * 60 * 60 * 24));

    if (daysUntil <= 3) return 'Urgent';
    if (daysUntil <= 7) return 'High';
    if (daysUntil <= 14) return 'Normal';
    return 'Low';
  }

  private getStandardEquipmentForRole(role?: string): Array<{ type: string; description: string }> {
    // Standard equipment for all roles
    const standard = [
      { type: 'Laptop', description: 'Standard business laptop' },
      { type: 'Monitor', description: 'External display monitor' },
      { type: 'Keyboard & Mouse', description: 'Wireless keyboard and mouse set' },
      { type: 'Headset', description: 'USB/Bluetooth headset for calls' },
      { type: 'Security Badge', description: 'Building access badge' }
    ];

    // Role-specific additions
    const roleSpecific: Record<string, Array<{ type: string; description: string }>> = {
      'Developer': [
        { type: 'Second Monitor', description: 'Additional display for development' },
        { type: 'Docking Station', description: 'USB-C docking station' }
      ],
      'Designer': [
        { type: 'Graphics Tablet', description: 'Drawing/design tablet' },
        { type: '4K Monitor', description: 'High-resolution display for design work' }
      ],
      'Executive': [
        { type: 'Mobile Phone', description: 'Company mobile device' },
        { type: 'Docking Station', description: 'Premium docking station' }
      ],
      'Field Worker': [
        { type: 'Mobile Phone', description: 'Company mobile device' },
        { type: 'Tablet', description: 'Field tablet device' }
      ]
    };

    const equipment = [...standard];

    if (role && roleSpecific[role]) {
      equipment.push(...roleSpecific[role]);
    }

    return equipment;
  }

  private async getEmployeeAssignedAssets(employeeEmail: string): Promise<Array<{
    assetTag: string;
    category: string;
    description: string;
  }>> {
    try {
      // Try to get from JML_AssignedAssets list
      const assets = await this.sp.web.lists
        .getByTitle('JML_AssignedAssets')
        .items
        .filter(`AssignedToEmail eq '${employeeEmail}'`)
        .select('AssetTag', 'Category', 'Description')();

      return assets.map((a: { AssetTag: string; Category: string; Description: string }) => ({
        assetTag: a.AssetTag,
        category: a.Category,
        description: a.Description
      }));
    } catch {
      // List might not exist or no assets found
      logger.info('AssetActionHandler', `No assigned assets found for ${employeeEmail}`);
      return [];
    }
  }

  private async getEmployeeLicenses(employeeEmail: string): Promise<string[]> {
    try {
      // Try to get from JML_LicenseAssignments list
      const licenses = await this.sp.web.lists
        .getByTitle('JML_LicenseAssignments')
        .items
        .filter(`AssignedToEmail eq '${employeeEmail}'`)
        .select('LicenseType')();

      return licenses.map((l: { LicenseType: string }) => l.LicenseType);
    } catch {
      // List might not exist or no licenses found
      // Return common license types as defaults
      return ['Microsoft 365', 'Teams', 'SharePoint'];
    }
  }
}

export default AssetActionHandler;
