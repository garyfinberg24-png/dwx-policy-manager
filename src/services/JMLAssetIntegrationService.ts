// @ts-nocheck
// JML Asset Integration Service
// Integrates asset tracking with JML onboarding/offboarding processes

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import {
  IAsset,
  IAssetType,
  AssetCategory,
  AssetStatus,
  AssetCondition,
  IJMLAssetAssignment
} from '../models/IAsset';
import { IJmlProcess, ProcessType } from '../models/IJmlProcess';
import { AssetService } from './AssetService';
import { AssetTrackingService } from './AssetTrackingService';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export interface IJMLAssetConfig {
  // Asset types to assign for different roles/departments
  assetTypesByRole: Map<string, number[]>; // JobTitle -> AssetTypeIds
  assetTypesByDepartment: Map<string, number[]>; // Department -> AssetTypeIds
  defaultAssetTypes: number[]; // Default assets for all new employees

  // Hardware assignments
  laptopRequired: boolean;
  monitorRequired: boolean;
  phoneRequired: boolean;
  dockingStationRequired: boolean;
  headsetRequired: boolean;

  // Software licenses
  softwareLicenses: number[]; // AssetTypeIds for software
}

export interface IJMLAssetAssignmentResult {
  success: boolean;
  processId: number;
  employeeId: number;
  assignedAssets: IAsset[];
  failedAssets: {
    assetTypeId: number;
    reason: string;
  }[];
  notes: string[];
}

export interface IJMLAssetRetrievalResult {
  success: boolean;
  processId: number;
  employeeId: number;
  retrievedAssets: IAsset[];
  missingAssets: IAsset[];
  notes: string[];
}

export class JMLAssetIntegrationService {
  private sp: SPFI;
  private assetService: AssetService;
  private trackingService: AssetTrackingService;
  private readonly PM_PROCESSES_LIST = 'PM_Processes';
  private readonly PM_ASSET_CONFIG_LIST = 'PM_Asset_Configuration';

  constructor(sp: SPFI) {
    this.sp = sp;
    this.assetService = new AssetService(sp);
    this.trackingService = new AssetTrackingService(sp);
  }

  // ==================== Onboarding Integration ====================

  /**
   * Automatically assign assets when a new joiner process is created
   */
  public async assignAssetsForOnboarding(
    process: IJmlProcess,
    assignedById: number,
    customAssetTypeIds?: number[]
  ): Promise<IJMLAssetAssignmentResult> {
    try {
      const validProcessId = ValidationUtils.validateInteger(process.Id!, 'processId', 1);
      const validAssignedById = ValidationUtils.validateInteger(assignedById, 'assignedById', 1);

      if (process.ProcessType !== ProcessType.Joiner) {
        throw new Error('This method is only for joiner processes');
      }

      const result: IJMLAssetAssignmentResult = {
        success: false,
        processId: validProcessId,
        employeeId: 0, // Will be set after employee user is created
        assignedAssets: [],
        failedAssets: [],
        notes: []
      };

      // Get asset configuration
      const config = await this.getAssetConfigForEmployee(
        process.JobTitle,
        process.Department,
        process.Location
      );

      // Determine which asset types to assign
      let assetTypesToAssign: number[] = customAssetTypeIds || config.defaultAssetTypes;

      // Add role-specific assets
      if (config.assetTypesByRole.has(process.JobTitle)) {
        assetTypesToAssign = [
          ...assetTypesToAssign,
          ...config.assetTypesByRole.get(process.JobTitle)!
        ];
      }

      // Add department-specific assets
      if (config.assetTypesByDepartment.has(process.Department)) {
        assetTypesToAssign = [
          ...assetTypesToAssign,
          ...config.assetTypesByDepartment.get(process.Department)!
        ];
      }

      // Remove duplicates - use Array.from to avoid downlevelIteration issues
      assetTypesToAssign = Array.from(new Set(assetTypesToAssign));

      result.notes.push(`Assigning ${assetTypesToAssign.length} asset types`);

      // Assign each asset type
      for (const assetTypeId of assetTypesToAssign) {
        try {
          const validAssetTypeId = ValidationUtils.validateInteger(assetTypeId, 'assetTypeId', 1);

          // Find available asset of this type
          const availableAssets = await this.assetService.getAssets({
            assetTypeId: validAssetTypeId,
            status: [AssetStatus.Available],
            condition: [AssetCondition.New, AssetCondition.Excellent, AssetCondition.Good]
          });

          if (availableAssets.length > 0) {
            // Reserve the asset (assign to a placeholder until employee user is created)
            const asset = availableAssets[0];

            await this.assetService.updateAsset(asset.Id!, {
              Status: AssetStatus.Reserved,
              Notes: `Reserved for ${process.EmployeeName} (Process #${validProcessId})`
            });

            result.assignedAssets.push(asset);
            result.notes.push(`Reserved asset ${asset.AssetTag} for employee`);

            // Create asset assignment record linked to process
            const assignmentId = await this.sp.web.lists.getByTitle('Asset Assignments').items.add({
              AssetId: asset.Id,
              ProcessId: validProcessId,
              AssignmentReason: `Onboarding for ${process.EmployeeName}`,
              AssignedById: validAssignedById,
              AssignedDate: new Date().toISOString(),
              Status: 'Reserved',
              IsActive: false // Will be activated when employee user is created
            });

            result.notes.push(`Created assignment record #${assignmentId.data.Id}`);
          } else {
            result.failedAssets.push({
              assetTypeId: validAssetTypeId,
              reason: 'No available assets of this type'
            });
            result.notes.push(`WARNING: No available assets for type ID ${validAssetTypeId}`);
          }
        } catch (error: any) {
          result.failedAssets.push({
            assetTypeId: assetTypeId,
            reason: error.message
          });
          result.notes.push(`ERROR: Failed to assign asset type ${assetTypeId}: ${error.message}`);
        }
      }

      result.success = result.assignedAssets.length > 0;

      // Update process with asset notes
      await this.updateProcessWithAssetInfo(validProcessId, result);

      return result;
    } catch (error) {
      logger.error('JMLAssetIntegrationService', 'Error assigning assets for onboarding:', error);
      throw error;
    }
  }

  /**
   * Activate reserved assets when employee user account is created
   */
  public async activateReservedAssets(
    processId: number,
    employeeUserId: number,
    assignedById: number
  ): Promise<number> {
    try {
      const validProcessId = ValidationUtils.validateInteger(processId, 'processId', 1);
      const validEmployeeUserId = ValidationUtils.validateInteger(employeeUserId, 'employeeUserId', 1);
      const validAssignedById = ValidationUtils.validateInteger(assignedById, 'assignedById', 1);

      // Get all reserved assets for this process
      const filter = `${ValidationUtils.buildFilter('ProcessId', 'eq', validProcessId)} and Status eq 'Reserved'`;

      const assignments = await this.sp.web.lists.getByTitle('Asset Assignments').items
        .filter(filter)
        .select('Id', 'AssetId')();

      let activatedCount = 0;

      for (const assignment of assignments) {
        // Assign asset to employee
        await this.assetService.assignAsset(
          assignment.AssetId,
          validEmployeeUserId,
          validAssignedById,
          `Onboarding assignment for Process #${validProcessId}`
        );

        // Update assignment record
        await this.sp.web.lists.getByTitle('Asset Assignments').items
          .getById(assignment.Id)
          .update({
            AssignedToId: validEmployeeUserId,
            IsActive: true,
            Status: 'Active'
          });

        activatedCount++;
      }

      return activatedCount;
    } catch (error) {
      logger.error('JMLAssetIntegrationService', 'Error activating reserved assets:', error);
      throw error;
    }
  }

  // ==================== Offboarding Integration ====================

  /**
   * Retrieve all assets from employee during offboarding
   */
  public async retrieveAssetsForOffboarding(
    process: IJmlProcess,
    retrievedById: number,
    assetCondition?: AssetCondition,
    notes?: string
  ): Promise<IJMLAssetRetrievalResult> {
    try {
      const validProcessId = ValidationUtils.validateInteger(process.Id!, 'processId', 1);
      const validRetrievedById = ValidationUtils.validateInteger(retrievedById, 'retrievedById', 1);

      if (process.ProcessType !== ProcessType.Leaver) {
        throw new Error('This method is only for leaver processes');
      }

      const result: IJMLAssetRetrievalResult = {
        success: false,
        processId: validProcessId,
        employeeId: 0, // Employee user ID if available
        retrievedAssets: [],
        missingAssets: [],
        notes: []
      };

      // Get all assets assigned to this employee (if we have their user ID)
      // For now, we'll search by employee name if user ID is not available
      const employeeAssets = await this.findAssetsByEmployeeName(process.EmployeeName);

      result.notes.push(`Found ${employeeAssets.length} assets assigned to ${process.EmployeeName}`);

      const returnCondition = assetCondition || AssetCondition.Good;

      for (const asset of employeeAssets) {
        try {
          // Retrieve the asset
          await this.assetService.unassignAsset(
            asset.Id!,
            returnCondition,
            notes || `Retrieved during offboarding (Process #${validProcessId})`
          );

          result.retrievedAssets.push(asset);
          result.notes.push(`Retrieved asset ${asset.AssetTag}`);

          // Link retrieval to process
          await this.sp.web.lists.getByTitle('Asset Assignments').items.add({
            AssetId: asset.Id,
            ProcessId: validProcessId,
            AssignmentReason: `Offboarding retrieval for ${process.EmployeeName}`,
            AssignedById: validRetrievedById,
            ReturnDate: new Date().toISOString(),
            ReturnCondition: returnCondition,
            ReturnNotes: notes,
            Status: 'Returned'
          });
        } catch (error: any) {
          result.missingAssets.push(asset);
          result.notes.push(`ERROR: Failed to retrieve asset ${asset.AssetTag}: ${error.message}`);
        }
      }

      result.success = result.retrievedAssets.length === employeeAssets.length;

      // Update process with retrieval info
      await this.updateProcessWithAssetInfo(validProcessId, result);

      return result;
    } catch (error) {
      logger.error('JMLAssetIntegrationService', 'Error retrieving assets for offboarding:', error);
      throw error;
    }
  }

  // ==================== Mover Integration ====================

  /**
   * Transfer assets when employee moves to new role/location
   */
  public async handleAssetTransferForMover(
    process: IJmlProcess,
    requestedById: number,
    newLocation?: string,
    newDepartment?: string
  ): Promise<number> {
    try {
      const validProcessId = ValidationUtils.validateInteger(process.Id!, 'processId', 1);
      const validRequestedById = ValidationUtils.validateInteger(requestedById, 'requestedById', 1);

      if (process.ProcessType !== ProcessType.Mover) {
        throw new Error('This method is only for mover processes');
      }

      const employeeAssets = await this.findAssetsByEmployeeName(process.EmployeeName);
      let transferCount = 0;

      for (const asset of employeeAssets) {
        // Request transfer for each asset
        await this.trackingService.requestTransfer(
          asset.Id!,
          validRequestedById,
          asset.AssignedToId, // Keep same user
          newLocation || process.Location,
          newDepartment || process.Department,
          `Internal move (Process #${validProcessId})`
        );

        transferCount++;
      }

      return transferCount;
    } catch (error) {
      logger.error('JMLAssetIntegrationService', 'Error handling asset transfer for mover:', error);
      throw error;
    }
  }

  // ==================== Asset Configuration ====================

  /**
   * Get asset configuration for specific employee profile
   */
  private async getAssetConfigForEmployee(
    jobTitle: string,
    department: string,
    location: string
  ): Promise<IJMLAssetConfig> {
    try {
      // For now, return a default configuration
      // In production, this would query the PM_Asset_Configuration list
      const config: IJMLAssetConfig = {
        assetTypesByRole: new Map(),
        assetTypesByDepartment: new Map(),
        defaultAssetTypes: [],
        laptopRequired: true,
        monitorRequired: true,
        phoneRequired: false,
        dockingStationRequired: true,
        headsetRequired: true,
        softwareLicenses: []
      };

      // Example role-based assignments (would come from SharePoint in production)
      if (jobTitle.toLowerCase().includes('developer') || jobTitle.toLowerCase().includes('engineer')) {
        config.assetTypesByRole.set(jobTitle, [1, 2, 3]); // Laptop, Monitor, Docking Station IDs
      }

      if (jobTitle.toLowerCase().includes('manager')) {
        config.assetTypesByRole.set(jobTitle, [1, 2, 4]); // Laptop, Monitor, Phone IDs
      }

      // Example department-based assignments
      if (department.toLowerCase() === 'it' || department.toLowerCase() === 'engineering') {
        config.assetTypesByDepartment.set(department, [1, 2, 3, 5]); // High-spec laptop, dual monitors, etc.
      }

      return config;
    } catch (error) {
      logger.error('JMLAssetIntegrationService', 'Error getting asset configuration:', error);
      throw error;
    }
  }

  /**
   * Find assets by employee name (fallback when user ID is not available)
   */
  private async findAssetsByEmployeeName(employeeName: string): Promise<IAsset[]> {
    try {
      if (!employeeName || typeof employeeName !== 'string') {
        return [];
      }

      const sanitizedName = ValidationUtils.sanitizeForOData(employeeName.substring(0, 100));

      // Get all assigned assets and filter by assignee name
      const assets = await this.assetService.getAssets({
        status: [AssetStatus.Assigned]
      });

      // Filter by employee name (case-insensitive)
      return assets.filter(asset =>
        asset.AssignedTo?.Title?.toLowerCase().includes(sanitizedName.toLowerCase())
      );
    } catch (error) {
      logger.error('JMLAssetIntegrationService', 'Error finding assets by employee name:', error);
      return [];
    }
  }

  /**
   * Update JML process with asset assignment/retrieval information
   */
  private async updateProcessWithAssetInfo(
    processId: number,
    result: IJMLAssetAssignmentResult | IJMLAssetRetrievalResult
  ): Promise<void> {
    try {
      const validProcessId = ValidationUtils.validateInteger(processId, 'processId', 1);

      // Get current process
      const process = await this.sp.web.lists.getByTitle(this.PM_PROCESSES_LIST).items
        .getById(validProcessId)
        .select('Id', 'Comments', 'CustomFields')();

      // Append asset information to notes
      const assetNotes = result.notes.join('\n');
      const updatedNotes = process.Notes
        ? `${process.Notes}\n\n[Asset Update ${new Date().toLocaleString()}]\n${assetNotes}`
        : assetNotes;

      // Update custom fields with asset data
      const customFields = process.CustomFields ? JSON.parse(process.CustomFields) : {};
      customFields.assetInfo = {
        lastUpdate: new Date().toISOString(),
        assignedCount: 'assignedAssets' in result ? result.assignedAssets.length : 0,
        retrievedCount: 'retrievedAssets' in result ? result.retrievedAssets.length : 0,
        notes: result.notes
      };

      await this.sp.web.lists.getByTitle(this.PM_PROCESSES_LIST).items
        .getById(validProcessId)
        .update({
          Notes: ValidationUtils.sanitizeHtml(updatedNotes),
          CustomFields: JSON.stringify(customFields)
        });
    } catch (error) {
      logger.error('JMLAssetIntegrationService', 'Error updating process with asset info:', error);
      // Don't throw - this is a non-critical operation
    }
  }

  // ==================== Reporting ====================

  /**
   * Get asset summary for a JML process
   */
  public async getAssetSummaryForProcess(processId: number): Promise<{
    reservedAssets: IAsset[];
    assignedAssets: IAsset[];
    retrievedAssets: IAsset[];
  }> {
    try {
      const validProcessId = ValidationUtils.validateInteger(processId, 'processId', 1);

      const filter = ValidationUtils.buildFilter('ProcessId', 'eq', validProcessId);

      const assignments = await this.sp.web.lists.getByTitle('Asset Assignments').items
        .filter(filter)
        .select('Id', 'AssetId', 'Status')();

      const assetIds = assignments.map(a => a.AssetId);
      const assets = await Promise.all(
        assetIds.map(id => this.assetService.getAssetById(id))
      );

      return {
        reservedAssets: assets.filter((_, i) => assignments[i].Status === 'Reserved'),
        assignedAssets: assets.filter((_, i) => assignments[i].Status === 'Active'),
        retrievedAssets: assets.filter((_, i) => assignments[i].Status === 'Returned')
      };
    } catch (error) {
      logger.error('JMLAssetIntegrationService', 'Error getting asset summary for process:', error);
      throw error;
    }
  }
}
