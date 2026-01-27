// @ts-nocheck
// Asset Tracking Service
// Advanced asset tracking, checkout, maintenance, and lifecycle management

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import {
  IAsset,
  IAssetCheckout,
  IAssetMaintenance,
  IAssetTransfer,
  IAssetAudit,
  IAssetAuditItem,
  IAssetRequest,
  IAssetHistoryEntry,
  IJMLAssetAssignment,
  IAssetBulkOperation,
  AssetStatus,
  AssetCondition,
  CheckoutStatus,
  MaintenanceType
} from '../models/IAsset';
import { AssetService } from './AssetService';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class AssetTrackingService {
  private sp: SPFI;
  private assetService: AssetService;
  private readonly ASSET_CHECKOUTS_LIST = 'Asset Checkouts';
  private readonly ASSET_MAINTENANCE_LIST = 'Asset Maintenance';
  private readonly ASSET_TRANSFERS_LIST = 'Asset Transfers';
  private readonly ASSET_AUDITS_LIST = 'Asset Audits';
  private readonly ASSET_AUDIT_ITEMS_LIST = 'Asset Audit Items';
  private readonly ASSET_REQUESTS_LIST = 'Asset Requests';

  constructor(sp: SPFI) {
    this.sp = sp;
    this.assetService = new AssetService(sp);
  }

  // ==================== Checkout/Check-in Operations ====================

  public async checkoutAsset(
    assetId: number,
    userId: number,
    checkedOutById: number,
    expectedReturnDate: Date,
    purpose?: string,
    location?: string
  ): Promise<number> {
    try {
      const validAssetId = ValidationUtils.validateInteger(assetId, 'assetId', 1);
      const validUserId = ValidationUtils.validateInteger(userId, 'userId', 1);
      const validCheckedOutById = ValidationUtils.validateInteger(checkedOutById, 'checkedOutById', 1);
      ValidationUtils.validateDate(expectedReturnDate, 'expectedReturnDate');

      // Check if asset is available
      const asset = await this.assetService.getAssetById(validAssetId);
      if (asset.Status !== AssetStatus.Available) {
        throw new Error(`Asset ${asset.AssetTag} is not available for checkout`);
      }

      // Create checkout record
      const checkoutData: any = {
        AssetId: validAssetId,
        CheckedOutToId: validUserId,
        CheckedOutById: validCheckedOutById,
        CheckoutDate: new Date().toISOString(),
        ExpectedReturnDate: expectedReturnDate.toISOString(),
        Status: CheckoutStatus.CheckedOut,
        IsOverdue: false,
        ReminderSent: false,
        OverdueNotificationSent: false
      };

      if (purpose) {
        checkoutData.Purpose = ValidationUtils.sanitizeHtml(purpose);
      }

      if (location) {
        checkoutData.Location = ValidationUtils.sanitizeHtml(location);
      }

      const result = await this.sp.web.lists.getByTitle(this.ASSET_CHECKOUTS_LIST).items.add(checkoutData);

      // Update asset status
      await this.assetService.updateAsset(validAssetId, {
        Status: AssetStatus.Assigned,
        AssignedToId: validUserId,
        AssignedDate: new Date()
      });

      return result.data.Id;
    } catch (error) {
      logger.error('AssetTrackingService', 'Error checking out asset:', error);
      throw error;
    }
  }

  public async checkinAsset(
    checkoutId: number,
    checkedInById: number,
    returnCondition: AssetCondition,
    returnNotes?: string
  ): Promise<void> {
    try {
      const validCheckoutId = ValidationUtils.validateInteger(checkoutId, 'checkoutId', 1);
      const validCheckedInById = ValidationUtils.validateInteger(checkedInById, 'checkedInById', 1);
      ValidationUtils.validateEnum(returnCondition, AssetCondition, 'returnCondition');

      // Get checkout record
      const checkout = await this.sp.web.lists.getByTitle(this.ASSET_CHECKOUTS_LIST).items
        .getById(validCheckoutId)
        .select('Id', 'AssetId', 'Status')();

      if (checkout.Status === CheckoutStatus.CheckedIn) {
        throw new Error('Asset has already been checked in');
      }

      // Update checkout record
      const updateData: any = {
        CheckedInDate: new Date().toISOString(),
        CheckedInById: validCheckedInById,
        ActualReturnDate: new Date().toISOString(),
        Status: CheckoutStatus.CheckedIn,
        ReturnCondition: returnCondition
      };

      if (returnNotes) {
        updateData.ReturnNotes = ValidationUtils.sanitizeHtml(returnNotes);
      }

      await this.sp.web.lists.getByTitle(this.ASSET_CHECKOUTS_LIST).items.getById(validCheckoutId).update(updateData);

      // Update asset status
      await this.assetService.updateAsset(checkout.AssetId, {
        Status: returnCondition === AssetCondition.Broken || returnCondition === AssetCondition.Poor
          ? AssetStatus.InMaintenance
          : AssetStatus.Available,
        Condition: returnCondition,
        AssignedToId: undefined,
        AssignedDate: undefined
      });
    } catch (error) {
      logger.error('AssetTrackingService', 'Error checking in asset:', error);
      throw error;
    }
  }

  public async getCheckouts(filter?: {
    assetId?: number;
    userId?: number;
    status?: CheckoutStatus;
    overdueOnly?: boolean;
  }): Promise<IAssetCheckout[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.ASSET_CHECKOUTS_LIST).items
        .select(
          'Id', 'AssetId', 'CheckedOutToId', 'CheckedOutTo/Title', 'CheckedOutTo/EMail',
          'CheckedOutById', 'CheckedOutBy/Title', 'CheckoutDate', 'ExpectedReturnDate',
          'ActualReturnDate', 'Purpose', 'Location', 'Status', 'IsOverdue',
          'CheckedInDate', 'CheckedInById', 'CheckedInBy/Title', 'ReturnCondition',
          'ReturnNotes', 'ReminderSent', 'OverdueNotificationSent', 'Comments',
          'Created', 'Modified'
        )
        .expand('CheckedOutTo', 'CheckedOutBy', 'CheckedInBy')
        .orderBy('CheckoutDate', false);

      const filters: string[] = [];

      if (filter) {
        if (filter.assetId !== undefined) {
          const validAssetId = ValidationUtils.validateInteger(filter.assetId, 'assetId', 1);
          filters.push(ValidationUtils.buildFilter('AssetId', 'eq', validAssetId));
        }

        if (filter.userId !== undefined) {
          const validUserId = ValidationUtils.validateInteger(filter.userId, 'userId', 1);
          filters.push(ValidationUtils.buildFilter('CheckedOutToId', 'eq', validUserId));
        }

        if (filter.status) {
          ValidationUtils.validateEnum(filter.status, CheckoutStatus, 'status');
          filters.push(ValidationUtils.buildFilter('Status', 'eq', filter.status));
        }

        if (filter.overdueOnly) {
          filters.push('IsOverdue eq 1');
        }
      }

      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      const items = await query.top(1000)();
      return items.map(this.mapCheckoutFromSP);
    } catch (error) {
      logger.error('AssetTrackingService', 'Error getting checkouts:', error);
      throw error;
    }
  }

  public async updateOverdueCheckouts(): Promise<number> {
    try {
      const now = new Date();
      const filter = `Status eq '${CheckoutStatus.CheckedOut}' and IsOverdue eq 0`;

      const checkouts = await this.sp.web.lists.getByTitle(this.ASSET_CHECKOUTS_LIST).items
        .filter(filter)
        .select('Id', 'ExpectedReturnDate')
        .top(1000)();

      let updatedCount = 0;

      for (const checkout of checkouts) {
        if (new Date(checkout.ExpectedReturnDate) < now) {
          await this.sp.web.lists.getByTitle(this.ASSET_CHECKOUTS_LIST).items
            .getById(checkout.Id)
            .update({ IsOverdue: true });
          updatedCount++;
        }
      }

      return updatedCount;
    } catch (error) {
      logger.error('AssetTrackingService', 'Error updating overdue checkouts:', error);
      throw error;
    }
  }

  // ==================== Maintenance Operations ====================

  public async scheduleMaintenance(
    assetId: number,
    maintenanceType: MaintenanceType,
    scheduledDate: Date,
    description: string,
    vendor?: string,
    estimatedCost?: number
  ): Promise<number> {
    try {
      const validAssetId = ValidationUtils.validateInteger(assetId, 'assetId', 1);
      ValidationUtils.validateEnum(maintenanceType, MaintenanceType, 'maintenanceType');
      ValidationUtils.validateDate(scheduledDate, 'scheduledDate');

      const maintenanceData: any = {
        AssetId: validAssetId,
        MaintenanceType: maintenanceType,
        ScheduledDate: scheduledDate.toISOString(),
        Description: ValidationUtils.sanitizeHtml(description),
        IsCompleted: false,
        IsCancelled: false
      };

      if (vendor) {
        maintenanceData.Vendor = ValidationUtils.sanitizeHtml(vendor);
      }

      if (estimatedCost !== undefined) {
        maintenanceData.Cost = ValidationUtils.validateInteger(estimatedCost, 'estimatedCost', 0);
      }

      const result = await this.sp.web.lists.getByTitle(this.ASSET_MAINTENANCE_LIST).items.add(maintenanceData);
      return result.data.Id;
    } catch (error) {
      logger.error('AssetTrackingService', 'Error scheduling maintenance:', error);
      throw error;
    }
  }

  public async completeMaintenance(
    maintenanceId: number,
    performedById: number,
    actualCost?: number,
    notes?: string,
    nextMaintenanceDate?: Date
  ): Promise<void> {
    try {
      const validMaintenanceId = ValidationUtils.validateInteger(maintenanceId, 'maintenanceId', 1);
      const validPerformedById = ValidationUtils.validateInteger(performedById, 'performedById', 1);

      // Get maintenance record
      const maintenance = await this.sp.web.lists.getByTitle(this.ASSET_MAINTENANCE_LIST).items
        .getById(validMaintenanceId)
        .select('Id', 'AssetId', 'IsCompleted')();

      if (maintenance.IsCompleted) {
        throw new Error('Maintenance has already been completed');
      }

      // Update maintenance record
      const updateData: any = {
        CompletedDate: new Date().toISOString(),
        PerformedById: validPerformedById,
        IsCompleted: true
      };

      if (actualCost !== undefined) {
        updateData.Cost = ValidationUtils.validateInteger(actualCost, 'actualCost', 0);
      }

      if (notes) {
        updateData.Notes = ValidationUtils.sanitizeHtml(notes);
      }

      if (nextMaintenanceDate) {
        ValidationUtils.validateDate(nextMaintenanceDate, 'nextMaintenanceDate');
        updateData.NextMaintenanceDate = nextMaintenanceDate.toISOString();
      }

      await this.sp.web.lists.getByTitle(this.ASSET_MAINTENANCE_LIST).items
        .getById(validMaintenanceId)
        .update(updateData);

      // Update asset
      const assetUpdates: any = {
        LastMaintenanceDate: new Date(),
        Status: AssetStatus.Available
      };

      if (nextMaintenanceDate) {
        assetUpdates.NextMaintenanceDate = nextMaintenanceDate;
      }

      await this.assetService.updateAsset(maintenance.AssetId, assetUpdates);
    } catch (error) {
      logger.error('AssetTrackingService', 'Error completing maintenance:', error);
      throw error;
    }
  }

  public async getMaintenanceRecords(assetId?: number, upcomingOnly?: boolean): Promise<IAssetMaintenance[]> {
    try {
      let query = this.sp.web.lists.getByTitle(this.ASSET_MAINTENANCE_LIST).items
        .select(
          'Id', 'AssetId', 'MaintenanceType', 'ScheduledDate', 'CompletedDate',
          'PerformedById', 'PerformedBy/Title', 'Vendor', 'Cost', 'Description',
          'IsCompleted', 'IsCancelled', 'NextMaintenanceDate', 'PartsUsed',
          'LaborHours', 'Comments', 'Attachments', 'Created', 'Modified'
        )
        .expand('PerformedBy')
        .orderBy('ScheduledDate', false);

      const filters: string[] = [];

      if (assetId !== undefined) {
        const validAssetId = ValidationUtils.validateInteger(assetId, 'assetId', 1);
        filters.push(ValidationUtils.buildFilter('AssetId', 'eq', validAssetId));
      }

      if (upcomingOnly) {
        filters.push('IsCompleted eq 0');
        filters.push('IsCancelled eq 0');
        const now = new Date();
        filters.push(ValidationUtils.buildFilter('ScheduledDate', 'ge', now));
      }

      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      const items = await query.top(1000)();
      return items.map(this.mapMaintenanceFromSP);
    } catch (error) {
      logger.error('AssetTrackingService', 'Error getting maintenance records:', error);
      throw error;
    }
  }

  // ==================== Transfer Operations ====================

  public async requestTransfer(
    assetId: number,
    requestedById: number,
    toUserId?: number,
    toLocation?: string,
    toDepartment?: string,
    transferReason?: string
  ): Promise<number> {
    try {
      const validAssetId = ValidationUtils.validateInteger(assetId, 'assetId', 1);
      const validRequestedById = ValidationUtils.validateInteger(requestedById, 'requestedById', 1);

      // Get current asset info
      const asset = await this.assetService.getAssetById(validAssetId);

      const transferData: any = {
        AssetId: validAssetId,
        FromUserId: asset.AssignedToId,
        FromLocation: asset.Location,
        FromDepartment: asset.Department,
        TransferDate: new Date().toISOString(),
        RequestedById: validRequestedById,
        Status: 'Pending'
      };

      if (toUserId !== undefined) {
        transferData.ToUserId = ValidationUtils.validateInteger(toUserId, 'toUserId', 1);
      }

      if (toLocation) {
        transferData.ToLocation = ValidationUtils.sanitizeHtml(toLocation);
      }

      if (toDepartment) {
        transferData.ToDepartment = ValidationUtils.sanitizeHtml(toDepartment);
      }

      if (transferReason) {
        transferData.TransferReason = ValidationUtils.sanitizeHtml(transferReason);
      }

      const result = await this.sp.web.lists.getByTitle(this.ASSET_TRANSFERS_LIST).items.add(transferData);
      return result.data.Id;
    } catch (error) {
      logger.error('AssetTrackingService', 'Error requesting transfer:', error);
      throw error;
    }
  }

  public async approveTransfer(transferId: number, approvedById: number): Promise<void> {
    try {
      const validTransferId = ValidationUtils.validateInteger(transferId, 'transferId', 1);
      const validApprovedById = ValidationUtils.validateInteger(approvedById, 'approvedById', 1);

      // Get transfer record
      const transfer = await this.sp.web.lists.getByTitle(this.ASSET_TRANSFERS_LIST).items
        .getById(validTransferId)
        .select('Id', 'AssetId', 'ToUserId', 'ToLocation', 'ToDepartment', 'Status')();

      if (transfer.Status !== 'Pending') {
        throw new Error('Transfer is not in pending status');
      }

      // Update transfer record
      await this.sp.web.lists.getByTitle(this.ASSET_TRANSFERS_LIST).items
        .getById(validTransferId)
        .update({
          Status: 'Approved',
          ApprovedById: validApprovedById,
          ApprovalDate: new Date().toISOString()
        });

      // Update asset
      const assetUpdates: any = {};
      if (transfer.ToUserId) assetUpdates.AssignedToId = transfer.ToUserId;
      if (transfer.ToLocation) assetUpdates.Location = transfer.ToLocation;
      if (transfer.ToDepartment) assetUpdates.Department = transfer.ToDepartment;

      await this.assetService.updateAsset(transfer.AssetId, assetUpdates);

      // Mark transfer as completed
      await this.sp.web.lists.getByTitle(this.ASSET_TRANSFERS_LIST).items
        .getById(validTransferId)
        .update({ Status: 'Completed' });
    } catch (error) {
      logger.error('AssetTrackingService', 'Error approving transfer:', error);
      throw error;
    }
  }

  // ==================== JML Integration ====================

  public async assignAssetsForOnboarding(
    processId: number,
    employeeId: number,
    assetTypeIds: number[],
    assignedById: number
  ): Promise<IJMLAssetAssignment> {
    try {
      const validProcessId = ValidationUtils.validateInteger(processId, 'processId', 1);
      const validEmployeeId = ValidationUtils.validateInteger(employeeId, 'employeeId', 1);
      const validAssignedById = ValidationUtils.validateInteger(assignedById, 'assignedById', 1);

      const assignedAssets: number[] = [];

      // Find and assign available assets of the requested types
      for (const typeId of assetTypeIds) {
        const validTypeId = ValidationUtils.validateInteger(typeId, 'assetTypeId', 1);

        // Find available asset of this type
        const availableAssets = await this.assetService.getAssets({
          assetTypeId: validTypeId,
          status: [AssetStatus.Available]
        });

        if (availableAssets.length > 0) {
          const asset = availableAssets[0];
          const assignmentId = await this.assetService.assignAsset(
            asset.Id!,
            validEmployeeId,
            validAssignedById,
            `Assigned for JML Process #${validProcessId}`
          );

          // Update assignment with process link
          await this.sp.web.lists.getByTitle('Asset Assignments').items
            .getById(assignmentId)
            .update({ ProcessId: validProcessId });

          assignedAssets.push(asset.Id!);
        }
      }

      return {
        processId: validProcessId,
        employeeId: validEmployeeId,
        assetIds: assignedAssets,
        assignmentDate: new Date(),
        assignedById: validAssignedById
      };
    } catch (error) {
      logger.error('AssetTrackingService', 'Error assigning assets for onboarding:', error);
      throw error;
    }
  }

  public async retrieveAssetsForOffboarding(
    processId: number,
    employeeId: number,
    returnCondition: AssetCondition,
    returnNotes?: string
  ): Promise<number[]> {
    try {
      const validProcessId = ValidationUtils.validateInteger(processId, 'processId', 1);
      const validEmployeeId = ValidationUtils.validateInteger(employeeId, 'employeeId', 1);
      ValidationUtils.validateEnum(returnCondition, AssetCondition, 'returnCondition');

      // Get all assets assigned to this employee
      const assignedAssets = await this.assetService.getAssets({
        assignedToId: validEmployeeId,
        status: [AssetStatus.Assigned]
      });

      const retrievedAssetIds: number[] = [];

      for (const asset of assignedAssets) {
        await this.assetService.unassignAsset(asset.Id!, returnCondition, returnNotes);
        retrievedAssetIds.push(asset.Id!);
      }

      return retrievedAssetIds;
    } catch (error) {
      logger.error('AssetTrackingService', 'Error retrieving assets for offboarding:', error);
      throw error;
    }
  }

  // ==================== Bulk Operations ====================

  public async executeBulkOperation(operation: IAssetBulkOperation): Promise<number> {
    try {
      let successCount = 0;

      for (const assetId of operation.assetIds) {
        try {
          const validAssetId = ValidationUtils.validateInteger(assetId, 'assetId', 1);

          switch (operation.operation) {
            case 'Assign':
              if (operation.parameters?.assignToId) {
                await this.assetService.assignAsset(
                  validAssetId,
                  operation.parameters.assignToId,
                  operation.parameters.assignToId
                );
                successCount++;
              }
              break;

            case 'Unassign':
              await this.assetService.unassignAsset(validAssetId);
              successCount++;
              break;

            case 'ChangeStatus':
              if (operation.parameters?.status) {
                await this.assetService.updateAsset(validAssetId, {
                  Status: operation.parameters.status
                });
                successCount++;
              }
              break;

            case 'ChangeLocation':
              if (operation.parameters?.location) {
                await this.assetService.updateAsset(validAssetId, {
                  Location: operation.parameters.location
                });
                successCount++;
              }
              break;

            case 'Retire':
              await this.assetService.updateAsset(validAssetId, {
                Status: AssetStatus.Retired,
                RetirementDate: operation.parameters?.retirementDate || new Date(),
                RetirementReason: operation.parameters?.retirementReason
              });
              successCount++;
              break;

            case 'Delete':
              await this.assetService.deleteAsset(validAssetId);
              successCount++;
              break;
          }
        } catch (error) {
          logger.error('AssetTrackingService', `Error processing asset ${assetId}:`, error);
          // Continue with next asset
        }
      }

      return successCount;
    } catch (error) {
      logger.error('AssetTrackingService', 'Error executing bulk operation:', error);
      throw error;
    }
  }

  // ==================== Asset History ====================

  public async getAssetHistory(assetId: number): Promise<IAssetHistoryEntry[]> {
    try {
      const validAssetId = ValidationUtils.validateInteger(assetId, 'assetId', 1);

      const history: IAssetHistoryEntry[] = [];

      // Get assignments
      const assignments = await this.assetService.getAssetAssignments(validAssetId);
      for (const assignment of assignments) {
        history.push({
          id: `assignment-${assignment.Id}`,
          timestamp: assignment.AssignedDate,
          action: 'Assigned',
          performedById: assignment.AssignedById,
          performedBy: assignment.AssignedBy?.Title || 'Unknown',
          description: `Assigned to ${assignment.AssignedTo?.Title || 'Unknown User'}`,
          details: assignment
        });

        if (assignment.ActualReturnDate) {
          history.push({
            id: `return-${assignment.Id}`,
            timestamp: assignment.ActualReturnDate,
            action: 'Unassigned',
            performedById: assignment.AssignedById,
            performedBy: assignment.AssignedBy?.Title || 'Unknown',
            description: `Returned from ${assignment.AssignedTo?.Title || 'Unknown User'}`,
            details: assignment
          });
        }
      }

      // Get checkouts
      const checkouts = await this.getCheckouts({ assetId: validAssetId });
      for (const checkout of checkouts) {
        history.push({
          id: `checkout-${checkout.Id}`,
          timestamp: checkout.CheckoutDate,
          action: 'CheckedOut',
          performedById: checkout.CheckedOutById,
          performedBy: checkout.CheckedOutBy?.Title || 'Unknown',
          description: `Checked out to ${checkout.CheckedOutTo?.Title || 'Unknown User'}`,
          details: checkout
        });

        if (checkout.CheckedInDate) {
          history.push({
            id: `checkin-${checkout.Id}`,
            timestamp: checkout.CheckedInDate,
            action: 'CheckedIn',
            performedById: checkout.CheckedInById || checkout.CheckedOutById,
            performedBy: checkout.CheckedInBy?.Title || 'Unknown',
            description: `Checked in by ${checkout.CheckedInBy?.Title || 'Unknown User'}`,
            details: checkout
          });
        }
      }

      // Get maintenance
      const maintenance = await this.getMaintenanceRecords(validAssetId);
      for (const maint of maintenance) {
        history.push({
          id: `maintenance-${maint.Id}`,
          timestamp: maint.ScheduledDate,
          action: 'Maintenance',
          performedById: maint.PerformedById || 0,
          performedBy: maint.PerformedBy?.Title || 'Scheduled',
          description: `${maint.MaintenanceType} maintenance - ${maint.Description}`,
          details: maint
        });
      }

      // Sort by timestamp descending
      history.sort((a, b) => b.timestamp.getTime() - a.timestamp.getTime());

      return history;
    } catch (error) {
      logger.error('AssetTrackingService', 'Error getting asset history:', error);
      throw error;
    }
  }

  // ==================== Mapping Functions ====================

  private mapCheckoutFromSP(item: any): IAssetCheckout {
    return {
      Id: item.Id,
      AssetId: item.AssetId,
      CheckedOutToId: item.CheckedOutToId,
      CheckedOutTo: item.CheckedOutTo,
      CheckedOutById: item.CheckedOutById,
      CheckedOutBy: item.CheckedOutBy,
      CheckoutDate: new Date(item.CheckoutDate),
      ExpectedReturnDate: new Date(item.ExpectedReturnDate),
      ActualReturnDate: item.ActualReturnDate ? new Date(item.ActualReturnDate) : undefined,
      Purpose: item.Purpose,
      Location: item.Location,
      Status: item.Status as CheckoutStatus,
      IsOverdue: item.IsOverdue,
      CheckedInDate: item.CheckedInDate ? new Date(item.CheckedInDate) : undefined,
      CheckedInById: item.CheckedInById,
      CheckedInBy: item.CheckedInBy,
      ReturnCondition: item.ReturnCondition as AssetCondition,
      ReturnNotes: item.ReturnNotes,
      ReminderSent: item.ReminderSent,
      OverdueNotificationSent: item.OverdueNotificationSent,
      Notes: item.Notes,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }

  private mapMaintenanceFromSP(item: any): IAssetMaintenance {
    return {
      Id: item.Id,
      AssetId: item.AssetId,
      MaintenanceType: item.MaintenanceType as MaintenanceType,
      ScheduledDate: new Date(item.ScheduledDate),
      CompletedDate: item.ActualCompletionDate ? new Date(item.ActualCompletionDate) : undefined,
      PerformedById: item.PerformedById,
      PerformedBy: item.PerformedBy,
      Vendor: item.Vendor,
      Cost: item.Cost,
      Description: item.Description,
      IsCompleted: item.IsCompleted,
      IsCancelled: item.IsCancelled,
      NextMaintenanceDate: item.NextMaintenanceDate ? new Date(item.NextMaintenanceDate) : undefined,
      PartsUsed: item.PartsUsed,
      LaborHours: item.LaborHours,
      Notes: item.Notes,
      Attachments: item.Attachments,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }
}
