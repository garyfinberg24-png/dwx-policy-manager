// Request Service
// Asset request workflow management

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import { IAssetRequest, AssetCategory } from '../models/IAsset';
import { AM_LISTS } from '../constants/SharePointListNames';

export interface IRequestStatistics {
  total: number;
  pending: number;
  approved: number;
  rejected: number;
  fulfilled: number;
  byPriority: { [key: string]: number };
}

export class RequestService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  public async getRequests(): Promise<IAssetRequest[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(AM_LISTS.ASSET_REQUESTS).items
        .select(
          'Id', 'Title', 'RequestedById', 'RequestedBy/Title', 'RequestedBy/EMail',
          'AssetCategory', 'Description', 'Justification', 'Quantity',
          'Priority', 'RequiredByDate', 'EstimatedCost', 'BudgetCode',
          'Status', 'ApprovedById', 'ApprovedBy/Title', 'Department',
          'RejectionReason', 'FulfilledDate', 'FulfilledAssetIds', 'Notes',
          'Created', 'Modified'
        )
        .expand('RequestedBy', 'ApprovedBy')
        .orderBy('Created', false)
        .top(5000)();

      return items.map(this.mapFromSP);
    } catch (error) {
      console.error('[RequestService] Error getting requests:', error);
      throw error;
    }
  }

  public async getMyRequests(): Promise<IAssetRequest[]> {
    try {
      const currentUser = await this.sp.web.currentUser();
      const items = await this.sp.web.lists.getByTitle(AM_LISTS.ASSET_REQUESTS).items
        .select(
          'Id', 'Title', 'RequestedById', 'RequestedBy/Title', 'RequestedBy/EMail',
          'AssetCategory', 'Description', 'Justification', 'Quantity',
          'Priority', 'RequiredByDate', 'EstimatedCost', 'BudgetCode',
          'Status', 'ApprovedById', 'ApprovedBy/Title', 'Department',
          'RejectionReason', 'FulfilledDate', 'FulfilledAssetIds', 'Notes',
          'Created', 'Modified'
        )
        .expand('RequestedBy', 'ApprovedBy')
        .filter(`RequestedById eq ${currentUser.Id}`)
        .orderBy('Created', false)
        .top(500)();

      return items.map(this.mapFromSP);
    } catch (error) {
      console.error('[RequestService] Error getting my requests:', error);
      throw error;
    }
  }

  public async submitRequest(request: Partial<IAssetRequest>): Promise<number> {
    try {
      if (!request.Category || !request.Description) {
        throw new Error('Category and Description are required');
      }

      const currentUser = await this.sp.web.currentUser();

      const itemData: any = {
        Title: `${request.Category} Request - ${new Date().toLocaleDateString()}`,
        RequestedById: currentUser.Id,
        AssetCategory: request.Category,
        Description: request.Description,
        Priority: request.Priority || 'Medium',
        Status: 'Pending',
        Quantity: request.Quantity || 1
      };

      if (request.Justification) itemData.Justification = request.Justification;
      if (request.RequiredByDate) itemData.RequiredByDate = request.RequiredByDate;
      if (request.EstimatedCost !== undefined) itemData.EstimatedCost = request.EstimatedCost;
      if (request.BudgetCode) itemData.BudgetCode = request.BudgetCode;
      if (request.Notes) itemData.Notes = request.Notes;

      const result = await this.sp.web.lists.getByTitle(AM_LISTS.ASSET_REQUESTS).items.add(itemData);
      return result.data.Id;
    } catch (error) {
      console.error('[RequestService] Error submitting request:', error);
      throw error;
    }
  }

  public async approveRequest(id: number): Promise<void> {
    try {
      const currentUser = await this.sp.web.currentUser();

      await this.sp.web.lists.getByTitle(AM_LISTS.ASSET_REQUESTS).items.getById(id).update({
        Status: 'Approved',
        ApprovedById: currentUser.Id
      });
    } catch (error) {
      console.error('[RequestService] Error approving request:', error);
      throw error;
    }
  }

  public async rejectRequest(id: number, reason: string): Promise<void> {
    try {
      const currentUser = await this.sp.web.currentUser();

      await this.sp.web.lists.getByTitle(AM_LISTS.ASSET_REQUESTS).items.getById(id).update({
        Status: 'Rejected',
        ApprovedById: currentUser.Id,
        RejectionReason: reason
      });
    } catch (error) {
      console.error('[RequestService] Error rejecting request:', error);
      throw error;
    }
  }

  public async fulfillRequest(id: number, assetIds: number[]): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(AM_LISTS.ASSET_REQUESTS).items.getById(id).update({
        Status: 'Fulfilled',
        FulfilledDate: new Date().toISOString(),
        FulfilledAssetIds: assetIds.join(',')
      });
    } catch (error) {
      console.error('[RequestService] Error fulfilling request:', error);
      throw error;
    }
  }

  public async getRequestStats(): Promise<IRequestStatistics> {
    try {
      const requests = await this.getRequests();

      return {
        total: requests.length,
        pending: requests.filter(r => r.Status === 'Pending').length,
        approved: requests.filter(r => r.Status === 'Approved').length,
        rejected: requests.filter(r => r.Status === 'Rejected').length,
        fulfilled: requests.filter(r => r.Status === 'Fulfilled').length,
        byPriority: {
          Low: requests.filter(r => r.Priority === 'Low').length,
          Medium: requests.filter(r => r.Priority === 'Medium').length,
          High: requests.filter(r => r.Priority === 'High').length,
          Urgent: requests.filter(r => r.Priority === 'Urgent').length
        }
      };
    } catch (error) {
      console.error('[RequestService] Error getting request stats:', error);
      throw error;
    }
  }

  private mapFromSP(item: any): IAssetRequest {
    return {
      Id: item.Id,
      RequestedById: item.RequestedById,
      RequestedBy: item.RequestedBy,
      RequestDate: item.Created ? new Date(item.Created) : new Date(),
      AssetTypeId: item.AssetTypeId,
      Category: item.AssetCategory as AssetCategory,
      Description: item.Description || '',
      Justification: item.Justification,
      Quantity: item.Quantity || 1,
      Priority: item.Priority || 'Medium',
      RequiredByDate: item.RequiredByDate ? new Date(item.RequiredByDate) : undefined,
      EstimatedCost: item.EstimatedCost,
      BudgetCode: item.BudgetCode,
      Status: item.Status || 'Pending',
      ApprovedById: item.ApprovedById,
      ApprovedBy: item.ApprovedBy,
      ApprovalDate: item.ApprovalDate ? new Date(item.ApprovalDate) : undefined,
      RejectionReason: item.RejectionReason,
      FulfilledDate: item.FulfilledDate ? new Date(item.FulfilledDate) : undefined,
      FulfilledAssetIds: item.FulfilledAssetIds,
      Notes: item.Notes,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }
}
