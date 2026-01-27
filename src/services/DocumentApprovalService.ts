// @ts-nocheck
/**
 * Document Approval Service
 * Manages approval workflows for generated documents
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users';
import { IUser } from '../models';
import { logger } from './LoggingService';

/**
 * Approval status enumeration
 */
export enum ApprovalStatus {
  Draft = 'Draft',
  PendingApproval = 'PendingApproval',
  Approved = 'Approved',
  Rejected = 'Rejected',
  Cancelled = 'Cancelled'
}

/**
 * Approver information
 */
export interface IApprover {
  user: IUser;
  status: 'Pending' | 'Approved' | 'Rejected';
  responseDate?: Date;
  comments?: string;
  order: number;
}

/**
 * Approval request interface
 */
export interface IApprovalRequest {
  id: number;
  documentId: number;
  documentTitle: string;
  templateId: number;
  templateName: string;
  processId?: number;
  status: ApprovalStatus;
  requestedBy: IUser;
  requestedDate: Date;
  approvers: IApprover[];
  dueDate?: Date;
  requestComments?: string;
  priority: 'Low' | 'Normal' | 'High' | 'Urgent';
  documentUrl?: string;
}

/**
 * Create approval request options
 */
export interface ICreateApprovalOptions {
  documentId: number;
  documentTitle: string;
  templateId: number;
  templateName: string;
  processId?: number;
  approverIds: number[];
  dueDate?: Date;
  comments?: string;
  priority?: 'Low' | 'Normal' | 'High' | 'Urgent';
  documentUrl?: string;
}

/**
 * Document Approval Service
 */
export class DocumentApprovalService {
  private sp: SPFI;
  private readonly approvalListName: string;

  constructor(sp: SPFI, approvalListName?: string) {
    this.sp = sp;
    this.approvalListName = approvalListName || 'JML_DocumentApprovals';
  }

  /**
   * Create a new approval request
   */
  public async createApprovalRequest(options: ICreateApprovalOptions): Promise<IApprovalRequest> {
    try {
      const approvers: IApprover[] = [];
      for (let i = 0; i < options.approverIds.length; i++) {
        const userId = options.approverIds[i];
        try {
          const user = await this.sp.web.siteUsers.getById(userId)();
          approvers.push({
            user: { Id: user.Id, Title: user.Title, EMail: user.Email },
            status: 'Pending',
            order: i + 1
          });
        } catch {
          logger.warn('DocumentApprovalService', `Could not find user ${userId}`);
        }
      }

      if (approvers.length === 0) {
        throw new Error('No valid approvers found');
      }

      const newItemResult = await this.sp.web.lists
        .getByTitle(this.approvalListName)
        .items.add({
          Title: `Approval: ${options.documentTitle}`,
          DocumentId: options.documentId,
          DocumentTitle: options.documentTitle,
          TemplateId: options.templateId,
          TemplateName: options.templateName,
          ProcessId: options.processId,
          Status: ApprovalStatus.PendingApproval,
          Approvers: JSON.stringify(approvers),
          DueDate: options.dueDate,
          RequestComments: options.comments,
          Priority: options.priority || 'Normal',
          DocumentUrl: options.documentUrl
        });

      const newItemId = (newItemResult.data as { Id: number }).Id;
      logger.info('DocumentApprovalService', `Created approval request ${newItemId}`);
      return this.getApprovalRequest(newItemId);
    } catch (error) {
      logger.error('DocumentApprovalService', 'Failed to create approval request:', error);
      throw error;
    }
  }

  /**
   * Get an approval request by ID
   */
  public async getApprovalRequest(requestId: number): Promise<IApprovalRequest> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.approvalListName)
        .items.getById(requestId)
        .select('*', 'Author/Id', 'Author/Title', 'Author/EMail')
        .expand('Author')();

      return this.mapToApprovalRequest(item);
    } catch (error) {
      logger.error('DocumentApprovalService', 'Failed to get approval request:', error);
      throw error;
    }
  }

  /**
   * Get pending approvals for the current user
   */
  public async getMyPendingApprovals(): Promise<IApprovalRequest[]> {
    try {
      const currentUser = await this.sp.web.currentUser();
      const items = await this.sp.web.lists
        .getByTitle(this.approvalListName)
        .items.filter(`Status eq '${ApprovalStatus.PendingApproval}'`)
        .select('*', 'Author/Id', 'Author/Title', 'Author/EMail')
        .expand('Author')
        .orderBy('Created', false)();

      const requests: IApprovalRequest[] = [];
      for (let i = 0; i < items.length; i++) {
        const request = this.mapToApprovalRequest(items[i]);
        const isApprover = request.approvers.some(
          a => a.user.Id === currentUser.Id && a.status === 'Pending'
        );
        if (isApprover) {
          requests.push(request);
        }
      }

      return requests;
    } catch (error) {
      logger.error('DocumentApprovalService', 'Failed to get pending approvals:', error);
      return [];
    }
  }

  /**
   * Submit an approval response
   */
  public async submitResponse(
    requestId: number,
    approved: boolean,
    comments?: string
  ): Promise<IApprovalRequest> {
    try {
      const currentUser = await this.sp.web.currentUser();
      const request = await this.getApprovalRequest(requestId);

      const approverIndex = request.approvers.findIndex(
        a => a.user.Id === currentUser.Id && a.status === 'Pending'
      );

      if (approverIndex === -1) {
        throw new Error('You are not a pending approver for this request');
      }

      request.approvers[approverIndex] = {
        ...request.approvers[approverIndex],
        status: approved ? 'Approved' : 'Rejected',
        responseDate: new Date(),
        comments
      };

      let newStatus = request.status;
      if (!approved) {
        newStatus = ApprovalStatus.Rejected;
      } else {
        const allApproved = request.approvers.every(a => a.status === 'Approved');
        if (allApproved) {
          newStatus = ApprovalStatus.Approved;
        }
      }

      await this.sp.web.lists
        .getByTitle(this.approvalListName)
        .items.getById(requestId)
        .update({
          Status: newStatus,
          Approvers: JSON.stringify(request.approvers)
        });

      logger.info('DocumentApprovalService', `Response submitted: ${approved ? 'Approved' : 'Rejected'}`);
      return this.getApprovalRequest(requestId);
    } catch (error) {
      logger.error('DocumentApprovalService', 'Failed to submit response:', error);
      throw error;
    }
  }

  /**
   * Cancel an approval request
   */
  public async cancelRequest(requestId: number, reason?: string): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.approvalListName)
        .items.getById(requestId)
        .update({
          Status: ApprovalStatus.Cancelled,
          RequestComments: reason ? `Cancelled: ${reason}` : 'Cancelled'
        });

      logger.info('DocumentApprovalService', `Request ${requestId} cancelled`);
    } catch (error) {
      logger.error('DocumentApprovalService', 'Failed to cancel request:', error);
      throw error;
    }
  }

  /**
   * Get approval statistics
   */
  public async getApprovalStats(): Promise<{
    pending: number;
    approved: number;
    rejected: number;
    myPending: number;
  }> {
    try {
      const currentUser = await this.sp.web.currentUser();
      const items = await this.sp.web.lists
        .getByTitle(this.approvalListName)
        .items.select('Status', 'Approvers')();

      let pending = 0, approved = 0, rejected = 0, myPending = 0;

      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        switch (item.Status) {
          case ApprovalStatus.PendingApproval:
            pending++;
            try {
              const approvers = JSON.parse(item.Approvers || '[]');
              if (approvers.some((a: IApprover) => a.user.Id === currentUser.Id && a.status === 'Pending')) {
                myPending++;
              }
            } catch { /* ignore */ }
            break;
          case ApprovalStatus.Approved:
            approved++;
            break;
          case ApprovalStatus.Rejected:
            rejected++;
            break;
        }
      }

      return { pending, approved, rejected, myPending };
    } catch (error) {
      logger.error('DocumentApprovalService', 'Failed to get stats:', error);
      return { pending: 0, approved: 0, rejected: 0, myPending: 0 };
    }
  }

  private mapToApprovalRequest(item: Record<string, unknown>): IApprovalRequest {
    let approvers: IApprover[] = [];
    try {
      approvers = JSON.parse(String(item.Approvers) || '[]');
    } catch { /* ignore */ }

    const author = item.Author as { Id?: number; Title?: string; EMail?: string } | undefined;

    return {
      id: Number(item.Id),
      documentId: Number(item.DocumentId),
      documentTitle: String(item.DocumentTitle || ''),
      templateId: Number(item.TemplateId),
      templateName: String(item.TemplateName || ''),
      processId: item.ProcessId ? Number(item.ProcessId) : undefined,
      status: String(item.Status) as ApprovalStatus,
      requestedBy: {
        Id: author?.Id || Number(item.AuthorId) || 0,
        Title: author?.Title || 'Unknown',
        EMail: author?.EMail || ''
      },
      requestedDate: new Date(String(item.Created)),
      approvers,
      dueDate: item.DueDate ? new Date(String(item.DueDate)) : undefined,
      requestComments: item.RequestComments ? String(item.RequestComments) : undefined,
      priority: (String(item.Priority) || 'Normal') as 'Low' | 'Normal' | 'High' | 'Urgent',
      documentUrl: item.DocumentUrl ? String(item.DocumentUrl) : undefined
    };
  }
}

export function createDocumentApprovalService(sp: SPFI, approvalListName?: string): DocumentApprovalService {
  return new DocumentApprovalService(sp, approvalListName);
}
