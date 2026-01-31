// @ts-nocheck
/**
 * Policy Request Service
 * Handles CRUD operations for policy requests submitted through the Request Policy wizard.
 * Persists to the PM_PolicyRequests SharePoint list.
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import {
  IPolicyRequest,
  IPolicyRequestFormData,
  IPolicyRequestSubmitResult,
  PolicyRequestStatus
} from '../models/IPolicyRequest';
import { PolicyLists } from '../constants/SharePointListNames';
import { NotificationLists } from '../constants/SharePointListNames';
import { logger } from './LoggingService';

/**
 * Generates a unique reference number for a policy request.
 * Format: PR-YYYYMMDD-XXXXX (e.g. PR-20260131-8A3F2)
 */
function generateReferenceNumber(): string {
  const now = new Date();
  const datePart = now.getFullYear().toString() +
    (now.getMonth() + 1).toString().padStart(2, '0') +
    now.getDate().toString().padStart(2, '0');
  const randomPart = Math.random().toString(36).substring(2, 7).toUpperCase();
  return `PR-${datePart}-${randomPart}`;
}

export class PolicyRequestService {
  private sp: SPFI;
  private readonly LIST_NAME = PolicyLists.POLICY_REQUESTS;
  private currentUserId: number = 0;
  private currentUserEmail: string = '';
  private currentUserName: string = '';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Initialize current user context. Call once before submitting.
   */
  public async initCurrentUser(): Promise<void> {
    try {
      const user = await this.sp.web.currentUser();
      this.currentUserId = user.Id;
      this.currentUserEmail = user.Email || '';
      this.currentUserName = user.Title || '';
    } catch (err) {
      logger.error('PolicyRequestService', 'Failed to get current user', err);
    }
  }

  /**
   * Submit a new policy request from the wizard form data.
   * Converts form data to a SharePoint list item and persists it.
   */
  public async submitRequest(
    formData: IPolicyRequestFormData,
    userName?: string,
    userEmail?: string
  ): Promise<IPolicyRequestSubmitResult> {
    try {
      // Ensure we have user info
      if (!this.currentUserName) {
        await this.initCurrentUser();
      }

      const referenceNumber = generateReferenceNumber();
      const requestedBy = userName || this.currentUserName || 'Unknown';
      const requestedByEmail = userEmail || this.currentUserEmail || '';

      const listItem: Record<string, any> = {
        Title: formData.policyTitle,
        RequestedBy: requestedBy,
        RequestedByEmail: requestedByEmail,
        RequestedByDepartment: '', // Will be populated if department info is available
        PolicyCategory: formData.policyCategory,
        PolicyType: formData.policyType,
        Priority: formData.priority,
        TargetAudience: formData.targetAudience,
        BusinessJustification: formData.businessJustification,
        RegulatoryDriver: formData.regulatoryDriver || '',
        DesiredEffectiveDate: formData.desiredEffectiveDate || null,
        ReadTimeframeDays: parseInt(formData.readTimeframeDays, 10) || 7,
        RequiresAcknowledgement: formData.requiresAcknowledgement,
        RequiresQuiz: formData.requiresQuiz,
        AdditionalNotes: formData.additionalNotes || '',
        NotifyAuthors: formData.notifyAuthors,
        PreferredAuthor: formData.preferredAuthor || '',
        Status: 'New' as PolicyRequestStatus,
        ReferenceNumber: referenceNumber,
        AssignedAuthor: '',
        AssignedAuthorEmail: ''
      };

      const result = await this.sp.web.lists.getByTitle(this.LIST_NAME).items.add(listItem);

      logger.info('PolicyRequestService', `Policy request submitted: ${referenceNumber}`, {
        itemId: result.Id,
        title: formData.policyTitle,
        requestedBy
      });

      // Log notification for policy authors (non-blocking)
      if (formData.notifyAuthors) {
        this.logRequestNotification(referenceNumber, formData, requestedBy).catch(err => {
          logger.warn('PolicyRequestService', 'Failed to log author notification (non-blocking)', err);
        });
      }

      return {
        success: true,
        referenceNumber,
        itemId: result.Id
      };
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Unknown error occurred';
      logger.error('PolicyRequestService', 'Failed to submit policy request', err);

      return {
        success: false,
        error: errorMessage
      };
    }
  }

  /**
   * Get all policy requests (for Author/Admin views)
   */
  public async getRequests(status?: PolicyRequestStatus): Promise<IPolicyRequest[]> {
    try {
      let items;
      if (status) {
        items = await this.sp.web.lists.getByTitle(this.LIST_NAME).items
          .filter(`Status eq '${status}'`)
          .orderBy('Created', false)
          .top(200)();
      } else {
        items = await this.sp.web.lists.getByTitle(this.LIST_NAME).items
          .orderBy('Created', false)
          .top(200)();
      }

      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title || '',
        RequestedBy: item.RequestedBy || '',
        RequestedByEmail: item.RequestedByEmail || '',
        RequestedByDepartment: item.RequestedByDepartment || '',
        PolicyCategory: item.PolicyCategory || '',
        PolicyType: item.PolicyType || 'New Policy',
        Priority: item.Priority || 'Medium',
        TargetAudience: item.TargetAudience || '',
        BusinessJustification: item.BusinessJustification || '',
        RegulatoryDriver: item.RegulatoryDriver || '',
        DesiredEffectiveDate: item.DesiredEffectiveDate || '',
        ReadTimeframeDays: item.ReadTimeframeDays || 7,
        RequiresAcknowledgement: item.RequiresAcknowledgement ?? true,
        RequiresQuiz: item.RequiresQuiz ?? false,
        AdditionalNotes: item.AdditionalNotes || '',
        NotifyAuthors: item.NotifyAuthors ?? true,
        PreferredAuthor: item.PreferredAuthor || '',
        AttachmentUrls: item.AttachmentUrls ? JSON.parse(item.AttachmentUrls) : [],
        Status: item.Status || 'New',
        AssignedAuthor: item.AssignedAuthor || '',
        AssignedAuthorEmail: item.AssignedAuthorEmail || '',
        ReferenceNumber: item.ReferenceNumber || '',
        Created: item.Created,
        Modified: item.Modified
      }));
    } catch (err) {
      logger.error('PolicyRequestService', 'Failed to get policy requests', err);
      return [];
    }
  }

  /**
   * Get requests submitted by a specific user
   */
  public async getMyRequests(email: string): Promise<IPolicyRequest[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.LIST_NAME).items
        .filter(`RequestedByEmail eq '${email}'`)
        .orderBy('Created', false)
        .top(100)();

      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title || '',
        RequestedBy: item.RequestedBy || '',
        RequestedByEmail: item.RequestedByEmail || '',
        RequestedByDepartment: item.RequestedByDepartment || '',
        PolicyCategory: item.PolicyCategory || '',
        PolicyType: item.PolicyType || 'New Policy',
        Priority: item.Priority || 'Medium',
        TargetAudience: item.TargetAudience || '',
        BusinessJustification: item.BusinessJustification || '',
        RegulatoryDriver: item.RegulatoryDriver || '',
        DesiredEffectiveDate: item.DesiredEffectiveDate || '',
        ReadTimeframeDays: item.ReadTimeframeDays || 7,
        RequiresAcknowledgement: item.RequiresAcknowledgement ?? true,
        RequiresQuiz: item.RequiresQuiz ?? false,
        AdditionalNotes: item.AdditionalNotes || '',
        NotifyAuthors: item.NotifyAuthors ?? true,
        PreferredAuthor: item.PreferredAuthor || '',
        AttachmentUrls: item.AttachmentUrls ? JSON.parse(item.AttachmentUrls) : [],
        Status: item.Status || 'New',
        AssignedAuthor: item.AssignedAuthor || '',
        AssignedAuthorEmail: item.AssignedAuthorEmail || '',
        ReferenceNumber: item.ReferenceNumber || '',
        Created: item.Created,
        Modified: item.Modified
      }));
    } catch (err) {
      logger.error('PolicyRequestService', 'Failed to get user requests', err);
      return [];
    }
  }

  /**
   * Update request status (for Author/Admin use)
   */
  public async updateRequestStatus(
    itemId: number,
    status: PolicyRequestStatus,
    assignedAuthor?: string,
    assignedAuthorEmail?: string
  ): Promise<boolean> {
    try {
      const updateFields: Record<string, any> = { Status: status };
      if (assignedAuthor) updateFields.AssignedAuthor = assignedAuthor;
      if (assignedAuthorEmail) updateFields.AssignedAuthorEmail = assignedAuthorEmail;

      await this.sp.web.lists.getByTitle(this.LIST_NAME).items.getById(itemId).update(updateFields);

      logger.info('PolicyRequestService', `Request ${itemId} status updated to ${status}`);
      return true;
    } catch (err) {
      logger.error('PolicyRequestService', `Failed to update request ${itemId}`, err);
      return false;
    }
  }

  /**
   * Log a notification entry for the Policy Authoring team about a new request.
   * Uses the PM_PolicyNotifications list for audit and in-app notification tracking.
   */
  private async logRequestNotification(
    referenceNumber: string,
    formData: IPolicyRequestFormData,
    requestedBy: string
  ): Promise<void> {
    try {
      const notificationList = NotificationLists.POLICY_NOTIFICATIONS;
      const priorityLabel = formData.priority === 'Critical' ? 'CRITICAL: ' : formData.priority === 'High' ? 'URGENT: ' : '';

      await this.sp.web.lists.getByTitle(notificationList).items.add({
        Title: `${priorityLabel}New Policy Request: ${formData.policyTitle}`,
        NotificationType: 'NewPolicy',
        Subject: `New Policy Request [${referenceNumber}]: ${formData.policyTitle}`,
        Body: [
          `A new policy request has been submitted by ${requestedBy}.`,
          ``,
          `Reference: ${referenceNumber}`,
          `Title: ${formData.policyTitle}`,
          `Category: ${formData.policyCategory}`,
          `Priority: ${formData.priority}`,
          `Type: ${formData.policyType}`,
          `Target Audience: ${formData.targetAudience}`,
          ``,
          `Business Justification:`,
          formData.businessJustification,
          formData.regulatoryDriver ? `\nRegulatory Driver: ${formData.regulatoryDriver}` : '',
          formData.preferredAuthor ? `\nPreferred Author: ${formData.preferredAuthor}` : '',
        ].join('\n'),
        Status: 'Pending',
        Priority: formData.priority,
        SendEmail: formData.notifyAuthors,
        SendInApp: true
      });

      logger.info('PolicyRequestService', `Notification logged for request ${referenceNumber}`);
    } catch (err) {
      // Non-critical failure — log but don't throw
      logger.warn('PolicyRequestService', 'Failed to log notification entry', err);
    }
  }
}
