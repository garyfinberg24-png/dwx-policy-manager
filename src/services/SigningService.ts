// @ts-nocheck
/* eslint-disable @typescript-eslint/no-explicit-any */
// TODO: Fix CanViewOtherSignatures missing from ISigner interface
// Signing Service
// Core service for managing document signing requests and workflows
// Note: Some fields may not exist in the SharePoint list - mapping handles this gracefully

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/attachments';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import { IItemAddResult, IItemUpdateResult } from '@pnp/sp/items';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../utils/pnpjsConfig';

import {
  ISigningRequest,
  ISigningDocument,
  ISigningChain,
  ISigningLevel,
  ISigner,
  ISigningTemplate,
  ISigningAuditLog,
  ISignatureProviderConfig,
  ISigningBlock,
  ISigningSummary,
  ISigningAnalytics,
  ICreateSigningRequest,
  ICreateSignerConfig,
  ISignDocumentRequest,
  IDeclineSigningRequest,
  IDelegateSigningRequest,
  IVoidSigningRequest,
  IResendSigningRequest,
  IUpdateSigningRequest,
  ISigningRequestFilter,
  ISigningCertificate,
  SigningRequestStatus,
  SigningWorkflowType,
  SigningRequestType,
  SignerStatus,
  SignerRole,
  SignatureType,
  SignatureProvider,
  SigningAuditAction,
  SigningTemplateCategory,
  SignerAuthenticationMethod,
  SigningEscalationAction
} from '../models/ISigning';
import { IUser } from '../models/ICommon';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

/**
 * Signing Service - Core operations for signing requests
 */
export class SigningService {
  // Singleton instance and context tracking
  private static instance: SigningService | null = null;
  private static currentContextUrl: string | null = null;

  private sp: SPFI;
  private currentUserId: number = 0;
  private currentUserEmail: string = '';

  // SharePoint List Names
  private readonly REQUESTS_LIST = 'JML_SigningRequests';
  private readonly CHAINS_LIST = 'JML_SigningChains';
  private readonly SIGNERS_LIST = 'JML_Signers';
  private readonly TEMPLATES_LIST = 'JML_SigningTemplates';
  private readonly AUDIT_LIST = 'JML_SigningAuditLog';
  private readonly CONFIG_LIST = 'JML_SignatureConfig';
  private readonly DOCUMENTS_LIBRARY = 'JML_SigningDocuments';

  private listsVerified: boolean = false;
  private listsExist: boolean = false;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Get or create singleton instance of SigningService
   * @param context WebPart context for SharePoint access
   * @returns SigningService instance
   */
  public static getInstance(context: WebPartContext): SigningService {
    const siteUrl = context.pageContext?.web?.absoluteUrl || '';

    // Create new instance if none exists or if context changed (different site)
    if (!SigningService.instance || SigningService.currentContextUrl !== siteUrl) {
      const sp = getSP(context);
      SigningService.instance = new SigningService(sp);
      SigningService.currentContextUrl = siteUrl;
      logger.info('SigningService', `Created new singleton instance for site: ${siteUrl}`);
    }

    return SigningService.instance;
  }

  /**
   * Reset the singleton instance (useful for testing or site changes)
   */
  public static resetInstance(): void {
    SigningService.instance = null;
    SigningService.currentContextUrl = null;
  }

  /**
   * Initialize the service with current user context
   * @param forceRecheck - If true, forces re-verification of lists even if already checked
   */
  public async initialize(forceRecheck: boolean = false): Promise<void> {
    try {
      const currentUser = await this.sp.web.currentUser();
      this.currentUserId = currentUser.Id;
      this.currentUserEmail = currentUser.Email;
      logger.info('SigningService', `Initialized for user: ${currentUser.Email}`);

      // Reset verification state if forcing recheck
      if (forceRecheck) {
        this.listsVerified = false;
        this.listsExist = false;
      }

      // Verify required lists exist
      await this.verifyRequiredLists();
    } catch (error) {
      logger.error('SigningService', 'Failed to initialize:', error);
      throw error;
    }
  }

  /**
   * Check if required SharePoint lists exist
   */
  private async verifyRequiredLists(): Promise<void> {
    if (this.listsVerified) {
      console.log('[SigningService] Lists already verified, skipping check');
      return;
    }

    try {
      console.log(`[SigningService] Checking for list: "${this.REQUESTS_LIST}"`);
      // Try to access the main requests list
      const listInfo = await this.sp.web.lists.getByTitle(this.REQUESTS_LIST).select('Title', 'ItemCount')();
      console.log(`[SigningService] List found: "${listInfo.Title}" with ${listInfo.ItemCount} items`);
      this.listsExist = true;
      this.listsVerified = true;
      logger.info('SigningService', 'Required lists verified');
    } catch (error: any) {
      console.error('[SigningService] List verification error:', error);
      console.log('[SigningService] Error message:', error?.message);
      console.log('[SigningService] Error status:', error?.status);
      // List doesn't exist - this is expected for first-time setup
      if (error?.message?.includes('does not exist') || error?.status === 404) {
        this.listsExist = false;
        this.listsVerified = true;
        logger.warn('SigningService', `Required list "${this.REQUESTS_LIST}" not found - provisioning needed`);
        console.log(`[SigningService] List "${this.REQUESTS_LIST}" not found`);
      } else {
        console.error('[SigningService] Unexpected error during list verification:', error);
        throw error;
      }
    }
  }

  /**
   * Check if the service is properly configured with required lists
   */
  public isConfigured(): boolean {
    return this.listsExist;
  }

  /**
   * Get list of required SharePoint lists for provisioning
   */
  public getRequiredLists(): string[] {
    return [
      this.REQUESTS_LIST,
      this.CHAINS_LIST,
      this.SIGNERS_LIST,
      this.TEMPLATES_LIST,
      this.AUDIT_LIST,
      this.CONFIG_LIST,
      this.DOCUMENTS_LIBRARY
    ];
  }

  // ============================================
  // REQUEST OPERATIONS
  // ============================================

  /**
   * Create a new signing request
   */
  public async createSigningRequest(request: ICreateSigningRequest): Promise<ISigningRequest> {
    try {
      logger.info('SigningService', 'Creating signing request:', { title: request.title });

      // Generate request number
      const requestNumber = await this.generateRequestNumber();

      // Build signing chain
      const signingChain = this.buildSigningChain(request.workflowType, request.signers);

      // Calculate dates
      const now = new Date();
      const dueDate = request.dueDate || this.addDays(now, request.expirationDays || 30);
      const expirationDate = this.addDays(now, request.expirationDays || 30);

      // Create the request record
      const requestData = {
        Title: request.title,
        RequestNumber: requestNumber,
        Description: request.description || '',
        Status: request.sendImmediately ? SigningRequestStatus.Pending : SigningRequestStatus.Draft,
        RequestType: this.determineRequestType(request.signers),
        WorkflowType: request.workflowType,
        Priority: 'Medium',
        RequesterId: this.currentUserId,
        DocumentIds: JSON.stringify(request.documentIds),
        ProcessId: request.processId || null,
        ProcessType: request.processType || null,
        TemplateId: request.templateId || null,
        Provider: request.provider || SignatureProvider.Internal,
        SigningChainJSON: JSON.stringify(signingChain),
        SigningBlocksJSON: request.signingBlocks ? JSON.stringify(request.signingBlocks) : null,
        DueDate: dueDate.toISOString(),
        ExpirationDate: expirationDate.toISOString(),
        ReminderEnabled: request.reminderEnabled !== false,
        ReminderDays: request.reminderDays || 3,
        EscalationEnabled: request.escalationEnabled || false,
        EscalationDays: request.escalationDays || 7,
        EscalationAction: request.escalationAction || SigningEscalationAction.Notify,
        AllowDelegation: request.allowDelegation !== false,
        AllowDecline: request.allowDecline !== false,
        RequireComments: request.requireComments || false,
        RequireAccessCode: request.requireAccessCode || false,
        AccessCode: request.accessCode || null,
        EmailSubject: request.emailSubject || `Please sign: ${request.title}`,
        EmailMessage: request.emailMessage || '',
        Tags: request.tags ? JSON.stringify(request.tags) : null,
        Category: request.category || SigningTemplateCategory.Custom,
        MetadataJSON: request.metadata ? JSON.stringify(request.metadata) : null
      };

      const result: IItemAddResult = await this.sp.web.lists
        .getByTitle(this.REQUESTS_LIST)
        .items.add(requestData);

      const createdRequestId = result.data.Id;

      // Create signing chain record
      await this.createSigningChainRecord(createdRequestId, signingChain);

      // Create signer records
      await this.createSignerRecords(createdRequestId, request.signers);

      // Log audit entry
      await this.logAuditEntry({
        RequestId: createdRequestId,
        RequestNumber: requestNumber,
        Action: SigningAuditAction.Created,
        ActionById: this.currentUserId,
        Description: `Signing request created: ${request.title}`,
        Details: { documentCount: request.documentIds.length, signerCount: request.signers.length }
      });

      // If send immediately, trigger the workflow
      if (request.sendImmediately) {
        await this.sendForSignature(createdRequestId);
      }

      // Return the created request
      return await this.getSigningRequestById(createdRequestId);
    } catch (error) {
      logger.error('SigningService', 'Failed to create signing request:', error);
      throw error;
    }
  }

  /**
   * Get signing requests with optional filtering
   */
  public async getSigningRequests(filter?: ISigningRequestFilter): Promise<ISigningRequest[]> {
    try {
      let query = this.sp.web.lists
        .getByTitle(this.REQUESTS_LIST)
        .items.select(
          'Id', 'Title', 'RequestNumber', 'Description', 'Status', 'RequestType', 'WorkflowType',
          'Priority', 'RequesterId', 'Requester/Id', 'Requester/Title', 'Requester/EMail',
          'DocumentIds', 'ProcessId', 'ProcessType', 'TemplateId', 'Provider', 'ExternalEnvelopeId',
          'SigningChainJSON', 'SigningBlocksJSON', 'DueDate', 'ExpirationDate', 'CompletedDate',
          'SentDate', 'ReminderEnabled', 'ReminderDays', 'EscalationEnabled', 'EscalationDays',
          'EscalationAction', 'AllowDelegation', 'AllowDecline', 'RequireComments',
          'RequireAccessCode', 'EmailSubject', 'EmailMessage', 'Tags', 'Category', 'MetadataJSON',
          'CertificateUrl', 'Created', 'Modified', 'Author/Id', 'Author/Title', 'Editor/Id', 'Editor/Title'
        )
        .expand('Requester', 'Author', 'Editor');

      // Build filter
      const filters: string[] = [];

      if (filter) {
        if (filter.searchTerm) {
          const searchTerm = ValidationUtils.sanitizeForOData(filter.searchTerm);
          filters.push(`(substringof('${searchTerm}', Title) or substringof('${searchTerm}', RequestNumber))`);
        }

        if (filter.status && filter.status.length > 0) {
          const statusFilters = filter.status.map(s => `Status eq '${s}'`).join(' or ');
          filters.push(`(${statusFilters})`);
        }

        if (filter.workflowType && filter.workflowType.length > 0) {
          const workflowFilters = filter.workflowType.map(w => `WorkflowType eq '${w}'`).join(' or ');
          filters.push(`(${workflowFilters})`);
        }

        if (filter.provider && filter.provider.length > 0) {
          const providerFilters = filter.provider.map(p => `Provider eq '${p}'`).join(' or ');
          filters.push(`(${providerFilters})`);
        }

        if (filter.requesterId) {
          filters.push(`RequesterId eq ${filter.requesterId}`);
        }

        if (filter.processId) {
          filters.push(`ProcessId eq ${filter.processId}`);
        }

        if (filter.templateId) {
          filters.push(`TemplateId eq ${filter.templateId}`);
        }

        if (filter.fromDate) {
          filters.push(`Created ge datetime'${filter.fromDate.toISOString()}'`);
        }

        if (filter.toDate) {
          filters.push(`Created le datetime'${filter.toDate.toISOString()}'`);
        }

        if (filter.isOverdue) {
          const now = new Date().toISOString();
          filters.push(`DueDate lt datetime'${now}'`);
          filters.push(`Status ne '${SigningRequestStatus.Completed}'`);
          filters.push(`Status ne '${SigningRequestStatus.Cancelled}'`);
        }

        if (filter.priority && filter.priority.length > 0) {
          const priorityFilters = filter.priority.map(p => `Priority eq '${p}'`).join(' or ');
          filters.push(`(${priorityFilters})`);
        }
      }

      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      // Apply sorting
      const sortBy = filter?.sortBy || 'Created';
      const sortDir = filter?.sortDirection === 'asc';
      query = query.orderBy(sortBy, sortDir);

      // Apply pagination
      const pageSize = filter?.pageSize || 50;
      query = query.top(pageSize);

      if (filter?.pageNumber && filter.pageNumber > 1) {
        query = query.skip((filter.pageNumber - 1) * pageSize);
      }

      const items = await query();

      return items.map(item => this.mapRequestFromSP(item));
    } catch (error) {
      logger.error('SigningService', 'Failed to get signing requests:', error);
      throw error;
    }
  }

  /**
   * Get a single signing request by ID
   */
  public async getSigningRequestById(id: number): Promise<ISigningRequest> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.REQUESTS_LIST)
        .items.getById(id)
        .select(
          'Id', 'Title', 'RequestNumber', 'Description', 'Status', 'RequestType', 'WorkflowType',
          'Priority', 'RequesterId', 'Requester/Id', 'Requester/Title', 'Requester/EMail',
          'DocumentIds', 'ProcessId', 'ProcessType', 'TemplateId', 'Provider', 'ExternalEnvelopeId',
          'SigningChainJSON', 'SigningBlocksJSON', 'DueDate', 'ExpirationDate', 'CompletedDate',
          'SentDate', 'ReminderEnabled', 'ReminderDays', 'EscalationEnabled', 'EscalationDays',
          'EscalationAction', 'AllowDelegation', 'AllowDecline', 'RequireComments',
          'RequireAccessCode', 'EmailSubject', 'EmailMessage', 'Tags', 'Category', 'MetadataJSON',
          'CertificateUrl', 'Created', 'Modified', 'Author/Id', 'Author/Title', 'Editor/Id', 'Editor/Title'
        )
        .expand('Requester', 'Author', 'Editor')();

      const request = this.mapRequestFromSP(item);

      // Load signers
      request.SigningChain.Levels = await this.loadSignersForRequest(id);

      return request;
    } catch (error) {
      logger.error('SigningService', `Failed to get signing request ${id}:`, error);
      throw error;
    }
  }

  /**
   * Get signing requests pending current user's signature
   */
  public async getMyPendingSignatures(): Promise<ISigner[]> {
    // If lists don't exist, return empty array
    if (!this.listsExist) {
      return [];
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.select(
          'Id', 'Title', 'RequestId', 'ChainId', 'Level', 'Order',
          'SignerUserId', 'SignerEmail', 'SignerName', 'SignerRole', 'Status',
          'SignatureType', 'SentDate', 'DueDate', 'Request/Title', 'Request/RequestNumber'
        )
        .expand('Request')
        .filter(`SignerEmail eq ${ValidationUtils.sanitizeEmailForOData(this.currentUserEmail)} and (Status eq 'Sent' or Status eq 'Viewed')`)
        .orderBy('SentDate', true)();

      return items.map(item => this.mapSignerFromSP(item));
    } catch (error) {
      logger.error('SigningService', 'Failed to get pending signatures:', error);
      throw error;
    }
  }

  /**
   * Update a signing request
   */
  public async updateSigningRequest(id: number, updates: IUpdateSigningRequest): Promise<ISigningRequest> {
    try {
      const updateData: any = {};

      if (updates.title) updateData.Title = updates.title;
      if (updates.description !== undefined) updateData.Description = updates.description;
      if (updates.dueDate) updateData.DueDate = updates.dueDate.toISOString();
      if (updates.emailMessage !== undefined) updateData.EmailMessage = updates.emailMessage;
      if (updates.reminderDays !== undefined) updateData.ReminderDays = updates.reminderDays;
      if (updates.tags) updateData.Tags = JSON.stringify(updates.tags);
      if (updates.metadata) updateData.MetadataJSON = JSON.stringify(updates.metadata);

      await this.sp.web.lists
        .getByTitle(this.REQUESTS_LIST)
        .items.getById(id)
        .update(updateData);

      // Log audit entry
      await this.logAuditEntry({
        RequestId: id,
        Action: SigningAuditAction.Updated,
        ActionById: this.currentUserId,
        Description: 'Signing request updated',
        Details: updates
      });

      return await this.getSigningRequestById(id);
    } catch (error) {
      logger.error('SigningService', `Failed to update signing request ${id}:`, error);
      throw error;
    }
  }

  /**
   * Cancel/void a signing request
   */
  public async cancelSigningRequest(request: IVoidSigningRequest): Promise<void> {
    try {
      const currentRequest = await this.getSigningRequestById(request.requestId);

      // Validate request can be cancelled
      if ([SigningRequestStatus.Completed, SigningRequestStatus.Expired].includes(currentRequest.Status)) {
        throw new Error(`Cannot cancel request in ${currentRequest.Status} status`);
      }

      // Update status
      await this.sp.web.lists
        .getByTitle(this.REQUESTS_LIST)
        .items.getById(request.requestId)
        .update({
          Status: SigningRequestStatus.Cancelled
        });

      // Update all pending signers
      await this.updateSignersStatus(request.requestId, SignerStatus.Voided);

      // If using external provider, void the envelope
      if (currentRequest.Provider !== SignatureProvider.Internal && currentRequest.ExternalEnvelopeId) {
        await this.voidExternalEnvelope(currentRequest);
      }

      // Log audit entry
      await this.logAuditEntry({
        RequestId: request.requestId,
        Action: SigningAuditAction.Cancelled,
        ActionById: this.currentUserId,
        Description: `Request cancelled: ${request.reason}`,
        Details: { reason: request.reason, notifySigners: request.notifySigners }
      });

      // Send notifications if requested
      if (request.notifySigners) {
        // Notification logic would go here
      }

      logger.info('SigningService', `Cancelled signing request ${request.requestId}`);
    } catch (error) {
      logger.error('SigningService', `Failed to cancel signing request ${request.requestId}:`, error);
      throw error;
    }
  }

  /**
   * Delete a signing request (draft only)
   */
  public async deleteSigningRequest(id: number): Promise<void> {
    try {
      const request = await this.getSigningRequestById(id);

      if (request.Status !== SigningRequestStatus.Draft) {
        throw new Error('Only draft requests can be deleted');
      }

      // Delete signers
      const signers = await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.filter(`RequestId eq ${id}`)();

      for (const signer of signers) {
        await this.sp.web.lists
          .getByTitle(this.SIGNERS_LIST)
          .items.getById(signer.Id)
          .delete();
      }

      // Delete chain
      const chains = await this.sp.web.lists
        .getByTitle(this.CHAINS_LIST)
        .items.filter(`RequestId eq ${id}`)();

      for (const chain of chains) {
        await this.sp.web.lists
          .getByTitle(this.CHAINS_LIST)
          .items.getById(chain.Id)
          .delete();
      }

      // Delete request
      await this.sp.web.lists
        .getByTitle(this.REQUESTS_LIST)
        .items.getById(id)
        .delete();

      logger.info('SigningService', `Deleted signing request ${id}`);
    } catch (error) {
      logger.error('SigningService', `Failed to delete signing request ${id}:`, error);
      throw error;
    }
  }

  // ============================================
  // WORKFLOW OPERATIONS
  // ============================================

  /**
   * Send request for signature
   */
  public async sendForSignature(requestId: number): Promise<void> {
    try {
      const request = await this.getSigningRequestById(requestId);

      if (request.Status !== SigningRequestStatus.Draft && request.Status !== SigningRequestStatus.Pending) {
        throw new Error(`Cannot send request in ${request.Status} status`);
      }

      // Update request status
      await this.sp.web.lists
        .getByTitle(this.REQUESTS_LIST)
        .items.getById(requestId)
        .update({
          Status: SigningRequestStatus.InProgress,
          SentDate: new Date().toISOString()
        });

      // Get first level signers
      const firstLevelSigners = await this.getSignersForLevel(requestId, 1);

      // Send to external provider or handle internally
      if (request.Provider !== SignatureProvider.Internal) {
        await this.sendToExternalProvider(request, firstLevelSigners);
      } else {
        // Send internal notifications
        for (const signer of firstLevelSigners) {
          await this.sendSignerNotification(requestId, signer.Id!);
        }
      }

      // Update signers status to Sent
      for (const signer of firstLevelSigners) {
        await this.sp.web.lists
          .getByTitle(this.SIGNERS_LIST)
          .items.getById(signer.Id!)
          .update({
            Status: SignerStatus.Sent,
            SentDate: new Date().toISOString()
          });
      }

      // Log audit entry
      await this.logAuditEntry({
        RequestId: requestId,
        Action: SigningAuditAction.Sent,
        ActionById: this.currentUserId,
        Description: `Request sent to ${firstLevelSigners.length} signer(s)`,
        Details: { signerCount: firstLevelSigners.length, provider: request.Provider }
      });

      logger.info('SigningService', `Sent signing request ${requestId} to ${firstLevelSigners.length} signers`);
    } catch (error) {
      logger.error('SigningService', `Failed to send signing request ${requestId}:`, error);
      throw error;
    }
  }

  /**
   * Resend to a specific signer
   */
  public async resendToSigner(request: IResendSigningRequest): Promise<void> {
    try {
      if (!request.signerId) {
        throw new Error('Signer ID is required');
      }

      const signer = await this.getSignerById(request.signerId);

      if (signer.Status === SignerStatus.Signed) {
        throw new Error('Cannot resend to signer who has already signed');
      }

      // Send notification
      await this.sendSignerNotification(request.requestId, request.signerId, request.message);

      // Update reminder count
      await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.getById(request.signerId)
        .update({
          RemindersSent: (signer.RemindersSent || 0) + 1,
          LastReminderDate: new Date().toISOString()
        });

      // Log audit entry
      await this.logAuditEntry({
        RequestId: request.requestId,
        SignerId: request.signerId,
        Action: SigningAuditAction.Resent,
        ActionById: this.currentUserId,
        Description: `Reminder sent to ${signer.SignerEmail}`
      });

      logger.info('SigningService', `Resent to signer ${request.signerId}`);
    } catch (error) {
      logger.error('SigningService', `Failed to resend to signer ${request.signerId}:`, error);
      throw error;
    }
  }

  /**
   * Recall/void a request
   */
  public async recallRequest(requestId: number, reason: string): Promise<void> {
    return this.cancelSigningRequest({ requestId, reason, notifySigners: true });
  }

  // ============================================
  // SIGNER OPERATIONS
  // ============================================

  /**
   * Sign a document
   */
  public async signDocument(request: ISignDocumentRequest): Promise<void> {
    try {
      const signer = await this.getSignerById(request.signerId);
      const signingRequest = await this.getSigningRequestById(request.requestId);

      // Validate signer can sign
      if (signer.Status !== SignerStatus.Sent && signer.Status !== SignerStatus.Viewed) {
        throw new Error(`Cannot sign - signer status is ${signer.Status}`);
      }

      // Validate access code if required
      if (signingRequest.RequireAccessCode && request.accessCode !== signer.AccessCode) {
        await this.logAuditEntry({
          RequestId: request.requestId,
          SignerId: request.signerId,
          Action: SigningAuditAction.AuthenticationFailed,
          Description: 'Invalid access code provided'
        });
        throw new Error('Invalid access code');
      }

      // Update signer record
      await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.getById(request.signerId)
        .update({
          Status: SignerStatus.Signed,
          SignedDate: new Date().toISOString(),
          SignatureDataJSON: JSON.stringify(request.signatureData),
          Comments: request.comments || null,
          IPAddress: this.getClientIP(),
          UserAgent: this.getUserAgent()
        });

      // Update completed blocks
      if (request.completedBlocks) {
        const completedBlockIds = request.completedBlocks.map(b => b.blockId);
        await this.sp.web.lists
          .getByTitle(this.SIGNERS_LIST)
          .items.getById(request.signerId)
          .update({
            CompletedBlockIds: JSON.stringify(completedBlockIds)
          });
      }

      // Log audit entry
      await this.logAuditEntry({
        RequestId: request.requestId,
        SignerId: request.signerId,
        SignerEmail: signer.SignerEmail,
        SignerName: signer.SignerName,
        Action: SigningAuditAction.Signed,
        ActionById: this.currentUserId,
        Description: `Document signed by ${signer.SignerName}`,
        Details: {
          signatureType: request.signatureData.type,
          blocksCompleted: request.completedBlocks?.length || 0
        }
      });

      // Check if level is complete and advance workflow
      await this.evaluateAndAdvanceWorkflow(request.requestId, signer.Level);

      logger.info('SigningService', `Document signed by ${signer.SignerEmail}`);
    } catch (error) {
      logger.error('SigningService', `Failed to sign document:`, error);
      throw error;
    }
  }

  /**
   * Decline to sign
   */
  public async declineToSign(request: IDeclineSigningRequest): Promise<void> {
    try {
      const signer = await this.getSignerById(request.signerId);
      const signingRequest = await this.getSigningRequestById(request.requestId);

      if (!signingRequest.AllowDecline) {
        throw new Error('Declining is not allowed for this request');
      }

      // Update signer record
      await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.getById(request.signerId)
        .update({
          Status: SignerStatus.Declined,
          DeclinedDate: new Date().toISOString(),
          DeclineReason: request.reason,
          Comments: request.comments || null
        });

      // Update request status
      await this.sp.web.lists
        .getByTitle(this.REQUESTS_LIST)
        .items.getById(request.requestId)
        .update({
          Status: SigningRequestStatus.Declined
        });

      // Log audit entry
      await this.logAuditEntry({
        RequestId: request.requestId,
        SignerId: request.signerId,
        SignerEmail: signer.SignerEmail,
        SignerName: signer.SignerName,
        Action: SigningAuditAction.Declined,
        ActionById: this.currentUserId,
        Description: `Document declined by ${signer.SignerName}: ${request.reason}`
      });

      // Notify requester
      // Notification logic would go here

      logger.info('SigningService', `Document declined by ${signer.SignerEmail}`);
    } catch (error) {
      logger.error('SigningService', `Failed to decline signing:`, error);
      throw error;
    }
  }

  /**
   * Delegate signature to another person
   */
  public async delegateSignature(request: IDelegateSigningRequest): Promise<void> {
    try {
      const signer = await this.getSignerById(request.signerId);
      const signingRequest = await this.getSigningRequestById(request.requestId);

      if (!signingRequest.AllowDelegation) {
        throw new Error('Delegation is not allowed for this request');
      }

      // Update original signer
      await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.getById(request.signerId)
        .update({
          Status: SignerStatus.Delegated,
          DelegatedToEmail: request.delegateToEmail,
          DelegatedToName: request.delegateToName,
          DelegationReason: request.reason,
          DelegationDate: new Date().toISOString()
        });

      // Create new signer record for delegate
      await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.add({
          Title: request.delegateToName,
          RequestId: request.requestId,
          ChainId: signer.ChainId,
          Level: signer.Level,
          Order: signer.Order,
          SignerEmail: request.delegateToEmail,
          SignerName: request.delegateToName,
          SignerPhone: request.delegateToPhone || null,
          SignerRole: signer.Role,
          Status: SignerStatus.Sent,
          SignatureType: signer.SignatureType,
          AuthenticationMethod: signer.AuthenticationMethod,
          DelegatedById: signer.Id,
          SentDate: new Date().toISOString()
        });

      // Log audit entry
      await this.logAuditEntry({
        RequestId: request.requestId,
        SignerId: request.signerId,
        Action: SigningAuditAction.Delegated,
        ActionById: this.currentUserId,
        Description: `Signature delegated from ${signer.SignerEmail} to ${request.delegateToEmail}`,
        Details: { reason: request.reason }
      });

      // Send notification to delegate
      // Notification logic would go here

      logger.info('SigningService', `Signature delegated to ${request.delegateToEmail}`);
    } catch (error) {
      logger.error('SigningService', `Failed to delegate signature:`, error);
      throw error;
    }
  }

  /**
   * Record document viewed by signer
   */
  public async recordDocumentViewed(requestId: number, signerId: number): Promise<void> {
    try {
      const signer = await this.getSignerById(signerId);

      if (signer.Status === SignerStatus.Sent) {
        await this.sp.web.lists
          .getByTitle(this.SIGNERS_LIST)
          .items.getById(signerId)
          .update({
            Status: SignerStatus.Viewed,
            ViewedDate: new Date().toISOString()
          });

        await this.logAuditEntry({
          RequestId: requestId,
          SignerId: signerId,
          SignerEmail: signer.SignerEmail,
          Action: SigningAuditAction.Viewed,
          Description: `Document viewed by ${signer.SignerEmail}`
        });
      }
    } catch (error) {
      logger.error('SigningService', `Failed to record document viewed:`, error);
    }
  }

  // ============================================
  // TEMPLATE OPERATIONS
  // ============================================

  /**
   * Get signing templates
   */
  public async getSigningTemplates(category?: SigningTemplateCategory): Promise<ISigningTemplate[]> {
    try {
      let query = this.sp.web.lists
        .getByTitle(this.TEMPLATES_LIST)
        .items.select(
          'Id', 'Title', 'Description', 'Category', 'Tags', 'WorkflowType',
          'DefaultSignersJSON', 'SigningBlocksJSON', 'DefaultDueDays', 'DefaultExpirationDays',
          'ReminderEnabled', 'ReminderDays', 'EscalationEnabled', 'EscalationDays',
          'EscalationAction', 'RequireComments', 'AllowDelegation', 'AllowDecline',
          'EmailSubject', 'EmailMessage', 'PreferredProvider', 'IsActive', 'UsageCount',
          'ProcessTypes', 'Created', 'Modified', 'Author/Id', 'Author/Title'
        )
        .expand('Author')
        .filter('IsActive eq true')
        .orderBy('Title', true);

      if (category) {
        query = query.filter(`Category eq '${category}'`);
      }

      const items = await query();

      return items.map(item => this.mapTemplateFromSP(item));
    } catch (error) {
      logger.error('SigningService', 'Failed to get signing templates:', error);
      throw error;
    }
  }

  /**
   * Get template by ID
   */
  public async getSigningTemplateById(id: number): Promise<ISigningTemplate> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.TEMPLATES_LIST)
        .items.getById(id)
        .select(
          'Id', 'Title', 'Description', 'Category', 'Tags', 'WorkflowType',
          'DefaultSignersJSON', 'SigningBlocksJSON', 'DefaultDueDays', 'DefaultExpirationDays',
          'ReminderEnabled', 'ReminderDays', 'EscalationEnabled', 'EscalationDays',
          'EscalationAction', 'RequireComments', 'AllowDelegation', 'AllowDecline',
          'EmailSubject', 'EmailMessage', 'PreferredProvider', 'IsActive', 'UsageCount',
          'ProcessTypes', 'Created', 'Modified', 'Author/Id', 'Author/Title'
        )
        .expand('Author')();

      return this.mapTemplateFromSP(item);
    } catch (error) {
      logger.error('SigningService', `Failed to get template ${id}:`, error);
      throw error;
    }
  }

  /**
   * Create signing template
   */
  public async createSigningTemplate(template: Partial<ISigningTemplate>): Promise<ISigningTemplate> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.TEMPLATES_LIST)
        .items.add({
          Title: template.Title,
          Description: template.Description || '',
          Category: template.Category || SigningTemplateCategory.Custom,
          Tags: template.Tags ? JSON.stringify(template.Tags) : null,
          WorkflowType: template.WorkflowType || SigningWorkflowType.Sequential,
          DefaultSignersJSON: template.DefaultSigners ? JSON.stringify(template.DefaultSigners) : null,
          SigningBlocksJSON: template.SigningBlocks ? JSON.stringify(template.SigningBlocks) : null,
          DefaultDueDays: template.DefaultDueDays || 7,
          DefaultExpirationDays: template.DefaultExpirationDays || 30,
          ReminderEnabled: template.ReminderEnabled !== false,
          ReminderDays: template.ReminderDays || 3,
          EscalationEnabled: template.EscalationEnabled || false,
          EscalationDays: template.EscalationDays || 7,
          EscalationAction: template.EscalationAction || SigningEscalationAction.Notify,
          RequireComments: template.RequireComments || false,
          AllowDelegation: template.AllowDelegation !== false,
          AllowDecline: template.AllowDecline !== false,
          EmailSubject: template.EmailSubject || '',
          EmailMessage: template.EmailMessage || '',
          PreferredProvider: template.PreferredProvider || SignatureProvider.Internal,
          IsActive: true,
          UsageCount: 0,
          ProcessTypes: template.ProcessTypes ? JSON.stringify(template.ProcessTypes) : null
        });

      return await this.getSigningTemplateById(result.data.Id);
    } catch (error) {
      logger.error('SigningService', 'Failed to create template:', error);
      throw error;
    }
  }

  /**
   * Update signing template
   */
  public async updateSigningTemplate(id: number, updates: Partial<ISigningTemplate>): Promise<ISigningTemplate> {
    try {
      const updateData: any = {};

      if (updates.Title) updateData.Title = updates.Title;
      if (updates.Description !== undefined) updateData.Description = updates.Description;
      if (updates.Category) updateData.Category = updates.Category;
      if (updates.Tags) updateData.Tags = JSON.stringify(updates.Tags);
      if (updates.WorkflowType) updateData.WorkflowType = updates.WorkflowType;
      if (updates.DefaultSigners) updateData.DefaultSignersJSON = JSON.stringify(updates.DefaultSigners);
      if (updates.SigningBlocks) updateData.SigningBlocksJSON = JSON.stringify(updates.SigningBlocks);
      if (updates.DefaultDueDays !== undefined) updateData.DefaultDueDays = updates.DefaultDueDays;
      if (updates.DefaultExpirationDays !== undefined) updateData.DefaultExpirationDays = updates.DefaultExpirationDays;
      if (updates.ReminderEnabled !== undefined) updateData.ReminderEnabled = updates.ReminderEnabled;
      if (updates.ReminderDays !== undefined) updateData.ReminderDays = updates.ReminderDays;
      if (updates.EscalationEnabled !== undefined) updateData.EscalationEnabled = updates.EscalationEnabled;
      if (updates.EscalationDays !== undefined) updateData.EscalationDays = updates.EscalationDays;
      if (updates.EscalationAction) updateData.EscalationAction = updates.EscalationAction;
      if (updates.IsActive !== undefined) updateData.IsActive = updates.IsActive;

      await this.sp.web.lists
        .getByTitle(this.TEMPLATES_LIST)
        .items.getById(id)
        .update(updateData);

      return await this.getSigningTemplateById(id);
    } catch (error) {
      logger.error('SigningService', `Failed to update template ${id}:`, error);
      throw error;
    }
  }

  /**
   * Delete signing template
   */
  public async deleteSigningTemplate(id: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.TEMPLATES_LIST)
        .items.getById(id)
        .update({ IsActive: false });

      logger.info('SigningService', `Deactivated template ${id}`);
    } catch (error) {
      logger.error('SigningService', `Failed to delete template ${id}:`, error);
      throw error;
    }
  }

  // ============================================
  // AUDIT OPERATIONS
  // ============================================

  /**
   * Get audit log for a request
   */
  public async getAuditLog(requestId: number): Promise<ISigningAuditLog[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.AUDIT_LIST)
        .items.select(
          'Id', 'Title', 'RequestId', 'RequestNumber', 'SignerId', 'SignerEmail', 'SignerName',
          'Action', 'ActionById', 'ActionBy/Id', 'ActionBy/Title', 'ActionByEmail', 'ActionByName',
          'ActionDate', 'PreviousStatus', 'NewStatus', 'Description', 'DetailsJSON',
          'IPAddress', 'UserAgent', 'GeoLocation', 'IsSystemAction', 'TriggerSource', 'Created'
        )
        .expand('ActionBy')
        .filter(`RequestId eq ${requestId}`)
        .orderBy('ActionDate', false)();

      return items.map(item => this.mapAuditLogFromSP(item));
    } catch (error) {
      logger.error('SigningService', `Failed to get audit log for request ${requestId}:`, error);
      throw error;
    }
  }

  /**
   * Log an audit entry
   */
  public async logAuditEntry(entry: Partial<ISigningAuditLog>): Promise<void> {
    try {
      const request = entry.RequestId ? await this.getBasicRequestInfo(entry.RequestId) : null;

      await this.sp.web.lists
        .getByTitle(this.AUDIT_LIST)
        .items.add({
          Title: `${entry.Action} - ${request?.RequestNumber || entry.RequestId}`,
          RequestId: entry.RequestId,
          RequestNumber: request?.RequestNumber || entry.RequestNumber,
          SignerId: entry.SignerId || null,
          SignerEmail: entry.SignerEmail || null,
          SignerName: entry.SignerName || null,
          Action: entry.Action,
          ActionById: entry.ActionById || this.currentUserId,
          ActionByEmail: entry.ActionByEmail || this.currentUserEmail,
          ActionDate: new Date().toISOString(),
          PreviousStatus: entry.PreviousStatus || null,
          NewStatus: entry.NewStatus || null,
          Description: entry.Description || null,
          DetailsJSON: entry.Details ? JSON.stringify(entry.Details) : null,
          IPAddress: entry.IPAddress || this.getClientIP(),
          UserAgent: entry.UserAgent || this.getUserAgent(),
          IsSystemAction: entry.IsSystemAction || false,
          TriggerSource: entry.TriggerSource || 'User'
        });
    } catch (error) {
      logger.error('SigningService', 'Failed to log audit entry:', error);
      // Don't throw - audit logging should not break main operations
    }
  }

  // ============================================
  // STATISTICS
  // ============================================

  /**
   * Get signing summary statistics
   */
  public async getSigningSummary(): Promise<ISigningSummary> {
    // If lists don't exist, return empty summary
    if (!this.listsExist) {
      return this.getEmptySummary();
    }

    try {
      // Get all requests
      const allRequests = await this.sp.web.lists
        .getByTitle(this.REQUESTS_LIST)
        .items.select('Id', 'Status', 'Provider', 'Category', 'DueDate', 'CompletedDate', 'Created')
        .top(5000)();

      // Get my pending signatures
      const myPending = await this.getMyPendingSignatures();

      // Calculate counts
      const now = new Date();
      const summary: ISigningSummary = {
        totalRequests: allRequests.length,
        draftRequests: allRequests.filter(r => r.Status === SigningRequestStatus.Draft).length,
        pendingRequests: allRequests.filter(r => r.Status === SigningRequestStatus.Pending).length,
        inProgressRequests: allRequests.filter(r => r.Status === SigningRequestStatus.InProgress).length,
        completedRequests: allRequests.filter(r => r.Status === SigningRequestStatus.Completed).length,
        declinedRequests: allRequests.filter(r => r.Status === SigningRequestStatus.Declined).length,
        expiredRequests: allRequests.filter(r => r.Status === SigningRequestStatus.Expired).length,
        cancelledRequests: allRequests.filter(r => r.Status === SigningRequestStatus.Cancelled).length,
        overdueRequests: allRequests.filter(r =>
          r.DueDate && new Date(r.DueDate) < now &&
          ![SigningRequestStatus.Completed, SigningRequestStatus.Cancelled, SigningRequestStatus.Expired]
            .includes(r.Status)
        ).length,

        myPendingSignatures: myPending.length,
        myCompletedSignatures: 0, // Would need separate query
        myRequestsCount: allRequests.filter(r => r.AuthorId === this.currentUserId).length,

        avgCompletionTimeHours: this.calculateAvgCompletionTime(allRequests),
        completionRate: this.calculateCompletionRate(allRequests),
        declineRate: this.calculateDeclineRate(allRequests),

        completedToday: allRequests.filter(r =>
          r.CompletedDate && this.isToday(new Date(r.CompletedDate))
        ).length,
        completedThisWeek: allRequests.filter(r =>
          r.CompletedDate && this.isThisWeek(new Date(r.CompletedDate))
        ).length,
        completedThisMonth: allRequests.filter(r =>
          r.CompletedDate && this.isThisMonth(new Date(r.CompletedDate))
        ).length,
        sentToday: allRequests.filter(r => this.isToday(new Date(r.Created))).length,
        sentThisWeek: allRequests.filter(r => this.isThisWeek(new Date(r.Created))).length,

        byStatus: this.groupByField(allRequests, 'Status'),
        byProvider: this.groupByField(allRequests, 'Provider'),
        byCategory: this.groupByField(allRequests, 'Category'),
        byDepartment: [],

        completionTrend: [],
        recentActivity: [],
        upcomingDue: [],
        overdueList: []
      };

      // Get recent activity
      summary.recentActivity = await this.getRecentAuditLog(10);

      return summary;
    } catch (error) {
      logger.error('SigningService', 'Failed to get signing summary:', error);
      throw error;
    }
  }

  // ============================================
  // PROVIDER OPERATIONS
  // ============================================

  /**
   * Get provider configurations
   */
  public async getProviderConfigs(): Promise<ISignatureProviderConfig[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.CONFIG_LIST)
        .items.filter('IsActive eq true')();

      return items.map(item => ({
        Id: item.Id,
        Title: item.Title,
        Provider: item.Provider,
        IsActive: item.IsActive,
        IsDefault: item.IsDefault,
        ApiBaseUrl: item.ApiBaseUrl,
        AccountId: item.AccountId,
        ClientId: item.ClientId,
        WebhookUrl: item.WebhookUrl,
        Settings: item.SettingsJSON ? JSON.parse(item.SettingsJSON) : {},
        LastSyncDate: item.LastSyncDate ? new Date(item.LastSyncDate) : undefined
      }));
    } catch (error) {
      logger.error('SigningService', 'Failed to get provider configs:', error);
      throw error;
    }
  }

  /**
   * Sync request status with external provider
   */
  public async syncWithProvider(requestId: number): Promise<void> {
    try {
      const request = await this.getSigningRequestById(requestId);

      if (request.Provider === SignatureProvider.Internal) {
        return; // No sync needed for internal
      }

      // Get provider config
      const configs = await this.getProviderConfigs();
      const config = configs.find(c => c.Provider === request.Provider);

      if (!config) {
        throw new Error(`No configuration found for provider ${request.Provider}`);
      }

      // Sync based on provider
      switch (request.Provider) {
        case SignatureProvider.DocuSign:
          await this.syncDocuSignStatus(request, config);
          break;
        case SignatureProvider.AdobeSign:
          await this.syncAdobeSignStatus(request, config);
          break;
        case SignatureProvider.SigningHub:
          await this.syncSigningHubStatus(request, config);
          break;
        default:
          logger.warn('SigningService', `Sync not implemented for provider ${request.Provider}`);
      }

      // Log audit entry
      await this.logAuditEntry({
        RequestId: requestId,
        Action: SigningAuditAction.SyncedWithProvider,
        Description: `Synced with ${request.Provider}`,
        IsSystemAction: true
      });
    } catch (error) {
      logger.error('SigningService', `Failed to sync with provider for request ${requestId}:`, error);
      throw error;
    }
  }

  // ============================================
  // PRIVATE HELPER METHODS
  // ============================================

  private async generateRequestNumber(): Promise<string> {
    const year = new Date().getFullYear();
    const prefix = `SIG-${year}-`;

    // Get the latest request number for this year
    const items = await this.sp.web.lists
      .getByTitle(this.REQUESTS_LIST)
      .items.filter(`substringof('${prefix}', RequestNumber)`)
      .orderBy('Id', false)
      .top(1)();

    let nextNumber = 1;
    if (items.length > 0) {
      const lastNumber = items[0].RequestNumber.replace(prefix, '');
      nextNumber = parseInt(lastNumber, 10) + 1;
    }

    return `${prefix}${nextNumber.toString().padStart(4, '0')}`;
  }

  private buildSigningChain(workflowType: SigningWorkflowType, signers: ICreateSignerConfig[]): ISigningChain {
    // Group signers by level
    const levelMap = new Map<number, ICreateSignerConfig[]>();
    signers.forEach(s => {
      if (!levelMap.has(s.level)) {
        levelMap.set(s.level, []);
      }
      levelMap.get(s.level)!.push(s);
    });

    const levels: ISigningLevel[] = [];
    levelMap.forEach((levelSigners, levelNum) => {
      levels.push({
        level: levelNum,
        signers: levelSigners.map(s => ({
          SignerEmail: s.email,
          SignerName: s.name,
          Role: s.role,
          Level: s.level,
          Order: s.order,
          Status: SignerStatus.Pending,
          SignatureType: s.signatureType || SignatureType.Electronic,
          AuthenticationMethod: s.authenticationMethod || SignerAuthenticationMethod.Email,
          RequireIdVerification: s.requireIdVerification || false,
          CanDelegate: s.canDelegate !== false,
          CanDecline: s.canDecline !== false
        } as ISigner)),
        workflowType: workflowType,
        dueDays: 7,
        status: SigningRequestStatus.Pending
      });
    });

    // Sort levels
    levels.sort((a, b) => a.level - b.level);

    return {
      WorkflowType: workflowType,
      CurrentLevel: 1,
      TotalLevels: levels.length,
      Status: SigningRequestStatus.Pending,
      Levels: levels
    };
  }

  private determineRequestType(signers: ICreateSignerConfig[]): SigningRequestType {
    if (signers.length === 1) {
      return SigningRequestType.SingleSigner;
    }

    // Check for counter-sign pattern (same person signs twice)
    const emails = signers.map(s => s.email.toLowerCase());
    const uniqueEmails = new Set(emails);
    if (uniqueEmails.size < signers.length) {
      return SigningRequestType.CounterSign;
    }

    return SigningRequestType.MultiSigner;
  }

  private async createSigningChainRecord(requestId: number, chain: ISigningChain): Promise<number> {
    const result = await this.sp.web.lists
      .getByTitle(this.CHAINS_LIST)
      .items.add({
        Title: `Chain-${requestId}`,
        RequestId: requestId,
        WorkflowType: chain.WorkflowType,
        CurrentLevel: chain.CurrentLevel,
        TotalLevels: chain.TotalLevels,
        Status: chain.Status,
        LevelsJSON: JSON.stringify(chain.Levels)
      });

    return result.data.Id;
  }

  private async createSignerRecords(requestId: number, signers: ICreateSignerConfig[]): Promise<void> {
    for (const signer of signers) {
      await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.add({
          Title: signer.name,
          RequestId: requestId,
          Level: signer.level,
          Order: signer.order,
          SignerEmail: signer.email,
          SignerName: signer.name,
          SignerPhone: signer.phone || null,
          SignerCompany: signer.company || null,
          SignerTitle: signer.title || null,
          SignerRole: signer.role,
          Status: SignerStatus.Pending,
          SignatureType: signer.signatureType || SignatureType.Electronic,
          AuthenticationMethod: signer.authenticationMethod || SignerAuthenticationMethod.Email,
          RequireIdVerification: signer.requireIdVerification || false,
          AccessCode: signer.accessCode || null,
          CanDelegate: signer.canDelegate !== false,
          CanDecline: signer.canDecline !== false,
          NotificationPreference: signer.notificationPreference || 'Email',
          AssignedBlockIds: signer.assignedBlockIds ? JSON.stringify(signer.assignedBlockIds) : null,
          MetadataJSON: signer.metadata ? JSON.stringify(signer.metadata) : null
        });
    }
  }

  private async loadSignersForRequest(requestId: number): Promise<ISigningLevel[]> {
    const signers = await this.sp.web.lists
      .getByTitle(this.SIGNERS_LIST)
      .items.filter(`RequestId eq ${requestId}`)
      .orderBy('Level', true)
      .orderBy('Order', true)();

    // Group by level
    const levelMap = new Map<number, ISigner[]>();
    signers.forEach(s => {
      const signer = this.mapSignerFromSP(s);
      if (!levelMap.has(signer.Level)) {
        levelMap.set(signer.Level, []);
      }
      levelMap.get(signer.Level)!.push(signer);
    });

    const levels: ISigningLevel[] = [];
    levelMap.forEach((levelSigners, levelNum) => {
      levels.push({
        level: levelNum,
        signers: levelSigners,
        workflowType: SigningWorkflowType.Sequential,
        dueDays: 7,
        status: this.determineLevelStatus(levelSigners)
      });
    });

    return levels.sort((a, b) => a.level - b.level);
  }

  private determineLevelStatus(signers: ISigner[]): SigningRequestStatus {
    const allSigned = signers.every(s => s.Status === SignerStatus.Signed);
    if (allSigned) return SigningRequestStatus.Completed;

    const anyDeclined = signers.some(s => s.Status === SignerStatus.Declined);
    if (anyDeclined) return SigningRequestStatus.Declined;

    const anyInProgress = signers.some(s =>
      [SignerStatus.Sent, SignerStatus.Viewed].includes(s.Status)
    );
    if (anyInProgress) return SigningRequestStatus.InProgress;

    return SigningRequestStatus.Pending;
  }

  private async getSignersForLevel(requestId: number, level: number): Promise<ISigner[]> {
    const items = await this.sp.web.lists
      .getByTitle(this.SIGNERS_LIST)
      .items.filter(`RequestId eq ${requestId} and Level eq ${level}`)
      .orderBy('Order', true)();

    return items.map(item => this.mapSignerFromSP(item));
  }

  private async getSignerById(signerId: number): Promise<ISigner> {
    const item = await this.sp.web.lists
      .getByTitle(this.SIGNERS_LIST)
      .items.getById(signerId)();

    return this.mapSignerFromSP(item);
  }

  private async updateSignersStatus(requestId: number, status: SignerStatus): Promise<void> {
    const signers = await this.sp.web.lists
      .getByTitle(this.SIGNERS_LIST)
      .items.filter(`RequestId eq ${requestId} and Status ne 'Signed'`)();

    for (const signer of signers) {
      await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.getById(signer.Id)
        .update({ Status: status });
    }
  }

  private async evaluateAndAdvanceWorkflow(requestId: number, completedLevel: number): Promise<void> {
    const request = await this.getSigningRequestById(requestId);
    const levels = request.SigningChain.Levels;

    // Check if current level is complete
    const currentLevelSigners = levels.find(l => l.level === completedLevel)?.signers || [];
    const allSigned = currentLevelSigners.every(s => s.Status === SignerStatus.Signed);

    if (!allSigned) {
      // Level not complete yet
      if (request.WorkflowType === SigningWorkflowType.FirstSigner) {
        // First signer wins - complete the request
        await this.completeRequest(requestId);
      }
      return;
    }

    // Level is complete - check if there are more levels
    const nextLevel = completedLevel + 1;
    const nextLevelExists = levels.some(l => l.level === nextLevel);

    if (nextLevelExists) {
      // Advance to next level
      await this.activateLevel(requestId, nextLevel);

      // Update chain
      await this.sp.web.lists
        .getByTitle(this.CHAINS_LIST)
        .items.filter(`RequestId eq ${requestId}`)
        .top(1)()
        .then(async (chains) => {
          if (chains.length > 0) {
            await this.sp.web.lists
              .getByTitle(this.CHAINS_LIST)
              .items.getById(chains[0].Id)
              .update({ CurrentLevel: nextLevel });
          }
        });
    } else {
      // All levels complete - complete the request
      await this.completeRequest(requestId);
    }
  }

  private async activateLevel(requestId: number, level: number): Promise<void> {
    const signers = await this.getSignersForLevel(requestId, level);

    for (const signer of signers) {
      await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.getById(signer.Id!)
        .update({
          Status: SignerStatus.Sent,
          SentDate: new Date().toISOString()
        });

      // Send notification
      await this.sendSignerNotification(requestId, signer.Id!);
    }

    await this.logAuditEntry({
      RequestId: requestId,
      Action: SigningAuditAction.Sent,
      Description: `Level ${level} activated - ${signers.length} signers notified`,
      IsSystemAction: true
    });
  }

  private async completeRequest(requestId: number): Promise<void> {
    await this.sp.web.lists
      .getByTitle(this.REQUESTS_LIST)
      .items.getById(requestId)
      .update({
        Status: SigningRequestStatus.Completed,
        CompletedDate: new Date().toISOString()
      });

    await this.logAuditEntry({
      RequestId: requestId,
      Action: SigningAuditAction.Completed,
      Description: 'All signatures collected - request completed',
      IsSystemAction: true
    });

    // Generate certificate of completion
    // await this.generateCertificate(requestId);

    // Send completion notifications
    // Notification logic would go here
  }

  private async sendSignerNotification(requestId: number, signerId: number, message?: string): Promise<void> {
    // This would integrate with the notification service
    logger.info('SigningService', `Sending notification to signer ${signerId} for request ${requestId}`);
    // Implementation would call NotificationService
  }

  private async sendToExternalProvider(request: ISigningRequest, signers: ISigner[]): Promise<void> {
    // Get provider config
    const configs = await this.getProviderConfigs();
    const config = configs.find(c => c.Provider === request.Provider);

    if (!config) {
      throw new Error(`No configuration found for provider ${request.Provider}`);
    }

    switch (request.Provider) {
      case SignatureProvider.DocuSign:
        await this.sendToDocuSign(request, signers, config);
        break;
      case SignatureProvider.AdobeSign:
        await this.sendToAdobeSign(request, signers, config);
        break;
      case SignatureProvider.SigningHub:
        await this.sendToSigningHub(request, signers, config);
        break;
      default:
        throw new Error(`Provider ${request.Provider} not implemented`);
    }
  }

  private async sendToDocuSign(request: ISigningRequest, signers: ISigner[], config: ISignatureProviderConfig): Promise<void> {
    // DocuSign integration logic
    logger.info('SigningService', `Sending to DocuSign: ${request.RequestNumber}`);
    // Implementation would call DocuSign API
  }

  private async sendToAdobeSign(request: ISigningRequest, signers: ISigner[], config: ISignatureProviderConfig): Promise<void> {
    // Adobe Sign integration logic
    logger.info('SigningService', `Sending to Adobe Sign: ${request.RequestNumber}`);
    // Implementation would call Adobe Sign API
  }

  private async sendToSigningHub(request: ISigningRequest, signers: ISigner[], config: ISignatureProviderConfig): Promise<void> {
    // Signing Hub integration logic
    logger.info('SigningService', `Sending to Signing Hub: ${request.RequestNumber}`);
    // Implementation would call Signing Hub API
  }

  private async voidExternalEnvelope(request: ISigningRequest): Promise<void> {
    // Void envelope with external provider
    logger.info('SigningService', `Voiding external envelope: ${request.ExternalEnvelopeId}`);
    // Implementation would call provider API
  }

  private async syncDocuSignStatus(request: ISigningRequest, config: ISignatureProviderConfig): Promise<void> {
    // Sync status from DocuSign
    logger.info('SigningService', `Syncing DocuSign status for: ${request.RequestNumber}`);
  }

  private async syncAdobeSignStatus(request: ISigningRequest, config: ISignatureProviderConfig): Promise<void> {
    // Sync status from Adobe Sign
    logger.info('SigningService', `Syncing Adobe Sign status for: ${request.RequestNumber}`);
  }

  private async syncSigningHubStatus(request: ISigningRequest, config: ISignatureProviderConfig): Promise<void> {
    // Sync status from Signing Hub
    logger.info('SigningService', `Syncing Signing Hub status for: ${request.RequestNumber}`);
  }

  private async getBasicRequestInfo(requestId: number): Promise<{ RequestNumber: string } | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.REQUESTS_LIST)
        .items.getById(requestId)
        .select('RequestNumber')();
      return item;
    } catch {
      return null;
    }
  }

  private async getRecentAuditLog(count: number): Promise<ISigningAuditLog[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.AUDIT_LIST)
        .items.orderBy('ActionDate', false)
        .top(count)();

      return items.map(item => this.mapAuditLogFromSP(item));
    } catch {
      return [];
    }
  }

  // ============================================
  // MAPPING METHODS
  // ============================================

  private mapRequestFromSP(item: any): ISigningRequest {
    return {
      Id: item.Id,
      Title: item.Title,
      RequestNumber: item.RequestNumber,
      Description: item.Description,
      Status: item.Status as SigningRequestStatus,
      RequestType: item.RequestType as SigningRequestType,
      WorkflowType: item.WorkflowType as SigningWorkflowType,
      Priority: item.Priority,
      RequesterId: item.RequesterId,
      Requester: item.Requester ? {
        Id: item.Requester.Id,
        Title: item.Requester.Title,
        EMail: item.Requester.EMail
      } : undefined,
      DocumentIds: item.DocumentIds ? JSON.parse(item.DocumentIds) : [],
      ProcessId: item.ProcessId,
      ProcessType: item.ProcessType,
      TemplateId: item.TemplateId,
      Provider: item.Provider as SignatureProvider,
      ExternalEnvelopeId: item.ExternalEnvelopeId,
      SigningChain: item.SigningChainJSON ? JSON.parse(item.SigningChainJSON) : {
        WorkflowType: item.WorkflowType,
        CurrentLevel: 1,
        TotalLevels: 1,
        Status: item.Status,
        Levels: []
      },
      SigningBlocks: item.SigningBlocksJSON ? JSON.parse(item.SigningBlocksJSON) : [],
      DueDate: item.DueDate ? new Date(item.DueDate) : undefined,
      ExpirationDate: item.ExpirationDate ? new Date(item.ExpirationDate) : undefined,
      CompletedDate: item.CompletedDate ? new Date(item.CompletedDate) : undefined,
      SentDate: item.SentDate ? new Date(item.SentDate) : undefined,
      ReminderEnabled: item.ReminderEnabled,
      ReminderDays: item.ReminderDays,
      EscalationEnabled: item.EscalationEnabled,
      EscalationDays: item.EscalationDays,
      EscalationAction: item.EscalationAction as SigningEscalationAction,
      AllowDelegation: item.AllowDelegation,
      AllowDecline: item.AllowDecline,
      RequireComments: item.RequireComments,
      RequireReason: item.RequireReason,
      AllowReassignment: item.AllowReassignment,
      RequireAccessCode: item.RequireAccessCode,
      EmailSubject: item.EmailSubject,
      EmailMessage: item.EmailMessage,
      CertificateUrl: item.CertificateUrl,
      Tags: item.Tags ? JSON.parse(item.Tags) : [],
      Category: item.Category as SigningTemplateCategory,
      Metadata: item.MetadataJSON ? JSON.parse(item.MetadataJSON) : {},
      Created: item.Created ? new Date(item.Created) : undefined,
      CreatedBy: item.Author ? {
        Id: item.Author.Id,
        Title: item.Author.Title
      } : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined,
      ModifiedBy: item.Editor ? {
        Id: item.Editor.Id,
        Title: item.Editor.Title
      } : undefined
    };
  }

  private mapSignerFromSP(item: any): ISigner {
    return {
      Id: item.Id,
      RequestId: item.RequestId,
      ChainId: item.ChainId,
      SignerUserId: item.SignerUserId,
      SignerEmail: item.SignerEmail,
      SignerName: item.SignerName,
      SignerPhone: item.SignerPhone,
      SignerCompany: item.SignerCompany,
      SignerTitle: item.SignerTitle,
      Role: item.SignerRole as SignerRole,
      Level: item.Level,
      Order: item.Order,
      Status: item.Status as SignerStatus,
      SignatureType: item.SignatureType as SignatureType,
      AuthenticationMethod: item.AuthenticationMethod as SignerAuthenticationMethod,
      RequireIdVerification: item.RequireIdVerification,
      AccessCode: item.AccessCode,
      CanDelegate: item.CanDelegate,
      CanDecline: item.CanDecline,
      CanAddComments: item.CanAddComments,
      DelegatedToEmail: item.DelegatedToEmail,
      DelegatedToName: item.DelegatedToName,
      DelegatedById: item.DelegatedById,
      DelegationReason: item.DelegationReason,
      DelegationDate: item.DelegationDate ? new Date(item.DelegationDate) : undefined,
      SentDate: item.SentDate ? new Date(item.SentDate) : undefined,
      ViewedDate: item.ViewedDate ? new Date(item.ViewedDate) : undefined,
      SignedDate: item.SignedDate ? new Date(item.SignedDate) : undefined,
      DeclinedDate: item.DeclinedDate ? new Date(item.DeclinedDate) : undefined,
      DeclineReason: item.DeclineReason,
      Comments: item.Comments,
      SignatureData: item.SignatureDataJSON ? JSON.parse(item.SignatureDataJSON) : undefined,
      IPAddress: item.IPAddress,
      UserAgent: item.UserAgent,
      AssignedBlockIds: item.AssignedBlockIds ? JSON.parse(item.AssignedBlockIds) : [],
      CompletedBlockIds: item.CompletedBlockIds ? JSON.parse(item.CompletedBlockIds) : [],
      RemindersSent: item.RemindersSent,
      LastReminderDate: item.LastReminderDate ? new Date(item.LastReminderDate) : undefined,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }

  private mapTemplateFromSP(item: any): ISigningTemplate {
    return {
      Id: item.Id,
      Title: item.Title,
      Description: item.Description,
      Category: item.Category as SigningTemplateCategory,
      Tags: item.Tags ? JSON.parse(item.Tags) : [],
      WorkflowType: item.WorkflowType as SigningWorkflowType,
      DefaultSigners: item.DefaultSignersJSON ? JSON.parse(item.DefaultSignersJSON) : [],
      SigningBlocks: item.SigningBlocksJSON ? JSON.parse(item.SigningBlocksJSON) : [],
      DefaultDueDays: item.DefaultDueDays,
      DefaultExpirationDays: item.DefaultExpirationDays,
      ReminderEnabled: item.ReminderEnabled,
      ReminderDays: item.ReminderDays,
      EscalationEnabled: item.EscalationEnabled,
      EscalationDays: item.EscalationDays,
      EscalationAction: item.EscalationAction as SigningEscalationAction,
      RequireComments: item.RequireComments,
      AllowDelegation: item.AllowDelegation,
      AllowDecline: item.AllowDecline,
      EmailSubject: item.EmailSubject,
      EmailMessage: item.EmailMessage,
      PreferredProvider: item.PreferredProvider as SignatureProvider,
      IsActive: item.IsActive,
      UsageCount: item.UsageCount,
      ProcessTypes: item.ProcessTypes ? JSON.parse(item.ProcessTypes) : [],
      Created: item.Created ? new Date(item.Created) : undefined,
      CreatedBy: item.Author ? {
        Id: item.Author.Id,
        Title: item.Author.Title
      } : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }

  private mapAuditLogFromSP(item: any): ISigningAuditLog {
    return {
      Id: item.Id,
      RequestId: item.RequestId,
      RequestNumber: item.RequestNumber,
      SignerId: item.SignerId,
      SignerEmail: item.SignerEmail,
      SignerName: item.SignerName,
      Action: item.Action as SigningAuditAction,
      ActionById: item.ActionById,
      ActionBy: item.ActionBy ? {
        Id: item.ActionBy.Id,
        Title: item.ActionBy.Title
      } : undefined,
      ActionByEmail: item.ActionByEmail,
      ActionByName: item.ActionByName,
      ActionDate: new Date(item.ActionDate),
      PreviousStatus: item.PreviousStatus,
      NewStatus: item.NewStatus,
      Description: item.Description,
      Details: item.DetailsJSON ? JSON.parse(item.DetailsJSON) : {},
      IPAddress: item.IPAddress,
      UserAgent: item.UserAgent,
      GeoLocation: item.GeoLocation,
      SessionId: item.SessionId,
      IsSystemAction: item.IsSystemAction,
      TriggerSource: item.TriggerSource,
      Created: item.Created ? new Date(item.Created) : undefined
    };
  }

  // ============================================
  // UTILITY METHODS
  // ============================================

  private addDays(date: Date, days: number): Date {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  }

  private isToday(date: Date): boolean {
    const today = new Date();
    return date.toDateString() === today.toDateString();
  }

  private isThisWeek(date: Date): boolean {
    const now = new Date();
    const weekStart = new Date(now.setDate(now.getDate() - now.getDay()));
    const weekEnd = new Date(now.setDate(now.getDate() - now.getDay() + 6));
    return date >= weekStart && date <= weekEnd;
  }

  private isThisMonth(date: Date): boolean {
    const now = new Date();
    return date.getMonth() === now.getMonth() && date.getFullYear() === now.getFullYear();
  }

  private calculateAvgCompletionTime(requests: any[]): number {
    const completed = requests.filter(r =>
      r.Status === SigningRequestStatus.Completed && r.CompletedDate && r.Created
    );

    if (completed.length === 0) return 0;

    const totalHours = completed.reduce((sum, r) => {
      const created = new Date(r.Created).getTime();
      const completedDate = new Date(r.CompletedDate).getTime();
      return sum + (completedDate - created) / (1000 * 60 * 60);
    }, 0);

    return Math.round(totalHours / completed.length);
  }

  private calculateCompletionRate(requests: any[]): number {
    const total = requests.filter(r => r.Status !== SigningRequestStatus.Draft).length;
    if (total === 0) return 0;

    const completed = requests.filter(r => r.Status === SigningRequestStatus.Completed).length;
    return Math.round((completed / total) * 100);
  }

  private calculateDeclineRate(requests: any[]): number {
    const total = requests.filter(r => r.Status !== SigningRequestStatus.Draft).length;
    if (total === 0) return 0;

    const declined = requests.filter(r => r.Status === SigningRequestStatus.Declined).length;
    return Math.round((declined / total) * 100);
  }

  private groupByField(items: any[], field: string): any[] {
    const grouped = new Map<string, number>();

    items.forEach(item => {
      const value = item[field] || 'Unknown';
      grouped.set(value, (grouped.get(value) || 0) + 1);
    });

    return Array.from(grouped.entries()).map(([key, count]) => ({
      [field.toLowerCase()]: key,
      count
    }));
  }

  private getClientIP(): string {
    // In a real implementation, this would get the client IP
    // For SPFx, this is typically not directly available
    return 'Unknown';
  }

  private getUserAgent(): string {
    return typeof navigator !== 'undefined' ? navigator.userAgent : 'Unknown';
  }

  /**
   * Returns an empty summary for when lists are not provisioned
   */
  private getEmptySummary(): ISigningSummary {
    return {
      totalRequests: 0,
      draftRequests: 0,
      pendingRequests: 0,
      inProgressRequests: 0,
      completedRequests: 0,
      declinedRequests: 0,
      expiredRequests: 0,
      cancelledRequests: 0,
      overdueRequests: 0,
      myPendingSignatures: 0,
      myCompletedSignatures: 0,
      myRequestsCount: 0,
      avgCompletionTimeHours: 0,
      completionRate: 0,
      declineRate: 0,
      completedToday: 0,
      completedThisWeek: 0,
      completedThisMonth: 0,
      sentToday: 0,
      sentThisWeek: 0,
      byStatus: [],
      byProvider: [],
      byCategory: [],
      byDepartment: [],
      completionTrend: [],
      recentActivity: [],
      upcomingDue: [],
      overdueList: []
    };
  }
}

export default SigningService;
