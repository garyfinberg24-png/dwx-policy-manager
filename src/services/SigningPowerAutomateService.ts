// @ts-nocheck
/* eslint-disable @typescript-eslint/no-explicit-any */
// TODO: Fix flatMap (needs es2019+) and iterator issues
// Signing Power Automate Service
// Handles Power Automate integration, webhooks, and HTTP triggers
// Note: Some fields may not exist in the SharePoint list - mapping handles this gracefully

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import {
  ISigningRequest,
  ISigner,
  ISigningWebhookPayload,
  SigningRequestStatus,
  SignerStatus,
  SigningAuditAction,
  SignatureProvider
} from '../models/ISigning';
import { SigningService } from './SigningService';
import { logger } from './LoggingService';

/**
 * Webhook Configuration
 */
export interface IWebhookConfig {
  id: string;
  name: string;
  url: string;
  secret?: string;
  events: SigningAuditAction[];
  isActive: boolean;
  headers?: Record<string, string>;
  retryCount?: number;
  timeoutSeconds?: number;
}

/**
 * Power Automate Trigger Response
 */
export interface ITriggerResponse {
  success: boolean;
  triggerId?: string;
  error?: string;
  timestamp: Date;
}

/**
 * Power Automate Action Request
 */
export interface IActionRequest {
  action: string;
  requestId?: number;
  signerId?: number;
  parameters?: Record<string, any>;
  correlationId?: string;
}

/**
 * Power Automate Action Response
 */
export interface IActionResponse {
  success: boolean;
  result?: any;
  error?: string;
  correlationId?: string;
}

/**
 * Signing Power Automate Service
 */
export class SigningPowerAutomateService {
  private sp: SPFI;
  private signingService: SigningService;
  private webhookConfigs: Map<string, IWebhookConfig> = new Map();

  private readonly WEBHOOK_CONFIG_LIST = 'JML_SigningWebhooks';
  private readonly WEBHOOK_LOG_LIST = 'JML_SigningWebhookLog';

  constructor(sp: SPFI) {
    this.sp = sp;
    this.signingService = new SigningService(sp);
  }

  /**
   * Initialize the service and load webhook configurations
   */
  public async initialize(): Promise<void> {
    try {
      await this.signingService.initialize();
      await this.loadWebhookConfigs();
      logger.info('SigningPowerAutomateService', 'Initialized');
    } catch (error) {
      logger.error('SigningPowerAutomateService', 'Failed to initialize:', error);
    }
  }

  /**
   * Load webhook configurations from SharePoint
   */
  private async loadWebhookConfigs(): Promise<void> {
    try {
      const configs = await this.sp.web.lists
        .getByTitle(this.WEBHOOK_CONFIG_LIST)
        .items.filter('IsActive eq true')();

      this.webhookConfigs.clear();
      for (const config of configs) {
        this.webhookConfigs.set(config.Id.toString(), {
          id: config.Id.toString(),
          name: config.Title,
          url: config.WebhookUrl,
          secret: config.WebhookSecret,
          events: config.Events ? JSON.parse(config.Events) : [],
          isActive: config.IsActive,
          headers: config.Headers ? JSON.parse(config.Headers) : {},
          retryCount: config.RetryCount || 3,
          timeoutSeconds: config.TimeoutSeconds || 30
        });
      }

      logger.info('SigningPowerAutomateService', `Loaded ${this.webhookConfigs.size} webhook configurations`);
    } catch (error) {
      logger.warn('SigningPowerAutomateService', 'Failed to load webhook configs - list may not exist:', error);
    }
  }

  // ============================================
  // HTTP TRIGGERS FOR POWER AUTOMATE
  // ============================================

  /**
   * Trigger: When a signing request is created
   */
  public async triggerOnRequestCreated(request: ISigningRequest): Promise<ITriggerResponse[]> {
    return this.sendWebhooks(SigningAuditAction.Created, {
      eventId: this.generateEventId(),
      eventType: SigningAuditAction.Created,
      timestamp: new Date(),
      request: this.buildRequestPayload(request)
    });
  }

  /**
   * Trigger: When a signing request is sent
   */
  public async triggerOnRequestSent(request: ISigningRequest): Promise<ITriggerResponse[]> {
    return this.sendWebhooks(SigningAuditAction.Sent, {
      eventId: this.generateEventId(),
      eventType: SigningAuditAction.Sent,
      timestamp: new Date(),
      request: this.buildRequestPayload(request)
    });
  }

  /**
   * Trigger: When a signature is completed
   */
  public async triggerOnSignatureCompleted(request: ISigningRequest, signer: ISigner): Promise<ITriggerResponse[]> {
    return this.sendWebhooks(SigningAuditAction.Signed, {
      eventId: this.generateEventId(),
      eventType: SigningAuditAction.Signed,
      timestamp: new Date(),
      request: this.buildRequestPayload(request),
      signer: this.buildSignerPayload(signer)
    });
  }

  /**
   * Trigger: When a signing request is completed (all signatures collected)
   */
  public async triggerOnRequestCompleted(request: ISigningRequest): Promise<ITriggerResponse[]> {
    return this.sendWebhooks(SigningAuditAction.Completed, {
      eventId: this.generateEventId(),
      eventType: SigningAuditAction.Completed,
      timestamp: new Date(),
      request: this.buildRequestPayload(request)
    });
  }

  /**
   * Trigger: When a signing request is declined
   */
  public async triggerOnRequestDeclined(request: ISigningRequest, signer: ISigner): Promise<ITriggerResponse[]> {
    return this.sendWebhooks(SigningAuditAction.Declined, {
      eventId: this.generateEventId(),
      eventType: SigningAuditAction.Declined,
      timestamp: new Date(),
      request: this.buildRequestPayload(request),
      signer: this.buildSignerPayload(signer),
      details: {
        declineReason: signer.DeclineReason
      }
    });
  }

  /**
   * Trigger: When a signing request expires
   */
  public async triggerOnRequestExpired(request: ISigningRequest): Promise<ITriggerResponse[]> {
    return this.sendWebhooks(SigningAuditAction.Expired, {
      eventId: this.generateEventId(),
      eventType: SigningAuditAction.Expired,
      timestamp: new Date(),
      request: this.buildRequestPayload(request)
    });
  }

  /**
   * Trigger: When a signing request is cancelled/voided
   */
  public async triggerOnRequestCancelled(request: ISigningRequest, reason: string): Promise<ITriggerResponse[]> {
    return this.sendWebhooks(SigningAuditAction.Cancelled, {
      eventId: this.generateEventId(),
      eventType: SigningAuditAction.Cancelled,
      timestamp: new Date(),
      request: this.buildRequestPayload(request),
      details: {
        cancelReason: reason
      }
    });
  }

  /**
   * Trigger: When a signing request is escalated
   */
  public async triggerOnRequestEscalated(request: ISigningRequest): Promise<ITriggerResponse[]> {
    return this.sendWebhooks(SigningAuditAction.Escalated, {
      eventId: this.generateEventId(),
      eventType: SigningAuditAction.Escalated,
      timestamp: new Date(),
      request: this.buildRequestPayload(request)
    });
  }

  /**
   * Trigger: When a signature is delegated
   */
  public async triggerOnSignatureDelegated(
    request: ISigningRequest,
    originalSigner: ISigner,
    delegateEmail: string,
    delegateName: string
  ): Promise<ITriggerResponse[]> {
    return this.sendWebhooks(SigningAuditAction.Delegated, {
      eventId: this.generateEventId(),
      eventType: SigningAuditAction.Delegated,
      timestamp: new Date(),
      request: this.buildRequestPayload(request),
      signer: this.buildSignerPayload(originalSigner),
      details: {
        delegateEmail,
        delegateName
      }
    });
  }

  /**
   * Trigger: When a document is viewed
   */
  public async triggerOnDocumentViewed(request: ISigningRequest, signer: ISigner): Promise<ITriggerResponse[]> {
    return this.sendWebhooks(SigningAuditAction.Viewed, {
      eventId: this.generateEventId(),
      eventType: SigningAuditAction.Viewed,
      timestamp: new Date(),
      request: this.buildRequestPayload(request),
      signer: this.buildSignerPayload(signer)
    });
  }

  // ============================================
  // INCOMING ACTIONS FROM POWER AUTOMATE
  // ============================================

  /**
   * Handle incoming action from Power Automate
   */
  public async handleAction(actionRequest: IActionRequest): Promise<IActionResponse> {
    try {
      logger.info('SigningPowerAutomateService', `Handling action: ${actionRequest.action}`, actionRequest);

      switch (actionRequest.action.toLowerCase()) {
        case 'createsigningrequest':
          return await this.handleCreateSigningRequest(actionRequest);

        case 'getrequeststatus':
          return await this.handleGetRequestStatus(actionRequest);

        case 'sendreminder':
          return await this.handleSendReminder(actionRequest);

        case 'recallrequest':
          return await this.handleRecallRequest(actionRequest);

        case 'addsigner':
          return await this.handleAddSigner(actionRequest);

        case 'updateduedate':
          return await this.handleUpdateDueDate(actionRequest);

        case 'getsigners':
          return await this.handleGetSigners(actionRequest);

        case 'voidrequest':
          return await this.handleVoidRequest(actionRequest);

        case 'resendtosigner':
          return await this.handleResendToSigner(actionRequest);

        case 'delegatesignature':
          return await this.handleDelegateSignature(actionRequest);

        case 'getauditlog':
          return await this.handleGetAuditLog(actionRequest);

        default:
          return {
            success: false,
            error: `Unknown action: ${actionRequest.action}`,
            correlationId: actionRequest.correlationId
          };
      }
    } catch (error) {
      logger.error('SigningPowerAutomateService', `Action failed: ${actionRequest.action}`, error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error',
        correlationId: actionRequest.correlationId
      };
    }
  }

  /**
   * Handle: Create Signing Request
   */
  private async handleCreateSigningRequest(actionRequest: IActionRequest): Promise<IActionResponse> {
    const params = actionRequest.parameters;

    if (!params?.title || !params?.documentIds || !params?.signers) {
      return {
        success: false,
        error: 'Missing required parameters: title, documentIds, signers',
        correlationId: actionRequest.correlationId
      };
    }

    const request = await this.signingService.createSigningRequest({
      title: params.title,
      description: params.description,
      documentIds: params.documentIds,
      workflowType: params.workflowType || 'Sequential',
      signers: params.signers,
      provider: params.provider || SignatureProvider.Internal,
      processId: params.processId,
      processType: params.processType,
      dueDate: params.dueDate ? new Date(params.dueDate) : undefined,
      expirationDays: params.expirationDays,
      reminderEnabled: params.reminderEnabled,
      reminderDays: params.reminderDays,
      emailSubject: params.emailSubject,
      emailMessage: params.emailMessage,
      allowDelegation: params.allowDelegation,
      allowDecline: params.allowDecline,
      tags: params.tags,
      sendImmediately: params.sendImmediately
    });

    return {
      success: true,
      result: {
        requestId: request.Id,
        requestNumber: request.RequestNumber,
        status: request.Status
      },
      correlationId: actionRequest.correlationId
    };
  }

  /**
   * Handle: Get Request Status
   */
  private async handleGetRequestStatus(actionRequest: IActionRequest): Promise<IActionResponse> {
    if (!actionRequest.requestId) {
      return {
        success: false,
        error: 'Missing required parameter: requestId',
        correlationId: actionRequest.correlationId
      };
    }

    const request = await this.signingService.getSigningRequestById(actionRequest.requestId);

    return {
      success: true,
      result: {
        requestId: request.Id,
        requestNumber: request.RequestNumber,
        status: request.Status,
        currentLevel: request.SigningChain.CurrentLevel,
        totalLevels: request.SigningChain.TotalLevels,
        signers: request.SigningChain.Levels.flatMap(l => l.signers.map(s => ({
          id: s.Id,
          name: s.SignerName,
          email: s.SignerEmail,
          status: s.Status,
          level: s.Level,
          signedDate: s.SignedDate
        }))),
        createdDate: request.Created,
        sentDate: request.SentDate,
        completedDate: request.CompletedDate,
        dueDate: request.DueDate
      },
      correlationId: actionRequest.correlationId
    };
  }

  /**
   * Handle: Send Reminder
   */
  private async handleSendReminder(actionRequest: IActionRequest): Promise<IActionResponse> {
    if (!actionRequest.requestId) {
      return {
        success: false,
        error: 'Missing required parameter: requestId',
        correlationId: actionRequest.correlationId
      };
    }

    await this.signingService.resendToSigner({
      requestId: actionRequest.requestId,
      signerId: actionRequest.signerId,
      message: actionRequest.parameters?.message
    });

    return {
      success: true,
      result: { message: 'Reminder sent successfully' },
      correlationId: actionRequest.correlationId
    };
  }

  /**
   * Handle: Recall Request
   */
  private async handleRecallRequest(actionRequest: IActionRequest): Promise<IActionResponse> {
    if (!actionRequest.requestId || !actionRequest.parameters?.reason) {
      return {
        success: false,
        error: 'Missing required parameters: requestId, reason',
        correlationId: actionRequest.correlationId
      };
    }

    await this.signingService.cancelSigningRequest({
      requestId: actionRequest.requestId,
      reason: actionRequest.parameters.reason,
      notifySigners: actionRequest.parameters.notifySigners !== false
    });

    return {
      success: true,
      result: { message: 'Request recalled successfully' },
      correlationId: actionRequest.correlationId
    };
  }

  /**
   * Handle: Add Signer
   */
  private async handleAddSigner(actionRequest: IActionRequest): Promise<IActionResponse> {
    const params = actionRequest.parameters;

    if (!actionRequest.requestId || !params?.email || !params?.name) {
      return {
        success: false,
        error: 'Missing required parameters: requestId, email, name',
        correlationId: actionRequest.correlationId
      };
    }

    // Add signer to request
    await this.sp.web.lists
      .getByTitle('JML_Signers')
      .items.add({
        Title: params.name,
        RequestId: actionRequest.requestId,
        SignerEmail: params.email,
        SignerName: params.name,
        SignerRole: params.role || 'Signer',
        Level: params.level || 1,
        Order: params.order || 1,
        Status: SignerStatus.Pending,
        SignatureType: params.signatureType || 'Electronic'
      });

    return {
      success: true,
      result: { message: 'Signer added successfully' },
      correlationId: actionRequest.correlationId
    };
  }

  /**
   * Handle: Update Due Date
   */
  private async handleUpdateDueDate(actionRequest: IActionRequest): Promise<IActionResponse> {
    if (!actionRequest.requestId || !actionRequest.parameters?.dueDate) {
      return {
        success: false,
        error: 'Missing required parameters: requestId, dueDate',
        correlationId: actionRequest.correlationId
      };
    }

    await this.signingService.updateSigningRequest(actionRequest.requestId, {
      requestId: actionRequest.requestId,
      dueDate: new Date(actionRequest.parameters.dueDate)
    });

    return {
      success: true,
      result: { message: 'Due date updated successfully' },
      correlationId: actionRequest.correlationId
    };
  }

  /**
   * Handle: Get Signers
   */
  private async handleGetSigners(actionRequest: IActionRequest): Promise<IActionResponse> {
    if (!actionRequest.requestId) {
      return {
        success: false,
        error: 'Missing required parameter: requestId',
        correlationId: actionRequest.correlationId
      };
    }

    const request = await this.signingService.getSigningRequestById(actionRequest.requestId);
    const signers = request.SigningChain.Levels.flatMap(l => l.signers);

    return {
      success: true,
      result: {
        requestId: actionRequest.requestId,
        signers: signers.map(s => ({
          id: s.Id,
          name: s.SignerName,
          email: s.SignerEmail,
          role: s.Role,
          status: s.Status,
          level: s.Level,
          order: s.Order,
          sentDate: s.SentDate,
          viewedDate: s.ViewedDate,
          signedDate: s.SignedDate,
          declinedDate: s.DeclinedDate,
          declineReason: s.DeclineReason
        }))
      },
      correlationId: actionRequest.correlationId
    };
  }

  /**
   * Handle: Void Request
   */
  private async handleVoidRequest(actionRequest: IActionRequest): Promise<IActionResponse> {
    if (!actionRequest.requestId) {
      return {
        success: false,
        error: 'Missing required parameter: requestId',
        correlationId: actionRequest.correlationId
      };
    }

    await this.signingService.cancelSigningRequest({
      requestId: actionRequest.requestId,
      reason: actionRequest.parameters?.reason || 'Voided via Power Automate',
      notifySigners: actionRequest.parameters?.notifySigners !== false
    });

    return {
      success: true,
      result: { message: 'Request voided successfully' },
      correlationId: actionRequest.correlationId
    };
  }

  /**
   * Handle: Resend to Signer
   */
  private async handleResendToSigner(actionRequest: IActionRequest): Promise<IActionResponse> {
    if (!actionRequest.requestId || !actionRequest.signerId) {
      return {
        success: false,
        error: 'Missing required parameters: requestId, signerId',
        correlationId: actionRequest.correlationId
      };
    }

    await this.signingService.resendToSigner({
      requestId: actionRequest.requestId,
      signerId: actionRequest.signerId,
      message: actionRequest.parameters?.message
    });

    return {
      success: true,
      result: { message: 'Request resent successfully' },
      correlationId: actionRequest.correlationId
    };
  }

  /**
   * Handle: Delegate Signature
   */
  private async handleDelegateSignature(actionRequest: IActionRequest): Promise<IActionResponse> {
    const params = actionRequest.parameters;

    if (!actionRequest.requestId || !actionRequest.signerId || !params?.delegateEmail || !params?.delegateName) {
      return {
        success: false,
        error: 'Missing required parameters: requestId, signerId, delegateEmail, delegateName',
        correlationId: actionRequest.correlationId
      };
    }

    await this.signingService.delegateSignature({
      requestId: actionRequest.requestId,
      signerId: actionRequest.signerId,
      delegateToEmail: params.delegateEmail,
      delegateToName: params.delegateName,
      reason: params.reason
    });

    return {
      success: true,
      result: { message: 'Signature delegated successfully' },
      correlationId: actionRequest.correlationId
    };
  }

  /**
   * Handle: Get Audit Log
   */
  private async handleGetAuditLog(actionRequest: IActionRequest): Promise<IActionResponse> {
    if (!actionRequest.requestId) {
      return {
        success: false,
        error: 'Missing required parameter: requestId',
        correlationId: actionRequest.correlationId
      };
    }

    const auditLog = await this.signingService.getAuditLog(actionRequest.requestId);

    return {
      success: true,
      result: {
        requestId: actionRequest.requestId,
        entries: auditLog.map(entry => ({
          id: entry.Id,
          action: entry.Action,
          actionBy: entry.ActionByName || entry.ActionByEmail,
          actionDate: entry.ActionDate,
          description: entry.Description,
          signerEmail: entry.SignerEmail,
          previousStatus: entry.PreviousStatus,
          newStatus: entry.NewStatus
        }))
      },
      correlationId: actionRequest.correlationId
    };
  }

  // ============================================
  // WEBHOOK PROCESSING
  // ============================================

  /**
   * Process incoming webhook from external provider
   */
  public async processIncomingWebhook(
    provider: SignatureProvider,
    payload: any,
    signature?: string
  ): Promise<{ success: boolean; message: string }> {
    try {
      logger.info('SigningPowerAutomateService', `Processing incoming webhook from ${provider}`);

      // Verify webhook signature if provided
      if (signature) {
        const isValid = await this.verifyWebhookSignature(provider, payload, signature);
        if (!isValid) {
          logger.warn('SigningPowerAutomateService', 'Invalid webhook signature');
          return { success: false, message: 'Invalid signature' };
        }
      }

      // Process based on provider
      switch (provider) {
        case SignatureProvider.DocuSign:
          return await this.processDocuSignWebhook(payload);
        case SignatureProvider.AdobeSign:
          return await this.processAdobeSignWebhook(payload);
        case SignatureProvider.SigningHub:
          return await this.processSigningHubWebhook(payload);
        default:
          return { success: false, message: `Unknown provider: ${provider}` };
      }
    } catch (error) {
      logger.error('SigningPowerAutomateService', 'Failed to process incoming webhook:', error);
      return { success: false, message: error instanceof Error ? error.message : 'Unknown error' };
    }
  }

  /**
   * Process DocuSign webhook
   */
  private async processDocuSignWebhook(payload: any): Promise<{ success: boolean; message: string }> {
    const event = payload.event;
    const envelopeId = payload.data?.envelopeId || payload.envelopeId;

    if (!envelopeId) {
      return { success: false, message: 'Missing envelope ID' };
    }

    // Find the request by external envelope ID
    const requests = await this.sp.web.lists
      .getByTitle('JML_SigningRequests')
      .items.filter(`ExternalEnvelopeId eq '${envelopeId}'`)
      .top(1)();

    if (requests.length === 0) {
      return { success: false, message: 'Request not found for envelope' };
    }

    const request = requests[0];

    // Map DocuSign event to our status
    switch (event) {
      case 'envelope-sent':
        await this.updateRequestStatus(request.Id, SigningRequestStatus.InProgress);
        break;
      case 'envelope-completed':
        await this.updateRequestStatus(request.Id, SigningRequestStatus.Completed);
        break;
      case 'envelope-declined':
        await this.updateRequestStatus(request.Id, SigningRequestStatus.Declined);
        break;
      case 'envelope-voided':
        await this.updateRequestStatus(request.Id, SigningRequestStatus.Voided);
        break;
      case 'recipient-completed':
        await this.updateSignerStatusByEmail(request.Id, payload.data.recipientEmail, SignerStatus.Signed);
        break;
      case 'recipient-declined':
        await this.updateSignerStatusByEmail(request.Id, payload.data.recipientEmail, SignerStatus.Declined);
        break;
    }

    // Log the webhook
    await this.logWebhook(SignatureProvider.DocuSign, event, request.Id, payload);

    return { success: true, message: `Processed DocuSign event: ${event}` };
  }

  /**
   * Process Adobe Sign webhook
   */
  private async processAdobeSignWebhook(payload: any): Promise<{ success: boolean; message: string }> {
    const event = payload.event;
    const agreementId = payload.agreement?.id;

    if (!agreementId) {
      return { success: false, message: 'Missing agreement ID' };
    }

    const requests = await this.sp.web.lists
      .getByTitle('JML_SigningRequests')
      .items.filter(`ExternalEnvelopeId eq '${agreementId}'`)
      .top(1)();

    if (requests.length === 0) {
      return { success: false, message: 'Request not found for agreement' };
    }

    const request = requests[0];

    switch (event) {
      case 'AGREEMENT_CREATED':
      case 'AGREEMENT_ACTION_REQUESTED':
        await this.updateRequestStatus(request.Id, SigningRequestStatus.InProgress);
        break;
      case 'AGREEMENT_ALL_SIGNED':
        await this.updateRequestStatus(request.Id, SigningRequestStatus.Completed);
        break;
      case 'AGREEMENT_RECALLED':
        await this.updateRequestStatus(request.Id, SigningRequestStatus.Cancelled);
        break;
      case 'AGREEMENT_EXPIRED':
        await this.updateRequestStatus(request.Id, SigningRequestStatus.Expired);
        break;
    }

    await this.logWebhook(SignatureProvider.AdobeSign, event, request.Id, payload);

    return { success: true, message: `Processed Adobe Sign event: ${event}` };
  }

  /**
   * Process Signing Hub webhook
   */
  private async processSigningHubWebhook(payload: any): Promise<{ success: boolean; message: string }> {
    const event = payload.eventType;
    const packageId = payload.packageId;

    if (!packageId) {
      return { success: false, message: 'Missing package ID' };
    }

    const requests = await this.sp.web.lists
      .getByTitle('JML_SigningRequests')
      .items.filter(`ExternalEnvelopeId eq '${packageId}'`)
      .top(1)();

    if (requests.length === 0) {
      return { success: false, message: 'Request not found for package' };
    }

    const request = requests[0];

    switch (event) {
      case 'PACKAGE_CREATED':
      case 'PACKAGE_SENT':
        await this.updateRequestStatus(request.Id, SigningRequestStatus.InProgress);
        break;
      case 'PACKAGE_COMPLETED':
        await this.updateRequestStatus(request.Id, SigningRequestStatus.Completed);
        break;
      case 'PACKAGE_DECLINED':
        await this.updateRequestStatus(request.Id, SigningRequestStatus.Declined);
        break;
      case 'PACKAGE_EXPIRED':
        await this.updateRequestStatus(request.Id, SigningRequestStatus.Expired);
        break;
      case 'PACKAGE_TRASHED':
        await this.updateRequestStatus(request.Id, SigningRequestStatus.Cancelled);
        break;
    }

    await this.logWebhook(SignatureProvider.SigningHub, event, request.Id, payload);

    return { success: true, message: `Processed Signing Hub event: ${event}` };
  }

  // ============================================
  // PRIVATE HELPER METHODS
  // ============================================

  /**
   * Send webhooks for an event
   */
  private async sendWebhooks(event: SigningAuditAction, payload: ISigningWebhookPayload): Promise<ITriggerResponse[]> {
    const responses: ITriggerResponse[] = [];

    for (const config of this.webhookConfigs.values()) {
      if (!config.isActive || !config.events.includes(event)) {
        continue;
      }

      try {
        const response = await this.sendWebhook(config, payload);
        responses.push(response);

        // Log webhook
        await this.logWebhookSent(config, event, payload, response);
      } catch (error) {
        logger.error('SigningPowerAutomateService', `Failed to send webhook to ${config.url}:`, error);
        responses.push({
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error',
          timestamp: new Date()
        });
      }
    }

    return responses;
  }

  /**
   * Send a single webhook
   */
  private async sendWebhook(config: IWebhookConfig, payload: ISigningWebhookPayload): Promise<ITriggerResponse> {
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      ...config.headers
    };

    // Add signature if secret is configured
    if (config.secret) {
      headers['X-Webhook-Signature'] = await this.generateWebhookSignature(payload, config.secret);
    }

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), (config.timeoutSeconds || 30) * 1000);

    try {
      const response = await fetch(config.url, {
        method: 'POST',
        headers,
        body: JSON.stringify(payload),
        signal: controller.signal
      });

      clearTimeout(timeoutId);

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      return {
        success: true,
        triggerId: payload.eventId,
        timestamp: new Date()
      };
    } catch (error) {
      clearTimeout(timeoutId);
      throw error;
    }
  }

  /**
   * Generate webhook signature
   */
  private async generateWebhookSignature(payload: ISigningWebhookPayload, secret: string): Promise<string> {
    const encoder = new TextEncoder();
    const data = encoder.encode(JSON.stringify(payload));
    const key = await crypto.subtle.importKey(
      'raw',
      encoder.encode(secret),
      { name: 'HMAC', hash: 'SHA-256' },
      false,
      ['sign']
    );
    const signature = await crypto.subtle.sign('HMAC', key, data);
    return btoa(String.fromCharCode(...new Uint8Array(signature)));
  }

  /**
   * Verify incoming webhook signature
   */
  private async verifyWebhookSignature(provider: SignatureProvider, payload: any, signature: string): Promise<boolean> {
    // Get provider config for secret
    const configs = await this.signingService.getProviderConfigs();
    const config = configs.find(c => c.Provider === provider);

    if (!config?.WebhookSecret) {
      return true; // No secret configured, allow
    }

    const expectedSignature = await this.generateWebhookSignature(payload, config.WebhookSecret);
    return signature === expectedSignature;
  }

  /**
   * Build request payload for webhook
   */
  private buildRequestPayload(request: ISigningRequest): ISigningWebhookPayload['request'] {
    return {
      id: request.Id!,
      requestNumber: request.RequestNumber,
      title: request.Title,
      status: request.Status,
      provider: request.Provider,
      externalEnvelopeId: request.ExternalEnvelopeId,
      requesterEmail: request.RequesterEmail || '',
      requesterName: request.RequesterName || ''
    };
  }

  /**
   * Build signer payload for webhook
   */
  private buildSignerPayload(signer: ISigner): ISigningWebhookPayload['signer'] {
    return {
      id: signer.Id!,
      name: signer.SignerName,
      email: signer.SignerEmail,
      role: signer.Role,
      status: signer.Status,
      level: signer.Level
    };
  }

  /**
   * Generate unique event ID
   */
  private generateEventId(): string {
    return `evt_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * Update request status
   */
  private async updateRequestStatus(requestId: number, status: SigningRequestStatus): Promise<void> {
    const updates: any = { Status: status };

    if (status === SigningRequestStatus.Completed) {
      updates.CompletedDate = new Date().toISOString();
    }

    await this.sp.web.lists
      .getByTitle('JML_SigningRequests')
      .items.getById(requestId)
      .update(updates);
  }

  /**
   * Update signer status by email
   */
  private async updateSignerStatusByEmail(requestId: number, email: string, status: SignerStatus): Promise<void> {
    const signers = await this.sp.web.lists
      .getByTitle('JML_Signers')
      .items.filter(`RequestId eq ${requestId} and SignerEmail eq '${email}'`)();

    for (const signer of signers) {
      const updates: any = { Status: status };

      if (status === SignerStatus.Signed) {
        updates.SignedDate = new Date().toISOString();
      } else if (status === SignerStatus.Declined) {
        updates.DeclinedDate = new Date().toISOString();
      }

      await this.sp.web.lists
        .getByTitle('JML_Signers')
        .items.getById(signer.Id)
        .update(updates);
    }
  }

  /**
   * Log incoming webhook
   */
  private async logWebhook(provider: SignatureProvider, event: string, requestId: number, payload: any): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.WEBHOOK_LOG_LIST)
        .items.add({
          Title: `${provider} - ${event}`,
          Provider: provider,
          Event: event,
          RequestId: requestId,
          Direction: 'Incoming',
          PayloadJSON: JSON.stringify(payload),
          ReceivedDate: new Date().toISOString(),
          Status: 'Processed'
        });
    } catch (error) {
      logger.warn('SigningPowerAutomateService', 'Failed to log webhook:', error);
    }
  }

  /**
   * Log outgoing webhook
   */
  private async logWebhookSent(
    config: IWebhookConfig,
    event: SigningAuditAction,
    payload: ISigningWebhookPayload,
    response: ITriggerResponse
  ): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.WEBHOOK_LOG_LIST)
        .items.add({
          Title: `${config.name} - ${event}`,
          WebhookConfigId: config.id,
          WebhookName: config.name,
          WebhookUrl: config.url,
          Event: event,
          RequestId: payload.request.id,
          Direction: 'Outgoing',
          PayloadJSON: JSON.stringify(payload),
          SentDate: new Date().toISOString(),
          Status: response.success ? 'Success' : 'Failed',
          ErrorMessage: response.error
        });
    } catch (error) {
      logger.warn('SigningPowerAutomateService', 'Failed to log outgoing webhook:', error);
    }
  }

  // ============================================
  // WEBHOOK MANAGEMENT
  // ============================================

  /**
   * Register a new webhook
   */
  public async registerWebhook(config: Omit<IWebhookConfig, 'id'>): Promise<IWebhookConfig> {
    const result = await this.sp.web.lists
      .getByTitle(this.WEBHOOK_CONFIG_LIST)
      .items.add({
        Title: config.name,
        WebhookUrl: config.url,
        WebhookSecret: config.secret || null,
        Events: JSON.stringify(config.events),
        IsActive: config.isActive,
        Headers: config.headers ? JSON.stringify(config.headers) : null,
        RetryCount: config.retryCount || 3,
        TimeoutSeconds: config.timeoutSeconds || 30
      });

    const newConfig: IWebhookConfig = {
      ...config,
      id: result.data.Id.toString()
    };

    this.webhookConfigs.set(newConfig.id, newConfig);

    return newConfig;
  }

  /**
   * Update webhook configuration
   */
  public async updateWebhook(id: string, updates: Partial<IWebhookConfig>): Promise<void> {
    const updateData: any = {};

    if (updates.name) updateData.Title = updates.name;
    if (updates.url) updateData.WebhookUrl = updates.url;
    if (updates.secret !== undefined) updateData.WebhookSecret = updates.secret;
    if (updates.events) updateData.Events = JSON.stringify(updates.events);
    if (updates.isActive !== undefined) updateData.IsActive = updates.isActive;
    if (updates.headers) updateData.Headers = JSON.stringify(updates.headers);
    if (updates.retryCount !== undefined) updateData.RetryCount = updates.retryCount;
    if (updates.timeoutSeconds !== undefined) updateData.TimeoutSeconds = updates.timeoutSeconds;

    await this.sp.web.lists
      .getByTitle(this.WEBHOOK_CONFIG_LIST)
      .items.getById(parseInt(id, 10))
      .update(updateData);

    // Reload configs
    await this.loadWebhookConfigs();
  }

  /**
   * Delete webhook configuration
   */
  public async deleteWebhook(id: string): Promise<void> {
    await this.sp.web.lists
      .getByTitle(this.WEBHOOK_CONFIG_LIST)
      .items.getById(parseInt(id, 10))
      .delete();

    this.webhookConfigs.delete(id);
  }

  /**
   * Get all webhook configurations
   */
  public getWebhookConfigs(): IWebhookConfig[] {
    return Array.from(this.webhookConfigs.values());
  }

  /**
   * Test a webhook configuration
   */
  public async testWebhook(id: string): Promise<ITriggerResponse> {
    const config = this.webhookConfigs.get(id);

    if (!config) {
      return {
        success: false,
        error: 'Webhook configuration not found',
        timestamp: new Date()
      };
    }

    const testPayload: ISigningWebhookPayload = {
      eventId: this.generateEventId(),
      eventType: SigningAuditAction.Created,
      timestamp: new Date(),
      request: {
        id: 0,
        requestNumber: 'TEST-0000',
        title: 'Test Webhook',
        status: SigningRequestStatus.Draft,
        provider: SignatureProvider.Internal,
        requesterEmail: 'test@example.com',
        requesterName: 'Test User'
      },
      details: {
        isTest: true
      }
    };

    return this.sendWebhook(config, testPayload);
  }
}

export default SigningPowerAutomateService;
