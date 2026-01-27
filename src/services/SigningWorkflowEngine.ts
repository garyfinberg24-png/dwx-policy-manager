// @ts-nocheck
/* eslint-disable @typescript-eslint/no-explicit-any */
// Signing Workflow Engine
// Handles workflow execution, escalation, reminders, and expiration logic
// Note: Some fields may not exist in the SharePoint list - mapping handles this gracefully

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import {
  ISigningRequest,
  ISigner,
  ISigningLevel,
  SigningRequestStatus,
  SigningWorkflowType,
  SignerStatus,
  SigningAuditAction,
  SigningEscalationAction
} from '../models/ISigning';
import { SigningService } from './SigningService';
import { SigningNotificationService } from './SigningNotificationService';
import { logger } from './LoggingService';

/**
 * Signing Workflow Engine
 * Manages workflow progression, reminders, escalations, and expirations
 */
export class SigningWorkflowEngine {
  private sp: SPFI;
  private signingService: SigningService;
  private notificationService: SigningNotificationService;

  private readonly REQUESTS_LIST = 'JML_SigningRequests';
  private readonly SIGNERS_LIST = 'JML_Signers';
  private readonly CHAINS_LIST = 'JML_SigningChains';

  constructor(sp: SPFI) {
    this.sp = sp;
    this.signingService = new SigningService(sp);
    this.notificationService = new SigningNotificationService(sp);
  }

  /**
   * Initialize the engine
   */
  public async initialize(): Promise<void> {
    await this.signingService.initialize();
    logger.info('SigningWorkflowEngine', 'Initialized');
  }

  // ============================================
  // WORKFLOW EXECUTION
  // ============================================

  /**
   * Execute workflow for a signing request
   */
  public async executeWorkflow(requestId: number): Promise<void> {
    try {
      const request = await this.signingService.getSigningRequestById(requestId);

      if (request.Status !== SigningRequestStatus.InProgress) {
        logger.warn('SigningWorkflowEngine', `Cannot execute workflow - request ${requestId} is ${request.Status}`);
        return;
      }

      // Get current level
      const currentLevel = request.SigningChain.CurrentLevel;
      const levelConfig = request.SigningChain.Levels.find(l => l.level === currentLevel);

      if (!levelConfig) {
        logger.error('SigningWorkflowEngine', `Level ${currentLevel} not found for request ${requestId}`);
        return;
      }

      // Execute based on workflow type
      switch (request.WorkflowType) {
        case SigningWorkflowType.Sequential:
          await this.executeSequentialLevel(request, levelConfig);
          break;
        case SigningWorkflowType.Parallel:
          await this.executeParallelLevel(request, levelConfig);
          break;
        case SigningWorkflowType.Hybrid:
          await this.executeHybridLevel(request, levelConfig);
          break;
        case SigningWorkflowType.FirstSigner:
          await this.executeFirstSignerLevel(request, levelConfig);
          break;
        case SigningWorkflowType.ApprovalThenSign:
          await this.executeApprovalThenSignLevel(request, levelConfig);
          break;
        default:
          await this.executeSequentialLevel(request, levelConfig);
      }

      logger.info('SigningWorkflowEngine', `Executed workflow for request ${requestId}`);
    } catch (error) {
      logger.error('SigningWorkflowEngine', `Failed to execute workflow for request ${requestId}:`, error);
      throw error;
    }
  }

  /**
   * Execute sequential workflow level
   * Signers sign one after another within the level
   */
  private async executeSequentialLevel(request: ISigningRequest, level: ISigningLevel): Promise<void> {
    const signers = level.signers.sort((a, b) => a.Order - b.Order);

    // Find the first pending signer
    const pendingSigner = signers.find(s =>
      s.Status === SignerStatus.Pending || s.Status === SignerStatus.Sent
    );

    if (pendingSigner && pendingSigner.Status === SignerStatus.Pending) {
      // Activate this signer
      await this.activateSigner(request.Id!, pendingSigner.Id!);
    }
  }

  /**
   * Execute parallel workflow level
   * All signers in the level can sign simultaneously
   */
  private async executeParallelLevel(request: ISigningRequest, level: ISigningLevel): Promise<void> {
    const pendingSigners = level.signers.filter(s => s.Status === SignerStatus.Pending);

    // Activate all pending signers
    for (const signer of pendingSigners) {
      await this.activateSigner(request.Id!, signer.Id!);
    }
  }

  /**
   * Execute hybrid workflow level
   * Uses the level's specific workflow type
   */
  private async executeHybridLevel(request: ISigningRequest, level: ISigningLevel): Promise<void> {
    if (level.workflowType === SigningWorkflowType.Parallel) {
      await this.executeParallelLevel(request, level);
    } else {
      await this.executeSequentialLevel(request, level);
    }
  }

  /**
   * Execute first-signer-wins workflow level
   * First person to sign completes the level
   */
  private async executeFirstSignerLevel(request: ISigningRequest, level: ISigningLevel): Promise<void> {
    // Activate all signers - first one to sign wins
    await this.executeParallelLevel(request, level);
  }

  /**
   * Execute approval-then-sign workflow
   * First levels are approvals, later levels are signatures
   */
  private async executeApprovalThenSignLevel(request: ISigningRequest, level: ISigningLevel): Promise<void> {
    // Determine if this is an approval or signing level based on signer roles
    const hasApprovers = level.signers.some(s => s.Role === 'Approver');

    if (hasApprovers) {
      // Treat as approval - can be parallel or sequential
      await this.executeParallelLevel(request, level);
    } else {
      // Treat as signing - usually sequential
      await this.executeSequentialLevel(request, level);
    }
  }

  /**
   * Activate a specific signer
   */
  private async activateSigner(requestId: number, signerId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.getById(signerId)
        .update({
          Status: SignerStatus.Sent,
          SentDate: new Date().toISOString()
        });

      // Send notification
      await this.notificationService.sendSignatureRequestNotification(requestId, signerId);

      logger.info('SigningWorkflowEngine', `Activated signer ${signerId} for request ${requestId}`);
    } catch (error) {
      logger.error('SigningWorkflowEngine', `Failed to activate signer ${signerId}:`, error);
      throw error;
    }
  }

  // ============================================
  // LEVEL EVALUATION
  // ============================================

  /**
   * Evaluate if a level is complete and advance if needed
   */
  public async evaluateLevelCompletion(requestId: number, level: number): Promise<boolean> {
    try {
      const request = await this.signingService.getSigningRequestById(requestId);
      const levelConfig = request.SigningChain.Levels.find(l => l.level === level);

      if (!levelConfig) {
        return false;
      }

      const signers = levelConfig.signers;
      let isComplete = false;

      switch (request.WorkflowType) {
        case SigningWorkflowType.Sequential:
        case SigningWorkflowType.Parallel:
        case SigningWorkflowType.Hybrid:
          // All signers must sign (or meet required count)
          if (levelConfig.requiredSignatures) {
            const signedCount = signers.filter(s => s.Status === SignerStatus.Signed).length;
            isComplete = signedCount >= levelConfig.requiredSignatures;
          } else {
            isComplete = signers.every(s =>
              s.Status === SignerStatus.Signed || s.Status === SignerStatus.Delegated
            );
          }
          break;

        case SigningWorkflowType.FirstSigner:
          // First signer to sign completes the level
          isComplete = signers.some(s => s.Status === SignerStatus.Signed);
          break;

        case SigningWorkflowType.ApprovalThenSign:
          // All must complete
          isComplete = signers.every(s =>
            s.Status === SignerStatus.Signed || s.Status === SignerStatus.Delegated
          );
          break;
      }

      if (isComplete) {
        await this.advanceToNextLevel(requestId);
      }

      return isComplete;
    } catch (error) {
      logger.error('SigningWorkflowEngine', `Failed to evaluate level completion:`, error);
      return false;
    }
  }

  /**
   * Advance workflow to the next level
   */
  public async advanceToNextLevel(requestId: number): Promise<void> {
    try {
      const request = await this.signingService.getSigningRequestById(requestId);
      const currentLevel = request.SigningChain.CurrentLevel;
      const nextLevel = currentLevel + 1;

      // Check if there are more levels
      const hasNextLevel = request.SigningChain.Levels.some(l => l.level === nextLevel);

      if (hasNextLevel) {
        // Update chain to next level
        await this.updateChainLevel(requestId, nextLevel);

        // Execute the next level
        await this.executeWorkflow(requestId);

        await this.signingService.logAuditEntry({
          RequestId: requestId,
          Action: SigningAuditAction.Sent,
          Description: `Advanced to level ${nextLevel}`,
          IsSystemAction: true
        });

        logger.info('SigningWorkflowEngine', `Advanced request ${requestId} to level ${nextLevel}`);
      } else {
        // All levels complete - complete the request
        await this.completeRequest(requestId);
      }
    } catch (error) {
      logger.error('SigningWorkflowEngine', `Failed to advance to next level:`, error);
      throw error;
    }
  }

  /**
   * Update the chain's current level
   */
  private async updateChainLevel(requestId: number, newLevel: number): Promise<void> {
    const chains = await this.sp.web.lists
      .getByTitle(this.CHAINS_LIST)
      .items.filter(`RequestId eq ${requestId}`)
      .top(1)();

    if (chains.length > 0) {
      await this.sp.web.lists
        .getByTitle(this.CHAINS_LIST)
        .items.getById(chains[0].Id)
        .update({
          CurrentLevel: newLevel
        });
    }
  }

  /**
   * Complete a signing request
   */
  private async completeRequest(requestId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.REQUESTS_LIST)
        .items.getById(requestId)
        .update({
          Status: SigningRequestStatus.Completed,
          CompletedDate: new Date().toISOString()
        });

      // Update chain status
      const chains = await this.sp.web.lists
        .getByTitle(this.CHAINS_LIST)
        .items.filter(`RequestId eq ${requestId}`)
        .top(1)();

      if (chains.length > 0) {
        await this.sp.web.lists
          .getByTitle(this.CHAINS_LIST)
          .items.getById(chains[0].Id)
          .update({
            Status: SigningRequestStatus.Completed,
            CompletedDate: new Date().toISOString()
          });
      }

      // Log audit entry
      await this.signingService.logAuditEntry({
        RequestId: requestId,
        Action: SigningAuditAction.Completed,
        Description: 'All signatures collected - request completed',
        IsSystemAction: true
      });

      // Send completion notifications
      await this.notificationService.sendCompletionNotification(requestId);

      // Generate certificate of completion
      await this.generateCertificateOfCompletion(requestId);

      logger.info('SigningWorkflowEngine', `Completed request ${requestId}`);
    } catch (error) {
      logger.error('SigningWorkflowEngine', `Failed to complete request ${requestId}:`, error);
      throw error;
    }
  }

  // ============================================
  // REMINDERS
  // ============================================

  /**
   * Process reminders for all overdue signers
   */
  public async processReminders(): Promise<number> {
    try {
      const now = new Date();
      let remindersSent = 0;

      // Get all active requests with reminders enabled
      const requests = await this.sp.web.lists
        .getByTitle(this.REQUESTS_LIST)
        .items.select('Id', 'ReminderEnabled', 'ReminderDays', 'Status')
        .filter(`Status eq '${SigningRequestStatus.InProgress}' and ReminderEnabled eq true`)();

      for (const request of requests) {
        const signersDueForReminder = await this.getSignersDueForReminder(
          request.Id,
          request.ReminderDays
        );

        for (const signer of signersDueForReminder) {
          await this.sendReminderToSigner(request.Id, signer.Id);
          remindersSent++;
        }
      }

      logger.info('SigningWorkflowEngine', `Processed ${remindersSent} reminders`);
      return remindersSent;
    } catch (error) {
      logger.error('SigningWorkflowEngine', 'Failed to process reminders:', error);
      return 0;
    }
  }

  /**
   * Get signers who are due for a reminder
   */
  private async getSignersDueForReminder(requestId: number, reminderDays: number): Promise<any[]> {
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - reminderDays);

    const signers = await this.sp.web.lists
      .getByTitle(this.SIGNERS_LIST)
      .items.filter(
        `RequestId eq ${requestId} and ` +
        `(Status eq '${SignerStatus.Sent}' or Status eq '${SignerStatus.Viewed}') and ` +
        `SentDate lt datetime'${cutoffDate.toISOString()}'`
      )();

    // Filter to signers who haven't been reminded recently
    const recentReminderCutoff = new Date();
    recentReminderCutoff.setDate(recentReminderCutoff.getDate() - 1); // Don't remind more than once per day

    return signers.filter(s =>
      !s.LastReminderDate || new Date(s.LastReminderDate) < recentReminderCutoff
    );
  }

  /**
   * Send reminder to a specific signer
   */
  public async sendReminderToSigner(requestId: number, signerId: number): Promise<void> {
    try {
      // Send notification
      await this.notificationService.sendReminderNotification(requestId, signerId);

      // Update signer record
      const signer = await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.getById(signerId)();

      await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.getById(signerId)
        .update({
          RemindersSent: (signer.RemindersSent || 0) + 1,
          LastReminderDate: new Date().toISOString()
        });

      // Log audit entry
      await this.signingService.logAuditEntry({
        RequestId: requestId,
        SignerId: signerId,
        SignerEmail: signer.SignerEmail,
        Action: SigningAuditAction.Reminded,
        Description: `Reminder sent to ${signer.SignerEmail}`,
        IsSystemAction: true
      });

      logger.info('SigningWorkflowEngine', `Sent reminder to signer ${signerId}`);
    } catch (error) {
      logger.error('SigningWorkflowEngine', `Failed to send reminder to signer ${signerId}:`, error);
    }
  }

  // ============================================
  // ESCALATION
  // ============================================

  /**
   * Process escalations for overdue requests
   */
  public async processEscalations(): Promise<number> {
    try {
      let escalationsProcessed = 0;

      // Get requests due for escalation
      const overdueRequests = await this.getRequestsDueForEscalation();

      for (const request of overdueRequests) {
        await this.escalateRequest(request.Id, request.EscalationAction);
        escalationsProcessed++;
      }

      logger.info('SigningWorkflowEngine', `Processed ${escalationsProcessed} escalations`);
      return escalationsProcessed;
    } catch (error) {
      logger.error('SigningWorkflowEngine', 'Failed to process escalations:', error);
      return 0;
    }
  }

  /**
   * Get requests that are due for escalation
   */
  private async getRequestsDueForEscalation(): Promise<any[]> {
    const requests = await this.sp.web.lists
      .getByTitle(this.REQUESTS_LIST)
      .items.select('Id', 'EscalationEnabled', 'EscalationDays', 'EscalationAction', 'SentDate', 'Status')
      .filter(`Status eq '${SigningRequestStatus.InProgress}' and EscalationEnabled eq true`)();

    const now = new Date();

    return requests.filter(r => {
      if (!r.SentDate) return false;
      const sentDate = new Date(r.SentDate);
      const daysSinceSent = Math.floor((now.getTime() - sentDate.getTime()) / (1000 * 60 * 60 * 24));
      return daysSinceSent >= r.EscalationDays;
    });
  }

  /**
   * Escalate a specific request
   */
  public async escalateRequest(requestId: number, action?: SigningEscalationAction): Promise<void> {
    try {
      const request = await this.signingService.getSigningRequestById(requestId);
      const escalationAction = action || request.EscalationAction || SigningEscalationAction.Notify;

      switch (escalationAction) {
        case SigningEscalationAction.Notify:
          await this.notificationService.sendEscalationNotification(requestId);
          break;

        case SigningEscalationAction.NotifyManager:
          await this.notifyManagerOfEscalation(requestId);
          break;

        case SigningEscalationAction.NotifyRequester:
          await this.notificationService.sendRequesterEscalationNotification(requestId);
          break;

        case SigningEscalationAction.Reassign:
          // This would need manager info to reassign
          await this.notificationService.sendEscalationNotification(requestId);
          break;

        case SigningEscalationAction.AutoApprove:
          await this.autoApproveRequest(requestId);
          break;

        case SigningEscalationAction.Cancel:
          await this.signingService.cancelSigningRequest({
            requestId,
            reason: 'Auto-cancelled due to escalation policy',
            notifySigners: true
          });
          break;

        default:
          await this.notificationService.sendEscalationNotification(requestId);
      }

      // Log audit entry
      await this.signingService.logAuditEntry({
        RequestId: requestId,
        Action: SigningAuditAction.Escalated,
        Description: `Request escalated - action: ${escalationAction}`,
        IsSystemAction: true
      });

      logger.info('SigningWorkflowEngine', `Escalated request ${requestId} with action ${escalationAction}`);
    } catch (error) {
      logger.error('SigningWorkflowEngine', `Failed to escalate request ${requestId}:`, error);
    }
  }

  /**
   * Notify manager about escalation
   */
  private async notifyManagerOfEscalation(requestId: number): Promise<void> {
    // This would integrate with your org structure to find the manager
    // For now, just send to requester
    await this.notificationService.sendRequesterEscalationNotification(requestId);
  }

  /**
   * Auto-approve a request (for escalation)
   */
  private async autoApproveRequest(requestId: number): Promise<void> {
    try {
      // Get pending signers and mark them as signed
      const signers = await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.filter(
          `RequestId eq ${requestId} and ` +
          `(Status eq '${SignerStatus.Sent}' or Status eq '${SignerStatus.Viewed}')`
        )();

      for (const signer of signers) {
        await this.sp.web.lists
          .getByTitle(this.SIGNERS_LIST)
          .items.getById(signer.Id)
          .update({
            Status: SignerStatus.Signed,
            SignedDate: new Date().toISOString(),
            Comments: 'Auto-approved due to escalation policy'
          });

        await this.signingService.logAuditEntry({
          RequestId: requestId,
          SignerId: signer.Id,
          SignerEmail: signer.SignerEmail,
          Action: SigningAuditAction.Signed,
          Description: 'Auto-approved due to escalation policy',
          IsSystemAction: true
        });
      }

      // Complete the request
      await this.completeRequest(requestId);
    } catch (error) {
      logger.error('SigningWorkflowEngine', `Failed to auto-approve request ${requestId}:`, error);
    }
  }

  // ============================================
  // EXPIRATION
  // ============================================

  /**
   * Process expirations for overdue requests
   */
  public async processExpirations(): Promise<number> {
    try {
      let expirationsProcessed = 0;
      const now = new Date();

      // Get expired requests
      const expiredRequests = await this.sp.web.lists
        .getByTitle(this.REQUESTS_LIST)
        .items.select('Id', 'ExpirationDate', 'Status')
        .filter(
          `Status eq '${SigningRequestStatus.InProgress}' or ` +
          `Status eq '${SigningRequestStatus.Pending}'`
        )();

      for (const request of expiredRequests) {
        if (request.ExpirationDate && new Date(request.ExpirationDate) < now) {
          await this.expireRequest(request.Id);
          expirationsProcessed++;
        }
      }

      logger.info('SigningWorkflowEngine', `Processed ${expirationsProcessed} expirations`);
      return expirationsProcessed;
    } catch (error) {
      logger.error('SigningWorkflowEngine', 'Failed to process expirations:', error);
      return 0;
    }
  }

  /**
   * Expire a specific request
   */
  public async expireRequest(requestId: number): Promise<void> {
    try {
      // Update request status
      await this.sp.web.lists
        .getByTitle(this.REQUESTS_LIST)
        .items.getById(requestId)
        .update({
          Status: SigningRequestStatus.Expired
        });

      // Update pending signers
      const signers = await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.filter(
          `RequestId eq ${requestId} and Status ne '${SignerStatus.Signed}'`
        )();

      for (const signer of signers) {
        await this.sp.web.lists
          .getByTitle(this.SIGNERS_LIST)
          .items.getById(signer.Id)
          .update({
            Status: SignerStatus.Expired
          });
      }

      // Log audit entry
      await this.signingService.logAuditEntry({
        RequestId: requestId,
        Action: SigningAuditAction.Expired,
        Description: 'Request expired',
        IsSystemAction: true
      });

      // Send expiration notifications
      await this.notificationService.sendExpirationNotification(requestId);

      logger.info('SigningWorkflowEngine', `Expired request ${requestId}`);
    } catch (error) {
      logger.error('SigningWorkflowEngine', `Failed to expire request ${requestId}:`, error);
    }
  }

  /**
   * Send expiration warning notifications
   */
  public async sendExpirationWarnings(warningDays: number = 3): Promise<number> {
    try {
      let warningsSent = 0;
      const warningDate = new Date();
      warningDate.setDate(warningDate.getDate() + warningDays);

      // Get requests expiring soon
      const expiringRequests = await this.sp.web.lists
        .getByTitle(this.REQUESTS_LIST)
        .items.select('Id', 'ExpirationDate', 'Status')
        .filter(`Status eq '${SigningRequestStatus.InProgress}'`)();

      for (const request of expiringRequests) {
        if (request.ExpirationDate) {
          const expDate = new Date(request.ExpirationDate);
          if (expDate <= warningDate && expDate > new Date()) {
            await this.notificationService.sendExpirationWarningNotification(request.Id);
            warningsSent++;
          }
        }
      }

      logger.info('SigningWorkflowEngine', `Sent ${warningsSent} expiration warnings`);
      return warningsSent;
    } catch (error) {
      logger.error('SigningWorkflowEngine', 'Failed to send expiration warnings:', error);
      return 0;
    }
  }

  // ============================================
  // CERTIFICATE GENERATION
  // ============================================

  /**
   * Generate certificate of completion
   */
  private async generateCertificateOfCompletion(requestId: number): Promise<void> {
    try {
      const request = await this.signingService.getSigningRequestById(requestId);

      // Get all signers with their signature data
      const signers = await this.sp.web.lists
        .getByTitle(this.SIGNERS_LIST)
        .items.filter(`RequestId eq ${requestId}`)
        .orderBy('Level', true)
        .orderBy('SignedDate', true)();

      // Build certificate data
      const certificateData = {
        requestId: request.Id,
        requestNumber: request.RequestNumber,
        title: request.Title,
        documents: request.Documents || [],
        signers: signers.map(s => ({
          name: s.SignerName,
          email: s.SignerEmail,
          role: s.SignerRole,
          signedDate: s.SignedDate,
          ipAddress: s.IPAddress,
          signatureType: s.SignatureType
        })),
        createdDate: request.Created,
        sentDate: request.SentDate,
        completedDate: request.CompletedDate,
        certificateId: `CERT-${request.RequestNumber}-${Date.now()}`,
        generatedDate: new Date()
      };

      // In a real implementation, this would:
      // 1. Generate a PDF certificate
      // 2. Store it in SharePoint
      // 3. Update the request with the certificate URL

      // Log certificate generation
      await this.signingService.logAuditEntry({
        RequestId: requestId,
        Action: SigningAuditAction.CertificateGenerated,
        Description: `Certificate of completion generated: ${certificateData.certificateId}`,
        Details: { certificateId: certificateData.certificateId },
        IsSystemAction: true
      });

      logger.info('SigningWorkflowEngine', `Generated certificate for request ${requestId}`);
    } catch (error) {
      logger.error('SigningWorkflowEngine', `Failed to generate certificate for request ${requestId}:`, error);
    }
  }

  // ============================================
  // BATCH PROCESSING
  // ============================================

  /**
   * Run all scheduled tasks (reminders, escalations, expirations)
   */
  public async runScheduledTasks(): Promise<{
    reminders: number;
    escalations: number;
    expirations: number;
    warnings: number;
  }> {
    logger.info('SigningWorkflowEngine', 'Running scheduled tasks...');

    const results = {
      reminders: 0,
      escalations: 0,
      expirations: 0,
      warnings: 0
    };

    try {
      // Process in order: expirations first, then escalations, then reminders
      results.expirations = await this.processExpirations();
      results.escalations = await this.processEscalations();
      results.reminders = await this.processReminders();
      results.warnings = await this.sendExpirationWarnings();

      logger.info('SigningWorkflowEngine', 'Scheduled tasks completed', results);
    } catch (error) {
      logger.error('SigningWorkflowEngine', 'Error running scheduled tasks:', error);
    }

    return results;
  }

  // ============================================
  // WORKFLOW STATUS
  // ============================================

  /**
   * Get workflow status for a request
   */
  public async getWorkflowStatus(requestId: number): Promise<{
    currentLevel: number;
    totalLevels: number;
    overallProgress: number;
    levelProgress: { level: number; progress: number; status: string }[];
    nextActions: string[];
    estimatedCompletionDate?: Date;
  }> {
    try {
      const request = await this.signingService.getSigningRequestById(requestId);
      const levels = request.SigningChain.Levels;

      // Calculate level progress
      const levelProgress = levels.map(level => {
        const totalSigners = level.signers.length;
        const signedSigners = level.signers.filter(s => s.Status === SignerStatus.Signed).length;
        return {
          level: level.level,
          progress: totalSigners > 0 ? Math.round((signedSigners / totalSigners) * 100) : 0,
          status: level.status
        };
      });

      // Calculate overall progress
      const totalSigners = levels.reduce((sum, l) => sum + l.signers.length, 0);
      const signedSigners = levels.reduce((sum, l) =>
        sum + l.signers.filter(s => s.Status === SignerStatus.Signed).length, 0);
      const overallProgress = totalSigners > 0 ? Math.round((signedSigners / totalSigners) * 100) : 0;

      // Determine next actions
      const nextActions: string[] = [];
      const currentLevel = levels.find(l => l.level === request.SigningChain.CurrentLevel);
      if (currentLevel) {
        const pendingSigners = currentLevel.signers.filter(s =>
          s.Status === SignerStatus.Sent || s.Status === SignerStatus.Viewed
        );
        pendingSigners.forEach(s => {
          nextActions.push(`Waiting for ${s.SignerName} to sign`);
        });
      }

      return {
        currentLevel: request.SigningChain.CurrentLevel,
        totalLevels: request.SigningChain.TotalLevels,
        overallProgress,
        levelProgress,
        nextActions,
        estimatedCompletionDate: this.estimateCompletionDate(request)
      };
    } catch (error) {
      logger.error('SigningWorkflowEngine', `Failed to get workflow status for request ${requestId}:`, error);
      throw error;
    }
  }

  /**
   * Estimate completion date based on average completion times
   */
  private estimateCompletionDate(request: ISigningRequest): Date | undefined {
    const remainingLevels = request.SigningChain.TotalLevels - request.SigningChain.CurrentLevel + 1;
    const avgDaysPerLevel = 2; // This could be calculated from historical data

    const estimatedDate = new Date();
    estimatedDate.setDate(estimatedDate.getDate() + (remainingLevels * avgDaysPerLevel));

    // Don't exceed due date
    if (request.DueDate && estimatedDate > request.DueDate) {
      return request.DueDate;
    }

    return estimatedDate;
  }
}

export default SigningWorkflowEngine;
