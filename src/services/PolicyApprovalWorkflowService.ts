// @ts-nocheck
/**
 * Policy Approval Workflow Service
 * Manages multi-stage approval workflows, delegation, escalation, and Teams notifications
 * for enterprise Policy Management
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import '@pnp/sp/sputilities';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../utils/pnpConfig';
import { logger } from './LoggingService';
import { IPolicy, PolicyStatus } from '../models/IPolicy';
import { PolicyLists, PolicyWorkflowLists } from '../constants/SharePointListNames';

// ============================================================================
// ENUMS
// ============================================================================

export enum PolicyApprovalStatus {
  Draft = 'Draft',
  PendingReview = 'Pending Review',
  PendingApproval = 'Pending Approval',
  Approved = 'Approved',
  Rejected = 'Rejected',
  Recalled = 'Recalled',
  Expired = 'Expired'
}

export enum ApprovalStageType {
  Review = 'Review',           // Content review
  LegalReview = 'Legal Review', // Legal compliance review
  Approval = 'Approval',       // Standard approval
  FinalApproval = 'Final Approval', // Executive sign-off
  Acknowledgement = 'Acknowledgement' // FYI notification
}

export enum ApprovalRuleType {
  AllMustApprove = 'All Must Approve',    // Unanimous
  AnyOneApproves = 'Any One Approves',    // First response wins
  MajorityApproves = 'Majority Approves', // 50%+ must approve
  QuorumApproves = 'Quorum Approves'      // Configurable threshold
}

export enum EscalationActionType {
  Notify = 'Notify',           // Send reminder
  NotifyManager = 'Notify Manager', // Escalate to manager
  AutoApprove = 'Auto Approve', // Auto-approve after timeout
  AutoReject = 'Auto Reject',   // Auto-reject after timeout
  Reassign = 'Reassign'         // Reassign to backup
}

export enum DelegationType {
  Temporary = 'Temporary',      // Time-bound delegation
  Permanent = 'Permanent',      // Ongoing delegation
  OutOfOffice = 'Out of Office' // OOO auto-delegation
}

// ============================================================================
// INTERFACES
// ============================================================================

export interface IApprovalStage {
  stageId: string;
  stageName: string;
  stageType: ApprovalStageType;
  stageOrder: number;
  approverIds: number[];
  approverRoles?: string[];     // Role-based approvers (e.g., "Legal Team", "HR Manager")
  approvalRule: ApprovalRuleType;
  quorumPercentage?: number;    // For QuorumApproves
  dueDays: number;
  requireComments: boolean;
  allowDelegation: boolean;
  escalationEnabled: boolean;
  escalationDays?: number;
  escalationAction?: EscalationActionType;
  escalationTargetIds?: number[];
  notifyOnComplete: boolean;
  instructions?: string;
}

export interface IApprovalWorkflowTemplate {
  id?: number;
  templateName: string;
  templateDescription?: string;
  category?: string;           // Policy category this applies to
  stages: IApprovalStage[];
  isActive: boolean;
  isDefault: boolean;
  createdById?: number;
  createdDate?: Date;
  modifiedDate?: Date;
}

export interface IApprovalWorkflowInstance {
  id?: number;
  policyId: number;
  policyNumber?: string;
  policyName?: string;
  templateId?: number;
  templateName?: string;
  currentStageId: string;
  currentStageOrder: number;
  overallStatus: PolicyApprovalStatus;
  stages: IApprovalStageInstance[];
  initiatedById: number;
  initiatedByName?: string;
  initiatedDate: Date;
  completedDate?: Date;
  completedById?: number;
  totalDurationDays?: number;
  comments?: string;
}

export interface IApprovalStageInstance {
  stageId: string;
  stageName: string;
  stageType: ApprovalStageType;
  stageOrder: number;
  status: PolicyApprovalStatus;
  approvalRule: ApprovalRuleType;
  approvals: IApprovalDecision[];
  startedDate?: Date;
  completedDate?: Date;
  dueDate?: Date;
  escalationLevel: number;
  lastEscalationDate?: Date;
}

export interface IApprovalDecision {
  id?: number;
  workflowInstanceId: number;
  stageId: string;
  approverId: number;
  approverName?: string;
  approverEmail?: string;
  originalApproverId?: number;  // If delegated
  delegatedById?: number;
  status: PolicyApprovalStatus;
  decision?: 'Approved' | 'Rejected' | 'Pending';
  comments?: string;
  requestedDate: Date;
  respondedDate?: Date;
  dueDate: Date;
  isOverdue: boolean;
  escalationLevel: number;
  notificationsSent: number;
  lastNotificationDate?: Date;
}

export interface IApprovalDelegation {
  id?: number;
  delegatorId: number;
  delegatorName?: string;
  delegatorEmail?: string;
  delegateId: number;
  delegateName?: string;
  delegateEmail?: string;
  delegationType: DelegationType;
  startDate: Date;
  endDate?: Date;
  reason?: string;
  policyCategories?: string[];  // Limit to specific categories
  isActive: boolean;
  createdDate: Date;
  approvedById?: number;
  approvedDate?: Date;
}

export interface IEscalationRule {
  id?: number;
  ruleName: string;
  ruleDescription?: string;
  triggerDays: number;          // Days after due date
  triggerCondition: 'Overdue' | 'NoResponse' | 'StageStuck';
  action: EscalationActionType;
  targetType: 'Manager' | 'SpecificUser' | 'Role' | 'BackupApprover';
  targetIds?: number[];
  targetRole?: string;
  notificationTemplate?: string;
  maxEscalations: number;
  escalationIntervalDays: number;
  isActive: boolean;
  appliesToCategories?: string[];
  priority: number;
}

export interface ITeamsAdaptiveCardPayload {
  type: 'AdaptiveCard';
  version: string;
  body: any[];
  actions?: any[];
  msTeams?: {
    width: 'Full';
  };
}

// ============================================================================
// POLICY APPROVAL WORKFLOW SERVICE
// ============================================================================

export class PolicyApprovalWorkflowService {
  private sp: SPFI;
  private context: WebPartContext;
  private siteUrl: string;
  private currentUserId: number = 0;

  // SharePoint List Names
  private readonly WORKFLOW_TEMPLATES_LIST = PolicyWorkflowLists.WORKFLOW_TEMPLATES;
  private readonly WORKFLOW_INSTANCES_LIST = PolicyWorkflowLists.WORKFLOW_INSTANCES;
  private readonly APPROVAL_DECISIONS_LIST = PolicyWorkflowLists.APPROVAL_DECISIONS;
  private readonly DELEGATIONS_LIST = PolicyWorkflowLists.DELEGATIONS;
  private readonly ESCALATION_RULES_LIST = PolicyWorkflowLists.ESCALATION_RULES;
  private readonly POLICIES_LIST = PolicyLists.POLICIES;
  private readonly WORKFLOW_HISTORY_LIST = PolicyWorkflowLists.WORKFLOW_HISTORY;

  constructor(context: WebPartContext) {
    this.context = context;
    this.sp = getSP(context);
    this.siteUrl = context.pageContext.web.absoluteUrl;
  }

  /**
   * Initialize service
   */
  public async initialize(): Promise<void> {
    const user = await this.sp.web.currentUser();
    this.currentUserId = user.Id;
  }

  // ============================================================================
  // WORKFLOW TEMPLATE MANAGEMENT
  // ============================================================================

  /**
   * Create a new approval workflow template
   */
  public async createWorkflowTemplate(
    template: IApprovalWorkflowTemplate
  ): Promise<number> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.WORKFLOW_TEMPLATES_LIST)
        .items.add({
          Title: template.templateName,
          TemplateName: template.templateName,
          TemplateDescription: template.templateDescription,
          Category: template.category,
          StagesJSON: JSON.stringify(template.stages),
          StageCount: template.stages.length,
          IsActive: template.isActive,
          IsDefault: template.isDefault,
          CreatedById: this.currentUserId
        });

      logger.info('PolicyApprovalWorkflowService', `Created workflow template: ${template.templateName}`);
      return result.data.Id;
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to create workflow template:', error);
      throw error;
    }
  }

  /**
   * Get workflow template by ID
   */
  public async getWorkflowTemplate(templateId: number): Promise<IApprovalWorkflowTemplate | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.WORKFLOW_TEMPLATES_LIST)
        .items.getById(templateId)
        .select('*', 'CreatedBy/Title', 'CreatedBy/EMail')
        .expand('CreatedBy')();

      return this.mapToWorkflowTemplate(item);
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to get workflow template:', error);
      return null;
    }
  }

  /**
   * Get all active workflow templates
   */
  public async getWorkflowTemplates(category?: string): Promise<IApprovalWorkflowTemplate[]> {
    try {
      let filter = 'IsActive eq 1';
      if (category) {
        filter += ` and (Category eq '${category}' or Category eq null)`;
      }

      const items = await this.sp.web.lists
        .getByTitle(this.WORKFLOW_TEMPLATES_LIST)
        .items.filter(filter)
        .orderBy('TemplateName')
        .top(100)();

      return items.map(item => this.mapToWorkflowTemplate(item));
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to get workflow templates:', error);
      return [];
    }
  }

  /**
   * Get default template for a category
   */
  public async getDefaultTemplate(category?: string): Promise<IApprovalWorkflowTemplate | null> {
    try {
      let filter = 'IsActive eq 1 and IsDefault eq 1';
      if (category) {
        filter += ` and (Category eq '${category}' or Category eq null)`;
      }

      const items = await this.sp.web.lists
        .getByTitle(this.WORKFLOW_TEMPLATES_LIST)
        .items.filter(filter)
        .top(1)();

      return items.length > 0 ? this.mapToWorkflowTemplate(items[0]) : null;
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to get default template:', error);
      return null;
    }
  }

  // ============================================================================
  // WORKFLOW INSTANCE MANAGEMENT
  // ============================================================================

  /**
   * Start approval workflow for a policy
   */
  public async startWorkflow(
    policyId: number,
    templateId?: number,
    customStages?: IApprovalStage[]
  ): Promise<IApprovalWorkflowInstance> {
    try {
      // Get policy details
      const policy = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(policyId)
        .select('Id', 'PolicyNumber', 'PolicyName', 'PolicyCategory')() as IPolicy;

      // Get or create stages
      let stages: IApprovalStage[];
      let templateName = 'Custom Workflow';

      if (templateId) {
        const template = await this.getWorkflowTemplate(templateId);
        if (!template) throw new Error('Template not found');
        stages = template.stages;
        templateName = template.templateName;
      } else if (customStages) {
        stages = customStages;
      } else {
        // Try to get default template
        const defaultTemplate = await this.getDefaultTemplate(policy.PolicyCategory);
        if (defaultTemplate) {
          stages = defaultTemplate.stages;
          templateName = defaultTemplate.templateName;
          templateId = defaultTemplate.id;
        } else {
          throw new Error('No workflow template specified and no default found');
        }
      }

      // Create workflow instance
      const firstStage = stages[0];
      const instanceResult = await this.sp.web.lists
        .getByTitle(this.WORKFLOW_INSTANCES_LIST)
        .items.add({
          Title: `${policy.PolicyNumber} - ${templateName}`,
          PolicyId: policyId,
          PolicyNumber: policy.PolicyNumber,
          PolicyName: policy.PolicyName,
          TemplateId: templateId,
          TemplateName: templateName,
          CurrentStageId: firstStage.stageId,
          CurrentStageOrder: 1,
          OverallStatus: PolicyApprovalStatus.PendingReview,
          StagesJSON: JSON.stringify(stages),
          StageCount: stages.length,
          InitiatedById: this.currentUserId,
          InitiatedDate: new Date().toISOString()
        });

      const workflowInstanceId = instanceResult.data.Id;

      // Create approval decisions for first stage
      await this.createStageApprovals(workflowInstanceId, firstStage, policyId);

      // Update policy status
      await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(policyId)
        .update({
          Status: PolicyStatus.InReview,
          WorkflowInstanceId: workflowInstanceId
        });

      // Log to history
      await this.logWorkflowHistory(workflowInstanceId, 'Started', `Workflow started: ${templateName}`);

      // Send notifications
      await this.sendStageNotifications(workflowInstanceId, firstStage);

      logger.info('PolicyApprovalWorkflowService', `Started workflow for policy ${policyId}`);

      return this.getWorkflowInstance(workflowInstanceId) as Promise<IApprovalWorkflowInstance>;
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to start workflow:', error);
      throw error;
    }
  }

  /**
   * Create approval decisions for a stage
   */
  private async createStageApprovals(
    workflowInstanceId: number,
    stage: IApprovalStage,
    policyId: number
  ): Promise<void> {
    const dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + stage.dueDays);

    for (const approverId of stage.approverIds) {
      // Check for active delegation
      const delegation = await this.getActiveDelegation(approverId);
      const actualApproverId = delegation ? delegation.delegateId : approverId;

      await this.sp.web.lists
        .getByTitle(this.APPROVAL_DECISIONS_LIST)
        .items.add({
          Title: `${stage.stageName} - Approver ${approverId}`,
          WorkflowInstanceId: workflowInstanceId,
          PolicyId: policyId,
          StageId: stage.stageId,
          StageName: stage.stageName,
          StageOrder: stage.stageOrder,
          ApproverId: actualApproverId,
          OriginalApproverId: delegation ? approverId : null,
          DelegatedById: delegation ? delegation.delegatorId : null,
          Status: PolicyApprovalStatus.PendingApproval,
          Decision: 'Pending',
          RequestedDate: new Date().toISOString(),
          DueDate: dueDate.toISOString(),
          IsOverdue: false,
          EscalationLevel: 0,
          NotificationsSent: 0,
          RequireComments: stage.requireComments
        });
    }
  }

  /**
   * Get workflow instance by ID
   */
  public async getWorkflowInstance(instanceId: number): Promise<IApprovalWorkflowInstance | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.WORKFLOW_INSTANCES_LIST)
        .items.getById(instanceId)
        .select('*', 'InitiatedBy/Title', 'InitiatedBy/EMail')
        .expand('InitiatedBy')();

      // Get all approval decisions for this instance
      const decisions = await this.sp.web.lists
        .getByTitle(this.APPROVAL_DECISIONS_LIST)
        .items.filter(`WorkflowInstanceId eq ${instanceId}`)
        .orderBy('StageOrder')
        .orderBy('Id')
        .top(500)();

      return this.mapToWorkflowInstance(item, decisions);
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to get workflow instance:', error);
      return null;
    }
  }

  /**
   * Submit approval decision
   */
  public async submitDecision(
    decisionId: number,
    approved: boolean,
    comments?: string
  ): Promise<void> {
    try {
      // Get the decision record
      const decision = await this.sp.web.lists
        .getByTitle(this.APPROVAL_DECISIONS_LIST)
        .items.getById(decisionId)();

      // Update decision
      await this.sp.web.lists
        .getByTitle(this.APPROVAL_DECISIONS_LIST)
        .items.getById(decisionId)
        .update({
          Status: approved ? PolicyApprovalStatus.Approved : PolicyApprovalStatus.Rejected,
          Decision: approved ? 'Approved' : 'Rejected',
          Comments: comments,
          RespondedDate: new Date().toISOString(),
          RespondedById: this.currentUserId
        });

      // Log to history
      await this.logWorkflowHistory(
        decision.WorkflowInstanceId,
        approved ? 'Approved' : 'Rejected',
        `Stage "${decision.StageName}" ${approved ? 'approved' : 'rejected'} by user ${this.currentUserId}${comments ? `: ${comments}` : ''}`
      );

      // Process stage completion
      await this.processStageCompletion(decision.WorkflowInstanceId, decision.StageId);

      logger.info('PolicyApprovalWorkflowService', `Decision submitted for ${decisionId}: ${approved ? 'Approved' : 'Rejected'}`);
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to submit decision:', error);
      throw error;
    }
  }

  /**
   * Process stage completion and advance workflow
   */
  private async processStageCompletion(
    workflowInstanceId: number,
    stageId: string
  ): Promise<void> {
    try {
      const instance = await this.getWorkflowInstance(workflowInstanceId);
      if (!instance) return;

      const currentStage = instance.stages.find(s => s.stageId === stageId);
      if (!currentStage) return;

      // Get all decisions for this stage
      const stageDecisions = currentStage.approvals;

      // Check if stage is complete based on approval rule
      const isComplete = this.isStageComplete(currentStage, stageDecisions);

      if (!isComplete) return;

      // Determine stage outcome
      const isApproved = this.isStageApproved(currentStage, stageDecisions);

      if (!isApproved) {
        // Stage rejected - end workflow
        await this.rejectWorkflow(workflowInstanceId, 'Stage rejected');
        return;
      }

      // Find next stage
      const stages: IApprovalStage[] = JSON.parse(
        (await this.sp.web.lists
          .getByTitle(this.WORKFLOW_INSTANCES_LIST)
          .items.getById(workflowInstanceId)
          .select('StagesJSON')()).StagesJSON
      );

      const currentIndex = stages.findIndex(s => s.stageId === stageId);
      const nextStage = stages[currentIndex + 1];

      if (nextStage) {
        // Advance to next stage
        await this.advanceToNextStage(workflowInstanceId, nextStage, instance.policyId);
      } else {
        // All stages complete - approve workflow
        await this.approveWorkflow(workflowInstanceId);
      }
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to process stage completion:', error);
      throw error;
    }
  }

  /**
   * Check if stage is complete based on approval rule
   */
  private isStageComplete(stage: IApprovalStageInstance, decisions: IApprovalDecision[]): boolean {
    const pendingCount = decisions.filter(d => d.decision === 'Pending').length;
    const approvedCount = decisions.filter(d => d.decision === 'Approved').length;
    const rejectedCount = decisions.filter(d => d.decision === 'Rejected').length;
    const totalCount = decisions.length;

    switch (stage.approvalRule) {
      case ApprovalRuleType.AllMustApprove:
        return pendingCount === 0; // All must respond

      case ApprovalRuleType.AnyOneApproves:
        return approvedCount > 0 || rejectedCount > 0; // Any response completes

      case ApprovalRuleType.MajorityApproves:
        const majority = Math.ceil(totalCount / 2);
        return approvedCount >= majority || rejectedCount >= majority;

      case ApprovalRuleType.QuorumApproves:
        // Would need quorum percentage from stage config
        const quorum = Math.ceil(totalCount * 0.6); // Default 60%
        return (approvedCount + rejectedCount) >= quorum;

      default:
        return pendingCount === 0;
    }
  }

  /**
   * Check if stage is approved based on approval rule
   */
  private isStageApproved(stage: IApprovalStageInstance, decisions: IApprovalDecision[]): boolean {
    const approvedCount = decisions.filter(d => d.decision === 'Approved').length;
    const rejectedCount = decisions.filter(d => d.decision === 'Rejected').length;
    const totalCount = decisions.length;

    switch (stage.approvalRule) {
      case ApprovalRuleType.AllMustApprove:
        return rejectedCount === 0 && approvedCount === totalCount;

      case ApprovalRuleType.AnyOneApproves:
        return approvedCount > 0;

      case ApprovalRuleType.MajorityApproves:
        return approvedCount > rejectedCount;

      case ApprovalRuleType.QuorumApproves:
        return approvedCount > rejectedCount;

      default:
        return approvedCount === totalCount;
    }
  }

  /**
   * Advance workflow to next stage
   */
  private async advanceToNextStage(
    workflowInstanceId: number,
    nextStage: IApprovalStage,
    policyId: number
  ): Promise<void> {
    // Update workflow instance
    await this.sp.web.lists
      .getByTitle(this.WORKFLOW_INSTANCES_LIST)
      .items.getById(workflowInstanceId)
      .update({
        CurrentStageId: nextStage.stageId,
        CurrentStageOrder: nextStage.stageOrder,
        OverallStatus: nextStage.stageType === ApprovalStageType.Review
          ? PolicyApprovalStatus.PendingReview
          : PolicyApprovalStatus.PendingApproval
      });

    // Create approvals for next stage
    await this.createStageApprovals(workflowInstanceId, nextStage, policyId);

    // Log to history
    await this.logWorkflowHistory(
      workflowInstanceId,
      'StageAdvanced',
      `Advanced to stage: ${nextStage.stageName}`
    );

    // Send notifications
    await this.sendStageNotifications(workflowInstanceId, nextStage);
  }

  /**
   * Approve workflow (all stages complete)
   */
  private async approveWorkflow(workflowInstanceId: number): Promise<void> {
    const instance = await this.getWorkflowInstance(workflowInstanceId);
    if (!instance) return;

    // Calculate total duration
    const startDate = new Date(instance.initiatedDate);
    const endDate = new Date();
    const durationDays = Math.ceil((endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24));

    // Update workflow instance
    await this.sp.web.lists
      .getByTitle(this.WORKFLOW_INSTANCES_LIST)
      .items.getById(workflowInstanceId)
      .update({
        OverallStatus: PolicyApprovalStatus.Approved,
        CompletedDate: endDate.toISOString(),
        CompletedById: this.currentUserId,
        TotalDurationDays: durationDays
      });

    // Update policy status
    await this.sp.web.lists
      .getByTitle(this.POLICIES_LIST)
      .items.getById(instance.policyId)
      .update({
        Status: PolicyStatus.Approved,
        ApprovedDate: endDate.toISOString(),
        ApprovedById: this.currentUserId
      });

    // Log to history
    await this.logWorkflowHistory(
      workflowInstanceId,
      'Completed',
      `Workflow completed and policy approved after ${durationDays} days`
    );

    logger.info('PolicyApprovalWorkflowService', `Workflow ${workflowInstanceId} approved`);
  }

  /**
   * Reject workflow
   */
  private async rejectWorkflow(workflowInstanceId: number, reason: string): Promise<void> {
    const instance = await this.getWorkflowInstance(workflowInstanceId);
    if (!instance) return;

    // Update workflow instance
    await this.sp.web.lists
      .getByTitle(this.WORKFLOW_INSTANCES_LIST)
      .items.getById(workflowInstanceId)
      .update({
        OverallStatus: PolicyApprovalStatus.Rejected,
        CompletedDate: new Date().toISOString(),
        Comments: reason
      });

    // Update policy status
    await this.sp.web.lists
      .getByTitle(this.POLICIES_LIST)
      .items.getById(instance.policyId)
      .update({
        Status: PolicyStatus.Draft,
        RejectedDate: new Date().toISOString(),
        RejectionReason: reason
      });

    // Log to history
    await this.logWorkflowHistory(workflowInstanceId, 'Rejected', reason);

    logger.info('PolicyApprovalWorkflowService', `Workflow ${workflowInstanceId} rejected: ${reason}`);
  }

  // ============================================================================
  // DELEGATION MANAGEMENT
  // ============================================================================

  /**
   * Create approval delegation
   */
  public async createDelegation(delegation: Omit<IApprovalDelegation, 'id' | 'createdDate'>): Promise<number> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.DELEGATIONS_LIST)
        .items.add({
          Title: `${delegation.delegatorName} → ${delegation.delegateName}`,
          DelegatorId: delegation.delegatorId,
          DelegatorName: delegation.delegatorName,
          DelegatorEmail: delegation.delegatorEmail,
          DelegateId: delegation.delegateId,
          DelegateName: delegation.delegateName,
          DelegateEmail: delegation.delegateEmail,
          DelegationType: delegation.delegationType,
          StartDate: delegation.startDate.toISOString(),
          EndDate: delegation.endDate?.toISOString(),
          Reason: delegation.reason,
          PolicyCategories: delegation.policyCategories?.join(';'),
          IsActive: delegation.isActive,
          CreatedDate: new Date().toISOString()
        });

      logger.info('PolicyApprovalWorkflowService', `Created delegation: ${delegation.delegatorId} → ${delegation.delegateId}`);
      return result.data.Id;
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to create delegation:', error);
      throw error;
    }
  }

  /**
   * Get active delegation for a user
   */
  public async getActiveDelegation(userId: number): Promise<IApprovalDelegation | null> {
    try {
      const now = new Date();
      const filter = `DelegatorId eq ${userId} and IsActive eq 1 and StartDate le datetime'${now.toISOString()}' and (EndDate ge datetime'${now.toISOString()}' or EndDate eq null)`;

      const items = await this.sp.web.lists
        .getByTitle(this.DELEGATIONS_LIST)
        .items.filter(filter)
        .top(1)();

      return items.length > 0 ? this.mapToDelegation(items[0]) : null;
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to get delegation:', error);
      return null;
    }
  }

  /**
   * Get delegations for a user (both delegating and receiving)
   */
  public async getUserDelegations(userId: number): Promise<{
    outgoing: IApprovalDelegation[];
    incoming: IApprovalDelegation[];
  }> {
    try {
      const [outgoing, incoming] = await Promise.all([
        this.sp.web.lists
          .getByTitle(this.DELEGATIONS_LIST)
          .items.filter(`DelegatorId eq ${userId}`)
          .orderBy('StartDate', false)
          .top(50)(),
        this.sp.web.lists
          .getByTitle(this.DELEGATIONS_LIST)
          .items.filter(`DelegateId eq ${userId}`)
          .orderBy('StartDate', false)
          .top(50)()
      ]);

      return {
        outgoing: outgoing.map(item => this.mapToDelegation(item)),
        incoming: incoming.map(item => this.mapToDelegation(item))
      };
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to get user delegations:', error);
      return { outgoing: [], incoming: [] };
    }
  }

  /**
   * Revoke delegation
   */
  public async revokeDelegation(delegationId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.DELEGATIONS_LIST)
        .items.getById(delegationId)
        .update({
          IsActive: false,
          EndDate: new Date().toISOString()
        });

      logger.info('PolicyApprovalWorkflowService', `Revoked delegation: ${delegationId}`);
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to revoke delegation:', error);
      throw error;
    }
  }

  // ============================================================================
  // ESCALATION ENGINE
  // ============================================================================

  /**
   * Create escalation rule
   */
  public async createEscalationRule(rule: Omit<IEscalationRule, 'id'>): Promise<number> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.ESCALATION_RULES_LIST)
        .items.add({
          Title: rule.ruleName,
          RuleName: rule.ruleName,
          RuleDescription: rule.ruleDescription,
          TriggerDays: rule.triggerDays,
          TriggerCondition: rule.triggerCondition,
          Action: rule.action,
          TargetType: rule.targetType,
          TargetIds: rule.targetIds?.join(';'),
          TargetRole: rule.targetRole,
          NotificationTemplate: rule.notificationTemplate,
          MaxEscalations: rule.maxEscalations,
          EscalationIntervalDays: rule.escalationIntervalDays,
          IsActive: rule.isActive,
          AppliesToCategories: rule.appliesToCategories?.join(';'),
          Priority: rule.priority
        });

      logger.info('PolicyApprovalWorkflowService', `Created escalation rule: ${rule.ruleName}`);
      return result.data.Id;
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to create escalation rule:', error);
      throw error;
    }
  }

  /**
   * Process escalations for all pending approvals
   */
  public async processEscalations(): Promise<{
    processed: number;
    escalated: number;
    autoApproved: number;
    autoRejected: number;
  }> {
    const stats = { processed: 0, escalated: 0, autoApproved: 0, autoRejected: 0 };

    try {
      // Get active escalation rules
      const rules = await this.sp.web.lists
        .getByTitle(this.ESCALATION_RULES_LIST)
        .items.filter('IsActive eq 1')
        .orderBy('Priority')
        .top(100)();

      // Get pending approvals that are overdue
      const now = new Date();
      const pendingApprovals = await this.sp.web.lists
        .getByTitle(this.APPROVAL_DECISIONS_LIST)
        .items.filter(`Decision eq 'Pending' and DueDate lt datetime'${now.toISOString()}'`)
        .top(500)();

      for (const approval of pendingApprovals) {
        stats.processed++;

        const dueDate = new Date(approval.DueDate);
        const daysOverdue = Math.ceil((now.getTime() - dueDate.getTime()) / (1000 * 60 * 60 * 24));

        // Find applicable rule
        const applicableRule = rules.find((r: any) => {
          if (r.TriggerDays > daysOverdue) return false;
          if (approval.EscalationLevel >= r.MaxEscalations) return false;
          return true;
        });

        if (!applicableRule) continue;

        // Check escalation interval
        if (approval.LastEscalationDate) {
          const lastEscalation = new Date(approval.LastEscalationDate);
          const daysSinceEscalation = Math.ceil((now.getTime() - lastEscalation.getTime()) / (1000 * 60 * 60 * 24));
          if (daysSinceEscalation < applicableRule.EscalationIntervalDays) continue;
        }

        // Execute escalation action
        await this.executeEscalationAction(approval, applicableRule, stats);
      }

      logger.info('PolicyApprovalWorkflowService', 'Processed escalations', stats);
      return stats;
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to process escalations:', error);
      return stats;
    }
  }

  /**
   * Execute escalation action
   */
  private async executeEscalationAction(
    approval: any,
    rule: any,
    stats: { escalated: number; autoApproved: number; autoRejected: number }
  ): Promise<void> {
    const action = rule.Action as EscalationActionType;

    switch (action) {
      case EscalationActionType.Notify:
        // Send reminder notification
        await this.sendEscalationNotification(approval, rule);
        stats.escalated++;
        break;

      case EscalationActionType.NotifyManager:
        // Get approver's manager and notify
        await this.sendManagerEscalation(approval, rule);
        stats.escalated++;
        break;

      case EscalationActionType.AutoApprove:
        // Auto-approve the decision
        await this.sp.web.lists
          .getByTitle(this.APPROVAL_DECISIONS_LIST)
          .items.getById(approval.Id)
          .update({
            Status: PolicyApprovalStatus.Approved,
            Decision: 'Approved',
            Comments: `Auto-approved after ${rule.TriggerDays} days overdue`,
            RespondedDate: new Date().toISOString()
          });
        await this.processStageCompletion(approval.WorkflowInstanceId, approval.StageId);
        stats.autoApproved++;
        break;

      case EscalationActionType.AutoReject:
        // Auto-reject the decision
        await this.sp.web.lists
          .getByTitle(this.APPROVAL_DECISIONS_LIST)
          .items.getById(approval.Id)
          .update({
            Status: PolicyApprovalStatus.Rejected,
            Decision: 'Rejected',
            Comments: `Auto-rejected after ${rule.TriggerDays} days overdue`,
            RespondedDate: new Date().toISOString()
          });
        await this.processStageCompletion(approval.WorkflowInstanceId, approval.StageId);
        stats.autoRejected++;
        break;

      case EscalationActionType.Reassign:
        // Reassign to backup approver
        await this.reassignApproval(approval, rule);
        stats.escalated++;
        break;
    }

    // Update escalation level
    await this.sp.web.lists
      .getByTitle(this.APPROVAL_DECISIONS_LIST)
      .items.getById(approval.Id)
      .update({
        EscalationLevel: approval.EscalationLevel + 1,
        LastEscalationDate: new Date().toISOString(),
        IsOverdue: true
      });
  }

  /**
   * Send escalation notification
   */
  private async sendEscalationNotification(approval: any, rule: any): Promise<void> {
    // Get approver details
    const approver = await this.sp.web.siteUsers.getById(approval.ApproverId)();

    // Send email notification
    await this.sp.utility.sendEmail({
      To: [approver.Email],
      Subject: `⚠️ Escalation: Policy Approval Overdue - ${approval.StageName}`,
      Body: this.buildEscalationEmailBody(approval, rule)
    });

    // Update notification count
    await this.sp.web.lists
      .getByTitle(this.APPROVAL_DECISIONS_LIST)
      .items.getById(approval.Id)
      .update({
        NotificationsSent: approval.NotificationsSent + 1,
        LastNotificationDate: new Date().toISOString()
      });
  }

  /**
   * Send manager escalation
   */
  private async sendManagerEscalation(approval: any, rule: any): Promise<void> {
    // This would typically integrate with Azure AD to get manager
    // For now, send to configured targets
    const targetIds = rule.TargetIds?.split(';').map((id: string) => parseInt(id, 10)) || [];

    for (const targetId of targetIds) {
      const target = await this.sp.web.siteUsers.getById(targetId)();

      await this.sp.utility.sendEmail({
        To: [target.Email],
        Subject: `🔔 Manager Alert: Policy Approval Escalation`,
        Body: this.buildManagerEscalationEmailBody(approval, rule)
      });
    }
  }

  /**
   * Reassign approval to backup
   */
  private async reassignApproval(approval: any, rule: any): Promise<void> {
    const targetIds = rule.TargetIds?.split(';').map((id: string) => parseInt(id, 10)) || [];
    if (targetIds.length === 0) return;

    const newApproverId = targetIds[0];
    const newApprover = await this.sp.web.siteUsers.getById(newApproverId)();

    await this.sp.web.lists
      .getByTitle(this.APPROVAL_DECISIONS_LIST)
      .items.getById(approval.Id)
      .update({
        OriginalApproverId: approval.ApproverId,
        ApproverId: newApproverId,
        Comments: `Reassigned from user ${approval.ApproverId} due to escalation`
      });

    // Notify new approver
    await this.sp.utility.sendEmail({
      To: [newApprover.Email],
      Subject: `📋 Policy Approval Reassigned to You`,
      Body: this.buildReassignmentEmailBody(approval, newApprover)
    });
  }

  // ============================================================================
  // TEAMS ADAPTIVE CARD NOTIFICATIONS
  // ============================================================================

  /**
   * Send Teams adaptive card for approval request
   */
  public async sendTeamsApprovalCard(
    approval: IApprovalDecision,
    policy: IPolicy,
    webhookUrl: string
  ): Promise<void> {
    try {
      const card = this.buildApprovalAdaptiveCard(approval, policy);

      const response = await fetch(webhookUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          type: 'message',
          attachments: [{
            contentType: 'application/vnd.microsoft.card.adaptive',
            contentUrl: null,
            content: card
          }]
        })
      });

      if (!response.ok) {
        throw new Error(`Teams webhook failed: ${response.statusText}`);
      }

      logger.info('PolicyApprovalWorkflowService', 'Sent Teams adaptive card notification');
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to send Teams card:', error);
      throw error;
    }
  }

  /**
   * Build approval request adaptive card
   */
  public buildApprovalAdaptiveCard(
    approval: IApprovalDecision,
    policy: IPolicy
  ): ITeamsAdaptiveCardPayload {
    const approvalUrl = `${this.siteUrl}/SitePages/PolicyApproval.aspx?decisionId=${approval.id}`;
    const policyUrl = `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`;

    return {
      type: 'AdaptiveCard',
      version: '1.4',
      msTeams: { width: 'Full' },
      body: [
        {
          type: 'Container',
          style: 'emphasis',
          items: [
            {
              type: 'ColumnSet',
              columns: [
                {
                  type: 'Column',
                  width: 'auto',
                  items: [
                    {
                      type: 'Image',
                      url: 'https://adaptivecards.io/content/pending.png',
                      size: 'Small'
                    }
                  ]
                },
                {
                  type: 'Column',
                  width: 'stretch',
                  items: [
                    {
                      type: 'TextBlock',
                      text: '📋 Policy Approval Request',
                      weight: 'Bolder',
                      size: 'Medium',
                      wrap: true
                    },
                    {
                      type: 'TextBlock',
                      text: `You have a pending policy approval`,
                      isSubtle: true,
                      spacing: 'None',
                      wrap: true
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          type: 'Container',
          items: [
            {
              type: 'FactSet',
              facts: [
                { title: 'Policy Number:', value: policy.PolicyNumber || 'N/A' },
                { title: 'Policy Name:', value: policy.PolicyName },
                { title: 'Category:', value: policy.PolicyCategory || 'General' },
                { title: 'Stage:', value: approval.stageId || 'Review' },
                { title: 'Due Date:', value: new Date(approval.dueDate).toLocaleDateString() }
              ]
            }
          ]
        },
        {
          type: 'Container',
          items: [
            {
              type: 'TextBlock',
              text: policy.Description || 'No description provided.',
              wrap: true,
              maxLines: 3
            }
          ]
        },
        {
          type: 'Container',
          style: approval.isOverdue ? 'attention' : 'default',
          items: approval.isOverdue ? [
            {
              type: 'TextBlock',
              text: '⚠️ This approval is overdue!',
              weight: 'Bolder',
              color: 'Attention',
              wrap: true
            }
          ] : []
        }
      ],
      actions: [
        {
          type: 'Action.OpenUrl',
          title: '✅ Review & Approve',
          url: approvalUrl,
          style: 'positive'
        },
        {
          type: 'Action.OpenUrl',
          title: '📄 View Policy',
          url: policyUrl
        },
        {
          type: 'Action.OpenUrl',
          title: '❌ Reject',
          url: `${approvalUrl}&action=reject`
        }
      ]
    };
  }

  /**
   * Build workflow status adaptive card
   */
  public buildWorkflowStatusCard(workflow: IApprovalWorkflowInstance): ITeamsAdaptiveCardPayload {
    const statusColors: Record<PolicyApprovalStatus, string> = {
      [PolicyApprovalStatus.Draft]: 'Default',
      [PolicyApprovalStatus.PendingReview]: 'Warning',
      [PolicyApprovalStatus.PendingApproval]: 'Warning',
      [PolicyApprovalStatus.Approved]: 'Good',
      [PolicyApprovalStatus.Rejected]: 'Attention',
      [PolicyApprovalStatus.Recalled]: 'Accent',
      [PolicyApprovalStatus.Expired]: 'Attention'
    };

    const stageItems = workflow.stages.map(stage => ({
      type: 'ColumnSet',
      columns: [
        {
          type: 'Column',
          width: 'auto',
          items: [{
            type: 'TextBlock',
            text: stage.status === PolicyApprovalStatus.Approved ? '✅' :
                  stage.status === PolicyApprovalStatus.Rejected ? '❌' :
                  stage.status === PolicyApprovalStatus.PendingApproval ? '⏳' : '⬜',
            size: 'Medium'
          }]
        },
        {
          type: 'Column',
          width: 'stretch',
          items: [
            {
              type: 'TextBlock',
              text: stage.stageName,
              weight: 'Bolder'
            },
            {
              type: 'TextBlock',
              text: stage.status,
              isSubtle: true,
              spacing: 'None'
            }
          ]
        }
      ]
    }));

    return {
      type: 'AdaptiveCard',
      version: '1.4',
      msTeams: { width: 'Full' },
      body: [
        {
          type: 'Container',
          style: statusColors[workflow.overallStatus] === 'Good' ? 'good' :
                 statusColors[workflow.overallStatus] === 'Attention' ? 'attention' : 'emphasis',
          items: [
            {
              type: 'TextBlock',
              text: `📋 Workflow Status: ${workflow.overallStatus}`,
              weight: 'Bolder',
              size: 'Medium',
              wrap: true
            }
          ]
        },
        {
          type: 'FactSet',
          facts: [
            { title: 'Policy:', value: `${workflow.policyNumber} - ${workflow.policyName}` },
            { title: 'Template:', value: workflow.templateName || 'Custom' },
            { title: 'Current Stage:', value: `${workflow.currentStageOrder} of ${workflow.stages.length}` },
            { title: 'Initiated:', value: new Date(workflow.initiatedDate).toLocaleDateString() }
          ]
        },
        {
          type: 'Container',
          items: [
            {
              type: 'TextBlock',
              text: 'Workflow Stages',
              weight: 'Bolder',
              spacing: 'Medium'
            },
            ...stageItems
          ]
        }
      ],
      actions: [
        {
          type: 'Action.OpenUrl',
          title: 'View Details',
          url: `${this.siteUrl}/SitePages/PolicyWorkflow.aspx?instanceId=${workflow.id}`
        }
      ]
    };
  }

  // ============================================================================
  // NOTIFICATION HELPERS
  // ============================================================================

  /**
   * Send stage notifications to approvers
   */
  private async sendStageNotifications(
    workflowInstanceId: number,
    stage: IApprovalStage
  ): Promise<void> {
    try {
      // Get instance details
      const instance = await this.getWorkflowInstance(workflowInstanceId);
      if (!instance) return;

      // Get policy details
      const policy = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(instance.policyId)() as IPolicy;

      // Get approvers
      for (const approverId of stage.approverIds) {
        const approver = await this.sp.web.siteUsers.getById(approverId)();

        // Send email
        await this.sp.utility.sendEmail({
          To: [approver.Email],
          Subject: `📋 Policy Approval Required: ${policy.PolicyNumber} - ${stage.stageName}`,
          Body: this.buildApprovalRequestEmailBody(policy, stage, approver)
        });
      }

      logger.info('PolicyApprovalWorkflowService', `Sent notifications for stage: ${stage.stageName}`);
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to send stage notifications:', error);
    }
  }

  /**
   * Build approval request email body
   */
  private buildApprovalRequestEmailBody(policy: IPolicy, stage: IApprovalStage, approver: any): string {
    const approvalUrl = `${this.siteUrl}/SitePages/PolicyApproval.aspx?policyId=${policy.Id}`;

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #0078d4 0%, #004578 100%); padding: 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: white; margin: 0;">📋 Policy Approval Request</h2>
          </div>

          <div style="padding: 24px; background: #f9f9f9; border-radius: 0 0 8px 8px;">
            <p>Hello ${approver.Title},</p>

            <p>You have been assigned to review and approve the following policy:</p>

            <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background: white; border-radius: 8px; overflow: hidden;">
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600; width: 140px;">Policy Number:</td>
                <td style="padding: 12px 16px;">${policy.PolicyNumber || 'N/A'}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Policy Name:</td>
                <td style="padding: 12px 16px;">${policy.PolicyName}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Category:</td>
                <td style="padding: 12px 16px;">${policy.PolicyCategory || 'General'}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Stage:</td>
                <td style="padding: 12px 16px;">${stage.stageName} (${stage.stageType})</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Due In:</td>
                <td style="padding: 12px 16px;">${stage.dueDays} days</td>
              </tr>
            </table>

            ${stage.instructions ? `
              <div style="background: #fff4ce; padding: 16px; border-radius: 8px; margin: 20px 0;">
                <strong>Instructions:</strong>
                <p style="margin: 8px 0 0 0;">${stage.instructions}</p>
              </div>
            ` : ''}

            <p style="text-align: center; margin: 24px 0;">
              <a href="${approvalUrl}"
                 style="background: #0078d4; color: white; padding: 14px 32px;
                        text-decoration: none; border-radius: 6px; display: inline-block;
                        font-weight: 600;">
                Review & Approve
              </a>
            </p>

            <p style="color: #605e5c; font-size: 12px; margin-top: 24px;">
              This is an automated notification from the Policy Management System.
            </p>
          </div>
        </body>
      </html>
    `;
  }

  /**
   * Build escalation email body
   */
  private buildEscalationEmailBody(approval: any, rule: any): string {
    const approvalUrl = `${this.siteUrl}/SitePages/PolicyApproval.aspx?decisionId=${approval.Id}`;

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #f7630c 0%, #d13438 100%); padding: 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: white; margin: 0;">⚠️ Escalation: Approval Overdue</h2>
          </div>

          <div style="padding: 24px; background: #f9f9f9; border-radius: 0 0 8px 8px;">
            <p>This is an escalation reminder for a policy approval that is overdue.</p>

            <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background: white; border-radius: 8px; overflow: hidden;">
              <tr>
                <td style="padding: 12px 16px; background: #fde7e9; font-weight: 600; width: 140px;">Stage:</td>
                <td style="padding: 12px 16px;">${approval.StageName}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #fde7e9; font-weight: 600;">Due Date:</td>
                <td style="padding: 12px 16px; color: #d13438; font-weight: 600;">${new Date(approval.DueDate).toLocaleDateString()}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #fde7e9; font-weight: 600;">Escalation Level:</td>
                <td style="padding: 12px 16px;">${approval.EscalationLevel + 1} of ${rule.MaxEscalations}</td>
              </tr>
            </table>

            <div style="background: #fff4ce; padding: 16px; border-radius: 8px; margin: 20px 0;">
              <strong>⚡ Action Required:</strong>
              <p style="margin: 8px 0 0 0;">Please review and respond to this approval request immediately.</p>
            </div>

            <p style="text-align: center; margin: 24px 0;">
              <a href="${approvalUrl}"
                 style="background: #d13438; color: white; padding: 14px 32px;
                        text-decoration: none; border-radius: 6px; display: inline-block;
                        font-weight: 600;">
                Respond Now
              </a>
            </p>
          </div>
        </body>
      </html>
    `;
  }

  /**
   * Build manager escalation email body
   */
  private buildManagerEscalationEmailBody(approval: any, rule: any): string {
    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #f7630c 0%, #d13438 100%); padding: 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: white; margin: 0;">🔔 Manager Alert: Approval Escalation</h2>
          </div>

          <div style="padding: 24px; background: #f9f9f9; border-radius: 0 0 8px 8px;">
            <p>A policy approval assigned to one of your team members is overdue and has been escalated.</p>

            <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background: white; border-radius: 8px; overflow: hidden;">
              <tr>
                <td style="padding: 12px 16px; background: #fde7e9; font-weight: 600; width: 140px;">Approver:</td>
                <td style="padding: 12px 16px;">User ID: ${approval.ApproverId}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #fde7e9; font-weight: 600;">Stage:</td>
                <td style="padding: 12px 16px;">${approval.StageName}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #fde7e9; font-weight: 600;">Due Date:</td>
                <td style="padding: 12px 16px; color: #d13438; font-weight: 600;">${new Date(approval.DueDate).toLocaleDateString()}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #fde7e9; font-weight: 600;">Days Overdue:</td>
                <td style="padding: 12px 16px;">${rule.TriggerDays}+</td>
              </tr>
            </table>

            <p>Please follow up with the approver to ensure timely response.</p>
          </div>
        </body>
      </html>
    `;
  }

  /**
   * Build reassignment email body
   */
  private buildReassignmentEmailBody(approval: any, newApprover: any): string {
    const approvalUrl = `${this.siteUrl}/SitePages/PolicyApproval.aspx?decisionId=${approval.Id}`;

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #0078d4 0%, #004578 100%); padding: 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: white; margin: 0;">📋 Policy Approval Reassigned to You</h2>
          </div>

          <div style="padding: 24px; background: #f9f9f9; border-radius: 0 0 8px 8px;">
            <p>Hello ${newApprover.Title},</p>

            <p>A policy approval has been reassigned to you due to escalation.</p>

            <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background: white; border-radius: 8px; overflow: hidden;">
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600; width: 140px;">Stage:</td>
                <td style="padding: 12px 16px;">${approval.StageName}</td>
              </tr>
              <tr>
                <td style="padding: 12px 16px; background: #f3f2f1; font-weight: 600;">Original Approver:</td>
                <td style="padding: 12px 16px;">User ID: ${approval.ApproverId}</td>
              </tr>
            </table>

            <p style="text-align: center; margin: 24px 0;">
              <a href="${approvalUrl}"
                 style="background: #0078d4; color: white; padding: 14px 32px;
                        text-decoration: none; border-radius: 6px; display: inline-block;
                        font-weight: 600;">
                Review & Approve
              </a>
            </p>
          </div>
        </body>
      </html>
    `;
  }

  // ============================================================================
  // WORKFLOW HISTORY
  // ============================================================================

  /**
   * Log workflow history entry
   */
  private async logWorkflowHistory(
    workflowInstanceId: number,
    action: string,
    details: string
  ): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.WORKFLOW_HISTORY_LIST)
        .items.add({
          Title: `${action} - ${new Date().toISOString()}`,
          WorkflowInstanceId: workflowInstanceId,
          Action: action,
          Details: details,
          PerformedById: this.currentUserId,
          PerformedDate: new Date().toISOString()
        });
    } catch (error) {
      logger.error('PolicyApprovalWorkflowService', 'Failed to log workflow history:', error);
    }
  }

  // ============================================================================
  // MAPPING HELPERS
  // ============================================================================

  private mapToWorkflowTemplate(item: any): IApprovalWorkflowTemplate {
    return {
      id: item.Id,
      templateName: item.TemplateName,
      templateDescription: item.TemplateDescription,
      category: item.Category,
      stages: JSON.parse(item.StagesJSON || '[]'),
      isActive: item.IsActive,
      isDefault: item.IsDefault,
      createdById: item.CreatedById,
      createdDate: item.Created ? new Date(item.Created) : undefined,
      modifiedDate: item.Modified ? new Date(item.Modified) : undefined
    };
  }

  private mapToWorkflowInstance(item: any, decisions: any[]): IApprovalWorkflowInstance {
    const stages: IApprovalStage[] = JSON.parse(item.StagesJSON || '[]');

    const stageInstances: IApprovalStageInstance[] = stages.map(stage => {
      const stageDecisions = decisions.filter(d => d.StageId === stage.stageId);

      return {
        stageId: stage.stageId,
        stageName: stage.stageName,
        stageType: stage.stageType,
        stageOrder: stage.stageOrder,
        status: this.calculateStageStatus(stageDecisions),
        approvalRule: stage.approvalRule,
        approvals: stageDecisions.map(d => this.mapToApprovalDecision(d)),
        startedDate: stageDecisions.length > 0 ? new Date(stageDecisions[0].RequestedDate) : undefined,
        completedDate: this.getStageCompletedDate(stageDecisions),
        dueDate: stageDecisions.length > 0 ? new Date(stageDecisions[0].DueDate) : undefined,
        escalationLevel: Math.max(...stageDecisions.map(d => d.EscalationLevel || 0), 0),
        lastEscalationDate: undefined
      };
    });

    return {
      id: item.Id,
      policyId: item.PolicyId,
      policyNumber: item.PolicyNumber,
      policyName: item.PolicyName,
      templateId: item.TemplateId,
      templateName: item.TemplateName,
      currentStageId: item.CurrentStageId,
      currentStageOrder: item.CurrentStageOrder,
      overallStatus: item.OverallStatus as PolicyApprovalStatus,
      stages: stageInstances,
      initiatedById: item.InitiatedById,
      initiatedByName: item.InitiatedBy?.Title,
      initiatedDate: new Date(item.InitiatedDate),
      completedDate: item.CompletedDate ? new Date(item.CompletedDate) : undefined,
      completedById: item.CompletedById,
      totalDurationDays: item.TotalDurationDays,
      comments: item.Comments
    };
  }

  private mapToApprovalDecision(item: any): IApprovalDecision {
    return {
      id: item.Id,
      workflowInstanceId: item.WorkflowInstanceId,
      stageId: item.StageId,
      approverId: item.ApproverId,
      approverName: item.Approver?.Title,
      approverEmail: item.Approver?.EMail,
      originalApproverId: item.OriginalApproverId,
      delegatedById: item.DelegatedById,
      status: item.Status as PolicyApprovalStatus,
      decision: item.Decision,
      comments: item.Comments,
      requestedDate: new Date(item.RequestedDate),
      respondedDate: item.RespondedDate ? new Date(item.RespondedDate) : undefined,
      dueDate: new Date(item.DueDate),
      isOverdue: item.IsOverdue,
      escalationLevel: item.EscalationLevel || 0,
      notificationsSent: item.NotificationsSent || 0,
      lastNotificationDate: item.LastNotificationDate ? new Date(item.LastNotificationDate) : undefined
    };
  }

  private mapToDelegation(item: any): IApprovalDelegation {
    return {
      id: item.Id,
      delegatorId: item.DelegatorId,
      delegatorName: item.DelegatorName,
      delegatorEmail: item.DelegatorEmail,
      delegateId: item.DelegateId,
      delegateName: item.DelegateName,
      delegateEmail: item.DelegateEmail,
      delegationType: item.DelegationType as DelegationType,
      startDate: new Date(item.StartDate),
      endDate: item.EndDate ? new Date(item.EndDate) : undefined,
      reason: item.Reason,
      policyCategories: item.PolicyCategories?.split(';'),
      isActive: item.IsActive,
      createdDate: new Date(item.CreatedDate),
      approvedById: item.ApprovedById,
      approvedDate: item.ApprovedDate ? new Date(item.ApprovedDate) : undefined
    };
  }

  private calculateStageStatus(decisions: any[]): PolicyApprovalStatus {
    if (decisions.length === 0) return PolicyApprovalStatus.Draft;

    const hasRejected = decisions.some(d => d.Decision === 'Rejected');
    if (hasRejected) return PolicyApprovalStatus.Rejected;

    const allApproved = decisions.every(d => d.Decision === 'Approved');
    if (allApproved) return PolicyApprovalStatus.Approved;

    return PolicyApprovalStatus.PendingApproval;
  }

  private getStageCompletedDate(decisions: any[]): Date | undefined {
    const allResponded = decisions.every(d => d.RespondedDate);
    if (!allResponded) return undefined;

    const dates = decisions.map(d => new Date(d.RespondedDate).getTime());
    return new Date(Math.max(...dates));
  }
}
