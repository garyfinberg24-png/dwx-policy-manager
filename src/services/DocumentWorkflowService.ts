// @ts-nocheck
// Document Workflow Service
// Service for managing document workflows and stages

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import {
  IDocumentWorkflow,
  IWorkflowStage,
  WorkflowType,
  WorkflowStatus,
  StageStatus,
  StageType,
  Priority,
  ActivityType,
  ActivitySeverity
} from '../models';
import { logger } from './LoggingService';
import { DocumentRegistryService } from './DocumentRegistryService';

/**
 * Service for Document Workflow operations
 */
export class DocumentWorkflowService {
  private sp: SPFI;
  private registryService: DocumentRegistryService;

  private readonly WORKFLOWS_LIST = 'PM_DocumentWorkflows';
  private readonly STAGES_LIST = 'PM_WorkflowStages';

  constructor(sp: SPFI) {
    this.sp = sp;
    this.registryService = new DocumentRegistryService(sp);
  }

  // ============================================================================
  // WORKFLOW CRUD
  // ============================================================================

  /**
   * Get workflow by ID
   */
  public async getById(id: number): Promise<IDocumentWorkflow | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.WORKFLOWS_LIST)
        .items
        .getById(id)
        .select(this.getWorkflowSelectFields())();

      return this.mapToWorkflow(item);
    } catch (error) {
      logger.error('DocumentWorkflowService', `Failed to get workflow ${id}:`, error);
      return null;
    }
  }

  /**
   * Get workflows for a document
   */
  public async getByDocumentId(documentRegistryId: number): Promise<IDocumentWorkflow[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.WORKFLOWS_LIST)
        .items
        .filter(`DocumentRegistryId eq ${documentRegistryId}`)
        .select(this.getWorkflowSelectFields())
        .orderBy('Created', false)();

      return items.map(this.mapToWorkflow);
    } catch (error) {
      logger.error('DocumentWorkflowService', `Failed to get workflows for document ${documentRegistryId}:`, error);
      return [];
    }
  }

  /**
   * Get user's pending workflows
   */
  public async getUserPendingWorkflows(userId: number): Promise<IDocumentWorkflow[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.WORKFLOWS_LIST)
        .items
        .filter(`CurrentAssigneesId eq ${userId} and (WorkflowStatus eq 'In Progress' or WorkflowStatus eq 'Pending Approval')`)
        .select(this.getWorkflowSelectFields())
        .orderBy('DueDate')();

      return items.map(this.mapToWorkflow);
    } catch (error) {
      logger.error('DocumentWorkflowService', `Failed to get pending workflows for user ${userId}:`, error);
      return [];
    }
  }

  /**
   * Get overdue workflows
   */
  public async getOverdueWorkflows(): Promise<IDocumentWorkflow[]> {
    try {
      const now = new Date().toISOString();
      const items = await this.sp.web.lists
        .getByTitle(this.WORKFLOWS_LIST)
        .items
        .filter(`DueDate lt '${now}' and (WorkflowStatus eq 'In Progress' or WorkflowStatus eq 'Pending Approval')`)
        .select(this.getWorkflowSelectFields())
        .orderBy('DueDate')();

      return items.map(this.mapToWorkflow);
    } catch (error) {
      logger.error('DocumentWorkflowService', 'Failed to get overdue workflows:', error);
      return [];
    }
  }

  /**
   * Get all active workflows (In Progress or Pending Approval)
   */
  public async getActiveWorkflows(): Promise<IDocumentWorkflow[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.WORKFLOWS_LIST)
        .items
        .filter(`WorkflowStatus eq 'In Progress' or WorkflowStatus eq 'Pending Approval'`)
        .select(this.getWorkflowSelectFields())
        .orderBy('DueDate')
        .top(100)();

      return items.map(this.mapToWorkflow);
    } catch (error) {
      logger.error('DocumentWorkflowService', 'Failed to get active workflows:', error);
      return [];
    }
  }

  /**
   * Create a new workflow
   */
  public async createWorkflow(
    documentRegistryId: number,
    workflowType: WorkflowType,
    stages: Partial<IWorkflowStage>[],
    options: {
      priority?: Priority;
      dueDate?: Date;
      initiatedById: number;
      notifyOnCompletion?: boolean;
    }
  ): Promise<IDocumentWorkflow> {
    try {
      // Get document info
      const doc = await this.registryService.getById(documentRegistryId);
      if (!doc) throw new Error('Document not found');

      // Calculate due date if not provided
      const dueDate = options.dueDate || this.calculateDefaultDueDate(stages.length);

      // Create workflow
      const workflowResult = await this.sp.web.lists
        .getByTitle(this.WORKFLOWS_LIST)
        .items
        .add({
          Title: `${workflowType} - ${doc.Title}`,
          DocumentRegistryId: documentRegistryId,
          DocumentTitle: doc.Title,
          WorkflowType: workflowType,
          WorkflowStatus: WorkflowStatus.Draft,
          CurrentStage: 1,
          TotalStages: stages.length,
          DueDate: dueDate.toISOString(),
          Priority: options.priority || Priority.Medium,
          EscalationLevel: 0,
          IsOverdue: false,
          InitiatedById: options.initiatedById,
          RemindersSent: 0,
          NotifyOnCompletion: options.notifyOnCompletion || true,
          AllowDelegation: true,
          AllowReassignment: true,
          RequireComments: false
        });

      const workflowId = workflowResult.data.Id;

      // Create stages
      for (let i = 0; i < stages.length; i++) {
        const stage = stages[i];
        const stageDueDate = this.calculateStageDueDate(dueDate, i, stages.length);

        await this.sp.web.lists
          .getByTitle(this.STAGES_LIST)
          .items
          .add({
            Title: stage.Title || `Stage ${i + 1}`,
            WorkflowId: workflowId,
            StageNumber: i + 1,
            StageType: stage.StageType || StageType.Approval,
            AssigneeType: stage.AssigneeType || 'User',
            AssigneeIds: stage.AssigneeIds ? { results: stage.AssigneeIds } : undefined,
            AssigneeRole: stage.AssigneeRole,
            StageStatus: i === 0 ? StageStatus.Pending : StageStatus.Pending,
            RequiredAction: stage.RequiredAction || 'Approve',
            DueDays: stage.DueDays || 5,
            StageDueDate: stageDueDate.toISOString(),
            StageRemindersSent: 0
          });
      }

      // Log activity
      await this.registryService.logActivity({
        DocumentRegistryId: documentRegistryId,
        DocumentTitle: doc.Title,
        ActivityDocumentId: doc.DocumentId,
        ActivityType: ActivityType.WorkflowStarted,
        ActivityById: options.initiatedById,
        ActivityDate: new Date(),
        ActivitySeverity: ActivitySeverity.Info,
        IsSystemAction: false,
        ActivityDetails: JSON.stringify({ workflowId, workflowType })
      });

      return await this.getById(workflowId) as IDocumentWorkflow;
    } catch (error) {
      logger.error('DocumentWorkflowService', 'Failed to create workflow:', error);
      throw error;
    }
  }

  /**
   * Start a workflow (move from Draft to In Progress)
   */
  public async startWorkflow(workflowId: number): Promise<void> {
    try {
      const workflow = await this.getById(workflowId);
      if (!workflow) throw new Error('Workflow not found');

      if (workflow.WorkflowStatus !== WorkflowStatus.Draft) {
        throw new Error('Workflow is not in Draft status');
      }

      // Get first stage assignees
      const stages = await this.getWorkflowStages(workflowId);
      const firstStage = stages.find(s => s.StageNumber === 1);

      await this.sp.web.lists
        .getByTitle(this.WORKFLOWS_LIST)
        .items
        .getById(workflowId)
        .update({
          WorkflowStatus: WorkflowStatus.InProgress,
          StartedDate: new Date().toISOString(),
          CurrentAssigneesId: firstStage?.AssigneeIds ? { results: firstStage.AssigneeIds } : undefined
        });

      // Update first stage to In Progress
      if (firstStage) {
        await this.sp.web.lists
          .getByTitle(this.STAGES_LIST)
          .items
          .getById(firstStage.Id!)
          .update({
            StageStatus: StageStatus.InProgress,
            StageStartedDate: new Date().toISOString()
          });
      }
    } catch (error) {
      logger.error('DocumentWorkflowService', `Failed to start workflow ${workflowId}:`, error);
      throw error;
    }
  }

  // ============================================================================
  // STAGE OPERATIONS
  // ============================================================================

  /**
   * Get stages for a workflow
   */
  public async getWorkflowStages(workflowId: number): Promise<IWorkflowStage[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.STAGES_LIST)
        .items
        .filter(`WorkflowId eq ${workflowId}`)
        .select(this.getStageSelectFields())
        .orderBy('StageNumber')();

      return items.map(this.mapToStage);
    } catch (error) {
      logger.error('DocumentWorkflowService', `Failed to get stages for workflow ${workflowId}:`, error);
      return [];
    }
  }

  /**
   * Get current stage of a workflow
   */
  public async getCurrentStage(workflowId: number): Promise<IWorkflowStage | null> {
    try {
      const workflow = await this.getById(workflowId);
      if (!workflow) return null;

      const stages = await this.getWorkflowStages(workflowId);
      return stages.find(s => s.StageNumber === workflow.CurrentStage) || null;
    } catch (error) {
      logger.error('DocumentWorkflowService', `Failed to get current stage for workflow ${workflowId}:`, error);
      return null;
    }
  }

  /**
   * Complete a stage
   */
  public async completeStage(
    stageId: number,
    action: 'Approved' | 'Rejected' | 'Completed',
    userId: number,
    comments?: string
  ): Promise<void> {
    try {
      // Get stage info
      const stageItem = await this.sp.web.lists
        .getByTitle(this.STAGES_LIST)
        .items
        .getById(stageId)
        .select('Id', 'WorkflowId', 'StageNumber', 'Title')();

      const workflowId = stageItem.WorkflowId;
      const workflow = await this.getById(workflowId);
      if (!workflow) throw new Error('Workflow not found');

      // Update stage
      await this.sp.web.lists
        .getByTitle(this.STAGES_LIST)
        .items
        .getById(stageId)
        .update({
          StageStatus: action === 'Rejected' ? StageStatus.Rejected : StageStatus.Completed,
          ActionTaken: action,
          StageCompletedDate: new Date().toISOString(),
          CompletedById: userId,
          StageComments: comments
        });

      // Handle workflow progression
      if (action === 'Rejected') {
        // Workflow rejected
        await this.sp.web.lists
          .getByTitle(this.WORKFLOWS_LIST)
          .items
          .getById(workflowId)
          .update({
            WorkflowStatus: WorkflowStatus.Rejected,
            CompletedDate: new Date().toISOString(),
            Outcome: 'Rejected',
            OutcomeComments: comments
          });

        // Log rejection
        await this.registryService.logActivity({
          DocumentRegistryId: workflow.DocumentRegistryId,
          DocumentTitle: workflow.DocumentTitle,
          ActivityType: ActivityType.WorkflowRejected,
          ActivityById: userId,
          ActivityDate: new Date(),
          ActivitySeverity: ActivitySeverity.Warning,
          IsSystemAction: false,
          ActivityDetails: JSON.stringify({ workflowId, stageId, comments })
        });
      } else {
        // Check if there are more stages
        const stages = await this.getWorkflowStages(workflowId);
        const currentStageNumber = stageItem.StageNumber;
        const nextStage = stages.find(s => s.StageNumber === currentStageNumber + 1);

        if (nextStage) {
          // Move to next stage
          await this.sp.web.lists
            .getByTitle(this.WORKFLOWS_LIST)
            .items
            .getById(workflowId)
            .update({
              CurrentStage: nextStage.StageNumber,
              CurrentAssigneesId: nextStage.AssigneeIds ? { results: nextStage.AssigneeIds } : undefined
            });

          // Start next stage
          await this.sp.web.lists
            .getByTitle(this.STAGES_LIST)
            .items
            .getById(nextStage.Id!)
            .update({
              StageStatus: StageStatus.InProgress,
              StageStartedDate: new Date().toISOString()
            });
        } else {
          // Workflow completed
          const finalStatus = action === 'Approved' ? WorkflowStatus.Approved : WorkflowStatus.Completed;
          await this.sp.web.lists
            .getByTitle(this.WORKFLOWS_LIST)
            .items
            .getById(workflowId)
            .update({
              WorkflowStatus: finalStatus,
              CompletedDate: new Date().toISOString(),
              Outcome: action,
              FinalApproverIds: { results: [userId] }
            });

          // Log completion
          await this.registryService.logActivity({
            DocumentRegistryId: workflow.DocumentRegistryId,
            DocumentTitle: workflow.DocumentTitle,
            ActivityType: action === 'Approved' ? ActivityType.WorkflowApproved : ActivityType.WorkflowCompleted,
            ActivityById: userId,
            ActivityDate: new Date(),
            ActivitySeverity: ActivitySeverity.Info,
            IsSystemAction: false,
            ActivityDetails: JSON.stringify({ workflowId })
          });
        }
      }
    } catch (error) {
      logger.error('DocumentWorkflowService', `Failed to complete stage ${stageId}:`, error);
      throw error;
    }
  }

  /**
   * Delegate a stage to another user
   */
  public async delegateStage(
    stageId: number,
    toUserId: number,
    fromUserId: number,
    reason: string
  ): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.STAGES_LIST)
        .items
        .getById(stageId)
        .update({
          AssigneeIds: { results: [toUserId] },
          DelegatedTo: toUserId,
          DelegatedDate: new Date().toISOString(),
          DelegationReason: reason
        });

      // Update workflow current assignees
      const stageItem = await this.sp.web.lists
        .getByTitle(this.STAGES_LIST)
        .items
        .getById(stageId)
        .select('WorkflowId')();

      await this.sp.web.lists
        .getByTitle(this.WORKFLOWS_LIST)
        .items
        .getById(stageItem.WorkflowId)
        .update({
          CurrentAssigneesId: { results: [toUserId] }
        });
    } catch (error) {
      logger.error('DocumentWorkflowService', `Failed to delegate stage ${stageId}:`, error);
      throw error;
    }
  }

  /**
   * Skip a stage
   */
  public async skipStage(stageId: number, userId: number, reason: string): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.STAGES_LIST)
        .items
        .getById(stageId)
        .update({
          StageStatus: StageStatus.Skipped,
          ActionTaken: 'Skipped',
          StageCompletedDate: new Date().toISOString(),
          CompletedById: userId,
          StageComments: reason
        });

      // Move workflow to next stage
      await this.completeStage(stageId, 'Completed', userId, `Skipped: ${reason}`);
    } catch (error) {
      logger.error('DocumentWorkflowService', `Failed to skip stage ${stageId}:`, error);
      throw error;
    }
  }

  // ============================================================================
  // WORKFLOW MANAGEMENT
  // ============================================================================

  /**
   * Cancel a workflow
   */
  public async cancelWorkflow(workflowId: number, userId: number, reason: string): Promise<void> {
    try {
      const workflow = await this.getById(workflowId);
      if (!workflow) throw new Error('Workflow not found');

      await this.sp.web.lists
        .getByTitle(this.WORKFLOWS_LIST)
        .items
        .getById(workflowId)
        .update({
          WorkflowStatus: WorkflowStatus.Cancelled,
          CompletedDate: new Date().toISOString(),
          Outcome: 'Cancelled',
          OutcomeComments: reason
        });

      // Log cancellation
      await this.registryService.logActivity({
        DocumentRegistryId: workflow.DocumentRegistryId,
        DocumentTitle: workflow.DocumentTitle,
        ActivityType: ActivityType.WorkflowCompleted,
        ActivityById: userId,
        ActivityDate: new Date(),
        ActivitySeverity: ActivitySeverity.Warning,
        IsSystemAction: false,
        ActivityDetails: JSON.stringify({ workflowId, action: 'cancelled', reason })
      });
    } catch (error) {
      logger.error('DocumentWorkflowService', `Failed to cancel workflow ${workflowId}:`, error);
      throw error;
    }
  }

  /**
   * Put workflow on hold
   */
  public async putOnHold(workflowId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.WORKFLOWS_LIST)
        .items
        .getById(workflowId)
        .update({
          WorkflowStatus: WorkflowStatus.OnHold
        });
    } catch (error) {
      logger.error('DocumentWorkflowService', `Failed to put workflow ${workflowId} on hold:`, error);
      throw error;
    }
  }

  /**
   * Resume workflow from hold
   */
  public async resumeWorkflow(workflowId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.WORKFLOWS_LIST)
        .items
        .getById(workflowId)
        .update({
          WorkflowStatus: WorkflowStatus.InProgress
        });
    } catch (error) {
      logger.error('DocumentWorkflowService', `Failed to resume workflow ${workflowId}:`, error);
      throw error;
    }
  }

  // ============================================================================
  // PRIVATE HELPER METHODS
  // ============================================================================

  private calculateDefaultDueDate(stageCount: number): Date {
    const date = new Date();
    date.setDate(date.getDate() + (stageCount * 5)); // 5 days per stage
    return date;
  }

  private calculateStageDueDate(workflowDueDate: Date, stageIndex: number, totalStages: number): Date {
    const totalDays = Math.ceil((workflowDueDate.getTime() - Date.now()) / (1000 * 60 * 60 * 24));
    const daysPerStage = Math.floor(totalDays / totalStages);
    const date = new Date();
    date.setDate(date.getDate() + (daysPerStage * (stageIndex + 1)));
    return date;
  }

  private getWorkflowSelectFields(): string {
    return `Id,Title,DocumentRegistryId,DocumentTitle,WorkflowType,TemplateId,
      WorkflowStatus,CurrentStage,TotalStages,CurrentAssigneesId,DueDate,
      StartedDate,CompletedDate,Duration,Priority,EscalationLevel,IsOverdue,
      InitiatedById,Outcome,OutcomeComments,FinalApproverIds,RemindersSent,
      LastReminderDate,NotifyOnCompletion,NotificationRecipientsId,WorkflowConfig,
      AllowDelegation,AllowReassignment,RequireComments,Created,Modified`.replace(/\s/g, '');
  }

  private getStageSelectFields(): string {
    return `Id,Title,WorkflowId,StageNumber,StageType,RequiredAction,AssigneeType,
      AssigneeIds,AssigneeRole,StageStatus,ActionTaken,DueDays,StageDueDate,
      StageStartedDate,StageCompletedDate,CompletedById,StageComments,
      StageRemindersSent,StageLastReminderDate,DelegatedTo,DelegatedDate,
      DelegationReason`.replace(/\s/g, '');
  }

  private mapToWorkflow(item: any): IDocumentWorkflow {
    return {
      Id: item.Id,
      Title: item.Title,
      DocumentRegistryId: item.DocumentRegistryId,
      DocumentTitle: item.DocumentTitle,
      WorkflowType: item.WorkflowType as WorkflowType,
      TemplateId: item.TemplateId,
      WorkflowStatus: item.WorkflowStatus as WorkflowStatus,
      CurrentStage: item.CurrentStage,
      TotalStages: item.TotalStages,
      CurrentAssigneeIds: item.CurrentAssigneesId,
      DueDate: item.DueDate ? new Date(item.DueDate) : undefined,
      StartedDate: item.StartedDate ? new Date(item.StartedDate) : undefined,
      CompletedDate: item.CompletedDate ? new Date(item.CompletedDate) : undefined,
      Duration: item.Duration,
      Priority: item.Priority as Priority,
      EscalationLevel: item.EscalationLevel || 0,
      IsOverdue: item.IsOverdue,
      InitiatedById: item.InitiatedById,
      Outcome: item.Outcome,
      OutcomeComments: item.OutcomeComments,
      FinalApproverIds: item.FinalApproverIds,
      RemindersSent: item.RemindersSent || 0,
      LastReminderDate: item.LastReminderDate ? new Date(item.LastReminderDate) : undefined,
      NotifyOnCompletion: item.NotifyOnCompletion,
      NotificationRecipientIds: item.NotificationRecipientsId,
      WorkflowConfig: item.WorkflowConfig,
      AllowDelegation: item.AllowDelegation,
      AllowReassignment: item.AllowReassignment,
      RequireComments: item.RequireComments
    };
  }

  private mapToStage(item: any): IWorkflowStage {
    return {
      Id: item.Id,
      Title: item.Title,
      WorkflowId: item.WorkflowId,
      StageNumber: item.StageNumber,
      StageType: item.StageType as StageType,
      RequiredAction: item.RequiredAction,
      AssigneeType: item.AssigneeType,
      AssigneeIds: item.AssigneeIds?.results || item.AssigneeIds,
      AssigneeRole: item.AssigneeRole,
      StageStatus: item.StageStatus as StageStatus,
      ActionTaken: item.ActionTaken,
      DueDays: item.DueDays,
      StageDueDate: item.StageDueDate ? new Date(item.StageDueDate) : undefined,
      StageStartedDate: item.StageStartedDate ? new Date(item.StageStartedDate) : undefined,
      StageCompletedDate: item.StageCompletedDate ? new Date(item.StageCompletedDate) : undefined,
      CompletedById: item.CompletedById,
      StageComments: item.StageComments,
      StageRemindersSent: item.StageRemindersSent || 0,
      StageLastReminderDate: item.StageLastReminderDate ? new Date(item.StageLastReminderDate) : undefined,
      DelegatedToId: item.DelegatedTo,
      DelegatedDate: item.DelegatedDate ? new Date(item.DelegatedDate) : undefined,
      DelegationReason: item.DelegationReason
    };
  }
}
