// @ts-nocheck
/* eslint-disable */
/**
 * WorkflowEngineService
 * Core workflow execution engine - the heart of the internal workflow system
 * Implements state machine pattern for workflow orchestration
 */

import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  IWorkflowDefinition,
  IWorkflowInstance,
  IWorkflowStep,
  IWorkflowContext,
  IWorkflowStepStatus,
  IActionContext,
  IActionResult,
  IServiceContainer,
  WorkflowInstanceStatus,
  StepStatus,
  StepType,
  TransitionType,
  LogLevel
} from '../../models/IWorkflow';
import { ProcessType } from '../../models/ICommon';
import { WorkflowDefinitionService, IParsedWorkflowDefinition } from './WorkflowDefinitionService';
import { WorkflowInstanceService } from './WorkflowInstanceService';
import { WorkflowConditionEvaluator } from './WorkflowConditionEvaluator';
import { WorkflowActionDispatcher } from './WorkflowActionDispatcher';
import { WorkflowNotificationService } from './WorkflowNotificationService';
import {
  ForEachHandler,
  SubWorkflowHandler,
  WebhookHandler,
  RetryHandler,
  IRetryContext
} from './handlers';
import { logger } from '../LoggingService';
import {
  retryWithDLQ,
  PROCESS_SYNC_RETRY_OPTIONS,
  workflowSyncDLQ,
  IRetryResult
} from '../../utils/retryUtils';

/**
 * Options for starting a workflow
 */
export interface IStartWorkflowOptions {
  definitionId?: number;
  definitionCode?: string;
  processId: number;
  processType: ProcessType;
  employeeName: string;
  employeeEmail?: string;
  department: string;
  managerId?: number;
  managerEmail?: string;
  startedByUserId: number;
  startedByUserName: string;
  customContext?: Record<string, unknown>;
}

/**
 * Result of workflow execution step
 */
export interface IExecutionResult {
  success: boolean;
  instanceId: number;
  status: WorkflowInstanceStatus;
  currentStepId?: string;
  currentStepName?: string;
  message?: string;
  nextAction?: 'continue' | 'wait' | 'complete' | 'error';
  error?: string;
}

/**
 * Callback type for process status synchronization
 */
export type ProcessStatusSyncCallback = (
  processId: number,
  workflowStatus: WorkflowInstanceStatus,
  workflowInstanceId: number
) => Promise<void>;

export class WorkflowEngineService {
  private sp: SPFI;
  private context: WebPartContext;
  private definitionService: WorkflowDefinitionService;
  private instanceService: WorkflowInstanceService;
  private conditionEvaluator: WorkflowConditionEvaluator;
  private actionDispatcher: WorkflowActionDispatcher;
  private notificationService: WorkflowNotificationService;

  // Enhanced handlers (Phase 1)
  private forEachHandler: ForEachHandler;
  private subWorkflowHandler: SubWorkflowHandler;
  private webhookHandler: WebhookHandler;
  private retryHandler: RetryHandler;

  // Callback for process status synchronization
  private processSyncCallback?: ProcessStatusSyncCallback;

  constructor(sp: SPFI, context: WebPartContext) {
    this.sp = sp;
    this.context = context;
    this.definitionService = new WorkflowDefinitionService(sp);
    this.instanceService = new WorkflowInstanceService(sp);
    this.conditionEvaluator = new WorkflowConditionEvaluator();
    this.actionDispatcher = new WorkflowActionDispatcher(sp, context);
    this.notificationService = new WorkflowNotificationService(sp, context);

    // Initialize enhanced handlers
    this.forEachHandler = new ForEachHandler();
    this.subWorkflowHandler = new SubWorkflowHandler(sp, context);
    this.webhookHandler = new WebhookHandler(context);
    this.retryHandler = new RetryHandler();
  }

  /**
   * Register a callback for process status synchronization
   * This callback is called whenever the workflow status changes
   */
  public onProcessStatusSync(callback: ProcessStatusSyncCallback): void {
    this.processSyncCallback = callback;
  }

  /**
   * Notify process of workflow status change with retry and dead letter queue
   * Critical for maintaining process-workflow status synchronization
   */
  private async notifyProcessStatusChange(
    processId: number,
    workflowStatus: WorkflowInstanceStatus,
    workflowInstanceId: number
  ): Promise<IRetryResult<void> | void> {
    if (!this.processSyncCallback) {
      return;
    }

    const syncPayload = {
      processId,
      workflowStatus,
      workflowInstanceId,
      timestamp: new Date().toISOString()
    };

    const result = await retryWithDLQ<void>(
      async () => {
        await this.processSyncCallback!(processId, workflowStatus, workflowInstanceId);
      },
      'process-status-sync',
      syncPayload,
      PROCESS_SYNC_RETRY_OPTIONS,
      workflowSyncDLQ,
      {
        source: 'WorkflowEngineService',
        operation: 'notifyProcessStatusChange'
      }
    );

    if (!result.success) {
      logger.error(
        'WorkflowEngineService',
        `Process sync failed after ${result.attempts} attempts. Added to DLQ: ${result.deadLetterItemId}`,
        result.error
      );

      // Log for monitoring/alerting
      await this.instanceService.addLog(
        workflowInstanceId,
        undefined,
        undefined,
        'Process Sync Failed',
        LogLevel.Warning,
        `Failed to sync process ${processId} status to ${workflowStatus}. DLQ ID: ${result.deadLetterItemId}`,
        { ...syncPayload, error: result.error?.message }
      );
    } else {
      logger.info(
        'WorkflowEngineService',
        `Process ${processId} synced to status ${workflowStatus} (${result.attempts} attempt(s), ${result.totalDurationMs}ms)`
      );
    }

    return result;
  }

  // ============================================================================
  // WORKFLOW LIFECYCLE
  // ============================================================================

  /**
   * Start a new workflow instance
   */
  public async startWorkflow(options: IStartWorkflowOptions): Promise<IExecutionResult> {
    try {
      // Get workflow definition
      let definition: IWorkflowDefinition | undefined;

      if (options.definitionId) {
        definition = await this.definitionService.getById(options.definitionId);
      } else if (options.definitionCode) {
        definition = await this.definitionService.getByCode(options.definitionCode);
      } else {
        // Get default for process type
        definition = await this.definitionService.getDefaultForProcessType(options.processType);
      }

      if (!definition) {
        throw new Error(`No workflow definition found for ${options.processType}`);
      }

      if (!definition.IsActive) {
        throw new Error(`Workflow "${definition.Title}" is not active`);
      }

      // Check if process already has an active workflow
      const existingInstance = await this.instanceService.getActiveForProcess(options.processId);
      if (existingInstance) {
        throw new Error(`Process ${options.processId} already has an active workflow (Instance: ${existingInstance.Id})`);
      }

      // Parse definition
      const parsedDef = this.definitionService.parseDefinition(definition);

      // Create workflow context
      const workflowContext: IWorkflowContext = {
        processId: options.processId,
        processType: options.processType,
        employeeName: options.employeeName,
        employeeEmail: options.employeeEmail || '',
        department: options.department,
        managerId: options.managerId,
        managerEmail: options.managerEmail,
        startedBy: options.startedByUserName,
        startedAt: new Date(),
        customFields: options.customContext
      };

      // Find start step
      const startStep = parsedDef.steps.find(s => s.type === StepType.Start);
      if (!startStep) {
        throw new Error('Workflow definition has no Start step');
      }

      // Calculate estimated completion
      const estimatedHours = definition.EstimatedDuration || this.estimateWorkflowDuration(parsedDef.steps);
      const estimatedCompletion = new Date();
      estimatedCompletion.setHours(estimatedCompletion.getHours() + estimatedHours);

      // Create instance
      const instance = await this.instanceService.create({
        Title: `${definition.Title} - ${options.employeeName}`,
        WorkflowDefinitionId: definition.Id,
        ProcessId: options.processId,
        Status: WorkflowInstanceStatus.Running,
        CurrentStepId: startStep.id,
        CurrentStepName: startStep.name,
        TotalSteps: parsedDef.steps.length,
        CompletedSteps: 0,
        ProgressPercentage: 0,
        EstimatedCompletionDate: estimatedCompletion,
        Context: JSON.stringify(workflowContext),
        Variables: JSON.stringify(parsedDef.variables.reduce((acc, v) => {
          acc[v.name] = v.defaultValue;
          return acc;
        }, {} as Record<string, unknown>)),
        StartedById: options.startedByUserId
      });

      // Create step status records for all steps
      for (const step of parsedDef.steps) {
        await this.instanceService.createStepStatus({
          WorkflowInstanceId: instance.Id,
          StepId: step.id,
          StepName: step.name,
          Status: step.id === startStep.id ? StepStatus.InProgress : StepStatus.Pending,
          Order: step.order
        });
      }

      // Log start
      await this.instanceService.addLog(
        instance.Id,
        startStep.id,
        startStep.name,
        'Workflow Started',
        LogLevel.Info,
        `Started workflow "${definition.Title}" for ${options.employeeName}`,
        { processId: options.processId, definitionId: definition.Id },
        options.startedByUserId
      );

      // Increment usage counter
      await this.definitionService.incrementUsageCount(definition.Id);

      // Send workflow started notification
      try {
        const recipientIds: number[] = [];
        const recipientEmails: string[] = [];

        // Notify the employee and manager
        if (options.managerId) {
          recipientIds.push(options.managerId);
        }
        if (options.managerEmail) {
          recipientEmails.push(options.managerEmail);
        }
        if (options.employeeEmail) {
          recipientEmails.push(options.employeeEmail);
        }

        if (recipientIds.length > 0 || recipientEmails.length > 0) {
          await this.notificationService.notifyWorkflowStarted(
            { ...instance, ProcessType: options.processType } as IWorkflowInstance,
            recipientIds,
            recipientEmails
          );
        }
      } catch (notifyError) {
        // Log but don't fail the workflow for notification errors
        logger.warn('WorkflowEngineService', 'Failed to send workflow started notification', notifyError);
      }

      // Notify process of workflow start
      await this.notifyProcessStatusChange(
        options.processId,
        WorkflowInstanceStatus.Running,
        instance.Id
      );

      // Execute start step and continue
      const result = await this.executeStep(instance.Id, startStep.id);

      return result;
    } catch (error) {
      logger.error('WorkflowEngineService', 'Error starting workflow', error);
      throw error;
    }
  }

  /**
   * Resume workflow execution (for waiting workflows)
   */
  public async resumeWorkflow(instanceId: number, triggerData?: Record<string, unknown>): Promise<IExecutionResult> {
    try {
      const instance = await this.instanceService.getById(instanceId);

      // Validate status allows resume
      const resumableStatuses = [
        WorkflowInstanceStatus.Paused,
        WorkflowInstanceStatus.WaitingForInput,
        WorkflowInstanceStatus.WaitingForApproval,
        WorkflowInstanceStatus.WaitingForTask
      ];

      if (!resumableStatuses.includes(instance.Status)) {
        throw new Error(`Workflow cannot be resumed from status: ${instance.Status}`);
      }

      // Update status to running
      await this.instanceService.updateStatus(instanceId, WorkflowInstanceStatus.Running);

      // If trigger data provided, merge into variables
      if (triggerData) {
        await this.instanceService.updateVariables(instanceId, triggerData);
      }

      // Log resume
      await this.instanceService.addLog(
        instanceId,
        instance.CurrentStepId,
        instance.CurrentStepName,
        'Workflow Resumed',
        LogLevel.Info,
        'Workflow execution resumed'
      );

      // Continue from current step
      if (instance.CurrentStepId) {
        return await this.executeStep(instanceId, instance.CurrentStepId);
      }

      return {
        success: true,
        instanceId,
        status: WorkflowInstanceStatus.Running,
        message: 'Workflow resumed'
      };
    } catch (error) {
      logger.error('WorkflowEngineService', `Error resuming workflow ${instanceId}`, error);
      throw error;
    }
  }

  /**
   * Complete a waiting step (task completed, approval received, etc.)
   * CRITICAL: Includes idempotency check to prevent race conditions from dual resume paths
   */
  public async completeWaitingStep(
    instanceId: number,
    stepId: string,
    result: Record<string, unknown>
  ): Promise<IExecutionResult> {
    try {
      const instance = await this.instanceService.getById(instanceId);

      // IDEMPOTENCY CHECK: Prevent race conditions from dual resume paths
      // (e.g., direct call from task completion + polling service both trying to resume)
      const stepStatus = await this.instanceService.getStepStatus(instanceId, stepId);
      if (stepStatus && (stepStatus.Status === StepStatus.Completed || stepStatus.Status === StepStatus.Skipped)) {
        logger.info(
          'WorkflowEngineService',
          `Idempotency check: Step ${stepId} already completed (status: ${stepStatus.Status}). Skipping duplicate completion.`
        );
        // Return current state without re-executing
        return {
          success: true,
          instanceId,
          status: instance.Status,
          currentStepId: instance.CurrentStepId,
          currentStepName: instance.CurrentStepName,
          message: 'Step already completed (idempotent)',
          nextAction: instance.Status === WorkflowInstanceStatus.Completed ? 'complete' : 'continue'
        };
      }

      const definition = await this.definitionService.getParsed(instance.WorkflowDefinitionId);
      const step = definition.steps.find(s => s.id === stepId);

      if (!step) {
        throw new Error(`Step ${stepId} not found in workflow`);
      }

      // Complete the step
      await this.instanceService.completeStep(instanceId, stepId, result);

      // Log completion
      await this.instanceService.addLog(
        instanceId,
        stepId,
        step.name,
        'Step Completed',
        LogLevel.Info,
        `Step "${step.name}" completed`,
        result
      );

      // Determine next step
      const nextStepId = this.resolveTransition(step, result, definition);

      if (nextStepId) {
        // Update status and continue
        await this.instanceService.updateStatus(instanceId, WorkflowInstanceStatus.Running);
        return await this.executeStep(instanceId, nextStepId);
      } else {
        // No next step - workflow complete
        return await this.completeWorkflow(instanceId);
      }
    } catch (error) {
      logger.error('WorkflowEngineService', `Error completing waiting step ${stepId}`, error);
      throw error;
    }
  }

  // ============================================================================
  // STEP EXECUTION
  // ============================================================================

  /**
   * Execute a workflow step
   */
  public async executeStep(instanceId: number, stepId: string): Promise<IExecutionResult> {
    try {
      // Get instance and definition
      const instance = await this.instanceService.getById(instanceId);
      const definition = await this.definitionService.getParsed(instance.WorkflowDefinitionId);
      const step = definition.steps.find(s => s.id === stepId);

      if (!step) {
        throw new Error(`Step ${stepId} not found in workflow definition`);
      }

      // Parse context and variables
      const context: IWorkflowContext = instance.Context ? JSON.parse(instance.Context) : {};
      const variables: Record<string, unknown> = instance.Variables ? JSON.parse(instance.Variables) : {};

      // Update progress
      const completedCount = await this.getCompletedStepCount(instanceId);
      await this.instanceService.updateProgress(
        instanceId,
        stepId,
        step.name,
        completedCount,
        definition.steps.length
      );

      // Log step start
      await this.instanceService.addLog(
        instanceId,
        stepId,
        step.name,
        'Step Started',
        LogLevel.Info,
        `Starting step "${step.name}" (Type: ${step.type})`
      );

      // Start step execution
      await this.instanceService.startStep(instanceId, stepId);

      // Check entry conditions
      if (step.conditions && step.conditions.length > 0) {
        const conditionsContext = { ...context, variables, process: { id: instance.ProcessId } };
        const conditionsMet = this.conditionEvaluator.evaluateConditions(step.conditions, conditionsContext);

        if (!conditionsMet) {
          // Skip step
          await this.instanceService.skipStep(instanceId, stepId, 'Entry conditions not met');
          await this.instanceService.addLog(
            instanceId,
            stepId,
            step.name,
            'Step Skipped',
            LogLevel.Info,
            'Entry conditions not met - step skipped'
          );

          // Move to next step
          const nextStepId = this.resolveTransition(step, { skipped: true }, definition);
          if (nextStepId) {
            return await this.executeStep(instanceId, nextStepId);
          } else {
            return await this.completeWorkflow(instanceId);
          }
        }
      }

      // Execute step based on type
      const actionContext: IActionContext = {
        workflowInstance: instance,
        currentStep: step,
        stepStatus: (await this.instanceService.getStepStatus(instanceId, stepId))!,
        process: { id: instance.ProcessId, ...context },
        variables,
        services: this.createServiceContainer()
      };

      let actionResult: IActionResult;

      switch (step.type) {
        case StepType.Start:
          actionResult = { success: true, nextAction: 'continue' };
          break;

        case StepType.End:
          actionResult = { success: true, nextAction: 'continue' };
          await this.instanceService.completeStep(instanceId, stepId, { completed: true });
          return await this.completeWorkflow(instanceId);

        case StepType.AssignTasks:
        case StepType.CreateTask:
        case StepType.Approval:
        case StepType.Action:
        case StepType.Notification:
        case StepType.SetVariable:
          actionResult = await this.actionDispatcher.dispatch(step.type, step.config, actionContext);
          break;

        case StepType.Condition:
          // Evaluate condition and determine branch
          const branchResult = this.evaluateConditionStep(step, actionContext);
          actionResult = {
            success: true,
            nextAction: 'continue',
            outputVariables: { branchTaken: branchResult.branchName }
          };
          // Override normal transition with branch result
          if (branchResult.targetStepId) {
            await this.instanceService.completeStep(instanceId, stepId, { branch: branchResult.branchName });
            return await this.executeStep(instanceId, branchResult.targetStepId);
          }
          break;

        case StepType.Wait:
          actionResult = await this.actionDispatcher.dispatch(step.type, step.config, actionContext);
          break;

        case StepType.WaitForTasks:
          actionResult = await this.handleWaitForTasks(step, actionContext);
          break;

        case StepType.Parallel:
          actionResult = await this.handleParallelSteps(step, actionContext);
          break;

        // ============================================================================
        // ENHANCED STEP TYPES (Phase 1 Implementation)
        // ============================================================================

        case StepType.ForEach:
          // Execute ForEach loop over a collection
          actionResult = await this.forEachHandler.execute(
            step,
            actionContext,
            async (iterStep, iterContext) => {
              // Execute nested step within the loop
              return await this.actionDispatcher.dispatch(
                iterStep.type,
                iterStep.config,
                iterContext
              );
            }
          );
          break;

        case StepType.CallWorkflow:
          // Execute sub-workflow
          actionResult = await this.subWorkflowHandler.execute(
            step,
            actionContext,
            async (subInstanceId, subStepId) => {
              // Execute sub-workflow step using this engine
              const result = await this.executeStep(subInstanceId, subStepId);
              return {
                success: result.success,
                status: result.status
              };
            }
          );
          break;

        case StepType.Webhook:
          // Call external webhook
          actionResult = await this.webhookHandler.execute(step, actionContext);
          break;

        default:
          actionResult = { success: false, error: `Unknown step type: ${step.type}` };
      }

      // ============================================================================
      // ENHANCED ERROR HANDLING WITH RETRY & EXPONENTIAL BACKOFF
      // ============================================================================
      if (!actionResult.success && step.errorConfig) {
        // Build retry context from step status
        const existingRetryContext: IRetryContext | undefined = actionContext.stepStatus?.RetryCount
          ? {
              attemptNumber: actionContext.stepStatus.RetryCount,
              maxAttempts: step.errorConfig.retryCount || 3,
              lastError: actionContext.stepStatus.ErrorMessage || '',
              totalDelayMs: 0
            }
          : undefined;

        // Create a proper executor that actually re-executes the step action
        const stepExecutor = async (s: IWorkflowStep, ctx: IActionContext): Promise<IActionResult> => {
          // Re-dispatch the step action based on type
          switch (s.type) {
            case StepType.AssignTasks:
            case StepType.CreateTask:
            case StepType.Approval:
            case StepType.Action:
            case StepType.Notification:
            case StepType.SetVariable:
            case StepType.Wait:
              return await this.actionDispatcher.dispatch(s.type, s.config, ctx);

            case StepType.Webhook:
              return await this.webhookHandler.execute(s, ctx);

            case StepType.ForEach:
              return await this.forEachHandler.execute(
                s,
                ctx,
                async (iterStep, iterContext) => {
                  return await this.actionDispatcher.dispatch(iterStep.type, iterStep.config, iterContext);
                }
              );

            case StepType.CallWorkflow:
              return await this.subWorkflowHandler.execute(
                s,
                ctx,
                async (subInstanceId, subStepId) => {
                  const result = await this.executeStep(subInstanceId, subStepId);
                  return { success: result.success, status: result.status };
                }
              );

            default:
              // For non-retryable step types, return the original failure
              return actionResult;
          }
        };

        // Apply retry/error handling configuration with exponential backoff
        const retryResult = await this.retryHandler.executeWithRetry(
          step,
          actionContext,
          stepExecutor,
          existingRetryContext
        );

        // Update action result with retry decision
        actionResult = retryResult;

        // Store retry count and schedule next retry if needed
        if (retryResult.shouldRetry && retryResult.retryContext && actionContext.stepStatus) {
          await this.instanceService.updateStepStatus(actionContext.stepStatus.Id, {
            RetryCount: retryResult.retryContext.attemptNumber,
            ErrorMessage: retryResult.retryContext.lastError
          });

          // Log retry scheduling with exponential backoff details
          const backoffDelay = retryResult.retryDelayMs || 0;
          logger.info(
            'WorkflowEngineService',
            `Retry scheduled for step "${step.name}" (attempt ${retryResult.retryContext.attemptNumber + 1}/${retryResult.retryContext.maxAttempts + 1}) ` +
            `in ${Math.round(backoffDelay / 1000)}s (exponential backoff)`
          );
        }
      }

      // Handle action result
      if (!actionResult.success) {
        await this.instanceService.failStep(instanceId, stepId, actionResult.error || 'Unknown error');
        await this.instanceService.setError(instanceId, stepId, actionResult.error || 'Unknown error');

        return {
          success: false,
          instanceId,
          status: WorkflowInstanceStatus.Failed,
          currentStepId: stepId,
          currentStepName: step.name,
          error: actionResult.error,
          nextAction: 'error'
        };
      }

      // Update variables if any output
      if (actionResult.outputVariables) {
        await this.instanceService.updateVariables(instanceId, actionResult.outputVariables);
      }

      // Handle next action
      switch (actionResult.nextAction) {
        case 'wait':
          // Update status based on what we're waiting for
          let waitStatus = WorkflowInstanceStatus.WaitingForInput;
          if (actionResult.waitForItemType === 'task') {
            waitStatus = WorkflowInstanceStatus.WaitingForTask;
          } else if (actionResult.waitForItemType === 'approval') {
            waitStatus = WorkflowInstanceStatus.WaitingForApproval;
          }

          await this.instanceService.updateStatus(instanceId, waitStatus);

          // Notify process of workflow waiting
          await this.notifyProcessStatusChange(
            instance.ProcessId,
            waitStatus,
            instanceId
          );

          return {
            success: true,
            instanceId,
            status: waitStatus,
            currentStepId: stepId,
            currentStepName: step.name,
            message: `Waiting for ${actionResult.waitForItemType || 'input'}`,
            nextAction: 'wait'
          };

        case 'continue':
        default:
          // Complete step
          await this.instanceService.completeStep(instanceId, stepId, { success: true }, actionResult.outputVariables);

          // Determine next step
          const nextStepId = this.resolveTransition(step, actionResult, definition);

          if (nextStepId) {
            // Continue to next step
            return await this.executeStep(instanceId, nextStepId);
          } else {
            // No next step - workflow complete
            return await this.completeWorkflow(instanceId);
          }
      }
    } catch (error) {
      logger.error('WorkflowEngineService', `Error executing step ${stepId}`, error);

      const errorMessage = error instanceof Error ? error.message : 'Unknown error';

      // Set error state
      await this.instanceService.setError(instanceId, stepId, errorMessage);

      // Get instance for notifications and process sync
      let failedInstance: IWorkflowInstance | undefined;

      // Send workflow failed notification
      try {
        failedInstance = await this.instanceService.getById(instanceId);
        const context = failedInstance.Context ? JSON.parse(failedInstance.Context) : {};
        const recipientIds: number[] = [];
        const recipientEmails: string[] = [];

        // Notify stakeholders
        if (context.managerId) {
          recipientIds.push(context.managerId);
        }
        if (context.managerEmail) {
          recipientEmails.push(context.managerEmail);
        }

        if (recipientIds.length > 0 || recipientEmails.length > 0) {
          await this.notificationService.notifyWorkflowFailed(
            failedInstance,
            recipientIds,
            errorMessage,
            recipientEmails
          );
        }
      } catch (notifyError) {
        logger.warn('WorkflowEngineService', 'Failed to send workflow failed notification', notifyError);
      }

      // Notify process of workflow failure
      if (failedInstance) {
        await this.notifyProcessStatusChange(
          failedInstance.ProcessId,
          WorkflowInstanceStatus.Failed,
          instanceId
        );
      }

      return {
        success: false,
        instanceId,
        status: WorkflowInstanceStatus.Failed,
        currentStepId: stepId,
        error: errorMessage,
        nextAction: 'error'
      };
    }
  }

  // ============================================================================
  // WORKFLOW COMPLETION
  // ============================================================================

  /**
   * Complete workflow instance
   */
  private async completeWorkflow(instanceId: number): Promise<IExecutionResult> {
    try {
      const instance = await this.instanceService.getById(instanceId);

      // Update status
      await this.instanceService.update(instanceId, {
        Status: WorkflowInstanceStatus.Completed,
        CompletedDate: new Date(),
        ProgressPercentage: 100
      });

      // Log completion
      await this.instanceService.addLog(
        instanceId,
        undefined,
        undefined,
        'Workflow Completed',
        LogLevel.Info,
        'Workflow execution completed successfully'
      );

      // Update definition statistics
      const stats = await this.instanceService.getDefinitionStats(instance.WorkflowDefinitionId);
      await this.definitionService.updateSuccessRate(instance.WorkflowDefinitionId, stats.successRate);
      await this.definitionService.updateAverageCompletionTime(instance.WorkflowDefinitionId, stats.averageDurationHours);

      // Send workflow completed notification
      try {
        const context = instance.Context ? JSON.parse(instance.Context) : {};
        const recipientIds: number[] = [];
        const recipientEmails: string[] = [];

        // Notify stakeholders from context
        if (context.managerId) {
          recipientIds.push(context.managerId);
        }
        if (context.managerEmail) {
          recipientEmails.push(context.managerEmail);
        }
        if (context.employeeEmail) {
          recipientEmails.push(context.employeeEmail);
        }

        if (recipientIds.length > 0 || recipientEmails.length > 0) {
          await this.notificationService.notifyWorkflowCompleted(
            instance,
            recipientIds,
            recipientEmails
          );
        }
      } catch (notifyError) {
        // Log but don't fail for notification errors
        logger.warn('WorkflowEngineService', 'Failed to send workflow completed notification', notifyError);
      }

      // Notify process of workflow completion
      await this.notifyProcessStatusChange(
        instance.ProcessId,
        WorkflowInstanceStatus.Completed,
        instanceId
      );

      return {
        success: true,
        instanceId,
        status: WorkflowInstanceStatus.Completed,
        message: 'Workflow completed successfully',
        nextAction: 'complete'
      };
    } catch (error) {
      logger.error('WorkflowEngineService', `Error completing workflow ${instanceId}`, error);
      throw error;
    }
  }

  // ============================================================================
  // TRANSITION RESOLUTION
  // ============================================================================

  /**
   * Resolve next step based on transition configuration
   */
  private resolveTransition(
    currentStep: IWorkflowStep,
    result: IActionResult | Record<string, unknown>,
    definition: IParsedWorkflowDefinition
  ): string | undefined {
    const transition = currentStep.onComplete;

    if (!transition) {
      // Find next step by order
      const nextByOrder = definition.steps
        .filter(s => s.order > currentStep.order)
        .sort((a, b) => a.order - b.order)[0];
      return nextByOrder?.id;
    }

    switch (transition.type) {
      case TransitionType.Next:
        // Go to next step by order
        const nextStep = definition.steps
          .filter(s => s.order > currentStep.order)
          .sort((a, b) => a.order - b.order)[0];
        return nextStep?.id;

      case TransitionType.Goto:
        return transition.targetStepId;

      case TransitionType.Branch:
        // Evaluate branches and find matching one
        if (transition.branches) {
          for (const branch of transition.branches) {
            if (branch.isDefault) continue;

            const conditionsMet = this.conditionEvaluator.evaluateConditionGroups(
              branch.conditions,
              result as Record<string, unknown>
            );

            if (conditionsMet) {
              return branch.targetStepId;
            }
          }

          // Use default branch if no conditions matched
          const defaultBranch = transition.branches.find(b => b.isDefault);
          return defaultBranch?.targetStepId;
        }
        break;

      case TransitionType.End:
        return undefined;

      case TransitionType.Parallel:
        // For parallel, we'd need special handling
        // Return first parallel step for now
        return transition.parallelStepIds?.[0];
    }

    return undefined;
  }

  // ============================================================================
  // CONDITION STEP HANDLING
  // ============================================================================

  /**
   * Evaluate condition step and determine branch
   */
  private evaluateConditionStep(
    step: IWorkflowStep,
    context: IActionContext
  ): { branchName: string; targetStepId?: string } {
    if (!step.config.conditionGroups) {
      return { branchName: 'default' };
    }

    const evalContext = {
      ...context.process,
      ...context.variables
    };

    // Check each condition group
    const result = this.conditionEvaluator.evaluateConditionGroups(
      step.config.conditionGroups,
      evalContext
    );

    if (result && step.onComplete?.branches) {
      // Find matching branch
      for (const branch of step.onComplete.branches) {
        if (!branch.isDefault) {
          const branchMet = this.conditionEvaluator.evaluateConditionGroups(
            branch.conditions,
            evalContext
          );
          if (branchMet) {
            return { branchName: branch.name, targetStepId: branch.targetStepId };
          }
        }
      }

      // Default branch
      const defaultBranch = step.onComplete.branches.find(b => b.isDefault);
      if (defaultBranch) {
        return { branchName: defaultBranch.name, targetStepId: defaultBranch.targetStepId };
      }
    }

    return { branchName: 'none' };
  }

  // ============================================================================
  // SPECIAL STEP HANDLERS
  // ============================================================================

  /**
   * Handle WaitForTasks step with timeout support
   * Waits for tasks created by previous steps with configurable timeout
   */
  private async handleWaitForTasks(
    step: IWorkflowStep,
    context: IActionContext
  ): Promise<IActionResult> {
    const taskStepIds = step.config.waitForTaskIds || [];
    const waitCondition = step.config.waitCondition || 'all'; // 'all' | 'any'
    const timeoutHours = step.config.timeoutHours || step.config.slaHours || 0; // 0 = no timeout
    const onTimeout = step.config.onTimeout || 'escalate'; // 'escalate' | 'skip' | 'fail'
    const instanceId = context.workflowInstance.Id;

    // Get step status to check for timeout
    const stepStatus = context.stepStatus;
    const stepStartedAt = stepStatus?.StartedDate ? new Date(stepStatus.StartedDate) : new Date();

    // Check if timeout has occurred
    if (timeoutHours > 0) {
      const timeoutMs = timeoutHours * 60 * 60 * 1000;
      const elapsedMs = Date.now() - stepStartedAt.getTime();

      if (elapsedMs >= timeoutMs) {
        // Timeout has occurred
        logger.warn(
          'WorkflowEngineService',
          `WaitForTasks step "${step.name}" has timed out after ${timeoutHours} hours`
        );

        // Handle timeout based on configuration
        return await this.handleWaitTimeout(step, context, onTimeout, elapsedMs);
      }
    }

    // Check actual task completion status
    const taskCompletionResult = await this.checkTasksCompletion(instanceId, taskStepIds, waitCondition);

    if (taskCompletionResult.allComplete || (waitCondition === 'any' && taskCompletionResult.anyComplete)) {
      // Tasks are complete - continue workflow
      return {
        success: true,
        nextAction: 'continue',
        outputVariables: {
          waitingForSteps: taskStepIds,
          waitCondition,
          completedTasks: taskCompletionResult.completedTasks,
          pendingTasks: taskCompletionResult.pendingTasks,
          taskResults: taskCompletionResult.taskResults
        }
      };
    }

    // Still waiting - return wait action with timeout info
    const remainingTimeMs = timeoutHours > 0
      ? (timeoutHours * 60 * 60 * 1000) - (Date.now() - stepStartedAt.getTime())
      : undefined;

    return {
      success: true,
      nextAction: 'wait',
      waitForItemType: 'task',
      waitForItemIds: taskCompletionResult.pendingTaskIds,
      outputVariables: {
        waitingForSteps: taskStepIds,
        waitCondition,
        completedTasks: taskCompletionResult.completedTasks,
        pendingTasks: taskCompletionResult.pendingTasks,
        pendingTaskIds: taskCompletionResult.pendingTaskIds,
        timeoutHours: timeoutHours || undefined,
        remainingTimeMs,
        onTimeout
      }
    };
  }

  /**
   * Handle timeout for WaitForTasks step
   */
  private async handleWaitTimeout(
    step: IWorkflowStep,
    context: IActionContext,
    onTimeout: string,
    elapsedMs: number
  ): Promise<IActionResult> {
    const instanceId = context.workflowInstance.Id;

    switch (onTimeout) {
      case 'skip':
        // Skip the waiting step and continue
        logger.info('WorkflowEngineService', `Timeout: Skipping WaitForTasks step "${step.name}"`);
        await this.instanceService.skipStep(instanceId, step.id, `Timeout after ${Math.floor(elapsedMs / 3600000)} hours`);

        return {
          success: true,
          nextAction: 'continue',
          outputVariables: {
            timeoutAction: 'skipped',
            elapsedMs
          }
        };

      case 'fail':
        // Fail the workflow
        logger.error('WorkflowEngineService', `Timeout: Failing workflow at step "${step.name}"`);

        return {
          success: false,
          error: `WaitForTasks step "${step.name}" timed out after ${Math.floor(elapsedMs / 3600000)} hours`,
          outputVariables: {
            timeoutAction: 'failed',
            elapsedMs
          }
        };

      case 'escalate':
      default:
        // Escalate - send notification and continue waiting
        logger.warn('WorkflowEngineService', `Timeout: Escalating WaitForTasks step "${step.name}"`);

        // Send escalation notification
        await this.sendTimeoutEscalation(step, context, elapsedMs);

        // Update timeout counter to prevent repeat escalations
        const currentTimeoutCount = (context.stepStatus?.RetryCount || 0) + 1;
        if (context.stepStatus?.Id) {
          await this.instanceService.updateStepStatus(context.stepStatus.Id, {
            RetryCount: currentTimeoutCount,
            ErrorMessage: `Escalated at ${new Date().toISOString()} after ${Math.floor(elapsedMs / 3600000)} hours`
          });
        }

        return {
          success: true,
          nextAction: 'wait',
          waitForItemType: 'task',
          outputVariables: {
            timeoutAction: 'escalated',
            escalationCount: currentTimeoutCount,
            elapsedMs
          }
        };
    }
  }

  /**
   * Check if tasks are complete for WaitForTasks step
   */
  private async checkTasksCompletion(
    instanceId: number,
    taskStepIds: string[],
    _waitCondition: string
  ): Promise<{
    allComplete: boolean;
    anyComplete: boolean;
    completedTasks: string[];
    pendingTasks: string[];
    pendingTaskIds: number[];
    taskResults: Array<{ stepId: string; status: string }>
  }> {
    const completedTasks: string[] = [];
    const pendingTasks: string[] = [];
    const pendingTaskIds: number[] = [];
    const taskResults: Array<{ stepId: string; status: string }> = [];

    for (const stepId of taskStepIds) {
      const stepStatus = await this.instanceService.getStepStatus(instanceId, stepId);

      if (stepStatus) {
        taskResults.push({ stepId, status: stepStatus.Status });

        if (stepStatus.Status === StepStatus.Completed || stepStatus.Status === StepStatus.Skipped) {
          completedTasks.push(stepId);
        } else if (stepStatus.Status !== StepStatus.Failed) {
          pendingTasks.push(stepId);
          // Extract actual task assignment IDs from step status
          if (stepStatus.TaskAssignmentIds) {
            const taskIds = stepStatus.TaskAssignmentIds.split(',')
              .map(id => parseInt(id.trim(), 10))
              .filter(id => !isNaN(id));
            pendingTaskIds.push(...taskIds);
          }
        }
      } else {
        pendingTasks.push(stepId);
        taskResults.push({ stepId, status: 'NotFound' });
      }
    }

    return {
      allComplete: completedTasks.length === taskStepIds.length,
      anyComplete: completedTasks.length > 0,
      completedTasks,
      pendingTasks,
      pendingTaskIds,
      taskResults
    };
  }

  /**
   * Send timeout escalation notification
   */
  private async sendTimeoutEscalation(
    step: IWorkflowStep,
    context: IActionContext,
    elapsedMs: number
  ): Promise<void> {
    try {
      const workflowContext = context.workflowInstance.Context
        ? JSON.parse(context.workflowInstance.Context)
        : {};

      const recipientIds: number[] = [];
      const recipientEmails: string[] = [];

      // Add manager if available
      if (workflowContext.managerId) {
        recipientIds.push(workflowContext.managerId);
      }
      if (workflowContext.managerEmail) {
        recipientEmails.push(workflowContext.managerEmail);
      }

      // Add escalation recipients from step config
      if (step.config.escalateToUserIds) {
        recipientIds.push(...step.config.escalateToUserIds);
      }
      if (step.config.escalateToEmails) {
        recipientEmails.push(...step.config.escalateToEmails);
      }

      if (recipientIds.length > 0 || recipientEmails.length > 0) {
        const elapsedHours = Math.floor(elapsedMs / 3600000);
        // Use SLA breach notification for timeout escalation
        await this.notificationService.notifySLABreach(
          context.workflowInstance,
          step.name,
          elapsedHours,
          recipientIds,
          recipientEmails
        );
      }
    } catch (error) {
      logger.warn('WorkflowEngineService', 'Failed to send timeout escalation notification', error);
    }
  }

  /**
   * Handle parallel execution steps
   * Executes multiple steps concurrently using Promise.all
   */
  private async handleParallelSteps(
    step: IWorkflowStep,
    context: IActionContext
  ): Promise<IActionResult> {
    const parallelStepIds = step.config.parallelStepIds || step.onComplete?.parallelStepIds || [];

    if (parallelStepIds.length === 0) {
      return { success: true, nextAction: 'continue' };
    }

    const instanceId = context.workflowInstance.Id;
    const startTime = Date.now();

    logger.info(
      'WorkflowEngineService',
      `Starting parallel execution of ${parallelStepIds.length} steps: ${parallelStepIds.join(', ')}`
    );

    // Get the workflow definition for step lookup
    const definition = await this.definitionService.getParsed(context.workflowInstance.WorkflowDefinitionId);

    // Create execution promises for all parallel steps
    const parallelPromises = parallelStepIds.map(async (stepId: string) => {
      try {
        const parallelStep = definition.steps.find(s => s.id === stepId);
        if (!parallelStep) {
          logger.warn('WorkflowEngineService', `Parallel step ${stepId} not found in definition`);
          return {
            stepId,
            success: false,
            error: `Step ${stepId} not found`
          };
        }

        // Mark step as in progress
        await this.instanceService.startStep(instanceId, stepId);

        // Build context for this parallel step
        const parallelContext: IActionContext = {
          ...context,
          currentStep: parallelStep,
          stepStatus: (await this.instanceService.getStepStatus(instanceId, stepId))!
        };

        // Execute the step action
        let result: IActionResult;

        switch (parallelStep.type) {
          case StepType.AssignTasks:
          case StepType.CreateTask:
          case StepType.Approval:
          case StepType.Action:
          case StepType.Notification:
          case StepType.SetVariable:
            result = await this.actionDispatcher.dispatch(parallelStep.type, parallelStep.config, parallelContext);
            break;

          case StepType.Webhook:
            result = await this.webhookHandler.execute(parallelStep, parallelContext);
            break;

          default:
            result = { success: true, nextAction: 'continue' };
        }

        // Mark step as complete or failed
        if (result.success) {
          await this.instanceService.completeStep(instanceId, stepId, { success: true }, result.outputVariables);
        } else {
          await this.instanceService.failStep(instanceId, stepId, result.error || 'Unknown error');
        }

        return {
          stepId,
          success: result.success,
          error: result.error,
          outputVariables: result.outputVariables,
          nextAction: result.nextAction,
          waitForItemIds: result.waitForItemIds
        };

      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        logger.error('WorkflowEngineService', `Parallel step ${stepId} failed`, error);

        await this.instanceService.failStep(instanceId, stepId, errorMessage);

        return {
          stepId,
          success: false,
          error: errorMessage,
          waitForItemIds: undefined
        };
      }
    });

    // Execute all steps in parallel and wait for all to complete
    const results = await Promise.all(parallelPromises);

    const executionTimeMs = Date.now() - startTime;
    const successCount = results.filter(r => r.success).length;
    const failedCount = results.filter(r => !r.success).length;

    logger.info(
      'WorkflowEngineService',
      `Parallel execution completed: ${successCount} succeeded, ${failedCount} failed (${executionTimeMs}ms)`
    );

    // Aggregate output variables from all parallel steps
    const aggregatedOutputs = results.reduce((acc, r) => {
      if (r.outputVariables) {
        acc[r.stepId] = r.outputVariables;
      }
      return acc;
    }, {} as Record<string, unknown>);

    // Check if any steps are waiting for tasks/approvals
    const waitingSteps = results.filter(r => r.nextAction === 'wait');
    if (waitingSteps.length > 0) {
      // Collect actual task IDs from waiting steps (if any)
      // Using reduce instead of flatMap for ES5 compatibility
      const waitingTaskIds: number[] = waitingSteps
        .map(s => s.waitForItemIds || [])
        .reduce<number[]>((acc, ids) => acc.concat(ids), [])
        .filter((id): id is number => typeof id === 'number');

      // Some parallel steps are waiting - workflow must wait
      return {
        success: true,
        nextAction: 'wait',
        waitForItemType: 'task',
        waitForItemIds: waitingTaskIds,
        outputVariables: {
          parallelSteps: parallelStepIds,
          executionMode: 'parallel',
          executionTimeMs,
          results: results.map(r => ({ stepId: r.stepId, success: r.success, error: r.error })),
          waitingSteps: waitingSteps.map(s => s.stepId),
          aggregatedOutputs
        }
      };
    }

    // All parallel steps completed
    // Determine success based on parallelConfig
    const failOnAny = step.config.failOnAnyError !== false; // Default to fail on any error
    const allSucceeded = failedCount === 0;

    if (!allSucceeded && failOnAny) {
      const failedSteps = results.filter(r => !r.success);
      return {
        success: false,
        error: `Parallel execution failed: ${failedSteps.map(s => `${s.stepId}: ${s.error}`).join('; ')}`,
        outputVariables: {
          parallelSteps: parallelStepIds,
          executionMode: 'parallel',
          executionTimeMs,
          results: results.map(r => ({ stepId: r.stepId, success: r.success, error: r.error })),
          aggregatedOutputs
        }
      };
    }

    return {
      success: true,
      nextAction: 'continue',
      outputVariables: {
        parallelSteps: parallelStepIds,
        executionMode: 'parallel',
        executionTimeMs,
        successCount,
        failedCount,
        results: results.map(r => ({ stepId: r.stepId, success: r.success, error: r.error })),
        aggregatedOutputs
      }
    };
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Get count of completed steps
   */
  private async getCompletedStepCount(instanceId: number): Promise<number> {
    const statuses = await this.instanceService.getStepStatuses(instanceId);
    return statuses.filter(s => s.Status === StepStatus.Completed || s.Status === StepStatus.Skipped).length;
  }

  /**
   * Estimate workflow duration based on steps
   */
  private estimateWorkflowDuration(steps: IWorkflowStep[]): number {
    let totalHours = 0;

    steps.forEach(step => {
      switch (step.type) {
        case StepType.Approval:
          totalHours += 24; // Assume 1 day for approvals
          break;
        case StepType.CreateTask:
        case StepType.AssignTasks:
          totalHours += step.config.dueDaysFromNow ? step.config.dueDaysFromNow * 8 : 8;
          break;
        case StepType.Wait:
          totalHours += step.config.waitHours || 24;
          break;
        default:
          totalHours += 0.5; // 30 minutes for other steps
      }
    });

    return totalHours;
  }

  /**
   * Create service container for action handlers
   */
  private createServiceContainer(): IServiceContainer {
    return {
      sp: this.sp,
      context: this.context,
      getService: <T>(serviceName: string): T => {
        // Service locator pattern - would be expanded based on needs
        switch (serviceName) {
          case 'WorkflowDefinitionService':
            return this.definitionService as unknown as T;
          case 'WorkflowInstanceService':
            return this.instanceService as unknown as T;
          default:
            throw new Error(`Unknown service: ${serviceName}`);
        }
      }
    };
  }

  // ============================================================================
  // DEAD LETTER QUEUE MANAGEMENT
  // ============================================================================

  /**
   * Get statistics for failed sync operations
   */
  public getSyncFailureStats(): { total: number; byType: Record<string, number> } {
    return workflowSyncDLQ.getStats();
  }

  /**
   * Get all failed sync operations
   */
  public getFailedSyncOperations(): Array<{
    id: string;
    operationType: string;
    payload: unknown;
    error: string;
    attempts: number;
    createdAt: Date;
    lastAttemptAt: Date;
  }> {
    return workflowSyncDLQ.getAll();
  }

  /**
   * Retry a specific failed sync operation from the DLQ
   */
  public async retryFailedSync(dlqItemId: string): Promise<IRetryResult<void>> {
    const items = workflowSyncDLQ.getAll();
    const item = items.find(i => i.id === dlqItemId);

    if (!item) {
      return {
        success: false,
        error: new Error(`DLQ item ${dlqItemId} not found`),
        attempts: 0,
        totalDurationMs: 0
      };
    }

    const payload = item.payload as {
      processId: number;
      workflowStatus: WorkflowInstanceStatus;
      workflowInstanceId: number;
    };

    // Update attempt count
    workflowSyncDLQ.updateAttempt(dlqItemId);

    // Retry the sync operation
    const result = await retryWithDLQ<void>(
      async () => {
        if (this.processSyncCallback) {
          await this.processSyncCallback(
            payload.processId,
            payload.workflowStatus,
            payload.workflowInstanceId
          );
        }
      },
      'process-status-sync-retry',
      payload,
      { ...PROCESS_SYNC_RETRY_OPTIONS, maxRetries: 1 }, // Single retry for manual retries
      workflowSyncDLQ,
      {
        source: 'WorkflowEngineService',
        operation: 'retryFailedSync',
        originalDlqId: dlqItemId
      }
    );

    if (result.success) {
      // Remove from DLQ on success
      workflowSyncDLQ.remove(dlqItemId);
      logger.info('WorkflowEngineService', `Successfully retried sync for DLQ item ${dlqItemId}`);
    }

    return result;
  }

  /**
   * Retry all failed sync operations
   */
  public async retryAllFailedSyncs(): Promise<{
    total: number;
    succeeded: number;
    failed: number;
    results: Array<{ id: string; success: boolean; error?: string }>
  }> {
    const items = workflowSyncDLQ.getByType('process-status-sync');
    const results: Array<{ id: string; success: boolean; error?: string }> = [];

    for (const item of items) {
      const result = await this.retryFailedSync(item.id);
      results.push({
        id: item.id,
        success: result.success,
        error: result.error?.message
      });
    }

    return {
      total: items.length,
      succeeded: results.filter(r => r.success).length,
      failed: results.filter(r => !r.success).length,
      results
    };
  }

  /**
   * Clear all items from the sync failure DLQ (use with caution)
   */
  public clearSyncFailureDLQ(): void {
    const stats = workflowSyncDLQ.getStats();
    workflowSyncDLQ.clear();
    logger.warn('WorkflowEngineService', `Cleared ${stats.total} items from sync failure DLQ`);
  }
}
