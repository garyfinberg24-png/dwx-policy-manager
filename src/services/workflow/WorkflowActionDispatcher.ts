// @ts-nocheck
/* eslint-disable */
/**
 * WorkflowActionDispatcher
 * Dispatches workflow actions to appropriate handlers
 * Central hub for executing different action types
 */

import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  IStepConfig,
  IActionContext,
  IActionResult,
  StepType,
  ActionType
} from '../../models/IWorkflow';
import { WorkflowConditionEvaluator } from './WorkflowConditionEvaluator';
import { logger } from '../LoggingService';

// Import action handlers
import { TaskActionHandler } from './handlers/TaskActionHandler';
import { ApprovalActionHandler } from './handlers/ApprovalActionHandler';
import { NotificationActionHandler } from './handlers/NotificationActionHandler';
import { ListActionHandler } from './handlers/ListActionHandler';
import { AzureADActionHandler } from './handlers/AzureADActionHandler';

/**
 * Interface for action handlers
 */
export interface IActionHandler {
  execute(config: IStepConfig, context: IActionContext): Promise<IActionResult>;
}

export class WorkflowActionDispatcher {
  private sp: SPFI;
  private webPartContext: WebPartContext;
  private conditionEvaluator: WorkflowConditionEvaluator;

  // Handler instances
  private taskHandler: TaskActionHandler;
  private approvalHandler: ApprovalActionHandler;
  private notificationHandler: NotificationActionHandler;
  private listHandler: ListActionHandler;
  private azureADHandler: AzureADActionHandler;

  constructor(sp: SPFI, context: WebPartContext) {
    this.sp = sp;
    this.webPartContext = context;
    this.conditionEvaluator = new WorkflowConditionEvaluator();

    // Initialize handlers
    this.taskHandler = new TaskActionHandler(sp);
    this.approvalHandler = new ApprovalActionHandler(sp);
    this.notificationHandler = new NotificationActionHandler(sp, context);
    this.listHandler = new ListActionHandler(sp);
    this.azureADHandler = new AzureADActionHandler(context);
  }

  // ============================================================================
  // DISPATCH METHODS
  // ============================================================================

  /**
   * Dispatch action based on step type
   */
  public async dispatch(
    stepType: StepType,
    config: IStepConfig,
    context: IActionContext
  ): Promise<IActionResult> {
    try {
      // Process config - replace tokens with actual values
      const processedConfig = this.processConfig(config, context);

      switch (stepType) {
        case StepType.CreateTask:
          return await this.taskHandler.createTask(processedConfig, context);

        case StepType.AssignTasks:
          return await this.taskHandler.assignTasksFromTemplate(processedConfig, context);

        case StepType.WaitForTasks:
          return await this.taskHandler.waitForTasks(processedConfig, context);

        case StepType.Approval:
          return await this.approvalHandler.createApproval(processedConfig, context);

        case StepType.Notification:
          return await this.notificationHandler.sendNotification(processedConfig, context);

        case StepType.Action:
          return await this.dispatchAction(processedConfig, context);

        case StepType.SetVariable:
          return await this.setVariable(processedConfig, context);

        case StepType.Wait:
          return await this.handleWait(processedConfig, context);

        default:
          logger.warn('WorkflowActionDispatcher', `Unhandled step type: ${stepType}`);
          return { success: true, nextAction: 'continue' };
      }
    } catch (error) {
      logger.error('WorkflowActionDispatcher', `Error dispatching action for ${stepType}`, error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error occurred'
      };
    }
  }

  /**
   * Dispatch action by action type
   */
  private async dispatchAction(
    config: IStepConfig,
    context: IActionContext
  ): Promise<IActionResult> {
    const actionType = config.actionType;
    const actionConfig = config.actionConfig || {};

    if (!actionType) {
      return { success: false, error: 'Action type not specified' };
    }

    switch (actionType) {
      case ActionType.CreateTask:
        return await this.taskHandler.createTask(config, context);

      case ActionType.AssignTasksFromTemplate:
        return await this.taskHandler.assignTasksFromTemplate(config, context);

      case ActionType.UpdateTaskStatus:
        return await this.taskHandler.updateTaskStatus(actionConfig as Record<string, unknown>, context);

      case ActionType.CreateApproval:
        return await this.approvalHandler.createApproval(config, context);

      case ActionType.SendNotification:
        return await this.notificationHandler.sendNotification(config, context);

      case ActionType.SendEmail:
        return await this.notificationHandler.sendEmail(actionConfig, context);

      case ActionType.UpdateListItem:
        return await this.listHandler.updateItem(actionConfig, context);

      case ActionType.CreateListItem:
        return await this.listHandler.createItem(actionConfig, context);

      case ActionType.SetVariable:
        return await this.setVariable(config, context);

      case ActionType.CallWebhook:
        return await this.callWebhook(actionConfig as Record<string, unknown>, context);

      case ActionType.SendTeamsMessage:
        return await this.notificationHandler.sendTeamsMessage(actionConfig, context);

      case ActionType.Wait:
        return await this.handleWait(config, context);

      // Azure AD / Entra ID Actions (IT Provisioning/Deprovisioning)
      case ActionType.DisableUserAccount:
        return await this.azureADHandler.disableUserAccount(actionConfig as Record<string, unknown>, context);

      case ActionType.EnableUserAccount:
        return await this.azureADHandler.enableUserAccount(actionConfig as Record<string, unknown>, context);

      case ActionType.AddUserToGroup:
        return await this.azureADHandler.addUserToGroup(actionConfig as Record<string, unknown>, context);

      case ActionType.RemoveUserFromGroup:
        return await this.azureADHandler.removeUserFromGroup(actionConfig as Record<string, unknown>, context);

      case ActionType.UpdateUserProfile:
        return await this.azureADHandler.updateUserProfile(actionConfig as Record<string, unknown>, context);

      case ActionType.ReclaimLicense:
        return await this.azureADHandler.revokeLicenses(actionConfig as Record<string, unknown>, context);

      // Asset & Equipment Actions
      case ActionType.CreateEquipmentRequest:
        return await this.listHandler.createItem({
          listName: 'JML_EquipmentRequests',
          ...actionConfig
        }, context);

      case ActionType.CreateAssetReturnRequest:
        return await this.listHandler.createItem({
          listName: 'JML_AssetReturns',
          ...actionConfig
        }, context);

      default:
        logger.warn('WorkflowActionDispatcher', `Unhandled action type: ${actionType}`);
        return { success: false, error: `Unknown action type: ${actionType}` };
    }
  }

  // ============================================================================
  // BUILT-IN ACTION HANDLERS
  // ============================================================================

  /**
   * Set workflow variable
   */
  private async setVariable(
    config: IStepConfig,
    context: IActionContext
  ): Promise<IActionResult> {
    try {
      const variableName = config.variableName;
      if (!variableName) {
        return { success: false, error: 'Variable name not specified' };
      }

      let value: unknown;

      if (config.variableExpression) {
        // Evaluate expression
        value = this.conditionEvaluator.evaluateExpression(
          config.variableExpression,
          { ...context.process, ...context.variables }
        );
      } else {
        value = config.variableValue;
      }

      return {
        success: true,
        nextAction: 'continue',
        outputVariables: {
          [variableName]: value
        }
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to set variable'
      };
    }
  }

  /**
   * Handle wait step
   */
  private async handleWait(
    config: IStepConfig,
    context: IActionContext
  ): Promise<IActionResult> {
    try {
      let waitUntil: Date | undefined;

      if (config.waitHours) {
        waitUntil = new Date();
        waitUntil.setHours(waitUntil.getHours() + config.waitHours);
      } else if (config.waitUntilField) {
        const fieldValue = context.process[config.waitUntilField] || context.variables[config.waitUntilField];
        if (fieldValue) {
          waitUntil = new Date(fieldValue as string);
        }
      }

      if (waitUntil) {
        // Check if wait time has passed
        if (waitUntil <= new Date()) {
          return { success: true, nextAction: 'continue' };
        }

        // Schedule resume
        return {
          success: true,
          nextAction: 'wait',
          outputVariables: {
            waitUntil: waitUntil.toISOString(),
            waitType: 'time'
          }
        };
      }

      return { success: true, nextAction: 'continue' };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to process wait'
      };
    }
  }

  /**
   * Call external webhook
   */
  private async callWebhook(
    config: Record<string, unknown>,
    context: IActionContext
  ): Promise<IActionResult> {
    try {
      const url = config.url as string;
      const method = (config.method as string) || 'POST';
      const headers = (config.headers as Record<string, string>) || {};
      const bodyTemplate = config.bodyTemplate as string;

      if (!url) {
        return { success: false, error: 'Webhook URL not specified' };
      }

      // Build request body
      let body: string | undefined;
      if (bodyTemplate) {
        body = this.conditionEvaluator.replaceTokens(
          bodyTemplate,
          { ...context.process, ...context.variables }
        );
      } else {
        // Default body with context data
        body = JSON.stringify({
          workflowInstanceId: context.workflowInstance.Id,
          stepId: context.currentStep.id,
          stepName: context.currentStep.name,
          processId: context.workflowInstance.ProcessId,
          timestamp: new Date().toISOString()
        });
      }

      // Make HTTP request
      const response = await fetch(url, {
        method,
        headers: {
          'Content-Type': 'application/json',
          ...headers
        },
        body: method !== 'GET' ? body : undefined
      });

      if (!response.ok) {
        return {
          success: false,
          error: `Webhook failed: ${response.status} ${response.statusText}`
        };
      }

      let responseData: unknown;
      try {
        responseData = await response.json();
      } catch {
        responseData = await response.text();
      }

      return {
        success: true,
        nextAction: 'continue',
        outputVariables: {
          webhookResponse: responseData,
          webhookStatus: response.status
        }
      };
    } catch (error) {
      logger.error('WorkflowActionDispatcher', 'Webhook call failed', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Webhook call failed'
      };
    }
  }

  // ============================================================================
  // CONFIG PROCESSING
  // ============================================================================

  /**
   * Process config - replace tokens and resolve field references
   */
  private processConfig(config: IStepConfig, context: IActionContext): IStepConfig {
    const evalContext = {
      ...context.process,
      ...context.variables,
      workflowInstance: context.workflowInstance,
      currentStep: context.currentStep
    };

    const processed: IStepConfig = { ...config };

    // Process string fields that might contain tokens
    if (processed.taskTitle) {
      processed.taskTitle = this.conditionEvaluator.replaceTokens(processed.taskTitle, evalContext);
    }
    if (processed.messageTemplate) {
      processed.messageTemplate = this.conditionEvaluator.replaceTokens(processed.messageTemplate, evalContext);
    }

    // Resolve field references for assignee
    if (processed.assigneeField) {
      const assigneeId = this.conditionEvaluator.evaluateExpression(processed.assigneeField, evalContext);
      if (typeof assigneeId === 'number') {
        processed.assigneeId = assigneeId;
      }
    }

    // Resolve field references for approver
    if (processed.approverField) {
      const approverId = this.conditionEvaluator.evaluateExpression(processed.approverField, evalContext);
      if (typeof approverId === 'number') {
        processed.approverId = approverId;
      }
    }

    // Resolve field references for recipient
    if (processed.recipientField) {
      const recipientId = this.conditionEvaluator.evaluateExpression(processed.recipientField, evalContext);
      if (typeof recipientId === 'number') {
        processed.recipientId = recipientId;
      }
    }

    // Calculate due date
    if (processed.dueDaysField) {
      const dueDays = this.conditionEvaluator.evaluateExpression(processed.dueDaysField, evalContext);
      if (typeof dueDays === 'number') {
        processed.dueDaysFromNow = dueDays;
      }
    }

    return processed;
  }

  // ============================================================================
  // HANDLER REGISTRATION
  // ============================================================================

  /**
   * Register a custom action handler
   */
  public registerHandler(actionType: string, handler: IActionHandler): void {
    // Store custom handlers for extension
    logger.info('WorkflowActionDispatcher', `Registered custom handler for: ${actionType}`);
  }
}
