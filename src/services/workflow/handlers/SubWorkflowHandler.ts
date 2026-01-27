// @ts-nocheck
/**
 * SubWorkflowHandler
 *
 * Handles CallWorkflow step execution - invoking sub-workflows.
 * Supports variable mapping between parent and child workflows.
 *
 * @author JML Development Team
 * @version 1.0.0
 */

import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  IWorkflowStep,
  IWorkflowInstance,
  IActionContext,
  IActionResult,
  WorkflowInstanceStatus
} from '../../../models/IWorkflow';
import { WorkflowDefinitionService } from '../WorkflowDefinitionService';
import { WorkflowInstanceService } from '../WorkflowInstanceService';
import { logger } from '../../LoggingService';

/**
 * Result of sub-workflow execution
 */
export interface ISubWorkflowResult extends IActionResult {
  subWorkflowInstanceId?: number;
  subWorkflowStatus?: WorkflowInstanceStatus;
  mappedOutputVariables?: Record<string, unknown>;
}

/**
 * Handler for CallWorkflow step type
 */
export class SubWorkflowHandler {
  private sp: SPFI;
  private context: WebPartContext;
  private definitionService: WorkflowDefinitionService;
  private instanceService: WorkflowInstanceService;

  constructor(sp: SPFI, context: WebPartContext) {
    this.sp = sp;
    this.context = context;
    this.definitionService = new WorkflowDefinitionService(sp);
    this.instanceService = new WorkflowInstanceService(sp);
  }

  /**
   * Execute CallWorkflow step
   */
  public async execute(
    step: IWorkflowStep,
    parentContext: IActionContext,
    engineExecuteStep: (instanceId: number, stepId: string) => Promise<{ success: boolean; status: WorkflowInstanceStatus }>
  ): Promise<ISubWorkflowResult> {
    const config = step.config;

    // Validate configuration
    if (!config.subWorkflowCode && !config.subWorkflowId) {
      return {
        success: false,
        error: 'CallWorkflow step requires subWorkflowCode or subWorkflowId',
        nextAction: 'fail'
      };
    }

    try {
      // Get sub-workflow definition
      let definition;
      if (config.subWorkflowId) {
        definition = await this.definitionService.getById(config.subWorkflowId);
      } else if (config.subWorkflowCode) {
        definition = await this.definitionService.getByCode(config.subWorkflowCode);
      }

      if (!definition) {
        return {
          success: false,
          error: `Sub-workflow not found: ${config.subWorkflowCode || config.subWorkflowId}`,
          nextAction: 'fail'
        };
      }

      if (!definition.IsActive) {
        return {
          success: false,
          error: `Sub-workflow "${definition.Title}" is not active`,
          nextAction: 'fail'
        };
      }

      logger.info('SubWorkflowHandler', `Starting sub-workflow: ${definition.Title}`);

      // Map input variables from parent to child
      const childVariables = this.mapInputVariables(
        config.inputMappings || {},
        parentContext.variables,
        parentContext.process
      );

      // Parse definition steps
      const parsedDef = this.definitionService.parseDefinition(definition);

      // Create sub-workflow instance
      const subInstance = await this.instanceService.create({
        Title: `${definition.Title} (Sub-workflow of ${parentContext.workflowInstance.Id})`,
        WorkflowDefinitionId: definition.Id,
        ProcessId: parentContext.workflowInstance.ProcessId,
        Status: WorkflowInstanceStatus.Running,
        CurrentStepId: parsedDef.steps[0]?.id,
        CurrentStepName: parsedDef.steps[0]?.name,
        TotalSteps: parsedDef.steps.length,
        CompletedSteps: 0,
        ProgressPercentage: 0,
        Context: JSON.stringify({
          ...parentContext.process,
          parentWorkflowInstanceId: parentContext.workflowInstance.Id,
          parentStepId: step.id
        }),
        Variables: JSON.stringify(childVariables)
      });

      logger.info('SubWorkflowHandler', `Sub-workflow instance created: ${subInstance.Id}`);

      // If not waiting, return immediately
      if (config.waitForSubWorkflow === false) {
        return {
          success: true,
          nextAction: 'continue',
          subWorkflowInstanceId: subInstance.Id,
          subWorkflowStatus: WorkflowInstanceStatus.Running,
          outputVariables: {
            subWorkflowId: subInstance.Id,
            subWorkflowStatus: 'Running'
          }
        };
      }

      // Execute sub-workflow synchronously (wait for completion)
      const result = await this.executeSubWorkflowToCompletion(
        subInstance.Id,
        parsedDef.steps[0]?.id || '',
        engineExecuteStep
      );

      // Get final instance state
      const finalInstance = await this.instanceService.getById(subInstance.Id);

      // Map output variables from child to parent
      const mappedOutputs = this.mapOutputVariables(
        config.outputMappings || {},
        finalInstance.Variables ? JSON.parse(finalInstance.Variables) : {}
      );

      if (result.success) {
        logger.info('SubWorkflowHandler', `Sub-workflow completed successfully: ${subInstance.Id}`);

        return {
          success: true,
          nextAction: 'continue',
          subWorkflowInstanceId: subInstance.Id,
          subWorkflowStatus: WorkflowInstanceStatus.Completed,
          mappedOutputVariables: mappedOutputs,
          outputVariables: {
            subWorkflowId: subInstance.Id,
            subWorkflowStatus: 'Completed',
            ...mappedOutputs
          }
        };
      } else {
        logger.warn('SubWorkflowHandler', `Sub-workflow failed: ${subInstance.Id}`);

        return {
          success: false,
          error: `Sub-workflow failed: ${finalInstance.ErrorMessage || 'Unknown error'}`,
          nextAction: 'fail',
          subWorkflowInstanceId: subInstance.Id,
          subWorkflowStatus: finalInstance.Status
        };
      }

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      logger.error('SubWorkflowHandler', 'Error executing sub-workflow', error);

      return {
        success: false,
        error: `Failed to execute sub-workflow: ${errorMessage}`,
        nextAction: 'fail'
      };
    }
  }

  /**
   * Execute sub-workflow steps until completion or failure
   */
  private async executeSubWorkflowToCompletion(
    instanceId: number,
    startStepId: string,
    engineExecuteStep: (instanceId: number, stepId: string) => Promise<{ success: boolean; status: WorkflowInstanceStatus }>
  ): Promise<{ success: boolean; status: WorkflowInstanceStatus }> {
    const maxIterations = 100; // Safety limit
    let currentStepId = startStepId;
    let iterations = 0;

    while (iterations < maxIterations) {
      iterations++;

      const result = await engineExecuteStep(instanceId, currentStepId);

      // Check terminal states
      if (result.status === WorkflowInstanceStatus.Completed) {
        return { success: true, status: result.status };
      }

      if (result.status === WorkflowInstanceStatus.Failed ||
          result.status === WorkflowInstanceStatus.Cancelled) {
        return { success: false, status: result.status };
      }

      // If waiting, we can't continue synchronously
      if (result.status === WorkflowInstanceStatus.WaitingForApproval ||
          result.status === WorkflowInstanceStatus.WaitingForTask ||
          result.status === WorkflowInstanceStatus.WaitingForInput) {
        logger.warn('SubWorkflowHandler', 'Sub-workflow entered waiting state, cannot complete synchronously');
        return { success: false, status: result.status };
      }

      // Get current step ID from instance
      const instance = await this.instanceService.getById(instanceId);
      if (!instance.CurrentStepId || instance.CurrentStepId === currentStepId) {
        // No progress made, likely completed or stuck
        break;
      }

      currentStepId = instance.CurrentStepId;
    }

    logger.warn('SubWorkflowHandler', `Sub-workflow reached max iterations (${maxIterations})`);
    return { success: false, status: WorkflowInstanceStatus.Failed };
  }

  /**
   * Map parent variables to child input variables
   */
  private mapInputVariables(
    mappings: Record<string, string>,
    parentVariables: Record<string, unknown>,
    parentProcess: Record<string, unknown>
  ): Record<string, unknown> {
    const childVariables: Record<string, unknown> = {};

    for (const [childKey, parentPath] of Object.entries(mappings)) {
      const value = this.resolveVariablePath(parentPath, parentVariables, parentProcess);
      childVariables[childKey] = value;
    }

    return childVariables;
  }

  /**
   * Map child output variables back to parent
   */
  private mapOutputVariables(
    mappings: Record<string, string>,
    childVariables: Record<string, unknown>
  ): Record<string, unknown> {
    const parentVariables: Record<string, unknown> = {};

    for (const [parentKey, childPath] of Object.entries(mappings)) {
      const parts = childPath.split('.');
      let value: unknown = childVariables;

      for (const part of parts) {
        if (value === null || value === undefined) break;
        value = (value as Record<string, unknown>)[part];
      }

      parentVariables[parentKey] = value;
    }

    return parentVariables;
  }

  /**
   * Resolve variable path from parent context
   */
  private resolveVariablePath(
    path: string,
    variables: Record<string, unknown>,
    process: Record<string, unknown>
  ): unknown {
    const parts = path.split('.');

    // Check for prefixes
    if (parts[0] === 'variables') {
      let current: unknown = variables;
      for (let i = 1; i < parts.length; i++) {
        if (current === null || current === undefined) return undefined;
        current = (current as Record<string, unknown>)[parts[i]];
      }
      return current;
    }

    if (parts[0] === 'process') {
      let current: unknown = process;
      for (let i = 1; i < parts.length; i++) {
        if (current === null || current === undefined) return undefined;
        current = (current as Record<string, unknown>)[parts[i]];
      }
      return current;
    }

    // Default: try variables first, then process
    if (variables[parts[0]] !== undefined) {
      let current: unknown = variables;
      for (const part of parts) {
        if (current === null || current === undefined) return undefined;
        current = (current as Record<string, unknown>)[part];
      }
      return current;
    }

    return process[parts[0]];
  }

  /**
   * Resume a waiting sub-workflow (called when sub-workflow completes asynchronously)
   */
  public async onSubWorkflowCompleted(
    parentInstanceId: number,
    parentStepId: string,
    subWorkflowInstanceId: number,
    outputMappings: Record<string, string>
  ): Promise<IActionResult> {
    try {
      // Get sub-workflow final state
      const subInstance = await this.instanceService.getById(subWorkflowInstanceId);

      if (subInstance.Status !== WorkflowInstanceStatus.Completed) {
        return {
          success: false,
          error: `Sub-workflow did not complete successfully: ${subInstance.Status}`,
          nextAction: 'fail'
        };
      }

      // Map output variables
      const childVariables = subInstance.Variables ? JSON.parse(subInstance.Variables) : {};
      const mappedOutputs = this.mapOutputVariables(outputMappings, childVariables);

      return {
        success: true,
        nextAction: 'continue',
        outputVariables: {
          subWorkflowId: subWorkflowInstanceId,
          subWorkflowStatus: 'Completed',
          ...mappedOutputs
        }
      };

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      logger.error('SubWorkflowHandler', 'Error handling sub-workflow completion', error);

      return {
        success: false,
        error: errorMessage,
        nextAction: 'fail'
      };
    }
  }
}
