// @ts-nocheck
/* eslint-disable */
/**
 * WorkflowValidationService
 * Validates workflow definitions before publish/activation
 * Ensures workflow integrity and identifies configuration issues
 *
 * Phase 6: Validation & Selection
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import {
  IWorkflowDefinition,
  IWorkflowStep,
  IStepConfig,
  ITransition,
  IValidationResult,
  IValidationError,
  StepType,
  TransitionType,
  ActionType
} from '../../models/IWorkflow';
import { logger } from '../LoggingService';

/**
 * Validation severity levels
 */
export enum ValidationSeverity {
  Error = 'Error',       // Blocks publishing
  Warning = 'Warning',   // Can publish but should fix
  Info = 'Info'          // Best practice suggestion
}

/**
 * Extended validation error with severity
 */
export interface IExtendedValidationError extends IValidationError {
  severity: ValidationSeverity;
  stepName?: string;
  suggestion?: string;
}

/**
 * Comprehensive validation result
 */
export interface IComprehensiveValidationResult {
  valid: boolean;
  canPublish: boolean;
  errors: IExtendedValidationError[];
  warnings: IExtendedValidationError[];
  info: IExtendedValidationError[];
  summary: {
    totalIssues: number;
    errorCount: number;
    warningCount: number;
    infoCount: number;
    stepsValidated: number;
  };
}

/**
 * Step graph node for connectivity analysis
 */
interface IStepGraphNode {
  stepId: string;
  incomingConnections: string[];
  outgoingConnections: string[];
  visited: boolean;
  reachableFromStart: boolean;
  canReachEnd: boolean;
}

export class WorkflowValidationService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ============================================================================
  // MAIN VALIDATION METHODS
  // ============================================================================

  /**
   * Validate a workflow definition comprehensively
   */
  public async validateWorkflowDefinition(
    definition: IWorkflowDefinition
  ): Promise<IComprehensiveValidationResult> {
    const errors: IExtendedValidationError[] = [];
    const warnings: IExtendedValidationError[] = [];
    const info: IExtendedValidationError[] = [];

    try {
      // Parse steps from JSON
      const steps: IWorkflowStep[] = definition.Steps
        ? JSON.parse(definition.Steps)
        : [];

      // 1. Validate basic structure
      this.validateBasicStructure(definition, steps, errors, warnings);

      // 2. Validate each step
      for (const step of steps) {
        await this.validateStep(step, steps, errors, warnings, info);
      }

      // 3. Validate workflow connectivity
      this.validateConnectivity(steps, errors, warnings);

      // 4. Validate no circular references
      this.validateNoCycles(steps, errors);

      // 5. Validate variables
      this.validateVariables(definition, steps, errors, warnings);

      // 6. Business rules validation
      await this.validateBusinessRules(definition, steps, errors, warnings);

    } catch (error) {
      errors.push({
        code: 'PARSE_ERROR',
        message: `Failed to parse workflow definition: ${error instanceof Error ? error.message : 'Unknown error'}`,
        severity: ValidationSeverity.Error
      });
    }

    const errorCount = errors.length;
    const warningCount = warnings.length;
    const infoCount = info.length;

    return {
      valid: errorCount === 0,
      canPublish: errorCount === 0, // Can publish if no errors (warnings allowed)
      errors,
      warnings,
      info,
      summary: {
        totalIssues: errorCount + warningCount + infoCount,
        errorCount,
        warningCount,
        infoCount,
        stepsValidated: definition.Steps ? JSON.parse(definition.Steps).length : 0
      }
    };
  }

  /**
   * Quick validation for UI feedback (subset of full validation)
   */
  public validateStepQuick(step: IWorkflowStep): IValidationError[] {
    const errors: IValidationError[] = [];

    // Required fields
    if (!step.id) {
      errors.push({ code: 'MISSING_ID', message: 'Step ID is required', stepId: step.id });
    }
    if (!step.name) {
      errors.push({ code: 'MISSING_NAME', message: 'Step name is required', stepId: step.id });
    }
    if (!step.type) {
      errors.push({ code: 'MISSING_TYPE', message: 'Step type is required', stepId: step.id });
    }

    // Type-specific validation
    switch (step.type) {
      case StepType.AssignTasks:
      case StepType.CreateTask:
        if (!step.config.taskTemplateId && !step.config.assigneeRole && !step.config.assigneeId) {
          errors.push({
            code: 'TASK_NO_ASSIGNEE',
            message: 'Task step requires either a template, role, or specific assignee',
            stepId: step.id
          });
        }
        break;

      case StepType.Approval:
        if (!step.config.approverId && !step.config.approverField && !step.config.approverRole) {
          errors.push({
            code: 'APPROVAL_NO_APPROVER',
            message: 'Approval step requires an approver (ID, field, or role)',
            stepId: step.id
          });
        }
        break;

      case StepType.Notification:
        if (!step.config.messageTemplate) {
          errors.push({
            code: 'NOTIFICATION_NO_MESSAGE',
            message: 'Notification step requires a message template',
            stepId: step.id
          });
        }
        if (!step.config.recipientId && !step.config.recipientField && !step.config.recipientSource) {
          errors.push({
            code: 'NOTIFICATION_NO_RECIPIENT',
            message: 'Notification step requires a recipient',
            stepId: step.id
          });
        }
        break;

      case StepType.Condition:
        if (!step.config.conditionGroups && !step.config.conditions) {
          errors.push({
            code: 'CONDITION_NO_RULES',
            message: 'Condition step requires at least one condition',
            stepId: step.id
          });
        }
        break;

      case StepType.Webhook:
        if (!step.config.webhookUrl) {
          errors.push({
            code: 'WEBHOOK_NO_URL',
            message: 'Webhook step requires a URL',
            stepId: step.id
          });
        }
        break;

      case StepType.ForEach:
        if (!step.config.collectionPath) {
          errors.push({
            code: 'FOREACH_NO_COLLECTION',
            message: 'ForEach step requires a collection path',
            stepId: step.id
          });
        }
        break;

      case StepType.CallWorkflow:
        if (!step.config.subWorkflowCode && !step.config.subWorkflowId) {
          errors.push({
            code: 'SUBWORKFLOW_NO_REFERENCE',
            message: 'CallWorkflow step requires a workflow code or ID',
            stepId: step.id
          });
        }
        break;
    }

    return errors;
  }

  // ============================================================================
  // STRUCTURE VALIDATION
  // ============================================================================

  /**
   * Validate basic workflow structure
   */
  private validateBasicStructure(
    definition: IWorkflowDefinition,
    steps: IWorkflowStep[],
    errors: IExtendedValidationError[],
    warnings: IExtendedValidationError[]
  ): void {
    // Must have at least one step
    if (steps.length === 0) {
      errors.push({
        code: 'NO_STEPS',
        message: 'Workflow must have at least one step',
        severity: ValidationSeverity.Error
      });
      return;
    }

    // Must have exactly one Start step
    const startSteps = steps.filter(s => s.type === StepType.Start);
    if (startSteps.length === 0) {
      errors.push({
        code: 'NO_START',
        message: 'Workflow must have a Start step',
        severity: ValidationSeverity.Error
      });
    } else if (startSteps.length > 1) {
      errors.push({
        code: 'MULTIPLE_STARTS',
        message: 'Workflow can only have one Start step',
        severity: ValidationSeverity.Error
      });
    }

    // Must have at least one End step
    const endSteps = steps.filter(s => s.type === StepType.End);
    if (endSteps.length === 0) {
      warnings.push({
        code: 'NO_END',
        message: 'Workflow should have an End step for proper termination',
        severity: ValidationSeverity.Warning,
        suggestion: 'Add an End step to clearly define workflow completion'
      });
    }

    // Check for duplicate step IDs
    const stepIds = steps.map(s => s.id);
    const duplicates = stepIds.filter((id, index) => stepIds.indexOf(id) !== index);
    if (duplicates.length > 0) {
      const uniqueDuplicates = Array.from(new Set(duplicates));
      errors.push({
        code: 'DUPLICATE_IDS',
        message: `Duplicate step IDs found: ${uniqueDuplicates.join(', ')}`,
        severity: ValidationSeverity.Error
      });
    }

    // Workflow code format
    if (definition.WorkflowCode && !/^[A-Z0-9_-]+$/i.test(definition.WorkflowCode)) {
      warnings.push({
        code: 'INVALID_CODE_FORMAT',
        message: 'Workflow code should only contain letters, numbers, underscores, and hyphens',
        severity: ValidationSeverity.Warning
      });
    }

    // Version format
    if (definition.Version && !/^\d+\.\d+\.\d+$/.test(definition.Version)) {
      warnings.push({
        code: 'INVALID_VERSION_FORMAT',
        message: 'Version should follow semantic versioning (e.g., 1.0.0)',
        severity: ValidationSeverity.Warning
      });
    }
  }

  // ============================================================================
  // STEP VALIDATION
  // ============================================================================

  /**
   * Validate individual step
   */
  private async validateStep(
    step: IWorkflowStep,
    allSteps: IWorkflowStep[],
    errors: IExtendedValidationError[],
    warnings: IExtendedValidationError[],
    info: IExtendedValidationError[]
  ): Promise<void> {
    // Quick validation
    const quickErrors = this.validateStepQuick(step);
    for (const err of quickErrors) {
      errors.push({
        ...err,
        severity: ValidationSeverity.Error,
        stepName: step.name
      });
    }

    // Validate transition
    this.validateTransition(step, allSteps, errors, warnings);

    // Validate SLA configuration
    if (step.sla) {
      if (step.sla.warningHours >= step.sla.breachHours) {
        warnings.push({
          code: 'SLA_WARNING_AFTER_BREACH',
          message: `SLA warning (${step.sla.warningHours}h) should be less than breach time (${step.sla.breachHours}h)`,
          stepId: step.id,
          stepName: step.name,
          severity: ValidationSeverity.Warning
        });
      }
    }

    // Validate timeout configuration
    if (step.timeoutHours && step.timeoutHours < 0) {
      errors.push({
        code: 'INVALID_TIMEOUT',
        message: 'Timeout hours cannot be negative',
        stepId: step.id,
        stepName: step.name,
        severity: ValidationSeverity.Error
      });
    }

    // Type-specific deep validation
    await this.validateStepConfig(step, errors, warnings, info);
  }

  /**
   * Validate step configuration deeply
   */
  private async validateStepConfig(
    step: IWorkflowStep,
    errors: IExtendedValidationError[],
    warnings: IExtendedValidationError[],
    info: IExtendedValidationError[]
  ): Promise<void> {
    const config = step.config;

    switch (step.type) {
      case StepType.AssignTasks:
      case StepType.CreateTask:
        // Validate task template exists
        if (config.taskTemplateId) {
          const templateExists = await this.checkItemExists('PM_ChecklistTemplates', config.taskTemplateId);
          if (!templateExists) {
            errors.push({
              code: 'TASK_TEMPLATE_NOT_FOUND',
              message: `Task template ID ${config.taskTemplateId} does not exist`,
              stepId: step.id,
              stepName: step.name,
              severity: ValidationSeverity.Error
            });
          }
        }

        // Validate due days
        if (config.dueDaysFromNow !== undefined && config.dueDaysFromNow < 0) {
          errors.push({
            code: 'INVALID_DUE_DAYS',
            message: 'Due days from now cannot be negative',
            stepId: step.id,
            stepName: step.name,
            severity: ValidationSeverity.Error
          });
        }
        break;

      case StepType.Approval:
        // Validate approver configuration
        if (config.approverId && typeof config.approverId === 'number') {
          // Could validate user exists, but might be expensive
          info.push({
            code: 'APPROVER_SPECIFIED',
            message: `Fixed approver ID ${config.approverId} specified - ensure this user exists`,
            stepId: step.id,
            stepName: step.name,
            severity: ValidationSeverity.Info
          });
        }
        break;

      case StepType.Webhook:
        // Validate URL format
        if (config.webhookUrl) {
          try {
            new URL(config.webhookUrl);
          } catch {
            errors.push({
              code: 'INVALID_WEBHOOK_URL',
              message: 'Webhook URL is not a valid URL format',
              stepId: step.id,
              stepName: step.name,
              severity: ValidationSeverity.Error
            });
          }
        }

        // Validate HTTP method
        const validMethods = ['GET', 'POST', 'PUT', 'PATCH', 'DELETE'];
        if (config.webhookMethod && !validMethods.includes(config.webhookMethod)) {
          errors.push({
            code: 'INVALID_HTTP_METHOD',
            message: `Invalid HTTP method: ${config.webhookMethod}`,
            stepId: step.id,
            stepName: step.name,
            severity: ValidationSeverity.Error
          });
        }

        // Warn about no timeout
        if (!config.webhookTimeout) {
          warnings.push({
            code: 'NO_WEBHOOK_TIMEOUT',
            message: 'No timeout specified for webhook - may hang indefinitely',
            stepId: step.id,
            stepName: step.name,
            severity: ValidationSeverity.Warning,
            suggestion: 'Add a timeout (e.g., 30000ms) to prevent indefinite waiting'
          });
        }
        break;

      case StepType.CallWorkflow:
        // Validate sub-workflow exists
        if (config.subWorkflowCode) {
          const subworkflowExists = await this.checkWorkflowCodeExists(config.subWorkflowCode);
          if (!subworkflowExists) {
            errors.push({
              code: 'SUBWORKFLOW_NOT_FOUND',
              message: `Sub-workflow with code "${config.subWorkflowCode}" does not exist`,
              stepId: step.id,
              stepName: step.name,
              severity: ValidationSeverity.Error
            });
          }
        }
        break;

      case StepType.ForEach:
        // Validate inner steps
        if (config.innerSteps && config.innerSteps.length > 0) {
          for (const innerStep of config.innerSteps) {
            const innerErrors = this.validateStepQuick(innerStep);
            for (const err of innerErrors) {
              errors.push({
                ...err,
                code: `INNER_${err.code}`,
                message: `Inner step "${innerStep.name}": ${err.message}`,
                stepId: step.id,
                stepName: step.name,
                severity: ValidationSeverity.Error
              });
            }
          }
        } else {
          warnings.push({
            code: 'FOREACH_EMPTY',
            message: 'ForEach step has no inner steps defined',
            stepId: step.id,
            stepName: step.name,
            severity: ValidationSeverity.Warning
          });
        }
        break;
    }
  }

  /**
   * Validate step transition configuration
   */
  private validateTransition(
    step: IWorkflowStep,
    allSteps: IWorkflowStep[],
    errors: IExtendedValidationError[],
    warnings: IExtendedValidationError[]
  ): void {
    // End step doesn't need transition
    if (step.type === StepType.End) {
      return;
    }

    const transition = step.onComplete;

    // Must have a transition (except for End)
    if (!transition) {
      warnings.push({
        code: 'NO_TRANSITION',
        message: 'Step has no transition defined - will use order-based next step',
        stepId: step.id,
        stepName: step.name,
        severity: ValidationSeverity.Warning
      });
      return;
    }

    const stepIds = allSteps.map(s => s.id);

    switch (transition.type) {
      case TransitionType.Goto:
        if (!transition.targetStepId) {
          errors.push({
            code: 'GOTO_NO_TARGET',
            message: 'Goto transition requires a target step ID',
            stepId: step.id,
            stepName: step.name,
            severity: ValidationSeverity.Error
          });
        } else if (!stepIds.includes(transition.targetStepId)) {
          errors.push({
            code: 'GOTO_INVALID_TARGET',
            message: `Goto target "${transition.targetStepId}" does not exist in workflow`,
            stepId: step.id,
            stepName: step.name,
            severity: ValidationSeverity.Error
          });
        }
        break;

      case TransitionType.Branch:
        if (!transition.branches || transition.branches.length === 0) {
          errors.push({
            code: 'BRANCH_NO_PATHS',
            message: 'Branch transition requires at least one branch path',
            stepId: step.id,
            stepName: step.name,
            severity: ValidationSeverity.Error
          });
        } else {
          // Validate each branch
          for (const branch of transition.branches) {
            if (!branch.targetStepId) {
              errors.push({
                code: 'BRANCH_NO_TARGET',
                message: `Branch "${branch.name}" has no target step`,
                stepId: step.id,
                stepName: step.name,
                severity: ValidationSeverity.Error
              });
            } else if (!stepIds.includes(branch.targetStepId)) {
              errors.push({
                code: 'BRANCH_INVALID_TARGET',
                message: `Branch "${branch.name}" target "${branch.targetStepId}" does not exist`,
                stepId: step.id,
                stepName: step.name,
                severity: ValidationSeverity.Error
              });
            }
          }

          // Check for default branch
          const hasDefault = transition.branches.some(b => b.isDefault);
          if (!hasDefault) {
            warnings.push({
              code: 'BRANCH_NO_DEFAULT',
              message: 'Branch transition has no default path - may fail if no conditions match',
              stepId: step.id,
              stepName: step.name,
              severity: ValidationSeverity.Warning,
              suggestion: 'Add a default branch to handle unmatched conditions'
            });
          }
        }
        break;

      case TransitionType.Parallel:
        if (!transition.parallelStepIds || transition.parallelStepIds.length === 0) {
          errors.push({
            code: 'PARALLEL_NO_STEPS',
            message: 'Parallel transition requires at least one parallel step',
            stepId: step.id,
            stepName: step.name,
            severity: ValidationSeverity.Error
          });
        } else {
          for (const parallelId of transition.parallelStepIds) {
            if (!stepIds.includes(parallelId)) {
              errors.push({
                code: 'PARALLEL_INVALID_STEP',
                message: `Parallel step "${parallelId}" does not exist`,
                stepId: step.id,
                stepName: step.name,
                severity: ValidationSeverity.Error
              });
            }
          }
        }
        break;
    }
  }

  // ============================================================================
  // CONNECTIVITY VALIDATION
  // ============================================================================

  /**
   * Validate workflow connectivity - all steps reachable and can reach end
   */
  private validateConnectivity(
    steps: IWorkflowStep[],
    errors: IExtendedValidationError[],
    warnings: IExtendedValidationError[]
  ): void {
    if (steps.length === 0) return;

    // Build graph
    const graph = new Map<string, IStepGraphNode>();
    for (const step of steps) {
      graph.set(step.id, {
        stepId: step.id,
        incomingConnections: [],
        outgoingConnections: this.getOutgoingConnections(step),
        visited: false,
        reachableFromStart: false,
        canReachEnd: step.type === StepType.End
      });
    }

    // Build incoming connections
    for (const step of steps) {
      const outgoing = this.getOutgoingConnections(step);
      for (const targetId of outgoing) {
        const targetNode = graph.get(targetId);
        if (targetNode) {
          targetNode.incomingConnections.push(step.id);
        }
      }
    }

    // Find start step
    const startStep = steps.find(s => s.type === StepType.Start);
    if (!startStep) return;

    // Mark reachable from start (BFS)
    const queue: string[] = [startStep.id];
    while (queue.length > 0) {
      const currentId = queue.shift()!;
      const node = graph.get(currentId);
      if (!node || node.reachableFromStart) continue;

      node.reachableFromStart = true;
      for (const nextId of node.outgoingConnections) {
        queue.push(nextId);
      }
    }

    // Mark can reach end (reverse BFS)
    const endSteps = steps.filter(s => s.type === StepType.End);
    const reverseQueue: string[] = endSteps.map(s => s.id);
    while (reverseQueue.length > 0) {
      const currentId = reverseQueue.shift()!;
      const node = graph.get(currentId);
      if (!node) continue;

      for (const prevId of node.incomingConnections) {
        const prevNode = graph.get(prevId);
        if (prevNode && !prevNode.canReachEnd) {
          prevNode.canReachEnd = true;
          reverseQueue.push(prevId);
        }
      }
    }

    // Check for unreachable steps
    graph.forEach((node, stepId) => {
      const step = steps.find(s => s.id === stepId);

      if (!node.reachableFromStart && step?.type !== StepType.Start) {
        warnings.push({
          code: 'UNREACHABLE_STEP',
          message: `Step "${step?.name}" is not reachable from the Start step`,
          stepId,
          stepName: step?.name,
          severity: ValidationSeverity.Warning,
          suggestion: 'Either connect this step or remove it'
        });
      }

      if (!node.canReachEnd && step?.type !== StepType.End) {
        warnings.push({
          code: 'DEAD_END_STEP',
          message: `Step "${step?.name}" cannot reach any End step`,
          stepId,
          stepName: step?.name,
          severity: ValidationSeverity.Warning,
          suggestion: 'Add a transition to an End step or remove this step'
        });
      }
    });
  }

  /**
   * Get outgoing connections from a step
   */
  private getOutgoingConnections(step: IWorkflowStep): string[] {
    const connections: string[] = [];

    if (!step.onComplete) {
      // Will use order-based next - we'd need to calculate this
      return connections;
    }

    switch (step.onComplete.type) {
      case TransitionType.Next:
        // Order-based - would need step context to determine
        break;
      case TransitionType.Goto:
        if (step.onComplete.targetStepId) {
          connections.push(step.onComplete.targetStepId);
        }
        break;
      case TransitionType.Branch:
        if (step.onComplete.branches) {
          for (const branch of step.onComplete.branches) {
            if (branch.targetStepId) {
              connections.push(branch.targetStepId);
            }
          }
        }
        break;
      case TransitionType.Parallel:
        if (step.onComplete.parallelStepIds) {
          connections.push(...step.onComplete.parallelStepIds);
        }
        break;
      case TransitionType.End:
        // No outgoing connection
        break;
    }

    // Add timeout transition
    if (step.onTimeout?.targetStepId) {
      connections.push(step.onTimeout.targetStepId);
    }

    return connections;
  }

  // ============================================================================
  // CYCLE DETECTION
  // ============================================================================

  /**
   * Detect cycles in workflow (infinite loops)
   */
  private validateNoCycles(
    steps: IWorkflowStep[],
    errors: IExtendedValidationError[]
  ): void {
    const WHITE = 0; // Not visited
    const GRAY = 1;  // Being processed
    const BLACK = 2; // Finished

    const color = new Map<string, number>();
    const parent = new Map<string, string>();

    for (const step of steps) {
      color.set(step.id, WHITE);
    }

    const hasCycle = (stepId: string): boolean => {
      color.set(stepId, GRAY);

      const step = steps.find(s => s.id === stepId);
      if (!step) return false;

      const neighbors = this.getOutgoingConnections(step);

      for (const neighbor of neighbors) {
        if (color.get(neighbor) === GRAY) {
          // Found a cycle - build path
          const cyclePath = [neighbor, stepId];
          let current = stepId;
          while (parent.has(current) && parent.get(current) !== neighbor) {
            current = parent.get(current)!;
            cyclePath.push(current);
          }

          errors.push({
            code: 'CYCLE_DETECTED',
            message: `Circular reference detected: ${cyclePath.reverse().join(' -> ')}`,
            severity: ValidationSeverity.Error,
            suggestion: 'Remove one of the transitions in the cycle'
          });
          return true;
        }

        if (color.get(neighbor) === WHITE) {
          parent.set(neighbor, stepId);
          if (hasCycle(neighbor)) return true;
        }
      }

      color.set(stepId, BLACK);
      return false;
    };

    // Start DFS from each unvisited node
    for (const step of steps) {
      if (color.get(step.id) === WHITE) {
        hasCycle(step.id);
      }
    }
  }

  // ============================================================================
  // VARIABLE VALIDATION
  // ============================================================================

  /**
   * Validate workflow variables
   */
  private validateVariables(
    definition: IWorkflowDefinition,
    steps: IWorkflowStep[],
    errors: IExtendedValidationError[],
    warnings: IExtendedValidationError[]
  ): void {
    // Parse defined variables
    const definedVariables = new Set<string>();
    if (definition.Variables) {
      try {
        const variables = JSON.parse(definition.Variables);
        for (const v of variables) {
          definedVariables.add(v.name);
        }
      } catch {
        errors.push({
          code: 'INVALID_VARIABLES_JSON',
          message: 'Variables JSON is invalid',
          severity: ValidationSeverity.Error
        });
      }
    }

    // Find variables used in steps
    const usedVariables = new Set<string>();
    const variablePattern = /\{\{([^}]+)\}\}/g;

    // Helper function to extract variables from a string
    const extractVariables = (text: string): void => {
      let match: RegExpExecArray | null;
      while ((match = variablePattern.exec(text)) !== null) {
        usedVariables.add(match[1].trim());
      }
      // Reset lastIndex for next use
      variablePattern.lastIndex = 0;
    };

    for (const step of steps) {
      // Check message templates for variable usage
      if (step.config.messageTemplate) {
        extractVariables(step.config.messageTemplate);
      }

      // Check webhook body template
      if (step.config.webhookBodyTemplate) {
        extractVariables(step.config.webhookBodyTemplate);
      }

      // SetVariable step defines a variable
      if (step.type === StepType.SetVariable && step.config.variableName) {
        definedVariables.add(step.config.variableName);
      }
    }

    // System variables that are always available
    const systemVariablesList = [
      'currentDate', 'currentUser', 'workflowInstanceId', 'processId',
      'employeeName', 'employeeEmail', 'department', 'managerId', 'managerEmail',
      'startDate', 'processType'
    ];
    const systemVariables = new Set<string>(systemVariablesList);

    // Check for undefined variables
    Array.from(usedVariables).forEach(used => {
      if (!definedVariables.has(used) && !systemVariables.has(used)) {
        warnings.push({
          code: 'UNDEFINED_VARIABLE',
          message: `Variable "${used}" is used but not defined`,
          severity: ValidationSeverity.Warning,
          suggestion: `Add "${used}" to workflow variables or check for typos`
        });
      }
    });
  }

  // ============================================================================
  // BUSINESS RULES VALIDATION
  // ============================================================================

  /**
   * Validate business-specific rules
   */
  private async validateBusinessRules(
    definition: IWorkflowDefinition,
    steps: IWorkflowStep[],
    errors: IExtendedValidationError[],
    warnings: IExtendedValidationError[]
  ): Promise<void> {
    // Check if workflow has meaningful actions
    const actionSteps = steps.filter(s =>
      s.type !== StepType.Start &&
      s.type !== StepType.End &&
      s.type !== StepType.Condition
    );

    if (actionSteps.length === 0) {
      warnings.push({
        code: 'NO_ACTIONS',
        message: 'Workflow has no action steps - will do nothing',
        severity: ValidationSeverity.Warning
      });
    }

    // Check for approval before tasks (common pattern issue)
    const taskSteps = steps.filter(s =>
      s.type === StepType.AssignTasks || s.type === StepType.CreateTask
    );
    const approvalSteps = steps.filter(s => s.type === StepType.Approval);

    if (taskSteps.length > 0 && approvalSteps.length > 0) {
      const firstTask = taskSteps.sort((a, b) => a.order - b.order)[0];
      const firstApproval = approvalSteps.sort((a, b) => a.order - b.order)[0];

      if (firstTask.order < firstApproval.order) {
        // Tasks assigned before approval - this might be intentional, just inform
        // Actually, this could be valid for "pre-approval tasks", so just skip this check
      }
    }

    // Check default workflow conflict
    if (definition.IsDefault) {
      const existingDefault = await this.checkExistingDefaultWorkflow(
        definition.ProcessType,
        definition.Id
      );

      if (existingDefault) {
        warnings.push({
          code: 'EXISTING_DEFAULT',
          message: `Another workflow is already set as default for ${definition.ProcessType} processes`,
          severity: ValidationSeverity.Warning,
          suggestion: 'Publishing this as default will replace the current default workflow'
        });
      }
    }
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Check if an item exists in a list
   */
  private async checkItemExists(listName: string, itemId: number): Promise<boolean> {
    try {
      await this.sp.web.lists.getByTitle(listName).items.getById(itemId).select('Id')();
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Check if a workflow code exists
   */
  private async checkWorkflowCodeExists(code: string): Promise<boolean> {
    try {
      const items = await this.sp.web.lists.getByTitle('PM_WorkflowDefinitions').items
        .filter(`WorkflowCode eq '${code}' and IsActive eq true`)
        .select('Id')
        .top(1)();
      return items.length > 0;
    } catch {
      return false;
    }
  }

  /**
   * Check if there's an existing default workflow for a process type
   */
  private async checkExistingDefaultWorkflow(
    processType: string,
    excludeId?: number
  ): Promise<boolean> {
    try {
      let filter = `ProcessType eq '${processType}' and IsDefault eq true and IsActive eq true`;
      if (excludeId) {
        filter += ` and Id ne ${excludeId}`;
      }

      const items = await this.sp.web.lists.getByTitle('PM_WorkflowDefinitions').items
        .filter(filter)
        .select('Id')
        .top(1)();

      return items.length > 0;
    } catch {
      return false;
    }
  }

  // ============================================================================
  // CONVERSION HELPERS
  // ============================================================================

  /**
   * Convert extended validation result to simple validation result
   */
  public toSimpleResult(result: IComprehensiveValidationResult): IValidationResult {
    return {
      valid: result.valid,
      errors: result.errors.map(e => ({
        field: e.field,
        stepId: e.stepId,
        code: e.code,
        message: e.message
      }))
    };
  }
}
