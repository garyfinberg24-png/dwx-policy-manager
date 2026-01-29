// @ts-nocheck
/* eslint-disable */
/**
 * WorkflowDefinitionService
 * Service for managing workflow definitions in SharePoint
 * Handles CRUD operations for workflow templates/blueprints
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import {
  IWorkflowDefinition,
  IWorkflowDefinitionSummary,
  IWorkflowStep,
  IWorkflowVariable,
  ITriggerCondition,
  WorkflowStatus,
  IValidationResult,
  IValidationError,
  StepType,
  TransitionType
} from '../../models/IWorkflow';
import { ProcessType } from '../../models/ICommon';
import { logger } from '../LoggingService';

// List name constant
const LIST_NAME = 'PM_WorkflowDefinitions';

// Default select fields for workflow definitions
const SELECT_FIELDS = [
  'Id', 'Title', 'WorkflowCode', 'Description', 'Version',
  'ProcessType', 'IsActive', 'IsDefault',
  'TriggerConditions', 'Steps', 'Variables',
  'Category', 'Tags', 'EstimatedDuration',
  'TimesUsed', 'AverageCompletionTime', 'SuccessRate',
  'PublishedDate', 'PublishedById',
  'Created', 'Modified', 'Author/Id', 'Author/Title', 'Editor/Id', 'Editor/Title'
].join(',');

/**
 * Parsed workflow definition with typed JSON fields
 */
export interface IParsedWorkflowDefinition extends Omit<IWorkflowDefinition, 'Steps' | 'Variables' | 'TriggerConditions'> {
  steps: IWorkflowStep[];
  variables: IWorkflowVariable[];
  triggerConditions: ITriggerCondition[];
}

export class WorkflowDefinitionService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ============================================================================
  // READ OPERATIONS
  // ============================================================================

  /**
   * Get all workflow definitions
   */
  public async getAll(includeInactive: boolean = false): Promise<IWorkflowDefinition[]> {
    try {
      let query = this.sp.web.lists.getByTitle(LIST_NAME).items
        .select(SELECT_FIELDS)
        .expand('Author', 'Editor')
        .orderBy('Title', true);

      if (!includeInactive) {
        query = query.filter('IsActive eq true');
      }

      const items = await query();
      return items as IWorkflowDefinition[];
    } catch (error) {
      logger.error('WorkflowDefinitionService', 'Error fetching workflow definitions', error);
      throw new Error('Unable to retrieve workflow definitions. Please try again.');
    }
  }

  /**
   * Get workflow definitions summary for listings
   */
  public async getSummaries(processType?: ProcessType): Promise<IWorkflowDefinitionSummary[]> {
    try {
      let query = this.sp.web.lists.getByTitle(LIST_NAME).items
        .select('Id', 'Title', 'WorkflowCode', 'ProcessType', 'Version', 'IsActive', 'IsDefault', 'Steps', 'TimesUsed', 'SuccessRate')
        .filter('IsActive eq true')
        .orderBy('Title', true);

      if (processType) {
        query = query.filter(`ProcessType eq '${processType}'`);
      }

      const items = await query();

      return items.map((item: IWorkflowDefinition) => ({
        Id: item.Id,
        Title: item.Title,
        WorkflowCode: item.WorkflowCode,
        ProcessType: item.ProcessType,
        Version: item.Version,
        IsActive: item.IsActive,
        IsDefault: item.IsDefault,
        StepCount: this.countSteps(item.Steps),
        TimesUsed: item.TimesUsed || 0,
        SuccessRate: item.SuccessRate
      }));
    } catch (error) {
      logger.error('WorkflowDefinitionService', 'Error fetching workflow summaries', error);
      throw new Error('Unable to retrieve workflow summaries. Please try again.');
    }
  }

  /**
   * Get workflow definition by ID
   */
  public async getById(id: number): Promise<IWorkflowDefinition> {
    try {
      const item = await this.sp.web.lists.getByTitle(LIST_NAME).items
        .getById(id)
        .select(SELECT_FIELDS)
        .expand('Author', 'Editor')();

      return item as IWorkflowDefinition;
    } catch (error) {
      logger.error('WorkflowDefinitionService', `Error fetching workflow definition ${id}`, error);
      throw new Error('Workflow definition not found.');
    }
  }

  /**
   * Get workflow definition by code
   */
  public async getByCode(code: string): Promise<IWorkflowDefinition | undefined> {
    try {
      const items = await this.sp.web.lists.getByTitle(LIST_NAME).items
        .select(SELECT_FIELDS)
        .expand('Author', 'Editor')
        .filter(`WorkflowCode eq '${code}'`)
        .top(1)();

      return items.length > 0 ? items[0] as IWorkflowDefinition : undefined;
    } catch (error) {
      logger.error('WorkflowDefinitionService', `Error fetching workflow by code ${code}`, error);
      throw new Error('Unable to retrieve workflow definition.');
    }
  }

  /**
   * Get default workflow for a process type
   */
  public async getDefaultForProcessType(processType: ProcessType): Promise<IWorkflowDefinition | undefined> {
    try {
      const items = await this.sp.web.lists.getByTitle(LIST_NAME).items
        .select(SELECT_FIELDS)
        .expand('Author', 'Editor')
        .filter(`ProcessType eq '${processType}' and IsDefault eq true and IsActive eq true`)
        .top(1)();

      return items.length > 0 ? items[0] as IWorkflowDefinition : undefined;
    } catch (error) {
      logger.error('WorkflowDefinitionService', `Error fetching default workflow for ${processType}`, error);
      throw new Error('Unable to retrieve default workflow.');
    }
  }

  /**
   * Get workflows by process type
   */
  public async getByProcessType(processType: ProcessType): Promise<IWorkflowDefinition[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(LIST_NAME).items
        .select(SELECT_FIELDS)
        .expand('Author', 'Editor')
        .filter(`ProcessType eq '${processType}' and IsActive eq true`)
        .orderBy('Title', true)();

      return items as IWorkflowDefinition[];
    } catch (error) {
      logger.error('WorkflowDefinitionService', `Error fetching workflows for ${processType}`, error);
      throw new Error('Unable to retrieve workflows.');
    }
  }

  /**
   * Get parsed workflow definition with typed JSON fields
   */
  public async getParsed(id: number): Promise<IParsedWorkflowDefinition> {
    const definition = await this.getById(id);
    return this.parseDefinition(definition);
  }

  // ============================================================================
  // CREATE/UPDATE OPERATIONS
  // ============================================================================

  /**
   * Create new workflow definition
   */
  public async create(definition: Partial<IWorkflowDefinition>): Promise<IWorkflowDefinition> {
    try {
      // Validate before saving
      const validation = this.validateDefinition(definition);
      if (!validation.valid) {
        throw new Error(`Validation failed: ${validation.errors.map(e => e.message).join(', ')}`);
      }

      // Ensure JSON fields are stringified
      const itemData = {
        Title: definition.Title,
        WorkflowCode: definition.WorkflowCode,
        Description: definition.Description,
        Version: definition.Version || '1.0.0',
        ProcessType: definition.ProcessType,
        IsActive: definition.IsActive ?? false,
        IsDefault: definition.IsDefault ?? false,
        TriggerConditions: typeof definition.TriggerConditions === 'string'
          ? definition.TriggerConditions
          : JSON.stringify(definition.TriggerConditions || []),
        Steps: typeof definition.Steps === 'string'
          ? definition.Steps
          : JSON.stringify(definition.Steps || []),
        Variables: typeof definition.Variables === 'string'
          ? definition.Variables
          : JSON.stringify(definition.Variables || []),
        Category: definition.Category,
        Tags: definition.Tags,
        EstimatedDuration: definition.EstimatedDuration,
        TimesUsed: 0,
        AverageCompletionTime: 0,
        SuccessRate: 0
      };

      const result = await this.sp.web.lists.getByTitle(LIST_NAME).items.add(itemData);

      logger.info('WorkflowDefinitionService', `Created workflow definition: ${result.data.Id}`);
      return await this.getById(result.data.Id);
    } catch (error) {
      logger.error('WorkflowDefinitionService', 'Error creating workflow definition', error);
      throw error;
    }
  }

  /**
   * Update workflow definition
   */
  public async update(id: number, updates: Partial<IWorkflowDefinition>): Promise<void> {
    try {
      // If Steps are being updated, validate the full definition
      if (updates.Steps !== undefined) {
        // Get existing definition and merge with updates for validation
        const existing = await this.getById(id);
        const mergedDefinition = {
          ...existing,
          ...updates,
          Steps: updates.Steps
        };

        const validation = this.validateDefinition(mergedDefinition);
        if (!validation.valid) {
          throw new Error(`Validation failed: ${validation.errors.map(e => e.message).join(', ')}`);
        }
      }

      // If Steps or Variables are provided as objects, stringify them
      const itemData: Record<string, unknown> = { ...updates };

      if (updates.Steps !== undefined && typeof updates.Steps !== 'string') {
        itemData.Steps = JSON.stringify(updates.Steps);
      }
      if (updates.Variables !== undefined && typeof updates.Variables !== 'string') {
        itemData.Variables = JSON.stringify(updates.Variables);
      }
      if (updates.TriggerConditions !== undefined && typeof updates.TriggerConditions !== 'string') {
        itemData.TriggerConditions = JSON.stringify(updates.TriggerConditions);
      }

      await this.sp.web.lists.getByTitle(LIST_NAME).items
        .getById(id)
        .update(itemData);

      logger.info('WorkflowDefinitionService', `Updated workflow definition: ${id}`);
    } catch (error) {
      logger.error('WorkflowDefinitionService', `Error updating workflow definition ${id}`, error);
      throw error instanceof Error && error.message.includes('Validation failed')
        ? error
        : new Error('Unable to update workflow definition.');
    }
  }

  /**
   * Publish workflow definition (make active)
   */
  public async publish(id: number, userId: number): Promise<void> {
    try {
      // Validate the workflow before publishing
      const definition = await this.getById(id);
      const validation = this.validateDefinition(definition);

      if (!validation.valid) {
        throw new Error(`Cannot publish: ${validation.errors.map(e => e.message).join(', ')}`);
      }

      await this.sp.web.lists.getByTitle(LIST_NAME).items
        .getById(id)
        .update({
          IsActive: true,
          PublishedDate: new Date().toISOString(),
          PublishedById: userId
        });

      logger.info('WorkflowDefinitionService', `Published workflow definition: ${id}`);
    } catch (error) {
      logger.error('WorkflowDefinitionService', `Error publishing workflow definition ${id}`, error);
      throw error;
    }
  }

  /**
   * Unpublish workflow definition (make inactive)
   */
  public async unpublish(id: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(LIST_NAME).items
        .getById(id)
        .update({
          IsActive: false
        });

      logger.info('WorkflowDefinitionService', `Unpublished workflow definition: ${id}`);
    } catch (error) {
      logger.error('WorkflowDefinitionService', `Error unpublishing workflow definition ${id}`, error);
      throw new Error('Unable to unpublish workflow definition.');
    }
  }

  /**
   * Set as default workflow for process type
   */
  public async setAsDefault(id: number): Promise<void> {
    try {
      const definition = await this.getById(id);

      // Remove default flag from other workflows of same type
      const existingDefaults = await this.sp.web.lists.getByTitle(LIST_NAME).items
        .filter(`ProcessType eq '${definition.ProcessType}' and IsDefault eq true`)
        .select('Id')();

      for (const item of existingDefaults) {
        if (item.Id !== id) {
          await this.sp.web.lists.getByTitle(LIST_NAME).items
            .getById(item.Id)
            .update({ IsDefault: false });
        }
      }

      // Set this one as default
      await this.sp.web.lists.getByTitle(LIST_NAME).items
        .getById(id)
        .update({ IsDefault: true });

      logger.info('WorkflowDefinitionService', `Set workflow ${id} as default for ${definition.ProcessType}`);
    } catch (error) {
      logger.error('WorkflowDefinitionService', `Error setting default workflow ${id}`, error);
      throw new Error('Unable to set default workflow.');
    }
  }

  /**
   * Clone workflow definition
   */
  public async clone(id: number, newTitle: string, newCode: string): Promise<IWorkflowDefinition> {
    try {
      const original = await this.getById(id);

      const cloned: Partial<IWorkflowDefinition> = {
        Title: newTitle,
        WorkflowCode: newCode,
        Description: original.Description,
        Version: '1.0.0',
        ProcessType: original.ProcessType,
        IsActive: false,
        IsDefault: false,
        TriggerConditions: original.TriggerConditions,
        Steps: original.Steps,
        Variables: original.Variables,
        Category: original.Category,
        Tags: original.Tags,
        EstimatedDuration: original.EstimatedDuration
      };

      return await this.create(cloned);
    } catch (error) {
      logger.error('WorkflowDefinitionService', `Error cloning workflow definition ${id}`, error);
      throw new Error('Unable to clone workflow definition.');
    }
  }

  /**
   * Create new version of workflow
   */
  public async createVersion(id: number, newVersion: string): Promise<IWorkflowDefinition> {
    try {
      const original = await this.getById(id);

      // Generate new code with version suffix
      const baseCode = original.WorkflowCode.replace(/-v\d+$/, '');
      const versionSuffix = newVersion.replace(/\./g, '');

      const newDefinition: Partial<IWorkflowDefinition> = {
        Title: `${original.Title} (v${newVersion})`,
        WorkflowCode: `${baseCode}-v${versionSuffix}`,
        Description: original.Description,
        Version: newVersion,
        ProcessType: original.ProcessType,
        IsActive: false,
        IsDefault: false,
        TriggerConditions: original.TriggerConditions,
        Steps: original.Steps,
        Variables: original.Variables,
        Category: original.Category,
        Tags: original.Tags,
        EstimatedDuration: original.EstimatedDuration
      };

      return await this.create(newDefinition);
    } catch (error) {
      logger.error('WorkflowDefinitionService', `Error creating version for workflow ${id}`, error);
      throw new Error('Unable to create new workflow version.');
    }
  }

  // ============================================================================
  // DELETE OPERATIONS
  // ============================================================================

  /**
   * Delete workflow definition
   */
  public async delete(id: number): Promise<void> {
    try {
      // Check if workflow has any instances
      const instanceCount = await this.getInstanceCount(id);
      if (instanceCount > 0) {
        throw new Error(`Cannot delete: workflow has ${instanceCount} existing instances. Unpublish instead.`);
      }

      await this.sp.web.lists.getByTitle(LIST_NAME).items
        .getById(id)
        .delete();

      logger.info('WorkflowDefinitionService', `Deleted workflow definition: ${id}`);
    } catch (error) {
      logger.error('WorkflowDefinitionService', `Error deleting workflow definition ${id}`, error);
      throw error;
    }
  }

  // ============================================================================
  // VALIDATION
  // ============================================================================

  /**
   * Validate workflow definition comprehensively
   * Checks: required fields, step structure, reachability, loops, step-specific config
   */
  public validateDefinition(definition: Partial<IWorkflowDefinition>): IValidationResult {
    const errors: IValidationError[] = [];
    const warnings: IValidationError[] = [];

    // Required fields
    if (!definition.Title?.trim()) {
      errors.push({ field: 'Title', code: 'REQUIRED', message: 'Title is required' });
    }
    if (!definition.WorkflowCode?.trim()) {
      errors.push({ field: 'WorkflowCode', code: 'REQUIRED', message: 'Workflow code is required' });
    }
    if (!definition.ProcessType) {
      errors.push({ field: 'ProcessType', code: 'REQUIRED', message: 'Process type is required' });
    }

    // Parse and validate steps
    let steps: IWorkflowStep[] = [];
    try {
      steps = typeof definition.Steps === 'string'
        ? JSON.parse(definition.Steps || '[]')
        : definition.Steps || [];
    } catch {
      errors.push({ field: 'Steps', code: 'INVALID_JSON', message: 'Steps contains invalid JSON' });
      return { valid: false, errors, warnings };
    }

    if (steps.length === 0) {
      errors.push({ field: 'Steps', code: 'EMPTY', message: 'At least one step is required' });
      return { valid: false, errors, warnings };
    }

    // Build step map and validate basic structure
    const stepMap = new Map<string, IWorkflowStep>();
    const stepIds = new Set<string>();
    let startStep: IWorkflowStep | undefined;
    let endStep: IWorkflowStep | undefined;

    steps.forEach((step, index) => {
      // Check for duplicate IDs
      if (stepIds.has(step.id)) {
        errors.push({
          stepId: step.id,
          code: 'DUPLICATE_ID',
          message: `Duplicate step ID: ${step.id}`
        });
      }
      stepIds.add(step.id);
      stepMap.set(step.id, step);

      // Check required step fields
      if (!step.name?.trim()) {
        errors.push({
          stepId: step.id,
          code: 'MISSING_NAME',
          message: `Step ${index + 1} is missing a name`
        });
      }
      if (!step.type) {
        errors.push({
          stepId: step.id,
          code: 'MISSING_TYPE',
          message: `Step ${step.name || index + 1} is missing a type`
        });
      }

      // Track start/end steps
      if (step.type === StepType.Start) startStep = step;
      if (step.type === StepType.End) endStep = step;

      // Validate step-specific configuration
      this.validateStepConfig(step, errors, warnings);
    });

    if (!startStep) {
      errors.push({ code: 'NO_START', message: 'Workflow must have a Start step' });
    }
    if (!endStep) {
      errors.push({ code: 'NO_END', message: 'Workflow must have an End step' });
    }

    // Validate all transitions point to existing steps
    this.validateTransitions(steps, stepIds, errors);

    // Only run advanced validation if basic structure is valid
    if (errors.length === 0 && startStep && endStep) {
      // Check reachability - all steps must be reachable from Start
      const unreachableSteps = this.findUnreachableSteps(steps, startStep, stepMap);
      unreachableSteps.forEach(stepId => {
        const step = stepMap.get(stepId);
        warnings.push({
          stepId,
          code: 'UNREACHABLE',
          message: `Step "${step?.name || stepId}" is not reachable from Start step`
        });
      });

      // Check for infinite loops (cycles without exit conditions)
      const loopSteps = this.detectInfiniteLoops(steps, stepMap);
      loopSteps.forEach(stepId => {
        const step = stepMap.get(stepId);
        warnings.push({
          stepId,
          code: 'POTENTIAL_LOOP',
          message: `Step "${step?.name || stepId}" may be part of an infinite loop`
        });
      });

      // Check that End step is reachable from Start
      const canReachEnd = this.canReachStep(startStep.id, endStep.id, stepMap, steps);
      if (!canReachEnd) {
        errors.push({
          code: 'END_UNREACHABLE',
          message: 'End step is not reachable from Start step - workflow will never complete'
        });
      }
    }

    return {
      valid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Validate step-specific configuration
   */
  private validateStepConfig(
    step: IWorkflowStep,
    errors: IValidationError[],
    warnings: IValidationError[]
  ): void {
    const config = step.config || {};

    switch (step.type) {
      case StepType.CreateTask:
      case StepType.AssignTasks:
        // Task steps require assignee configuration
        if (!config.assigneeId && !config.assigneeType && !config.assigneeRole && !config.assigneeField) {
          errors.push({
            stepId: step.id,
            code: 'MISSING_ASSIGNEE',
            message: `Task step "${step.name}" must have an assignee configured (assigneeId, assigneeType, assigneeRole, or assigneeField)`
          });
        }
        // Warn if no due date configured
        if (!config.dueDaysFromNow && !config.dueDaysField) {
          warnings.push({
            stepId: step.id,
            code: 'NO_DUE_DATE',
            message: `Task step "${step.name}" has no due date configured`
          });
        }
        break;

      case StepType.Approval:
        // Approval steps require approver configuration
        if (!config.approverId && !config.approverRole && !config.approverEmail && !config.approverField) {
          errors.push({
            stepId: step.id,
            code: 'MISSING_APPROVER',
            message: `Approval step "${step.name}" must have approvers configured`
          });
        }
        break;

      case StepType.Notification:
        // Notification steps require recipients and template/message
        if (!config.recipientIds && !config.recipientEmails && !config.recipientRole && !config.recipientField) {
          errors.push({
            stepId: step.id,
            code: 'MISSING_RECIPIENTS',
            message: `Notification step "${step.name}" must have recipients configured`
          });
        }
        if (!config.messageTemplate && !config.notificationSubject) {
          warnings.push({
            stepId: step.id,
            code: 'MISSING_TEMPLATE',
            message: `Notification step "${step.name}" has no template or subject configured`
          });
        }
        break;

      case StepType.Condition:
        // Condition steps require condition groups
        if (!config.conditionGroups || config.conditionGroups.length === 0) {
          errors.push({
            stepId: step.id,
            code: 'MISSING_CONDITIONS',
            message: `Condition step "${step.name}" has no conditions defined`
          });
        }
        break;

      case StepType.SetVariable:
        // SetVariable steps require variable name and value
        if (!config.variableName) {
          errors.push({
            stepId: step.id,
            code: 'MISSING_VARIABLE_NAME',
            message: `SetVariable step "${step.name}" must specify a variable name`
          });
        }
        break;

      case StepType.Wait:
        // Wait steps should have duration configured
        if (!config.waitHours && !config.waitUntilField) {
          warnings.push({
            stepId: step.id,
            code: 'NO_WAIT_DURATION',
            message: `Wait step "${step.name}" has no duration configured`
          });
        }
        break;

      case StepType.WaitForTasks:
        // WaitForTasks should reference task step IDs
        if (!config.waitForTaskIds || config.waitForTaskIds.length === 0) {
          errors.push({
            stepId: step.id,
            code: 'MISSING_TASK_IDS',
            message: `WaitForTasks step "${step.name}" must reference task step IDs`
          });
        }
        break;

      case StepType.Parallel:
        // Parallel steps need parallel step IDs
        if ((!config.parallelStepIds || config.parallelStepIds.length === 0) &&
            (!step.onComplete?.parallelStepIds || step.onComplete.parallelStepIds.length === 0)) {
          errors.push({
            stepId: step.id,
            code: 'MISSING_PARALLEL_STEPS',
            message: `Parallel step "${step.name}" must define parallel step IDs`
          });
        }
        break;

      case StepType.Webhook:
        // Webhook requires URL
        if (!config.webhookUrl) {
          errors.push({
            stepId: step.id,
            code: 'MISSING_WEBHOOK_URL',
            message: `Webhook step "${step.name}" must have a URL configured`
          });
        }
        break;

      case StepType.CallWorkflow:
        // CallWorkflow requires workflow ID or code
        if (!config.subWorkflowId && !config.subWorkflowCode) {
          errors.push({
            stepId: step.id,
            code: 'MISSING_WORKFLOW_REF',
            message: `CallWorkflow step "${step.name}" must reference a workflow (subWorkflowId or subWorkflowCode)`
          });
        }
        break;
    }
  }

  /**
   * Validate all transitions point to existing steps
   */
  private validateTransitions(
    steps: IWorkflowStep[],
    stepIds: Set<string>,
    errors: IValidationError[]
  ): void {
    steps.forEach(step => {
      if (step.onComplete) {
        // Validate Goto transitions
        if (step.onComplete.type === TransitionType.Goto && step.onComplete.targetStepId) {
          if (!stepIds.has(step.onComplete.targetStepId)) {
            errors.push({
              stepId: step.id,
              code: 'INVALID_TRANSITION',
              message: `Step "${step.name}" references non-existent step: ${step.onComplete.targetStepId}`
            });
          }
        }

        // Validate Branch transitions
        if (step.onComplete.type === TransitionType.Branch && step.onComplete.branches) {
          step.onComplete.branches.forEach((branch, idx) => {
            if (branch.targetStepId && !stepIds.has(branch.targetStepId)) {
              errors.push({
                stepId: step.id,
                code: 'INVALID_BRANCH_TARGET',
                message: `Step "${step.name}" branch ${idx + 1} references non-existent step: ${branch.targetStepId}`
              });
            }
          });
        }

        // Validate Parallel transitions
        if (step.onComplete.type === TransitionType.Parallel && step.onComplete.parallelStepIds) {
          step.onComplete.parallelStepIds.forEach(parallelId => {
            if (!stepIds.has(parallelId)) {
              errors.push({
                stepId: step.id,
                code: 'INVALID_PARALLEL_TARGET',
                message: `Step "${step.name}" references non-existent parallel step: ${parallelId}`
              });
            }
          });
        }
      }

      // Validate WaitForTasks references
      if (step.type === StepType.WaitForTasks && step.config?.waitForTaskIds) {
        step.config.waitForTaskIds.forEach((taskStepId: string) => {
          if (!stepIds.has(taskStepId)) {
            errors.push({
              stepId: step.id,
              code: 'INVALID_TASK_REFERENCE',
              message: `WaitForTasks step "${step.name}" references non-existent step: ${taskStepId}`
            });
          }
        });
      }

      // Validate Parallel step references
      if (step.type === StepType.Parallel && step.config?.parallelStepIds) {
        step.config.parallelStepIds.forEach((parallelId: string) => {
          if (!stepIds.has(parallelId)) {
            errors.push({
              stepId: step.id,
              code: 'INVALID_PARALLEL_STEP',
              message: `Parallel step "${step.name}" references non-existent step: ${parallelId}`
            });
          }
        });
      }
    });
  }

  /**
   * Find steps that are not reachable from the Start step
   */
  private findUnreachableSteps(
    steps: IWorkflowStep[],
    startStep: IWorkflowStep,
    stepMap: Map<string, IWorkflowStep>
  ): string[] {
    const reachable = new Set<string>();
    const queue: string[] = [startStep.id];

    while (queue.length > 0) {
      const currentId = queue.shift()!;
      if (reachable.has(currentId)) continue;

      reachable.add(currentId);
      const current = stepMap.get(currentId);
      if (!current) continue;

      // Get all possible next steps
      const nextSteps = this.getNextStepIds(current, steps);
      nextSteps.forEach(nextId => {
        if (!reachable.has(nextId)) {
          queue.push(nextId);
        }
      });
    }

    // Find steps not in reachable set
    return steps
      .filter(s => !reachable.has(s.id))
      .map(s => s.id);
  }

  /**
   * Detect potential infinite loops (cycles without proper exit)
   */
  private detectInfiniteLoops(
    steps: IWorkflowStep[],
    stepMap: Map<string, IWorkflowStep>
  ): string[] {
    const loopSteps: string[] = [];
    const visited = new Set<string>();
    const recursionStack = new Set<string>();

    const detectCycle = (stepId: string): boolean => {
      if (recursionStack.has(stepId)) {
        return true; // Found a cycle
      }
      if (visited.has(stepId)) {
        return false;
      }

      visited.add(stepId);
      recursionStack.add(stepId);

      const step = stepMap.get(stepId);
      if (step && step.type !== StepType.End) {
        const nextSteps = this.getNextStepIds(step, steps);
        for (const nextId of nextSteps) {
          if (detectCycle(nextId)) {
            // Check if this is an intentional retry loop (has conditions)
            if (!step.onComplete?.branches?.some(b => b.conditions && b.conditions.length > 0)) {
              loopSteps.push(stepId);
            }
          }
        }
      }

      recursionStack.delete(stepId);
      return false;
    };

    steps.forEach(step => {
      if (!visited.has(step.id)) {
        detectCycle(step.id);
      }
    });

    return loopSteps;
  }

  /**
   * Check if a target step is reachable from a source step
   */
  private canReachStep(
    sourceId: string,
    targetId: string,
    stepMap: Map<string, IWorkflowStep>,
    steps: IWorkflowStep[]
  ): boolean {
    const visited = new Set<string>();
    const queue: string[] = [sourceId];

    while (queue.length > 0) {
      const currentId = queue.shift()!;
      if (currentId === targetId) return true;
      if (visited.has(currentId)) continue;

      visited.add(currentId);
      const current = stepMap.get(currentId);
      if (!current) continue;

      const nextSteps = this.getNextStepIds(current, steps);
      nextSteps.forEach(nextId => {
        if (!visited.has(nextId)) {
          queue.push(nextId);
        }
      });
    }

    return false;
  }

  /**
   * Get all possible next step IDs from a step
   */
  private getNextStepIds(step: IWorkflowStep, allSteps: IWorkflowStep[]): string[] {
    const nextIds: string[] = [];

    if (!step.onComplete) {
      // Default: next by order
      const currentOrder = step.order || 0;
      const nextByOrder = allSteps
        .filter(s => (s.order || 0) > currentOrder)
        .sort((a, b) => (a.order || 0) - (b.order || 0))[0];
      if (nextByOrder) {
        nextIds.push(nextByOrder.id);
      }
      return nextIds;
    }

    switch (step.onComplete.type) {
      case TransitionType.Next:
        const currentOrder = step.order || 0;
        const nextByOrder = allSteps
          .filter(s => (s.order || 0) > currentOrder)
          .sort((a, b) => (a.order || 0) - (b.order || 0))[0];
        if (nextByOrder) {
          nextIds.push(nextByOrder.id);
        }
        break;

      case TransitionType.Goto:
        if (step.onComplete.targetStepId) {
          nextIds.push(step.onComplete.targetStepId);
        }
        break;

      case TransitionType.Branch:
        if (step.onComplete.branches) {
          step.onComplete.branches.forEach(branch => {
            if (branch.targetStepId) {
              nextIds.push(branch.targetStepId);
            }
          });
        }
        break;

      case TransitionType.Parallel:
        if (step.onComplete.parallelStepIds) {
          nextIds.push(...step.onComplete.parallelStepIds);
        }
        break;

      case TransitionType.End:
        // No next steps
        break;
    }

    // Also check parallel step config
    if (step.config?.parallelStepIds) {
      nextIds.push(...step.config.parallelStepIds);
    }

    return nextIds;
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Parse workflow definition JSON fields
   */
  public parseDefinition(definition: IWorkflowDefinition): IParsedWorkflowDefinition {
    let steps: IWorkflowStep[] = [];
    let variables: IWorkflowVariable[] = [];
    let triggerConditions: ITriggerCondition[] = [];

    try {
      steps = definition.Steps ? JSON.parse(definition.Steps) : [];
    } catch (error) {
      logger.warn('WorkflowDefinitionService', `Failed to parse steps for workflow ${definition.Id}`, error);
    }

    try {
      variables = definition.Variables ? JSON.parse(definition.Variables) : [];
    } catch (error) {
      logger.warn('WorkflowDefinitionService', `Failed to parse variables for workflow ${definition.Id}`, error);
    }

    try {
      triggerConditions = definition.TriggerConditions ? JSON.parse(definition.TriggerConditions) : [];
    } catch (error) {
      logger.warn('WorkflowDefinitionService', `Failed to parse trigger conditions for workflow ${definition.Id}`, error);
    }

    const { Steps: _s, Variables: _v, TriggerConditions: _t, ...rest } = definition;

    return {
      ...rest,
      steps,
      variables,
      triggerConditions
    };
  }

  /**
   * Count steps in a workflow
   */
  private countSteps(stepsJson: string | undefined): number {
    if (!stepsJson) return 0;
    try {
      const steps = JSON.parse(stepsJson);
      return Array.isArray(steps) ? steps.length : 0;
    } catch {
      return 0;
    }
  }

  /**
   * Get count of workflow instances for this definition
   */
  private async getInstanceCount(definitionId: number): Promise<number> {
    try {
      const items = await this.sp.web.lists.getByTitle('PM_WorkflowInstances').items
        .filter(`WorkflowDefinitionId eq ${definitionId}`)
        .select('Id')
        .top(1)();

      return items.length;
    } catch {
      // List may not exist yet
      return 0;
    }
  }

  /**
   * Increment usage counter
   */
  public async incrementUsageCount(id: number): Promise<void> {
    try {
      const definition = await this.getById(id);
      await this.sp.web.lists.getByTitle(LIST_NAME).items
        .getById(id)
        .update({
          TimesUsed: (definition.TimesUsed || 0) + 1
        });
    } catch (error) {
      logger.warn('WorkflowDefinitionService', `Failed to increment usage count for workflow ${id}`, error);
      // Don't throw - this is non-critical
    }
  }

  /**
   * Update success rate based on completed instances
   */
  public async updateSuccessRate(id: number, successRate: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(LIST_NAME).items
        .getById(id)
        .update({
          SuccessRate: successRate
        });
    } catch (error) {
      logger.warn('WorkflowDefinitionService', `Failed to update success rate for workflow ${id}`, error);
    }
  }

  /**
   * Update average completion time
   */
  public async updateAverageCompletionTime(id: number, avgTimeHours: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(LIST_NAME).items
        .getById(id)
        .update({
          AverageCompletionTime: avgTimeHours
        });
    } catch (error) {
      logger.warn('WorkflowDefinitionService', `Failed to update avg completion time for workflow ${id}`, error);
    }
  }
}
