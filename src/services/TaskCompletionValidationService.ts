// @ts-nocheck
/**
 * TaskCompletionValidationService
 * Pre-completion validation to ensure data quality and business rules
 * CRITICAL: Prevents invalid task completions that could cause workflow issues
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import { IJmlTaskAssignment } from '../models';
import { TaskStatus, Priority } from '../models/ICommon';
import { logger } from './LoggingService';

/**
 * Validation rule configuration
 */
export interface IValidationRule {
  id: string;
  name: string;
  description: string;
  severity: 'error' | 'warning' | 'info';
  enabled: boolean;
  category: 'required' | 'dependency' | 'business' | 'data';
  validate: (task: IJmlTaskAssignment, context: IValidationContext) => Promise<IValidationResult>;
}

/**
 * Context for validation
 */
export interface IValidationContext {
  sp: SPFI;
  userId: number;
  skipWarnings?: boolean;
  forceComplete?: boolean;
  completionData?: {
    notes?: string;
    hours?: number;
    attachments?: number;
  };
}

/**
 * Individual validation result
 */
export interface IValidationResult {
  ruleId: string;
  passed: boolean;
  severity: 'error' | 'warning' | 'info';
  message: string;
  details?: Record<string, unknown>;
  canOverride?: boolean;
}

/**
 * Complete validation response
 */
export interface ITaskValidationResponse {
  taskId: number;
  valid: boolean;
  canComplete: boolean;
  errors: IValidationResult[];
  warnings: IValidationResult[];
  info: IValidationResult[];
  checkedAt: Date;
  overridableErrors: string[];
}

/**
 * Validation rule definitions
 */
const DEFAULT_VALIDATION_RULES: Omit<IValidationRule, 'validate'>[] = [
  {
    id: 'task-status-valid',
    name: 'Valid Task Status',
    description: 'Task must be in a completable status',
    severity: 'error',
    enabled: true,
    category: 'required'
  },
  {
    id: 'task-not-blocked',
    name: 'Task Not Blocked',
    description: 'Task must not be blocked',
    severity: 'error',
    enabled: true,
    category: 'required'
  },
  {
    id: 'dependencies-complete',
    name: 'Dependencies Complete',
    description: 'All blocking dependencies must be complete',
    severity: 'error',
    enabled: true,
    category: 'dependency'
  },
  {
    id: 'approval-received',
    name: 'Approval Received',
    description: 'Required approval must be obtained',
    severity: 'error',
    enabled: true,
    category: 'required'
  },
  {
    id: 'required-attachments',
    name: 'Required Attachments',
    description: 'Check if task requires attachments',
    severity: 'warning',
    enabled: true,
    category: 'business'
  },
  {
    id: 'completion-notes',
    name: 'Completion Notes',
    description: 'Completion notes are recommended',
    severity: 'info',
    enabled: true,
    category: 'data'
  },
  {
    id: 'time-tracking',
    name: 'Time Tracking',
    description: 'Actual hours should be recorded',
    severity: 'warning',
    enabled: true,
    category: 'data'
  },
  {
    id: 'overdue-acknowledgment',
    name: 'Overdue Acknowledgment',
    description: 'Task is overdue - acknowledge before completing',
    severity: 'warning',
    enabled: true,
    category: 'business'
  }
];

export class TaskCompletionValidationService {
  private sp: SPFI;
  private rules: Map<string, IValidationRule> = new Map();
  private readonly tasksListTitle = 'JML_TaskAssignments';
  private readonly dependenciesListTitle = 'JML_TaskDependencies';

  constructor(sp: SPFI) {
    this.sp = sp;
    this.initializeRules();
  }

  /**
   * Initialize validation rules
   */
  private initializeRules(): void {
    // Task Status Valid
    this.registerRule({
      ...DEFAULT_VALIDATION_RULES.find(r => r.id === 'task-status-valid')!,
      validate: async (task: IJmlTaskAssignment): Promise<IValidationResult> => {
        const completableStatuses = [TaskStatus.NotStarted, TaskStatus.InProgress, TaskStatus.Waiting];
        const isValid = completableStatuses.includes(task.Status);

        return {
          ruleId: 'task-status-valid',
          passed: isValid,
          severity: 'error',
          message: isValid
            ? 'Task status allows completion'
            : `Task cannot be completed from status: ${task.Status}`,
          details: { currentStatus: task.Status }
        };
      }
    });

    // Task Not Blocked
    this.registerRule({
      ...DEFAULT_VALIDATION_RULES.find(r => r.id === 'task-not-blocked')!,
      validate: async (task: IJmlTaskAssignment): Promise<IValidationResult> => {
        return {
          ruleId: 'task-not-blocked',
          passed: !task.IsBlocked,
          severity: 'error',
          message: task.IsBlocked
            ? `Task is blocked: ${task.BlockedReason || 'No reason provided'}`
            : 'Task is not blocked',
          details: { isBlocked: task.IsBlocked, reason: task.BlockedReason }
        };
      }
    });

    // Dependencies Complete
    this.registerRule({
      ...DEFAULT_VALIDATION_RULES.find(r => r.id === 'dependencies-complete')!,
      validate: async (task: IJmlTaskAssignment, context: IValidationContext): Promise<IValidationResult> => {
        try {
          // Check for blocking dependencies
          const dependencies = await context.sp.web.lists
            .getByTitle(this.dependenciesListTitle)
            .items
            .filter(`DependentTaskId eq ${task.Id} and DependencyType eq 'BlockedBy'`)
            .select('Id', 'BlockingTaskId', 'DependencyType')();

          if (dependencies.length === 0) {
            return {
              ruleId: 'dependencies-complete',
              passed: true,
              severity: 'error',
              message: 'No blocking dependencies'
            };
          }

          // Check if all blocking tasks are complete
          const blockingTaskIds = dependencies.map(d => d.BlockingTaskId);
          const blockingTasks = await context.sp.web.lists
            .getByTitle(this.tasksListTitle)
            .items
            .filter(blockingTaskIds.map(id => `Id eq ${id}`).join(' or '))
            .select('Id', 'Title', 'Status')();

          const incompleteTasks = blockingTasks.filter(t =>
            t.Status !== TaskStatus.Completed && t.Status !== TaskStatus.Skipped
          );

          return {
            ruleId: 'dependencies-complete',
            passed: incompleteTasks.length === 0,
            severity: 'error',
            message: incompleteTasks.length === 0
              ? 'All dependencies are complete'
              : `${incompleteTasks.length} blocking task(s) not complete`,
            details: {
              totalDependencies: dependencies.length,
              incompleteTasks: incompleteTasks.map(t => ({ id: t.Id, title: t.Title, status: t.Status }))
            }
          };
        } catch (error) {
          logger.warn('TaskCompletionValidationService', 'Error checking dependencies', error);
          return {
            ruleId: 'dependencies-complete',
            passed: true,
            severity: 'error',
            message: 'Could not verify dependencies - proceeding',
            details: { error: 'Dependency check failed' }
          };
        }
      }
    });

    // Approval Received
    this.registerRule({
      ...DEFAULT_VALIDATION_RULES.find(r => r.id === 'approval-received')!,
      validate: async (task: IJmlTaskAssignment): Promise<IValidationResult> => {
        if (!task.RequiresApproval) {
          return {
            ruleId: 'approval-received',
            passed: true,
            severity: 'error',
            message: 'Task does not require approval'
          };
        }

        const isApproved = task.ApprovalStatus === 'Approved';
        return {
          ruleId: 'approval-received',
          passed: isApproved,
          severity: 'error',
          message: isApproved
            ? 'Task approval received'
            : `Task requires approval (current status: ${task.ApprovalStatus || 'Pending'})`,
          details: { approvalStatus: task.ApprovalStatus }
        };
      }
    });

    // Required Attachments
    this.registerRule({
      ...DEFAULT_VALIDATION_RULES.find(r => r.id === 'required-attachments')!,
      validate: async (task: IJmlTaskAssignment, context: IValidationContext): Promise<IValidationResult> => {
        // Check task template for attachment requirements
        const attachmentCount = context.completionData?.attachments || 0;

        // For now, just warn if no attachments (could be enhanced with task-specific requirements)
        return {
          ruleId: 'required-attachments',
          passed: true, // Pass but warn
          severity: 'warning',
          message: attachmentCount > 0
            ? `${attachmentCount} attachment(s) present`
            : 'No attachments uploaded - consider if documents are required',
          canOverride: true,
          details: { attachmentCount }
        };
      }
    });

    // Completion Notes
    this.registerRule({
      ...DEFAULT_VALIDATION_RULES.find(r => r.id === 'completion-notes')!,
      validate: async (task: IJmlTaskAssignment, context: IValidationContext): Promise<IValidationResult> => {
        const hasNotes = !!(context.completionData?.notes || task.CompletionNotes);

        return {
          ruleId: 'completion-notes',
          passed: true, // Info only
          severity: 'info',
          message: hasNotes
            ? 'Completion notes provided'
            : 'Consider adding completion notes for future reference',
          canOverride: true
        };
      }
    });

    // Time Tracking
    this.registerRule({
      ...DEFAULT_VALIDATION_RULES.find(r => r.id === 'time-tracking')!,
      validate: async (task: IJmlTaskAssignment, context: IValidationContext): Promise<IValidationResult> => {
        const hasHours = (context.completionData?.hours || task.ActualHours) > 0;

        return {
          ruleId: 'time-tracking',
          passed: hasHours,
          severity: 'warning',
          message: hasHours
            ? 'Actual hours recorded'
            : 'Actual hours not recorded - recommended for time tracking',
          canOverride: true,
          details: { actualHours: context.completionData?.hours || task.ActualHours }
        };
      }
    });

    // Overdue Acknowledgment
    this.registerRule({
      ...DEFAULT_VALIDATION_RULES.find(r => r.id === 'overdue-acknowledgment')!,
      validate: async (task: IJmlTaskAssignment): Promise<IValidationResult> => {
        const isOverdue = task.DueDate && new Date(task.DueDate) < new Date();

        return {
          ruleId: 'overdue-acknowledgment',
          passed: !isOverdue,
          severity: 'warning',
          message: isOverdue
            ? `Task is overdue (due: ${new Date(task.DueDate!).toLocaleDateString()})`
            : 'Task is not overdue',
          canOverride: true,
          details: { dueDate: task.DueDate, isOverdue }
        };
      }
    });
  }

  /**
   * Register a validation rule
   */
  public registerRule(rule: IValidationRule): void {
    this.rules.set(rule.id, rule);
  }

  /**
   * Enable or disable a rule
   */
  public setRuleEnabled(ruleId: string, enabled: boolean): void {
    const rule = this.rules.get(ruleId);
    if (rule) {
      rule.enabled = enabled;
    }
  }

  /**
   * Validate a task before completion
   */
  public async validateForCompletion(
    taskId: number,
    userId: number,
    completionData?: {
      notes?: string;
      hours?: number;
      attachments?: number;
    }
  ): Promise<ITaskValidationResponse> {
    try {
      // Get the task
      const task = await this.getTask(taskId);
      if (!task) {
        return {
          taskId,
          valid: false,
          canComplete: false,
          errors: [{
            ruleId: 'task-exists',
            passed: false,
            severity: 'error',
            message: 'Task not found'
          }],
          warnings: [],
          info: [],
          checkedAt: new Date(),
          overridableErrors: []
        };
      }

      const context: IValidationContext = {
        sp: this.sp,
        userId,
        completionData
      };

      const errors: IValidationResult[] = [];
      const warnings: IValidationResult[] = [];
      const info: IValidationResult[] = [];
      const overridableErrors: string[] = [];

      // Run all enabled rules
      for (const rule of Array.from(this.rules.values())) {
        if (!rule.enabled) continue;

        try {
          const result = await rule.validate(task, context);

          if (!result.passed) {
            switch (result.severity) {
              case 'error':
                errors.push(result);
                if (result.canOverride) {
                  overridableErrors.push(result.ruleId);
                }
                break;
              case 'warning':
                warnings.push(result);
                break;
              case 'info':
                info.push(result);
                break;
            }
          } else if (result.severity === 'info') {
            info.push(result);
          }
        } catch (ruleError) {
          logger.warn('TaskCompletionValidationService',
            `Rule ${rule.id} failed`, ruleError);
        }
      }

      // Task can complete if no blocking errors
      const blockingErrors = errors.filter(e => !overridableErrors.includes(e.ruleId));

      return {
        taskId,
        valid: errors.length === 0,
        canComplete: blockingErrors.length === 0,
        errors,
        warnings,
        info,
        checkedAt: new Date(),
        overridableErrors
      };
    } catch (error) {
      logger.error('TaskCompletionValidationService',
        `Error validating task ${taskId}`, error);

      return {
        taskId,
        valid: false,
        canComplete: false,
        errors: [{
          ruleId: 'validation-error',
          passed: false,
          severity: 'error',
          message: 'Validation failed: ' + (error instanceof Error ? error.message : 'Unknown error')
        }],
        warnings: [],
        info: [],
        checkedAt: new Date(),
        overridableErrors: []
      };
    }
  }

  /**
   * Quick validation - just check if task can be completed (no details)
   */
  public async canComplete(taskId: number, userId: number): Promise<boolean> {
    const result = await this.validateForCompletion(taskId, userId);
    return result.canComplete;
  }

  /**
   * Validate multiple tasks at once (for bulk operations)
   */
  public async validateBulkCompletion(
    taskIds: number[],
    userId: number
  ): Promise<{
    validTasks: number[];
    invalidTasks: Array<{ taskId: number; reason: string }>;
    warningTasks: Array<{ taskId: number; warnings: string[] }>;
  }> {
    const validTasks: number[] = [];
    const invalidTasks: Array<{ taskId: number; reason: string }> = [];
    const warningTasks: Array<{ taskId: number; warnings: string[] }> = [];

    for (const taskId of taskIds) {
      const result = await this.validateForCompletion(taskId, userId);

      if (result.canComplete) {
        validTasks.push(taskId);
        if (result.warnings.length > 0) {
          warningTasks.push({
            taskId,
            warnings: result.warnings.map(w => w.message)
          });
        }
      } else {
        invalidTasks.push({
          taskId,
          reason: result.errors.map(e => e.message).join('; ')
        });
      }
    }

    return { validTasks, invalidTasks, warningTasks };
  }

  /**
   * Get all validation rules
   */
  public getRules(): IValidationRule[] {
    return Array.from(this.rules.values());
  }

  /**
   * Get task by ID
   */
  private async getTask(taskId: number): Promise<IJmlTaskAssignment | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.getById(taskId)
        .select(
          'Id', 'Title', 'Status', 'Priority', 'DueDate', 'IsBlocked', 'BlockedReason',
          'RequiresApproval', 'ApprovalStatus', 'CompletionNotes', 'ActualHours',
          'AssignedToId', 'WorkflowInstanceId', 'WorkflowStepId'
        )();

      return {
        Id: item.Id,
        Title: item.Title,
        Status: item.Status as TaskStatus,
        Priority: item.Priority as Priority,
        DueDate: item.DueDate ? new Date(item.DueDate) : undefined,
        IsBlocked: item.IsBlocked,
        BlockedReason: item.BlockedReason,
        RequiresApproval: item.RequiresApproval,
        ApprovalStatus: item.ApprovalStatus,
        CompletionNotes: item.CompletionNotes,
        ActualHours: item.ActualHours,
        AssignedToId: item.AssignedToId,
        WorkflowInstanceId: item.WorkflowInstanceId,
        WorkflowStepId: item.WorkflowStepId
      } as IJmlTaskAssignment;
    } catch (error) {
      logger.error('TaskCompletionValidationService',
        `Error getting task ${taskId}`, error);
      return null;
    }
  }
}
