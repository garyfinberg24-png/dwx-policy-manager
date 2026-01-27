// @ts-nocheck
/**
 * RetryHandler
 *
 * Handles retry logic for failed workflow steps with configurable strategies.
 * Supports exponential backoff, max retries, and error-specific handling.
 *
 * @author JML Development Team
 * @version 1.0.0
 */

import {
  IWorkflowStep,
  IActionContext,
  IActionResult,
  IStepErrorConfig,
  StepStatus
} from '../../../models/IWorkflow';
import { logger } from '../../LoggingService';

/**
 * Retry execution context
 */
export interface IRetryContext {
  attemptNumber: number;
  maxAttempts: number;
  lastError: string;
  totalDelayMs: number;
  nextRetryAt?: Date;
}

/**
 * Retry result
 */
export interface IRetryResult extends IActionResult {
  retryContext?: IRetryContext;
  shouldRetry: boolean;
  retryDelayMs?: number;
}

/**
 * Step executor function type
 */
export type StepExecutor = (
  step: IWorkflowStep,
  context: IActionContext
) => Promise<IActionResult>;

/**
 * Handler for retry logic in workflow steps
 */
export class RetryHandler {
  private static readonly DEFAULT_MAX_RETRIES = 3;
  private static readonly DEFAULT_RETRY_DELAY_MS = 60000; // 1 minute
  private static readonly DEFAULT_BACKOFF_MULTIPLIER = 2;
  private static readonly MAX_DELAY_MS = 3600000; // 1 hour max

  /**
   * Execute a step with retry logic
   */
  public async executeWithRetry(
    step: IWorkflowStep,
    context: IActionContext,
    executor: StepExecutor,
    existingRetryContext?: IRetryContext
  ): Promise<IRetryResult> {
    const errorConfig = step.errorConfig || this.getDefaultErrorConfig();
    const maxRetries = errorConfig.retryCount ?? RetryHandler.DEFAULT_MAX_RETRIES;

    const retryContext: IRetryContext = existingRetryContext || {
      attemptNumber: 0,
      maxAttempts: maxRetries,
      lastError: '',
      totalDelayMs: 0
    };

    retryContext.attemptNumber++;

    logger.info(
      'RetryHandler',
      `Executing step "${step.name}" (attempt ${retryContext.attemptNumber}/${maxRetries + 1})`
    );

    try {
      // Execute the step
      const result = await executor(step, context);

      if (result.success) {
        // Success - no retry needed
        return {
          ...result,
          retryContext,
          shouldRetry: false
        };
      }

      // Step failed - determine if we should retry
      retryContext.lastError = result.error || 'Unknown error';

      return await this.handleFailure(step, context, errorConfig, retryContext, executor);

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      retryContext.lastError = errorMessage;

      logger.error('RetryHandler', `Step "${step.name}" threw exception`, error);

      return await this.handleFailure(step, context, errorConfig, retryContext, executor);
    }
  }

  /**
   * Handle step failure and determine next action
   */
  private async handleFailure(
    step: IWorkflowStep,
    context: IActionContext,
    errorConfig: IStepErrorConfig,
    retryContext: IRetryContext,
    _executor: StepExecutor
  ): Promise<IRetryResult> {
    const action = errorConfig.action || 'fail';

    switch (action) {
      case 'retry':
        return this.handleRetryAction(step, errorConfig, retryContext);

      case 'skip':
        return this.handleSkipAction(step, retryContext);

      case 'goto':
        return this.handleGotoAction(step, errorConfig, retryContext);

      case 'fail':
      default:
        return this.handleFailAction(step, errorConfig, retryContext, context);
    }
  }

  /**
   * Handle retry action
   */
  private handleRetryAction(
    step: IWorkflowStep,
    errorConfig: IStepErrorConfig,
    retryContext: IRetryContext
  ): IRetryResult {
    const maxRetries = errorConfig.retryCount ?? RetryHandler.DEFAULT_MAX_RETRIES;

    if (retryContext.attemptNumber > maxRetries) {
      // Max retries exceeded
      logger.warn(
        'RetryHandler',
        `Max retries (${maxRetries}) exceeded for step "${step.name}"`
      );

      return {
        success: false,
        error: `Max retries exceeded: ${retryContext.lastError}`,
        nextAction: 'fail',
        retryContext,
        shouldRetry: false
      };
    }

    // Calculate delay with exponential backoff
    const baseDelay = (errorConfig.retryDelayMinutes ?? 1) * 60000;
    const multiplier = errorConfig.retryBackoffMultiplier ?? RetryHandler.DEFAULT_BACKOFF_MULTIPLIER;
    const delayMs = Math.min(
      baseDelay * Math.pow(multiplier, retryContext.attemptNumber - 1),
      RetryHandler.MAX_DELAY_MS
    );

    retryContext.totalDelayMs += delayMs;
    retryContext.nextRetryAt = new Date(Date.now() + delayMs);

    logger.info(
      'RetryHandler',
      `Scheduling retry for step "${step.name}" in ${delayMs / 1000}s (attempt ${retryContext.attemptNumber + 1})`
    );

    return {
      success: false,
      error: retryContext.lastError,
      nextAction: 'wait',
      retryContext,
      shouldRetry: true,
      retryDelayMs: delayMs,
      outputVariables: {
        retryScheduled: true,
        retryAttempt: retryContext.attemptNumber,
        nextRetryAt: retryContext.nextRetryAt.toISOString()
      }
    };
  }

  /**
   * Handle skip action
   */
  private handleSkipAction(
    step: IWorkflowStep,
    retryContext: IRetryContext
  ): IRetryResult {
    logger.info(
      'RetryHandler',
      `Skipping failed step "${step.name}" as per error configuration`
    );

    return {
      success: true,
      nextAction: 'continue',
      retryContext,
      shouldRetry: false,
      outputVariables: {
        stepSkipped: true,
        skipReason: retryContext.lastError,
        stepStatus: StepStatus.Skipped
      }
    };
  }

  /**
   * Handle goto action
   */
  private handleGotoAction(
    step: IWorkflowStep,
    errorConfig: IStepErrorConfig,
    retryContext: IRetryContext
  ): IRetryResult {
    if (!errorConfig.gotoStepId) {
      logger.warn(
        'RetryHandler',
        `Goto action configured but no gotoStepId specified for step "${step.name}"`
      );

      return {
        success: false,
        error: 'Goto action requires gotoStepId',
        nextAction: 'fail',
        retryContext,
        shouldRetry: false
      };
    }

    logger.info(
      'RetryHandler',
      `Redirecting from failed step "${step.name}" to step "${errorConfig.gotoStepId}"`
    );

    return {
      success: true,
      nextAction: 'continue',
      retryContext,
      shouldRetry: false,
      outputVariables: {
        errorRedirect: true,
        originalError: retryContext.lastError,
        redirectToStepId: errorConfig.gotoStepId
      }
    };
  }

  /**
   * Handle fail action (default)
   */
  private async handleFailAction(
    step: IWorkflowStep,
    errorConfig: IStepErrorConfig,
    retryContext: IRetryContext,
    context: IActionContext
  ): Promise<IRetryResult> {
    logger.error(
      'RetryHandler',
      `Step "${step.name}" failed permanently: ${retryContext.lastError}`
    );

    // Send error notifications if configured
    if (errorConfig.notifyOnError && errorConfig.notifyOnError.length > 0) {
      await this.sendErrorNotifications(step, errorConfig, retryContext, context);
    }

    return {
      success: false,
      error: retryContext.lastError,
      nextAction: 'fail',
      retryContext,
      shouldRetry: false,
      outputVariables: {
        stepFailed: true,
        failureReason: retryContext.lastError,
        attemptsMade: retryContext.attemptNumber
      }
    };
  }

  /**
   * Send error notifications
   */
  private async sendErrorNotifications(
    step: IWorkflowStep,
    errorConfig: IStepErrorConfig,
    retryContext: IRetryContext,
    context: IActionContext
  ): Promise<void> {
    if (!errorConfig.notifyOnError) return;

    try {
      const notificationService = context.services?.getService<{
        sendErrorNotification: (
          recipients: string[],
          stepName: string,
          error: string,
          instanceId: number
        ) => Promise<void>;
      }>('WorkflowNotificationService');

      if (notificationService) {
        await notificationService.sendErrorNotification(
          errorConfig.notifyOnError,
          step.name,
          retryContext.lastError,
          context.workflowInstance.Id
        );
      }
    } catch (notifyError) {
      logger.warn('RetryHandler', 'Failed to send error notifications', notifyError);
    }
  }

  /**
   * Get default error configuration
   */
  private getDefaultErrorConfig(): IStepErrorConfig {
    return {
      action: 'fail',
      retryCount: 0
    };
  }

  /**
   * Check if a step should be retried based on stored retry context
   */
  public shouldRetryNow(retryContext: IRetryContext): boolean {
    if (!retryContext.nextRetryAt) {
      return false;
    }

    return new Date() >= retryContext.nextRetryAt;
  }

  /**
   * Get remaining delay until next retry
   */
  public getRemainingDelay(retryContext: IRetryContext): number {
    if (!retryContext.nextRetryAt) {
      return 0;
    }

    const remaining = retryContext.nextRetryAt.getTime() - Date.now();
    return Math.max(0, remaining);
  }

  /**
   * Calculate total retry duration for a step
   */
  public calculateMaxRetryDuration(errorConfig: IStepErrorConfig): number {
    const maxRetries = errorConfig.retryCount ?? RetryHandler.DEFAULT_MAX_RETRIES;
    const baseDelayMs = (errorConfig.retryDelayMinutes ?? 1) * 60000;
    const multiplier = errorConfig.retryBackoffMultiplier ?? RetryHandler.DEFAULT_BACKOFF_MULTIPLIER;

    let totalMs = 0;
    for (let i = 0; i < maxRetries; i++) {
      const delay = Math.min(
        baseDelayMs * Math.pow(multiplier, i),
        RetryHandler.MAX_DELAY_MS
      );
      totalMs += delay;
    }

    return totalMs;
  }
}
