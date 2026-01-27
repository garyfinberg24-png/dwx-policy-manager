// @ts-nocheck
/**
 * ForEachHandler
 *
 * Handles ForEach/Loop step execution in workflows.
 * Iterates over collections and executes inner steps for each item.
 *
 * @author JML Development Team
 * @version 1.0.0
 */

import {
  IWorkflowStep,
  IActionContext,
  IActionResult,
  IStepConfig,
  StepType
} from '../../../models/IWorkflow';
import { logger } from '../../LoggingService';

/**
 * Context for ForEach loop iteration
 */
export interface IForEachIterationContext {
  item: unknown;
  index: number;
  total: number;
  isFirst: boolean;
  isLast: boolean;
}

/**
 * Result of ForEach execution
 */
export interface IForEachResult extends IActionResult {
  iterations: number;
  successfulIterations: number;
  failedIterations: number;
  iterationResults?: Array<{
    index: number;
    success: boolean;
    error?: string;
    outputVariables?: Record<string, unknown>;
  }>;
}

/**
 * Handler for ForEach step type
 */
export class ForEachHandler {

  /**
   * Execute ForEach step
   */
  public async execute(
    step: IWorkflowStep,
    context: IActionContext,
    stepExecutor: (step: IWorkflowStep, context: IActionContext) => Promise<IActionResult>
  ): Promise<IForEachResult> {
    const config = step.config;

    // Validate configuration
    if (!config.collectionPath) {
      return {
        success: false,
        error: 'ForEach step requires collectionPath configuration',
        iterations: 0,
        successfulIterations: 0,
        failedIterations: 0
      };
    }

    if (!config.innerSteps || config.innerSteps.length === 0) {
      return {
        success: false,
        error: 'ForEach step requires innerSteps configuration',
        iterations: 0,
        successfulIterations: 0,
        failedIterations: 0
      };
    }

    // Get collection from context/variables
    const collection = this.resolveCollectionPath(config.collectionPath, context);

    if (!Array.isArray(collection)) {
      logger.warn('ForEachHandler', `Collection at path '${config.collectionPath}' is not an array or is undefined`);
      return {
        success: true,
        nextAction: 'continue',
        iterations: 0,
        successfulIterations: 0,
        failedIterations: 0,
        outputVariables: {
          forEachCompleted: true,
          iterationCount: 0
        }
      };
    }

    logger.info('ForEachHandler', `Starting ForEach loop with ${collection.length} items`);

    const itemVariable = config.itemVariable || 'item';
    const indexVariable = config.indexVariable || 'index';
    const iterationResults: IForEachResult['iterationResults'] = [];
    let successfulIterations = 0;
    let failedIterations = 0;

    // Execute based on parallel or sequential mode
    if (config.parallelForEach && collection.length > 1) {
      // Parallel execution with max concurrency
      const maxParallel = config.maxParallel || 5;
      const results = await this.executeParallel(
        collection,
        config,
        context,
        stepExecutor,
        itemVariable,
        indexVariable,
        maxParallel
      );

      for (const result of results) {
        iterationResults.push(result);
        if (result.success) {
          successfulIterations++;
        } else {
          failedIterations++;
        }
      }
    } else {
      // Sequential execution
      for (let i = 0; i < collection.length; i++) {
        const item = collection[i];

        const iterationContext: IForEachIterationContext = {
          item,
          index: i,
          total: collection.length,
          isFirst: i === 0,
          isLast: i === collection.length - 1
        };

        try {
          const result = await this.executeIteration(
            item,
            i,
            collection.length,
            config,
            context,
            stepExecutor,
            itemVariable,
            indexVariable
          );

          iterationResults.push({
            index: i,
            success: result.success,
            error: result.error,
            outputVariables: result.outputVariables
          });

          if (result.success) {
            successfulIterations++;
          } else {
            failedIterations++;

            // Check if we should stop on error
            if (config.onError?.action === 'fail') {
              logger.warn('ForEachHandler', `Stopping ForEach loop due to error at index ${i}`);
              break;
            }
          }
        } catch (error) {
          const errorMessage = error instanceof Error ? error.message : 'Unknown error';
          logger.error('ForEachHandler', `Error in iteration ${i}`, error);

          iterationResults.push({
            index: i,
            success: false,
            error: errorMessage
          });
          failedIterations++;

          if (config.onError?.action === 'fail') {
            break;
          }
        }
      }
    }

    logger.info('ForEachHandler', `ForEach completed: ${successfulIterations}/${collection.length} successful`);

    return {
      success: failedIterations === 0 || config.onError?.action !== 'fail',
      nextAction: 'continue',
      iterations: collection.length,
      successfulIterations,
      failedIterations,
      iterationResults,
      outputVariables: {
        forEachCompleted: true,
        iterationCount: collection.length,
        successfulIterations,
        failedIterations
      }
    };
  }

  /**
   * Execute a single iteration
   */
  private async executeIteration(
    item: unknown,
    index: number,
    total: number,
    config: IStepConfig,
    parentContext: IActionContext,
    stepExecutor: (step: IWorkflowStep, context: IActionContext) => Promise<IActionResult>,
    itemVariable: string,
    indexVariable: string
  ): Promise<IActionResult> {
    // Create iteration context with loop variables
    const iterationVariables = {
      ...parentContext.variables,
      [itemVariable]: item,
      [indexVariable]: index,
      [`${itemVariable}_total`]: total,
      [`${itemVariable}_isFirst`]: index === 0,
      [`${itemVariable}_isLast`]: index === total - 1
    };

    const iterationContext: IActionContext = {
      ...parentContext,
      variables: iterationVariables
    };

    // Execute each inner step
    let lastResult: IActionResult = { success: true, nextAction: 'continue' };

    for (const innerStep of config.innerSteps || []) {
      // Skip Start/End steps in inner steps
      if (innerStep.type === StepType.Start || innerStep.type === StepType.End) {
        continue;
      }

      lastResult = await stepExecutor(innerStep, iterationContext);

      if (!lastResult.success) {
        return lastResult;
      }

      // Merge output variables back to iteration context
      if (lastResult.outputVariables) {
        Object.assign(iterationVariables, lastResult.outputVariables);
      }

      // Handle wait actions within loop
      if (lastResult.nextAction === 'wait') {
        // For now, we don't support waiting within loops
        // This would require more complex state management
        logger.warn('ForEachHandler', 'Wait actions within ForEach loops are not fully supported');
      }
    }

    return {
      success: true,
      nextAction: 'continue',
      outputVariables: lastResult.outputVariables
    };
  }

  /**
   * Execute iterations in parallel with max concurrency
   */
  private async executeParallel(
    collection: unknown[],
    config: IStepConfig,
    parentContext: IActionContext,
    stepExecutor: (step: IWorkflowStep, context: IActionContext) => Promise<IActionResult>,
    itemVariable: string,
    indexVariable: string,
    maxParallel: number
  ): Promise<Array<{
    index: number;
    success: boolean;
    error?: string;
    outputVariables?: Record<string, unknown>;
  }>> {
    const results: Array<{
      index: number;
      success: boolean;
      error?: string;
      outputVariables?: Record<string, unknown>;
    }> = [];

    // Process in batches
    for (let i = 0; i < collection.length; i += maxParallel) {
      const batch = collection.slice(i, Math.min(i + maxParallel, collection.length));

      const batchPromises = batch.map(async (item, batchIndex) => {
        const actualIndex = i + batchIndex;

        try {
          const result = await this.executeIteration(
            item,
            actualIndex,
            collection.length,
            config,
            parentContext,
            stepExecutor,
            itemVariable,
            indexVariable
          );

          return {
            index: actualIndex,
            success: result.success,
            error: result.error,
            outputVariables: result.outputVariables
          };
        } catch (error) {
          return {
            index: actualIndex,
            success: false,
            error: error instanceof Error ? error.message : 'Unknown error'
          };
        }
      });

      const batchResults = await Promise.all(batchPromises);
      results.push(...batchResults);
    }

    // Sort by index to maintain order
    return results.sort((a, b) => a.index - b.index);
  }

  /**
   * Resolve collection path from context/variables
   */
  private resolveCollectionPath(path: string, context: IActionContext): unknown {
    const parts = path.split('.');
    let current: unknown = context;

    for (const part of parts) {
      if (current === null || current === undefined) {
        return undefined;
      }

      // Check variables first
      if (part === 'variables' || (current === context && context.variables[part] !== undefined)) {
        current = part === 'variables' ? context.variables : context.variables[part];
      } else if (typeof current === 'object' && current !== null) {
        current = (current as Record<string, unknown>)[part];
      } else {
        return undefined;
      }
    }

    return current;
  }
}
