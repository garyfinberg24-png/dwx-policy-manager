// @ts-nocheck
/**
 * ListActionHandler
 * Handles SharePoint list operations within workflow execution
 * Creates, updates, and queries list items
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import {
  IActionContext,
  IActionResult,
  IActionConfig,
  IFieldUpdate
} from '../../../models/IWorkflow';
import { WorkflowConditionEvaluator } from '../WorkflowConditionEvaluator';
import { logger } from '../../LoggingService';

export class ListActionHandler {
  private sp: SPFI;
  private conditionEvaluator: WorkflowConditionEvaluator;

  constructor(sp: SPFI) {
    this.sp = sp;
    this.conditionEvaluator = new WorkflowConditionEvaluator();
  }

  /**
   * Create a new list item
   */
  public async createItem(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const listName = config.listName as string;
      const updates = config.updates as IFieldUpdate[] || [];

      if (!listName) {
        return { success: false, error: 'List name not specified' };
      }

      // Build item data from field updates
      const itemData = this.buildItemData(updates, context);

      // Create item
      const result = await this.sp.web.lists.getByTitle(listName).items.add(itemData);

      logger.info('ListActionHandler', `Created item in ${listName}: ${result.data.Id}`);

      return {
        success: true,
        nextAction: 'continue',
        createdItemIds: [result.data.Id],
        outputVariables: {
          createdItemId: result.data.Id,
          createdItemList: listName
        }
      };
    } catch (error) {
      logger.error('ListActionHandler', 'Error creating list item', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to create list item'
      };
    }
  }

  /**
   * Update an existing list item
   */
  public async updateItem(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const listName = config.listName as string;
      let itemId = config.itemId as number;
      const itemIdField = config.itemIdField as string;
      const updates = config.updates as IFieldUpdate[] || [];

      if (!listName) {
        return { success: false, error: 'List name not specified' };
      }

      // Resolve item ID from field if specified
      if (!itemId && itemIdField) {
        const fieldValue = context.process[itemIdField] || context.variables[itemIdField];
        if (typeof fieldValue === 'number') {
          itemId = fieldValue;
        } else if (typeof fieldValue === 'string') {
          itemId = parseInt(fieldValue, 10);
        }
      }

      if (!itemId) {
        return { success: false, error: 'Item ID not specified or resolved' };
      }

      // Build update data from field updates
      const updateData = this.buildItemData(updates, context);

      // Update item
      await this.sp.web.lists.getByTitle(listName).items
        .getById(itemId)
        .update(updateData);

      logger.info('ListActionHandler', `Updated item ${itemId} in ${listName}`);

      return {
        success: true,
        nextAction: 'continue',
        outputVariables: {
          updatedItemId: itemId,
          updatedItemList: listName
        }
      };
    } catch (error) {
      logger.error('ListActionHandler', 'Error updating list item', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to update list item'
      };
    }
  }

  /**
   * Delete a list item
   */
  public async deleteItem(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const listName = config.listName as string;
      let itemId = config.itemId as number;
      const itemIdField = config.itemIdField as string;

      if (!listName) {
        return { success: false, error: 'List name not specified' };
      }

      // Resolve item ID from field if specified
      if (!itemId && itemIdField) {
        const fieldValue = context.process[itemIdField] || context.variables[itemIdField];
        if (typeof fieldValue === 'number') {
          itemId = fieldValue;
        }
      }

      if (!itemId) {
        return { success: false, error: 'Item ID not specified' };
      }

      // Delete item
      await this.sp.web.lists.getByTitle(listName).items
        .getById(itemId)
        .delete();

      logger.info('ListActionHandler', `Deleted item ${itemId} from ${listName}`);

      return {
        success: true,
        nextAction: 'continue',
        outputVariables: {
          deletedItemId: itemId,
          deletedItemList: listName
        }
      };
    } catch (error) {
      logger.error('ListActionHandler', 'Error deleting list item', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to delete list item'
      };
    }
  }

  /**
   * Query list items
   */
  public async queryItems(config: Record<string, unknown>, context: IActionContext): Promise<IActionResult> {
    try {
      const listName = config.listName as string;
      const filter = config.filter as string;
      const select = config.select as string[] || ['Id', 'Title'];
      const top = config.top as number || 100;

      if (!listName) {
        return { success: false, error: 'List name not specified' };
      }

      // Process filter with context values
      let processedFilter = filter;
      if (filter) {
        const evalContext = { ...context.process, ...context.variables };
        processedFilter = this.conditionEvaluator.replaceTokens(filter, evalContext);
      }

      // Query items
      let query = this.sp.web.lists.getByTitle(listName).items
        .select(...select)
        .top(top);

      if (processedFilter) {
        query = query.filter(processedFilter);
      }

      const items = await query();

      logger.info('ListActionHandler', `Queried ${items.length} items from ${listName}`);

      return {
        success: true,
        nextAction: 'continue',
        outputVariables: {
          queryResults: items,
          queryCount: items.length,
          queriedList: listName
        }
      };
    } catch (error) {
      logger.error('ListActionHandler', 'Error querying list items', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to query list items'
      };
    }
  }

  /**
   * Update process status
   */
  public async updateProcessStatus(
    processId: number,
    status: string,
    additionalUpdates?: Record<string, unknown>
  ): Promise<IActionResult> {
    try {
      const updateData: Record<string, unknown> = {
        ProcessStatus: status,
        ...additionalUpdates
      };

      await this.sp.web.lists.getByTitle('JML_Processes').items
        .getById(processId)
        .update(updateData);

      logger.info('ListActionHandler', `Updated process ${processId} status to ${status}`);

      return {
        success: true,
        nextAction: 'continue',
        outputVariables: {
          processId,
          newStatus: status
        }
      };
    } catch (error) {
      logger.error('ListActionHandler', 'Error updating process status', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to update process status'
      };
    }
  }

  /**
   * Update process progress
   */
  public async updateProcessProgress(
    processId: number,
    completedTasks: number,
    totalTasks: number
  ): Promise<IActionResult> {
    try {
      const progressPercentage = totalTasks > 0 ? Math.round((completedTasks / totalTasks) * 100) : 0;

      await this.sp.web.lists.getByTitle('JML_Processes').items
        .getById(processId)
        .update({
          CompletedTasks: completedTasks,
          TotalTasks: totalTasks,
          ProgressPercentage: progressPercentage
        });

      logger.info('ListActionHandler', `Updated process ${processId} progress to ${progressPercentage}%`);

      return {
        success: true,
        nextAction: 'continue',
        outputVariables: {
          processId,
          completedTasks,
          totalTasks,
          progressPercentage
        }
      };
    } catch (error) {
      logger.error('ListActionHandler', 'Error updating process progress', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to update process progress'
      };
    }
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Build item data from field updates
   */
  private buildItemData(updates: IFieldUpdate[], context: IActionContext): Record<string, unknown> {
    const itemData: Record<string, unknown> = {};
    const evalContext = { ...context.process, ...context.variables };

    for (const update of updates) {
      let value: unknown;

      if (update.expression) {
        // Evaluate expression
        value = this.conditionEvaluator.evaluateExpression(update.expression, evalContext);
      } else if (update.valueField) {
        // Get value from field
        value = evalContext[update.valueField];
      } else {
        // Use direct value
        value = update.value;
      }

      // Handle special field name suffixes for SharePoint
      let fieldName = update.fieldName;
      if (typeof value === 'number' && fieldName.endsWith('Id')) {
        // Person or lookup field - ensure it's treated as ID
        itemData[fieldName] = value;
      } else {
        itemData[fieldName] = value;
      }
    }

    return itemData;
  }
}
