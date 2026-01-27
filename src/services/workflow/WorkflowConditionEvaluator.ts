// @ts-nocheck
/* eslint-disable */
/**
 * WorkflowConditionEvaluator
 * Evaluates conditions for workflow branching and step entry
 * Supports field comparisons, date logic, and complex condition groups
 */

import {
  ICondition,
  IConditionGroup,
  ConditionOperator
} from '../../models/IWorkflow';
import { logger } from '../LoggingService';

/**
 * Context object containing data available for condition evaluation
 */
export interface IEvaluationContext {
  [key: string]: unknown;
}

export class WorkflowConditionEvaluator {

  // ============================================================================
  // PUBLIC EVALUATION METHODS
  // ============================================================================

  /**
   * Evaluate a list of conditions (all must be true - AND logic)
   */
  public evaluateConditions(conditions: ICondition[], context: IEvaluationContext): boolean {
    if (!conditions || conditions.length === 0) {
      return true;
    }

    return conditions.every(condition => this.evaluateCondition(condition, context));
  }

  /**
   * Evaluate condition groups with AND/OR logic
   */
  public evaluateConditionGroups(groups: IConditionGroup[], context: IEvaluationContext): boolean {
    if (!groups || groups.length === 0) {
      return true;
    }

    // Groups are connected with AND logic between them
    return groups.every(group => this.evaluateConditionGroup(group, context));
  }

  /**
   * Evaluate a single condition group
   */
  public evaluateConditionGroup(group: IConditionGroup, context: IEvaluationContext): boolean {
    if (!group.conditions || group.conditions.length === 0) {
      return true;
    }

    if (group.logic === 'OR') {
      return group.conditions.some(condition => this.evaluateCondition(condition, context));
    } else {
      // Default to AND
      return group.conditions.every(condition => this.evaluateCondition(condition, context));
    }
  }

  /**
   * Evaluate a single condition
   */
  public evaluateCondition(condition: ICondition, context: IEvaluationContext): boolean {
    try {
      // Get field value from context
      const fieldValue = this.getFieldValue(condition.field, context);

      // Get comparison value (either direct value or from another field)
      let compareValue: unknown = condition.value;
      if (condition.valueField) {
        compareValue = this.getFieldValue(condition.valueField, context);
      }

      // Evaluate based on operator
      return this.compare(fieldValue, condition.operator, compareValue);
    } catch (error) {
      logger.warn('WorkflowConditionEvaluator', `Error evaluating condition for field ${condition.field}`, error);
      return false;
    }
  }

  // ============================================================================
  // COMPARISON LOGIC
  // ============================================================================

  /**
   * Compare two values using the specified operator
   */
  private compare(
    fieldValue: unknown,
    operator: ConditionOperator,
    compareValue: unknown
  ): boolean {
    switch (operator) {
      case ConditionOperator.Equals:
        return this.equals(fieldValue, compareValue);

      case ConditionOperator.NotEquals:
        return !this.equals(fieldValue, compareValue);

      case ConditionOperator.Contains:
        return this.contains(fieldValue, compareValue);

      case ConditionOperator.StartsWith:
        return this.startsWith(fieldValue, compareValue);

      case ConditionOperator.EndsWith:
        return this.endsWith(fieldValue, compareValue);

      case ConditionOperator.GreaterThan:
        return this.greaterThan(fieldValue, compareValue);

      case ConditionOperator.GreaterThanOrEqual:
        return this.greaterThanOrEqual(fieldValue, compareValue);

      case ConditionOperator.LessThan:
        return this.lessThan(fieldValue, compareValue);

      case ConditionOperator.LessThanOrEqual:
        return this.lessThanOrEqual(fieldValue, compareValue);

      case ConditionOperator.IsEmpty:
        return this.isEmpty(fieldValue);

      case ConditionOperator.IsNotEmpty:
        return !this.isEmpty(fieldValue);

      case ConditionOperator.In:
        return this.inArray(fieldValue, compareValue);

      case ConditionOperator.NotIn:
        return !this.inArray(fieldValue, compareValue);

      case ConditionOperator.DateBefore:
        return this.dateBefore(fieldValue, compareValue);

      case ConditionOperator.DateAfter:
        return this.dateAfter(fieldValue, compareValue);

      case ConditionOperator.DateEquals:
        return this.dateEquals(fieldValue, compareValue);

      default:
        logger.warn('WorkflowConditionEvaluator', `Unknown operator: ${operator}`);
        return false;
    }
  }

  /**
   * Equality comparison (case-insensitive for strings)
   */
  private equals(fieldValue: unknown, compareValue: unknown): boolean {
    if (fieldValue === compareValue) return true;

    // Case-insensitive string comparison
    if (typeof fieldValue === 'string' && typeof compareValue === 'string') {
      return fieldValue.toLowerCase() === compareValue.toLowerCase();
    }

    // Number comparison with type coercion
    if (typeof fieldValue === 'number' || typeof compareValue === 'number') {
      return Number(fieldValue) === Number(compareValue);
    }

    // Boolean comparison with string conversion
    if (typeof fieldValue === 'boolean' || typeof compareValue === 'boolean') {
      return this.toBoolean(fieldValue) === this.toBoolean(compareValue);
    }

    return false;
  }

  /**
   * Contains check (string contains or array includes)
   */
  private contains(fieldValue: unknown, compareValue: unknown): boolean {
    if (typeof fieldValue === 'string' && typeof compareValue === 'string') {
      return fieldValue.toLowerCase().includes(compareValue.toLowerCase());
    }

    if (Array.isArray(fieldValue)) {
      return fieldValue.some(item =>
        this.equals(item, compareValue)
      );
    }

    return false;
  }

  /**
   * Starts with check
   */
  private startsWith(fieldValue: unknown, compareValue: unknown): boolean {
    if (typeof fieldValue === 'string' && typeof compareValue === 'string') {
      return fieldValue.toLowerCase().startsWith(compareValue.toLowerCase());
    }
    return false;
  }

  /**
   * Ends with check
   */
  private endsWith(fieldValue: unknown, compareValue: unknown): boolean {
    if (typeof fieldValue === 'string' && typeof compareValue === 'string') {
      return fieldValue.toLowerCase().endsWith(compareValue.toLowerCase());
    }
    return false;
  }

  /**
   * Greater than comparison
   */
  private greaterThan(fieldValue: unknown, compareValue: unknown): boolean {
    const fieldNum = Number(fieldValue);
    const compareNum = Number(compareValue);

    if (isNaN(fieldNum) || isNaN(compareNum)) {
      // Try date comparison
      const fieldDate = this.toDate(fieldValue);
      const compareDate = this.toDate(compareValue);
      if (fieldDate && compareDate) {
        return fieldDate.getTime() > compareDate.getTime();
      }
      return false;
    }

    return fieldNum > compareNum;
  }

  /**
   * Greater than or equal comparison
   */
  private greaterThanOrEqual(fieldValue: unknown, compareValue: unknown): boolean {
    return this.greaterThan(fieldValue, compareValue) || this.equals(fieldValue, compareValue);
  }

  /**
   * Less than comparison
   */
  private lessThan(fieldValue: unknown, compareValue: unknown): boolean {
    const fieldNum = Number(fieldValue);
    const compareNum = Number(compareValue);

    if (isNaN(fieldNum) || isNaN(compareNum)) {
      // Try date comparison
      const fieldDate = this.toDate(fieldValue);
      const compareDate = this.toDate(compareValue);
      if (fieldDate && compareDate) {
        return fieldDate.getTime() < compareDate.getTime();
      }
      return false;
    }

    return fieldNum < compareNum;
  }

  /**
   * Less than or equal comparison
   */
  private lessThanOrEqual(fieldValue: unknown, compareValue: unknown): boolean {
    return this.lessThan(fieldValue, compareValue) || this.equals(fieldValue, compareValue);
  }

  /**
   * Empty check
   */
  private isEmpty(fieldValue: unknown): boolean {
    if (fieldValue === null || fieldValue === undefined) return true;
    if (typeof fieldValue === 'string') return fieldValue.trim() === '';
    if (Array.isArray(fieldValue)) return fieldValue.length === 0;
    if (typeof fieldValue === 'object') return Object.keys(fieldValue).length === 0;
    return false;
  }

  /**
   * In array check
   */
  private inArray(fieldValue: unknown, compareValue: unknown): boolean {
    let arrayValue: unknown[] | undefined;

    if (Array.isArray(compareValue)) {
      arrayValue = compareValue;
    } else if (typeof compareValue === 'string') {
      // Try to parse as JSON array
      try {
        const parsed = JSON.parse(compareValue);
        if (Array.isArray(parsed)) {
          arrayValue = parsed;
        }
      } catch {
        // Treat as comma-separated
        arrayValue = compareValue.split(',').map(s => s.trim());
      }
    }

    if (arrayValue) {
      return arrayValue.some(item => this.equals(fieldValue, item));
    }

    return false;
  }

  /**
   * Date before comparison
   */
  private dateBefore(fieldValue: unknown, compareValue: unknown): boolean {
    const fieldDate = this.toDate(fieldValue);
    const compareDate = this.toDate(compareValue);

    if (!fieldDate || !compareDate) return false;

    return fieldDate.getTime() < compareDate.getTime();
  }

  /**
   * Date after comparison
   */
  private dateAfter(fieldValue: unknown, compareValue: unknown): boolean {
    const fieldDate = this.toDate(fieldValue);
    const compareDate = this.toDate(compareValue);

    if (!fieldDate || !compareDate) return false;

    return fieldDate.getTime() > compareDate.getTime();
  }

  /**
   * Date equals comparison (same day)
   */
  private dateEquals(fieldValue: unknown, compareValue: unknown): boolean {
    const fieldDate = this.toDate(fieldValue);
    const compareDate = this.toDate(compareValue);

    if (!fieldDate || !compareDate) return false;

    // Compare date parts only (ignore time)
    return (
      fieldDate.getFullYear() === compareDate.getFullYear() &&
      fieldDate.getMonth() === compareDate.getMonth() &&
      fieldDate.getDate() === compareDate.getDate()
    );
  }

  // ============================================================================
  // VALUE EXTRACTION
  // ============================================================================

  /**
   * Get value from context using dot notation path
   * e.g., "process.Department" or "variables.approvalLevel"
   */
  private getFieldValue(path: string, context: IEvaluationContext): unknown {
    if (!path) return undefined;

    // Handle special tokens
    const processedPath = this.processTokens(path, context);

    // Split path and navigate
    const parts = processedPath.split('.');
    let value: unknown = context;

    for (const part of parts) {
      if (value === null || value === undefined) {
        return undefined;
      }

      // Handle array index notation [0]
      const arrayMatch = part.match(/^(\w+)\[(\d+)\]$/);
      if (arrayMatch) {
        const [, propName, index] = arrayMatch;
        value = (value as Record<string, unknown>)[propName];
        if (Array.isArray(value)) {
          value = value[parseInt(index, 10)];
        } else {
          return undefined;
        }
      } else {
        value = (value as Record<string, unknown>)[part];
      }
    }

    return value;
  }

  /**
   * Process special tokens in field path
   */
  private processTokens(path: string, context: IEvaluationContext): string {
    // Replace @today with current date
    if (path.includes('@today')) {
      path = path.replace('@today', new Date().toISOString().split('T')[0]);
    }

    // Replace @now with current datetime
    if (path.includes('@now')) {
      path = path.replace('@now', new Date().toISOString());
    }

    // Replace @currentUser (if context has it)
    if (path.includes('@currentUser') && context.currentUserId) {
      path = path.replace('@currentUser', String(context.currentUserId));
    }

    return path;
  }

  // ============================================================================
  // TYPE CONVERSION HELPERS
  // ============================================================================

  /**
   * Convert value to boolean
   */
  private toBoolean(value: unknown): boolean {
    if (typeof value === 'boolean') return value;
    if (typeof value === 'string') {
      const lower = value.toLowerCase().trim();
      return lower === 'true' || lower === 'yes' || lower === '1';
    }
    if (typeof value === 'number') return value !== 0;
    return Boolean(value);
  }

  /**
   * Convert value to Date
   */
  private toDate(value: unknown): Date | undefined {
    if (!value) return undefined;

    if (value instanceof Date) return value;

    if (typeof value === 'string') {
      // Handle relative dates
      if (value.startsWith('@')) {
        return this.parseRelativeDate(value);
      }

      const date = new Date(value);
      return isNaN(date.getTime()) ? undefined : date;
    }

    if (typeof value === 'number') {
      return new Date(value);
    }

    return undefined;
  }

  /**
   * Parse relative date expressions
   * @today, @today+7, @today-30, etc.
   */
  private parseRelativeDate(expression: string): Date | undefined {
    const now = new Date();
    now.setHours(0, 0, 0, 0);

    if (expression === '@today') {
      return now;
    }

    const match = expression.match(/^@today([+-])(\d+)$/);
    if (match) {
      const [, operator, days] = match;
      const offset = parseInt(days, 10);
      now.setDate(now.getDate() + (operator === '+' ? offset : -offset));
      return now;
    }

    if (expression === '@now') {
      return new Date();
    }

    return undefined;
  }

  // ============================================================================
  // EXPRESSION HELPERS
  // ============================================================================

  /**
   * Evaluate a simple expression
   * Supports: field references, literals, basic arithmetic
   */
  public evaluateExpression(expression: string, context: IEvaluationContext): unknown {
    if (!expression) return undefined;

    // Check if it's a field reference (starts with field path)
    if (expression.match(/^[a-zA-Z_]/)) {
      const value = this.getFieldValue(expression, context);
      if (value !== undefined) return value;
    }

    // Try to parse as JSON literal
    try {
      return JSON.parse(expression);
    } catch {
      // Not JSON, return as string
      return expression;
    }
  }

  /**
   * Replace field tokens in a template string
   * e.g., "Hello {{employeeName}}, welcome to {{department}}"
   */
  public replaceTokens(template: string, context: IEvaluationContext): string {
    if (!template) return '';

    return template.replace(/\{\{([^}]+)\}\}/g, (match, path) => {
      const value = this.getFieldValue(path.trim(), context);
      return value !== undefined ? String(value) : match;
    });
  }
}
