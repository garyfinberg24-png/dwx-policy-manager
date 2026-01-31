/**
 * Validation and sanitization utilities for SharePoint OData queries.
 */
export class ValidationUtils {
  /**
   * Sanitize a string value for use in OData filter expressions.
   * Escapes single quotes to prevent OData injection.
   */
  public static sanitizeForOData(value: string): string {
    if (!value || typeof value !== 'string') return '';
    // OData uses doubled single quotes as escape
    return value.replace(/'/g, "''");
  }

  /**
   * Build an OData filter expression.
   */
  public static buildFilter(field: string, operator: string, value: string | number | boolean): string {
    if (typeof value === 'number' || typeof value === 'boolean') {
      return `${field} ${operator} ${value}`;
    }
    return `${field} ${operator} '${ValidationUtils.sanitizeForOData(String(value))}'`;
  }

  /**
   * Validate that a value is a positive integer.
   */
  public static validateInteger(value: unknown, fieldName: string, minValue: number = 0): number {
    const num = typeof value === 'number' ? value : parseInt(String(value), 10);
    if (isNaN(num) || !isFinite(num) || num < minValue) {
      throw new Error(`${fieldName} must be an integer >= ${minValue}, got: ${value}`);
    }
    return Math.floor(num);
  }
}
