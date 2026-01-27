// @ts-nocheck
import { useState, useCallback } from 'react';
import { ValidationUtils } from '../utils/ValidationUtils';

export interface IFieldValidationOptions {
  required?: boolean;
  minLength?: number;
  maxLength?: number;
  pattern?: RegExp;
  validator?: (value: any) => string | null;
  validateOnChange?: boolean;
  validateOnBlur?: boolean;
}

export interface IFieldValidationState {
  value: any;
  error: string | null;
  touched: boolean;
  dirty: boolean;
}

export interface IFieldValidationResult {
  value: any;
  error: string | null;
  touched: boolean;
  dirty: boolean;
  isValid: boolean;
  onChange: (value: any) => void;
  onBlur: () => void;
  reset: () => void;
  validate: () => boolean;
  setValue: (value: any) => void;
}

/**
 * Custom hook for field-level validation with real-time feedback
 *
 * @param initialValue - Initial value for the field
 * @param options - Validation options
 * @returns Field validation state and handlers
 *
 * @example
 * ```typescript
 * const email = useFieldValidation('', {
 *   required: true,
 *   validator: (value) => {
 *     try {
 *       ValidationUtils.validateEmail(value);
 *       return null;
 *     } catch (error) {
 *       return error.message;
 *     }
 *   },
 *   validateOnChange: true
 * });
 *
 * return (
 *   <Input
 *     value={email.value}
 *     onChange={(e, data) => email.onChange(data.value)}
 *     onBlur={email.onBlur}
 *     validationMessage={email.error}
 *     validationState={email.error ? 'error' : undefined}
 *   />
 * );
 * ```
 */
export const useFieldValidation = (
  initialValue: any = '',
  options: IFieldValidationOptions = {}
): IFieldValidationResult => {
  const {
    required = false,
    minLength,
    maxLength,
    pattern,
    validator,
    validateOnChange = false,
    validateOnBlur = true
  } = options;

  const [state, setState] = useState<IFieldValidationState>({
    value: initialValue,
    error: null,
    touched: false,
    dirty: false
  });

  /**
   * Validate the current value
   */
  const validateValue = useCallback((value: any): string | null => {
    // Required validation
    if (required && (value === null || value === undefined || value === '')) {
      return 'This field is required';
    }

    // Skip further validation if empty and not required
    if (!required && (value === null || value === undefined || value === '')) {
      return null;
    }

    // String length validation
    if (typeof value === 'string') {
      if (minLength !== undefined && value.length < minLength) {
        return `Minimum length is ${minLength} characters`;
      }
      if (maxLength !== undefined && value.length > maxLength) {
        return `Maximum length is ${maxLength} characters`;
      }
    }

    // Pattern validation
    if (pattern && typeof value === 'string' && !pattern.test(value)) {
      return 'Invalid format';
    }

    // Custom validator
    if (validator) {
      return validator(value);
    }

    return null;
  }, [required, minLength, maxLength, pattern, validator]);

  /**
   * Handle value change
   */
  const onChange = useCallback((newValue: any) => {
    const error = validateOnChange ? validateValue(newValue) : null;

    setState(prev => ({
      ...prev,
      value: newValue,
      error,
      dirty: true
    }));
  }, [validateOnChange, validateValue]);

  /**
   * Handle blur event
   */
  const onBlur = useCallback(() => {
    const error = validateOnBlur ? validateValue(state.value) : null;

    setState(prev => ({
      ...prev,
      error,
      touched: true
    }));
  }, [validateOnBlur, validateValue, state.value]);

  /**
   * Manually trigger validation
   */
  const validate = useCallback((): boolean => {
    const error = validateValue(state.value);

    setState(prev => ({
      ...prev,
      error,
      touched: true
    }));

    return error === null;
  }, [validateValue, state.value]);

  /**
   * Reset field to initial state
   */
  const reset = useCallback(() => {
    setState({
      value: initialValue,
      error: null,
      touched: false,
      dirty: false
    });
  }, [initialValue]);

  /**
   * Set value without validation
   */
  const setValue = useCallback((newValue: any) => {
    setState(prev => ({
      ...prev,
      value: newValue,
      dirty: true
    }));
  }, []);

  return {
    value: state.value,
    error: state.error,
    touched: state.touched,
    dirty: state.dirty,
    isValid: state.error === null,
    onChange,
    onBlur,
    reset,
    validate,
    setValue
  };
};
