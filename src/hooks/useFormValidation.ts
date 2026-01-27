import { useState, useCallback, useMemo } from 'react';

export interface IFormField {
  validate: () => boolean;
  reset: () => void;
  isValid: boolean;
  error: string | null;
}

export interface IFormValidationState {
  isSubmitting: boolean;
  submitCount: number;
  errors: Record<string, string>;
}

export interface IFormValidationResult {
  isSubmitting: boolean;
  submitCount: number;
  errors: Record<string, string>;
  isValid: boolean;
  hasErrors: boolean;
  register: (fieldName: string, field: IFormField) => void;
  unregister: (fieldName: string) => void;
  validateAll: () => boolean;
  handleSubmit: (onSubmit: () => void | Promise<void>) => () => Promise<void>;
  resetForm: () => void;
  setSubmitting: (isSubmitting: boolean) => void;
}

/**
 * Custom hook for form-level validation management
 *
 * Manages multiple field validations, form submission, and error display
 *
 * @returns Form validation state and handlers
 *
 * @example
 * ```typescript
 * const form = useFormValidation();
 *
 * const email = useFieldValidation('', {
 *   required: true,
 *   validator: (value) => {
 *     try {
 *       ValidationUtils.validateEmail(value);
 *       return null;
 *     } catch (error) {
 *       return error.message;
 *     }
 *   }
 * });
 *
 * const password = useFieldValidation('', {
 *   required: true,
 *   minLength: 8
 * });
 *
 * // Register fields
 * useEffect(() => {
 *   form.register('email', email);
 *   form.register('password', password);
 *
 *   return () => {
 *     form.unregister('email');
 *     form.unregister('password');
 *   };
 * }, []);
 *
 * const handleFormSubmit = form.handleSubmit(async () => {
 *   // Submit form data
 *   await saveUser({ email: email.value, password: password.value });
 * });
 *
 * return (
 *   <form onSubmit={handleFormSubmit}>
 *     <Input
 *       value={email.value}
 *       onChange={(e, data) => email.onChange(data.value)}
 *       onBlur={email.onBlur}
 *       validationMessage={email.error}
 *     />
 *     <Input
 *       type="password"
 *       value={password.value}
 *       onChange={(e, data) => password.onChange(data.value)}
 *       onBlur={password.onBlur}
 *       validationMessage={password.error}
 *     />
 *     {form.hasErrors && <FormErrorSummary errors={form.errors} />}
 *     <Button type="submit" disabled={form.isSubmitting}>
 *       {form.isSubmitting ? 'Submitting...' : 'Submit'}
 *     </Button>
 *   </form>
 * );
 * ```
 */
export const useFormValidation = (): IFormValidationResult => {
  const [fields] = useState<Map<string, IFormField>>(new Map());

  const [state, setState] = useState<IFormValidationState>({
    isSubmitting: false,
    submitCount: 0,
    errors: {}
  });

  /**
   * Register a field for validation
   */
  const register = useCallback((fieldName: string, field: IFormField) => {
    fields.set(fieldName, field);
  }, [fields]);

  /**
   * Unregister a field
   */
  const unregister = useCallback((fieldName: string) => {
    fields.delete(fieldName);
  }, [fields]);

  /**
   * Validate all registered fields
   */
  const validateAll = useCallback((): boolean => {
    const errors: Record<string, string> = {};
    let isValid = true;

    fields.forEach((field, fieldName) => {
      const fieldValid = field.validate();
      if (!fieldValid && field.error) {
        errors[fieldName] = field.error;
        isValid = false;
      }
    });

    setState(prev => ({
      ...prev,
      errors
    }));

    return isValid;
  }, [fields]);

  /**
   * Handle form submission with validation
   */
  const handleSubmit = useCallback((onSubmit: () => void | Promise<void>) => {
    return async (e?: React.FormEvent) => {
      if (e) {
        e.preventDefault();
      }

      setState(prev => ({
        ...prev,
        submitCount: prev.submitCount + 1
      }));

      // Validate all fields
      const isValid = validateAll();

      if (!isValid) {
        return;
      }

      // Set submitting state
      setState(prev => ({
        ...prev,
        isSubmitting: true
      }));

      try {
        await onSubmit();

        // Clear errors on successful submit
        setState(prev => ({
          ...prev,
          errors: {},
          isSubmitting: false
        }));
      } catch (error) {
        // Handle submission error
        console.error('Form submission error:', error);

        setState(prev => ({
          ...prev,
          isSubmitting: false,
          errors: {
            submit: error instanceof Error ? error.message : 'Submission failed'
          }
        }));
      }
    };
  }, [validateAll]);

  /**
   * Reset the entire form
   */
  const resetForm = useCallback(() => {
    fields.forEach(field => field.reset());

    setState({
      isSubmitting: false,
      submitCount: 0,
      errors: {}
    });
  }, [fields]);

  /**
   * Set submitting state manually
   */
  const setSubmitting = useCallback((isSubmitting: boolean) => {
    setState(prev => ({
      ...prev,
      isSubmitting
    }));
  }, []);

  /**
   * Check if form is valid
   */
  const isValid = useMemo(() => {
    return Array.from(fields.values()).every(field => field.isValid);
  }, [fields, state.errors]); // Re-compute when errors change

  /**
   * Check if form has errors
   */
  const hasErrors = useMemo(() => {
    return Object.keys(state.errors).length > 0;
  }, [state.errors]);

  return {
    isSubmitting: state.isSubmitting,
    submitCount: state.submitCount,
    errors: state.errors,
    isValid,
    hasErrors,
    register,
    unregister,
    validateAll,
    handleSubmit,
    resetForm,
    setSubmitting
  };
};
