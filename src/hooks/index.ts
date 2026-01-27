// Hooks Export
export { useProcesses } from './useProcesses';
export { useMyTasks } from './useMyTasks';
export { useFieldValidation } from './useFieldValidation';
export { useFormValidation } from './useFormValidation';
export { useLicense, useModuleAccess } from './useLicense';
export { useEmbeddedNavigation } from './useEmbeddedNavigation';

// Data Caching Hooks
export {
  useDataCache,
  invalidateCache,
  clearCache,
  getCacheEntry,
  setCacheEntry
} from './useDataCache';
export {
  useTemplates,
  useTemplatesByType,
  useTemplateById,
  useTemplatesForProcess,
  invalidateTemplateCache,
  invalidateTemplateById,
  TemplateCacheKeys
} from './useTemplateCache';

export type { IUseProcessesResult } from './useProcesses';
export type { IUseMyTasksResult } from './useMyTasks';
export type {
  IFieldValidationOptions,
  IFieldValidationState,
  IFieldValidationResult
} from './useFieldValidation';
export type {
  IFormField,
  IFormValidationState,
  IFormValidationResult
} from './useFormValidation';
export type { IUseLicenseResult } from './useLicense';
export type { IUseEmbeddedNavigationResult } from './useEmbeddedNavigation';
export type { IUseDataCacheOptions, IUseDataCacheResult } from './useDataCache';

// Dialog Hooks
export {
  useDialog,
  DialogProvider,
  createDialogManager
} from './useDialog';
export type {
  IAlertOptions,
  IConfirmOptions,
  IPromptOptions,
  DialogVariant
} from './useDialog';
