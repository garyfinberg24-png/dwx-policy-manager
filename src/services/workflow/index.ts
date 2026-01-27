// @ts-nocheck
/* eslint-disable */
/**
 * Workflow Engine - Module Exports
 * Central export point for all workflow engine services and handlers
 */

// Core Services
export { WorkflowDefinitionService, IParsedWorkflowDefinition } from './WorkflowDefinitionService';
export { WorkflowInstanceService } from './WorkflowInstanceService';
export { WorkflowEngineService, IStartWorkflowOptions, IExecutionResult } from './WorkflowEngineService';
export { WorkflowConditionEvaluator, IEvaluationContext } from './WorkflowConditionEvaluator';
export { WorkflowActionDispatcher, IActionHandler } from './WorkflowActionDispatcher';
export { WorkflowSchedulerService, ISchedulerResult } from './WorkflowSchedulerService';
export {
  WorkflowResumeService,
  IWorkflowResumeResult,
  IPollingResult,
  ICompletedItem,
  IPollingConfig
} from './WorkflowResumeService';
export {
  WorkflowValidationService,
  ValidationSeverity,
  IExtendedValidationError,
  IComprehensiveValidationResult
} from './WorkflowValidationService';

// Phase 7: Advanced Features
export {
  WorkflowAdvancedService,
  TaskDependencyService,
  MultiLevelApprovalService,
  ScheduledNotificationService,
  ParallelStepService,
  // Task Dependency interfaces
  ITaskDependency,
  IDependencyValidationResult,
  ICascadeUnblockResult,
  // Multi-level Approval interfaces
  IApprovalLevel,
  IMultiLevelApprovalStatus,
  IApprovalLevelStatus,
  IApprovalEscalationResult,
  IUserPendingApproval,
  // Scheduled Notification interfaces
  IScheduledNotification,
  IScheduledNotificationResult,
  // Parallel Step interfaces
  IParallelBranchStatus,
  IParallelExecutionContext,
  IParallelSyncResult
} from './WorkflowAdvancedService';

// Integration Enhancement Services (Phase 7.5)
export {
  PersistentDeadLetterQueueService,
  IPersistentDLQItem,
  IDLQRetryResult
} from './PersistentDeadLetterQueueService';
export {
  NotificationProcessorService,
  NotificationDeliveryStatus,
  INotificationItem,
  INotificationProcessorConfig,
  INotificationProcessingResult
} from './NotificationProcessorService';

// Phase 8: Reliability & Retry Services
export {
  WorkflowResumeRetryService,
  IFailedResumeOperation,
  IResumeRetryConfig,
  IResumeRetryResult,
  IBatchRetryResult
} from './WorkflowResumeRetryService';

// Notification Service
export {
  WorkflowNotificationService,
  WorkflowNotificationEvent,
  NotificationChannel,
  NotificationPriority,
  EmailSendMode,
  IWorkflowNotification,
  INotificationResult,
  IWorkflowNotificationServiceConfig
} from './WorkflowNotificationService';

// Default Workflow Definitions
export {
  DEFAULT_JOINER_WORKFLOW,
  DEFAULT_MOVER_WORKFLOW,
  DEFAULT_LEAVER_WORKFLOW,
  JOINER_WORKFLOW_STEPS,
  MOVER_WORKFLOW_STEPS,
  LEAVER_WORKFLOW_STEPS,
  JOINER_WORKFLOW_VARIABLES,
  MOVER_WORKFLOW_VARIABLES,
  LEAVER_WORKFLOW_VARIABLES,
  ALL_DEFAULT_WORKFLOWS,
  getDefaultWorkflowForProcessType
} from './DefaultWorkflowDefinitions';

// Notification Preferences Service (Phase 3)
export {
  NotificationPreferencesService,
  NotificationChannel as UserNotificationChannel,
  DigestFrequency,
  NotificationEventType,
  IEventChannelPreference,
  IQuietHours,
  IUserNotificationPreferences,
  IResolvedDeliverySettings,
  IDigestItem
} from './NotificationPreferencesService';

// Action Handlers - Core
export { TaskActionHandler } from './handlers/TaskActionHandler';
export { ApprovalActionHandler, ApprovalStatus } from './handlers/ApprovalActionHandler';
export { NotificationActionHandler } from './handlers/NotificationActionHandler';
export { ListActionHandler } from './handlers/ListActionHandler';

// Action Handlers - Phase 2: Process-Specific
export {
  AzureADActionHandler,
  IAzureADOperationResult,
  // Phase 4: User Provisioning interfaces
  IUserProvisioningRequest,
  IUserProvisioningResult,
  // Phase 4: License Management interfaces
  ILicenseAssignmentRequest,
  ILicenseRevocationRequest,
  ILicenseOperationResult,
  IAvailableLicense
} from './handlers/AzureADActionHandler';
export { CalendarActionHandler, ICalendarEventDetails, ICalendarOperationResult } from './handlers/CalendarActionHandler';
export {
  AssetActionHandler,
  EquipmentRequestStatus,
  AssetReturnStatus,
  LicenseReclamationStatus,
  IEquipmentRequest,
  IAssetReturn,
  ILicenseReclamation
} from './handlers/AssetActionHandler';

// Phase 4: Analytics & Intelligence
export {
  WorkflowAnalyticsService,
  IAnalyticsConfig,
  ITimePrediction,
  IPredictionFactor,
  IBottleneck,
  BottleneckType,
  BottleneckSeverity,
  IResourceSuggestion,
  ResourceType,
  SuggestionPriority,
  IAnalyticsReport,
  IWorkflowMetrics,
  IProcessTypeAnalysis,
  ITrendData
} from './WorkflowAnalyticsService';
