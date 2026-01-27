// @ts-nocheck
// Services Export
export { SPService } from './SPService';
export { GraphService } from './GraphService';
export { CacheService, CacheKeys, CacheDurations } from './CacheService';
export type { ICacheEntry } from './CacheService';
export { DocumentTemplateService } from './DocumentTemplateService';
export { DocxTemplateProcessor, docxTemplateProcessor } from './DocxTemplateProcessor';
export type { IDocxProcessingResult, ITemplateDataContext } from './DocxTemplateProcessor';
export { PdfConversionService, pdfConversionService } from './PdfConversionService';
export type { IPdfDocumentOptions, IPdfGenerationResult, IPdfSection } from './PdfConversionService';
export { TemplateVersioningService, createTemplateVersioningService } from './TemplateVersioningService';
export type { ITemplateVersion, IVersionComparison, IVersionChange, ICreateVersionOptions } from './TemplateVersioningService';
export { DocumentApprovalService, createDocumentApprovalService, ApprovalStatus } from './DocumentApprovalService';
export type { IApprovalRequest, IApprover, ICreateApprovalOptions } from './DocumentApprovalService';
export { BulkDocumentService, createBulkDocumentService } from './BulkDocumentService';
export type { IBulkOperationResult, IBulkOperationItem, IBulkGenerationOptions, BulkOperationStatus } from './BulkDocumentService';
export { PowerAutomateService, createPowerAutomateService, FlowTriggerType } from './PowerAutomateService';
export type { IFlowTriggerPayload, IFlowEndpoint, ITriggerResult, IDocumentGeneratedData, IApprovalTriggerData } from './PowerAutomateService';
export { ExcelGenerationService, excelGenerationService } from './ExcelGenerationService';
export type { IExcelGenerationResult, IExcelDocumentOptions, IExcelTableData, IExcelSheetConfig } from './ExcelGenerationService';
export { PowerPointGenerationService, powerPointGenerationService } from './PowerPointGenerationService';
export type { IPowerPointGenerationResult, IPowerPointDocumentOptions, ISlideConfig, SlideType } from './PowerPointGenerationService';
export { logger, LoggingService } from './LoggingService';
export { EmailQueueService, EmailPriority, EmailQueueStatus } from './EmailQueueService';
export type { IEmailQueueItem, IQueueEmailOptions, IQueueResult, IBatchQueueResult } from './EmailQueueService';
export { ImageGenerationService, imageGenerationService } from './ImageGenerationService';
export type { IImageGenerationResult, IImageDocumentOptions, IFloorplanConfig, IFloorplanRoom, IEmergencyRoute, IOrgChartNode } from './ImageGenerationService';
export { DiagramService } from './DiagramService';
export { ScheduledTaskProcessor } from './ScheduledTaskProcessor';
export type { IScheduledProcessingResult, ITaskSLAStatus, IScheduledProcessorConfig } from './ScheduledTaskProcessor';
export { TaskAssignmentService, ConcurrencyErrorType } from './TaskAssignmentService';
export type { ITaskCompletionResult, ITaskUpdateData, IConcurrentUpdateResult } from './TaskAssignmentService';

// Task Completion Validation
export { TaskCompletionValidationService } from './TaskCompletionValidationService';
export type {
  IValidationRule,
  IValidationContext,
  IValidationResult,
  ITaskValidationResponse
} from './TaskCompletionValidationService';

// Real-time Status Sync
export { StatusSyncService } from './StatusSyncService';
export type {
  ChangeType,
  EntityType,
  IChangeEvent,
  IChangeSubscription,
  ISyncConfig
} from './StatusSyncService';

// Financial Management Services
export { ExpenseService } from './ExpenseService';
export { PayrollService } from './PayrollService';

// Document Hub Services
export { DocumentHubService } from './DocumentHubService';
export { DocumentRegistryService } from './DocumentRegistryService';
export { DocumentWorkflowService } from './DocumentWorkflowService';
export { DocumentHubBridgeService } from './DocumentHubBridgeService';

// External Sharing Services
export { ExternalSharingService } from './ExternalSharingService';
export { ExternalSharingAuditService } from './ExternalSharingAuditService';
export { GuestUserService } from './GuestUserService';
export { SharedResourceService } from './SharedResourceService';
export { CrossTenantAccessService } from './CrossTenantAccessService';

// Task Escalation Scheduler - INTEGRATION FIX: Periodic escalation processing
export { TaskEscalationScheduler, createTaskEscalationScheduler } from './TaskEscalationScheduler';
export type { IEscalationRunResult, ISchedulerConfig } from './TaskEscalationScheduler';

// Stakeholder Notification Service - P3 INTEGRATION FIX: Centralized stakeholder notifications
export { StakeholderNotificationService, createStakeholderNotificationService, StakeholderRole, ProcessEventType } from './StakeholderNotificationService';
export type { IStakeholder, IProcessContext, INotificationEventData, IStakeholderNotificationResult } from './StakeholderNotificationService';

// Policy Notification Queue Processor - Processes policy social notifications (share, follow, etc.)
export { PolicyNotificationQueueProcessor, PolicyNotificationType, PolicyNotificationChannel, PolicyNotificationPriority, NotificationQueueStatus } from './PolicyNotificationQueueProcessor';
export type { IPolicyNotificationQueueItem, IPolicyNotificationProcessorConfig, IProcessingResult } from './PolicyNotificationQueueProcessor';
