// @ts-nocheck
/**
 * Workflow Action Handlers
 * Barrel export for all action handler classes
 *
 * @author JML Development Team
 * @version 2.0.0
 */

// Core Action Handlers
export { TaskActionHandler } from './TaskActionHandler';
export { ApprovalActionHandler, ApprovalStatus } from './ApprovalActionHandler';
export { NotificationActionHandler } from './NotificationActionHandler';
export { ListActionHandler } from './ListActionHandler';

// Phase 2: Process-Specific Action Handlers
export { AzureADActionHandler, IAzureADOperationResult } from './AzureADActionHandler';
export { CalendarActionHandler, ICalendarEventDetails, ICalendarOperationResult } from './CalendarActionHandler';
export {
  AssetActionHandler,
  EquipmentRequestStatus,
  AssetReturnStatus,
  LicenseReclamationStatus,
  IEquipmentRequest,
  IAssetReturn,
  ILicenseReclamation
} from './AssetActionHandler';

// Enhanced Workflow Handlers (Phase 1 Implementation)
export { ForEachHandler, IForEachResult } from './ForEachHandler';
export { SubWorkflowHandler, ISubWorkflowResult } from './SubWorkflowHandler';
export { WebhookHandler, IWebhookResult } from './WebhookHandler';
export { RetryHandler, IRetryContext, IRetryResult } from './RetryHandler';
