// JML Data Models
// Export all interfaces for easy import

export * from './IJmlProcess';
export * from './IJmlChecklistTemplate';
export * from './IJmlTask';
export * from './IJmlTaskAssignment';
export * from './IJmlConfiguration';
export * from './IJmlAuditLog';
export * from './IJmlNotification';
export * from './IJmlTemplateTaskMapping';
export * from './IJmlPresence';
export * from './IJmlSearch';
export * from './IJmlAnalytics';
export * from './IJmlDocument';
export * from './IJmlApproval';
export * from './IJmlTaskComment';
export * from './IJmlTaskTimeEntry';
export * from './IJmlTaskEscalation';
export * from './IJmlTaskAttachment';
export * from './IJmlSavedView';
export * from './ICommon';
export * from './ILicense';
export * from './IPolicy';
export * from './IITProvisioning';
// Export ISigning but exclude SignatureProvider (already exported from IJmlDocument)
export {
  SigningRequestStatus,
  SigningWorkflowType,
  SigningRequestType,
  SignerStatus,
  SignerRole,
  SignatureType,
  SignatureProvider as SigningSignatureProvider,
  SigningBlockType,
  SigningBlockValidation,
  SigningAuditAction,
  SigningEscalationAction,
  SigningTemplateCategory,
  SignerAuthenticationMethod,
  SigningNotificationType,
  type ISigningBlock,
  type ISigningBlockOption,
  type ISigningBlockValidationRule,
  type ISigningBlockCondition,
  type ISigningBlockTemplate,
  type ISigningRequest,
  type ISigningDocument,
  type ISigningChain,
  type ISigningLevel,
  type ILevelConditionalRule,
  type ISigner,
  type ISignatureData,
  type ISignatureStroke,
  type ISigningTemplate,
  type ITemplateSignerConfig,
  type ISigningAuditLog,
  type ISignatureProviderConfig,
  type IProviderSettings,
  type ICreateSigningRequest,
  type ICreateSignerConfig,
  type ISignDocumentRequest,
  type ICompletedBlock,
  type IDeclineSigningRequest,
  type IDelegateSigningRequest,
  type IVoidSigningRequest,
  type IResendSigningRequest,
  type IUpdateSigningRequest,
  type ISigningRequestFilter,
  type ISigningSummary,
  type ISigningAnalytics,
  type ISigningWebhookPayload,
  type ISigningCertificate,
  type ISigningServiceConfig,
} from './ISigning';

// Export IContractManagement excluding conflicting SignatureStatus and SignatureProvider
// (those are already exported from IJmlDocument)
export {
  ContractLifecycleStatus,
  ContractCategory,
  ContractPriority,
  ContractRiskLevel,
  ContractValueType,
  ContractRenewalType,
  ContractTerminationType,
  ClauseCategory,
  ClauseRiskLevel,
  ClauseNegotiability,
  ContractIndustry,
  ContractApprovalStatus,
  ContractApprovalAction,
  ObligationType,
  ObligationStatus,
  ObligationFrequency,
  ContractAuditAction,
  // Export SignatureStatus and SignatureProvider with Contract prefix to avoid conflicts
  SignatureStatus as ContractSignatureStatus,
  SignatureProvider as ContractSignatureProvider,
  // Interfaces
  type IContractRecord,
  type IContractParty,
  type IContractVersion,
  type IContractClause,
  type IContractClauseInstance,
  type IClauseCategoryDef,
  type IContractTemplate,
  type IContractApproval,
  type IContractApprovalRule,
  type IContractSignature,
  type IContractObligation,
  type IContractAuditLog,
  type IContractComment,
  type IContractDocument,
  type IContractNotification,
  type IContractStatistics,
  type IContractDashboard,
  type IContractAlert,
  type IExpiryTimelineItem,
  type IContractFilter,
  type IClauseFilter,
  type IObligationFilter,
  type IContractExportOptions,
  type IBulkOperationResult,
  type IContractFlowConfig,
  type IContractFlowTrigger,
  type IContractJMLIntegration,
  type IContractProcurementLink,
  type IContractAssetLink,
} from './IContractManagement';

// Diagram models
export * from './IDiagram';

// Document Hub models
export * from './IDocumentHub';
export * from './IModuleBridge';

// External Sharing models - explicitly export to avoid DataClassification conflict with IPolicy
export {
  // Enums
  TrustLevel,
  TrustStatus,
  GuestStatus,
  InvitationStatus,
  RiskLevel,
  ResourceType,
  SharingLevel,
  SharedResourceStatus,
  DataClassification as ExternalSharingDataClassification, // Renamed to avoid conflict with IPolicy
  AcknowledgmentStatus,
  RelatedModule,
  AuditActionType,
  AuditResult,
  AccessReviewStatus,
  ReviewStatus,
  AccessReviewType,
  ReviewType,
  ReviewDecision,
  ActionTaken,
  AccessLevel,
  SharingPolicyType,
  // Interfaces
  type ITrustedOrganization,
  type IExternalGuestUser,
  type ISharedResource,
  type IExternalSharingAuditLog,
  type IAuditLogEntryExternal,
  type ISharingPolicy,
  type IAccessReview,
  type ITrustConfig,
  type ITrustConfiguration,
  type ITrustHealthStatus,
  type IGuestFilter,
  type IInvitation,
  type IInvitationResult,
  type IBulkInviteResult,
  type IGuestAccessDetails,
  type IResourceFilter,
  type IShareRequest,
  type IBulkShareResult,
  type IAccessLogEntry,
  type IAuditFilter,
  type IReportOptions,
  type IComplianceReport,
  type IRiskContext,
  type ISecurityAlert,
  type IAccessReviewRequest,
  type IReviewDecision,
  type IReviewPolicy,
  type ICrossTenantPolicy,
  type IPartnerConfiguration,
  type ICrossTenantAccessSettings,
  type IInboundTrustSettings,
  type IExternalSharingKPIs,
  type IActivityFeedItem,
  type IExternalRecipient,
  type IValidationResult,
} from './IExternalSharing';
export * from './ICVManagement';

// Policy Author models
export * from './IPolicyAuthor';
export * from './IPolicyAuthorState';
