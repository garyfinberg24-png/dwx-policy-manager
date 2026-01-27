// @ts-nocheck
/**
 * Module Bridge Interfaces
 *
 * Defines integration interfaces between Document Hub and other JML Enterprise modules:
 * - Contract Manager: Link documents to contracts
 * - Signing Service: Send documents for e-signature
 * - Policy Hub: Apply policies and retention rules
 */

import { IDocumentRegistryEntry, DocumentClassification, ConfidentialityLevel } from './IDocumentHub';
import { ICV, CVStatus, CVSource, ExperienceLevel, EducationLevel, ICVSearchFilters } from './ICVManagement';

// ============================================================================
// CONTRACT MANAGER BRIDGE
// ============================================================================

/**
 * Links a document to a contract
 */
export interface IDocumentContractLink {
  id: number;
  documentId: number;
  contractId: number;
  contractNumber: string;
  contractTitle: string;
  linkType: DocumentContractLinkType;
  linkedDate: Date;
  linkedBy: string;
  notes?: string;
}

/**
 * Type of relationship between document and contract
 */
export enum DocumentContractLinkType {
  /** Main contract document */
  PrimaryContract = 'Primary Contract',
  /** Amendment or addendum */
  Amendment = 'Amendment',
  /** Supporting attachment */
  Attachment = 'Attachment',
  /** Exhibit or schedule */
  Exhibit = 'Exhibit',
  /** Correspondence related to contract */
  Correspondence = 'Correspondence',
  /** Proof of delivery or acceptance */
  ProofOfDelivery = 'Proof of Delivery',
  /** Insurance certificate */
  InsuranceCertificate = 'Insurance Certificate',
  /** Other related document */
  Other = 'Other'
}

/**
 * Contract summary for linking
 */
export interface IContractSummary {
  id: number;
  contractNumber: string;
  title: string;
  status: string;
  partyName: string;
  startDate?: Date;
  endDate?: Date;
  value?: number;
  currency?: string;
}

/**
 * Request to link document to contract
 */
export interface ILinkDocumentToContractRequest {
  documentId: number;
  contractId: number;
  linkType: DocumentContractLinkType;
  notes?: string;
}

/**
 * Contract Manager Bridge Interface
 */
export interface IContractManagerBridge {
  /**
   * Search for contracts to link
   * @param searchText Search term
   * @param maxResults Maximum results to return
   */
  searchContracts(searchText: string, maxResults?: number): Promise<IContractSummary[]>;

  /**
   * Link a document to a contract
   * @param request Link request details
   */
  linkDocumentToContract(request: ILinkDocumentToContractRequest): Promise<IDocumentContractLink>;

  /**
   * Remove link between document and contract
   * @param linkId Link ID to remove
   */
  unlinkDocumentFromContract(linkId: number): Promise<void>;

  /**
   * Get all contracts linked to a document
   * @param documentId Document ID
   */
  getLinkedContracts(documentId: number): Promise<IDocumentContractLink[]>;

  /**
   * Get all documents linked to a contract
   * @param contractId Contract ID
   */
  getLinkedDocuments(contractId: number): Promise<IDocumentRegistryEntry[]>;
}

// ============================================================================
// SIGNING SERVICE BRIDGE
// ============================================================================

/**
 * Signature provider options
 */
export enum SigningProvider {
  Internal = 'Internal',
  DocuSign = 'DocuSign',
  AdobeSign = 'AdobeSign'
}

/**
 * Signature request status
 */
export enum SigningStatus {
  Draft = 'Draft',
  Pending = 'Pending',
  InProgress = 'In Progress',
  Completed = 'Completed',
  Declined = 'Declined',
  Voided = 'Voided',
  Expired = 'Expired'
}

/**
 * Signer role in signing process
 */
export enum SignerRole {
  Signer = 'Signer',
  CarbonCopy = 'Carbon Copy',
  Approver = 'Approver',
  Witness = 'Witness'
}

/**
 * Signer information
 */
export interface IDocumentSigner {
  id?: number;
  name: string;
  email: string;
  role: SignerRole;
  order: number;
  status?: SigningStatus;
  signedDate?: Date;
  declineReason?: string;
}

/**
 * Signing request for a document
 */
export interface IDocumentSigningRequest {
  id?: number;
  documentId: number;
  provider: SigningProvider;
  status: SigningStatus;
  signers: IDocumentSigner[];
  emailSubject: string;
  emailMessage?: string;
  expirationDays?: number;
  reminderDays?: number;
  createdDate?: Date;
  createdBy?: string;
  completedDate?: Date;
  externalEnvelopeId?: string;
}

/**
 * Request to initiate document signing
 */
export interface IInitiateSigningRequest {
  documentId: number;
  provider: SigningProvider;
  signers: Omit<IDocumentSigner, 'id' | 'status' | 'signedDate' | 'declineReason'>[];
  emailSubject: string;
  emailMessage?: string;
  expirationDays?: number;
  reminderDays?: number;
}

/**
 * Signing Service Bridge Interface
 */
export interface ISigningServiceBridge {
  /**
   * Get available signing providers
   */
  getAvailableProviders(): Promise<SigningProvider[]>;

  /**
   * Check if document can be signed
   * @param documentId Document ID
   */
  canSignDocument(documentId: number): Promise<{ canSign: boolean; reason?: string }>;

  /**
   * Initiate signing request for document
   * @param request Signing request details
   */
  initiateSigningRequest(request: IInitiateSigningRequest): Promise<IDocumentSigningRequest>;

  /**
   * Get signing status for document
   * @param documentId Document ID
   */
  getSigningStatus(documentId: number): Promise<IDocumentSigningRequest | null>;

  /**
   * Get signing history for document
   * @param documentId Document ID
   */
  getSigningHistory(documentId: number): Promise<IDocumentSigningRequest[]>;

  /**
   * Cancel/void signing request
   * @param requestId Signing request ID
   * @param reason Reason for cancellation
   */
  cancelSigningRequest(requestId: number, reason: string): Promise<void>;

  /**
   * Resend signing request
   * @param requestId Signing request ID
   */
  resendSigningRequest(requestId: number): Promise<void>;

  /**
   * Download signed document
   * @param requestId Signing request ID
   */
  downloadSignedDocument(requestId: number): Promise<Blob>;
}

// ============================================================================
// POLICY HUB BRIDGE
// ============================================================================

/**
 * Policy classification levels
 */
export enum PolicyClassification {
  Public = 'Public',
  Internal = 'Internal',
  Confidential = 'Confidential',
  Restricted = 'Restricted'
}

/**
 * Policy summary for linking
 */
export interface IPolicySummary {
  id: number;
  policyNumber: string;
  title: string;
  category: string;
  classification: PolicyClassification;
  effectiveDate: Date;
  reviewDate?: Date;
  owner: string;
  status: string;
}

/**
 * Policy document link
 */
export interface IDocumentPolicyLink {
  id: number;
  documentId: number;
  policyId: number;
  policyNumber: string;
  policyTitle: string;
  linkType: DocumentPolicyLinkType;
  linkedDate: Date;
  linkedBy: string;
  notes?: string;
}

/**
 * Type of relationship between document and policy
 */
export enum DocumentPolicyLinkType {
  /** Document implements policy requirement */
  Implementation = 'Implementation',
  /** Document is evidence of policy compliance */
  Evidence = 'Evidence',
  /** Document references the policy */
  Reference = 'Reference',
  /** Document is an exception to policy */
  Exception = 'Exception',
  /** Document is a template for policy */
  Template = 'Template',
  /** Other relationship */
  Other = 'Other'
}

/**
 * Retention rule to apply from policy
 */
export interface IPolicyRetentionRule {
  id: number;
  policyId: number;
  retentionPeriod: number;
  retentionUnit: 'Days' | 'Months' | 'Years';
  dispositionAction: 'Delete' | 'Archive' | 'Review';
  triggerEvent: 'Creation' | 'LastModified' | 'PolicyLinked' | 'Custom';
  customTriggerDate?: Date;
}

/**
 * Request to apply policy to document
 */
export interface IApplyPolicyRequest {
  documentId: number;
  policyId: number;
  linkType: DocumentPolicyLinkType;
  applyRetention: boolean;
  applyClassification: boolean;
  notes?: string;
}

/**
 * Policy Hub Bridge Interface
 */
export interface IPolicyHubBridge {
  /**
   * Search for policies to link
   * @param searchText Search term
   * @param category Optional category filter
   * @param maxResults Maximum results
   */
  searchPolicies(searchText: string, category?: string, maxResults?: number): Promise<IPolicySummary[]>;

  /**
   * Get policy categories
   */
  getPolicyCategories(): Promise<string[]>;

  /**
   * Link document to policy
   * @param request Link request details
   */
  linkDocumentToPolicy(request: IApplyPolicyRequest): Promise<IDocumentPolicyLink>;

  /**
   * Remove link between document and policy
   * @param linkId Link ID to remove
   */
  unlinkDocumentFromPolicy(linkId: number): Promise<void>;

  /**
   * Get all policies linked to a document
   * @param documentId Document ID
   */
  getLinkedPolicies(documentId: number): Promise<IDocumentPolicyLink[]>;

  /**
   * Get retention rule from policy
   * @param policyId Policy ID
   */
  getPolicyRetentionRule(policyId: number): Promise<IPolicyRetentionRule | null>;

  /**
   * Apply policy classification to document
   * @param documentId Document ID
   * @param policyId Policy ID
   */
  applyPolicyClassification(documentId: number, policyId: number): Promise<void>;

  /**
   * Apply policy retention to document
   * @param documentId Document ID
   * @param policyId Policy ID
   */
  applyPolicyRetention(documentId: number, policyId: number): Promise<void>;

  /**
   * Check document compliance with linked policies
   * @param documentId Document ID
   */
  checkPolicyCompliance(documentId: number): Promise<IPolicyComplianceResult>;
}

/**
 * Policy compliance check result
 */
export interface IPolicyComplianceResult {
  documentId: number;
  isCompliant: boolean;
  checkedDate: Date;
  linkedPolicies: number;
  violations: IPolicyViolation[];
  warnings: IPolicyWarning[];
}

/**
 * Policy violation detail
 */
export interface IPolicyViolation {
  policyId: number;
  policyTitle: string;
  requirement: string;
  violationType: 'Classification' | 'Retention' | 'Access' | 'Review' | 'Other';
  severity: 'Critical' | 'High' | 'Medium' | 'Low';
  remediation: string;
}

/**
 * Policy warning detail
 */
export interface IPolicyWarning {
  policyId: number;
  policyTitle: string;
  message: string;
  warningType: 'UpcomingReview' | 'RetentionExpiring' | 'PolicyUpdated' | 'Other';
}

// ============================================================================
// CV MANAGEMENT BRIDGE
// ============================================================================

/**
 * CV summary for document linking
 */
export interface ICVSummary {
  id: number;
  candidateName: string;
  email: string;
  positionAppliedFor?: string;
  department?: string;
  status: CVStatus;
  source: CVSource;
  submissionDate: Date;
  qualificationScore?: number;
  experienceLevel?: ExperienceLevel;
  skills?: string[];
  cvFileUrl?: string;
  cvFileName?: string;
}

/**
 * Type of relationship between document and CV
 */
export enum DocumentCVLinkType {
  /** The actual CV/resume document */
  Resume = 'Resume',
  /** Cover letter */
  CoverLetter = 'Cover Letter',
  /** Certificate or qualification */
  Certificate = 'Certificate',
  /** Portfolio or work sample */
  Portfolio = 'Portfolio',
  /** Reference letter */
  Reference = 'Reference',
  /** Application form */
  Application = 'Application',
  /** Other supporting document */
  Other = 'Other'
}

/**
 * Links a document to a CV record
 */
export interface IDocumentCVLink {
  id: number;
  documentId: number;
  cvId: number;
  candidateName: string;
  candidateEmail: string;
  linkType: DocumentCVLinkType;
  linkedDate: Date;
  linkedBy: string;
  notes?: string;
}

/**
 * Request to link document to CV
 */
export interface ILinkDocumentToCVRequest {
  documentId: number;
  cvId: number;
  linkType: DocumentCVLinkType;
  notes?: string;
}

/**
 * CV search parameters for Document Hub
 */
export interface ICVSearchParams {
  keyword?: string;
  candidateName?: string;
  email?: string;
  status?: CVStatus[];
  source?: CVSource[];
  positionAppliedFor?: string;
  department?: string;
  experienceLevel?: ExperienceLevel[];
  skills?: string[];
  minQualificationScore?: number;
  submissionDateFrom?: Date;
  submissionDateTo?: Date;
  hasDocuments?: boolean;
  maxResults?: number;
}

/**
 * CV Management Bridge Interface
 */
export interface ICVManagementBridge {
  /**
   * Search for CVs
   * @param params Search parameters
   */
  searchCVs(params: ICVSearchParams): Promise<ICVSummary[]>;

  /**
   * Get CV by ID
   * @param cvId CV ID
   */
  getCVById(cvId: number): Promise<ICV | null>;

  /**
   * Get CV summary by ID
   * @param cvId CV ID
   */
  getCVSummary(cvId: number): Promise<ICVSummary | null>;

  /**
   * Link a document to a CV
   * @param request Link request details
   */
  linkDocumentToCV(request: ILinkDocumentToCVRequest): Promise<IDocumentCVLink>;

  /**
   * Remove link between document and CV
   * @param linkId Link ID to remove
   */
  unlinkDocumentFromCV(linkId: number): Promise<void>;

  /**
   * Get all CVs linked to a document
   * @param documentId Document ID
   */
  getLinkedCVs(documentId: number): Promise<IDocumentCVLink[]>;

  /**
   * Get all documents linked to a CV
   * @param cvId CV ID
   */
  getLinkedDocuments(cvId: number): Promise<IDocumentRegistryEntry[]>;

  /**
   * Get CV document (the actual resume file)
   * @param cvId CV ID
   */
  getCVDocument(cvId: number): Promise<{ url: string; fileName: string } | null>;

  /**
   * Get positions for filtering
   */
  getPositions(): Promise<string[]>;

  /**
   * Get departments for filtering
   */
  getDepartments(): Promise<string[]>;

  /**
   * Get available skills for filtering
   */
  getAvailableSkills(): Promise<string[]>;
}

// ============================================================================
// UNIFIED BRIDGE INTERFACE
// ============================================================================

/**
 * Module integration types available
 */
export enum BridgeModuleType {
  ContractManager = 'ContractManager',
  SigningService = 'SigningService',
  PolicyHub = 'PolicyHub',
  CVManagement = 'CVManagement'
}

/**
 * Bridge availability status
 */
export interface IBridgeAvailability {
  module: BridgeModuleType;
  isAvailable: boolean;
  reason?: string;
  version?: string;
}

/**
 * Document integration summary
 */
export interface IDocumentIntegrationSummary {
  documentId: number;
  linkedContracts: number;
  signingRequests: number;
  linkedPolicies: number;
  linkedCVs: number;
  activeSigningRequest: boolean;
  isCompliant: boolean;
  lastChecked: Date;
}

/**
 * Master Module Bridge Interface
 * Provides unified access to all module bridges
 */
export interface IDocumentHubModuleBridge {
  /**
   * Check which modules are available
   */
  getAvailableModules(): Promise<IBridgeAvailability[]>;

  /**
   * Get Contract Manager bridge
   */
  getContractManagerBridge(): IContractManagerBridge;

  /**
   * Get Signing Service bridge
   */
  getSigningServiceBridge(): ISigningServiceBridge;

  /**
   * Get Policy Hub bridge
   */
  getPolicyHubBridge(): IPolicyHubBridge;

  /**
   * Get CV Management bridge
   */
  getCVManagementBridge(): ICVManagementBridge;

  /**
   * Get integration summary for document
   * @param documentId Document ID
   */
  getDocumentIntegrationSummary(documentId: number): Promise<IDocumentIntegrationSummary>;

  /**
   * Get all integrations for a document
   * @param documentId Document ID
   */
  getDocumentIntegrations(documentId: number): Promise<{
    contracts: IDocumentContractLink[];
    signingRequests: IDocumentSigningRequest[];
    policies: IDocumentPolicyLink[];
    cvs: IDocumentCVLink[];
    compliance: IPolicyComplianceResult | null;
  }>;
}
