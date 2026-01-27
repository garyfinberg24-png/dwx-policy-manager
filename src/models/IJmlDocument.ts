// Document Management Models
// Models for document attachment, versioning, and e-signature integration

import { IUser } from './ICommon';

/**
 * Document types for JML processes
 */
export enum DocumentType {
  Contract = 'Contract',
  OfferLetter = 'OfferLetter',
  ExitForm = 'ExitForm',
  PolicyDocument = 'PolicyDocument',
  TrainingMaterial = 'TrainingMaterial',
  AccessForm = 'AccessForm',
  EquipmentForm = 'EquipmentForm',
  NDAAgreement = 'NDAAgreement',
  HandbookAcknowledgment = 'HandbookAcknowledgment',
  BackgroundCheck = 'BackgroundCheck',
  Other = 'Other'
}

/**
 * Signature status for documents requiring e-signature
 */
export enum SignatureStatus {
  NotRequired = 'NotRequired',
  Pending = 'Pending',
  Sent = 'Sent',
  Signed = 'Signed',
  Declined = 'Declined',
  Expired = 'Expired',
  Voided = 'Voided'
}

/**
 * E-signature provider
 */
export enum SignatureProvider {
  DocuSign = 'DocuSign',
  AdobeSign = 'AdobeSign',
  Internal = 'Internal'
}

/**
 * Document interface
 */
export interface IJmlDocument {
  Id: number;
  ProcessID: number;
  FileName: string;
  FileUrl: string;
  FileSize: number;
  ContentType: string;
  DocumentType: DocumentType;
  UploadedBy: IUser;
  UploadedById: number;
  UploadedDate: Date;
  RequiresSignature: boolean;
  SignatureStatus: SignatureStatus;
  SignatureProvider?: SignatureProvider;
  SignatureEnvelopeId?: string;
  SignedBy?: IUser;
  SignedById?: number;
  SignedDate?: Date;
  Version: string;
  VersionHistory?: IJmlDocumentVersion[];
  IsCurrentVersion: boolean;
  CheckedOut: boolean;
  CheckedOutBy?: IUser;
  CheckedOutById?: number;
  CheckedOutDate?: Date;
  Description?: string;
  Tags?: string[];
  ExpirationDate?: Date;
  IsConfidential: boolean;
  Modified: Date;
  ModifiedBy: IUser;
  ModifiedById: number;
}

/**
 * Document version history
 */
export interface IJmlDocumentVersion {
  VersionNumber: string;
  VersionDate: Date;
  VersionAuthor: IUser;
  VersionAuthorId: number;
  VersionComment?: string;
  FileUrl: string;
  FileSize: number;
}

/**
 * Document template
 */
export interface IJmlDocumentTemplate {
  Id: number;
  Title: string;
  Description?: string;
  DocumentType: DocumentType;
  TemplateUrl: string;
  ProcessTypes: string[];
  Placeholders: ITemplatePlaceholder[];
  RequiresSignature: boolean;
  SignatureProvider?: SignatureProvider;
  IsActive: boolean;
  Created: Date;
  CreatedBy: IUser;
  Modified: Date;
  ModifiedBy: IUser;
}

/**
 * Template placeholder for dynamic content
 */
export interface ITemplatePlaceholder {
  Key: string;
  Label: string;
  Description?: string;
  DataType: 'text' | 'date' | 'number' | 'boolean' | 'user' | 'department';
  DefaultValue?: string;
  Required: boolean;
  ValidationPattern?: string;
}

/**
 * Document upload options
 */
export interface IDocumentUploadOptions {
  processId: number;
  documentType: DocumentType;
  requiresSignature?: boolean;
  description?: string;
  tags?: string[];
  expirationDate?: Date;
  isConfidential?: boolean;
}

/**
 * E-signature request
 */
export interface ISignatureRequest {
  documentId: number;
  processId: number;
  provider: SignatureProvider;
  signers: ISignatureRecipient[];
  emailSubject: string;
  emailMessage: string;
  expirationDays?: number;
  reminderDays?: number;
}

/**
 * Signature recipient
 */
export interface ISignatureRecipient {
  email: string;
  name: string;
  role: 'Signer' | 'CC' | 'Approver';
  routingOrder: number;
  requireIdVerification?: boolean;
}

/**
 * Signature webhook event
 */
export interface ISignatureWebhookEvent {
  envelopeId: string;
  status: SignatureStatus;
  documentId: number;
  signedBy?: string;
  signedDate?: Date;
  declineReason?: string;
  voidReason?: string;
}

/**
 * Document search filters
 */
export interface IDocumentSearchFilters {
  processId?: number;
  documentTypes?: DocumentType[];
  signatureStatus?: SignatureStatus[];
  uploadedBy?: number;
  uploadedDateFrom?: Date;
  uploadedDateTo?: Date;
  searchText?: string;
  tags?: string[];
  isConfidential?: boolean;
}

/**
 * Document library view options
 */
export interface IDocumentLibraryViewOptions {
  viewType: 'list' | 'grid' | 'tiles';
  sortBy: 'FileName' | 'UploadedDate' | 'DocumentType' | 'FileSize';
  sortDirection: 'asc' | 'desc';
  groupBy?: 'DocumentType' | 'SignatureStatus' | 'UploadedBy';
  showVersionHistory: boolean;
}

/**
 * File upload progress
 */
export interface IFileUploadProgress {
  fileName: string;
  fileSize: number;
  uploadedBytes: number;
  percentage: number;
  status: 'queued' | 'uploading' | 'completed' | 'error';
  error?: string;
}
