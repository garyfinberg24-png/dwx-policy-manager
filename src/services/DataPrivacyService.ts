// @ts-nocheck
// Data Privacy, GDPR & POPIA Service
// Handles data retention, anonymization, deletion requests, exports, consent management
// Supports GDPR (EU), POPIA (South Africa), and multi-regional compliance

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/site-users/web';
import '@pnp/sp/batching';
import {
  IDataRetentionPolicy,
  IDataDeletionRequest,
  IDataExportRequest,
  IConsentRecord,
  IPrivacyImpactAssessment,
  IAnonymizationJob,
  IAuditLogEntry,
  IDataSubjectRequest,
  IPersonalDataField,
  EntityType,
  DeletionRequestStatus,
  DeletionRequestType,
  ExportFormat,
  ExportRequestStatus,
  AnonymizationMethod,
  JobStatus,
  AuditAction,
  RequestStatus,
  PERSONAL_DATA_FIELDS,
  PersonalDataType,
  ConsentType,
  ConsentMethod,
  PIAStatus,
  RiskLevel,
  // POPIA imports
  PrivacyRegulation,
  POPIACondition,
  POPIALawfulBasis,
  POPIASpecialCategory,
  POPIADataSubjectRight,
  IPOPIAComplianceRecord,
  POPIAComplianceStatus,
  IPOPIADataBreach,
  DataBreachType,
  BreachInvestigationStatus,
  IPOPIAInformationOfficer,
  IPOPIAProcessingRegister,
  IPOPIAConsentRecord,
  IPOPIADataSubjectRequest,
  IPOPIACrossBorderTransfer,
  POPIATransferMechanism,
  IPOPIAChecklistItem,
  IMultiRegulationCompliance
} from '../models/IDataPrivacy';
import * as CryptoJS from 'crypto-js';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export interface IDataPrivacyConfig {
  sp: SPFI;
  currentUserEmail: string;
  currentUserId: string;
  encryptionKey?: string; // For anonymization encryption
}

export interface IAnonymizationResult {
  success: boolean;
  recordsProcessed: number;
  recordsAnonymized: number;
  errors: string[];
}

export interface IExportResult {
  success: boolean;
  downloadUrl?: string;
  fileSize?: number;
  recordCount?: number;
  error?: string;
}

export interface IDeletionResult {
  success: boolean;
  itemsDeleted: number;
  itemsAnonymized: number;
  summary: string;
  errors: string[];
}

export class DataPrivacyService {
  private sp: SPFI;
  private currentUserEmail: string;
  private currentUserId: string;
  private encryptionKey: string;

  constructor(config: IDataPrivacyConfig) {
    this.sp = config.sp;

    // Validate inputs
    ValidationUtils.validateEmail(config.currentUserEmail);
    ValidationUtils.validateUserId(config.currentUserId);

    this.currentUserEmail = config.currentUserEmail;
    this.currentUserId = config.currentUserId;

    // CRITICAL SECURITY: Encryption key is required and must be strong
    if (!config.encryptionKey) {
      throw new Error('Encryption key is required. Configure in Azure Key Vault and pass via IDataPrivacyConfig.');
    }

    if (config.encryptionKey.length < 32) {
      throw new Error('Encryption key must be at least 32 characters for adequate security.');
    }

    // Validate it's not the old default key
    if (config.encryptionKey.includes('default') || config.encryptionKey.includes('change-in-production')) {
      throw new Error('Default encryption key detected. You must use a unique, secure key from Azure Key Vault.');
    }

    this.encryptionKey = config.encryptionKey;
  }

  /**
   * Initialize service and check permissions
   */
  public async initialize(): Promise<void> {
    try {
      // Verify required lists exist
      const lists = [
        'PM_DataRetentionPolicies',
        'PM_DataDeletionRequests',
        'PM_DataExportRequests',
        'PM_ConsentRecords',
        'PM_PrivacyImpactAssessments',
        'PM_AnonymizationJobs',
        'PM_AuditLog',
        'PM_DataSubjectRequests'
      ];

      for (const listTitle of lists) {
        await this.sp.web.lists.getByTitle(listTitle)();
      }

      logger.debug('DataPrivacyService', 'DataPrivacyService initialized successfully');
    } catch (error) {
      logger.error('DataPrivacyService', 'Failed to initialize DataPrivacyService:', error);
      throw new Error('DataPrivacyService initialization failed. Ensure all required lists are created.');
    }
  }

  // ==========================================
  // DATA RETENTION POLICIES
  // ==========================================

  /**
   * Get all active retention policies
   */
  public async getRetentionPolicies(): Promise<IDataRetentionPolicy[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle('PM_DataRetentionPolicies')
        .items.filter('IsActive eq true')
        .orderBy('EntityType')();

      return items.map(item => this.deserializeRetentionPolicy(item));
    } catch (error) {
      logger.error('DataPrivacyService', 'Error getting retention policies:', error);
      throw error;
    }
  }

  /**
   * Get retention policy for specific entity type
   */
  public async getRetentionPolicyByEntity(entityType: EntityType): Promise<IDataRetentionPolicy | null> {
    try {
      // Validate enum value to prevent injection
      ValidationUtils.validateEnum(entityType, EntityType, 'EntityType');

      // Build secure filter
      const filter = `${ValidationUtils.buildFilter('EntityType', 'eq', entityType)} and IsActive eq true`;

      const items = await this.sp.web.lists
        .getByTitle('PM_DataRetentionPolicies')
        .items.filter(filter)
        .top(1)();

      return items.length > 0 ? this.deserializeRetentionPolicy(items[0]) : null;
    } catch (error) {
      logger.error('DataPrivacyService', 'Error getting retention policy:', error);
      return null;
    }
  }

  /**
   * Execute retention policies (delete/anonymize old data)
   */
  public async executeRetentionPolicies(): Promise<{ [entityType: string]: IAnonymizationResult }> {
    try {
      const policies = await this.getRetentionPolicies();
      const results: { [entityType: string]: IAnonymizationResult } = {};

      for (const policy of policies.filter(p => p.AutoDeleteEnabled)) {
        logger.debug('DataPrivacyService', `Executing retention policy for ${policy.EntityType}`);

        const cutoffDate = new Date();
        cutoffDate.setDate(cutoffDate.getDate() - policy.RetentionPeriodDays);

        const result = await this.executeRetentionPolicy(policy, cutoffDate);
        results[policy.EntityType] = result;

        // Update policy execution tracking
        await this.sp.web.lists
          .getByTitle('PM_DataRetentionPolicies')
          .items.getById(policy.Id!)
          .update({
            LastExecuted: new Date().toISOString(),
            ItemsProcessed: (policy.ItemsProcessed || 0) + result.recordsProcessed
          });
      }

      return results;
    } catch (error) {
      logger.error('DataPrivacyService', 'Error executing retention policies:', error);
      throw error;
    }
  }

  /**
   * Execute single retention policy
   */
  private async executeRetentionPolicy(
    policy: IDataRetentionPolicy,
    cutoffDate: Date
  ): Promise<IAnonymizationResult> {
    const result: IAnonymizationResult = {
      success: true,
      recordsProcessed: 0,
      recordsAnonymized: 0,
      errors: []
    };

    try {
      const listName = this.getListNameForEntityType(policy.EntityType);
      if (!listName) {
        result.errors.push(`No list mapping for entity type: ${policy.EntityType}`);
        result.success = false;
        return result;
      }

      // Find items older than retention period
      const filterQuery = `Created lt datetime'${cutoffDate.toISOString()}'`;
      const items = await this.sp.web.lists.getByTitle(listName).items.filter(filterQuery)();

      result.recordsProcessed = items.length;

      if (policy.NotifyBeforeDeletion && policy.NotificationDays) {
        // Check if notification was already sent
        // In production, implement notification logic here
        logger.debug('DataPrivacyService', `Would notify before deletion for ${items.length} items`);
      }

      // Use batch operations for performance (fixes N+1 query problem)
      const [batch] = this.sp.web.batched();
      const batchList = this.sp.web.lists.getByTitle(listName).using(batch);

      for (const item of items) {
        try {
          if (policy.AnonymizeBeforeDelete) {
            // Anonymize personal data fields
            await this.anonymizeItem(listName, item.Id, policy.EntityType);
            result.recordsAnonymized++;
          } else {
            // Hard delete using batch
            batchList.items.getById(item.Id).delete();
            result.recordsProcessed++;
          }

          // Log audit entry (non-batched for tracking)
          await this.logAuditEntry({
            Action: policy.AnonymizeBeforeDelete ? AuditAction.DataAnonymized : AuditAction.DataDeleted,
            EntityType: policy.EntityType,
            EntityId: item.Id,
            Details: JSON.stringify({ policyId: policy.Id, cutoffDate }),
            Success: true
          });
        } catch (itemError) {
          result.errors.push(`Failed to process item ${item.Id}: ${itemError.message}`);
          result.success = false;
        }
      }

      // Execute batch operations - PnP JS v3 batches execute automatically when awaited
      if (!policy.AnonymizeBeforeDelete && items.length > 0) {
        try {
          await batch;
        } catch (batchError) {
          result.errors.push(`Batch delete failed: ${batchError.message}`);
          result.success = false;
        }
      }
    } catch (error) {
      result.errors.push(`Policy execution failed: ${error.message}`);
      result.success = false;
    }

    return result;
  }

  // ==========================================
  // ANONYMIZATION
  // ==========================================

  /**
   * Anonymize personal data for a specific item
   */
  public async anonymizeItem(
    listName: string,
    itemId: number,
    entityType: EntityType
  ): Promise<boolean> {
    try {
      const personalDataFields = PERSONAL_DATA_FIELDS.filter(f => f.listName === listName);
      const updateData: any = {};

      for (const field of personalDataFields) {
        if (field.canAnonymize) {
          const anonymizedValue = this.anonymizeValue(
            `original-value-${itemId}`, // In production, fetch actual value
            field.anonymizationMethod,
            field.dataType
          );
          updateData[field.fieldName] = anonymizedValue;
        }
      }

      await this.sp.web.lists.getByTitle(listName).items.getById(itemId).update(updateData);

      await this.logAuditEntry({
        Action: AuditAction.DataAnonymized,
        EntityType: listName,
        EntityId: itemId,
        Details: JSON.stringify({ fields: Object.keys(updateData) }),
        Success: true
      });

      return true;
    } catch (error) {
      logger.error('DataPrivacyService', 'Error anonymizing item:', error);
      return false;
    }
  }

  /**
   * Anonymize value based on method
   */
  private anonymizeValue(
    value: string,
    method: AnonymizationMethod,
    dataType: PersonalDataType
  ): string {
    if (!value) return value;

    switch (method) {
      case AnonymizationMethod.Hash:
        return CryptoJS.SHA256(value).toString();

      case AnonymizationMethod.Mask:
        if (dataType === PersonalDataType.Email) {
          // mask@*****.com
          const parts = value.split('@');
          return parts.length === 2 ? `${parts[0].substring(0, 2)}***@${parts[1]}` : '***@***.com';
        } else if (dataType === PersonalDataType.Phone) {
          // ***-***-1234
          return `***-***-${value.slice(-4)}`;
        } else {
          // Generic masking
          return value.substring(0, 2) + '*'.repeat(Math.max(value.length - 2, 3));
        }

      case AnonymizationMethod.Replace:
        if (dataType === PersonalDataType.Name) return '[Anonymized User]';
        if (dataType === PersonalDataType.Email) return 'anonymized@system.local';
        if (dataType === PersonalDataType.Phone) return '000-000-0000';
        if (dataType === PersonalDataType.Address) return '[Anonymized Address]';
        return '[Anonymized]';

      case AnonymizationMethod.Generalize:
        if (dataType === PersonalDataType.Address) return '[City/Region Removed]';
        if (dataType === PersonalDataType.LocationData) return '[Location Generalized]';
        return '[Generalized]';

      case AnonymizationMethod.Remove:
        return '';

      case AnonymizationMethod.Encrypt:
        return CryptoJS.AES.encrypt(value, this.encryptionKey).toString();

      default:
        return '[Anonymized]';
    }
  }

  /**
   * Create anonymization job
   */
  public async createAnonymizationJob(
    entityType: EntityType,
    userId?: string,
    dateRangeFrom?: Date,
    dateRangeTo?: Date
  ): Promise<number> {
    try {
      // Validate inputs
      ValidationUtils.validateEnum(entityType, EntityType, 'EntityType');

      if (userId) {
        ValidationUtils.validateUserId(userId);
      }

      if (dateRangeFrom && dateRangeTo) {
        ValidationUtils.validateDateRange(dateRangeFrom, dateRangeTo);
      }

      const fields = PERSONAL_DATA_FIELDS.filter(
        f => f.listName === this.getListNameForEntityType(entityType)
      ).map(f => f.fieldName);

      const job: Partial<IAnonymizationJob> = {
        Title: `Anonymization Job - ${entityType}`,
        EntityType: entityType,
        UserIdToAnonymize: userId,
        DateRangeFrom: dateRangeFrom,
        DateRangeTo: dateRangeTo,
        Fields: fields,
        Method: AnonymizationMethod.Replace,
        Status: JobStatus.Pending,
        RequestedBy: this.currentUserEmail
      };

      const result = await this.sp.web.lists
        .getByTitle('PM_AnonymizationJobs')
        .items.add(this.serializeAnonymizationJob(job));

      return result.data.Id;
    } catch (error) {
      logger.error('DataPrivacyService', 'Error creating anonymization job:', error);
      throw error;
    }
  }

  // ==========================================
  // RIGHT TO BE FORGOTTEN (Data Deletion)
  // ==========================================

  /**
   * Submit data deletion request
   */
  public async submitDeletionRequest(
    requestType: DeletionRequestType,
    subjectUserEmail?: string,
    reason?: string,
    entityTypes?: EntityType[]
  ): Promise<number> {
    try {
      // Validate inputs
      ValidationUtils.validateEnum(requestType, DeletionRequestType, 'DeletionRequestType');

      const targetEmail = subjectUserEmail || this.currentUserEmail;
      ValidationUtils.validateEmail(targetEmail);

      if (entityTypes) {
        entityTypes.forEach(et => ValidationUtils.validateEnum(et, EntityType, 'EntityType'));
      }

      const request: Partial<IDataDeletionRequest> = {
        Title: `Data Deletion Request - ${requestType}`,
        RequesterId: this.currentUserId,
        RequesterEmail: this.currentUserEmail,
        SubjectUserEmail: targetEmail,
        RequestType: requestType,
        Reason: ValidationUtils.sanitizeHtml(reason || ''),
        RequestDate: new Date(),
        Status: DeletionRequestStatus.Pending,
        EntityTypes: entityTypes,
        RetainAuditTrail: true
      };

      const result = await this.sp.web.lists
        .getByTitle('PM_DataDeletionRequests')
        .items.add(this.serializeDeletionRequest(request));

      await this.logAuditEntry({
        Action: AuditAction.DeletionRequested,
        EntityType: 'DataDeletionRequest',
        EntityId: result.data.Id,
        Details: JSON.stringify({ requestType, entityTypes }),
        Success: true
      });

      return result.data.Id;
    } catch (error) {
      logger.error('DataPrivacyService', 'Error submitting deletion request:', error);
      throw error;
    }
  }

  /**
   * Process deletion request (admin/DPO action)
   */
  public async processDeletionRequest(
    requestId: number,
    approved: boolean,
    rejectionReason?: string
  ): Promise<IDeletionResult> {
    const result: IDeletionResult = {
      success: false,
      itemsDeleted: 0,
      itemsAnonymized: 0,
      summary: '',
      errors: []
    };

    try {
      const request = await this.getDeletionRequest(requestId);
      if (!request) {
        throw new Error('Deletion request not found');
      }

      if (!approved) {
        await this.sp.web.lists
          .getByTitle('PM_DataDeletionRequests')
          .items.getById(requestId)
          .update({
            Status: DeletionRequestStatus.Rejected,
            RejectionReason: rejectionReason,
            ProcessedBy: this.currentUserEmail,
            ProcessedDate: new Date().toISOString()
          });

        result.summary = 'Request rejected';
        result.success = true;
        return result;
      }

      // Update status to In Progress
      await this.sp.web.lists
        .getByTitle('PM_DataDeletionRequests')
        .items.getById(requestId)
        .update({
          Status: DeletionRequestStatus.InProgress,
          ApprovedBy: this.currentUserEmail,
          ApprovedDate: new Date().toISOString()
        });

      const entityTypes = request.EntityTypes || Object.values(EntityType);
      const deletionSummary: any = {};

      for (const entityType of entityTypes) {
        const listName = this.getListNameForEntityType(entityType);
        if (!listName) continue;

        try {
          const items = await this.getUserDataItems(listName, request.SubjectUserEmail!);

          if (request.RequestType === DeletionRequestType.Anonymization) {
            // Anonymize items (no batch support for updates with complex logic)
            for (const item of items) {
              await this.anonymizeItem(listName, item.Id, entityType);
              result.itemsAnonymized++;
            }
          } else {
            // Use batch operations for deletion (fixes N+1 query problem)
            if (items.length > 0) {
              const [batch] = this.sp.web.batched();
              const batchList = this.sp.web.lists.getByTitle(listName).using(batch);

              items.forEach(item => {
                batchList.items.getById(item.Id).delete();
              });

              await batch; // PnP JS v3 batches execute automatically when awaited
              result.itemsDeleted += items.length;
            }
          }

          deletionSummary[entityType] = {
            itemsProcessed: items.length,
            action: request.RequestType === DeletionRequestType.Anonymization ? 'anonymized' : 'deleted'
          };
        } catch (error) {
          result.errors.push(`Error processing ${entityType}: ${error.message}`);
        }
      }

      result.summary = JSON.stringify(deletionSummary);
      result.success = result.errors.length === 0;

      // Update request status
      await this.sp.web.lists
        .getByTitle('PM_DataDeletionRequests')
        .items.getById(requestId)
        .update({
          Status: result.success ? DeletionRequestStatus.Completed : DeletionRequestStatus.PartiallyCompleted,
          ProcessedBy: this.currentUserEmail,
          ProcessedDate: new Date().toISOString(),
          CompletedDate: new Date().toISOString(),
          DeletionSummary: result.summary
        });

      await this.logAuditEntry({
        Action: AuditAction.DeletionCompleted,
        EntityType: 'DataDeletionRequest',
        EntityId: requestId,
        Details: result.summary,
        Success: result.success
      });

      return result;
    } catch (error) {
      result.errors.push(`Processing failed: ${error.message}`);
      return result;
    }
  }

  /**
   * Get deletion request by ID
   */
  public async getDeletionRequest(requestId: number): Promise<IDataDeletionRequest | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle('PM_DataDeletionRequests')
        .items.getById(requestId)();

      return this.deserializeDeletionRequest(item);
    } catch (error) {
      logger.error('DataPrivacyService', 'Error getting deletion request:', error);
      return null;
    }
  }

  // ==========================================
  // DATA EXPORT
  // ==========================================

  /**
   * Submit data export request
   */
  public async submitExportRequest(
    format: ExportFormat,
    includeAttachments: boolean = false,
    entityTypes?: EntityType[],
    dateFrom?: Date,
    dateTo?: Date
  ): Promise<number> {
    try {
      // Validate inputs
      ValidationUtils.validateEnum(format, ExportFormat, 'ExportFormat');

      if (dateFrom && dateTo) {
        ValidationUtils.validateDateRange(dateFrom, dateTo);
      }

      if (entityTypes) {
        entityTypes.forEach(et => ValidationUtils.validateEnum(et, EntityType, 'EntityType'));
      }

      const request: Partial<IDataExportRequest> = {
        Title: `Data Export Request - ${format}`,
        RequesterId: this.currentUserId,
        RequesterEmail: this.currentUserEmail,
        ExportFormat: format,
        IncludeAttachments: includeAttachments,
        EntityTypes: entityTypes,
        DateFrom: dateFrom,
        DateTo: dateTo,
        RequestDate: new Date(),
        Status: ExportRequestStatus.Pending
      };

      const result = await this.sp.web.lists
        .getByTitle('PM_DataExportRequests')
        .items.add(this.serializeExportRequest(request));

      await this.logAuditEntry({
        Action: AuditAction.DataExported,
        EntityType: 'DataExportRequest',
        EntityId: result.data.Id,
        Details: JSON.stringify({ format, entityTypes }),
        Success: true
      });

      // Process export asynchronously
      setTimeout(() => this.processExportRequest(result.data.Id), 1000);

      return result.data.Id;
    } catch (error) {
      logger.error('DataPrivacyService', 'Error submitting export request:', error);
      throw error;
    }
  }

  /**
   * Process export request
   */
  private async processExportRequest(requestId: number): Promise<IExportResult> {
    const result: IExportResult = {
      success: false,
      recordCount: 0
    };

    try {
      const item = await this.sp.web.lists
        .getByTitle('PM_DataExportRequests')
        .items.getById(requestId)();

      const request = this.deserializeExportRequest(item);

      // Update status to Processing
      await this.sp.web.lists
        .getByTitle('PM_DataExportRequests')
        .items.getById(requestId)
        .update({ Status: ExportRequestStatus.Processing });

      const entityTypes = request.EntityTypes || Object.values(EntityType);
      const exportData: any = {};
      let totalRecords = 0;

      for (const entityType of entityTypes) {
        const listName = this.getListNameForEntityType(entityType);
        if (!listName) continue;

        const items = await this.getUserDataItems(listName, request.RequesterEmail, request.DateFrom, request.DateTo);
        exportData[entityType] = items;
        totalRecords += items.length;
      }

      // Generate export file
      const exportContent = this.generateExportContent(exportData, request.ExportFormat);
      const fileName = `data-export-${requestId}-${Date.now()}.${this.getFileExtension(request.ExportFormat)}`;

      // In production, upload to document library and get download URL
      const downloadUrl = `/sites/jml/DataExports/${fileName}`;

      // Set expiry date (30 days from now)
      const expiryDate = new Date();
      expiryDate.setDate(expiryDate.getDate() + 30);

      await this.sp.web.lists
        .getByTitle('PM_DataExportRequests')
        .items.getById(requestId)
        .update({
          Status: ExportRequestStatus.Completed,
          ProcessedDate: new Date().toISOString(),
          DownloadUrl: downloadUrl,
          ExpiryDate: expiryDate.toISOString(),
          RecordCount: totalRecords,
          FileSize: exportContent.length
        });

      result.success = true;
      result.downloadUrl = downloadUrl;
      result.recordCount = totalRecords;
      result.fileSize = exportContent.length;

      return result;
    } catch (error) {
      await this.sp.web.lists
        .getByTitle('PM_DataExportRequests')
        .items.getById(requestId)
        .update({ Status: ExportRequestStatus.Failed });

      result.error = error.message;
      return result;
    }
  }

  /**
   * Generate export content based on format
   */
  private generateExportContent(data: any, format: ExportFormat): string {
    switch (format) {
      case ExportFormat.JSON:
        return JSON.stringify(data, null, 2);

      case ExportFormat.CSV:
        // Simple CSV generation (in production, use proper CSV library)
        let csv = '';
        Object.keys(data).forEach(entityType => {
          csv += `\n\n=== ${entityType} ===\n`;
          const items = data[entityType];
          if (items.length > 0) {
            const headers = Object.keys(items[0]).join(',');
            csv += headers + '\n';
            items.forEach((item: any) => {
              csv += Object.values(item).join(',') + '\n';
            });
          }
        });
        return csv;

      case ExportFormat.XML:
        // Simple XML generation
        let xml = '<?xml version="1.0" encoding="UTF-8"?>\n<DataExport>\n';
        Object.keys(data).forEach(entityType => {
          xml += `  <${entityType}>\n`;
          data[entityType].forEach((item: any) => {
            xml += '    <Item>\n';
            Object.keys(item).forEach(key => {
              xml += `      <${key}>${item[key]}</${key}>\n`;
            });
            xml += '    </Item>\n';
          });
          xml += `  </${entityType}>\n`;
        });
        xml += '</DataExport>';
        return xml;

      default:
        return JSON.stringify(data);
    }
  }

  /**
   * Get file extension for export format
   */
  private getFileExtension(format: ExportFormat): string {
    switch (format) {
      case ExportFormat.JSON: return 'json';
      case ExportFormat.CSV: return 'csv';
      case ExportFormat.XML: return 'xml';
      case ExportFormat.PDF: return 'pdf';
      case ExportFormat.Excel: return 'xlsx';
      default: return 'txt';
    }
  }

  // ==========================================
  // CONSENT MANAGEMENT
  // ==========================================

  /**
   * Record user consent
   */
  public async recordConsent(
    consentType: ConsentType,
    purpose: string,
    consentGiven: boolean,
    consentVersion: string,
    consentMethod: string,
    ipAddress?: string,
    userAgent?: string
  ): Promise<number> {
    try {
      // Validate inputs
      ValidationUtils.validateEnum(consentType, ConsentType, 'ConsentType');

      if (!purpose || purpose.trim().length === 0) {
        throw new Error('Consent purpose is required');
      }

      if (!consentVersion || consentVersion.trim().length === 0) {
        throw new Error('Consent version is required');
      }

      if (!consentMethod || consentMethod.trim().length === 0) {
        throw new Error('Consent method is required');
      }

      const consent: Partial<IConsentRecord> = {
        Title: `Consent - ${consentType}`,
        UserId: this.currentUserId,
        UserEmail: this.currentUserEmail,
        ConsentType: consentType,
        Purpose: ValidationUtils.sanitizeHtml(purpose),
        ConsentGiven: consentGiven,
        ConsentDate: new Date(),
        ConsentVersion: ValidationUtils.sanitizeHtml(consentVersion),
        ConsentMethod: ValidationUtils.sanitizeHtml(consentMethod) as ConsentMethod,
        IPAddress: ipAddress ? ValidationUtils.sanitizeHtml(ipAddress) : undefined,
        UserAgent: userAgent ? ValidationUtils.sanitizeHtml(userAgent) : undefined,
        IsActive: consentGiven
      };

      const result = await this.sp.web.lists
        .getByTitle('PM_ConsentRecords')
        .items.add(this.serializeConsentRecord(consent));

      await this.logAuditEntry({
        Action: consentGiven ? AuditAction.ConsentGiven : AuditAction.ConsentWithdrawn,
        EntityType: 'ConsentRecord',
        EntityId: result.data.Id,
        Details: JSON.stringify({ consentType, purpose }),
        Success: true
      });

      return result.data.Id;
    } catch (error) {
      logger.error('DataPrivacyService', 'Error recording consent:', error);
      throw error;
    }
  }

  /**
   * Withdraw consent
   */
  public async withdrawConsent(
    consentId: number,
    reason?: string
  ): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('PM_ConsentRecords')
        .items.getById(consentId)
        .update({
          IsActive: false,
          WithdrawnDate: new Date().toISOString(),
          WithdrawalReason: reason
        });

      await this.logAuditEntry({
        Action: AuditAction.ConsentWithdrawn,
        EntityType: 'ConsentRecord',
        EntityId: consentId,
        Details: JSON.stringify({ reason }),
        Success: true
      });
    } catch (error) {
      logger.error('DataPrivacyService', 'Error withdrawing consent:', error);
      throw error;
    }
  }

  /**
   * Get user consents
   */
  public async getUserConsents(userId?: string): Promise<IConsentRecord[]> {
    try {
      const userEmail = userId || this.currentUserEmail;

      // Validate email
      ValidationUtils.validateEmail(userEmail);

      // Build secure filter
      const filter = ValidationUtils.buildFilter('UserEmail', 'eq', userEmail);

      const items = await this.sp.web.lists
        .getByTitle('PM_ConsentRecords')
        .items.filter(filter)
        .orderBy('ConsentDate', false)();

      return items.map(item => this.deserializeConsentRecord(item));
    } catch (error) {
      logger.error('DataPrivacyService', 'Error getting user consents:', error);
      return [];
    }
  }

  // ==========================================
  // PRIVACY IMPACT ASSESSMENTS
  // ==========================================

  /**
   * Create Privacy Impact Assessment
   */
  public async createPIA(pia: Partial<IPrivacyImpactAssessment>): Promise<number> {
    try {
      // Validate required fields
      if (!pia.ProjectName || pia.ProjectName.trim().length === 0) {
        throw new Error('Project name is required for PIA');
      }

      if (!pia.ProjectDescription || pia.ProjectDescription.trim().length === 0) {
        throw new Error('Project description is required for PIA');
      }

      // Sanitize HTML inputs to prevent XSS
      const newPIA: Partial<IPrivacyImpactAssessment> = {
        ...pia,
        ProjectName: ValidationUtils.sanitizeHtml(pia.ProjectName),
        ProjectDescription: ValidationUtils.sanitizeHtml(pia.ProjectDescription),
        DataController: pia.DataController ? ValidationUtils.sanitizeHtml(pia.DataController) : undefined,
        DataProcessor: pia.DataProcessor ? ValidationUtils.sanitizeHtml(pia.DataProcessor) : undefined,
        ProcessingPurpose: pia.ProcessingPurpose ? ValidationUtils.sanitizeHtml(pia.ProcessingPurpose) : undefined,
        AssessmentDate: new Date(),
        Status: PIAStatus.Draft,
        RiskLevel: this.calculateOverallRiskLevel(pia.Risks || [])
      };

      const result = await this.sp.web.lists
        .getByTitle('PM_PrivacyImpactAssessments')
        .items.add(this.serializePIA(newPIA));

      await this.logAuditEntry({
        Action: AuditAction.PIACreated,
        EntityType: 'PrivacyImpactAssessment',
        EntityId: result.data.Id,
        Details: JSON.stringify({ projectName: pia.ProjectName }),
        Success: true
      });

      return result.data.Id;
    } catch (error) {
      logger.error('DataPrivacyService', 'Error creating PIA:', error);
      throw error;
    }
  }

  /**
   * Calculate overall risk level from individual risks
   */
  private calculateOverallRiskLevel(risks: any[]): RiskLevel {
    if (risks.length === 0) return RiskLevel.Low;

    const maxRiskScore = Math.max(...risks.map(r => r.riskScore || 1));

    if (maxRiskScore >= 7) return RiskLevel.Critical;
    if (maxRiskScore >= 5) return RiskLevel.High;
    if (maxRiskScore >= 3) return RiskLevel.Medium;
    return RiskLevel.Low;
  }

  // ==========================================
  // HELPER METHODS
  // ==========================================

  /**
   * Get list name for entity type
   */
  private getListNameForEntityType(entityType: EntityType): string | null {
    const mapping: { [key in EntityType]?: string } = {
      [EntityType.Process]: 'PM_Processes',
      [EntityType.Task]: 'PM_Tasks',
      [EntityType.Approval]: 'PM_Approvals',
      [EntityType.ApprovalHistory]: 'PM_ApprovalHistory',
      [EntityType.IntegrationLog]: 'PM_IntegrationLogs',
      [EntityType.AIUsageLog]: 'PM_AIUsageLogs',
      [EntityType.UserActivity]: 'PM_UserActivity',
      [EntityType.AuditLog]: 'PM_AuditLog',
      [EntityType.Notification]: 'PM_Notifications',
      [EntityType.Comment]: 'PM_Comments',
      [EntityType.Attachment]: 'PM_Attachments'
    };

    return mapping[entityType] || null;
  }

  /**
   * Get user's data items from a list
   */
  private async getUserDataItems(
    listName: string,
    userEmail: string,
    dateFrom?: Date,
    dateTo?: Date
  ): Promise<any[]> {
    try {
      // Validate inputs
      ValidationUtils.validateEmail(userEmail);
      if (dateFrom) ValidationUtils.validateDateRange(dateFrom, dateTo || new Date());

      // Build secure filter using ValidationUtils
      const emailFilters = [
        ValidationUtils.buildFilter('EmployeeEmail', 'eq', userEmail),
        ValidationUtils.buildFilter('CreatedBy', 'eq', userEmail),
        ValidationUtils.buildFilter('AssignedTo', 'eq', userEmail)
      ];
      let filter = `(${emailFilters.join(' or ')})`;

      if (dateFrom) {
        filter += ` and ${ValidationUtils.buildFilter('Created', 'ge', dateFrom)}`;
      }
      if (dateTo) {
        filter += ` and ${ValidationUtils.buildFilter('Created', 'le', dateTo)}`;
      }

      return await this.sp.web.lists.getByTitle(listName).items.filter(filter)();
    } catch (error) {
      logger.error('DataPrivacyService', `Error getting user data from ${listName}:`, error);
      return [];
    }
  }

  /**
   * Log audit entry
   */
  private async logAuditEntry(entry: Partial<IAuditLogEntry>): Promise<void> {
    try {
      const auditEntry: Partial<IAuditLogEntry> = {
        Title: `${entry.Action} - ${new Date().toISOString()}`,
        Timestamp: new Date(),
        UserId: this.currentUserId,
        UserEmail: this.currentUserEmail,
        ...entry
      };

      await this.sp.web.lists
        .getByTitle('PM_AuditLog')
        .items.add(this.serializeAuditEntry(auditEntry));
    } catch (error) {
      logger.error('DataPrivacyService', 'Error logging audit entry:', error);
    }
  }

  // ==========================================
  // SERIALIZATION HELPERS
  // ==========================================

  private serializeRetentionPolicy(policy: Partial<IDataRetentionPolicy>): any {
    return {
      Title: policy.PolicyName,
      PolicyName: policy.PolicyName,
      EntityType: policy.EntityType,
      RetentionPeriodDays: policy.RetentionPeriodDays,
      AutoDeleteEnabled: policy.AutoDeleteEnabled,
      AnonymizeBeforeDelete: policy.AnonymizeBeforeDelete,
      ApplyToStatus: policy.ApplyToStatus ? JSON.stringify(policy.ApplyToStatus) : null,
      Exceptions: policy.Exceptions,
      IsActive: policy.IsActive,
      LastExecuted: policy.LastExecuted,
      NextExecution: policy.NextExecution,
      ItemsProcessed: policy.ItemsProcessed,
      NotifyBeforeDeletion: policy.NotifyBeforeDeletion,
      NotificationDays: policy.NotificationDays
    };
  }

  private deserializeRetentionPolicy(item: any): IDataRetentionPolicy {
    return {
      ...item,
      ApplyToStatus: item.ApplyToStatus ? JSON.parse(item.ApplyToStatus) : undefined
    };
  }

  private serializeDeletionRequest(request: Partial<IDataDeletionRequest>): any {
    return {
      Title: request.Title,
      RequesterId: request.RequesterId,
      RequesterEmail: request.RequesterEmail,
      SubjectUserId: request.SubjectUserId,
      SubjectUserEmail: request.SubjectUserEmail,
      RequestType: request.RequestType,
      Reason: request.Reason,
      RequestDate: request.RequestDate,
      Status: request.Status,
      EntityTypes: request.EntityTypes ? JSON.stringify(request.EntityTypes) : null,
      RetainAuditTrail: request.RetainAuditTrail
    };
  }

  private deserializeDeletionRequest(item: any): IDataDeletionRequest {
    return {
      ...item,
      EntityTypes: item.EntityTypes ? JSON.parse(item.EntityTypes) : undefined
    };
  }

  private serializeExportRequest(request: Partial<IDataExportRequest>): any {
    return {
      Title: request.Title,
      RequesterId: request.RequesterId,
      RequesterEmail: request.RequesterEmail,
      SubjectUserId: request.SubjectUserId,
      SubjectUserEmail: request.SubjectUserEmail,
      ExportFormat: request.ExportFormat,
      IncludeAttachments: request.IncludeAttachments,
      EntityTypes: request.EntityTypes ? JSON.stringify(request.EntityTypes) : null,
      DateFrom: request.DateFrom,
      DateTo: request.DateTo,
      RequestDate: request.RequestDate,
      Status: request.Status
    };
  }

  private deserializeExportRequest(item: any): IDataExportRequest {
    return {
      ...item,
      EntityTypes: item.EntityTypes ? JSON.parse(item.EntityTypes) : undefined
    };
  }

  private serializeConsentRecord(consent: Partial<IConsentRecord>): any {
    return {
      Title: consent.Title,
      UserId: consent.UserId,
      UserEmail: consent.UserEmail,
      ConsentType: consent.ConsentType,
      Purpose: consent.Purpose,
      ConsentGiven: consent.ConsentGiven,
      ConsentDate: consent.ConsentDate,
      ConsentVersion: consent.ConsentVersion,
      ConsentMethod: consent.ConsentMethod,
      IPAddress: consent.IPAddress,
      UserAgent: consent.UserAgent,
      IsActive: consent.IsActive
    };
  }

  private deserializeConsentRecord(item: any): IConsentRecord {
    return item;
  }

  private serializePIA(pia: Partial<IPrivacyImpactAssessment>): any {
    return {
      Title: pia.ProjectName,
      ProjectName: pia.ProjectName,
      ProjectDescription: pia.ProjectDescription,
      DataController: pia.DataController,
      DataProcessor: pia.DataProcessor,
      AssessmentDate: pia.AssessmentDate,
      Status: pia.Status,
      RiskLevel: pia.RiskLevel,
      PersonalDataTypes: pia.PersonalDataTypes ? JSON.stringify(pia.PersonalDataTypes) : null,
      DataSubjects: pia.DataSubjects ? JSON.stringify(pia.DataSubjects) : null,
      ProcessingPurpose: pia.ProcessingPurpose,
      LegalBasis: pia.LegalBasis ? JSON.stringify(pia.LegalBasis) : null,
      DataSources: pia.DataSources ? JSON.stringify(pia.DataSources) : null,
      Risks: pia.Risks ? JSON.stringify(pia.Risks) : null,
      Mitigations: pia.Mitigations ? JSON.stringify(pia.Mitigations) : null
    };
  }

  private serializeAnonymizationJob(job: Partial<IAnonymizationJob>): any {
    return {
      Title: job.Title,
      EntityType: job.EntityType,
      UserIdToAnonymize: job.UserIdToAnonymize,
      DateRangeFrom: job.DateRangeFrom,
      DateRangeTo: job.DateRangeTo,
      Fields: job.Fields ? JSON.stringify(job.Fields) : null,
      Method: job.Method,
      Status: job.Status,
      RequestedBy: job.RequestedBy
    };
  }

  private serializeAuditEntry(entry: Partial<IAuditLogEntry>): any {
    return {
      Title: entry.Title,
      Timestamp: entry.Timestamp,
      UserId: entry.UserId,
      UserEmail: entry.UserEmail,
      Action: entry.Action,
      EntityType: entry.EntityType,
      EntityId: entry.EntityId,
      Details: entry.Details,
      Success: entry.Success,
      ErrorMessage: entry.ErrorMessage
    };
  }

  // ==========================================
  // POPIA (Protection of Personal Information Act) - South Africa
  // ==========================================

  /**
   * Create or update POPIA compliance record
   */
  public async createPOPIAComplianceRecord(record: Partial<IPOPIAComplianceRecord>): Promise<number> {
    try {
      // Validation
      if (!record.OrganizationName || !record.InformationOfficer) {
        throw new Error('Organization name and Information Officer are required for POPIA compliance');
      }

      ValidationUtils.validateEmail(record.InformationOfficerContact);

      const recordData: any = {
        Title: `POPIA Compliance - ${record.OrganizationName}`,
        OrganizationName: ValidationUtils.sanitizeInput(record.OrganizationName),
        InformationOfficer: ValidationUtils.sanitizeInput(record.InformationOfficer),
        InformationOfficerContact: record.InformationOfficerContact,
        DeputyInformationOfficers: record.DeputyInformationOfficers ? JSON.stringify(record.DeputyInformationOfficers) : null,
        RegistrationNumber: record.RegistrationNumber ? ValidationUtils.sanitizeInput(record.RegistrationNumber) : null,
        RegistrationDate: record.RegistrationDate,
        RegistrationStatus: record.RegistrationStatus || 'Pending',
        ComplianceStatus: record.ComplianceStatus || POPIAComplianceStatus.NotAssessed,
        LastAssessmentDate: record.LastAssessmentDate,
        NextAssessmentDate: record.NextAssessmentDate,
        AssessedBy: record.AssessedBy ? ValidationUtils.sanitizeInput(record.AssessedBy) : null,
        ConditionsMet: record.ConditionsMet ? JSON.stringify(record.ConditionsMet) : null,
        ConditionsNotMet: record.ConditionsNotMet ? JSON.stringify(record.ConditionsNotMet) : null,
        ActionPlanForCompliance: record.ActionPlanForCompliance ? ValidationUtils.sanitizeHtml(record.ActionPlanForCompliance) : null,
        POPIAManualUrl: record.POPIAManualUrl ? ValidationUtils.sanitizeInput(record.POPIAManualUrl) : null,
        PrivacyPolicyUrl: record.PrivacyPolicyUrl ? ValidationUtils.sanitizeInput(record.PrivacyPolicyUrl) : null,
        ProcessingRegisterUrl: record.ProcessingRegisterUrl ? ValidationUtils.sanitizeInput(record.ProcessingRegisterUrl) : null,
        DataBreachProtocolUrl: record.DataBreachProtocolUrl ? ValidationUtils.sanitizeInput(record.DataBreachProtocolUrl) : null,
        CrossBorderTransfersEnabled: record.CrossBorderTransfersEnabled || false,
        TransferMechanisms: record.TransferMechanisms ? JSON.stringify(record.TransferMechanisms) : null,
        TransferCountries: record.TransferCountries ? JSON.stringify(record.TransferCountries) : null,
        SecurityMeasures: record.SecurityMeasures ? JSON.stringify(record.SecurityMeasures) : null,
        EncryptionEnabled: record.EncryptionEnabled || false,
        AccessControlsImplemented: record.AccessControlsImplemented || false,
        IncidentResponsePlanExists: record.IncidentResponsePlanExists || false,
        DataBreachesRecorded: record.DataBreachesRecorded || 0,
        LastBreachDate: record.LastBreachDate,
        BreachNotificationsCompliant: record.BreachNotificationsCompliant || true,
        AuditLogRetentionDays: record.AuditLogRetentionDays || 365,
        ConsentRecordsRetained: record.ConsentRecordsRetained !== false,
        Notes: record.Notes ? ValidationUtils.sanitizeHtml(record.Notes) : null
      };

      const result = await this.sp.web.lists
        .getByTitle('PM_POPIAComplianceRecords')
        .items.add(recordData);

      await this.logAuditEntry({
        Action: AuditAction.ConfigurationChange,
        EntityType: 'POPIAComplianceRecord',
        EntityId: result.data.Id,
        Details: `POPIA compliance record created for ${record.OrganizationName}`,
        Success: true
      });

      return result.data.Id;
    } catch (error) {
      logger.error('DataPrivacyService', 'Error creating POPIA compliance record:', error);
      throw new Error(`Failed to create POPIA compliance record: ${error.message}`);
    }
  }

  /**
   * Record POPIA data breach (72-hour notification requirement)
   */
  public async recordPOPIADataBreach(breach: Partial<IPOPIADataBreach>): Promise<number> {
    try {
      // Validation
      if (!breach.BreachDate || !breach.BreachType || !breach.Description) {
        throw new Error('Breach date, type, and description are required');
      }

      ValidationUtils.validateEnum(breach.BreachType, DataBreachType, 'BreachType');
      ValidationUtils.validateDate(breach.BreachDate, 'BreachDate');

      // Calculate statutory deadline (72 hours from discovery)
      const discoveryDate = breach.DiscoveryDate || new Date();
      const deadline72Hours = new Date(discoveryDate);
      deadline72Hours.setHours(deadline72Hours.getHours() + 72);

      const breachId = `BREACH-${Date.now()}-${Math.random().toString(36).substr(2, 9).toUpperCase()}`;

      const breachData: any = {
        Title: `Data Breach - ${breach.BreachType}`,
        BreachId: breachId,
        BreachDate: breach.BreachDate,
        DiscoveryDate: discoveryDate,
        BreachType: breach.BreachType,
        Severity: breach.Severity || RiskLevel.Medium,
        DataTypesAffected: breach.DataTypesAffected ? JSON.stringify(breach.DataTypesAffected) : null,
        SpecialCategoriesAffected: breach.SpecialCategoriesAffected ? JSON.stringify(breach.SpecialCategoriesAffected) : null,
        NumberOfDataSubjectsAffected: breach.NumberOfDataSubjectsAffected || 0,
        DataSubjectCategories: breach.DataSubjectCategories ? JSON.stringify(breach.DataSubjectCategories) : null,
        Description: ValidationUtils.sanitizeHtml(breach.Description),
        CauseOfBreach: breach.CauseOfBreach ? ValidationUtils.sanitizeHtml(breach.CauseOfBreach) : null,
        SystemsAffected: breach.SystemsAffected ? JSON.stringify(breach.SystemsAffected) : null,
        UnauthorizedAccess: breach.UnauthorizedAccess || false,
        DataExfiltrated: breach.DataExfiltrated || false,
        ContainmentActions: breach.ContainmentActions ? ValidationUtils.sanitizeHtml(breach.ContainmentActions) : null,
        RemediationSteps: breach.RemediationSteps ? ValidationUtils.sanitizeHtml(breach.RemediationSteps) : null,
        PreventiveMeasures: breach.PreventiveMeasures ? ValidationUtils.sanitizeHtml(breach.PreventiveMeasures) : null,
        RegulatoryNotificationRequired: breach.RegulatoryNotificationRequired !== false,
        RegulatoryNotificationDate: breach.RegulatoryNotificationDate,
        RegulatoryNotificationReference: breach.RegulatoryNotificationReference,
        DataSubjectsNotificationRequired: breach.DataSubjectsNotificationRequired || false,
        DataSubjectsNotificationDate: breach.DataSubjectsNotificationDate,
        DataSubjectsNotificationMethod: breach.DataSubjectsNotificationMethod,
        LikelyToResultInHarm: breach.LikelyToResultInHarm !== false,
        HarmAssessmentDetails: breach.HarmAssessmentDetails ? ValidationUtils.sanitizeHtml(breach.HarmAssessmentDetails) : null,
        InvestigationStatus: breach.InvestigationStatus || BreachInvestigationStatus.Initiated,
        InvestigatingOfficer: breach.InvestigatingOfficer,
        InvestigationReport: breach.InvestigationReport ? ValidationUtils.sanitizeHtml(breach.InvestigationReport) : null,
        RootCauseAnalysis: breach.RootCauseAnalysis ? ValidationUtils.sanitizeHtml(breach.RootCauseAnalysis) : null,
        RegulatoryInquiryOpened: breach.RegulatoryInquiryOpened || false,
        EnforcementAction: breach.EnforcementAction,
        Penalties: breach.Penalties,
        Resolution: breach.Resolution ? ValidationUtils.sanitizeHtml(breach.Resolution) : null,
        LessonsLearned: breach.LessonsLearned ? ValidationUtils.sanitizeHtml(breach.LessonsLearned) : null,
        ClosedDate: breach.ClosedDate
      };

      const result = await this.sp.web.lists
        .getByTitle('PM_POPIADataBreaches')
        .items.add(breachData);

      // Create critical audit log entry
      await this.logAuditEntry({
        Action: AuditAction.SecurityIncident,
        EntityType: 'POPIADataBreach',
        EntityId: result.data.Id,
        Details: `POPIA data breach recorded: ${breachId}. Regulatory notification deadline: ${deadline72Hours.toISOString()}`,
        Success: true
      });

      // Warn if approaching 72-hour deadline
      const hoursUntilDeadline = (deadline72Hours.getTime() - new Date().getTime()) / (1000 * 60 * 60);
      if (hoursUntilDeadline < 24 && !breach.RegulatoryNotificationDate) {
        logger.warn('DataPrivacyService', 'URGENT: POPIA breach ${breachId} notification deadline in ${hoursUntilDeadline.toFixed(1)} hours!');
      }

      return result.data.Id;
    } catch (error) {
      logger.error('DataPrivacyService', 'Error recording POPIA data breach:', error);
      throw new Error(`Failed to record data breach: ${error.message}`);
    }
  }

  /**
   * Appoint POPIA Information Officer (Section 55-58)
   */
  public async appointInformationOfficer(officer: Partial<IPOPIAInformationOfficer>): Promise<number> {
    try {
      // Validation
      if (!officer.OfficerName || !officer.OfficerEmail || !officer.PublicContactEmail) {
        throw new Error('Officer name, email, and public contact email are required');
      }

      ValidationUtils.validateEmail(officer.OfficerEmail);
      ValidationUtils.validateEmail(officer.PublicContactEmail);

      const officerData: any = {
        Title: `${officer.OfficerType || 'Information Officer'} - ${officer.OfficerName}`,
        OfficerName: ValidationUtils.sanitizeInput(officer.OfficerName),
        OfficerEmail: officer.OfficerEmail,
        OfficerPhone: officer.OfficerPhone ? ValidationUtils.sanitizeInput(officer.OfficerPhone) : null,
        OfficerType: officer.OfficerType || 'Information Officer',
        Department: officer.Department ? ValidationUtils.sanitizeInput(officer.Department) : null,
        AppointmentDate: officer.AppointmentDate || new Date(),
        TerminationDate: officer.TerminationDate,
        IsActive: officer.IsActive !== false,
        Responsibilities: officer.Responsibilities ? JSON.stringify(officer.Responsibilities) : null,
        TrainingCompleted: officer.TrainingCompleted || false,
        TrainingDate: officer.TrainingDate,
        CertificationUrl: officer.CertificationUrl,
        PublicContactEmail: officer.PublicContactEmail,
        PublicContactPhone: officer.PublicContactPhone,
        OfficeAddress: officer.OfficeAddress ? ValidationUtils.sanitizeInput(officer.OfficeAddress) : null,
        CanApproveProcessing: officer.CanApproveProcessing !== false,
        CanHandleComplaints: officer.CanHandleComplaints !== false,
        CanAuthorizeTransfers: officer.CanAuthorizeTransfers !== false,
        Notes: officer.Notes ? ValidationUtils.sanitizeHtml(officer.Notes) : null
      };

      const result = await this.sp.web.lists
        .getByTitle('PM_POPIAInformationOfficers')
        .items.add(officerData);

      await this.logAuditEntry({
        Action: AuditAction.ConfigurationChange,
        EntityType: 'POPIAInformationOfficer',
        EntityId: result.data.Id,
        Details: `POPIA ${officer.OfficerType} appointed: ${officer.OfficerName}`,
        Success: true
      });

      return result.data.Id;
    } catch (error) {
      logger.error('DataPrivacyService', 'Error appointing Information Officer:', error);
      throw new Error(`Failed to appoint Information Officer: ${error.message}`);
    }
  }

  /**
   * Create processing register entry (POPIA Section 51)
   */
  public async createProcessingRegisterEntry(entry: Partial<IPOPIAProcessingRegister>): Promise<number> {
    try {
      // Validation
      if (!entry.ProcessingPurpose || !entry.ProcessingDescription) {
        throw new Error('Processing purpose and description are required');
      }

      const entryData: any = {
        Title: `Processing: ${entry.ProcessingPurpose}`,
        ProcessingPurpose: ValidationUtils.sanitizeInput(entry.ProcessingPurpose),
        ProcessingDescription: ValidationUtils.sanitizeHtml(entry.ProcessingDescription),
        LawfulBasis: entry.LawfulBasis ? JSON.stringify(entry.LawfulBasis) : null,
        PersonalDataCategories: entry.PersonalDataCategories ? JSON.stringify(entry.PersonalDataCategories) : null,
        SpecialPersonalInfo: entry.SpecialPersonalInfo ? JSON.stringify(entry.SpecialPersonalInfo) : null,
        DataSubjectCategories: entry.DataSubjectCategories ? JSON.stringify(entry.DataSubjectCategories) : null,
        ResponsibleParty: entry.ResponsibleParty ? ValidationUtils.sanitizeInput(entry.ResponsibleParty) : null,
        OperatorInvolved: entry.OperatorInvolved ? ValidationUtils.sanitizeInput(entry.OperatorInvolved) : null,
        DataRecipients: entry.DataRecipients ? JSON.stringify(entry.DataRecipients) : null,
        DataSources: entry.DataSources ? JSON.stringify(entry.DataSources) : null,
        StorageLocation: entry.StorageLocation ? JSON.stringify(entry.StorageLocation) : null,
        CrossBorderTransfer: entry.CrossBorderTransfer || false,
        TransferDestinations: entry.TransferDestinations ? JSON.stringify(entry.TransferDestinations) : null,
        TransferSafeguards: entry.TransferSafeguards ? ValidationUtils.sanitizeInput(entry.TransferSafeguards) : null,
        RetentionPeriod: entry.RetentionPeriod ? ValidationUtils.sanitizeInput(entry.RetentionPeriod) : null,
        RetentionJustification: entry.RetentionJustification ? ValidationUtils.sanitizeHtml(entry.RetentionJustification) : null,
        DisposalMethod: entry.DisposalMethod ? ValidationUtils.sanitizeInput(entry.DisposalMethod) : null,
        SecurityMeasures: entry.SecurityMeasures ? JSON.stringify(entry.SecurityMeasures) : null,
        AccessControls: entry.AccessControls ? ValidationUtils.sanitizeInput(entry.AccessControls) : null,
        EncryptionUsed: entry.EncryptionUsed || false,
        DataProtectionImpactAssessment: entry.DataProtectionImpactAssessment || false,
        PIAReference: entry.PIAReference,
        LastReviewDate: entry.LastReviewDate,
        NextReviewDate: entry.NextReviewDate,
        ConsentRequired: entry.ConsentRequired || false,
        ConsentMechanism: entry.ConsentMechanism,
        ConsentWithdrawalProcess: entry.ConsentWithdrawalProcess,
        IsActive: entry.IsActive !== false,
        Notes: entry.Notes ? ValidationUtils.sanitizeHtml(entry.Notes) : null
      };

      const result = await this.sp.web.lists
        .getByTitle('PM_POPIAProcessingRegister')
        .items.add(entryData);

      await this.logAuditEntry({
        Action: AuditAction.ConfigurationChange,
        EntityType: 'POPIAProcessingRegister',
        EntityId: result.data.Id,
        Details: `Processing register entry created: ${entry.ProcessingPurpose}`,
        Success: true
      });

      return result.data.Id;
    } catch (error) {
      logger.error('DataPrivacyService', 'Error creating processing register entry:', error);
      throw new Error(`Failed to create processing register entry: ${error.message}`);
    }
  }

  /**
   * Record POPIA-compliant consent
   */
  public async recordPOPIAConsent(consent: Partial<IPOPIAConsentRecord>): Promise<number> {
    try {
      // POPIA-specific validation
      if (!consent.ConsentVoluntary) {
        throw new Error('Consent must be voluntary under POPIA');
      }

      if (!consent.ConsentSpecific) {
        throw new Error('Consent must be specific to purpose under POPIA');
      }

      if (!consent.ConsentInformed) {
        throw new Error('Data subject must be informed under POPIA');
      }

      // If child data, parental consent required
      if (consent.DataSubjectIsChild && !consent.ParentalConsentObtained) {
        throw new Error('Parental consent required for processing children\'s personal information (POPIA Section 35-37)');
      }

      // If special category, explicit consent required
      if (consent.InvolvesSpecialCategory && !consent.ConsentGiven) {
        throw new Error('Explicit consent required for special personal information (POPIA Section 26-34)');
      }

      // Record standard consent first
      const consentId = await this.recordConsent(
        consent.ConsentType,
        consent.Purpose,
        consent.ConsentGiven,
        consent.ConsentVersion,
        consent.ConsentMethod,
        consent.IPAddress,
        consent.UserAgent
      );

      // Update with POPIA-specific fields
      await this.sp.web.lists
        .getByTitle('PM_ConsentRecords')
        .items.getById(consentId)
        .update({
          Regulation: PrivacyRegulation.POPIA,
          LawfulBasis: consent.LawfulBasis || POPIALawfulBasis.Consent,
          ConsentVoluntary: consent.ConsentVoluntary,
          ConsentSpecific: consent.ConsentSpecific,
          ConsentInformed: consent.ConsentInformed,
          InvolvesSpecialCategory: consent.InvolvesSpecialCategory || false,
          SpecialCategories: consent.SpecialCategories ? JSON.stringify(consent.SpecialCategories) : null,
          DirectMarketingConsent: consent.DirectMarketingConsent,
          OptOutMechanismProvided: consent.OptOutMechanismProvided !== false,
          DataSubjectIsChild: consent.DataSubjectIsChild || false,
          ParentalConsentObtained: consent.ParentalConsentObtained,
          AgeVerificationMethod: consent.AgeVerificationMethod,
          InvolvedInAutomatedDecisionMaking: consent.InvolvedInAutomatedDecisionMaking || false,
          ProfilingConsent: consent.ProfilingConsent,
          ConsentEvidenceUrl: consent.ConsentEvidenceUrl,
          ConsentLanguage: consent.ConsentLanguage || 'en'
        });

      return consentId;
    } catch (error) {
      logger.error('DataPrivacyService', 'Error recording POPIA consent:', error);
      throw new Error(`Failed to record POPIA consent: ${error.message}`);
    }
  }

  /**
   * Submit POPIA data subject request (30-day response deadline)
   */
  public async submitPOPIADataSubjectRequest(request: Partial<IPOPIADataSubjectRequest>): Promise<number> {
    try {
      // Validate POPIA-specific fields
      ValidationUtils.validateEnum(request.RequestType, POPIADataSubjectRight, 'RequestType');

      // Calculate statutory deadline (30 days from request)
      const requestDate = new Date();
      const deadline = new Date(requestDate);
      deadline.setDate(deadline.getDate() + 30);

      // Create POPIA-specific request directly
      const requestData = {
        Title: `POPIA Request - ${request.RequestType}`,
        RequesterId: this.currentUserId,
        RequesterEmail: this.currentUserEmail,
        RequestType: request.RequestType,
        SubjectUserEmail: request.SubjectUserEmail || this.currentUserEmail,
        Description: request.Description,
        RequestDate: requestDate,
        DueDate: deadline,
        Status: RequestStatus.Received,
        Regulation: PrivacyRegulation.POPIA,
        RequestLanguage: request.RequestLanguage,
        PrescribedForm: request.PrescribedForm,
        IdentityVerified: request.IdentityVerified,
        IdentityVerificationMethod: request.IdentityVerificationMethod,
        StatutoryDeadline: deadline,
        ExtensionGranted: false
      };

      const result = await this.sp.web.lists
        .getByTitle('PM_DataSubjectRequests')
        .items.add(requestData);

      const requestId = result.data.Id;

      // Note: POPIA-specific fields already set in requestData above
      // Update with additional fee-related fields if provided
      if (request.FeeRequired || request.FeeAmount) {
        await this.sp.web.lists
          .getByTitle('PM_DataSubjectRequests')
          .items.getById(requestId)
          .update({
            FeeRequired: request.FeeRequired || false,
            FeeAmount: request.FeeAmount,
            FeeJustification: request.FeeJustification,
            ResponseProvided: false
          });
      }

      await this.logAuditEntry({
        Action: AuditAction.DataAccessRequested,
        EntityType: 'POPIADataSubjectRequest',
        EntityId: requestId,
        Details: `POPIA data subject request submitted. Deadline: ${deadline.toISOString()}`,
        Success: true
      });

      return requestId;
    } catch (error) {
      logger.error('DataPrivacyService', 'Error submitting POPIA data subject request:', error);
      throw new Error(`Failed to submit POPIA data subject request: ${error.message}`);
    }
  }

  /**
   * Assess cross-border data transfer (POPIA Section 72)
   */
  public async assessCrossBorderTransfer(transfer: Partial<IPOPIACrossBorderTransfer>): Promise<number> {
    try {
      // Validation
      if (!transfer.TransferPurpose || !transfer.DestinationCountry || !transfer.DestinationOrganization) {
        throw new Error('Transfer purpose, destination country, and organization are required');
      }

      ValidationUtils.validateEnum(transfer.TransferMechanism, POPIATransferMechanism, 'TransferMechanism');

      const transferData: any = {
        Title: `Cross-Border Transfer - ${transfer.DestinationCountry}`,
        TransferPurpose: ValidationUtils.sanitizeInput(transfer.TransferPurpose),
        DataCategories: transfer.DataCategories ? JSON.stringify(transfer.DataCategories) : null,
        DataSubjectCount: transfer.DataSubjectCount || 0,
        DestinationCountry: ValidationUtils.sanitizeInput(transfer.DestinationCountry),
        DestinationOrganization: ValidationUtils.sanitizeInput(transfer.DestinationOrganization),
        DestinationContact: transfer.DestinationContact,
        AdequacyDecisionExists: transfer.AdequacyDecisionExists || false,
        AdequacyDecisionReference: transfer.AdequacyDecisionReference,
        AlternativeSafeguards: transfer.AlternativeSafeguards,
        TransferMechanism: transfer.TransferMechanism,
        ContractualClauses: transfer.ContractualClauses,
        BindingCorporateRules: transfer.BindingCorporateRules,
        DataSubjectConsentObtained: transfer.DataSubjectConsentObtained || false,
        ConsentRecords: transfer.ConsentRecords,
        DataProtectionGuarantees: transfer.DataProtectionGuarantees ? ValidationUtils.sanitizeHtml(transfer.DataProtectionGuarantees) : null,
        EncryptionInTransit: transfer.EncryptionInTransit !== false,
        SecurityCertifications: transfer.SecurityCertifications ? JSON.stringify(transfer.SecurityCertifications) : null,
        InformationOfficerApproval: transfer.InformationOfficerApproval || false,
        ApprovedBy: transfer.ApprovedBy,
        ApprovalDate: transfer.ApprovalDate,
        OngoingMonitoring: transfer.OngoingMonitoring !== false,
        LastAuditDate: transfer.LastAuditDate,
        NextAuditDate: transfer.NextAuditDate,
        IsActive: transfer.IsActive !== false,
        Notes: transfer.Notes ? ValidationUtils.sanitizeHtml(transfer.Notes) : null
      };

      const result = await this.sp.web.lists
        .getByTitle('PM_POPIACrossBorderTransfers')
        .items.add(transferData);

      await this.logAuditEntry({
        Action: AuditAction.ConfigurationChange,
        EntityType: 'POPIACrossBorderTransfer',
        EntityId: result.data.Id,
        Details: `Cross-border transfer assessment created for ${transfer.DestinationCountry}`,
        Success: true
      });

      return result.data.Id;
    } catch (error) {
      logger.error('DataPrivacyService', 'Error assessing cross-border transfer:', error);
      throw new Error(`Failed to assess cross-border transfer: ${error.message}`);
    }
  }

  /**
   * Generate POPIA compliance checklist
   */
  public generatePOPIAChecklist(): IPOPIAChecklistItem[] {
    return [
      // Condition 1: Accountability
      {
        section: 'Section 8',
        requirement: 'Ensure measures are in place to secure integrity and confidentiality of personal information',
        condition: POPIACondition.Accountability,
        compliant: false
      },
      {
        section: 'Section 55',
        requirement: 'Appoint an Information Officer',
        condition: POPIACondition.Accountability,
        compliant: false
      },
      // Condition 2: Processing Limitation
      {
        section: 'Section 9',
        requirement: 'Process personal information lawfully and reasonably',
        condition: POPIACondition.ProcessingLimitation,
        compliant: false
      },
      {
        section: 'Section 11',
        requirement: 'Obtain consent where required or ensure alternative lawful basis',
        condition: POPIACondition.ProcessingLimitation,
        compliant: false
      },
      // Condition 3: Purpose Specification
      {
        section: 'Section 13',
        requirement: 'Collect personal information for specific, explicitly defined purpose',
        condition: POPIACondition.PurposeSpecification,
        compliant: false
      },
      {
        section: 'Section 18',
        requirement: 'Inform data subjects of purpose of collection',
        condition: POPIACondition.PurposeSpecification,
        compliant: false
      },
      // Condition 4: Further Processing Limitation
      {
        section: 'Section 15',
        requirement: 'Do not use personal information for secondary purposes without consent',
        condition: POPIACondition.FurtherProcessingLimitation,
        compliant: false
      },
      // Condition 5: Information Quality
      {
        section: 'Section 16',
        requirement: 'Ensure personal information is complete, accurate, and not misleading',
        condition: POPIACondition.InformationQuality,
        compliant: false
      },
      {
        section: 'Section 17',
        requirement: 'Update personal information regularly',
        condition: POPIACondition.InformationQuality,
        compliant: false
      },
      // Condition 6: Openness
      {
        section: 'Section 18',
        requirement: 'Notify data subjects when collecting their personal information',
        condition: POPIACondition.Openness,
        compliant: false
      },
      {
        section: 'Section 51',
        requirement: 'Maintain POPIA Manual with processing details',
        condition: POPIACondition.Openness,
        compliant: false
      },
      // Condition 7: Security Safeguards
      {
        section: 'Section 19',
        requirement: 'Implement appropriate technical and organizational security measures',
        condition: POPIACondition.SecuritySafeguards,
        compliant: false
      },
      {
        section: 'Section 22',
        requirement: 'Notify Information Regulator of security breach within 72 hours',
        condition: POPIACondition.SecuritySafeguards,
        compliant: false
      },
      // Condition 8: Data Subject Participation
      {
        section: 'Section 23',
        requirement: 'Allow data subjects to request confirmation of processing and access to information',
        condition: POPIACondition.DataSubjectParticipation,
        compliant: false
      },
      {
        section: 'Section 24',
        requirement: 'Allow data subjects to request correction or deletion of information',
        condition: POPIACondition.DataSubjectParticipation,
        compliant: false
      },
      // Special Personal Information
      {
        section: 'Section 26-34',
        requirement: 'Obtain explicit consent for processing special personal information',
        condition: POPIACondition.ProcessingLimitation,
        compliant: false
      },
      // Cross-Border Transfers
      {
        section: 'Section 72',
        requirement: 'Ensure adequate level of protection for cross-border data transfers',
        condition: POPIACondition.SecuritySafeguards,
        compliant: false
      },
      // Direct Marketing
      {
        section: 'Section 69',
        requirement: 'Obtain consent before sending direct marketing communications',
        condition: POPIACondition.ProcessingLimitation,
        compliant: false
      }
    ];
  }
}
