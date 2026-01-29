// @ts-nocheck
// Document Registry Service
// Service for managing the central document registry

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';
import {
  IDocumentRegistryEntry,
  IDocumentSearchCriteria,
  IDocumentSearchResult,
  IDocumentRegistration,
  IDocumentRegistrationResult,
  IDocumentActivity,
  DocumentStatus,
  SourceModule,
  ConfidentialityLevel,
  ActivityType,
  ActivitySeverity,
  DeviceType
} from '../models';
import { logger } from './LoggingService';

/**
 * Service for Document Registry operations
 */
export class DocumentRegistryService {
  private sp: SPFI;

  private readonly REGISTRY_LIST = 'PM_DocumentRegistry';
  private readonly ACTIVITY_LIST = 'PM_DocumentActivity';
  private readonly SHARING_LIST = 'PM_DocumentSharing';
  private readonly SAVED_SEARCHES_LIST = 'PM_SavedSearches';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ============================================================================
  // BASIC DOCUMENT RETRIEVAL
  // ============================================================================

  /**
   * Get all documents (paginated)
   */
  public async getDocuments(pageSize: number = 100): Promise<IDocumentRegistryEntry[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .select(this.getSelectFields())
        .orderBy('Modified', false)
        .top(pageSize)();

      return items.map(this.mapToRegistryEntry);
    } catch (error) {
      logger.error('DocumentRegistryService', 'Failed to get documents:', error);
      throw error;
    }
  }

  /**
   * Search documents with criteria (alias for search)
   */
  public async searchDocuments(criteria: IDocumentSearchCriteria): Promise<IDocumentRegistryEntry[]> {
    try {
      const result = await this.search(criteria, 1, 100);
      return result.items;
    } catch (error) {
      logger.error('DocumentRegistryService', 'Failed to search documents:', error);
      return [];
    }
  }

  /**
   * Get recent activity across all documents
   */
  public async getRecentActivity(limit: number = 50, offset: number = 0): Promise<IDocumentActivity[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.ACTIVITY_LIST)
        .items
        .select(
          'Id', 'Title', 'DocumentRegistryId', 'DocumentTitle', 'ActivityDocumentId',
          'ActivityType', 'ActivityById', 'ActivityByEmail', 'ActivityDate',
          'ActivityDetails', 'PreviousValue', 'NewValue', 'IPAddress', 'UserAgent',
          'GeoLocation', 'DeviceType', 'SessionId', 'IsSystemAction',
          'RelatedEntityType', 'RelatedEntityId', 'ActivitySeverity'
        )
        .orderBy('ActivityDate', false)
        .skip(offset)
        .top(limit)();

      return items.map(this.mapToActivity);
    } catch (error) {
      logger.error('DocumentRegistryService', 'Failed to get recent activity:', error);
      return [];
    }
  }

  // ============================================================================
  // SAVED SEARCHES
  // ============================================================================

  /**
   * Save a search
   */
  public async saveSearch(search: {
    searchName: string;
    searchQuery: IDocumentSearchCriteria;
    notifyOnNewResults?: boolean;
  }): Promise<number> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.SAVED_SEARCHES_LIST)
        .items
        .add({
          Title: search.searchName,
          SearchQuery: JSON.stringify(search.searchQuery),
          NotifyOnNewResults: search.notifyOnNewResults || false,
          RunCount: 0
        });

      return result.data.Id;
    } catch (error) {
      logger.error('DocumentRegistryService', 'Failed to save search:', error);
      throw error;
    }
  }

  /**
   * Get saved searches for current user
   */
  public async getSavedSearches(): Promise<Array<{
    id: number;
    searchName: string;
    searchQuery: IDocumentSearchCriteria;
    notifyOnNewResults: boolean;
    lastRun?: Date;
    runCount: number;
    createdDate: Date;
    modifiedDate: Date;
  }>> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.SAVED_SEARCHES_LIST)
        .items
        .select('Id', 'Title', 'SearchQuery', 'NotifyOnNewResults', 'LastRun', 'RunCount', 'Created', 'Modified')
        .orderBy('Title')();

      return items.map((item: any) => ({
        id: item.Id,
        searchName: item.Title,
        searchQuery: item.SearchQuery ? JSON.parse(item.SearchQuery) : {},
        notifyOnNewResults: item.NotifyOnNewResults || false,
        lastRun: item.LastRun ? new Date(item.LastRun) : undefined,
        runCount: item.RunCount || 0,
        createdDate: new Date(item.Created),
        modifiedDate: new Date(item.Modified)
      }));
    } catch (error) {
      logger.error('DocumentRegistryService', 'Failed to get saved searches:', error);
      return [];
    }
  }

  // ============================================================================
  // CRUD OPERATIONS
  // ============================================================================

  /**
   * Get document by ID
   */
  public async getById(id: number): Promise<IDocumentRegistryEntry | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .getById(id)
        .select(this.getSelectFields())();

      return this.mapToRegistryEntry(item);
    } catch (error) {
      logger.error('DocumentRegistryService', `Failed to get document ${id}:`, error);
      return null;
    }
  }

  /**
   * Get document by DocumentId (unique identifier)
   */
  public async getByDocumentId(documentId: string): Promise<IDocumentRegistryEntry | null> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .filter(`DocumentId eq '${documentId}'`)
        .select(this.getSelectFields())
        .top(1)();

      return items.length > 0 ? this.mapToRegistryEntry(items[0]) : null;
    } catch (error) {
      logger.error('DocumentRegistryService', `Failed to get document by ID ${documentId}:`, error);
      return null;
    }
  }

  /**
   * Search documents with criteria
   */
  public async search(
    criteria: IDocumentSearchCriteria,
    pageNumber: number = 1,
    pageSize: number = 25
  ): Promise<IDocumentSearchResult> {
    try {
      const filterParts: string[] = [];

      // Build filter query
      if (criteria.searchText) {
        filterParts.push(`substringof('${criteria.searchText}', Title)`);
      }

      if (criteria.sourceModules && criteria.sourceModules.length > 0) {
        const moduleFilters = criteria.sourceModules.map(m => `SourceModule eq '${m}'`);
        filterParts.push(`(${moduleFilters.join(' or ')})`);
      }

      if (criteria.statuses && criteria.statuses.length > 0) {
        const statusFilters = criteria.statuses.map(s => `DocumentStatus eq '${s}'`);
        filterParts.push(`(${statusFilters.join(' or ')})`);
      }

      if (criteria.confidentialityLevels && criteria.confidentialityLevels.length > 0) {
        const confFilters = criteria.confidentialityLevels.map(c => `ConfidentialityLevel eq '${c}'`);
        filterParts.push(`(${confFilters.join(' or ')})`);
      }

      if (criteria.departments && criteria.departments.length > 0) {
        const deptFilters = criteria.departments.map(d => `Department eq '${d}'`);
        filterParts.push(`(${deptFilters.join(' or ')})`);
      }

      if (criteria.ownerId) {
        filterParts.push(`DocumentOwnerId eq ${criteria.ownerId}`);
      }

      if (criteria.onLegalHold !== undefined) {
        filterParts.push(`OnLegalHold eq ${criteria.onLegalHold ? 1 : 0}`);
      }

      if (criteria.isRecord !== undefined) {
        filterParts.push(`IsRecord eq ${criteria.isRecord ? 1 : 0}`);
      }

      if (criteria.dateFrom) {
        filterParts.push(`Created ge '${criteria.dateFrom.toISOString()}'`);
      }

      if (criteria.dateTo) {
        filterParts.push(`Created le '${criteria.dateTo.toISOString()}'`);
      }

      const filterQuery = filterParts.length > 0 ? filterParts.join(' and ') : '';
      const skip = (pageNumber - 1) * pageSize;

      // Get items
      let query = this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .select(this.getSelectFields())
        .orderBy('Modified', false)
        .top(pageSize)
        .skip(skip);

      if (filterQuery) {
        query = query.filter(filterQuery);
      }

      const items = await query();

      // Get total count (separate query for efficiency)
      let countQuery = this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .select('Id');

      if (filterQuery) {
        countQuery = countQuery.filter(filterQuery);
      }

      const allItems = await countQuery();
      const totalCount = allItems.length;

      return {
        items: items.map(this.mapToRegistryEntry),
        totalCount,
        pageNumber,
        pageSize,
        hasMore: skip + items.length < totalCount
      };
    } catch (error) {
      logger.error('DocumentRegistryService', 'Failed to search documents:', error);
      throw error;
    }
  }

  /**
   * Get documents by source module
   */
  public async getBySourceModule(
    module: SourceModule,
    pageSize: number = 50
  ): Promise<IDocumentRegistryEntry[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .filter(`SourceModule eq '${module}'`)
        .select(this.getSelectFields())
        .orderBy('Modified', false)
        .top(pageSize)();

      return items.map(this.mapToRegistryEntry);
    } catch (error) {
      logger.error('DocumentRegistryService', `Failed to get documents for module ${module}:`, error);
      throw error;
    }
  }

  /**
   * Get documents on legal hold
   */
  public async getDocumentsOnLegalHold(): Promise<IDocumentRegistryEntry[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .filter('OnLegalHold eq 1')
        .select(this.getSelectFields())
        .orderBy('LegalHoldAppliedDate', false)();

      return items.map(this.mapToRegistryEntry);
    } catch (error) {
      logger.error('DocumentRegistryService', 'Failed to get documents on legal hold:', error);
      throw error;
    }
  }

  /**
   * Get documents expiring within specified days
   */
  public async getExpiringDocuments(days: number): Promise<IDocumentRegistryEntry[]> {
    try {
      const futureDate = new Date();
      futureDate.setDate(futureDate.getDate() + days);

      const items = await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .filter(`RetentionExpiryDate le '${futureDate.toISOString()}' and DocumentStatus ne 'Archived'`)
        .select(this.getSelectFields())
        .orderBy('RetentionExpiryDate')();

      return items.map(this.mapToRegistryEntry);
    } catch (error) {
      logger.error('DocumentRegistryService', 'Failed to get expiring documents:', error);
      throw error;
    }
  }

  /**
   * Get declared records
   */
  public async getDeclaredRecords(): Promise<IDocumentRegistryEntry[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .filter('IsRecord eq 1')
        .select(this.getSelectFields())
        .orderBy('RecordDeclaredDate', false)();

      return items.map(this.mapToRegistryEntry);
    } catch (error) {
      logger.error('DocumentRegistryService', 'Failed to get declared records:', error);
      throw error;
    }
  }

  // ============================================================================
  // DOCUMENT REGISTRATION (Bridge)
  // ============================================================================

  /**
   * Register a document from another module
   */
  public async registerDocument(
    registration: IDocumentRegistration
  ): Promise<IDocumentRegistrationResult> {
    try {
      // Generate unique document ID
      const documentId = this.generateDocumentId(registration.sourceModule);

      // Create registry entry
      const result = await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .add({
          Title: registration.fileName,
          DocumentId: documentId,
          SourceModule: registration.sourceModule,
          SourceItemId: registration.sourceItemId,
          SourceUrl: registration.sourceUrl,
          FileName: registration.fileName,
          FileExtension: registration.fileExtension,
          FileSizeBytes: registration.fileSizeBytes,
          ClassificationId: registration.classification,
          ConfidentialityLevel: registration.confidentialityLevel || ConfidentialityLevel.Internal,
          Department: registration.department,
          DocumentOwnerId: registration.ownerId,
          DocumentStatus: DocumentStatus.Active,
          DocumentTags: registration.tags ? JSON.stringify(registration.tags) : undefined,
          IsRecord: false,
          OnLegalHold: false,
          ExternalAccessEnabled: false,
          ViewCount: 0,
          DownloadCount: 0,
          VersionCount: 1
        });

      // Log the registration activity
      await this.logActivity({
        DocumentRegistryId: result.data.Id,
        DocumentTitle: registration.fileName,
        ActivityDocumentId: documentId,
        ActivityType: ActivityType.Upload,
        ActivityById: registration.ownerId,
        ActivityDate: new Date(),
        ActivitySeverity: ActivitySeverity.Info,
        IsSystemAction: true,
        ActivityDetails: JSON.stringify({
          sourceModule: registration.sourceModule,
          sourceItemId: registration.sourceItemId
        })
      });

      return {
        success: true,
        documentId,
        registryEntryId: result.data.Id
      };
    } catch (error) {
      logger.error('DocumentRegistryService', 'Failed to register document:', error);
      return {
        success: false,
        errors: [error instanceof Error ? error.message : 'Unknown error']
      };
    }
  }

  /**
   * Sync document from source module (update registry with latest info)
   */
  public async syncFromSource(
    registryId: number,
    updates: Partial<IDocumentRegistryEntry>
  ): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .getById(registryId)
        .update({
          Title: updates.Title,
          FileName: updates.FileName,
          FileSizeBytes: updates.FileSizeBytes,
          LastVersionDate: updates.LastVersionDate?.toISOString(),
          VersionCount: updates.VersionCount
        });

      // Log sync activity
      const doc = await this.getById(registryId);
      if (doc) {
        await this.logActivity({
          DocumentRegistryId: registryId,
          DocumentTitle: doc.Title,
          ActivityDocumentId: doc.DocumentId,
          ActivityType: ActivityType.MetadataUpdated,
          ActivityDate: new Date(),
          ActivitySeverity: ActivitySeverity.Info,
          IsSystemAction: true,
          ActivityDetails: JSON.stringify({ action: 'sync_from_source' })
        });
      }
    } catch (error) {
      logger.error('DocumentRegistryService', `Failed to sync document ${registryId}:`, error);
      throw error;
    }
  }

  // ============================================================================
  // DOCUMENT OPERATIONS
  // ============================================================================

  /**
   * Update document status
   */
  public async updateStatus(
    id: number,
    status: DocumentStatus,
    userId: number
  ): Promise<void> {
    try {
      const doc = await this.getById(id);
      if (!doc) throw new Error('Document not found');

      await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .getById(id)
        .update({ DocumentStatus: status });

      await this.logActivity({
        DocumentRegistryId: id,
        DocumentTitle: doc.Title,
        ActivityDocumentId: doc.DocumentId,
        ActivityType: ActivityType.MetadataUpdated,
        ActivityById: userId,
        ActivityDate: new Date(),
        ActivitySeverity: ActivitySeverity.Info,
        IsSystemAction: false,
        PreviousValue: doc.DocumentStatus,
        NewValue: status,
        ActivityDetails: JSON.stringify({ field: 'DocumentStatus' })
      });
    } catch (error) {
      logger.error('DocumentRegistryService', `Failed to update status for ${id}:`, error);
      throw error;
    }
  }

  /**
   * Declare document as record
   */
  public async declareAsRecord(id: number, userId: number): Promise<void> {
    try {
      const doc = await this.getById(id);
      if (!doc) throw new Error('Document not found');

      if (doc.IsRecord) {
        throw new Error('Document is already declared as a record');
      }

      await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .getById(id)
        .update({
          IsRecord: true,
          RecordDeclaredDate: new Date().toISOString(),
          RecordDeclaredById: userId
        });

      await this.logActivity({
        DocumentRegistryId: id,
        DocumentTitle: doc.Title,
        ActivityDocumentId: doc.DocumentId,
        ActivityType: ActivityType.RecordDeclared,
        ActivityById: userId,
        ActivityDate: new Date(),
        ActivitySeverity: ActivitySeverity.Warning,
        IsSystemAction: false
      });
    } catch (error) {
      logger.error('DocumentRegistryService', `Failed to declare record ${id}:`, error);
      throw error;
    }
  }

  /**
   * Apply legal hold to document
   */
  public async applyLegalHold(
    id: number,
    legalHoldId: number,
    userId: number
  ): Promise<void> {
    try {
      const doc = await this.getById(id);
      if (!doc) throw new Error('Document not found');

      await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .getById(id)
        .update({
          OnLegalHold: true,
          LegalHoldAppliedDate: new Date().toISOString()
        });

      await this.logActivity({
        DocumentRegistryId: id,
        DocumentTitle: doc.Title,
        ActivityDocumentId: doc.DocumentId,
        ActivityType: ActivityType.LegalHoldApplied,
        ActivityById: userId,
        ActivityDate: new Date(),
        ActivitySeverity: ActivitySeverity.Critical,
        IsSystemAction: false,
        ActivityDetails: JSON.stringify({ legalHoldId })
      });
    } catch (error) {
      logger.error('DocumentRegistryService', `Failed to apply legal hold to ${id}:`, error);
      throw error;
    }
  }

  /**
   * Release legal hold from document
   */
  public async releaseLegalHold(
    id: number,
    userId: number,
    reason: string
  ): Promise<void> {
    try {
      const doc = await this.getById(id);
      if (!doc) throw new Error('Document not found');

      await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .getById(id)
        .update({
          OnLegalHold: false
        });

      await this.logActivity({
        DocumentRegistryId: id,
        DocumentTitle: doc.Title,
        ActivityDocumentId: doc.DocumentId,
        ActivityType: ActivityType.LegalHoldReleased,
        ActivityById: userId,
        ActivityDate: new Date(),
        ActivitySeverity: ActivitySeverity.Warning,
        IsSystemAction: false,
        ActivityDetails: JSON.stringify({ reason })
      });
    } catch (error) {
      logger.error('DocumentRegistryService', `Failed to release legal hold from ${id}:`, error);
      throw error;
    }
  }

  /**
   * Log document view
   */
  public async logView(
    id: number,
    userId: number,
    deviceType: DeviceType = DeviceType.Desktop
  ): Promise<void> {
    try {
      const doc = await this.getById(id);
      if (!doc) return;

      // Update view count
      await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .getById(id)
        .update({
          ViewCount: (doc.ViewCount || 0) + 1,
          LastAccessedDate: new Date().toISOString(),
          LastAccessedById: userId
        });

      // Log activity
      await this.logActivity({
        DocumentRegistryId: id,
        DocumentTitle: doc.Title,
        ActivityDocumentId: doc.DocumentId,
        ActivityType: ActivityType.View,
        ActivityById: userId,
        ActivityDate: new Date(),
        ActivitySeverity: ActivitySeverity.Info,
        IsSystemAction: false,
        DeviceType: deviceType
      });
    } catch (error) {
      logger.error('DocumentRegistryService', `Failed to log view for ${id}:`, error);
    }
  }

  /**
   * Log document download
   */
  public async logDownload(
    id: number,
    userId: number,
    deviceType: DeviceType = DeviceType.Desktop
  ): Promise<void> {
    try {
      const doc = await this.getById(id);
      if (!doc) return;

      // Update download count
      await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .getById(id)
        .update({
          DownloadCount: (doc.DownloadCount || 0) + 1,
          LastAccessedDate: new Date().toISOString(),
          LastAccessedById: userId
        });

      // Log activity
      await this.logActivity({
        DocumentRegistryId: id,
        DocumentTitle: doc.Title,
        ActivityDocumentId: doc.DocumentId,
        ActivityType: ActivityType.Download,
        ActivityById: userId,
        ActivityDate: new Date(),
        ActivitySeverity: ActivitySeverity.Info,
        IsSystemAction: false,
        DeviceType: deviceType
      });
    } catch (error) {
      logger.error('DocumentRegistryService', `Failed to log download for ${id}:`, error);
    }
  }

  // ============================================================================
  // ACTIVITY LOGGING
  // ============================================================================

  /**
   * Log document activity
   */
  public async logActivity(activity: Partial<IDocumentActivity>): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.ACTIVITY_LIST)
        .items
        .add({
          Title: `${activity.ActivityType} - ${activity.DocumentTitle}`,
          DocumentRegistryId: activity.DocumentRegistryId,
          DocumentTitle: activity.DocumentTitle,
          ActivityDocumentId: activity.ActivityDocumentId,
          ActivityType: activity.ActivityType,
          ActivityById: activity.ActivityById,
          ActivityByEmail: activity.ActivityByEmail,
          ActivityDate: activity.ActivityDate?.toISOString() || new Date().toISOString(),
          ActivityDetails: activity.ActivityDetails,
          PreviousValue: activity.PreviousValue,
          NewValue: activity.NewValue,
          IPAddress: activity.IPAddress,
          UserAgent: activity.UserAgent,
          GeoLocation: activity.GeoLocation,
          DeviceType: activity.DeviceType || DeviceType.Unknown,
          SessionId: activity.SessionId,
          IsSystemAction: activity.IsSystemAction || false,
          RelatedEntityType: activity.RelatedEntityType,
          RelatedEntityId: activity.RelatedEntityId,
          ActivitySeverity: activity.ActivitySeverity || ActivitySeverity.Info
        });
    } catch (error) {
      logger.error('DocumentRegistryService', 'Failed to log activity:', error);
    }
  }

  /**
   * Get activity for a document
   */
  public async getDocumentActivity(
    documentId: number,
    limit: number = 50
  ): Promise<IDocumentActivity[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.ACTIVITY_LIST)
        .items
        .filter(`DocumentRegistryId eq ${documentId}`)
        .select(
          'Id', 'Title', 'DocumentRegistryId', 'DocumentTitle', 'ActivityDocumentId',
          'ActivityType', 'ActivityById', 'ActivityByEmail', 'ActivityDate',
          'ActivityDetails', 'PreviousValue', 'NewValue', 'IPAddress', 'UserAgent',
          'GeoLocation', 'DeviceType', 'SessionId', 'IsSystemAction',
          'RelatedEntityType', 'RelatedEntityId', 'ActivitySeverity'
        )
        .orderBy('ActivityDate', false)
        .top(limit)();

      return items.map(this.mapToActivity);
    } catch (error) {
      logger.error('DocumentRegistryService', `Failed to get activity for document ${documentId}:`, error);
      return [];
    }
  }

  // ============================================================================
  // PRIVATE HELPER METHODS
  // ============================================================================

  private generateDocumentId(sourceModule: SourceModule): string {
    const prefix = this.getModulePrefix(sourceModule);
    const timestamp = Date.now().toString(36).toUpperCase();
    const random = Math.random().toString(36).substring(2, 6).toUpperCase();
    return `${prefix}-${timestamp}-${random}`;
  }

  private getModulePrefix(sourceModule: SourceModule): string {
    switch (sourceModule) {
      case SourceModule.ContractManager: return 'CON';
      case SourceModule.SigningService: return 'SIG';
      case SourceModule.PolicyHub: return 'POL';
      case SourceModule.ProcessDocuments: return 'PRC';
      case SourceModule.Training: return 'TRN';
      case SourceModule.DocumentHub: return 'DOC';
      default: return 'GEN';
    }
  }

  private getSelectFields(): string {
    return `Id,Title,DocumentId,SourceModule,SourceItemId,SourceUrl,
      ClassificationId,ConfidentialityLevel,Department,DocumentTags,
      DocumentOwnerId,Contributors,DocumentStatus,IsRecord,RecordDeclaredDate,
      RecordDeclaredById,RetentionPolicyId,RetentionStartDate,RetentionExpiryDate,
      DispositionStatus,DispositionDate,DispositionById,OnLegalHold,
      LegalHoldAppliedDate,FileName,FileExtension,FileSizeBytes,ContentHash,
      LastVersionDate,VersionCount,AISummary,AIClassificationConfidence,
      ExtractedEntities,ExtractedKeywords,LanguageDetected,ExternalAccessEnabled,
      ActiveShareCount,ViewCount,DownloadCount,LastAccessedDate,LastAccessedById,
      ReviewDate,ReviewedById,ParentDocumentId,SupersededByDocumentId,
      Created,Modified,AuthorId,EditorId`.replace(/\s/g, '');
  }

  private mapToRegistryEntry(item: any): IDocumentRegistryEntry {
    return {
      Id: item.Id,
      Title: item.Title,
      DocumentId: item.DocumentId,
      SourceModule: item.SourceModule as SourceModule,
      SourceItemId: item.SourceItemId,
      SourceUrl: item.SourceUrl,
      ClassificationId: item.ClassificationId,
      ConfidentialityLevel: item.ConfidentialityLevel as ConfidentialityLevel,
      Department: item.Department,
      DocumentTags: item.DocumentTags ? JSON.parse(item.DocumentTags) : [],
      DocumentOwnerId: item.DocumentOwnerId,
      ContributorIds: item.ContributorsId,
      DocumentStatus: item.DocumentStatus as DocumentStatus,
      IsRecord: item.IsRecord,
      RecordDeclaredDate: item.RecordDeclaredDate ? new Date(item.RecordDeclaredDate) : undefined,
      RecordDeclaredById: item.RecordDeclaredById,
      RetentionPolicyId: item.RetentionPolicyId,
      RetentionStartDate: item.RetentionStartDate ? new Date(item.RetentionStartDate) : undefined,
      RetentionExpiryDate: item.RetentionExpiryDate ? new Date(item.RetentionExpiryDate) : undefined,
      DispositionStatus: item.DispositionStatus,
      DispositionDate: item.DispositionDate ? new Date(item.DispositionDate) : undefined,
      DispositionById: item.DispositionById,
      OnLegalHold: item.OnLegalHold,
      LegalHoldAppliedDate: item.LegalHoldAppliedDate ? new Date(item.LegalHoldAppliedDate) : undefined,
      FileName: item.FileName,
      FileExtension: item.FileExtension,
      FileSizeBytes: item.FileSizeBytes,
      ContentHash: item.ContentHash,
      LastVersionDate: item.LastVersionDate ? new Date(item.LastVersionDate) : undefined,
      VersionCount: item.VersionCount || 1,
      AISummary: item.AISummary,
      AIClassificationConfidence: item.AIClassificationConfidence,
      ExtractedEntities: item.ExtractedEntities,
      ExtractedKeywords: item.ExtractedKeywords ? JSON.parse(item.ExtractedKeywords) : [],
      LanguageDetected: item.LanguageDetected,
      ExternalAccessEnabled: item.ExternalAccessEnabled,
      ActiveShareCount: item.ActiveShareCount || 0,
      ViewCount: item.ViewCount || 0,
      DownloadCount: item.DownloadCount || 0,
      LastAccessedDate: item.LastAccessedDate ? new Date(item.LastAccessedDate) : undefined,
      LastAccessedById: item.LastAccessedById,
      ReviewDate: item.ReviewDate ? new Date(item.ReviewDate) : undefined,
      ReviewedById: item.ReviewedById,
      ParentDocumentId: item.ParentDocumentId,
      SupersededByDocumentId: item.SupersededByDocumentId,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined,
      AuthorId: item.AuthorId,
      EditorId: item.EditorId
    };
  }

  private mapToActivity(item: any): IDocumentActivity {
    return {
      Id: item.Id,
      Title: item.Title,
      DocumentRegistryId: item.DocumentRegistryId,
      DocumentTitle: item.DocumentTitle,
      ActivityDocumentId: item.ActivityDocumentId,
      ActivityType: item.ActivityType as ActivityType,
      ActivityById: item.ActivityById,
      ActivityByEmail: item.ActivityByEmail,
      ActivityDate: new Date(item.ActivityDate),
      ActivityDetails: item.ActivityDetails,
      PreviousValue: item.PreviousValue,
      NewValue: item.NewValue,
      IPAddress: item.IPAddress,
      UserAgent: item.UserAgent,
      GeoLocation: item.GeoLocation,
      DeviceType: item.DeviceType as DeviceType,
      SessionId: item.SessionId,
      IsSystemAction: item.IsSystemAction,
      RelatedEntityType: item.RelatedEntityType,
      RelatedEntityId: item.RelatedEntityId,
      ActivitySeverity: item.ActivitySeverity as ActivitySeverity
    };
  }
}
