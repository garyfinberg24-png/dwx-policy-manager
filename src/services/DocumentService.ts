// @ts-nocheck
// Document Service
// Handles document upload, versioning, and management for JML processes

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/fields';
import {
  IJmlDocument,
  IDocumentUploadOptions,
  IDocumentSearchFilters,
  IFileUploadProgress,
  DocumentType,
  SignatureStatus
} from '../models';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class DocumentService {
  private sp: SPFI;
  private readonly DOCUMENT_LIBRARY = 'PM_Documents';
  private readonly CHUNK_SIZE = 10485760; // 10MB chunks for large files

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Upload a document to the process document library
   */
  public async uploadDocument(
    file: File,
    options: IDocumentUploadOptions,
    onProgress?: (progress: IFileUploadProgress) => void
  ): Promise<IJmlDocument> {
    try {
      const folderPath = `Process_${options.processId}`;

      // Ensure folder exists
      await this.ensureProcessFolder(options.processId);

      // Upload file
      const uploadResult = await this.uploadFileWithProgress(
        file,
        folderPath,
        onProgress
      );

      // Create metadata item
      const metadata = {
        ProcessID: options.processId,
        DocumentType: options.documentType,
        RequiresSignature: options.requiresSignature || false,
        SignatureStatus: options.requiresSignature ? SignatureStatus.Pending : SignatureStatus.NotRequired,
        Description: options.description,
        Tags: options.tags ? JSON.stringify(options.tags) : undefined,
        ExpirationDate: options.expirationDate?.toISOString(),
        IsConfidential: options.isConfidential || false,
        Version: '1.0',
        IsCurrentVersion: true,
        CheckedOut: false
      };

      // Update file metadata
      const item = await uploadResult.file.getItem();
      await item.update(metadata);

      // Get the created document
      return await this.getDocumentById(item.Id);
    } catch (error) {
      logger.error('DocumentService', 'Failed to upload document:', error);
      throw error;
    }
  }

  /**
   * Upload file with progress tracking
   */
  private async uploadFileWithProgress(
    file: File,
    folderPath: string,
    onProgress?: (progress: IFileUploadProgress) => void
  ): Promise<any> {
    const progress: IFileUploadProgress = {
      fileName: file.name,
      fileSize: file.size,
      uploadedBytes: 0,
      percentage: 0,
      status: 'uploading'
    };

    if (onProgress) {
      onProgress(progress);
    }

    // For small files, use simple upload
    if (file.size < this.CHUNK_SIZE) {
      const result = await this.sp.web.lists
        .getByTitle(this.DOCUMENT_LIBRARY)
        .rootFolder
        .folders.getByUrl(folderPath)
        .files.addUsingPath(file.name, file, { Overwrite: true });

      progress.uploadedBytes = file.size;
      progress.percentage = 100;
      progress.status = 'completed';

      if (onProgress) {
        onProgress(progress);
      }

      return result;
    }

    // For large files, use chunked upload
    return await this.uploadLargeFile(file, folderPath, onProgress);
  }

  /**
   * Upload large file in chunks
   */
  private async uploadLargeFile(
    file: File,
    folderPath: string,
    onProgress?: (progress: IFileUploadProgress) => void
  ): Promise<any> {
    const chunkSize = this.CHUNK_SIZE;
    const chunks = Math.ceil(file.size / chunkSize);

    let uploadedBytes = 0;
    const folder = this.sp.web.lists
      .getByTitle(this.DOCUMENT_LIBRARY)
      .rootFolder
      .folders.getByUrl(folderPath);

    // Start chunked upload
    const uploadId = await folder.files.addChunked(
      file.name,
      file,
      (data) => {
        uploadedBytes += data.blockNumber === chunks ? file.size % chunkSize : chunkSize;

        if (onProgress) {
          const progress: IFileUploadProgress = {
            fileName: file.name,
            fileSize: file.size,
            uploadedBytes,
            percentage: Math.round((uploadedBytes / file.size) * 100),
            status: 'uploading'
          };
          onProgress(progress);
        }
      },
      true
    );

    if (onProgress) {
      const progress: IFileUploadProgress = {
        fileName: file.name,
        fileSize: file.size,
        uploadedBytes: file.size,
        percentage: 100,
        status: 'completed'
      };
      onProgress(progress);
    }

    return uploadId;
  }

  /**
   * Ensure process folder exists
   */
  private async ensureProcessFolder(processId: number): Promise<void> {
    try {
      const folderName = `Process_${processId}`;
      const library = this.sp.web.lists.getByTitle(this.DOCUMENT_LIBRARY);

      try {
        await library.rootFolder.folders.getByUrl(folderName)();
      } catch {
        // Folder doesn't exist, create it
        await library.rootFolder.folders.addUsingPath(folderName);
      }
    } catch (error) {
      logger.error('DocumentService', 'Failed to ensure process folder:', error);
      throw error;
    }
  }

  /**
   * Get document by ID
   */
  public async getDocumentById(documentId: number): Promise<IJmlDocument> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.DOCUMENT_LIBRARY)
        .items.getById(documentId)
        .select(
          '*',
          'File/Name',
          'File/ServerRelativeUrl',
          'File/Length',
          'File/TimeLastModified',
          'UploadedBy/Title',
          'UploadedBy/EMail',
          'SignedBy/Title',
          'SignedBy/EMail',
          'CheckedOutBy/Title',
          'CheckedOutBy/EMail',
          'Author/Title',
          'Author/EMail',
          'Editor/Title',
          'Editor/EMail'
        )
        .expand(
          'File',
          'UploadedBy',
          'SignedBy',
          'CheckedOutBy',
          'Author',
          'Editor'
        )();

      return this.mapToDocument(item);
    } catch (error) {
      logger.error('DocumentService', 'Failed to get document:', error);
      throw error;
    }
  }

  /**
   * Get all documents for a process
   */
  public async getProcessDocuments(processId: number, filters?: IDocumentSearchFilters): Promise<IJmlDocument[]> {
    try {
      // Validate process ID
      const validProcessId = ValidationUtils.validateInteger(processId, 'processId', 1);

      // Build base filter
      const baseFilter = ValidationUtils.buildFilter('ProcessID', 'eq', validProcessId);
      let filterQuery = baseFilter;

      if (filters) {
        // Build document type filters (enum values - validated)
        if (filters.documentTypes && filters.documentTypes.length > 0) {
          const typeFilters: string[] = [];
          for (let i = 0; i < filters.documentTypes.length; i++) {
            ValidationUtils.validateEnum(filters.documentTypes[i], DocumentType, 'DocumentType');
            typeFilters.push(ValidationUtils.buildFilter('DocumentType', 'eq', filters.documentTypes[i]));
          }
          filterQuery += ` and (${typeFilters.join(' or ')})`;
        }

        // Build signature status filters (enum values - validated)
        if (filters.signatureStatus && filters.signatureStatus.length > 0) {
          const statusFilters: string[] = [];
          for (let i = 0; i < filters.signatureStatus.length; i++) {
            ValidationUtils.validateEnum(filters.signatureStatus[i], SignatureStatus, 'SignatureStatus');
            statusFilters.push(ValidationUtils.buildFilter('SignatureStatus', 'eq', filters.signatureStatus[i]));
          }
          filterQuery += ` and (${statusFilters.join(' or ')})`;
        }

        // Build search text filter (sanitized for OData substringof)
        if (filters.searchText) {
          const sanitizedSearch = ValidationUtils.sanitizeForOData(filters.searchText.substring(0, 100));
          filterQuery += ` and substringof('${sanitizedSearch}', File/Name)`;
        }
      }

      const items = await this.sp.web.lists
        .getByTitle(this.DOCUMENT_LIBRARY)
        .items.select(
          '*',
          'File/Name',
          'File/ServerRelativeUrl',
          'File/Length',
          'File/TimeLastModified',
          'UploadedBy/Title',
          'UploadedBy/EMail',
          'Author/Title',
          'Author/EMail'
        )
        .expand('File', 'UploadedBy', 'Author')
        .filter(filterQuery)
        .orderBy('UploadedDate', false)();

      const documents: IJmlDocument[] = [];

      for (let i = 0; i < items.length; i++) {
        documents.push(this.mapToDocument(items[i]));
      }

      return documents;
    } catch (error) {
      logger.error('DocumentService', 'Failed to get process documents:', error);
      return [];
    }
  }

  /**
   * Delete a document
   */
  public async deleteDocument(documentId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.DOCUMENT_LIBRARY)
        .items.getById(documentId)
        .delete();
    } catch (error) {
      logger.error('DocumentService', 'Failed to delete document:', error);
      throw error;
    }
  }

  /**
   * Check out document
   */
  public async checkOutDocument(documentId: number): Promise<void> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.DOCUMENT_LIBRARY)
        .items.getById(documentId)
        .select('File/ServerRelativeUrl')
        .expand('File')();

      await this.sp.web.getFileByServerRelativePath(item.File.ServerRelativeUrl).checkout();
    } catch (error) {
      logger.error('DocumentService', 'Failed to check out document:', error);
      throw error;
    }
  }

  /**
   * Check in document
   */
  public async checkInDocument(documentId: number, comment?: string): Promise<void> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.DOCUMENT_LIBRARY)
        .items.getById(documentId)
        .select('File/ServerRelativeUrl')
        .expand('File')();

      await this.sp.web.getFileByServerRelativePath(item.File.ServerRelativeUrl)
        .checkin(comment || '', 1); // 1 = Major version
    } catch (error) {
      logger.error('DocumentService', 'Failed to check in document:', error);
      throw error;
    }
  }

  /**
   * Get document version history
   */
  public async getVersionHistory(documentId: number): Promise<any[]> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.DOCUMENT_LIBRARY)
        .items.getById(documentId)
        .select('File/ServerRelativeUrl')
        .expand('File')();

      const versions = await this.sp.web
        .getFileByServerRelativePath(item.File.ServerRelativeUrl)
        .versions();

      return versions;
    } catch (error) {
      logger.error('DocumentService', 'Failed to get version history:', error);
      return [];
    }
  }

  /**
   * Download document
   */
  public async downloadDocument(documentId: number): Promise<void> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.DOCUMENT_LIBRARY)
        .items.getById(documentId)
        .select('File/Name', 'File/ServerRelativeUrl')
        .expand('File')();

      const fileUrl = `${window.location.origin}${item.File.ServerRelativeUrl}`;

      // Create a temporary link and trigger download
      const link = document.createElement('a');
      link.href = fileUrl;
      link.download = item.File.Name;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (error) {
      logger.error('DocumentService', 'Failed to download document:', error);
      throw error;
    }
  }

  /**
   * Update document metadata
   */
  public async updateDocument(documentId: number, updates: Partial<IJmlDocument>): Promise<void> {
    try {
      const metadata: any = {};

      if (updates.DocumentType !== undefined) {
        metadata.DocumentType = updates.DocumentType;
      }
      if (updates.Description !== undefined) {
        metadata.Description = updates.Description;
      }
      if (updates.Tags !== undefined) {
        metadata.Tags = JSON.stringify(updates.Tags);
      }
      if (updates.ExpirationDate !== undefined) {
        metadata.ExpirationDate = updates.ExpirationDate?.toISOString();
      }
      if (updates.IsConfidential !== undefined) {
        metadata.IsConfidential = updates.IsConfidential;
      }
      if (updates.RequiresSignature !== undefined) {
        metadata.RequiresSignature = updates.RequiresSignature;
      }
      if (updates.SignatureStatus !== undefined) {
        metadata.SignatureStatus = updates.SignatureStatus;
      }

      await this.sp.web.lists
        .getByTitle(this.DOCUMENT_LIBRARY)
        .items.getById(documentId)
        .update(metadata);
    } catch (error) {
      logger.error('DocumentService', 'Failed to update document:', error);
      throw error;
    }
  }

  /**
   * Get document URL
   */
  public getDocumentUrl(documentId: number): Promise<string> {
    return this.sp.web.lists
      .getByTitle(this.DOCUMENT_LIBRARY)
      .items.getById(documentId)
      .select('File/ServerRelativeUrl')
      .expand('File')()
      .then((item: any) => `${window.location.origin}${item.File.ServerRelativeUrl}`);
  }

  /**
   * Map SharePoint item to IJmlDocument
   */
  private mapToDocument(item: any): IJmlDocument {
    return {
      Id: item.Id,
      ProcessID: item.ProcessID,
      FileName: item.File?.Name || item.Title,
      FileUrl: item.File?.ServerRelativeUrl || '',
      FileSize: item.File?.Length || 0,
      ContentType: item.File?.ContentType || '',
      DocumentType: item.DocumentType as DocumentType,
      UploadedBy: {
        Id: item.UploadedById || item.AuthorId,
        Title: item.UploadedBy?.Title || item.Author?.Title || '',
        EMail: item.UploadedBy?.EMail || item.Author?.EMail || ''
      },
      UploadedById: item.UploadedById || item.AuthorId,
      UploadedDate: new Date(item.UploadedDate || item.Created),
      RequiresSignature: item.RequiresSignature || false,
      SignatureStatus: (item.SignatureStatus as SignatureStatus) || SignatureStatus.NotRequired,
      SignatureProvider: item.SignatureProvider,
      SignatureEnvelopeId: item.SignatureEnvelopeId,
      SignedBy: item.SignedBy ? {
        Id: item.SignedById,
        Title: item.SignedBy.Title,
        EMail: item.SignedBy.EMail
      } : undefined,
      SignedById: item.SignedById,
      SignedDate: item.SignedDate ? new Date(item.SignedDate) : undefined,
      Version: item.Version || '1.0',
      IsCurrentVersion: item.IsCurrentVersion !== false,
      CheckedOut: item.CheckedOut || false,
      CheckedOutBy: item.CheckedOutBy ? {
        Id: item.CheckedOutById,
        Title: item.CheckedOutBy.Title,
        EMail: item.CheckedOutBy.EMail
      } : undefined,
      CheckedOutById: item.CheckedOutById,
      CheckedOutDate: item.CheckedOutDate ? new Date(item.CheckedOutDate) : undefined,
      Description: item.Description,
      Tags: item.Tags ? JSON.parse(item.Tags) : undefined,
      ExpirationDate: item.ExpirationDate ? new Date(item.ExpirationDate) : undefined,
      IsConfidential: item.IsConfidential || false,
      Modified: new Date(item.Modified),
      ModifiedBy: {
        Id: item.EditorId,
        Title: item.Editor?.Title || '',
        EMail: item.Editor?.EMail || ''
      },
      ModifiedById: item.EditorId
    };
  }
}
