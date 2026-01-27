// @ts-nocheck
// TaskAttachmentsService - Handles file attachments for tasks
// Provides upload, download, delete, and list functionality

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import {
  IJmlTaskAttachment,
  ITaskAttachmentView,
  ITaskAttachmentUpload,
  ITaskAttachmentSummary,
  AttachmentCategory
} from '../models';
import { logger } from './LoggingService';

export class TaskAttachmentsService {
  private sp: SPFI;
  private libraryTitle = 'JML_TaskAttachments';
  private listExists: boolean | null = null;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Check if the attachments library exists
   */
  private async checkListExists(): Promise<boolean> {
    if (this.listExists !== null) {
      return this.listExists;
    }

    try {
      await this.sp.web.lists.getByTitle(this.libraryTitle).select('Id')();
      this.listExists = true;
      return true;
    } catch {
      logger.warn('TaskAttachmentsService', `Library '${this.libraryTitle}' does not exist. Attachments feature will be unavailable.`);
      this.listExists = false;
      return false;
    }
  }

  /**
   * Get all attachments for a task
   */
  public async getTaskAttachments(taskAssignmentId: number, currentUserId: number): Promise<ITaskAttachmentView[]> {
    try {
      // Check if library exists first
      const exists = await this.checkListExists();
      if (!exists) {
        logger.info('TaskAttachmentsService', `Attachments library not found - returning empty array for task ${taskAssignmentId}`);
        return [];
      }

      logger.info('TaskAttachmentsService', `Fetching attachments for task ${taskAssignmentId}`);

      const items = await this.sp.web.lists
        .getByTitle(this.libraryTitle)
        .items.filter(`TaskAssignmentId eq ${taskAssignmentId}`)
        .select(
          'Id', 'TaskAssignmentId', 'FileName', 'FileUrl', 'FileSize', 'FileType',
          'UploadedById', 'UploadedDate', 'Description', 'Category', 'IsRequired',
          'VersionLabel', 'IsLatestVersion', 'ServerRelativeUrl', 'UniqueId',
          'Created', 'Modified',
          'UploadedBy/Id', 'UploadedBy/Title', 'UploadedBy/EMail',
          'Author/Id', 'Author/Title'
        )
        .expand('UploadedBy', 'Author')
        .orderBy('Created', false)();

      const attachments: ITaskAttachmentView[] = items.map((item: any) => ({
        Id: item.Id,
        Title: item.FileName || 'Attachment',
        TaskAssignmentId: item.TaskAssignmentId,
        FileName: item.FileName,
        FileUrl: item.FileUrl,
        FileSize: item.FileSize,
        FileType: item.FileType,
        UploadedById: item.UploadedById,
        UploadedBy: item.UploadedBy ? {
          Id: item.UploadedBy.Id,
          Title: item.UploadedBy.Title,
          Email: item.UploadedBy.EMail
        } : undefined,
        UploadedDate: new Date(item.UploadedDate || item.Created),
        Description: item.Description,
        Category: item.Category as AttachmentCategory,
        IsRequired: item.IsRequired || false,
        VersionLabel: item.VersionLabel,
        IsLatestVersion: item.IsLatestVersion !== false,
        ServerRelativeUrl: item.ServerRelativeUrl,
        UniqueId: item.UniqueId,
        Created: new Date(item.Created),
        Modified: new Date(item.Modified),
        Author: item.Author ? {
          Id: item.Author.Id,
          Title: item.Author.Title
        } : undefined,
        FormattedSize: this.formatFileSize(item.FileSize),
        IconName: this.getFileIcon(item.FileType),
        CanDelete: item.UploadedById === currentUserId || item.Author?.Id === currentUserId,
        CanDownload: true,
        IsImage: this.isImageFile(item.FileType),
        ThumbnailUrl: this.isImageFile(item.FileType) ? item.FileUrl : undefined
      }));

      logger.info('TaskAttachmentsService', `Retrieved ${attachments.length} attachments`);
      return attachments;
    } catch (error) {
      logger.error('TaskAttachmentsService', 'Error fetching attachments', error);
      throw error;
    }
  }

  /**
   * Upload a new attachment
   */
  public async uploadAttachment(upload: ITaskAttachmentUpload, userId: number): Promise<IJmlTaskAttachment> {
    try {
      logger.info('TaskAttachmentsService', `Uploading file: ${upload.File.name}`);

      // Create folder for task if it doesn't exist
      const folderName = `Task_${upload.TaskAssignmentId}`;
      await this.ensureTaskFolder(folderName);

      // Generate unique filename to avoid conflicts
      const timestamp = new Date().getTime();
      const fileName = `${timestamp}_${upload.File.name}`;
      const folderPath = `${this.libraryTitle}/${folderName}`;

      // Upload file
      const uploadResult = await this.sp.web.getFolderByServerRelativePath(folderPath)
        .files.addUsingPath(fileName, upload.File, { Overwrite: true });

      // Get the uploaded file item
      const file = await uploadResult.file();
      const item = await uploadResult.file.getItem();

      // Update list item with metadata
      await item.update({
        TaskAssignmentId: upload.TaskAssignmentId,
        FileName: upload.File.name,
        FileSize: upload.File.size,
        FileType: this.getFileExtension(upload.File.name),
        UploadedById: userId,
        UploadedDate: new Date().toISOString(),
        Description: upload.Description || '',
        Category: upload.Category || AttachmentCategory.Other,
        IsRequired: upload.IsRequired || false,
        IsLatestVersion: true
      });

      // Refresh item to get server-generated values
      const updatedItem = await this.sp.web.lists
        .getByTitle(this.libraryTitle)
        .items.getById((item as any).Id)
        .select('Id', 'FileName', 'FileSize', 'FileType', 'ServerRelativeUrl', 'UniqueId')();

      logger.info('TaskAttachmentsService', `File uploaded successfully: ${fileName}`);

      return {
        Id: updatedItem.Id,
        Title: upload.File.name,
        TaskAssignmentId: upload.TaskAssignmentId,
        FileName: upload.File.name,
        FileUrl: file.ServerRelativeUrl,
        FileSize: upload.File.size,
        FileType: this.getFileExtension(upload.File.name),
        UploadedById: userId,
        UploadedDate: new Date(),
        Description: upload.Description,
        Category: upload.Category || AttachmentCategory.Other,
        IsRequired: upload.IsRequired || false,
        IsLatestVersion: true,
        ServerRelativeUrl: file.ServerRelativeUrl,
        UniqueId: updatedItem.UniqueId,
        Created: new Date(),
        Modified: new Date()
      };
    } catch (error) {
      logger.error('TaskAttachmentsService', 'Error uploading file', error);
      throw error;
    }
  }

  /**
   * Delete an attachment
   */
  public async deleteAttachment(attachmentId: number): Promise<void> {
    try {
      logger.info('TaskAttachmentsService', `Deleting attachment ${attachmentId}`);

      // Get file URL before deleting list item
      const item = await this.sp.web.lists
        .getByTitle(this.libraryTitle)
        .items.getById(attachmentId)
        .select('ServerRelativeUrl', 'FileUrl')();

      const fileUrl = item.ServerRelativeUrl || item.FileUrl;

      if (fileUrl) {
        // Delete the actual file
        await this.sp.web.getFileByServerRelativePath(fileUrl).delete();
      }

      logger.info('TaskAttachmentsService', `Attachment ${attachmentId} deleted`);
    } catch (error) {
      logger.error('TaskAttachmentsService', 'Error deleting attachment', error);
      throw error;
    }
  }

  /**
   * Download an attachment
   */
  public async downloadAttachment(attachmentId: number): Promise<void> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.libraryTitle)
        .items.getById(attachmentId)
        .select('ServerRelativeUrl', 'FileName')();

      // Open in new window for download
      window.open(item.ServerRelativeUrl, '_blank');
    } catch (error) {
      logger.error('TaskAttachmentsService', 'Error downloading attachment', error);
      throw error;
    }
  }

  /**
   * Get attachment summary for a task
   */
  public async getTaskAttachmentSummary(taskAssignmentId: number): Promise<ITaskAttachmentSummary> {
    try {
      // Check if library exists first
      const exists = await this.checkListExists();
      if (!exists) {
        return {
          TaskAssignmentId: taskAssignmentId,
          TotalCount: 0,
          TotalSize: 0,
          FormattedTotalSize: '0 Bytes',
          RequiredCount: 0,
          ByCategory: {}
        };
      }

      const attachments = await this.sp.web.lists
        .getByTitle(this.libraryTitle)
        .items.filter(`TaskAssignmentId eq ${taskAssignmentId}`)
        .select('Id', 'FileSize', 'Category', 'IsRequired', 'Created')
        .orderBy('Created', false)();

      let totalSize = 0;
      let requiredCount = 0;
      const byCategory: { [key in AttachmentCategory]?: number } = {};

      attachments.forEach((item: any) => {
        totalSize += item.FileSize || 0;
        if (item.IsRequired) requiredCount++;

        const category = item.Category as AttachmentCategory || AttachmentCategory.Other;
        byCategory[category] = (byCategory[category] || 0) + 1;
      });

      return {
        TaskAssignmentId: taskAssignmentId,
        TotalCount: attachments.length,
        TotalSize: totalSize,
        FormattedTotalSize: this.formatFileSize(totalSize),
        RequiredCount: requiredCount,
        ByCategory: byCategory
      };
    } catch (error) {
      logger.error('TaskAttachmentsService', 'Error getting attachment summary', error);
      throw error;
    }
  }

  /**
   * Update attachment metadata
   */
  public async updateAttachment(
    attachmentId: number,
    updates: { Description?: string; Category?: AttachmentCategory; IsRequired?: boolean }
  ): Promise<void> {
    try {
      logger.info('TaskAttachmentsService', `Updating attachment ${attachmentId}`);

      const updateObj: any = {};
      if (updates.Description !== undefined) updateObj.Description = updates.Description;
      if (updates.Category) updateObj.Category = updates.Category;
      if (updates.IsRequired !== undefined) updateObj.IsRequired = updates.IsRequired;

      await this.sp.web.lists
        .getByTitle(this.libraryTitle)
        .items.getById(attachmentId)
        .update(updateObj);

      logger.info('TaskAttachmentsService', `Attachment ${attachmentId} updated`);
    } catch (error) {
      logger.error('TaskAttachmentsService', 'Error updating attachment', error);
      throw error;
    }
  }

  /**
   * Ensure task folder exists
   */
  private async ensureTaskFolder(folderName: string): Promise<void> {
    try {
      // Try to get the folder
      await this.sp.web.lists
        .getByTitle(this.libraryTitle)
        .rootFolder.folders.getByUrl(folderName)();
    } catch {
      // Folder doesn't exist, create it
      try {
        await this.sp.web.lists
          .getByTitle(this.libraryTitle)
          .rootFolder.folders.addUsingPath(folderName);
        logger.info('TaskAttachmentsService', `Created folder: ${folderName}`);
      } catch (error) {
        logger.warn('TaskAttachmentsService', `Folder may already exist: ${folderName}`);
      }
    }
  }

  /**
   * Format file size for display
   */
  private formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + ' ' + sizes[i];
  }

  /**
   * Get file extension
   */
  private getFileExtension(fileName: string): string {
    const lastDot = fileName.lastIndexOf('.');
    return lastDot > 0 ? fileName.substring(lastDot) : '';
  }

  /**
   * Get Fluent UI icon name for file type
   */
  private getFileIcon(fileType: string): string {
    const ext = fileType.toLowerCase();
    const iconMap: { [key: string]: string } = {
      '.pdf': 'PDF',
      '.doc': 'WordDocument',
      '.docx': 'WordDocument',
      '.xls': 'ExcelDocument',
      '.xlsx': 'ExcelDocument',
      '.ppt': 'PowerPointDocument',
      '.pptx': 'PowerPointDocument',
      '.txt': 'TextDocument',
      '.jpg': 'FileImage',
      '.jpeg': 'FileImage',
      '.png': 'FileImage',
      '.gif': 'FileImage',
      '.zip': 'ZipFolder',
      '.rar': 'ZipFolder',
      '.msg': 'Mail',
      '.eml': 'Mail'
    };
    return iconMap[ext] || 'Page';
  }

  /**
   * Check if file is an image
   */
  private isImageFile(fileType: string): boolean {
    const imageExtensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg'];
    return imageExtensions.includes(fileType.toLowerCase());
  }
}
