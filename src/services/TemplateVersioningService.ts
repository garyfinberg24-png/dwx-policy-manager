// @ts-nocheck
/**
 * Template Versioning Service
 * Manages version history for document templates
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import { IJmlDocumentTemplate, IUser } from '../models';
import { logger } from './LoggingService';

/**
 * Template version information
 */
export interface ITemplateVersion {
  /** Version ID */
  id: number;
  /** Version number (e.g., "1.0", "2.0") */
  versionNumber: string;
  /** Version label (e.g., "Draft", "Published", "Archived") */
  versionLabel: string;
  /** When this version was created */
  created: Date;
  /** Who created this version */
  createdBy: IUser;
  /** Size of the template file in bytes */
  size: number;
  /** Comments/notes for this version */
  comments?: string;
  /** Whether this is the current/active version */
  isCurrentVersion: boolean;
  /** URL to download this version */
  downloadUrl: string;
}

/**
 * Version comparison result
 */
export interface IVersionComparison {
  /** The two versions being compared */
  versions: [ITemplateVersion, ITemplateVersion];
  /** Fields that changed */
  changes: IVersionChange[];
  /** Summary of changes */
  summary: string;
}

/**
 * Individual version change
 */
export interface IVersionChange {
  field: string;
  oldValue: string;
  newValue: string;
  changeType: 'added' | 'removed' | 'modified';
}

/**
 * Options for creating a new version
 */
export interface ICreateVersionOptions {
  /** Comments for the new version */
  comments?: string;
  /** Whether to make this the current version */
  makeCurrent?: boolean;
  /** Version label */
  label?: 'Draft' | 'Published' | 'Archived';
}

/**
 * Template Versioning Service
 */
export class TemplateVersioningService {
  private sp: SPFI;
  private readonly templateLibrary: string;
  private readonly versionHistoryList: string;

  constructor(sp: SPFI, templateLibraryUrl?: string) {
    this.sp = sp;
    this.templateLibrary = templateLibraryUrl || 'PM_DocumentTemplates';
    this.versionHistoryList = 'PM_TemplateVersionHistory';
  }

  /**
   * Get version history for a template
   * @param templateId Template ID
   */
  public async getVersionHistory(templateId: number): Promise<ITemplateVersion[]> {
    try {
      // Get the template file
      const template = await this.sp.web.lists
        .getByTitle(this.templateLibrary)
        .items.getById(templateId)
        .select('File/ServerRelativeUrl')
        .expand('File')();

      if (!template.File?.ServerRelativeUrl) {
        throw new Error('Template file not found');
      }

      // Get file versions from SharePoint
      const fileVersions = await this.sp.web
        .getFileByServerRelativePath(template.File.ServerRelativeUrl)
        .versions();

      // Also get the current version
      const currentFile = await this.sp.web
        .getFileByServerRelativePath(template.File.ServerRelativeUrl)
        .select('*', 'ModifiedBy/Title', 'ModifiedBy/EMail')
        .expand('ModifiedBy')() as unknown as {
          UIVersionLabel?: string;
          TimeLastModified: string;
          ModifiedBy?: { Id?: number; Title?: string; EMail?: string };
          Length?: number;
          CheckInComment?: string;
          ServerRelativeUrl: string;
        };

      const versions: ITemplateVersion[] = [];

      // Add current version first
      versions.push({
        id: 0, // Current version
        versionNumber: currentFile.UIVersionLabel || '1.0',
        versionLabel: 'Current',
        created: new Date(currentFile.TimeLastModified),
        createdBy: {
          Id: currentFile.ModifiedBy?.Id || 0,
          Title: currentFile.ModifiedBy?.Title || 'Unknown',
          EMail: currentFile.ModifiedBy?.EMail || ''
        },
        size: Number(currentFile.Length) || 0,
        comments: currentFile.CheckInComment || undefined,
        isCurrentVersion: true,
        downloadUrl: currentFile.ServerRelativeUrl
      });

      // Add historical versions
      for (let i = 0; i < fileVersions.length; i++) {
        const version = fileVersions[i];
        versions.push({
          id: version.ID,
          versionNumber: version.VersionLabel || String(version.ID),
          versionLabel: 'Historical',
          created: new Date(version.Created),
          createdBy: {
            Id: 0,
            Title: version.CreatedBy?.Name || 'Unknown',
            EMail: version.CreatedBy?.Email || ''
          },
          size: version.Size || 0,
          comments: version.CheckInComment || undefined,
          isCurrentVersion: false,
          downloadUrl: version.Url || ''
        });
      }

      // Sort by version number descending
      versions.sort((a, b) => {
        const vA = parseFloat(a.versionNumber) || 0;
        const vB = parseFloat(b.versionNumber) || 0;
        return vB - vA;
      });

      logger.info('TemplateVersioningService', `Retrieved ${versions.length} versions for template ${templateId}`);
      return versions;
    } catch (error) {
      logger.error('TemplateVersioningService', 'Failed to get version history:', error);
      // Return empty array for templates without version history
      return [];
    }
  }

  /**
   * Create a new version of a template
   * @param templateId Template ID
   * @param newFile The new template file
   * @param options Version options
   */
  public async createVersion(
    templateId: number,
    newFile: File,
    options?: ICreateVersionOptions
  ): Promise<ITemplateVersion> {
    try {
      // Get the current template
      const template = await this.sp.web.lists
        .getByTitle(this.templateLibrary)
        .items.getById(templateId)
        .select('File/Name', 'File/ServerRelativeUrl')
        .expand('File')();

      if (!template.File?.ServerRelativeUrl) {
        throw new Error('Template file not found');
      }

      // Check out the file if required
      const file = this.sp.web.getFileByServerRelativePath(template.File.ServerRelativeUrl);

      try {
        await file.checkout();
      } catch {
        // File might not require checkout or already checked out
        logger.debug('TemplateVersioningService', 'Checkout not required or already checked out');
      }

      // Upload new version
      const folder = template.File.ServerRelativeUrl.substring(
        0,
        template.File.ServerRelativeUrl.lastIndexOf('/')
      );

      await this.sp.web
        .getFolderByServerRelativePath(folder)
        .files.addUsingPath(template.File.Name, newFile, { Overwrite: true });

      // Check in with comments
      const checkInComment = options?.comments || 'New version uploaded';
      await file.checkin(checkInComment, 1); // Major version

      // Record version in history list (for additional metadata)
      await this.recordVersionHistory(templateId, {
        comments: options?.comments,
        label: options?.label
      });

      // Get the new version info
      const versions = await this.getVersionHistory(templateId);
      const newVersion = versions[0]; // Most recent version

      logger.info('TemplateVersioningService', `Created new version for template ${templateId}`);
      return newVersion;
    } catch (error) {
      logger.error('TemplateVersioningService', 'Failed to create version:', error);
      throw error;
    }
  }

  /**
   * Restore a previous version
   * @param templateId Template ID
   * @param versionId Version ID to restore
   */
  public async restoreVersion(templateId: number, versionId: number): Promise<void> {
    try {
      // Get the template file
      const template = await this.sp.web.lists
        .getByTitle(this.templateLibrary)
        .items.getById(templateId)
        .select('File/ServerRelativeUrl')
        .expand('File')();

      if (!template.File?.ServerRelativeUrl) {
        throw new Error('Template file not found');
      }

      // Restore the version
      await this.sp.web
        .getFileByServerRelativePath(template.File.ServerRelativeUrl)
        .versions.restoreByLabel(String(versionId));

      // Record the restoration in history
      await this.recordVersionHistory(templateId, {
        comments: `Restored from version ${versionId}`,
        label: 'Draft'
      });

      logger.info('TemplateVersioningService', `Restored version ${versionId} for template ${templateId}`);
    } catch (error) {
      logger.error('TemplateVersioningService', 'Failed to restore version:', error);
      throw error;
    }
  }

  /**
   * Delete a specific version
   * @param templateId Template ID
   * @param versionId Version ID to delete
   */
  public async deleteVersion(templateId: number, versionId: number): Promise<void> {
    try {
      // Get the template file
      const template = await this.sp.web.lists
        .getByTitle(this.templateLibrary)
        .items.getById(templateId)
        .select('File/ServerRelativeUrl')
        .expand('File')();

      if (!template.File?.ServerRelativeUrl) {
        throw new Error('Template file not found');
      }

      // Delete the version
      await this.sp.web
        .getFileByServerRelativePath(template.File.ServerRelativeUrl)
        .versions.deleteById(versionId);

      logger.info('TemplateVersioningService', `Deleted version ${versionId} for template ${templateId}`);
    } catch (error) {
      logger.error('TemplateVersioningService', 'Failed to delete version:', error);
      throw error;
    }
  }

  /**
   * Download a specific version
   * @param templateId Template ID
   * @param versionId Version ID (0 for current)
   */
  public async downloadVersion(templateId: number, versionId: number): Promise<Blob> {
    try {
      // Get the template file
      const template = await this.sp.web.lists
        .getByTitle(this.templateLibrary)
        .items.getById(templateId)
        .select('File/ServerRelativeUrl', 'File/Name')
        .expand('File')();

      if (!template.File?.ServerRelativeUrl) {
        throw new Error('Template file not found');
      }

      let fileBlob: Blob;

      if (versionId === 0) {
        // Get current version
        fileBlob = await this.sp.web
          .getFileByServerRelativePath(template.File.ServerRelativeUrl)
          .getBlob();
      } else {
        // Get specific version
        const versions = await this.sp.web
          .getFileByServerRelativePath(template.File.ServerRelativeUrl)
          .versions();

        const version = versions.find(v => v.ID === versionId);
        if (!version) {
          throw new Error(`Version ${versionId} not found`);
        }

        // Get version content
        const versionUrl = version.Url;
        const response = await fetch(versionUrl);
        fileBlob = await response.blob();
      }

      logger.info('TemplateVersioningService', `Downloaded version ${versionId} for template ${templateId}`);
      return fileBlob;
    } catch (error) {
      logger.error('TemplateVersioningService', 'Failed to download version:', error);
      throw error;
    }
  }

  /**
   * Compare two versions
   * @param templateId Template ID
   * @param versionId1 First version ID
   * @param versionId2 Second version ID
   */
  public async compareVersions(
    templateId: number,
    versionId1: number,
    versionId2: number
  ): Promise<IVersionComparison> {
    try {
      const versions = await this.getVersionHistory(templateId);

      const v1 = versions.find(v => v.id === versionId1);
      const v2 = versions.find(v => v.id === versionId2);

      if (!v1 || !v2) {
        throw new Error('One or both versions not found');
      }

      // For now, compare basic metadata
      // Full content comparison would require more sophisticated tooling
      const changes: IVersionChange[] = [];

      if (v1.size !== v2.size) {
        changes.push({
          field: 'File Size',
          oldValue: `${v1.size} bytes`,
          newValue: `${v2.size} bytes`,
          changeType: 'modified'
        });
      }

      if (v1.createdBy.Title !== v2.createdBy.Title) {
        changes.push({
          field: 'Modified By',
          oldValue: v1.createdBy.Title,
          newValue: v2.createdBy.Title,
          changeType: 'modified'
        });
      }

      const summary = changes.length > 0
        ? `${changes.length} change(s) detected between versions ${v1.versionNumber} and ${v2.versionNumber}`
        : `No metadata changes between versions ${v1.versionNumber} and ${v2.versionNumber}`;

      return {
        versions: [v1, v2],
        changes,
        summary
      };
    } catch (error) {
      logger.error('TemplateVersioningService', 'Failed to compare versions:', error);
      throw error;
    }
  }

  /**
   * Get the latest version number
   * @param templateId Template ID
   */
  public async getLatestVersionNumber(templateId: number): Promise<string> {
    try {
      const versions = await this.getVersionHistory(templateId);
      return versions.length > 0 ? versions[0].versionNumber : '1.0';
    } catch {
      return '1.0';
    }
  }

  /**
   * Record version history in a separate list for additional metadata
   */
  private async recordVersionHistory(
    templateId: number,
    options: { comments?: string; label?: string }
  ): Promise<void> {
    try {
      // Try to add to version history list
      await this.sp.web.lists
        .getByTitle(this.versionHistoryList)
        .items.add({
          Title: `Template ${templateId} - ${new Date().toISOString()}`,
          TemplateId: templateId,
          VersionLabel: options.label || 'Draft',
          Comments: options.comments || '',
          Created: new Date()
        });
    } catch {
      // List might not exist - that's okay, SharePoint handles versions natively
      logger.debug('TemplateVersioningService', 'Version history list not available, using SharePoint native versioning only');
    }
  }
}

// Export singleton factory
export function createTemplateVersioningService(sp: SPFI, templateLibraryUrl?: string): TemplateVersioningService {
  return new TemplateVersioningService(sp, templateLibraryUrl);
}
