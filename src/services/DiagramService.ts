// @ts-nocheck
/**
 * Diagram Service
 * Handles diagram upload, storage, and management in SharePoint
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import {
  IDiagram,
  IDiagramUploadOptions,
  IDiagramSearchFilters,
  IDiagramAnnotations,
  DiagramType,
  DiagramStatus
} from '../models/IDiagram';
import { logger } from './LoggingService';
import { CacheService, CacheDurations } from './CacheService';

export class DiagramService {
  private sp: SPFI;
  private cache: CacheService;
  private readonly DIAGRAM_LIST = 'PM_Diagrams';
  private readonly DIAGRAM_LIBRARY = 'PM_DiagramImages';

  constructor(sp: SPFI) {
    this.sp = sp;
    this.cache = CacheService.getInstance();
  }

  /**
   * Upload a new diagram with base image
   */
  public async uploadDiagram(
    imageFile: File,
    options: IDiagramUploadOptions
  ): Promise<IDiagram> {
    try {
      logger.info('DiagramService', `Uploading diagram: ${options.title}`);

      // Upload the image file to the library
      const imageUrl = await this.uploadImageFile(imageFile, options.title);

      // Create thumbnail URL (SharePoint auto-generates)
      const thumbnailUrl = this.generateThumbnailUrl(imageUrl);

      // Create the diagram list item
      const itemData: any = {
        Title: options.title,
        DiagramType: options.diagramType,
        Status: DiagramStatus.Draft,
        Description: options.description || '',
        Building: options.building || '',
        Floor: options.floor || '',
        Department: options.department || '',
        ImageUrl: imageUrl,
        ThumbnailUrl: thumbnailUrl,
        AnnotationsJson: JSON.stringify({ markers: [], shapes: [], routes: [], zones: [] }),
        Version: '1.0',
        Tags: options.tags ? JSON.stringify(options.tags) : '[]',
        EffectiveDate: options.effectiveDate?.toISOString(),
        ExpirationDate: options.expirationDate?.toISOString(),
        IsActive: false
      };

      const result = await this.sp.web.lists
        .getByTitle(this.DIAGRAM_LIST)
        .items.add(itemData);

      // Invalidate cache
      this.cache.invalidatePattern('diagrams:');

      return await this.getDiagramById(result.data.Id);
    } catch (error) {
      logger.error('DiagramService', 'Failed to upload diagram:', error);
      throw error;
    }
  }

  /**
   * Upload image file to SharePoint library
   */
  private async uploadImageFile(file: File, title: string): Promise<string> {
    try {
      // Sanitise filename
      const timestamp = Date.now();
      const extension = file.name.split('.').pop() || 'png';
      const sanitisedTitle = title.replace(/[^a-zA-Z0-9]/g, '_');
      const fileName = `${sanitisedTitle}_${timestamp}.${extension}`;

      // Upload to library
      const result = await this.sp.web.lists
        .getByTitle(this.DIAGRAM_LIBRARY)
        .rootFolder
        .files.addUsingPath(fileName, file, { Overwrite: true });

      return result.data.ServerRelativeUrl;
    } catch (error) {
      logger.error('DiagramService', 'Failed to upload image file:', error);
      throw error;
    }
  }

  /**
   * Generate thumbnail URL from image URL
   */
  private generateThumbnailUrl(imageUrl: string): string {
    // SharePoint auto-generates thumbnails via REST API
    return `${imageUrl}?width=300&height=300`;
  }

  /**
   * Get diagram by ID
   */
  public async getDiagramById(diagramId: number): Promise<IDiagram> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.DIAGRAM_LIST)
        .items.getById(diagramId)
        .select(
          '*',
          'Author/Id',
          'Author/Title',
          'Author/EMail',
          'Editor/Id',
          'Editor/Title',
          'Editor/EMail',
          'ApprovedBy/Id',
          'ApprovedBy/Title',
          'ApprovedBy/EMail'
        )
        .expand('Author', 'Editor', 'ApprovedBy')();

      return this.mapToDiagram(item);
    } catch (error) {
      logger.error('DiagramService', `Failed to get diagram ${diagramId}:`, error);
      throw error;
    }
  }

  /**
   * Get all diagrams with optional filters
   */
  public async getDiagrams(filters?: IDiagramSearchFilters): Promise<IDiagram[]> {
    try {
      const cacheKey = `diagrams:list:${JSON.stringify(filters || {})}`;
      const cached = this.cache.get<IDiagram[]>(cacheKey);
      if (cached) {
        return cached;
      }

      let filterQuery = '';
      const filterParts: string[] = [];

      if (filters) {
        // Search text
        if (filters.searchText) {
          const searchText = filters.searchText.replace(/'/g, "''");
          filterParts.push(`(substringof('${searchText}', Title) or substringof('${searchText}', Description))`);
        }

        // Diagram types
        if (filters.diagramTypes && filters.diagramTypes.length > 0) {
          const typeFilters = filters.diagramTypes.map(t => `DiagramType eq '${t}'`);
          filterParts.push(`(${typeFilters.join(' or ')})`);
        }

        // Statuses
        if (filters.statuses && filters.statuses.length > 0) {
          const statusFilters = filters.statuses.map(s => `Status eq '${s}'`);
          filterParts.push(`(${statusFilters.join(' or ')})`);
        }

        // Buildings
        if (filters.buildings && filters.buildings.length > 0) {
          const buildingFilters = filters.buildings.map(b => `Building eq '${b.replace(/'/g, "''")}'`);
          filterParts.push(`(${buildingFilters.join(' or ')})`);
        }

        // Floors
        if (filters.floors && filters.floors.length > 0) {
          const floorFilters = filters.floors.map(f => `Floor eq '${f.replace(/'/g, "''")}'`);
          filterParts.push(`(${floorFilters.join(' or ')})`);
        }

        // Departments
        if (filters.departments && filters.departments.length > 0) {
          const deptFilters = filters.departments.map(d => `Department eq '${d.replace(/'/g, "''")}'`);
          filterParts.push(`(${deptFilters.join(' or ')})`);
        }

        // Active only
        if (filters.isActive !== undefined) {
          filterParts.push(`IsActive eq ${filters.isActive}`);
        }
      }

      filterQuery = filterParts.join(' and ');

      let query = this.sp.web.lists
        .getByTitle(this.DIAGRAM_LIST)
        .items.select(
          '*',
          'Author/Id',
          'Author/Title',
          'Author/EMail',
          'Editor/Id',
          'Editor/Title',
          'Editor/EMail'
        )
        .expand('Author', 'Editor')
        .orderBy('Modified', false)
        .top(500);

      if (filterQuery) {
        query = query.filter(filterQuery);
      }

      const items = await query();
      const diagrams = items.map((item: any) => this.mapToDiagram(item));

      this.cache.set(cacheKey, diagrams, CacheDurations.MEDIUM);

      return diagrams;
    } catch (error) {
      logger.error('DiagramService', 'Failed to get diagrams:', error);
      return [];
    }
  }

  /**
   * Get diagrams by type
   */
  public async getDiagramsByType(diagramType: DiagramType): Promise<IDiagram[]> {
    return this.getDiagrams({ diagramTypes: [diagramType] });
  }

  /**
   * Get active published diagrams
   */
  public async getPublishedDiagrams(filters?: IDiagramSearchFilters): Promise<IDiagram[]> {
    return this.getDiagrams({
      ...filters,
      statuses: [DiagramStatus.Published],
      isActive: true
    });
  }

  /**
   * Update diagram metadata
   */
  public async updateDiagram(diagramId: number, updates: Partial<IDiagram>): Promise<IDiagram> {
    try {
      const updateData: any = {};

      if (updates.Title !== undefined) updateData.Title = updates.Title;
      if (updates.DiagramType !== undefined) updateData.DiagramType = updates.DiagramType;
      if (updates.Description !== undefined) updateData.Description = updates.Description;
      if (updates.Building !== undefined) updateData.Building = updates.Building;
      if (updates.Floor !== undefined) updateData.Floor = updates.Floor;
      if (updates.Department !== undefined) updateData.Department = updates.Department;
      if (updates.Tags !== undefined) updateData.Tags = JSON.stringify(updates.Tags);
      if (updates.EffectiveDate !== undefined) updateData.EffectiveDate = updates.EffectiveDate?.toISOString();
      if (updates.ExpirationDate !== undefined) updateData.ExpirationDate = updates.ExpirationDate?.toISOString();
      if (updates.IsActive !== undefined) updateData.IsActive = updates.IsActive;

      await this.sp.web.lists
        .getByTitle(this.DIAGRAM_LIST)
        .items.getById(diagramId)
        .update(updateData);

      this.cache.invalidatePattern('diagrams:');

      return await this.getDiagramById(diagramId);
    } catch (error) {
      logger.error('DiagramService', `Failed to update diagram ${diagramId}:`, error);
      throw error;
    }
  }

  /**
   * Save diagram annotations
   */
  public async saveAnnotations(diagramId: number, annotations: IDiagramAnnotations): Promise<IDiagram> {
    try {
      logger.info('DiagramService', `Saving annotations for diagram ${diagramId}`);

      const annotationsJson = JSON.stringify(annotations);

      await this.sp.web.lists
        .getByTitle(this.DIAGRAM_LIST)
        .items.getById(diagramId)
        .update({
          AnnotationsJson: annotationsJson
        });

      this.cache.invalidatePattern('diagrams:');

      return await this.getDiagramById(diagramId);
    } catch (error) {
      logger.error('DiagramService', `Failed to save annotations for diagram ${diagramId}:`, error);
      throw error;
    }
  }

  /**
   * Update diagram status
   */
  public async updateStatus(diagramId: number, status: DiagramStatus): Promise<IDiagram> {
    try {
      const updateData: any = { Status: status };

      // If publishing, set IsActive
      if (status === DiagramStatus.Published) {
        updateData.IsActive = true;
      }

      // If archiving, set IsActive to false
      if (status === DiagramStatus.Archived) {
        updateData.IsActive = false;
      }

      await this.sp.web.lists
        .getByTitle(this.DIAGRAM_LIST)
        .items.getById(diagramId)
        .update(updateData);

      this.cache.invalidatePattern('diagrams:');

      return await this.getDiagramById(diagramId);
    } catch (error) {
      logger.error('DiagramService', `Failed to update diagram status ${diagramId}:`, error);
      throw error;
    }
  }

  /**
   * Approve diagram
   */
  public async approveDiagram(diagramId: number): Promise<IDiagram> {
    try {
      await this.sp.web.lists
        .getByTitle(this.DIAGRAM_LIST)
        .items.getById(diagramId)
        .update({
          Status: DiagramStatus.Approved,
          ApprovedDate: new Date().toISOString()
        });

      this.cache.invalidatePattern('diagrams:');

      return await this.getDiagramById(diagramId);
    } catch (error) {
      logger.error('DiagramService', `Failed to approve diagram ${diagramId}:`, error);
      throw error;
    }
  }

  /**
   * Publish diagram (make it active)
   */
  public async publishDiagram(diagramId: number): Promise<IDiagram> {
    try {
      // First, deactivate any other active diagrams of the same type/building/floor
      const diagram = await this.getDiagramById(diagramId);

      if (diagram.Building && diagram.Floor) {
        const existingActive = await this.getDiagrams({
          diagramTypes: [diagram.DiagramType],
          buildings: [diagram.Building],
          floors: [diagram.Floor],
          isActive: true
        });

        // Deactivate existing active diagrams
        for (const existing of existingActive) {
          if (existing.Id !== diagramId) {
            await this.sp.web.lists
              .getByTitle(this.DIAGRAM_LIST)
              .items.getById(existing.Id!)
              .update({ IsActive: false });
          }
        }
      }

      // Publish and activate the diagram
      await this.sp.web.lists
        .getByTitle(this.DIAGRAM_LIST)
        .items.getById(diagramId)
        .update({
          Status: DiagramStatus.Published,
          IsActive: true
        });

      this.cache.invalidatePattern('diagrams:');

      return await this.getDiagramById(diagramId);
    } catch (error) {
      logger.error('DiagramService', `Failed to publish diagram ${diagramId}:`, error);
      throw error;
    }
  }

  /**
   * Delete diagram
   */
  public async deleteDiagram(diagramId: number): Promise<void> {
    try {
      // Get the diagram to find the image URL
      const diagram = await this.getDiagramById(diagramId);

      // Delete the list item
      await this.sp.web.lists
        .getByTitle(this.DIAGRAM_LIST)
        .items.getById(diagramId)
        .delete();

      // Try to delete the image file
      if (diagram.ImageUrl) {
        try {
          await this.sp.web.getFileByServerRelativePath(diagram.ImageUrl).delete();
        } catch (fileError) {
          logger.warn('DiagramService', `Could not delete image file: ${diagram.ImageUrl}`, fileError);
        }
      }

      this.cache.invalidatePattern('diagrams:');
    } catch (error) {
      logger.error('DiagramService', `Failed to delete diagram ${diagramId}:`, error);
      throw error;
    }
  }

  /**
   * Replace diagram image
   */
  public async replaceImage(diagramId: number, newImageFile: File): Promise<IDiagram> {
    try {
      const diagram = await this.getDiagramById(diagramId);

      // Upload new image
      const newImageUrl = await this.uploadImageFile(newImageFile, diagram.Title);
      const thumbnailUrl = this.generateThumbnailUrl(newImageUrl);

      // Update diagram with new image URL
      await this.sp.web.lists
        .getByTitle(this.DIAGRAM_LIST)
        .items.getById(diagramId)
        .update({
          ImageUrl: newImageUrl,
          ThumbnailUrl: thumbnailUrl
        });

      // Delete old image file
      if (diagram.ImageUrl && diagram.ImageUrl !== newImageUrl) {
        try {
          await this.sp.web.getFileByServerRelativePath(diagram.ImageUrl).delete();
        } catch (fileError) {
          logger.warn('DiagramService', `Could not delete old image file: ${diagram.ImageUrl}`, fileError);
        }
      }

      this.cache.invalidatePattern('diagrams:');

      return await this.getDiagramById(diagramId);
    } catch (error) {
      logger.error('DiagramService', `Failed to replace image for diagram ${diagramId}:`, error);
      throw error;
    }
  }

  /**
   * Get unique buildings for filtering
   */
  public async getBuildings(): Promise<string[]> {
    try {
      const cacheKey = 'diagrams:buildings';
      const cached = this.cache.get<string[]>(cacheKey);
      if (cached) return cached;

      const items = await this.sp.web.lists
        .getByTitle(this.DIAGRAM_LIST)
        .items.select('Building')
        .filter("Building ne null and Building ne ''")
        .top(1000)();

      const buildings = Array.from(new Set(items.map((i: any) => i.Building as string))).sort();

      this.cache.set(cacheKey, buildings, CacheDurations.LONG);

      return buildings;
    } catch (error) {
      logger.error('DiagramService', 'Failed to get buildings:', error);
      return [];
    }
  }

  /**
   * Get unique floors for a building
   */
  public async getFloors(building?: string): Promise<string[]> {
    try {
      const cacheKey = `diagrams:floors:${building || 'all'}`;
      const cached = this.cache.get<string[]>(cacheKey);
      if (cached) return cached;

      let query = this.sp.web.lists
        .getByTitle(this.DIAGRAM_LIST)
        .items.select('Floor')
        .filter("Floor ne null and Floor ne ''")
        .top(1000);

      if (building) {
        query = query.filter(`Building eq '${building.replace(/'/g, "''")}'`);
      }

      const items = await query();

      const floors = Array.from(new Set(items.map((i: any) => i.Floor as string))).sort();

      this.cache.set(cacheKey, floors, CacheDurations.LONG);

      return floors;
    } catch (error) {
      logger.error('DiagramService', 'Failed to get floors:', error);
      return [];
    }
  }

  /**
   * Map SharePoint item to IDiagram
   */
  private mapToDiagram(item: any): IDiagram {
    let annotations: IDiagramAnnotations | undefined;
    let tags: string[] | undefined;

    try {
      if (item.AnnotationsJson) {
        annotations = JSON.parse(item.AnnotationsJson);
      }
    } catch {
      annotations = { markers: [], shapes: [], routes: [], zones: [] };
    }

    try {
      if (item.Tags) {
        tags = JSON.parse(item.Tags);
      }
    } catch {
      tags = [];
    }

    return {
      Id: item.Id,
      Title: item.Title,
      DiagramType: item.DiagramType as DiagramType,
      Status: item.Status as DiagramStatus,
      Description: item.Description,
      Building: item.Building,
      Floor: item.Floor,
      Department: item.Department,
      ImageUrl: item.ImageUrl,
      ThumbnailUrl: item.ThumbnailUrl,
      AnnotationsJson: item.AnnotationsJson,
      annotations,
      Version: item.Version || '1.0',
      Tags: tags,
      EffectiveDate: item.EffectiveDate ? new Date(item.EffectiveDate) : undefined,
      ExpirationDate: item.ExpirationDate ? new Date(item.ExpirationDate) : undefined,
      IsActive: item.IsActive || false,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined,
      CreatedBy: item.Author ? {
        Id: item.Author.Id,
        Title: item.Author.Title,
        EMail: item.Author.EMail
      } : undefined,
      CreatedById: item.AuthorId,
      ModifiedBy: item.Editor ? {
        Id: item.Editor.Id,
        Title: item.Editor.Title,
        EMail: item.Editor.EMail
      } : undefined,
      ModifiedById: item.EditorId,
      ApprovedBy: item.ApprovedBy ? {
        Id: item.ApprovedBy.Id,
        Title: item.ApprovedBy.Title,
        EMail: item.ApprovedBy.EMail
      } : undefined,
      ApprovedById: item.ApprovedById,
      ApprovedDate: item.ApprovedDate ? new Date(item.ApprovedDate) : undefined
    };
  }
}
