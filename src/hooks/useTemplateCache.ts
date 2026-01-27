/**
 * Template Caching Hooks
 * Specialized hooks for caching document templates and related data
 */

import * as React from 'react';
import { useDataCache, IUseDataCacheOptions, invalidateCache } from './useDataCache';
import { DocumentTemplateService } from '../services/DocumentTemplateService';
import { IJmlDocumentTemplate, DocumentType } from '../models';
import { SPFI } from '@pnp/sp';

/**
 * Cache key generators for templates
 */
export const TemplateCacheKeys = {
  all: () => 'templates:all',
  byType: (type: DocumentType) => `templates:type:${type}`,
  byId: (id: number) => `templates:id:${id}`,
  forProcess: (processType: string) => `templates:process:${processType}`,
  forProcessAndType: (processType: string, docType: DocumentType) =>
    `templates:process:${processType}:type:${docType}`
};

/**
 * Hook to fetch all templates
 */
export function useTemplates(
  sp: SPFI,
  templateLibraryUrl?: string,
  options?: IUseDataCacheOptions<IJmlDocumentTemplate[]>
): ReturnType<typeof useDataCache<IJmlDocumentTemplate[]>> {
  const service = React.useMemo(
    () => new DocumentTemplateService(sp, templateLibraryUrl),
    [sp, templateLibraryUrl]
  );

  const fetcher = React.useCallback(
    () => service.getTemplates(),
    [service]
  );

  return useDataCache<IJmlDocumentTemplate[]>(
    TemplateCacheKeys.all(),
    fetcher,
    {
      staleTime: 5 * 60 * 1000, // 5 minutes
      cacheTime: 30 * 60 * 1000, // 30 minutes
      ...options
    }
  );
}

/**
 * Hook to fetch templates by document type
 */
export function useTemplatesByType(
  sp: SPFI,
  documentType: DocumentType,
  templateLibraryUrl?: string,
  options?: IUseDataCacheOptions<IJmlDocumentTemplate[]>
): ReturnType<typeof useDataCache<IJmlDocumentTemplate[]>> {
  const service = React.useMemo(
    () => new DocumentTemplateService(sp, templateLibraryUrl),
    [sp, templateLibraryUrl]
  );

  const fetcher = React.useCallback(
    () => service.getTemplates(documentType),
    [service, documentType]
  );

  return useDataCache<IJmlDocumentTemplate[]>(
    TemplateCacheKeys.byType(documentType),
    fetcher,
    {
      staleTime: 5 * 60 * 1000,
      cacheTime: 30 * 60 * 1000,
      ...options
    }
  );
}

/**
 * Hook to fetch a single template by ID
 */
export function useTemplateById(
  sp: SPFI,
  templateId: number | null,
  templateLibraryUrl?: string,
  options?: IUseDataCacheOptions<IJmlDocumentTemplate>
): ReturnType<typeof useDataCache<IJmlDocumentTemplate>> {
  const service = React.useMemo(
    () => new DocumentTemplateService(sp, templateLibraryUrl),
    [sp, templateLibraryUrl]
  );

  const fetcher = React.useCallback(
    () => templateId ? service.getTemplateById(templateId) : Promise.reject(new Error('No template ID')),
    [service, templateId]
  );

  return useDataCache<IJmlDocumentTemplate>(
    templateId ? TemplateCacheKeys.byId(templateId) : null,
    fetcher,
    {
      staleTime: 10 * 60 * 1000, // 10 minutes - individual templates change less frequently
      cacheTime: 60 * 60 * 1000, // 1 hour
      enabled: templateId !== null,
      ...options
    }
  );
}

/**
 * Hook to fetch templates for a process type
 */
export function useTemplatesForProcess(
  sp: SPFI,
  processType: string,
  documentType?: DocumentType,
  templateLibraryUrl?: string,
  options?: IUseDataCacheOptions<IJmlDocumentTemplate[]>
): ReturnType<typeof useDataCache<IJmlDocumentTemplate[]>> {
  const service = React.useMemo(
    () => new DocumentTemplateService(sp, templateLibraryUrl),
    [sp, templateLibraryUrl]
  );

  const fetcher = React.useCallback(
    () => service.getTemplatesForProcess(processType, documentType),
    [service, processType, documentType]
  );

  const cacheKey = documentType
    ? TemplateCacheKeys.forProcessAndType(processType, documentType)
    : TemplateCacheKeys.forProcess(processType);

  return useDataCache<IJmlDocumentTemplate[]>(
    cacheKey,
    fetcher,
    {
      staleTime: 5 * 60 * 1000,
      cacheTime: 30 * 60 * 1000,
      ...options
    }
  );
}

/**
 * Invalidate all template caches
 */
export function invalidateTemplateCache(): void {
  invalidateCache('templates:');
}

/**
 * Invalidate cache for a specific template
 */
export function invalidateTemplateById(templateId: number): void {
  invalidateCache(TemplateCacheKeys.byId(templateId));
  // Also invalidate list caches since they may contain this template
  invalidateCache('templates:all');
  invalidateCache('templates:type:');
  invalidateCache('templates:process:');
}
