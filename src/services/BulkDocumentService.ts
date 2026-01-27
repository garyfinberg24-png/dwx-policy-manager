// @ts-nocheck
/**
 * Bulk Document Service
 * Handles bulk document generation operations
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/batching';
import { IJmlProcess, IJmlDocumentTemplate } from '../models';
import { DocumentTemplateService } from './DocumentTemplateService';
import { IDocxProcessingResult } from './DocxTemplateProcessor';
import { logger } from './LoggingService';

/**
 * Bulk operation status
 */
export type BulkOperationStatus = 'pending' | 'processing' | 'completed' | 'failed' | 'cancelled';

/**
 * Individual item result in a bulk operation
 */
export interface IBulkOperationItem {
  /** Process or item ID */
  id: number;
  /** Display name */
  name: string;
  /** Current status */
  status: BulkOperationStatus;
  /** Result if completed */
  result?: IDocxProcessingResult;
  /** Error message if failed */
  error?: string;
  /** Progress percentage (0-100) */
  progress: number;
}

/**
 * Bulk operation result
 */
export interface IBulkOperationResult {
  /** Total items to process */
  total: number;
  /** Successfully completed count */
  succeeded: number;
  /** Failed count */
  failed: number;
  /** Cancelled count */
  cancelled: number;
  /** Individual item results */
  items: IBulkOperationItem[];
  /** Overall operation status */
  status: BulkOperationStatus;
  /** Start time */
  startTime: Date;
  /** End time */
  endTime?: Date;
  /** Duration in milliseconds */
  duration?: number;
}

/**
 * Bulk generation options
 */
export interface IBulkGenerationOptions {
  /** Template to use */
  templateId: number;
  /** Processes to generate documents for */
  processIds: number[];
  /** Custom placeholder values per process (keyed by process ID) */
  placeholderValues?: { [processId: number]: { [key: string]: string } };
  /** Company info to use */
  companyInfo?: { name: string; address: string; phone: string };
  /** Maximum concurrent operations */
  concurrency?: number;
  /** Callback for progress updates */
  onProgress?: (result: IBulkOperationResult) => void;
  /** Whether to stop on first error */
  stopOnError?: boolean;
}

/**
 * Bulk Document Service
 */
export class BulkDocumentService {
  private sp: SPFI;
  private templateService: DocumentTemplateService;
  private isCancelled: boolean = false;

  constructor(sp: SPFI, templateLibraryUrl?: string) {
    this.sp = sp;
    this.templateService = new DocumentTemplateService(sp, templateLibraryUrl);
  }

  /**
   * Generate documents in bulk
   */
  public async generateBulk(options: IBulkGenerationOptions): Promise<IBulkOperationResult> {
    this.isCancelled = false;
    const startTime = new Date();

    const result: IBulkOperationResult = {
      total: options.processIds.length,
      succeeded: 0,
      failed: 0,
      cancelled: 0,
      items: options.processIds.map(id => ({
        id,
        name: `Process ${id}`,
        status: 'pending' as BulkOperationStatus,
        progress: 0
      })),
      status: 'processing',
      startTime
    };

    try {
      // Get the template
      const template = await this.templateService.getTemplateById(options.templateId);

      // Get all processes
      const processes = await this.getProcesses(options.processIds);

      // Update item names
      for (let i = 0; i < result.items.length; i++) {
        const process = processes.find(p => p.Id === result.items[i].id);
        if (process) {
          result.items[i].name = process.EmployeeName;
        }
      }

      // Process in batches based on concurrency
      const concurrency = options.concurrency || 3;
      const batches = this.chunkArray(options.processIds, concurrency);

      for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
        if (this.isCancelled) {
          // Mark remaining as cancelled
          for (let i = 0; i < result.items.length; i++) {
            if (result.items[i].status === 'pending') {
              result.items[i].status = 'cancelled';
              result.cancelled++;
            }
          }
          break;
        }

        const batch = batches[batchIndex];
        const batchPromises = batch.map(async (processId) => {
          const itemIndex = result.items.findIndex(item => item.id === processId);
          const process = processes.find(p => p.Id === processId);

          if (!process) {
            result.items[itemIndex].status = 'failed';
            result.items[itemIndex].error = 'Process not found';
            result.failed++;
            return;
          }

          result.items[itemIndex].status = 'processing';
          result.items[itemIndex].progress = 50;
          this.notifyProgress(options.onProgress, result);

          try {
            const placeholders = options.placeholderValues?.[processId] || {};
            const docResult = await this.templateService.createFromTemplate(
              options.templateId,
              process,
              placeholders,
              options.companyInfo
            );

            if (docResult.success) {
              result.items[itemIndex].status = 'completed';
              result.items[itemIndex].result = docResult;
              result.items[itemIndex].progress = 100;
              result.succeeded++;
            } else {
              result.items[itemIndex].status = 'failed';
              result.items[itemIndex].error = docResult.error;
              result.failed++;

              if (options.stopOnError) {
                this.isCancelled = true;
              }
            }
          } catch (error) {
            result.items[itemIndex].status = 'failed';
            result.items[itemIndex].error = error instanceof Error ? error.message : 'Unknown error';
            result.failed++;

            if (options.stopOnError) {
              this.isCancelled = true;
            }
          }
        });

        await Promise.all(batchPromises);
        this.notifyProgress(options.onProgress, result);
      }

      result.status = this.isCancelled ? 'cancelled' : (result.failed > 0 && result.succeeded === 0 ? 'failed' : 'completed');
      result.endTime = new Date();
      result.duration = result.endTime.getTime() - startTime.getTime();

      logger.info('BulkDocumentService', `Bulk generation completed: ${result.succeeded}/${result.total} succeeded`);
      return result;
    } catch (error) {
      result.status = 'failed';
      result.endTime = new Date();
      result.duration = result.endTime.getTime() - startTime.getTime();

      logger.error('BulkDocumentService', 'Bulk generation failed:', error);
      return result;
    }
  }

  /**
   * Cancel ongoing bulk operation
   */
  public cancel(): void {
    this.isCancelled = true;
    logger.info('BulkDocumentService', 'Bulk operation cancelled');
  }

  /**
   * Download all generated documents as a ZIP
   */
  public async downloadAsZip(result: IBulkOperationResult): Promise<Blob> {
    // For a full implementation, we would use JSZip
    // This is a simplified version that concatenates the blobs
    const successfulItems = result.items.filter(
      item => item.status === 'completed' && item.result?.blob
    );

    if (successfulItems.length === 0) {
      throw new Error('No documents to download');
    }

    // In a real implementation, use JSZip to create a proper ZIP
    // For now, return the first document as a placeholder
    if (successfulItems[0].result?.blob) {
      return successfulItems[0].result.blob;
    }

    throw new Error('No document blob available');
  }

  /**
   * Get processes by IDs
   */
  private async getProcesses(processIds: number[]): Promise<IJmlProcess[]> {
    try {
      const processes: IJmlProcess[] = [];

      // Batch get processes
      const batches = this.chunkArray(processIds, 100);

      for (let i = 0; i < batches.length; i++) {
        const batch = batches[i];
        const filterParts = batch.map(id => `Id eq ${id}`);
        const filter = filterParts.join(' or ');

        const items = await this.sp.web.lists
          .getByTitle('JML_Processes')
          .items.filter(filter)
          .select('*', 'Manager/Id', 'Manager/Title', 'Manager/EMail')
          .expand('Manager')();

        for (let j = 0; j < items.length; j++) {
          processes.push(this.mapToProcess(items[j]));
        }
      }

      return processes;
    } catch (error) {
      logger.error('BulkDocumentService', 'Failed to get processes:', error);
      return [];
    }
  }

  /**
   * Map SharePoint item to process
   */
  private mapToProcess(item: Record<string, unknown>): IJmlProcess {
    const manager = item.Manager as { Id?: number; Title?: string; EMail?: string } | undefined;

    return {
      Id: Number(item.Id),
      Title: String(item.Title || ''),
      ProcessType: String(item.ProcessType) as IJmlProcess['ProcessType'],
      ProcessStatus: String(item.ProcessStatus) as IJmlProcess['ProcessStatus'],
      EmployeeName: String(item.EmployeeName || ''),
      EmployeeEmail: String(item.EmployeeEmail || ''),
      EmployeeID: item.EmployeeID ? String(item.EmployeeID) : undefined,
      Department: String(item.Department || ''),
      JobTitle: String(item.JobTitle || ''),
      Location: String(item.Location || ''),
      Manager: manager ? {
        Id: manager.Id || 0,
        Title: manager.Title || '',
        EMail: manager.EMail || ''
      } : undefined,
      StartDate: new Date(String(item.StartDate)),
      TargetCompletionDate: new Date(String(item.TargetCompletionDate)),
      Priority: String(item.Priority || 'Normal') as IJmlProcess['Priority'],
      Created: new Date(String(item.Created)),
      Modified: new Date(String(item.Modified))
    };
  }

  /**
   * Split array into chunks
   */
  private chunkArray<T>(array: T[], chunkSize: number): T[][] {
    const chunks: T[][] = [];
    for (let i = 0; i < array.length; i += chunkSize) {
      chunks.push(array.slice(i, i + chunkSize));
    }
    return chunks;
  }

  /**
   * Notify progress callback
   */
  private notifyProgress(
    onProgress: ((result: IBulkOperationResult) => void) | undefined,
    result: IBulkOperationResult
  ): void {
    if (onProgress) {
      onProgress({ ...result });
    }
  }
}

export function createBulkDocumentService(sp: SPFI, templateLibraryUrl?: string): BulkDocumentService {
  return new BulkDocumentService(sp, templateLibraryUrl);
}
