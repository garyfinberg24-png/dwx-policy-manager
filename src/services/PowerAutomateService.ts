// @ts-nocheck
/**
 * Power Automate Integration Service
 * Provides triggers and webhooks for Power Automate flows
 */

import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { logger } from './LoggingService';

/**
 * Flow trigger types
 */
export enum FlowTriggerType {
  DocumentGenerated = 'DocumentGenerated',
  DocumentApproved = 'DocumentApproved',
  DocumentRejected = 'DocumentRejected',
  BulkOperationCompleted = 'BulkOperationCompleted',
  TemplateUpdated = 'TemplateUpdated',
  ApprovalRequested = 'ApprovalRequested'
}

/**
 * Flow trigger payload base
 */
export interface IFlowTriggerPayload {
  /** Trigger type */
  triggerType: FlowTriggerType;
  /** Timestamp */
  timestamp: string;
  /** Source system */
  source: 'JML-DocumentBuilder';
  /** Correlation ID for tracking */
  correlationId: string;
  /** Additional data */
  data: Record<string, unknown>;
}

/**
 * Document generated trigger data
 */
export interface IDocumentGeneratedData {
  documentId: number;
  documentName: string;
  documentUrl: string;
  templateId: number;
  templateName: string;
  processId?: number;
  employeeName: string;
  generatedBy: string;
  generatedByEmail: string;
}

/**
 * Approval trigger data
 */
export interface IApprovalTriggerData {
  requestId: number;
  documentId: number;
  documentName: string;
  status: 'Approved' | 'Rejected' | 'Requested';
  approverName?: string;
  approverEmail?: string;
  comments?: string;
  requestedBy: string;
  requestedByEmail: string;
}

/**
 * Flow endpoint configuration
 */
export interface IFlowEndpoint {
  /** Flow URL (HTTP trigger) */
  url: string;
  /** Trigger types this endpoint handles */
  triggerTypes: FlowTriggerType[];
  /** Whether endpoint is active */
  isActive: boolean;
  /** Display name */
  name: string;
}

/**
 * Trigger result
 */
export interface ITriggerResult {
  success: boolean;
  triggerType: FlowTriggerType;
  correlationId: string;
  response?: unknown;
  error?: string;
}

/**
 * Power Automate Integration Service
 */
export class PowerAutomateService {
  private httpClient: HttpClient;
  private endpoints: Map<FlowTriggerType, IFlowEndpoint[]>;
  private defaultTimeout: number = 30000;

  constructor(httpClient: HttpClient) {
    this.httpClient = httpClient;
    this.endpoints = new Map();
  }

  /**
   * Register a flow endpoint for specific trigger types
   */
  public registerEndpoint(endpoint: IFlowEndpoint): void {
    for (let i = 0; i < endpoint.triggerTypes.length; i++) {
      const triggerType = endpoint.triggerTypes[i];
      const existing = this.endpoints.get(triggerType) || [];
      existing.push(endpoint);
      this.endpoints.set(triggerType, existing);
    }

    logger.info('PowerAutomateService', `Registered endpoint: ${endpoint.name} for ${endpoint.triggerTypes.join(', ')}`);
  }

  /**
   * Unregister a flow endpoint
   */
  public unregisterEndpoint(url: string): void {
    this.endpoints.forEach((endpoints, triggerType) => {
      const filtered = endpoints.filter(e => e.url !== url);
      this.endpoints.set(triggerType, filtered);
    });

    logger.info('PowerAutomateService', `Unregistered endpoint: ${url}`);
  }

  /**
   * Trigger flow for document generated event
   */
  public async triggerDocumentGenerated(data: IDocumentGeneratedData): Promise<ITriggerResult[]> {
    const payload = this.createPayload(FlowTriggerType.DocumentGenerated, data as unknown as Record<string, unknown>);
    return this.triggerFlows(FlowTriggerType.DocumentGenerated, payload);
  }

  /**
   * Trigger flow for document approved event
   */
  public async triggerDocumentApproved(data: IApprovalTriggerData): Promise<ITriggerResult[]> {
    const approvalData = { ...data, status: 'Approved' as const };
    const payload = this.createPayload(FlowTriggerType.DocumentApproved, approvalData as unknown as Record<string, unknown>);
    return this.triggerFlows(FlowTriggerType.DocumentApproved, payload);
  }

  /**
   * Trigger flow for document rejected event
   */
  public async triggerDocumentRejected(data: IApprovalTriggerData): Promise<ITriggerResult[]> {
    const rejectionData = { ...data, status: 'Rejected' as const };
    const payload = this.createPayload(FlowTriggerType.DocumentRejected, rejectionData as unknown as Record<string, unknown>);
    return this.triggerFlows(FlowTriggerType.DocumentRejected, payload);
  }

  /**
   * Trigger flow for approval requested event
   */
  public async triggerApprovalRequested(data: IApprovalTriggerData): Promise<ITriggerResult[]> {
    const requestData = { ...data, status: 'Requested' as const };
    const payload = this.createPayload(FlowTriggerType.ApprovalRequested, requestData as unknown as Record<string, unknown>);
    return this.triggerFlows(FlowTriggerType.ApprovalRequested, payload);
  }

  /**
   * Trigger flow for bulk operation completed
   */
  public async triggerBulkOperationCompleted(data: {
    operationId: string;
    total: number;
    succeeded: number;
    failed: number;
    templateName: string;
    initiatedBy: string;
    initiatedByEmail: string;
  }): Promise<ITriggerResult[]> {
    const payload = this.createPayload(FlowTriggerType.BulkOperationCompleted, data);
    return this.triggerFlows(FlowTriggerType.BulkOperationCompleted, payload);
  }

  /**
   * Trigger flow for template updated
   */
  public async triggerTemplateUpdated(data: {
    templateId: number;
    templateName: string;
    version: string;
    updatedBy: string;
    updatedByEmail: string;
    changeType: 'Created' | 'Modified' | 'Deleted';
  }): Promise<ITriggerResult[]> {
    const payload = this.createPayload(FlowTriggerType.TemplateUpdated, data);
    return this.triggerFlows(FlowTriggerType.TemplateUpdated, payload);
  }

  /**
   * Test connection to a flow endpoint
   */
  public async testEndpoint(url: string): Promise<{ success: boolean; message: string }> {
    try {
      const testPayload: IFlowTriggerPayload = {
        triggerType: FlowTriggerType.DocumentGenerated,
        timestamp: new Date().toISOString(),
        source: 'JML-DocumentBuilder',
        correlationId: this.generateCorrelationId(),
        data: { test: true }
      };

      const response = await this.callEndpoint(url, testPayload);

      if (response.ok) {
        return { success: true, message: 'Connection successful' };
      } else {
        return { success: false, message: `HTTP ${response.status}: ${response.statusText}` };
      }
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, message: errorMessage };
    }
  }

  /**
   * Get all registered endpoints
   */
  public getRegisteredEndpoints(): IFlowEndpoint[] {
    const allEndpoints: IFlowEndpoint[] = [];
    const seenUrls = new Set<string>();

    this.endpoints.forEach((endpoints) => {
      for (let i = 0; i < endpoints.length; i++) {
        if (!seenUrls.has(endpoints[i].url)) {
          seenUrls.add(endpoints[i].url);
          allEndpoints.push(endpoints[i]);
        }
      }
    });

    return allEndpoints;
  }

  /**
   * Create trigger payload
   */
  private createPayload(triggerType: FlowTriggerType, data: Record<string, unknown>): IFlowTriggerPayload {
    return {
      triggerType,
      timestamp: new Date().toISOString(),
      source: 'JML-DocumentBuilder',
      correlationId: this.generateCorrelationId(),
      data
    };
  }

  /**
   * Trigger all flows for a specific trigger type
   */
  private async triggerFlows(
    triggerType: FlowTriggerType,
    payload: IFlowTriggerPayload
  ): Promise<ITriggerResult[]> {
    const endpoints = this.endpoints.get(triggerType) || [];
    const activeEndpoints = endpoints.filter(e => e.isActive);

    if (activeEndpoints.length === 0) {
      logger.debug('PowerAutomateService', `No active endpoints for trigger type: ${triggerType}`);
      return [];
    }

    const results: ITriggerResult[] = [];

    for (let i = 0; i < activeEndpoints.length; i++) {
      const endpoint = activeEndpoints[i];
      try {
        const response = await this.callEndpoint(endpoint.url, payload);

        if (response.ok) {
          const responseData = await response.json().catch(() => ({}));
          results.push({
            success: true,
            triggerType,
            correlationId: payload.correlationId,
            response: responseData
          });

          logger.info('PowerAutomateService', `Flow triggered successfully: ${endpoint.name}`);
        } else {
          results.push({
            success: false,
            triggerType,
            correlationId: payload.correlationId,
            error: `HTTP ${response.status}: ${response.statusText}`
          });

          logger.warn('PowerAutomateService', `Flow trigger failed: ${endpoint.name} - HTTP ${response.status}`);
        }
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        results.push({
          success: false,
          triggerType,
          correlationId: payload.correlationId,
          error: errorMessage
        });

        logger.error('PowerAutomateService', `Flow trigger error: ${endpoint.name}`, error);
      }
    }

    return results;
  }

  /**
   * Call a flow endpoint
   */
  private async callEndpoint(
    url: string,
    payload: IFlowTriggerPayload
  ): Promise<HttpClientResponse> {
    const options: IHttpClientOptions = {
      body: JSON.stringify(payload),
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      }
    };

    return this.httpClient.post(url, HttpClient.configurations.v1, options);
  }

  /**
   * Generate a correlation ID for tracking
   */
  private generateCorrelationId(): string {
    const timestamp = Date.now().toString(36);
    const random = Math.random().toString(36).substring(2, 9);
    return `jml-${timestamp}-${random}`;
  }
}

/**
 * Create Power Automate service instance
 */
export function createPowerAutomateService(httpClient: HttpClient): PowerAutomateService {
  return new PowerAutomateService(httpClient);
}
