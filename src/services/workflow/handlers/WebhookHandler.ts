// @ts-nocheck
/**
 * WebhookHandler
 *
 * Handles Webhook step execution - calling external HTTP endpoints.
 * Supports various HTTP methods, headers, body templates, and response handling.
 *
 * @author JML Development Team
 * @version 1.0.0
 */

import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  IWorkflowStep,
  IActionContext,
  IActionResult
} from '../../../models/IWorkflow';
import { logger } from '../../LoggingService';

/**
 * Result of webhook execution
 */
export interface IWebhookResult extends IActionResult {
  httpStatus?: number;
  responseBody?: unknown;
  responseHeaders?: Record<string, string>;
  executionTimeMs?: number;
}

/**
 * Handler for Webhook step type
 */
export class WebhookHandler {
  private context: WebPartContext;
  private httpClient: HttpClient;

  constructor(context: WebPartContext) {
    this.context = context;
    this.httpClient = context.httpClient;
  }

  /**
   * Execute Webhook step
   */
  public async execute(
    step: IWorkflowStep,
    actionContext: IActionContext
  ): Promise<IWebhookResult> {
    const config = step.config;
    const startTime = Date.now();

    // Validate configuration
    if (!config.webhookUrl) {
      return {
        success: false,
        error: 'Webhook step requires webhookUrl configuration',
        nextAction: 'fail'
      };
    }

    try {
      // Resolve URL with variable substitution
      const resolvedUrl = this.replaceTokens(config.webhookUrl, actionContext);

      // Build headers
      const headers: Record<string, string> = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        ...this.resolveHeaders(config.webhookHeaders || {}, actionContext)
      };

      // Build body (for POST/PUT/PATCH)
      let body: string | undefined;
      if (config.webhookBodyTemplate && ['POST', 'PUT', 'PATCH'].includes(config.webhookMethod || 'POST')) {
        body = this.replaceTokens(config.webhookBodyTemplate, actionContext);
      }

      logger.info('WebhookHandler', `Calling webhook: ${config.webhookMethod || 'POST'} ${resolvedUrl}`);

      // Configure request
      const httpOptions: IHttpClientOptions = {
        headers,
        body
      };

      // Set timeout if configured
      const timeout = config.webhookTimeout || 30000; // Default 30 seconds

      // Execute HTTP request with timeout
      const response = await this.executeWithTimeout(
        resolvedUrl,
        config.webhookMethod || 'POST',
        httpOptions,
        timeout
      );

      const executionTimeMs = Date.now() - startTime;

      // Parse response
      const responseText = await response.text();
      let responseBody: unknown;

      try {
        responseBody = JSON.parse(responseText);
      } catch {
        responseBody = responseText;
      }

      // Check for success (2xx status codes)
      if (response.ok) {
        logger.info('WebhookHandler', `Webhook successful: ${response.status} in ${executionTimeMs}ms`);

        // Store response in variable if configured
        const outputVariables: Record<string, unknown> = {
          webhookStatus: response.status,
          webhookSuccess: true,
          webhookExecutionTimeMs: executionTimeMs
        };

        if (config.webhookResponseVariable) {
          outputVariables[config.webhookResponseVariable] = responseBody;
        }

        return {
          success: true,
          nextAction: 'continue',
          httpStatus: response.status,
          responseBody,
          executionTimeMs,
          outputVariables
        };
      } else {
        logger.warn('WebhookHandler', `Webhook failed: ${response.status} - ${responseText}`);

        return {
          success: false,
          error: `Webhook returned ${response.status}: ${typeof responseBody === 'string' ? responseBody : JSON.stringify(responseBody)}`,
          nextAction: 'fail',
          httpStatus: response.status,
          responseBody,
          executionTimeMs,
          outputVariables: {
            webhookStatus: response.status,
            webhookSuccess: false,
            webhookError: responseText
          }
        };
      }

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      const executionTimeMs = Date.now() - startTime;

      logger.error('WebhookHandler', 'Webhook execution failed', error);

      return {
        success: false,
        error: `Webhook failed: ${errorMessage}`,
        nextAction: 'fail',
        executionTimeMs,
        outputVariables: {
          webhookSuccess: false,
          webhookError: errorMessage
        }
      };
    }
  }

  /**
   * Execute HTTP request with timeout
   */
  private async executeWithTimeout(
    url: string,
    method: string,
    options: IHttpClientOptions,
    timeout: number
  ): Promise<HttpClientResponse> {
    // Create timeout promise
    const timeoutPromise = new Promise<never>((_, reject) => {
      setTimeout(() => reject(new Error(`Request timeout after ${timeout}ms`)), timeout);
    });

    // Create request promise based on method
    let requestPromise: Promise<HttpClientResponse>;

    switch (method.toUpperCase()) {
      case 'GET':
        requestPromise = this.httpClient.get(
          url,
          HttpClient.configurations.v1,
          options
        );
        break;

      case 'POST':
        requestPromise = this.httpClient.post(
          url,
          HttpClient.configurations.v1,
          options
        );
        break;

      case 'PUT':
        requestPromise = this.httpClient.fetch(
          url,
          HttpClient.configurations.v1,
          { ...options, method: 'PUT' }
        );
        break;

      case 'PATCH':
        requestPromise = this.httpClient.fetch(
          url,
          HttpClient.configurations.v1,
          { ...options, method: 'PATCH' }
        );
        break;

      case 'DELETE':
        requestPromise = this.httpClient.fetch(
          url,
          HttpClient.configurations.v1,
          { ...options, method: 'DELETE' }
        );
        break;

      default:
        requestPromise = this.httpClient.post(
          url,
          HttpClient.configurations.v1,
          options
        );
    }

    // Race between request and timeout
    return Promise.race([requestPromise, timeoutPromise]);
  }

  /**
   * Replace {{tokens}} in string with actual values
   */
  private replaceTokens(template: string, context: IActionContext): string {
    return template.replace(/\{\{(\w+(?:\.\w+)*)\}\}/g, (match, path) => {
      const value = this.resolveTokenPath(path, context);
      if (value === undefined || value === null) {
        return match; // Keep original if not found
      }
      return typeof value === 'object' ? JSON.stringify(value) : String(value);
    });
  }

  /**
   * Resolve token path from context
   */
  private resolveTokenPath(path: string, context: IActionContext): unknown {
    const parts = path.split('.');
    let current: unknown;

    // Check common prefixes
    if (parts[0] === 'variables') {
      current = context.variables;
      parts.shift();
    } else if (parts[0] === 'process') {
      current = context.process;
      parts.shift();
    } else if (parts[0] === 'instance') {
      current = context.workflowInstance;
      parts.shift();
    } else if (parts[0] === 'step') {
      current = context.currentStep;
      parts.shift();
    } else {
      // Try variables first, then process
      if (context.variables[parts[0]] !== undefined) {
        current = context.variables;
      } else {
        current = context.process;
      }
    }

    // Navigate path
    for (const part of parts) {
      if (current === null || current === undefined) {
        return undefined;
      }
      current = (current as Record<string, unknown>)[part];
    }

    return current;
  }

  /**
   * Resolve headers with token replacement
   */
  private resolveHeaders(
    headers: Record<string, string>,
    context: IActionContext
  ): Record<string, string> {
    const resolved: Record<string, string> = {};

    for (const [key, value] of Object.entries(headers)) {
      resolved[key] = this.replaceTokens(value, context);
    }

    return resolved;
  }

  /**
   * Validate webhook URL
   */
  public static validateWebhookUrl(url: string): boolean {
    try {
      const parsed = new URL(url);
      return ['http:', 'https:'].includes(parsed.protocol);
    } catch {
      return false;
    }
  }
}
