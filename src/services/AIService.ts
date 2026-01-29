// @ts-nocheck
// AIService - Azure OpenAI integration for intelligent automation
// Provides AI-powered features for JML processes

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import {
  AIFeatureType,
  IAIAnalysisRequest,
  IAIAnalysisResponse,
  IAIConfig,
  ITaskRecommendation,
  ICompletionPrediction,
  ISentimentAnalysis,
  IGeneratedDocument,
  IDocumentGenerationRequest,
  DocumentTemplateType,
  IRiskAssessment,
  IAutoCategorization,
  IFeedbackAnalysis,
  IChatMessage,
  IChatContext
} from '../models/IAI';
import { IJmlTask } from '../models/IJmlTask';
import { IJmlProcess } from '../models/IJmlProcess';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class AIService {
  private sp: SPFI;
  private configs: Map<AIFeatureType, IAIConfig> = new Map();
  private initialized: boolean = false;

  // Azure OpenAI configuration
  private endpoint: string = '';
  private apiKey: string = '';
  private deploymentName: string = 'gpt-4';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Initialize AI service and load configurations
   */
  public async initialize(): Promise<void> {
    if (this.initialized) return;

    try {
      // Load AI configurations from SharePoint
      const configs = await this.sp.web.lists
        .getByTitle('PM_AIConfigs')
        .items
        .select('Id', 'Title', 'FeatureType', 'IsEnabled', 'Endpoint', 'ApiKey', 'DeploymentName', 'Configuration', 'SystemPrompt', 'Temperature', 'MaxTokens')
        .filter('IsEnabled eq true')();

      configs.forEach(config => {
        this.configs.set(config.FeatureType, config as IAIConfig);
      });

      // Set default configuration
      const defaultConfig = configs.find(c => c.FeatureType === AIFeatureType.TaskRecommendation);
      if (defaultConfig) {
        this.endpoint = defaultConfig.Endpoint || '';
        this.apiKey = defaultConfig.ApiKey || '';
        this.deploymentName = defaultConfig.DeploymentName || 'gpt-4';
      }

      this.initialized = true;
    } catch (error) {
      logger.error('AIService', 'Failed to initialize AIService:', error);
      this.initialized = true; // Continue without configs
    }
  }

  /**
   * Suggest tasks based on job role and department
   */
  public async suggestTasks(jobRole: string, department: string): Promise<IJmlTask[]> {
    await this.ensureInitialized();

    const startTime = Date.now();

    try {
      // Get all available tasks
      const allTasks = await this.sp.web.lists
        .getByTitle('PM_Tasks')
        .items
        .select('Id', 'Title', 'TaskCode', 'Category', 'Department', 'Description', 'EstimatedHours', 'RequiresApproval', 'Priority')
        .filter('IsActive eq true')();

      // Get historical data for similar roles
      const historicalProcesses = await this.getHistoricalProcesses(jobRole, department);

      // Build AI prompt
      const prompt = this.buildTaskRecommendationPrompt(jobRole, department, allTasks, historicalProcesses);

      // Call Azure OpenAI
      const response = await this.callAzureOpenAI({
        type: AIFeatureType.TaskRecommendation,
        data: { prompt },
        options: { maxResults: 20 }
      });

      if (!response.success || !response.data) {
        throw new Error(response.error || 'Failed to get task recommendations');
      }

      // Parse AI response
      const recommendations: ITaskRecommendation[] = response.data.recommendations || [];

      // Map to IJmlTask array
      const suggestedTasks = recommendations
        .filter(rec => rec.relevanceScore >= 0.6)
        .map(rec => rec.task)
        .slice(0, 15);

      // Log usage
      await this.logUsage({
        featureType: AIFeatureType.TaskRecommendation,
        request: { jobRole, department },
        response: { count: suggestedTasks.length },
        processingTime: Date.now() - startTime,
        success: true
      });

      return suggestedTasks;
    } catch (error) {
      await this.logUsage({
        featureType: AIFeatureType.TaskRecommendation,
        request: { jobRole, department },
        error: error instanceof Error ? error.message : 'Unknown error',
        processingTime: Date.now() - startTime,
        success: false
      });

      throw error;
    }
  }

  /**
   * Predict process completion date using AI
   */
  public async predictCompletionDate(processId: number): Promise<Date> {
    await this.ensureInitialized();

    const startTime = Date.now();

    try {
      // Validate process ID
      const validatedProcessId = ValidationUtils.validateInteger(processId, 'processId', 1);

      // Get process details
      const process = await this.sp.web.lists
        .getByTitle('PM_Processes')
        .items
        .getById(validatedProcessId)
        .select('Id', 'ProcessType', 'Priority', 'TotalTasks', 'CompletedTasks', 'StartDate', 'TargetCompletionDate', 'Department')();

      // Get process tasks with secure filter
      const filter = ValidationUtils.buildFilter('ProcessIDId', 'eq', validatedProcessId);
      const tasks = await this.sp.web.lists
        .getByTitle('PM_ProcessTasks')
        .items
        .select('TaskTitle', 'Status', 'EstimatedHours', 'ActualHours', 'AssignedToId')
        .filter(filter)();

      // Get historical completion data
      const historicalData = await this.getHistoricalCompletionData(process.ProcessType, process.Department);

      // Build AI prompt
      const prompt = this.buildPredictionPrompt(process, tasks, historicalData);

      // Call Azure OpenAI
      const response = await this.callAzureOpenAI({
        type: AIFeatureType.PredictiveAnalytics,
        data: { prompt },
        options: { includeExplanation: true }
      });

      if (!response.success || !response.data) {
        throw new Error(response.error || 'Failed to predict completion date');
      }

      const prediction: ICompletionPrediction = response.data;

      // Log usage
      await this.logUsage({
        featureType: AIFeatureType.PredictiveAnalytics,
        processId,
        request: { processId },
        response: { predictedDate: prediction.predictedDate, confidence: prediction.confidence },
        processingTime: Date.now() - startTime,
        success: true
      });

      return prediction.predictedDate;
    } catch (error) {
      await this.logUsage({
        featureType: AIFeatureType.PredictiveAnalytics,
        processId,
        request: { processId },
        error: error instanceof Error ? error.message : 'Unknown error',
        processingTime: Date.now() - startTime,
        success: false
      });

      throw error;
    }
  }

  /**
   * Generate document from template using AI
   */
  public async generateDocumentFromTemplate(template: string, data: any): Promise<string> {
    await this.ensureInitialized();

    const startTime = Date.now();

    try {
      const request: IDocumentGenerationRequest = {
        templateType: this.parseTemplateType(template),
        data,
        format: 'html',
        tone: 'professional'
      };

      // Build AI prompt
      const prompt = this.buildDocumentGenerationPrompt(request);

      // Call Azure OpenAI
      const response = await this.callAzureOpenAI({
        type: AIFeatureType.DocumentGeneration,
        data: { prompt },
        options: { maxResults: 1 }
      });

      if (!response.success || !response.data) {
        throw new Error(response.error || 'Failed to generate document');
      }

      const document: IGeneratedDocument = response.data;

      // Log usage
      await this.logUsage({
        featureType: AIFeatureType.DocumentGeneration,
        request: { template, dataKeys: Object.keys(data) },
        response: { wordCount: document.metadata.wordCount },
        processingTime: Date.now() - startTime,
        success: true
      });

      return document.content;
    } catch (error) {
      await this.logUsage({
        featureType: AIFeatureType.DocumentGeneration,
        request: { template, dataKeys: Object.keys(data) },
        error: error instanceof Error ? error.message : 'Unknown error',
        processingTime: Date.now() - startTime,
        success: false
      });

      throw error;
    }
  }

  /**
   * Analyze feedback sentiment using AI
   */
  public async analyzeFeedback(feedback: string): Promise<ISentimentAnalysis> {
    await this.ensureInitialized();

    const startTime = Date.now();

    try {
      // Build AI prompt
      const prompt = this.buildSentimentAnalysisPrompt(feedback);

      // Call Azure OpenAI
      const response = await this.callAzureOpenAI({
        type: AIFeatureType.SentimentAnalysis,
        data: { prompt },
        options: { includeExplanation: true }
      });

      if (!response.success || !response.data) {
        throw new Error(response.error || 'Failed to analyze feedback');
      }

      const analysis: ISentimentAnalysis = response.data;

      // Log usage
      await this.logUsage({
        featureType: AIFeatureType.SentimentAnalysis,
        request: { feedbackLength: feedback.length },
        response: { sentiment: analysis.sentiment, score: analysis.score },
        processingTime: Date.now() - startTime,
        success: true
      });

      return analysis;
    } catch (error) {
      await this.logUsage({
        featureType: AIFeatureType.SentimentAnalysis,
        request: { feedbackLength: feedback.length },
        error: error instanceof Error ? error.message : 'Unknown error',
        processingTime: Date.now() - startTime,
        success: false
      });

      throw error;
    }
  }

  /**
   * Auto-categorize process by risk and complexity
   */
  public async categorizeProcess(process: IJmlProcess): Promise<IAutoCategorization> {
    await this.ensureInitialized();

    const startTime = Date.now();

    try {
      const prompt = this.buildCategorizationPrompt(process);

      const response = await this.callAzureOpenAI({
        type: AIFeatureType.AutoCategorization,
        data: { prompt }
      });

      if (!response.success || !response.data) {
        throw new Error(response.error || 'Failed to categorize process');
      }

      const categorization: IAutoCategorization = response.data;

      await this.logUsage({
        featureType: AIFeatureType.AutoCategorization,
        processId: process.Id,
        request: { processType: process.ProcessType },
        response: { complexity: categorization.complexity, confidence: categorization.confidence },
        processingTime: Date.now() - startTime,
        success: true
      });

      return categorization;
    } catch (error) {
      await this.logUsage({
        featureType: AIFeatureType.AutoCategorization,
        processId: process.Id,
        request: { processType: process.ProcessType },
        error: error instanceof Error ? error.message : 'Unknown error',
        processingTime: Date.now() - startTime,
        success: false
      });

      throw error;
    }
  }

  /**
   * Assess process risk using AI
   */
  public async assessRisk(process: IJmlProcess): Promise<IRiskAssessment> {
    await this.ensureInitialized();

    const prompt = `Assess the risk level for this JML process:

Process Type: ${process.ProcessType}
Department: ${process.Department}
Job Title: ${process.JobTitle}
Priority: ${process.Priority}
Total Tasks: ${process.TotalTasks}
Target Completion: ${process.TargetCompletionDate}

Provide a comprehensive risk assessment including risk level, factors, and mitigation recommendations.`;

    const response = await this.callAzureOpenAI({
      type: AIFeatureType.RiskAssessment,
      data: { prompt }
    });

    return response.data;
  }

  /**
   * Analyze exit interview feedback
   */
  public async analyzeExitFeedback(feedbackItems: string[]): Promise<IFeedbackAnalysis> {
    await this.ensureInitialized();

    const combinedFeedback = feedbackItems.join('\n\n');

    const prompt = `Analyze the following exit interview feedback and provide insights:

${combinedFeedback}

Provide:
1. Overall sentiment
2. Common themes
3. Action items for HR
4. Retention risk assessment`;

    const response = await this.callAzureOpenAI({
      type: AIFeatureType.SentimentAnalysis,
      data: { prompt }
    });

    return response.data;
  }

  /**
   * Generate AI-powered report narrative (Reports Builder feature)
   */
  public async generateReportNarrative(request: {
    reportType: string;
    reportTitle: string;
    dataType: string;
    data: any;
    filters?: any;
    narrativeStyle: string;
    tone: 'professional' | 'friendly' | 'technical' | 'marketing';
    length: 'brief' | 'standard' | 'detailed';
    focusAreas: string[];
    customInstructions?: string;
  }): Promise<{
    content: string;
    keyMetrics: string[];
    insights: string[];
    recommendations: string[];
    sentiment: 'positive' | 'neutral' | 'cautionary' | 'negative';
    metadata: {
      wordCount: number;
      generatedAt: Date;
      tokensUsed?: number;
      narrativeStyle: string;
      focusAreas: string[];
    };
  }> {
    await this.ensureInitialized();

    const startTime = Date.now();

    try {
      const prompt = this.buildReportNarrativePrompt(request);

      const response = await this.callAzureOpenAI({
        type: AIFeatureType.ReportNarrative,
        data: { prompt },
        options: { includeExplanation: true, maxResults: 1 }
      });

      if (!response.success || !response.data) {
        throw new Error(response.error || 'Failed to generate report narrative');
      }

      const narrative = {
        content: response.data.content || '',
        keyMetrics: response.data.keyMetrics || [],
        insights: response.data.insights || [],
        recommendations: response.data.recommendations || [],
        sentiment: response.data.sentiment || 'neutral',
        metadata: {
          wordCount: this.countWords(response.data.content || ''),
          generatedAt: new Date(),
          tokensUsed: response.metadata.tokensUsed,
          narrativeStyle: request.narrativeStyle,
          focusAreas: request.focusAreas
        }
      };

      await this.logUsage({
        featureType: AIFeatureType.ReportNarrative,
        request: { reportType: request.reportType, reportTitle: request.reportTitle },
        response: { wordCount: narrative.metadata.wordCount, sentiment: narrative.sentiment },
        processingTime: Date.now() - startTime,
        success: true
      });

      return narrative;
    } catch (error) {
      await this.logUsage({
        featureType: AIFeatureType.ReportNarrative,
        request: { reportType: request.reportType },
        error: error instanceof Error ? error.message : 'Unknown error',
        processingTime: Date.now() - startTime,
        success: false
      });

      throw error;
    }
  }

  /**
   * Build report narrative AI prompt
   */
  private buildReportNarrativePrompt(request: {
    reportType: string;
    reportTitle: string;
    dataType: string;
    data: any;
    filters?: any;
    narrativeStyle: string;
    tone: string;
    length: string;
    focusAreas: string[];
    customInstructions?: string;
  }): string {
    const lengthWords = { brief: '100-150 words', standard: '200-300 words', detailed: '400-600 words' }[request.length] || '200-300 words';

    return `You are a business intelligence analyst writing a ${request.narrativeStyle} narrative for a ${request.reportType} report.

Report: ${request.reportTitle}
Data Type: ${request.dataType}
Tone: ${request.tone}
Length: ${lengthWords}

Data:
${JSON.stringify(request.data, null, 2)}

${request.filters ? `Filters: ${JSON.stringify(request.filters, null, 2)}` : ''}

Focus: ${request.focusAreas.join(', ')}
${request.customInstructions || ''}

Generate narrative with:
1. Key findings summary
2. Actionable insights
3. Trends and patterns
4. Risks and opportunities
5. Data-driven recommendations

Respond with valid JSON:
{
  "content": "Markdown narrative...",
  "keyMetrics": ["Metric: value"],
  "insights": ["Insight"],
  "recommendations": ["Recommendation"],
  "sentiment": "positive|neutral|cautionary|negative"
}`;
  }

  /**
   * Count words in text
   */
  private countWords(text: string): number {
    if (!text) return 0;
    return text.trim().split(/\s+/).filter(w => w.length > 0).length;
  }

  /**
   * Chatbot conversation handler
   */
  public async chat(message: string, context: IChatContext): Promise<IChatMessage> {
    await this.ensureInitialized();

    // Sanitize user message to prevent prompt injection
    const sanitizedMessage = this.sanitizePromptInput(message, 2000);

    if (!sanitizedMessage || sanitizedMessage.trim().length === 0) {
      throw new Error('Message cannot be empty after sanitization');
    }

    const config = this.configs.get(AIFeatureType.Chatbot);
    const systemPrompt = config?.SystemPrompt || this.getDefaultChatbotPrompt();

    // Build conversation history with sanitized messages
    const messages = [
      { role: 'system', content: systemPrompt },
      ...context.conversationHistory.map(m => ({
        role: m.role,
        // Sanitize historical messages as well (defense in depth)
        content: m.role === 'user' ? this.sanitizePromptInput(m.content, 2000) : m.content
      })),
      { role: 'user', content: sanitizedMessage }
    ];

    const response = await this.callAzureOpenAIChat(messages);

    const assistantMessage: IChatMessage = {
      id: this.generateId(),
      role: 'assistant',
      content: response.data.content,
      timestamp: new Date(),
      context: context,
      citations: response.data.citations,
      actions: response.data.actions
    };

    return assistantMessage;
  }

  /**
   * Call Azure OpenAI API
   */
  private async callAzureOpenAI(request: IAIAnalysisRequest): Promise<IAIAnalysisResponse> {
    const config = this.configs.get(request.type);

    if (!config) {
      throw new Error(`AI feature ${request.type} is not configured`);
    }

    const endpoint = config.Endpoint || this.endpoint;
    const apiKey = config.ApiKey || this.apiKey;
    const deploymentName = config.DeploymentName || this.deploymentName;

    const url = `${endpoint}/openai/deployments/${deploymentName}/chat/completions?api-version=2024-02-15-preview`;

    const systemPrompt = config.SystemPrompt || this.getDefaultSystemPrompt(request.type);

    const requestBody = {
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: request.data.prompt }
      ],
      temperature: config.Temperature ?? 0.7,
      max_tokens: config.MaxTokens ?? 2000,
      top_p: config.TopP ?? 0.95,
      frequency_penalty: config.FrequencyPenalty ?? 0,
      presence_penalty: config.PresencePenalty ?? 0,
      response_format: { type: 'json_object' }
    };

    const startTime = Date.now();

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'api-key': apiKey
        },
        body: JSON.stringify(requestBody)
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Azure OpenAI API error: ${response.status} - ${errorText}`);
      }

      const result = await response.json();

      const content = result.choices[0].message.content;
      const parsedData = JSON.parse(content);

      return {
        success: true,
        data: parsedData,
        metadata: {
          timestamp: new Date(),
          processingTime: Date.now() - startTime,
          model: result.model,
          tokensUsed: result.usage?.total_tokens
        }
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error',
        metadata: {
          timestamp: new Date(),
          processingTime: Date.now() - startTime,
          model: deploymentName
        }
      };
    }
  }

  /**
   * Call Azure OpenAI Chat API
   */
  private async callAzureOpenAIChat(messages: any[]): Promise<IAIAnalysisResponse> {
    const config = this.configs.get(AIFeatureType.Chatbot);
    const endpoint = config?.Endpoint || this.endpoint;
    const apiKey = config?.ApiKey || this.apiKey;
    const deploymentName = config?.DeploymentName || this.deploymentName;

    const url = `${endpoint}/openai/deployments/${deploymentName}/chat/completions?api-version=2024-02-15-preview`;

    const requestBody = {
      messages,
      temperature: config?.Temperature ?? 0.7,
      max_tokens: config?.MaxTokens ?? 1000
    };

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'api-key': apiKey
      },
      body: JSON.stringify(requestBody)
    });

    const result = await response.json();

    return {
      success: true,
      data: {
        content: result.choices[0].message.content,
        citations: [],
        actions: []
      },
      metadata: {
        timestamp: new Date(),
        processingTime: 0,
        model: result.model,
        tokensUsed: result.usage?.total_tokens
      }
    };
  }

  /**
   * Sanitize user input for AI prompts to prevent prompt injection
   */
  private sanitizePromptInput(input: string, maxLength: number = 500): string {
    if (!input || typeof input !== 'string') {
      return '';
    }

    // Remove potential prompt injection attempts
    let sanitized = input
      // Remove system/assistant/user role markers
      .replace(/system:|assistant:|user:/gi, '')
      // Remove special tokens
      .replace(/<\|.*?\|>/g, '')
      // Remove potential instruction overrides
      .replace(/ignore (previous|all) instructions?/gi, '')
      .replace(/new instructions?:/gi, '')
      .replace(/you are now/gi, '')
      .replace(/forget (everything|all|previous)/gi, '')
      // Remove script tags
      .replace(/<script[^>]*>.*?<\/script>/gi, '')
      // Remove common injection patterns
      .replace(/```[\s\S]*?```/g, '')
      .replace(/\[INST\].*?\[\/INST\]/gi, '');

    // Limit length to prevent token overflow
    sanitized = sanitized.substring(0, maxLength);

    // Trim and normalize whitespace
    return sanitized.trim().replace(/\s+/g, ' ');
  }

  /**
   * Build task recommendation prompt
   */
  private buildTaskRecommendationPrompt(jobRole: string, department: string, allTasks: any[], historicalData: any[]): string {
    // Sanitize user inputs to prevent prompt injection
    const safeJobRole = this.sanitizePromptInput(jobRole, 100);
    const safeDepartment = this.sanitizePromptInput(department, 100);

    return `You are an expert HR and IT process consultant. Recommend the most relevant tasks for a ${safeJobRole} in the ${safeDepartment} department.

Available Tasks:
${JSON.stringify(allTasks, null, 2)}

Historical Data from Similar Processes:
${JSON.stringify(historicalData, null, 2)}

Provide recommendations in the following JSON format:
{
  "recommendations": [
    {
      "task": { /* full task object */ },
      "relevanceScore": 0.95,
      "reasoning": "This task is critical for...",
      "estimatedHours": 2,
      "priority": "High"
    }
  ]
}

Prioritize tasks that:
1. Are commonly used for this role/department
2. Are required by policy or regulations
3. Have dependencies that should be completed early
4. Match the department and category`;
  }

  /**
   * Build prediction prompt
   */
  private buildPredictionPrompt(process: any, tasks: any[], historicalData: any[]): string {
    return `Predict the completion date for this JML process based on current progress and historical data.

Current Process:
${JSON.stringify(process, null, 2)}

Tasks:
${JSON.stringify(tasks, null, 2)}

Historical Completion Data:
${JSON.stringify(historicalData, null, 2)}

Provide prediction in JSON format:
{
  "predictedDate": "2025-01-15T00:00:00Z",
  "confidence": 0.85,
  "factors": [
    {
      "factor": "Current completion rate",
      "impact": "positive",
      "weight": 0.4
    }
  ],
  "risks": ["Pending approval may delay completion"],
  "recommendations": ["Expedite IT tasks to stay on track"]
}`;
  }

  /**
   * Build document generation prompt
   */
  private buildDocumentGenerationPrompt(request: IDocumentGenerationRequest): string {
    // Sanitize user inputs to prevent prompt injection
    const safeTemplateType = this.sanitizePromptInput(String(request.templateType), 50);
    const safeTone = this.sanitizePromptInput(request.tone || 'professional', 50);
    const safeFormat = this.sanitizePromptInput(request.format || 'html', 20);
    const safeCustomInstructions = this.sanitizePromptInput(request.customInstructions || '', 1000);

    return `Generate a professional ${safeTemplateType} document with the following data:

${JSON.stringify(request.data, null, 2)}

Tone: ${safeTone}
Format: ${safeFormat}

${safeCustomInstructions}

Provide response in JSON format:
{
  "content": "Generated document content...",
  "format": "html",
  "title": "Document Title",
  "metadata": {
    "generatedAt": "2025-01-01T00:00:00Z",
    "templateType": "${request.templateType}",
    "wordCount": 250,
    "estimatedReadTime": 2
  },
  "suggestions": ["Consider adding...", "You may want to..."]
}`;
  }

  /**
   * Build sentiment analysis prompt
   */
  private buildSentimentAnalysisPrompt(feedback: string): string {
    // Sanitize user feedback to prevent prompt injection
    const safeFeedback = this.sanitizePromptInput(feedback, 2000);

    return `Analyze the sentiment of the following employee feedback:

"${safeFeedback}"

Provide analysis in JSON format:
{
  "sentiment": "positive" | "neutral" | "negative",
  "score": 0.75,
  "confidence": 0.9,
  "keyPhrases": ["great team", "growth opportunities"],
  "emotions": {
    "joy": 0.6,
    "sadness": 0.1
  },
  "summary": "The employee expresses overall positive sentiment...",
  "recommendations": ["Continue team building activities", "Address work-life balance concerns"]
}`;
  }

  /**
   * Build categorization prompt
   */
  private buildCategorizationPrompt(process: IJmlProcess): string {
    return `Categorize this JML process by complexity, risk, and priority:

Process Type: ${process.ProcessType}
Department: ${process.Department}
Job Title: ${process.JobTitle}
Total Tasks: ${process.TotalTasks}
Priority: ${process.Priority}

Provide categorization in JSON format:
{
  "category": "Standard Onboarding",
  "confidence": 0.9,
  "reasoning": "Standard process with typical task count...",
  "tags": ["it-access", "hr-onboarding", "facilities"],
  "priority": "Medium",
  "complexity": "Moderate",
  "estimatedDuration": 14
}`;
  }

  /**
   * Get default system prompt for feature type
   */
  private getDefaultSystemPrompt(featureType: AIFeatureType): string {
    const prompts: Record<AIFeatureType, string> = {
      [AIFeatureType.TaskRecommendation]: 'You are an expert HR and IT consultant specializing in employee lifecycle management. Provide intelligent task recommendations based on job roles and historical data.',
      [AIFeatureType.PredictiveAnalytics]: 'You are a data scientist specializing in project management analytics. Provide accurate completion predictions based on current progress and historical trends.',
      [AIFeatureType.Chatbot]: this.getDefaultChatbotPrompt(),
      [AIFeatureType.AutoCategorization]: 'You are a process analyst. Categorize JML processes by complexity, risk, and priority with high accuracy.',
      [AIFeatureType.SentimentAnalysis]: 'You are an expert in sentiment analysis and employee feedback interpretation. Provide detailed emotional analysis and actionable insights.',
      [AIFeatureType.DocumentGeneration]: 'You are a professional business writer. Generate clear, professional documents tailored to the audience and purpose.',
      [AIFeatureType.RiskAssessment]: 'You are a risk management consultant. Assess process risks and provide mitigation strategies.',
      [AIFeatureType.ReportNarrative]: 'You are a professional report writer. Generate clear, engaging narratives that transform data into compelling stories for stakeholders.'
    };

    return prompts[featureType] || 'You are a helpful AI assistant for JML processes.';
  }

  /**
   * Get default chatbot system prompt
   */
  private getDefaultChatbotPrompt(): string {
    return `You are JML Assistant, an AI-powered helper for the Joiner/Mover/Leaver (JML) management system.

Your capabilities:
- Answer questions about JML processes, tasks, and policies
- Guide users through creating and managing processes
- Provide status updates on processes
- Explain task requirements and dependencies
- Suggest best practices for employee lifecycle management
- Help troubleshoot issues

Guidelines:
- Be helpful, professional, and concise
- Provide specific answers with examples when possible
- Offer to perform actions (create process, search tasks, etc.)
- Cite policies and documentation when relevant
- Ask clarifying questions if needed

Always respond in a friendly, professional tone.`;
  }

  /**
   * Get historical processes for similar roles
   */
  private async getHistoricalProcesses(jobRole: string, department: string): Promise<any[]> {
    try {
      // Validate and sanitize inputs
      if (!jobRole || typeof jobRole !== 'string') {
        throw new Error('Invalid job role');
      }
      if (!department || typeof department !== 'string') {
        throw new Error('Invalid department');
      }

      // Build secure filter
      const jobFilter = ValidationUtils.buildFilter('JobTitle', 'eq', jobRole);
      const deptFilter = ValidationUtils.buildFilter('Department', 'eq', department);
      const filter = `${jobFilter} or ${deptFilter}`;

      const processes = await this.sp.web.lists
        .getByTitle('PM_Processes')
        .items
        .select('Id', 'ProcessType', 'JobTitle', 'Department', 'TotalTasks', 'CompletedTasks', 'ActualCompletionDate')
        .filter(filter)
        .top(20)
        .orderBy('Created', false)();

      return processes;
    } catch {
      return [];
    }
  }

  /**
   * Get historical completion data
   */
  private async getHistoricalCompletionData(processType: string, department: string): Promise<any[]> {
    try {
      // Validate and sanitize inputs
      if (!processType || typeof processType !== 'string') {
        throw new Error('Invalid process type');
      }
      if (!department || typeof department !== 'string') {
        throw new Error('Invalid department');
      }

      // Build secure filter
      const typeFilter = ValidationUtils.buildFilter('ProcessType', 'eq', processType);
      const deptFilter = ValidationUtils.buildFilter('Department', 'eq', department);
      const filter = `${typeFilter} and ${deptFilter} and ActualCompletionDate ne null`;

      const processes = await this.sp.web.lists
        .getByTitle('PM_Processes')
        .items
        .select('ProcessType', 'Department', 'StartDate', 'ActualCompletionDate', 'TotalTasks', 'Priority')
        .filter(filter)
        .top(50)();

      return processes;
    } catch {
      return [];
    }
  }

  /**
   * Parse template type from string
   */
  private parseTemplateType(template: string): DocumentTemplateType {
    const templateMap: Record<string, DocumentTemplateType> = {
      'offer': DocumentTemplateType.OfferLetter,
      'welcome': DocumentTemplateType.WelcomeEmail,
      'exit': DocumentTemplateType.ExitSummary,
      'transfer': DocumentTemplateType.TransferNotification,
      'task': DocumentTemplateType.TaskInstructions,
      'approval': DocumentTemplateType.ApprovalRequest,
      'completion': DocumentTemplateType.CompletionReport,
      'feedback': DocumentTemplateType.FeedbackSurvey
    };

    const key = Object.keys(templateMap).find(k => template.toLowerCase().includes(k));
    return key ? templateMap[key] : DocumentTemplateType.TaskInstructions;
  }

  /**
   * Log AI usage
   */
  private async logUsage(log: {
    featureType: AIFeatureType;
    processId?: number;
    request: any;
    response?: any;
    error?: string;
    processingTime: number;
    success: boolean;
  }): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('PM_AIUsageLogs')
        .items
        .add({
          Title: `${log.featureType} - ${new Date().toISOString()}`,
          FeatureType: log.featureType,
          ProcessID: log.processId,
          RequestData: JSON.stringify(log.request),
          ResponseData: log.response ? JSON.stringify(log.response) : undefined,
          ErrorMessage: log.error,
          ProcessingTime: log.processingTime,
          Success: log.success
        });
    } catch (error) {
      logger.error('AIService', 'Failed to log AI usage:', error);
    }
  }

  /**
   * Ensure service is initialized
   */
  private async ensureInitialized(): Promise<void> {
    if (!this.initialized) {
      await this.initialize();
    }
  }

  /**
   * Generate unique ID
   */
  private generateId(): string {
    return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
  }
}
