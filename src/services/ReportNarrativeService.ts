// @ts-nocheck
// ReportNarrativeService - AI-powered report narrative generation
// Extends AIService with report-specific narrative generation capabilities

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import {
  AIFeatureType,
  IReportNarrativeRequest,
  IGeneratedReportNarrative,
  NarrativeStyle,
  FocusArea
} from '../models/IAI';
import { INarrativeTemplate, ReportCategory } from '../models/IReportBuilder';
import { AIService } from './AIService';
import { logger } from './LoggingService';

export class ReportNarrativeService {
  private sp: SPFI;
  private aiService: AIService;
  private narrativeTemplates: Map<string, INarrativeTemplate> = new Map();
  private initialized: boolean = false;

  constructor(sp: SPFI, aiService: AIService) {
    this.sp = sp;
    this.aiService = aiService;
  }

  /**
   * Initialize service and load narrative templates
   */
  public async initialize(): Promise<void> {
    if (this.initialized) return;

    try {
      await this.aiService.initialize();
      await this.loadNarrativeTemplates();
      this.initialized = true;
    } catch (error) {
      logger.error('ReportNarrativeService', 'Failed to initialize:', error);
      this.initialized = true; // Continue without templates
    }
  }

  /**
   * Generate AI-powered report narrative
   */
  public async generateReportNarrative(request: IReportNarrativeRequest): Promise<IGeneratedReportNarrative> {
    await this.ensureInitialized();

    const startTime = Date.now();

    try {
      // Build AI prompt for report narrative
      const prompt = this.buildReportNarrativePrompt(request);

      // Call AIService with ReportNarrative feature type
      const response = await (this.aiService as any).callAzureOpenAI({
        type: AIFeatureType.ReportNarrative,
        data: { prompt },
        options: {
          includeExplanation: true,
          maxResults: 1
        }
      });

      if (!response.success || !response.data) {
        throw new Error(response.error || 'Failed to generate report narrative');
      }

      // Parse and structure the narrative response
      const narrative = this.parseNarrativeResponse(response.data, request);

      // Log usage
      await this.logNarrativeUsage({
        request,
        narrative,
        processingTime: Date.now() - startTime,
        tokensUsed: response.metadata.tokensUsed,
        success: true
      });

      return narrative;
    } catch (error) {
      await this.logNarrativeUsage({
        request,
        error: error instanceof Error ? error.message : 'Unknown error',
        processingTime: Date.now() - startTime,
        success: false
      });

      throw error;
    }
  }

  /**
   * Generate narrative from template with variable substitution
   */
  public async generateFromTemplate(templateId: string, variables: Record<string, any>): Promise<string> {
    await this.ensureInitialized();

    const template = this.narrativeTemplates.get(templateId);

    if (!template) {
      throw new Error(`Narrative template ${templateId} not found`);
    }

    // Replace {{variable}} placeholders with actual values
    let content = template.content;

    template.variables.forEach(variableName => {
      const value = variables[variableName] || '';
      const regex = new RegExp(`\\{\\{${variableName}\\}\\}`, 'g');
      content = content.replace(regex, String(value));
    });

    return content;
  }

  /**
   * Generate hybrid narrative (template + AI enhancement)
   */
  public async generateHybridNarrative(
    templateId: string,
    variables: Record<string, any>,
    request: IReportNarrativeRequest
  ): Promise<IGeneratedReportNarrative> {
    await this.ensureInitialized();

    // Start with template content
    const templateContent = await this.generateFromTemplate(templateId, variables);

    // Build AI prompt to enhance the template
    const enhancementPrompt = this.buildHybridEnhancementPrompt(templateContent, request);

    const response = await (this.aiService as any).callAzureOpenAI({
      type: AIFeatureType.ReportNarrative,
      data: { prompt: enhancementPrompt },
      options: { includeExplanation: true }
    });

    if (!response.success || !response.data) {
      // If AI fails, return template-only narrative
      return {
        content: templateContent,
        keyMetrics: [],
        insights: [],
        recommendations: [],
        sentiment: 'neutral',
        metadata: {
          wordCount: this.countWords(templateContent),
          generatedAt: new Date(),
          narrativeStyle: request.narrativeStyle,
          focusAreas: request.focusAreas
        }
      };
    }

    return this.parseNarrativeResponse(response.data, request);
  }

  /**
   * Get available narrative templates
   */
  public async getTemplates(category?: string): Promise<INarrativeTemplate[]> {
    await this.ensureInitialized();

    const templates = Array.from(this.narrativeTemplates.values());

    if (category) {
      return templates.filter(t => t.category === category);
    }

    return templates;
  }

  /**
   * Get template by ID
   */
  public async getTemplate(templateId: string): Promise<INarrativeTemplate | undefined> {
    await this.ensureInitialized();
    return this.narrativeTemplates.get(templateId);
  }

  /**
   * Analyze report data and suggest focus areas
   */
  public suggestFocusAreas(data: any, reportType: string): FocusArea[] {
    const focusAreas: FocusArea[] = [];

    // Analyze data to suggest relevant focus areas
    if (data.historicalData || data.trends) {
      focusAreas.push('trends');
    }

    if (data.benchmarks || data.comparison) {
      focusAreas.push('comparisons');
    }

    if (data.risks || data.compliance) {
      focusAreas.push('risks');
    }

    if (data.opportunities || data.improvements) {
      focusAreas.push('opportunities');
    }

    // Always include insights and recommendations
    focusAreas.push('insights', 'recommendations');

    return focusAreas;
  }

  /**
   * Build AI prompt for report narrative generation
   */
  private buildReportNarrativePrompt(request: IReportNarrativeRequest): string {
    const lengthGuidance = this.getLengthGuidance(request.length);
    const styleGuidance = this.getStyleGuidance(request.narrativeStyle);
    const focusGuidance = this.getFocusGuidance(request.focusAreas);

    return `You are a business intelligence analyst writing a report narrative for a ${request.reportType} report.

Report Title: ${request.reportTitle}
Data Type: ${request.dataType}
Narrative Style: ${request.narrativeStyle}
Tone: ${request.tone}
Length: ${request.length} (${lengthGuidance})

Report Data:
${JSON.stringify(request.data, null, 2)}

${request.filters ? `Applied Filters:\n${JSON.stringify(request.filters, null, 2)}\n` : ''}

Style Guidelines:
${styleGuidance}

Focus Areas:
${focusGuidance}

${request.customInstructions || ''}

Generate a compelling narrative that:
1. Summarizes key findings from the data
2. Provides actionable insights
3. Identifies trends and patterns
4. Highlights risks and opportunities
5. Makes data-driven recommendations

Respond in the following JSON format:
{
  "content": "The narrative text in markdown format...",
  "keyMetrics": ["Metric 1: value", "Metric 2: value"],
  "insights": ["Insight 1...", "Insight 2..."],
  "recommendations": ["Recommendation 1...", "Recommendation 2..."],
  "sentiment": "positive" | "neutral" | "cautionary" | "negative",
  "citations": [
    {
      "dataSource": "PM_Processes",
      "metric": "Total Processes",
      "value": 247
    }
  ]
}

Make the narrative engaging, data-driven, and actionable for ${request.reportType} stakeholders.`;
  }

  /**
   * Build hybrid enhancement prompt
   */
  private buildHybridEnhancementPrompt(templateContent: string, request: IReportNarrativeRequest): string {
    return `You are enhancing a report narrative template with AI-generated insights.

Original Template Content:
${templateContent}

Report Data:
${JSON.stringify(request.data, null, 2)}

Instructions:
1. Keep the structure and format of the template
2. Enhance the content with specific insights from the data
3. Add depth and context to the narrative
4. Identify additional patterns or trends not obvious in the template
5. Maintain the ${request.tone} tone

Respond in the following JSON format:
{
  "content": "Enhanced narrative with template structure preserved...",
  "keyMetrics": ["Key metrics highlighted"],
  "insights": ["Additional insights discovered"],
  "recommendations": ["Data-driven recommendations"],
  "sentiment": "positive" | "neutral" | "cautionary" | "negative"
}`;
  }

  /**
   * Parse AI response into structured narrative
   */
  private parseNarrativeResponse(data: any, request: IReportNarrativeRequest): IGeneratedReportNarrative {
    return {
      content: data.content || '',
      keyMetrics: data.keyMetrics || [],
      insights: data.insights || [],
      recommendations: data.recommendations || [],
      sentiment: data.sentiment || 'neutral',
      metadata: {
        wordCount: this.countWords(data.content || ''),
        generatedAt: new Date(),
        tokensUsed: data.tokensUsed,
        narrativeStyle: request.narrativeStyle,
        focusAreas: request.focusAreas
      },
      citations: data.citations || []
    };
  }

  /**
   * Get length guidance
   */
  private getLengthGuidance(length: 'brief' | 'standard' | 'detailed'): string {
    const guidance = {
      brief: '100-150 words, focus on top-line summary',
      standard: '200-300 words, balanced overview with key details',
      detailed: '400-600 words, comprehensive analysis with context'
    };

    return guidance[length];
  }

  /**
   * Get style guidance
   */
  private getStyleGuidance(style: NarrativeStyle): string {
    const guidance = {
      [NarrativeStyle.ExecutiveSummary]: 'High-level overview for C-suite. Focus on business impact, strategic implications, and key decisions needed.',
      [NarrativeStyle.DetailedAnalysis]: 'In-depth analysis for managers and analysts. Include methodology, detailed findings, and supporting evidence.',
      [NarrativeStyle.ActionItems]: 'Prioritized action list format. Clear next steps with owners, deadlines, and expected outcomes.',
      [NarrativeStyle.TrendAnalysis]: 'Focus on patterns over time. Identify trends, forecast future states, and explain drivers of change.',
      [NarrativeStyle.Comparison]: 'Comparative analysis format. Highlight differences, similarities, and relative performance.',
      [NarrativeStyle.Storytelling]: 'Narrative arc with beginning, middle, and end. Make data relatable through context and examples.'
    };

    return guidance[style];
  }

  /**
   * Get focus guidance
   */
  private getFocusGuidance(focusAreas: FocusArea[]): string {
    const descriptions: Record<FocusArea, string> = {
      trends: 'Identify and explain patterns, trajectories, and changes over time',
      insights: 'Surface non-obvious findings, correlations, and key takeaways',
      risks: 'Highlight potential issues, threats, and areas of concern',
      opportunities: 'Identify improvement areas, growth potential, and quick wins',
      recommendations: 'Provide specific, actionable advice based on the data',
      comparisons: 'Compare to benchmarks, targets, or historical performance'
    };

    return focusAreas.map(area => `- ${area}: ${descriptions[area]}`).join('\n');
  }

  /**
   * Load narrative templates from files
   */
  private async loadNarrativeTemplates(): Promise<void> {
    // Template IDs matching the JSON files
    const templateIds = [
      'executive-summary',
      'quarterly-business-review',
      'department-performance',
      'manager-workload',
      'compliance-audit',
      'first-day-readiness',
      'roi-business-case',
      'cost-per-process',
      'process-efficiency',
      'task-completion',
      'employee-experience',
      'turnover-retention',
      'action-items',
      'improvement-recommendations',
      'sla-performance',
      'automation-impact',
      'risk-assessment',
      'trend-analysis',
      'benchmark-comparison',
      'year-end-summary'
    ];

    // Note: In production, these would be loaded from SharePoint library or embedded as modules
    // For now, store template metadata for reference
    const templateMetadata: INarrativeTemplate[] = [
      {
        id: 'executive-summary',
        name: 'Executive Dashboard Summary',
        description: 'High-level overview for C-suite stakeholders',
        category: ReportCategory.Executive,
        content: '',
        variables: ['report_title', 'start_date', 'end_date', 'total_processes', 'percent_change'],
        suggestedFor: [ReportCategory.Executive, ReportCategory.Operational]
      }
      // Additional templates would be loaded here
    ];

    templateMetadata.forEach(template => {
      this.narrativeTemplates.set(template.id, template);
    });
  }

  /**
   * Count words in text
   */
  private countWords(text: string): number {
    return text.trim().split(/\s+/).length;
  }

  /**
   * Log narrative generation usage
   */
  private async logNarrativeUsage(log: {
    request?: IReportNarrativeRequest;
    narrative?: IGeneratedReportNarrative;
    error?: string;
    processingTime: number;
    tokensUsed?: number;
    success: boolean;
  }): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('PM_AIUsageLogs')
        .items
        .add({
          Title: `Report Narrative - ${new Date().toISOString()}`,
          FeatureType: AIFeatureType.ReportNarrative,
          RequestData: log.request ? JSON.stringify({
            reportType: log.request.reportType,
            reportTitle: log.request.reportTitle,
            narrativeStyle: log.request.narrativeStyle,
            length: log.request.length,
            focusAreas: log.request.focusAreas
          }) : undefined,
          ResponseData: log.narrative ? JSON.stringify({
            wordCount: log.narrative.metadata.wordCount,
            sentiment: log.narrative.sentiment,
            insightCount: log.narrative.insights.length,
            recommendationCount: log.narrative.recommendations.length
          }) : undefined,
          ErrorMessage: log.error,
          ProcessingTime: log.processingTime,
          TokensUsed: log.tokensUsed,
          Success: log.success
        });
    } catch (error) {
      logger.error('ReportNarrativeService', 'Failed to log usage:', error);
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
}
