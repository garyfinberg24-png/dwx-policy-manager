// @ts-nocheck
// SearchAIService - AI-powered search enhancements
// Provides natural language search, spell correction, and query understanding

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { AIService } from './AIService';
import { logger } from './LoggingService';

/**
 * Parsed natural language query result
 */
export interface IParsedQuery {
  originalQuery: string;
  structuredQuery: string;
  intent: SearchIntent;
  entities: IQueryEntity[];
  filters: IQueryFilter[];
  dateRange?: IDateRange;
  sortBy?: string;
  sortDirection?: 'asc' | 'desc';
  confidence: number;
}

/**
 * Search intent types
 */
export enum SearchIntent {
  FindContent = 'find_content',
  FindPerson = 'find_person',
  FindTask = 'find_task',
  FindProcess = 'find_process',
  FindDocument = 'find_document',
  StatusCheck = 'status_check',
  DueDateQuery = 'due_date_query',
  AssignmentQuery = 'assignment_query',
  Unknown = 'unknown'
}

/**
 * Extracted entity from query
 */
export interface IQueryEntity {
  type: 'person' | 'department' | 'date' | 'status' | 'process_type' | 'priority' | 'keyword';
  value: string;
  normalizedValue?: string;
  confidence: number;
}

/**
 * Filter extracted from natural language
 */
export interface IQueryFilter {
  field: string;
  operator: 'eq' | 'ne' | 'gt' | 'lt' | 'ge' | 'le' | 'contains';
  value: string | number | Date;
}

/**
 * Date range for queries
 */
export interface IDateRange {
  start?: Date;
  end?: Date;
  relative?: string; // "this week", "last month", etc.
}

/**
 * Spell correction suggestion
 */
export interface ISpellCorrection {
  original: string;
  corrected: string;
  confidence: number;
  alternatives: string[];
}

/**
 * Search analytics data
 */
export interface ISearchAnalytics {
  query: string;
  userId: string;
  timestamp: Date;
  resultCount: number;
  clickedResults: string[];
  refinersUsed: string[];
  scope: string;
  duration: number;
}

/**
 * Trending search item
 */
export interface ITrendingSearch {
  query: string;
  count: number;
  trend: 'rising' | 'stable' | 'falling';
  lastSearched: Date;
}

/**
 * Related content suggestion
 */
export interface IRelatedContent {
  title: string;
  url: string;
  type: string;
  relevanceScore: number;
  reason: string;
}

/**
 * AI-powered search service
 */
export class SearchAIService {
  private sp: SPFI;
  private aiService: AIService;
  private commonMisspellings: Map<string, string>;
  private jmlTerms: Set<string>;

  constructor(sp: SPFI) {
    this.sp = sp;
    this.aiService = new AIService(sp);
    this.commonMisspellings = this.buildMisspellingsMap();
    this.jmlTerms = this.buildJmlTermsSet();
  }

  /**
   * Initialize the service
   */
  public async initialize(): Promise<void> {
    await this.aiService.initialize();
  }

  /**
   * Parse natural language search query
   */
  public async parseNaturalLanguageQuery(query: string): Promise<IParsedQuery> {
    const startTime = Date.now();

    try {
      // First, try rule-based parsing for common patterns
      const ruleBasedResult = this.parseWithRules(query);

      if (ruleBasedResult.confidence >= 0.8) {
        return ruleBasedResult;
      }

      // For complex queries, use AI
      const aiResult = await this.parseWithAI(query);

      // Combine results, preferring AI for complex queries
      return aiResult.confidence > ruleBasedResult.confidence ? aiResult : ruleBasedResult;
    } catch (error) {
      logger.error('SearchAIService', 'Failed to parse natural language query:', error);

      // Return basic structured query on error
      return {
        originalQuery: query,
        structuredQuery: query,
        intent: SearchIntent.FindContent,
        entities: [],
        filters: [],
        confidence: 0.5
      };
    }
  }

  /**
   * Rule-based query parsing for common patterns
   */
  private parseWithRules(query: string): IParsedQuery {
    const lowerQuery = query.toLowerCase().trim();
    const entities: IQueryEntity[] = [];
    const filters: IQueryFilter[] = [];
    let intent = SearchIntent.FindContent;
    let dateRange: IDateRange | undefined;
    let confidence = 0.6;

    // Detect intent from keywords
    if (/\b(find|search|show|get|list)\s+(all\s+)?tasks?\b/i.test(query)) {
      intent = SearchIntent.FindTask;
      confidence = 0.85;
    } else if (/\b(find|search|show|get|list)\s+(all\s+)?process(es)?\b/i.test(query)) {
      intent = SearchIntent.FindProcess;
      confidence = 0.85;
    } else if (/\b(find|search|show|get|list)\s+(all\s+)?document(s)?\b/i.test(query)) {
      intent = SearchIntent.FindDocument;
      confidence = 0.85;
    } else if (/\bwho\b|\bperson\b|\bemployee\b|\buser\b/i.test(query)) {
      intent = SearchIntent.FindPerson;
      confidence = 0.8;
    } else if (/\bstatus\b|\bprogress\b|\bhow\s+is\b/i.test(query)) {
      intent = SearchIntent.StatusCheck;
      confidence = 0.8;
    } else if (/\bdue\b|\bdeadline\b|\bexpir(e|ing|ed)\b/i.test(query)) {
      intent = SearchIntent.DueDateQuery;
      confidence = 0.8;
    } else if (/\bassigned\s+to\b|\bmy\s+tasks?\b|\bowner\b/i.test(query)) {
      intent = SearchIntent.AssignmentQuery;
      confidence = 0.8;
    }

    // Extract date-related phrases
    const datePatterns: Array<{ pattern: RegExp; handler: (match: RegExpMatchArray) => IDateRange }> = [
      {
        pattern: /\bthis\s+week\b/i,
        handler: () => this.getThisWeekRange()
      },
      {
        pattern: /\blast\s+week\b/i,
        handler: () => this.getLastWeekRange()
      },
      {
        pattern: /\bnext\s+week\b/i,
        handler: () => this.getNextWeekRange()
      },
      {
        pattern: /\bthis\s+month\b/i,
        handler: () => this.getThisMonthRange()
      },
      {
        pattern: /\blast\s+month\b/i,
        handler: () => this.getLastMonthRange()
      },
      {
        pattern: /\btoday\b/i,
        handler: () => this.getTodayRange()
      },
      {
        pattern: /\btomorrow\b/i,
        handler: () => this.getTomorrowRange()
      },
      {
        pattern: /\boverdue\b/i,
        handler: () => ({ end: new Date(), relative: 'overdue' })
      },
      {
        pattern: /\bdue\s+(?:in\s+)?(\d+)\s+days?\b/i,
        handler: (match) => this.getDueDaysRange(parseInt(match[1], 10))
      }
    ];

    for (const { pattern, handler } of datePatterns) {
      const match = query.match(pattern);
      if (match) {
        dateRange = handler(match);
        entities.push({
          type: 'date',
          value: match[0],
          normalizedValue: dateRange.relative,
          confidence: 0.9
        });
        break;
      }
    }

    // Extract status keywords
    const statusPatterns = [
      { pattern: /\b(not\s+started|pending|new)\b/i, value: 'Not Started' },
      { pattern: /\b(in\s+progress|active|ongoing)\b/i, value: 'In Progress' },
      { pattern: /\b(completed?|done|finished)\b/i, value: 'Completed' },
      { pattern: /\b(blocked|stuck|waiting)\b/i, value: 'Blocked' },
      { pattern: /\b(cancelled?|canceled?)\b/i, value: 'Cancelled' }
    ];

    for (const { pattern, value } of statusPatterns) {
      if (pattern.test(query)) {
        entities.push({
          type: 'status',
          value: value,
          confidence: 0.9
        });
        filters.push({
          field: 'Status',
          operator: 'eq',
          value: value
        });
      }
    }

    // Extract priority keywords
    const priorityPatterns = [
      { pattern: /\b(high|urgent|critical)\s+priority\b/i, value: 'High' },
      { pattern: /\b(medium|normal)\s+priority\b/i, value: 'Medium' },
      { pattern: /\b(low)\s+priority\b/i, value: 'Low' }
    ];

    for (const { pattern, value } of priorityPatterns) {
      if (pattern.test(query)) {
        entities.push({
          type: 'priority',
          value: value,
          confidence: 0.85
        });
        filters.push({
          field: 'Priority',
          operator: 'eq',
          value: value
        });
      }
    }

    // Extract process types
    const processTypes = ['onboarding', 'offboarding', 'transfer', 'promotion', 'leaver', 'joiner', 'mover'];
    for (const processType of processTypes) {
      if (lowerQuery.includes(processType)) {
        entities.push({
          type: 'process_type',
          value: processType,
          normalizedValue: this.capitalizeFirst(processType),
          confidence: 0.9
        });
        filters.push({
          field: 'ProcessType',
          operator: 'eq',
          value: this.capitalizeFirst(processType)
        });
      }
    }

    // Extract department mentions
    const departments = ['hr', 'it', 'finance', 'marketing', 'sales', 'operations', 'legal', 'engineering'];
    for (const dept of departments) {
      if (lowerQuery.includes(dept)) {
        entities.push({
          type: 'department',
          value: dept,
          normalizedValue: dept.toUpperCase(),
          confidence: 0.85
        });
        filters.push({
          field: 'Department',
          operator: 'eq',
          value: dept.toUpperCase()
        });
      }
    }

    // Build structured query
    const structuredQuery = this.buildStructuredQuery(query, filters, dateRange);

    return {
      originalQuery: query,
      structuredQuery,
      intent,
      entities,
      filters,
      dateRange,
      confidence
    };
  }

  /**
   * AI-powered query parsing for complex queries
   */
  private async parseWithAI(query: string): Promise<IParsedQuery> {
    try {
      const prompt = `Parse this natural language search query for a JML (Joiner/Mover/Leaver) HR management system:

Query: "${query}"

Extract:
1. Search intent (find_content, find_person, find_task, find_process, find_document, status_check, due_date_query, assignment_query)
2. Entities (person names, departments, dates, statuses, process types, priorities)
3. Filters that should be applied
4. Date ranges if mentioned
5. Sort preferences if mentioned

JML Context:
- Processes: Onboarding, Offboarding, Transfer, Promotion
- Statuses: Not Started, In Progress, Completed, Blocked
- Priorities: High, Medium, Low
- Common tasks: IT setup, Badge creation, Access provisioning, Training

Respond with valid JSON:
{
  "intent": "find_task",
  "entities": [
    {"type": "date", "value": "this week", "normalizedValue": "2025-01-27/2025-02-02", "confidence": 0.9}
  ],
  "filters": [
    {"field": "DueDate", "operator": "le", "value": "2025-02-02"}
  ],
  "dateRange": {"start": "2025-01-27", "end": "2025-02-02", "relative": "this week"},
  "sortBy": "DueDate",
  "sortDirection": "asc",
  "structuredQuery": "tasks due this week",
  "confidence": 0.85
}`;

      const response = await this.aiService.chat(prompt, {
        conversationHistory: [],
        sessionId: `nlp_parse_${Date.now()}`
      });

      // Parse AI response
      const parsed = JSON.parse(response.content);

      return {
        originalQuery: query,
        structuredQuery: parsed.structuredQuery || query,
        intent: parsed.intent as SearchIntent || SearchIntent.FindContent,
        entities: parsed.entities || [],
        filters: parsed.filters || [],
        dateRange: parsed.dateRange ? {
          start: parsed.dateRange.start ? new Date(parsed.dateRange.start) : undefined,
          end: parsed.dateRange.end ? new Date(parsed.dateRange.end) : undefined,
          relative: parsed.dateRange.relative
        } : undefined,
        sortBy: parsed.sortBy,
        sortDirection: parsed.sortDirection,
        confidence: parsed.confidence || 0.7
      };
    } catch (error) {
      logger.error('SearchAIService', 'AI parsing failed:', error);
      return this.parseWithRules(query);
    }
  }

  /**
   * Get spell correction suggestions
   */
  public getSpellCorrections(query: string): ISpellCorrection | null {
    const words = query.toLowerCase().split(/\s+/);
    const corrections: Array<{ original: string; corrected: string }> = [];

    for (const word of words) {
      // Check common misspellings
      if (this.commonMisspellings.has(word)) {
        corrections.push({
          original: word,
          corrected: this.commonMisspellings.get(word)!
        });
      }

      // Check JML terms using fuzzy matching
      if (word.length >= 4 && !this.jmlTerms.has(word)) {
        const closestMatch = this.findClosestMatch(word, Array.from(this.jmlTerms));
        if (closestMatch && this.levenshteinDistance(word, closestMatch) <= 2) {
          corrections.push({
            original: word,
            corrected: closestMatch
          });
        }
      }
    }

    if (corrections.length === 0) {
      return null;
    }

    // Build corrected query
    let correctedQuery = query.toLowerCase();
    for (const { original, corrected } of corrections) {
      correctedQuery = correctedQuery.replace(new RegExp(`\\b${original}\\b`, 'gi'), corrected);
    }

    // Generate alternatives
    const alternatives = this.generateAlternatives(query, corrections);

    return {
      original: query,
      corrected: correctedQuery,
      confidence: Math.min(0.95, 0.6 + (corrections.length * 0.15)),
      alternatives
    };
  }

  /**
   * Get related content suggestions based on search results
   */
  public async getRelatedContent(searchQuery: string, resultIds: string[]): Promise<IRelatedContent[]> {
    try {
      // Get metadata for clicked results to find patterns
      const relatedItems: IRelatedContent[] = [];

      // Suggest related process types
      if (searchQuery.toLowerCase().includes('onboarding')) {
        relatedItems.push({
          title: 'Onboarding Checklist Templates',
          url: '/sites/jml/Lists/Templates',
          type: 'template',
          relevanceScore: 0.85,
          reason: 'Related to onboarding processes'
        });
      }

      // Suggest related documentation
      if (searchQuery.toLowerCase().includes('task')) {
        relatedItems.push({
          title: 'Task Management Guide',
          url: '/sites/jml/Docs/TaskGuide.pdf',
          type: 'document',
          relevanceScore: 0.75,
          reason: 'Helpful documentation for task management'
        });
      }

      return relatedItems;
    } catch (error) {
      logger.error('SearchAIService', 'Failed to get related content:', error);
      return [];
    }
  }

  /**
   * Track search analytics
   */
  public async trackSearch(analytics: ISearchAnalytics): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('JML_SearchAnalytics')
        .items
        .add({
          Title: analytics.query.substring(0, 255),
          Query: analytics.query,
          UserId: analytics.userId,
          SearchTimestamp: analytics.timestamp.toISOString(),
          ResultCount: analytics.resultCount,
          ClickedResults: JSON.stringify(analytics.clickedResults),
          RefinersUsed: JSON.stringify(analytics.refinersUsed),
          SearchScope: analytics.scope,
          Duration: analytics.duration
        });
    } catch (error) {
      logger.error('SearchAIService', 'Failed to track search:', error);
    }
  }

  /**
   * Get trending searches
   */
  public async getTrendingSearches(limit: number = 10): Promise<ITrendingSearch[]> {
    try {
      // Get searches from last 7 days
      const sevenDaysAgo = new Date();
      sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

      const searches = await this.sp.web.lists
        .getByTitle('JML_SearchAnalytics')
        .items
        .select('Query', 'SearchTimestamp', 'ResultCount')
        .filter(`SearchTimestamp ge datetime'${sevenDaysAgo.toISOString()}'`)
        .top(1000)();

      // Aggregate searches
      const searchCounts = new Map<string, { count: number; lastSearched: Date; totalResults: number }>();

      for (const search of searches) {
        const query = search.Query.toLowerCase().trim();
        const existing = searchCounts.get(query);
        const timestamp = new Date(search.SearchTimestamp);

        if (existing) {
          existing.count++;
          existing.totalResults += search.ResultCount;
          if (timestamp > existing.lastSearched) {
            existing.lastSearched = timestamp;
          }
        } else {
          searchCounts.set(query, {
            count: 1,
            lastSearched: timestamp,
            totalResults: search.ResultCount
          });
        }
      }

      // Sort by count and return top N
      const trending: ITrendingSearch[] = Array.from(searchCounts.entries())
        .sort((a, b) => b[1].count - a[1].count)
        .slice(0, limit)
        .map(([query, data]) => ({
          query,
          count: data.count,
          trend: this.determineTrend(data.count, data.lastSearched),
          lastSearched: data.lastSearched
        }));

      return trending;
    } catch (error) {
      logger.error('SearchAIService', 'Failed to get trending searches:', error);
      return [];
    }
  }

  /**
   * Get popular searches for suggestions
   */
  public async getPopularSearches(limit: number = 5): Promise<string[]> {
    const trending = await this.getTrendingSearches(limit);
    return trending.map(t => t.query);
  }

  // Helper methods

  private buildStructuredQuery(query: string, filters: IQueryFilter[], dateRange?: IDateRange): string {
    let structuredQuery = query;

    // Remove processed filter terms from query for cleaner search
    const filterTerms = filters.map(f => String(f.value).toLowerCase());
    for (const term of filterTerms) {
      structuredQuery = structuredQuery.replace(new RegExp(`\\b${term}\\b`, 'gi'), '').trim();
    }

    // Clean up extra spaces
    structuredQuery = structuredQuery.replace(/\s+/g, ' ').trim();

    return structuredQuery || query;
  }

  private buildMisspellingsMap(): Map<string, string> {
    return new Map([
      // Common typos for JML terms
      ['onbording', 'onboarding'],
      ['onbaording', 'onboarding'],
      ['ofboarding', 'offboarding'],
      ['offbording', 'offboarding'],
      ['transferr', 'transfer'],
      ['promtion', 'promotion'],
      ['emplyee', 'employee'],
      ['employe', 'employee'],
      ['emploee', 'employee'],
      ['taks', 'task'],
      ['tsak', 'task'],
      ['porcess', 'process'],
      ['proccess', 'process'],
      ['deparment', 'department'],
      ['departmant', 'department'],
      ['asigned', 'assigned'],
      ['assignd', 'assigned'],
      ['compleate', 'complete'],
      ['complted', 'completed'],
      ['penidng', 'pending'],
      ['pendng', 'pending'],
      ['aprroval', 'approval'],
      ['aprovall', 'approval'],
      ['documnet', 'document'],
      ['docuemnt', 'document'],
      ['tempalte', 'template'],
      ['templat', 'template'],
      ['traning', 'training'],
      ['trainig', 'training'],
      ['managment', 'management'],
      ['managament', 'management'],
      ['resouces', 'resources'],
      ['resoruces', 'resources']
    ]);
  }

  private buildJmlTermsSet(): Set<string> {
    return new Set([
      'onboarding', 'offboarding', 'transfer', 'promotion', 'joiner', 'mover', 'leaver',
      'task', 'process', 'employee', 'manager', 'department', 'assigned', 'completed',
      'pending', 'blocked', 'approval', 'document', 'template', 'training', 'checklist',
      'badge', 'access', 'equipment', 'laptop', 'provisioning', 'deprovisioning',
      'exit', 'interview', 'survey', 'feedback', 'hr', 'it', 'facilities',
      'orientation', 'induction', 'handover', 'knowledge', 'clearance', 'resignation',
      'termination', 'retirement', 'contract', 'compliance', 'policy', 'procedure'
    ]);
  }

  private findClosestMatch(word: string, candidates: string[]): string | null {
    let closest: string | null = null;
    let minDistance = Infinity;

    for (const candidate of candidates) {
      const distance = this.levenshteinDistance(word, candidate);
      if (distance < minDistance) {
        minDistance = distance;
        closest = candidate;
      }
    }

    return minDistance <= 2 ? closest : null;
  }

  private levenshteinDistance(a: string, b: string): number {
    const matrix: number[][] = [];

    for (let i = 0; i <= b.length; i++) {
      matrix[i] = [i];
    }

    for (let j = 0; j <= a.length; j++) {
      matrix[0][j] = j;
    }

    for (let i = 1; i <= b.length; i++) {
      for (let j = 1; j <= a.length; j++) {
        if (b.charAt(i - 1) === a.charAt(j - 1)) {
          matrix[i][j] = matrix[i - 1][j - 1];
        } else {
          matrix[i][j] = Math.min(
            matrix[i - 1][j - 1] + 1,
            matrix[i][j - 1] + 1,
            matrix[i - 1][j] + 1
          );
        }
      }
    }

    return matrix[b.length][a.length];
  }

  private generateAlternatives(query: string, corrections: Array<{ original: string; corrected: string }>): string[] {
    const alternatives: string[] = [];

    // Generate alternative corrections
    for (const { original, corrected } of corrections) {
      const similar = Array.from(this.jmlTerms)
        .filter(term => this.levenshteinDistance(original, term) <= 3 && term !== corrected)
        .slice(0, 2);

      for (const alt of similar) {
        alternatives.push(query.replace(new RegExp(`\\b${original}\\b`, 'gi'), alt));
      }
    }

    return alternatives.slice(0, 3);
  }

  private determineTrend(count: number, lastSearched: Date): 'rising' | 'stable' | 'falling' {
    const daysSinceLastSearch = (Date.now() - lastSearched.getTime()) / (1000 * 60 * 60 * 24);

    if (daysSinceLastSearch < 1 && count > 5) return 'rising';
    if (daysSinceLastSearch > 3) return 'falling';
    return 'stable';
  }

  private capitalizeFirst(str: string): string {
    return str.charAt(0).toUpperCase() + str.slice(1);
  }

  // Date range helpers
  private getThisWeekRange(): IDateRange {
    const now = new Date();
    const start = new Date(now);
    start.setDate(now.getDate() - now.getDay());
    start.setHours(0, 0, 0, 0);

    const end = new Date(start);
    end.setDate(start.getDate() + 6);
    end.setHours(23, 59, 59, 999);

    return { start, end, relative: 'this week' };
  }

  private getLastWeekRange(): IDateRange {
    const now = new Date();
    const start = new Date(now);
    start.setDate(now.getDate() - now.getDay() - 7);
    start.setHours(0, 0, 0, 0);

    const end = new Date(start);
    end.setDate(start.getDate() + 6);
    end.setHours(23, 59, 59, 999);

    return { start, end, relative: 'last week' };
  }

  private getNextWeekRange(): IDateRange {
    const now = new Date();
    const start = new Date(now);
    start.setDate(now.getDate() - now.getDay() + 7);
    start.setHours(0, 0, 0, 0);

    const end = new Date(start);
    end.setDate(start.getDate() + 6);
    end.setHours(23, 59, 59, 999);

    return { start, end, relative: 'next week' };
  }

  private getThisMonthRange(): IDateRange {
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth(), 1);
    const end = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59, 999);

    return { start, end, relative: 'this month' };
  }

  private getLastMonthRange(): IDateRange {
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const end = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59, 999);

    return { start, end, relative: 'last month' };
  }

  private getTodayRange(): IDateRange {
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const end = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);

    return { start, end, relative: 'today' };
  }

  private getTomorrowRange(): IDateRange {
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
    const end = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1, 23, 59, 59, 999);

    return { start, end, relative: 'tomorrow' };
  }

  private getDueDaysRange(days: number): IDateRange {
    const now = new Date();
    const end = new Date(now);
    end.setDate(now.getDate() + days);
    end.setHours(23, 59, 59, 999);

    return { start: now, end, relative: `due in ${days} days` };
  }
}
