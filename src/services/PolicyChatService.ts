// ============================================================================
// PolicyChatService — Client-side RAG orchestrator for AI Chat Assistant
// ============================================================================
// Searches policies via PolicyHubService, builds compact context, calls the
// Azure Function chat endpoint, and returns structured responses.
// ============================================================================

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { PolicyHubService } from './PolicyHubService';
import { logger } from './LoggingService';
import { ConfigKeys } from '../models/IJmlConfiguration';
import { IPolicy, PolicyStatus } from '../models/IPolicy';

// ── Types ──

export type ChatMode = 'policy-qa' | 'author-assist' | 'general-help';

export interface IChatMessageLocal {
  id: string;
  role: 'user' | 'assistant';
  content: string;
  timestamp: Date;
  citations?: IChatCitation[];
  suggestedActions?: IChatSuggestedAction[];
}

export interface IChatCitation {
  policyId: number;
  title: string;
  excerpt: string;
}

export interface IChatSuggestedAction {
  type: 'navigate' | 'search';
  label: string;
  url: string;
}

interface ChatFunctionRequest {
  message: string;
  conversationHistory: { role: 'user' | 'assistant'; content: string }[];
  mode: ChatMode;
  policyContext?: {
    policies: {
      id: number;
      title: string;
      summary: string;
      keyPoints: string[];
      category: string;
      complianceRisk: string;
      effectiveDate: string;
      status: string;
    }[];
  };
  userRole: string;
  maxTokens?: number;
}

interface ChatFunctionResponse {
  message: string;
  citations?: IChatCitation[];
  suggestedActions?: IChatSuggestedAction[];
  metadata: {
    model: string;
    tokensUsed: number;
    processingTimeMs: number;
  };
}

// ── Constants ──

const SESSION_STORAGE_KEY = 'pm_chat_session';
const LOCAL_STORAGE_URL_KEY = 'PM_AI_ChatFunctionUrl';
const MAX_CONTEXT_CHARS = 20000;
const MAX_CONTEXT_POLICIES = 5;

// ── Service ──

export class PolicyChatService {
  private sp: SPFI;
  private hubService: PolicyHubService;
  private functionUrl: string = '';
  private isEnabled: boolean = false;
  private maxTokens: number = 1000;
  private initialized: boolean = false;

  constructor(sp: SPFI) {
    this.sp = sp;
    this.hubService = new PolicyHubService(sp);
  }

  // ──────────── Initialization ────────────

  /**
   * Load configuration from PM_Configuration list, with localStorage fallback.
   */
  public async initialize(): Promise<void> {
    if (this.initialized) return;

    try {
      // Try loading from PM_Configuration list
      const configs = await this.sp.web.lists
        .getByTitle('PM_Configuration')
        .items.filter(
          `substringof('Integration.AI.Chat', ConfigKey)`
        )
        .select('ConfigKey', 'ConfigValue')
        .top(5)();

      const configMap: Record<string, string> = {};
      configs.forEach((item: any) => {
        configMap[item.ConfigKey] = item.ConfigValue;
      });

      this.isEnabled = configMap[ConfigKeys.AI_CHAT_ENABLED] === 'true';
      this.functionUrl = configMap[ConfigKeys.AI_CHAT_FUNCTION_URL] || '';
      this.maxTokens = parseInt(configMap[ConfigKeys.AI_CHAT_MAX_TOKENS] || '1000', 10);

    } catch (error) {
      logger.warn('PolicyChatService', 'PM_Configuration load failed, checking localStorage:', error);
    }

    // Fallback: localStorage
    if (!this.functionUrl) {
      const localUrl = localStorage.getItem(LOCAL_STORAGE_URL_KEY);
      if (localUrl) {
        this.functionUrl = localUrl;
        this.isEnabled = true;
      }
    }

    // Clamp maxTokens
    if (this.maxTokens < 200) this.maxTokens = 200;
    if (this.maxTokens > 2000) this.maxTokens = 2000;

    // Initialize hub service for policy search
    try {
      await this.hubService.initialize();
    } catch {
      logger.warn('PolicyChatService', 'PolicyHubService init failed — policy search unavailable');
    }

    this.initialized = true;
    logger.info('PolicyChatService', `Initialized — enabled=${this.isEnabled}, hasUrl=${!!this.functionUrl}`);
  }

  /**
   * Whether the chat assistant is available (configured + enabled).
   */
  public isAvailable(): boolean {
    return this.isEnabled && !!this.functionUrl;
  }

  // ──────────── Main API ────────────

  /**
   * Send a message to the AI assistant.
   * For policy-qa and author-assist modes, performs client-side RAG:
   *   1. Search published policies matching the user's question
   *   2. Build compact context (summaries + key points)
   *   3. Call the Azure Function with context + message
   */
  public async sendMessage(
    message: string,
    mode: ChatMode,
    userRole: string,
    conversationHistory: IChatMessageLocal[]
  ): Promise<IChatMessageLocal> {
    if (!this.isAvailable()) {
      throw new Error('AI Chat Assistant is not configured. Please contact your administrator.');
    }

    const startTime = Date.now();

    // Step 1: Search relevant policies (RAG — only for policy-aware modes)
    let policyContext: ChatFunctionRequest['policyContext'] | undefined;
    if (mode !== 'general-help') {
      try {
        policyContext = await this.searchRelevantPolicies(message);
      } catch (error) {
        logger.warn('PolicyChatService', 'Policy search failed, proceeding without context:', error);
      }
    }

    // Step 2: Build conversation history for the function
    const history = conversationHistory
      .slice(-10) // max 10 messages
      .map(msg => ({
        role: msg.role as 'user' | 'assistant',
        content: msg.content.substring(0, 2000),
      }));

    // Step 3: Call Azure Function
    const request: ChatFunctionRequest = {
      message,
      conversationHistory: history,
      mode,
      policyContext,
      userRole,
      maxTokens: this.maxTokens,
    };

    const response = await this.callChatFunction(request);

    logger.info('PolicyChatService', `Chat response: ${response.metadata.tokensUsed} tokens, ${Date.now() - startTime}ms total`);

    // Step 4: Build local message object
    return {
      id: this.generateId(),
      role: 'assistant',
      content: response.message,
      timestamp: new Date(),
      citations: response.citations,
      suggestedActions: response.suggestedActions,
    };
  }

  // ──────────── RAG: Policy Search ────────────

  /**
   * Search published policies matching the user's question.
   * Returns a compact context object (summaries + key points, ≤20K chars).
   */
  private async searchRelevantPolicies(
    query: string
  ): Promise<ChatFunctionRequest['policyContext'] | undefined> {
    try {
      const results = await this.hubService.searchPolicyHub({
        searchText: query,
        filters: {
          statuses: [PolicyStatus.Published],
          isActive: true,
        },
        page: 1,
        pageSize: MAX_CONTEXT_POLICIES,
        includeDocuments: false,
        includeFacets: false,
      });

      if (!results.policies || results.policies.length === 0) {
        return undefined;
      }

      return this.buildPolicyContext(results.policies);
    } catch (error) {
      logger.error('PolicyChatService', 'searchRelevantPolicies failed:', error);
      return undefined;
    }
  }

  /**
   * Build a compact policy context from full policy objects.
   * Extracts summaries, key points, and essential metadata.
   * Enforces a total character budget to stay within token limits.
   */
  private buildPolicyContext(
    policies: IPolicy[]
  ): ChatFunctionRequest['policyContext'] {
    let totalChars = 0;
    const contextPolicies: Array<{
      id: number; title: string; summary: string; keyPoints: string[];
      category: string; complianceRisk: string; effectiveDate: string; status: string;
    }> = [];

    for (const policy of policies) {
      // Extract summary — strip HTML if present
      const summary = this.stripHtml(policy.PolicySummary || policy.Description || '').substring(0, 500);

      // Extract key points from KeyPoints field or build from content
      let keyPoints: string[] = [];
      if (policy.KeyPoints) {
        try {
          keyPoints = typeof policy.KeyPoints === 'string'
            ? JSON.parse(policy.KeyPoints)
            : policy.KeyPoints;
        } catch {
          keyPoints = [String(policy.KeyPoints)];
        }
      }
      // Limit key points
      keyPoints = (keyPoints || []).slice(0, 5).map(kp => String(kp).substring(0, 200));

      const entry = {
        id: policy.Id || 0,
        title: policy.Title || '',
        summary,
        keyPoints,
        category: String(policy.PolicyCategory || ''),
        complianceRisk: String(policy.ComplianceRisk || 'Informational'),
        effectiveDate: policy.EffectiveDate ? new Date(policy.EffectiveDate).toISOString().split('T')[0] : '',
        status: String(policy.PolicyStatus || 'Published'),
      };

      // Check character budget
      const entryChars = JSON.stringify(entry).length;
      if (totalChars + entryChars > MAX_CONTEXT_CHARS) break;

      totalChars += entryChars;
      contextPolicies.push(entry);
    }

    return { policies: contextPolicies };
  }

  // ──────────── Azure Function Call ────────────

  /**
   * Call the Azure Function chat endpoint.
   */
  private async callChatFunction(request: ChatFunctionRequest): Promise<ChatFunctionResponse> {
    const response = await fetch(this.functionUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(request),
    });

    if (!response.ok) {
      const errorText = await response.text().catch(() => 'Unknown error');
      if (response.status === 429) {
        throw new Error('Too many requests. Please wait a moment before trying again.');
      }
      logger.error('PolicyChatService', `Function call failed: ${response.status} — ${errorText}`);
      throw new Error('AI service temporarily unavailable. Please try again in a moment.');
    }

    return response.json() as Promise<ChatFunctionResponse>;
  }

  // ──────────── Session Persistence ────────────

  /**
   * Save conversation to sessionStorage (survives page navigations within the tab).
   */
  public saveSession(messages: IChatMessageLocal[], mode: ChatMode): void {
    try {
      const data = JSON.stringify({ messages, mode, savedAt: new Date().toISOString() });
      sessionStorage.setItem(SESSION_STORAGE_KEY, data);
    } catch {
      // sessionStorage full or unavailable — silently ignore
    }
  }

  /**
   * Restore conversation from sessionStorage.
   */
  public restoreSession(): { messages: IChatMessageLocal[]; mode: ChatMode } | null {
    try {
      const raw = sessionStorage.getItem(SESSION_STORAGE_KEY);
      if (!raw) return null;

      const data = JSON.parse(raw);
      // Rehydrate Date objects
      const messages: IChatMessageLocal[] = (data.messages || []).map((m: any) => ({
        ...m,
        timestamp: new Date(m.timestamp),
      }));

      return { messages, mode: data.mode || 'policy-qa' };
    } catch {
      return null;
    }
  }

  /**
   * Clear the saved session.
   */
  public clearSession(): void {
    sessionStorage.removeItem(SESSION_STORAGE_KEY);
  }

  // ──────────── Suggested Prompts ────────────

  /**
   * Get role-aware suggested prompts for the given mode.
   */
  public getSuggestedPrompts(role: string, mode: ChatMode): string[] {
    switch (mode) {
      case 'policy-qa':
        return [
          "What's our data retention policy?",
          'Which policies require acknowledgement?',
          'When is the next policy review due?',
          'What are the key compliance requirements?',
        ];
      case 'author-assist': {
        const prompts = [
          'Help me draft an introduction for a new policy',
          'Review this section for clarity and compliance',
          'Suggest improvements for the scope section',
        ];
        if (role === 'Manager' || role === 'Admin') {
          prompts.push('What approval steps does this policy need?');
        }
        return prompts;
      }
      case 'general-help':
        return [
          'How do I create a new policy?',
          'Where can I see my acknowledgements?',
          'How do policy approvals work?',
          'How do I search for policies?',
        ];
      default:
        return [];
    }
  }

  // ──────────── Utilities ────────────

  private stripHtml(html: string): string {
    return html
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/\s+/g, ' ')
      .trim();
  }

  private generateId(): string {
    return `msg_${Date.now()}_${Math.random().toString(36).substring(2, 8)}`;
  }
}
