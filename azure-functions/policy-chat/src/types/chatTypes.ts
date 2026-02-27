// ============================================================================
// DWx Policy Manager — AI Chat Assistant Types
// ============================================================================

export type ChatMode = 'policy-qa' | 'author-assist' | 'general-help';
export type UserRole = 'User' | 'Author' | 'Manager' | 'Admin';

// ── Request ──

export interface PolicyContext {
  id: number;
  title: string;
  summary: string;
  keyPoints: string[];
  category: string;
  complianceRisk: string;
  effectiveDate: string;
  status: string;
}

export interface ChatRequest {
  message: string;
  conversationHistory: { role: 'user' | 'assistant'; content: string }[];
  mode: ChatMode;
  policyContext?: { policies: PolicyContext[] };
  userRole: UserRole;
  maxTokens?: number;
}

// ── Response ──

export interface ChatCitation {
  policyId: number;
  title: string;
  excerpt: string;
}

export interface ChatSuggestedAction {
  type: 'navigate' | 'search';
  label: string;
  url: string;
}

export interface ChatResponse {
  message: string;
  citations?: ChatCitation[];
  suggestedActions?: ChatSuggestedAction[];
  metadata: {
    model: string;
    tokensUsed: number;
    processingTimeMs: number;
  };
}

// ── Validation limits ──

export const LIMITS = {
  MAX_MESSAGE_LENGTH: 2000,
  MAX_HISTORY_MESSAGES: 10,
  MAX_POLICY_CONTEXT: 5,
  MAX_TOKENS_DEFAULT: 1000,
  MAX_TOKENS_CEILING: 2000,
  MAX_CONTEXT_CHARS: 20000,
} as const;
