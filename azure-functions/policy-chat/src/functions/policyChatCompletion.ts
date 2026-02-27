// ============================================================================
// DWx Policy Manager — AI Chat Completion (Azure Function HTTP Trigger)
// ============================================================================
// Accepts a chat message + policy context (RAG), calls Azure OpenAI GPT-4o,
// returns a structured response with citations and suggested actions.
// ============================================================================

import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { ChatRequest, ChatResponse, LIMITS } from '../types/chatTypes';
import { getSystemPrompt, buildPolicyContextMessage } from '../prompts/systemPrompts';

// ── Rate limiting (in-memory, per Function instance) ──

const rateLimitMap = new Map<string, { count: number; resetAt: number }>();
const RATE_LIMIT_WINDOW_MS = 60_000;
const RATE_LIMIT_MAX = 20;

function checkRateLimit(ip: string): boolean {
  const now = Date.now();
  const entry = rateLimitMap.get(ip);
  if (!entry || now > entry.resetAt) {
    rateLimitMap.set(ip, { count: 1, resetAt: now + RATE_LIMIT_WINDOW_MS });
    return true;
  }
  entry.count++;
  return entry.count <= RATE_LIMIT_MAX;
}

// ── Prompt sanitization ──

function sanitizeInput(text: string): string {
  return text
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<[^>]+>/g, '')
    .replace(/\b(system|assistant|ignore previous|disregard|new instructions?)\b/gi, '[$1]')
    .trim();
}

// ── Validation ──

interface ValidationError {
  status: number;
  message: string;
}

function validateRequest(body: ChatRequest): ValidationError | null {
  if (!body.message || typeof body.message !== 'string') {
    return { status: 400, message: 'Missing or invalid "message" field' };
  }
  if (body.message.length > LIMITS.MAX_MESSAGE_LENGTH) {
    return { status: 400, message: `Message exceeds ${LIMITS.MAX_MESSAGE_LENGTH} characters` };
  }
  if (!['policy-qa', 'author-assist', 'general-help'].includes(body.mode)) {
    return { status: 400, message: 'Invalid "mode". Must be: policy-qa, author-assist, or general-help' };
  }
  if (!['User', 'Author', 'Manager', 'Admin'].includes(body.userRole)) {
    return { status: 400, message: 'Invalid "userRole". Must be: User, Author, Manager, or Admin' };
  }
  if (body.conversationHistory && body.conversationHistory.length > LIMITS.MAX_HISTORY_MESSAGES) {
    return { status: 400, message: `Conversation history exceeds ${LIMITS.MAX_HISTORY_MESSAGES} messages` };
  }
  if (body.policyContext?.policies && body.policyContext.policies.length > LIMITS.MAX_POLICY_CONTEXT) {
    return { status: 400, message: `Policy context exceeds ${LIMITS.MAX_POLICY_CONTEXT} policies` };
  }
  return null;
}

// ── Main handler ──

async function policyChatCompletion(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const startTime = Date.now();
  context.log('Policy Chat — request received');

  // CORS preflight
  if (request.method === 'OPTIONS') {
    return { status: 204, headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'POST, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' } };
  }

  // Rate limit
  const clientIp = request.headers.get('x-forwarded-for') || 'unknown';
  if (!checkRateLimit(clientIp)) {
    return {
      status: 429,
      jsonBody: { error: 'Too many requests. Please wait a moment before trying again.' },
    };
  }

  try {
    // Parse and validate
    const body = (await request.json()) as ChatRequest;
    const validationError = validateRequest(body);
    if (validationError) {
      return { status: validationError.status, jsonBody: { error: validationError.message } };
    }

    // Environment variables
    const endpoint = process.env.AZURE_OPENAI_ENDPOINT;
    const apiKey = process.env.AZURE_OPENAI_API_KEY;
    const deployment = process.env.AZURE_OPENAI_DEPLOYMENT || 'gpt-4o';
    const apiVersion = process.env.AZURE_OPENAI_API_VERSION || '2024-02-15-preview';

    if (!endpoint || !apiKey) {
      context.log('Missing Azure OpenAI configuration');
      return { status: 500, jsonBody: { error: 'AI service not configured. Contact your administrator.' } };
    }

    // Build messages array
    const systemPrompt = getSystemPrompt(body.mode);
    const messages: { role: string; content: string }[] = [
      { role: 'system', content: systemPrompt },
    ];

    // Inject policy context (for policy-qa and author-assist modes)
    if (body.mode !== 'general-help' && body.policyContext?.policies) {
      const contextMsg = buildPolicyContextMessage(body.policyContext.policies);
      messages.push({ role: 'system', content: contextMsg });
    }

    // Conversation history
    if (body.conversationHistory) {
      for (const msg of body.conversationHistory.slice(-LIMITS.MAX_HISTORY_MESSAGES)) {
        messages.push({
          role: msg.role,
          content: sanitizeInput(msg.content).substring(0, LIMITS.MAX_MESSAGE_LENGTH),
        });
      }
    }

    // User message
    messages.push({ role: 'user', content: sanitizeInput(body.message) });

    // Token budget
    const maxTokens = Math.min(body.maxTokens || LIMITS.MAX_TOKENS_DEFAULT, LIMITS.MAX_TOKENS_CEILING);

    // Call Azure OpenAI
    const openAiUrl = `${endpoint}openai/deployments/${deployment}/chat/completions?api-version=${apiVersion}`;

    context.log(`Calling Azure OpenAI: model=${deployment}, mode=${body.mode}, messages=${messages.length}, maxTokens=${maxTokens}`);

    const openAiResponse = await fetch(openAiUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'api-key': apiKey,
      },
      body: JSON.stringify({
        messages,
        temperature: 0.7,
        max_tokens: maxTokens,
        top_p: 0.95,
      }),
    });

    if (!openAiResponse.ok) {
      const errorText = await openAiResponse.text();
      context.log(`Azure OpenAI error: ${openAiResponse.status} — ${errorText}`);
      return {
        status: 502,
        jsonBody: { error: 'AI service temporarily unavailable. Please try again in a moment.' },
      };
    }

    const openAiResult = await openAiResponse.json() as any;
    const rawContent = openAiResult.choices?.[0]?.message?.content || '';
    const tokensUsed = openAiResult.usage?.total_tokens || 0;

    context.log(`Azure OpenAI response: ${tokensUsed} tokens, ${Date.now() - startTime}ms`);

    // Parse structured response
    let parsedResponse: { message: string; citations?: any[]; suggestedActions?: any[] };
    try {
      // Try to extract JSON from the response (may be wrapped in markdown fences)
      const jsonMatch = rawContent.match(/```(?:json)?\s*([\s\S]*?)```/) ||
                        rawContent.match(/(\{[\s\S]*\})/);
      const jsonStr = jsonMatch ? jsonMatch[1].trim() : rawContent;
      parsedResponse = JSON.parse(jsonStr);
    } catch {
      // If JSON parsing fails, treat the entire response as the message
      parsedResponse = {
        message: rawContent,
        citations: [],
        suggestedActions: [],
      };
    }

    // Build response
    const chatResponse: ChatResponse = {
      message: parsedResponse.message || rawContent,
      citations: Array.isArray(parsedResponse.citations) ? parsedResponse.citations : [],
      suggestedActions: Array.isArray(parsedResponse.suggestedActions) ? parsedResponse.suggestedActions : [],
      metadata: {
        model: deployment,
        tokensUsed,
        processingTimeMs: Date.now() - startTime,
      },
    };

    return {
      status: 200,
      jsonBody: chatResponse,
      headers: { 'Content-Type': 'application/json' },
    };

  } catch (error: any) {
    context.log(`Unhandled error: ${error.message || error}`);
    return {
      status: 500,
      jsonBody: { error: 'An unexpected error occurred. Please try again.' },
    };
  }
}

// ── Register ──

app.http('policyChatCompletion', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'function',
  handler: policyChatCompletion,
});
