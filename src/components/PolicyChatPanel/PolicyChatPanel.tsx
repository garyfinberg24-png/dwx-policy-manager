// ============================================================================
// PolicyChatPanel — AI Chat Assistant (Fluent UI Panel)
// ============================================================================
// Three-mode chat assistant: Policy Q&A, Author Assist, General Help.
// Uses PolicyChatService for client-side RAG + Azure Function calls.
// ============================================================================

import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { SPFI } from '@pnp/sp';
import styles from './PolicyChatPanel.module.scss';
import { PolicyChatService, ChatMode, IChatMessageLocal } from '../../services/PolicyChatService';
import { PolicyManagerRole } from '../../services/PolicyRoleService';

// ── Props ──

export interface IPolicyChatPanelProps {
  isOpen: boolean;
  onDismiss: () => void;
  sp: SPFI;
  userRole: PolicyManagerRole;
  userName: string;
}

// ── Mode config ──

interface ModeConfig {
  key: ChatMode;
  label: string;
  minRole?: PolicyManagerRole;
}

const MODES: ModeConfig[] = [
  { key: 'policy-qa', label: 'Policy Q&A' },
  { key: 'author-assist', label: 'Author Assistant', minRole: PolicyManagerRole.Author },
  { key: 'general-help', label: 'Help' },
];

const ROLE_RANK: Record<string, number> = {
  User: 0,
  Author: 1,
  Manager: 2,
  Admin: 3,
};

// ── Component ──

export const PolicyChatPanel: React.FC<IPolicyChatPanelProps> = ({
  isOpen,
  onDismiss,
  sp,
  userRole,
  userName,
}) => {
  // State
  const [messages, setMessages] = React.useState<IChatMessageLocal[]>([]);
  const [inputText, setInputText] = React.useState('');
  const [isLoading, setIsLoading] = React.useState(false);
  const [mode, setMode] = React.useState<ChatMode>('policy-qa');
  const [error, setError] = React.useState<string | null>(null);
  const [isAvailable, setIsAvailable] = React.useState(false);
  const [isInitialized, setIsInitialized] = React.useState(false);

  // Refs
  const chatServiceRef = React.useRef<PolicyChatService | null>(null);
  const threadEndRef = React.useRef<HTMLDivElement>(null);
  const textareaRef = React.useRef<HTMLTextAreaElement>(null);

  // ── Initialize service on first open ──

  React.useEffect(() => {
    if (!isOpen || isInitialized) return;

    const init = async (): Promise<void> => {
      const svc = new PolicyChatService(sp);
      await svc.initialize();
      chatServiceRef.current = svc;
      setIsAvailable(svc.isAvailable());
      setIsInitialized(true);

      // Restore session
      const session = svc.restoreSession();
      if (session) {
        setMessages(session.messages);
        setMode(session.mode);
      }
    };

    init().catch(() => {
      setIsInitialized(true);
      setIsAvailable(false);
    });
  }, [isOpen, isInitialized, sp]);

  // ── Auto-scroll on new messages ──

  React.useEffect(() => {
    threadEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, isLoading]);

  // ── Save session when messages or mode change ──

  React.useEffect(() => {
    if (chatServiceRef.current && messages.length > 0) {
      chatServiceRef.current.saveSession(messages, mode);
    }
  }, [messages, mode]);

  // ── Focus input when panel opens ──

  React.useEffect(() => {
    if (isOpen && isAvailable) {
      setTimeout(() => textareaRef.current?.focus(), 300);
    }
  }, [isOpen, isAvailable]);

  // ── Send message ──

  const handleSend = async (): Promise<void> => {
    const text = inputText.trim();
    if (!text || isLoading || !chatServiceRef.current) return;

    setError(null);

    // Add user message
    const userMessage: IChatMessageLocal = {
      id: `msg_${Date.now()}_user`,
      role: 'user',
      content: text,
      timestamp: new Date(),
    };
    const updatedMessages = [...messages, userMessage];
    setMessages(updatedMessages);
    setInputText('');
    setIsLoading(true);

    // Reset textarea height
    if (textareaRef.current) {
      textareaRef.current.style.height = 'auto';
    }

    try {
      const response = await chatServiceRef.current.sendMessage(
        text,
        mode,
        userRole,
        updatedMessages
      );
      setMessages(prev => [...prev, response]);
    } catch (err: any) {
      setError(err.message || 'Something went wrong. Please try again.');
    } finally {
      setIsLoading(false);
    }
  };

  // ── Key handler ──

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>): void => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  // ── Auto-resize textarea ──

  const handleInputChange = (e: React.ChangeEvent<HTMLTextAreaElement>): void => {
    setInputText(e.target.value);
    const el = e.target;
    el.style.height = 'auto';
    el.style.height = Math.min(el.scrollHeight, 80) + 'px';
  };

  // ── Clear conversation ──

  const handleClear = (): void => {
    setMessages([]);
    setError(null);
    chatServiceRef.current?.clearSession();
  };

  // ── Use suggested prompt ──

  const handleSuggestedPrompt = (prompt: string): void => {
    setInputText(prompt);
    // Auto-send it
    setTimeout(() => {
      setInputText(prompt);
      // Let the effect handle it by focusing and then user can press Enter
      textareaRef.current?.focus();
    }, 50);
  };

  // ── Visible modes (role-filtered) ──

  const visibleModes = MODES.filter(m => {
    if (!m.minRole) return true;
    return (ROLE_RANK[userRole] || 0) >= (ROLE_RANK[m.minRole] || 0);
  });

  // ── User initials ──

  const userInitials = userName
    .split(' ')
    .map(w => w[0])
    .join('')
    .substring(0, 2)
    .toUpperCase();

  // ── Render markdown (simple) ──

  const renderMarkdown = (text: string): React.ReactNode => {
    // Bold
    let html = text.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    // Inline code
    html = html.replace(/`([^`]+)`/g, '<code>$1</code>');
    // Line breaks to <br>
    html = html.replace(/\n/g, '<br/>');
    // Simple bullet lists: lines starting with "- " or "* "
    html = html.replace(/((?:^|\<br\/\>)[\s]*[-*]\s.+(?:\<br\/\>[\s]*[-*]\s.+)*)/g, (match) => {
      const items = match
        .split('<br/>')
        .filter(line => line.trim().match(/^[-*]\s/))
        .map(line => `<li>${line.trim().replace(/^[-*]\s/, '')}</li>`)
        .join('');
      return `<ul>${items}</ul>`;
    });

    return <div className={styles.messageContent} dangerouslySetInnerHTML={{ __html: html }} />;
  };

  // ── Render a single message ──

  const renderMessage = (msg: IChatMessageLocal): React.ReactNode => {
    const isUser = msg.role === 'user';

    return (
      <div
        key={msg.id}
        className={`${styles.messageRow} ${isUser ? styles.messageRowUser : styles.messageRowAssistant}`}
      >
        {/* Avatar */}
        <div className={`${styles.messageAvatar} ${isUser ? styles.avatarUser : styles.avatarAssistant}`}>
          {isUser ? userInitials : (
            <svg viewBox="0 0 24 24" fill="none">
              <path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5"
                stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          )}
        </div>

        {/* Bubble */}
        <div>
          <div className={`${styles.messageBubble} ${isUser ? styles.bubbleUser : styles.bubbleAssistant}`}>
            {isUser ? msg.content : renderMarkdown(msg.content)}
          </div>

          {/* Citations */}
          {!isUser && msg.citations && msg.citations.length > 0 && (
            <div className={styles.citations}>
              {msg.citations.map((cite, i) => (
                <a
                  key={i}
                  className={styles.citationPill}
                  href={`/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${cite.policyId}`}
                  title={cite.excerpt}
                >
                  <svg viewBox="0 0 24 24" fill="none">
                    <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"
                      stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                    <polyline points="14 2 14 8 20 8"
                      stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                  </svg>
                  {cite.title}
                </a>
              ))}
            </div>
          )}

          {/* Suggested Actions */}
          {!isUser && msg.suggestedActions && msg.suggestedActions.length > 0 && (
            <div className={styles.suggestedActions}>
              {msg.suggestedActions.map((action, i) => (
                <a key={i} className={styles.actionChip} href={action.url}>
                  {action.label}
                </a>
              ))}
            </div>
          )}
        </div>
      </div>
    );
  };

  // ── Panel content ──

  const renderPanelContent = (): React.ReactNode => {
    // Not available state
    if (isInitialized && !isAvailable) {
      return (
        <div className={styles.chatPanelContent}>
          <div className={styles.notAvailable}>
            <svg viewBox="0 0 24 24" fill="none">
              <path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z"
                stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
            <p><strong>AI Assistant not configured</strong></p>
            <p>Please contact your administrator to enable the AI Chat Assistant and configure the Function URL in Admin Settings.</p>
          </div>
        </div>
      );
    }

    const suggestedPrompts = chatServiceRef.current?.getSuggestedPrompts(userRole, mode) || [];

    return (
      <div className={styles.chatPanelContent}>
        {/* Mode Tabs */}
        <div className={styles.modeTabs}>
          {visibleModes.map(m => (
            <button
              key={m.key}
              type="button"
              className={`${styles.modeTab} ${mode === m.key ? styles.modeTabActive : ''}`}
              onClick={() => setMode(m.key)}
            >
              {m.label}
            </button>
          ))}
        </div>

        {/* Error Banner */}
        {error && (
          <div className={styles.errorBanner}>
            <svg viewBox="0 0 24 24" fill="none">
              <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="2" />
              <line x1="12" y1="8" x2="12" y2="12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
              <line x1="12" y1="16" x2="12.01" y2="16" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
            </svg>
            <span>{error}</span>
            <button className={styles.errorDismiss} onClick={() => setError(null)} type="button">&times;</button>
          </div>
        )}

        {/* Message Thread */}
        <div className={styles.messageThread}>
          {messages.length === 0 ? (
            // Welcome / Empty state
            <div className={styles.welcomeSection}>
              <div className={styles.welcomeIcon}>
                <svg viewBox="0 0 24 24" fill="none">
                  <path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5"
                    stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
              </div>
              <div className={styles.welcomeTitle}>
                {mode === 'policy-qa' && 'Policy Q&A'}
                {mode === 'author-assist' && 'Author Assistant'}
                {mode === 'general-help' && 'Help & Guidance'}
              </div>
              <div className={styles.welcomeSubtitle}>
                {mode === 'policy-qa' && 'Ask questions about your organization\'s policies. I\'ll search the policy library and provide answers with citations.'}
                {mode === 'author-assist' && 'I can help you draft policy sections, improve clarity, and check for compliance requirements.'}
                {mode === 'general-help' && 'Need help navigating the app? Ask me about features, workflows, or how to accomplish tasks.'}
              </div>
              <div className={styles.suggestedPrompts}>
                {suggestedPrompts.map((prompt, i) => (
                  <button
                    key={i}
                    type="button"
                    className={styles.suggestedPrompt}
                    onClick={() => handleSuggestedPrompt(prompt)}
                  >
                    {prompt}
                  </button>
                ))}
              </div>
            </div>
          ) : (
            // Messages
            <>
              {messages.map(msg => renderMessage(msg))}
              {isLoading && (
                <div className={styles.typingIndicator}>
                  <div className={`${styles.messageAvatar} ${styles.avatarAssistant}`}>
                    <svg viewBox="0 0 24 24" fill="none">
                      <path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5"
                        stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                  </div>
                  <div className={styles.typingDots}>
                    <div className={styles.typingDot} />
                    <div className={styles.typingDot} />
                    <div className={styles.typingDot} />
                  </div>
                </div>
              )}
            </>
          )}
          <div ref={threadEndRef} />
        </div>

        {/* Input Bar */}
        <div className={styles.inputBar}>
          {messages.length > 0 && (
            <button
              type="button"
              className={styles.clearButton}
              onClick={handleClear}
              title="Clear conversation"
            >
              <svg viewBox="0 0 24 24" fill="none">
                <polyline points="3 6 5 6 21 6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                <path d="M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"
                  stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </button>
          )}
          <div className={styles.inputWrapper}>
            <textarea
              ref={textareaRef}
              className={styles.inputField}
              value={inputText}
              onChange={handleInputChange}
              onKeyDown={handleKeyDown}
              placeholder={
                mode === 'policy-qa' ? 'Ask about a policy...' :
                mode === 'author-assist' ? 'Describe what you need help with...' :
                'Ask a question...'
              }
              rows={1}
              disabled={isLoading}
            />
          </div>
          <button
            type="button"
            className={styles.sendButton}
            onClick={handleSend}
            disabled={!inputText.trim() || isLoading}
            title="Send message"
          >
            <svg viewBox="0 0 24 24" fill="none">
              <line x1="22" y1="2" x2="11" y2="13" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              <polygon points="22 2 15 22 11 13 2 9 22 2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          </button>
        </div>
      </div>
    );
  };

  // ── Render ──

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="Policy Assistant"
      isLightDismiss={true}
      closeButtonAriaLabel="Close"
      styles={{
        main: { background: '#f8fafc' },
        scrollableContent: { display: 'flex', flexDirection: 'column', height: '100%' },
        content: { padding: 0, flex: 1, display: 'flex', flexDirection: 'column' },
      }}
    >
      {renderPanelContent()}
    </Panel>
  );
};
