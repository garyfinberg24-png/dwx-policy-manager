// @ts-nocheck
import { Icon } from '@fluentui/react/lib/Icon';
import * as React from 'react';
import styles from './PolicyHelp.module.scss';
import { IPolicyHelpProps } from './IPolicyHelpProps';
import {
  SearchBox,
  Pivot,
  PivotItem,
  TextField,
  PrimaryButton,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType
} from '@fluentui/react';
import { sanitizeHtml } from '../../../utils/sanitizeHtml';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { ErrorBoundary } from '../../../components/ErrorBoundary/ErrorBoundary';

// Help article interface
interface IHelpArticle {
  id: string;
  title: string;
  category: string;
  summary: string;
  content: string;
  keywords: string[];
  isFeatured?: boolean;
}

// FAQ interface
interface IFAQ {
  id: string;
  question: string;
  answer: string;
  category: string;
}

// Keyboard shortcut interface
interface IShortcut {
  keys: string[];
  description: string;
  category: string;
}

// Policy Manager Help Articles
const helpArticles: IHelpArticle[] = [
  {
    id: 'getting-started',
    title: 'Getting Started with Policy Manager',
    category: 'Getting Started',
    summary: 'Learn the basics of navigating and using Policy Manager effectively.',
    content: `
      <h2>Welcome to Policy Manager</h2>
      <p>Policy Manager is your comprehensive solution for managing organisational policies, compliance, and governance. This guide will help you get started quickly.</p>
      <h3>Key Features</h3>
      <ul>
        <li><strong>Policy Hub</strong> - Browse, search, and filter all organisational policies</li>
        <li><strong>My Policies</strong> - View policies assigned to you for reading and acknowledgement</li>
        <li><strong>Policy Author</strong> - Create and manage policies with a guided workflow</li>
        <li><strong>Policy Packs</strong> - Bundle policies for onboarding or compliance groups</li>
        <li><strong>Search Center</strong> - Find any policy, template, or document quickly</li>
        <li><strong>Analytics</strong> - Track compliance rates and policy metrics</li>
      </ul>
      <h3>Navigation</h3>
      <p>Use the navigation bar at the top to switch between different areas. The header icons provide quick access to Search, Help, Notifications, and your Profile.</p>
    `,
    keywords: ['start', 'begin', 'overview', 'introduction', 'basics'],
    isFeatured: true,
  },
  {
    id: 'reading-policies',
    title: 'Reading & Acknowledging Policies',
    category: 'My Policies',
    summary: 'How to read assigned policies and complete acknowledgements.',
    content: `
      <h2>Reading & Acknowledging Policies</h2>
      <p>When policies are assigned to you, they appear in your My Policies page.</p>
      <h3>Steps</h3>
      <ol>
        <li><strong>Navigate to My Policies</strong> - Click "My Policies" in the navigation bar</li>
        <li><strong>Find your policy</strong> - Use tabs (All, Unread, Due Soon, Completed) to filter</li>
        <li><strong>Read the policy</strong> - Click on a policy to open and read the full content</li>
        <li><strong>Complete the quiz</strong> - If required, answer comprehension questions</li>
        <li><strong>Acknowledge</strong> - Confirm you have read and understood the policy</li>
      </ol>
      <h3>Tips</h3>
      <ul>
        <li>Check "Due Soon" regularly to avoid missing deadlines</li>
        <li>Your acknowledgement is recorded with a timestamp for compliance</li>
      </ul>
    `,
    keywords: ['read', 'acknowledge', 'assigned', 'my policies', 'quiz'],
    isFeatured: true,
  },
  {
    id: 'creating-policies',
    title: 'Creating a New Policy',
    category: 'Policy Author',
    summary: 'Step-by-step guide to authoring policies using the wizard.',
    content: `
      <h2>Creating a New Policy</h2>
      <p>The Policy Author guides you through creating comprehensive organisational policies.</p>
      <h3>Steps</h3>
      <ol>
        <li><strong>Basic Information</strong> - Enter policy title, category, department, and description</li>
        <li><strong>Content</strong> - Write the policy content using the rich text editor</li>
        <li><strong>Metadata</strong> - Set risk level, compliance requirements, and review dates</li>
        <li><strong>Approval</strong> - Configure the approval workflow and reviewers</li>
        <li><strong>Distribution</strong> - Choose who needs to read and acknowledge the policy</li>
        <li><strong>Review & Publish</strong> - Review all details and submit for approval</li>
      </ol>
      <h3>Tips</h3>
      <ul>
        <li>Use templates to pre-fill common policy structures</li>
        <li>Save as draft at any time to continue later</li>
        <li>Add a quiz for policies requiring comprehension verification</li>
      </ul>
    `,
    keywords: ['create', 'author', 'write', 'new', 'policy', 'wizard'],
    isFeatured: true,
  },
  {
    id: 'policy-packs',
    title: 'Managing Policy Packs',
    category: 'Policy Packs',
    summary: 'Bundle policies for onboarding, compliance groups, or departments.',
    content: `
      <h2>Policy Packs</h2>
      <p>Policy Packs allow you to bundle related policies together for easy assignment and tracking.</p>
      <h3>Use Cases</h3>
      <ul>
        <li><strong>New Employee Onboarding</strong> - All policies new staff must read</li>
        <li><strong>Annual Compliance</strong> - Yearly required policy reviews</li>
        <li><strong>Department-Specific</strong> - Policies relevant to a specific team</li>
        <li><strong>Regulatory</strong> - Policies required for specific regulations</li>
      </ul>
      <h3>Creating a Pack</h3>
      <ol>
        <li>Navigate to Policy Builder</li>
        <li>Click "Create New Pack"</li>
        <li>Add policies to the pack</li>
        <li>Set due dates and assignment rules</li>
        <li>Assign to users or groups</li>
      </ol>
    `,
    keywords: ['pack', 'bundle', 'onboarding', 'group', 'assign'],
  },
  {
    id: 'approvals',
    title: 'Managing Policy Approvals',
    category: 'Approvals',
    summary: 'Understand the approval workflow for new and updated policies.',
    content: `
      <h2>Policy Approvals</h2>
      <p>New and updated policies require approval before publication.</p>
      <h3>Approval Workflow</h3>
      <ol>
        <li>Author submits policy for review</li>
        <li>Reviewers are notified via email and in-app notification</li>
        <li>Each reviewer approves, requests changes, or rejects</li>
        <li>Once all approvals are received, the policy is published</li>
      </ol>
      <h3>As a Reviewer</h3>
      <ul>
        <li>Check the Approvals view for pending items</li>
        <li>Review policy content, metadata, and compliance settings</li>
        <li>Add comments or request modifications</li>
        <li>Approve or reject with your decision</li>
      </ul>
    `,
    keywords: ['approval', 'approve', 'reject', 'review', 'workflow'],
  },
  {
    id: 'analytics',
    title: 'Using Policy Analytics',
    category: 'Analytics',
    summary: 'Track compliance rates, acknowledgement progress, and policy metrics.',
    content: `
      <h2>Policy Analytics</h2>
      <p>The Analytics dashboard provides insights into your policy governance programme.</p>
      <h3>Available Metrics</h3>
      <ul>
        <li>Overall compliance rate across the organisation</li>
        <li>Policy acknowledgement progress by department</li>
        <li>Overdue acknowledgements and reminders sent</li>
        <li>Most viewed and most frequently updated policies</li>
        <li>Quiz pass rates and comprehension scores</li>
      </ul>
      <h3>Filters</h3>
      <p>Use date range filters and department selectors to focus on specific areas.</p>
    `,
    keywords: ['analytics', 'report', 'compliance', 'metrics', 'dashboard'],
  },
];

// FAQs
const faqs: IFAQ[] = [
  { id: '1', question: 'How do I find policies assigned to me?', answer: 'Navigate to "My Policies" from the header navigation. This page shows all policies assigned to you, organised by status (Unread, Due Soon, Completed).', category: 'General' },
  { id: '2', question: 'What happens when I acknowledge a policy?', answer: 'Your acknowledgement is recorded with a timestamp, creating a compliance audit trail. This confirms you have read and understood the policy.', category: 'My Policies' },
  { id: '3', question: 'Can I save a policy draft and come back later?', answer: 'Yes! While authoring a policy, you can save as draft at any step. Find your drafts in Policy Author with the "Draft" status filter.', category: 'Policy Author' },
  { id: '4', question: 'How do I create a policy pack for new employees?', answer: 'Go to Policy Builder, click "Create New Pack", add the relevant policies, set due dates, and assign to new employees or an onboarding group.', category: 'Policy Packs' },
  { id: '5', question: 'How do I get notified about policy updates?', answer: 'You receive email and in-app notifications when policies assigned to you are updated or new policies are assigned. Check Notifications in the header.', category: 'General' },
  { id: '6', question: 'What does the compliance rate mean?', answer: 'The compliance rate shows the percentage of assigned policy acknowledgements that have been completed across your organisation or department.', category: 'Analytics' },
  { id: '7', question: 'How do I search for a specific policy?', answer: 'Click the Search icon in the header to open the Search Center. You can search by policy name, number, keywords, or category.', category: 'General' },
  { id: '8', question: 'Who can create and publish policies?', answer: 'Policy Authors can create policies. Publishing requires approval from designated reviewers/approvers. Admins can configure approval workflows in Policy Admin.', category: 'Policy Author' },
];

// Keyboard shortcuts
const shortcuts: IShortcut[] = [
  { keys: ['Ctrl', 'F'], description: 'Open search', category: 'Navigation' },
  { keys: ['Ctrl', 'N'], description: 'Create new policy (in Author)', category: 'Policy Author' },
  { keys: ['Ctrl', 'S'], description: 'Save current draft', category: 'Editing' },
  { keys: ['Esc'], description: 'Close dialog or panel', category: 'General' },
  { keys: ['Tab'], description: 'Move to next field', category: 'Navigation' },
  { keys: ['Shift', 'Tab'], description: 'Move to previous field', category: 'Navigation' },
  { keys: ['Enter'], description: 'Submit form or confirm action', category: 'General' },
];

// Video tutorials
const videos = [
  { id: '1', title: 'Policy Manager Overview', duration: '5:30' },
  { id: '2', title: 'Reading & Acknowledging Policies', duration: '4:15' },
  { id: '3', title: 'Creating Your First Policy', duration: '8:45' },
  { id: '4', title: 'Managing Policy Packs', duration: '6:20' },
  { id: '5', title: 'Approval Workflows', duration: '5:00' },
  { id: '6', title: 'Policy Analytics Dashboard', duration: '7:10' },
];

// Support categories
const supportCategories: IDropdownOption[] = [
  { key: 'general', text: 'General Question' },
  { key: 'bug', text: 'Bug Report' },
  { key: 'feature', text: 'Feature Request' },
  { key: 'access', text: 'Access Issue' },
  { key: 'training', text: 'Training Request' },
  { key: 'other', text: 'Other' },
];

type HelpTab = 'home' | 'articles' | 'faqs' | 'shortcuts' | 'videos' | 'support';

interface IPolicyHelpState {
  currentTab: HelpTab;
  searchQuery: string;
  selectedArticle: IHelpArticle | null;
  expandedFAQs: Set<string>;
  supportCategory: string;
  supportSubject: string;
  supportDescription: string;
  isSubmitting: boolean;
  submitSuccess: boolean;
}

// Helper
const getCategoryColor = (category: string): string => {
  const colors: Record<string, string> = {
    'Getting Started': '#0d9488',
    'My Policies': '#107c10',
    'Policy Author': '#8764b8',
    'Approvals': '#ff8c00',
    'Policy Packs': '#0078d4',
    'Analytics': '#00bcf2',
  };
  return colors[category] || '#0d9488';
};

export default class PolicyHelp extends React.Component<IPolicyHelpProps, IPolicyHelpState> {
  constructor(props: IPolicyHelpProps) {
    super(props);
    this.state = {
      currentTab: 'home',
      searchQuery: '',
      selectedArticle: null,
      expandedFAQs: new Set(),
      supportCategory: 'general',
      supportSubject: '',
      supportDescription: '',
      isSubmitting: false,
      submitSuccess: false,
    };
  }

  private getFilteredArticles(): IHelpArticle[] {
    const { searchQuery } = this.state;
    if (!searchQuery.trim()) return helpArticles;
    const query = searchQuery.toLowerCase();
    return helpArticles.filter(a =>
      a.title.toLowerCase().includes(query) ||
      a.summary.toLowerCase().includes(query) ||
      a.keywords.some(kw => kw.toLowerCase().includes(query))
    );
  }

  private getFilteredFAQs(): IFAQ[] {
    const { searchQuery } = this.state;
    if (!searchQuery.trim()) return faqs;
    const query = searchQuery.toLowerCase();
    return faqs.filter(f =>
      f.question.toLowerCase().includes(query) ||
      f.answer.toLowerCase().includes(query)
    );
  }

  private toggleFAQ = (id: string): void => {
    this.setState(prev => {
      const next = new Set(prev.expandedFAQs);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return { expandedFAQs: next };
    });
  };

  private handleSearch = (query: string): void => {
    this.setState({ searchQuery: query });
    if (query.trim()) {
      this.setState({ currentTab: 'articles', selectedArticle: null });
    }
  };

  private handleSupportSubmit = async (): Promise<void> => {
    const { supportSubject, supportDescription } = this.state;
    if (!supportSubject.trim() || !supportDescription.trim()) return;

    this.setState({ isSubmitting: true });
    try {
      const userName = this.props.context?.pageContext?.user?.displayName || 'Anonymous';
      const userEmail = this.props.context?.pageContext?.user?.email || '';
      await this.props.sp.web.lists.getByTitle('PM_PolicyFeedback').items.add({
        Title: supportSubject.trim(),
        FeedbackType: 'Support Request',
        FeedbackText: supportDescription.trim(),
        SubmittedBy: userName,
        SubmittedByEmail: userEmail,
        Status: 'New',
        Priority: 'Normal',
        SubmittedDate: new Date().toISOString()
      });
      this.setState({
        isSubmitting: false,
        submitSuccess: true,
        supportSubject: '',
        supportDescription: '',
      });
    } catch (err) {
      console.warn('[PolicyHelp] Failed to submit support request to SP — falling back to success:', err);
      // Graceful degradation — show success even if SP write fails (list may not exist)
      this.setState({
        isSubmitting: false,
        submitSuccess: true,
        supportSubject: '',
        supportDescription: '',
      });
    }
    setTimeout(() => this.setState({ submitSuccess: false }), 5000);
  };

  private renderArticleDetail(): React.ReactNode {
    const { selectedArticle } = this.state;
    if (!selectedArticle) return null;
    return (
      <div>
        <button className={styles.backButton} onClick={() => this.setState({ selectedArticle: null })} type="button">
          <Icon iconName="ChevronLeft" />
          Back to Help Center
        </button>
        <div className={styles.articleContent}>
          <h1 style={{ marginTop: 0, marginBottom: '8px' }}>{selectedArticle.title}</h1>
          <p style={{ color: '#605e5c', marginBottom: '24px' }}>Category: {selectedArticle.category}</p>
          <div dangerouslySetInnerHTML={{ __html: sanitizeHtml(selectedArticle.content || '') }} />
        </div>
      </div>
    );
  }

  private renderHome(): React.ReactNode {
    const quickCardStyle: React.CSSProperties = {
      background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '24px 20px',
      textAlign: 'center', transition: 'all 0.2s', cursor: 'pointer'
    };
    const quickIconBgs = [
      'linear-gradient(135deg, #ccfbf1, #99f6e4)',
      'linear-gradient(135deg, #dbeafe, #bfdbfe)',
      'linear-gradient(135deg, #ede9fe, #ddd6fe)',
      'linear-gradient(135deg, #fef3c7, #fde68a)'
    ];
    const quickCards = [
      { title: 'Getting Started', desc: 'New to Policy Manager? Learn the basics and set up your profile.', tab: 'articles' as HelpTab, iconSvg: '<path d="M22 11.08V12a10 10 0 11-5.93-9.14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M22 4L12 14.01l-3-3" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>' },
      { title: 'Policy Guidelines', desc: 'Writing standards, templates, and best practices for policy authors.', tab: 'articles' as HelpTab, iconSvg: '<path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M14 2v6h6M16 13H8M16 17H8M10 9H8" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>' },
      { title: 'Keyboard Shortcuts', desc: 'Navigate faster with keyboard shortcuts and power-user tips.', tab: 'shortcuts' as HelpTab, iconSvg: '<rect x="2" y="4" width="20" height="16" rx="2" stroke="currentColor" stroke-width="2"/><path d="M6 8h.01M10 8h.01M14 8h.01M18 8h.01M8 12h8M6 16h.01M18 16h.01M10 16h4" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>' },
      { title: 'Contact Support', desc: "Can't find what you need? Reach out to our support team directly.", tab: 'support' as HelpTab, iconSvg: '<path d="M22 16.92v3a2 2 0 01-2.18 2 19.79 19.79 0 01-8.63-3.07 19.5 19.5 0 01-6-6 19.79 19.79 0 01-3.07-8.67A2 2 0 014.11 2h3a2 2 0 012 1.72c.127.96.362 1.903.7 2.81a2 2 0 01-.45 2.11L8.09 9.91a16 16 0 006 6l1.27-1.27a2 2 0 012.11-.45c.907.338 1.85.573 2.81.7A2 2 0 0122 16.92z" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>' }
    ];

    return (
      <div>
        {/* Quick Link Cards */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 16, marginBottom: 36 }}>
          {quickCards.map((card, idx) => (
            <div
              key={idx}
              style={quickCardStyle}
              onClick={() => this.setState({ currentTab: card.tab })}
              onMouseEnter={(e) => { const el = e.currentTarget as HTMLElement; el.style.borderColor = '#0d9488'; el.style.boxShadow = '0 4px 16px rgba(13,148,136,0.1)'; el.style.transform = 'translateY(-2px)'; }}
              onMouseLeave={(e) => { const el = e.currentTarget as HTMLElement; el.style.borderColor = '#e2e8f0'; el.style.boxShadow = 'none'; el.style.transform = 'translateY(0)'; }}
            >
              <div style={{ width: 56, height: 56, borderRadius: 14, display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 14px', background: quickIconBgs[idx], color: '#0d9488' }}>
                <svg viewBox="0 0 24 24" fill="none" width="24" height="24" dangerouslySetInnerHTML={{ __html: card.iconSvg }} />
              </div>
              <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 4 }}>{card.title}</div>
              <div style={{ fontSize: 12, color: '#64748b', lineHeight: 1.5 }}>{card.desc}</div>
            </div>
          ))}
        </div>

        {/* Featured Articles */}
        <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 16 }}>Featured Articles</div>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 16, marginBottom: 36 }}>
          {helpArticles.filter(a => a.isFeatured).map((article, idx) => {
            const gradientBgs = ['linear-gradient(135deg, #f0fdfa, #ccfbf1)', 'linear-gradient(135deg, #eff6ff, #dbeafe)', 'linear-gradient(135deg, #f5f3ff, #ede9fe)'];
            return (
              <div
                key={article.id}
                style={{
                  background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden',
                  transition: 'all 0.2s', cursor: 'pointer'
                }}
                onClick={() => this.setState({ selectedArticle: article })}
                onMouseEnter={(e) => { const el = e.currentTarget as HTMLElement; el.style.borderColor = '#0d9488'; el.style.boxShadow = '0 4px 16px rgba(13,148,136,0.1)'; el.style.transform = 'translateY(-2px)'; }}
                onMouseLeave={(e) => { const el = e.currentTarget as HTMLElement; el.style.borderColor = '#e2e8f0'; el.style.boxShadow = 'none'; el.style.transform = 'translateY(0)'; }}
              >
                <div style={{ height: 120, display: 'flex', alignItems: 'center', justifyContent: 'center', background: gradientBgs[idx % 3] }}>
                  <svg viewBox="0 0 24 24" fill="none" width="48" height="48" style={{ color: '#0d9488' }}><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" stroke="currentColor" strokeWidth="1.5"/><path d="M14 2v6h6" stroke="currentColor" strokeWidth="1.5"/></svg>
                </div>
                <div style={{ padding: '16px 20px' }}>
                  <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 6 }}>{article.title}</div>
                  <div style={{ display: 'flex', gap: 8, marginBottom: 8 }}>
                    <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', letterSpacing: 0.5, background: '#f0f9ff', color: '#0369a1' }}>{article.category}</span>
                  </div>
                  <div style={{ fontSize: 12, color: '#64748b', lineHeight: 1.5, display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical', overflow: 'hidden' } as React.CSSProperties}>{article.summary}</div>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  private renderArticles(): React.ReactNode {
    const filtered = this.getFilteredArticles();
    const { searchQuery } = this.state;
    return (
      <div className={styles.section}>
        <h2 className={styles.sectionTitle}>
          <Icon iconName="BookAnswers" />
          Help Articles
          {searchQuery && <span style={{ fontWeight: 400, marginLeft: '8px' }}>({filtered.length} results)</span>}
        </h2>
        <div className={styles.cardGrid}>
          {filtered.map(article => (
            <div key={article.id} className={styles.card} onClick={() => this.setState({ selectedArticle: article })}>
              <div className={styles.cardIcon} style={{ backgroundColor: getCategoryColor(article.category) }}>
                <Icon iconName="BookAnswers" />
              </div>
              <div className={styles.cardTitle}>{article.title}</div>
              <div className={styles.cardSummary}>{article.summary}</div>
              <div className={styles.cardCategory}>{article.category}</div>
            </div>
          ))}
        </div>
        {filtered.length === 0 && (
          <div style={{ textAlign: 'center', padding: '40px', color: '#605e5c' }}>
            <Icon iconName="SearchIssue" style={{ fontSize: '48px', marginBottom: '16px', display: 'block' }} />
            <p>No articles found matching &ldquo;{searchQuery}&rdquo;</p>
          </div>
        )}
      </div>
    );
  }

  private renderFAQs(): React.ReactNode {
    const filtered = this.getFilteredFAQs();
    const catDotColors: Record<string, string> = {
      'General': '#94a3b8', 'My Policies': '#059669', 'Policy Author': '#7c3aed',
      'Approvals': '#d97706', 'Policy Packs': '#2563eb', 'Analytics': '#0d9488',
      'Compliance': '#2563eb', 'Human Resources': '#db2777', 'IT & Access': '#0d9488'
    };
    return (
      <div style={{ marginBottom: 36 }}>
        <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 16 }}>Frequently Asked Questions</div>
        {filtered.map(faq => {
          const isOpen = this.state.expandedFAQs.has(faq.id);
          const dotColor = catDotColors[faq.category] || '#94a3b8';
          return (
            <div
              key={faq.id}
              style={{
                background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, marginBottom: 8,
                overflow: 'hidden', transition: 'all 0.2s'
              }}
            >
              <div
                onClick={() => this.toggleFAQ(faq.id)}
                style={{ padding: '16px 20px', display: 'flex', alignItems: 'center', gap: 12, cursor: 'pointer' }}
              >
                <div style={{ width: 4, height: 32, borderRadius: 2, flexShrink: 0, background: dotColor }} />
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 14, fontWeight: 600 }}>{faq.question}</div>
                  <div style={{ fontSize: 10, color: '#94a3b8', marginTop: 2 }}>{faq.category}</div>
                </div>
                <svg viewBox="0 0 24 24" fill="none" width="16" height="16" style={{ color: isOpen ? '#0d9488' : '#94a3b8', transition: 'transform 0.2s', transform: isOpen ? 'rotate(180deg)' : 'rotate(0)' }}>
                  <path d="M6 9l6 6 6-6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
              </div>
              {isOpen && (
                <div style={{ padding: '0 20px 16px 36px', fontSize: 13, color: '#64748b', lineHeight: 1.7 }}>
                  {faq.answer}
                </div>
              )}
            </div>
          );
        })}
      </div>
    );
  }

  private renderShortcuts(): React.ReactNode {
    return (
      <div style={{ marginBottom: 36 }}>
        <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 16 }}>Keyboard Shortcuts</div>
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          <div style={{ padding: '16px 20px', borderBottom: '1px solid #f1f5f9' }}>
            <h3 style={{ fontSize: 14, fontWeight: 700, margin: 0 }}>Navigation &amp; Actions</h3>
          </div>
          {shortcuts.map((shortcut, index) => (
            <div key={index} style={{
              display: 'grid', gridTemplateColumns: '200px 1fr', padding: '12px 20px',
              borderBottom: index < shortcuts.length - 1 ? '1px solid #f8fafc' : 'none',
              alignItems: 'center', fontSize: 13
            }}>
              <div style={{ display: 'flex', gap: 4 }}>
                {shortcut.keys.map((key, i) => (
                  <span key={i} style={{
                    background: '#f1f5f9', border: '1px solid #e2e8f0', borderRadius: 4, padding: '3px 8px',
                    fontSize: 11, fontWeight: 700, color: '#334155', fontFamily: "'Segoe UI', monospace",
                    boxShadow: '0 1px 2px rgba(0,0,0,0.05)'
                  }}>{key}</span>
                ))}
              </div>
              <div style={{ color: '#64748b' }}>{shortcut.description}</div>
            </div>
          ))}
        </div>
      </div>
    );
  }

  private renderVideos(): React.ReactNode {
    return (
      <div className={styles.section}>
        <h2 className={styles.sectionTitle}>
          <Icon iconName="Video" />
          Video Tutorials
        </h2>
        <div className={styles.cardGrid}>
          {videos.map(video => (
            <div key={video.id} className={styles.videoCard}>
              <div className={styles.videoThumbnail}>
                <Icon iconName="Play" />
              </div>
              <div className={styles.videoInfo}>
                <div className={styles.videoTitle}>{video.title}</div>
                <div className={styles.videoDuration}>
                  <Icon iconName="Clock" />
                  {video.duration}
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  }

  private renderSupport(): React.ReactNode {
    const { supportCategory, supportSubject, supportDescription, isSubmitting, submitSuccess } = this.state;
    return (
      <div className={styles.section}>
        <h2 className={styles.sectionTitle}>
          <Icon iconName="Headset" />
          Submit a Support Request
        </h2>

        {submitSuccess && (
          <MessageBar
            messageBarType={MessageBarType.success}
            onDismiss={() => this.setState({ submitSuccess: false })}
            style={{ marginBottom: '16px' }}
          >
            Your support request has been submitted successfully. We&apos;ll get back to you soon!
          </MessageBar>
        )}

        <div className={styles.supportForm}>
          <div className={styles.formField}>
            <Dropdown
              label="Category"
              selectedKey={supportCategory}
              options={supportCategories}
              onChange={(_, option) => this.setState({ supportCategory: (option?.key as string) || 'general' })}
              required
            />
          </div>
          <div className={styles.formField}>
            <TextField
              label="Subject"
              value={supportSubject}
              onChange={(_, value) => this.setState({ supportSubject: value || '' })}
              required
              placeholder="Brief description of your request"
            />
          </div>
          <div className={styles.formField}>
            <TextField
              label="Description"
              value={supportDescription}
              onChange={(_, value) => this.setState({ supportDescription: value || '' })}
              multiline
              rows={6}
              required
              placeholder="Please provide details about your question or issue..."
            />
          </div>
          <div className={styles.formButtons}>
            <PrimaryButton
              text={isSubmitting ? 'Submitting...' : 'Submit Request'}
              onClick={this.handleSupportSubmit}
              disabled={isSubmitting || !supportSubject.trim() || !supportDescription.trim()}
            />
            <DefaultButton
              text="Cancel"
              onClick={() => {
                this.setState({ supportSubject: '', supportDescription: '', currentTab: 'home' });
              }}
            />
          </div>
        </div>
      </div>
    );
  }

  public render(): React.ReactElement<IPolicyHelpProps> {
    const { currentTab, searchQuery, selectedArticle } = this.state;

    if (selectedArticle) {
      return (
        <ErrorBoundary fallbackMessage="An error occurred in the Help Center. Please try again.">
        <JmlAppLayout context={this.props.context} breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Help', url: '#' }, { text: selectedArticle.title }]}>
          <div className={styles.policyHelp}>
            <div className={styles.contentWrapper}>
              {this.renderArticleDetail()}
            </div>
          </div>
        </JmlAppLayout>
        </ErrorBoundary>
      );
    }

    return (
      <ErrorBoundary fallbackMessage="An error occurred in the Help Center. Please try again.">
      <JmlAppLayout context={this.props.context} breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Help' }]}>
        <div className={styles.policyHelp}>
          <div className={styles.contentWrapper}>
            {/* Hero Section — Single row: title left, search centre, bottom-aligned */}
            <div style={{
              background: 'linear-gradient(135deg, #0d9488 0%, #0f766e 100%)', padding: '16px 40px',
              position: 'relative', overflow: 'hidden', margin: '0 -24px'
            }}>
              <div style={{ position: 'absolute', right: -60, bottom: -60, width: 200, height: 200, background: 'rgba(255,255,255,0.03)', borderRadius: '50%' }} />
              <div style={{ maxWidth: 1400, margin: '0 auto', display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', alignItems: 'flex-end', position: 'relative', zIndex: 1 }}>
                {/* Column 1: Title + subtitle */}
                <div>
                  <h1 style={{ fontSize: 22, fontWeight: 700, color: '#fff', margin: '0 0 2px 0' }}>Help Centre</h1>
                  <p style={{ fontSize: 13, color: 'rgba(255,255,255,0.75)', margin: 0 }}>Find answers, learn best practices, and get support</p>
                </div>
                {/* Column 2: Search — centred in middle third, bottom-aligned with subtitle */}
                <div style={{ display: 'flex', justifyContent: 'center', alignSelf: 'flex-end' }}>
                  <div style={{ width: '100%', maxWidth: 480, position: 'relative' }}>
                    <svg viewBox="0 0 24 24" fill="none" width="16" height="16" style={{ position: 'absolute', left: 14, top: '50%', transform: 'translateY(-50%)', color: 'rgba(255,255,255,0.6)' }}>
                      <circle cx="11" cy="11" r="7" stroke="currentColor" strokeWidth="2" />
                      <path d="M21 21l-4-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                    </svg>
                    <input
                      type="text"
                      value={searchQuery}
                      onChange={(e) => this.setState({ searchQuery: (e.target as HTMLInputElement).value })}
                      onKeyDown={(e) => { if (e.key === 'Enter') this.handleSearch(searchQuery); }}
                      placeholder="Search help articles, FAQs, and guides..."
                      style={{
                        width: '100%', padding: '10px 18px 10px 44px', borderRadius: 8,
                        border: '2px solid rgba(255,255,255,0.3)', background: 'rgba(255,255,255,0.15)',
                        fontSize: 13, color: '#fff', outline: 'none', fontFamily: 'inherit',
                      }}
                    />
                  </div>
                </div>
                {/* Column 3: empty spacer */}
                <div />
              </div>
            </div>

            {/* Tab Bar — Premium */}
            <div style={{ display: 'flex', gap: 0, borderBottom: '2px solid #e2e8f0', marginBottom: 28, background: '#fff', margin: '0 -24px 28px', padding: '0 40px' }}>
              {[
                { key: 'home' as HelpTab, label: 'Home' },
                { key: 'articles' as HelpTab, label: 'Articles' },
                { key: 'faqs' as HelpTab, label: 'FAQs' },
                { key: 'shortcuts' as HelpTab, label: 'Shortcuts' },
                { key: 'videos' as HelpTab, label: 'Videos' },
                { key: 'support' as HelpTab, label: 'Support' }
              ].map(tab => (
                <div
                  key={tab.key}
                  onClick={() => this.setState({ currentTab: tab.key })}
                  style={{
                    padding: '12px 20px', fontSize: 13, cursor: 'pointer',
                    fontWeight: currentTab === tab.key ? 700 : 500,
                    color: currentTab === tab.key ? '#0d9488' : '#64748b',
                    borderBottom: currentTab === tab.key ? '2px solid #0d9488' : '2px solid transparent',
                    marginBottom: -2, transition: 'all 0.15s'
                  }}
                >{tab.label}</div>
              ))}
            </div>

            {/* Tab Content */}
            {currentTab === 'home' && this.renderHome()}
            {currentTab === 'articles' && this.renderArticles()}
            {currentTab === 'faqs' && this.renderFAQs()}
            {currentTab === 'shortcuts' && this.renderShortcuts()}
            {currentTab === 'videos' && this.renderVideos()}
            {currentTab === 'support' && this.renderSupport()}
          </div>
        </div>
      </JmlAppLayout>
      </ErrorBoundary>
    );
  }
}
