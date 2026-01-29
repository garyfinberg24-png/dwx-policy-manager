// @ts-nocheck
import * as React from 'react';
import styles from './PolicyHelp.module.scss';
import { IPolicyHelpProps } from './IPolicyHelpProps';
import {
  SearchBox,
  Icon,
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
import { JmlAppLayout } from '../../../components/JmlAppLayout';

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
    await new Promise(resolve => setTimeout(resolve, 1500));
    this.setState({
      isSubmitting: false,
      submitSuccess: true,
      supportSubject: '',
      supportDescription: '',
    });
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
          <div dangerouslySetInnerHTML={{ __html: selectedArticle.content }} />
        </div>
      </div>
    );
  }

  private renderHome(): React.ReactNode {
    return (
      <div>
        <div className={styles.section}>
          <h2 className={styles.sectionTitle}>
            <Icon iconName="FavoriteStar" />
            Featured Articles
          </h2>
          <div className={styles.cardGrid}>
            {helpArticles.filter(a => a.isFeatured).map(article => (
              <div key={article.id} className={styles.card} onClick={() => this.setState({ selectedArticle: article })}>
                <div className={styles.cardIcon} style={{ backgroundColor: getCategoryColor(article.category) }}>
                  <Icon iconName="BookAnswers" />
                </div>
                <div className={styles.cardTitle}>{article.title}</div>
                <div className={styles.cardSummary}>{article.summary}</div>
              </div>
            ))}
          </div>
        </div>

        <div className={styles.quickLinks}>
          <div className={styles.quickLink} style={{ background: 'linear-gradient(135deg, #0d9488 0%, #0f766e 100%)' }}
            onClick={() => this.setState({ currentTab: 'articles' })}>
            <Icon iconName="BookAnswers" className={styles.quickLinkIcon} />
            <div className={styles.quickLinkTitle}>Browse All Articles</div>
            <div className={styles.quickLinkText}>Explore our complete knowledge base</div>
          </div>
          <div className={styles.quickLink} style={{ background: 'linear-gradient(135deg, #107c10 0%, #0b5a08 100%)' }}
            onClick={() => this.setState({ currentTab: 'faqs' })}>
            <Icon iconName="Unknown" className={styles.quickLinkIcon} />
            <div className={styles.quickLinkTitle}>FAQs</div>
            <div className={styles.quickLinkText}>Quick answers to common questions</div>
          </div>
          <div className={styles.quickLink} style={{ background: 'linear-gradient(135deg, #8764b8 0%, #6b4fa0 100%)' }}
            onClick={() => this.setState({ currentTab: 'shortcuts' })}>
            <Icon iconName="KeyboardClassic" className={styles.quickLinkIcon} />
            <div className={styles.quickLinkTitle}>Keyboard Shortcuts</div>
            <div className={styles.quickLinkText}>Work faster with shortcuts</div>
          </div>
          <div className={styles.quickLink} style={{ background: 'linear-gradient(135deg, #f7630c 0%, #ca5010 100%)' }}
            onClick={() => this.setState({ currentTab: 'support' })}>
            <Icon iconName="Headset" className={styles.quickLinkIcon} />
            <div className={styles.quickLinkTitle}>Get Support</div>
            <div className={styles.quickLinkText}>Contact our support team</div>
          </div>
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
    return (
      <div className={styles.section}>
        <h2 className={styles.sectionTitle}>
          <Icon iconName="Unknown" />
          Frequently Asked Questions
        </h2>
        {filtered.map(faq => (
          <div key={faq.id} className={styles.faqItem}>
            <div className={styles.faqQuestion} onClick={() => this.toggleFAQ(faq.id)}>
              <span>{faq.question}</span>
              <Icon iconName={this.state.expandedFAQs.has(faq.id) ? 'ChevronUp' : 'ChevronDown'} />
            </div>
            {this.state.expandedFAQs.has(faq.id) && (
              <div className={styles.faqAnswer}>{faq.answer}</div>
            )}
          </div>
        ))}
      </div>
    );
  }

  private renderShortcuts(): React.ReactNode {
    return (
      <div className={styles.section}>
        <h2 className={styles.sectionTitle}>
          <Icon iconName="KeyboardClassic" />
          Keyboard Shortcuts
        </h2>
        <table className={styles.shortcutTable}>
          <thead>
            <tr style={{ backgroundColor: '#f3f2f1' }}>
              <th className={styles.shortcutCell}>Shortcut</th>
              <th className={styles.shortcutCell}>Action</th>
              <th className={styles.shortcutCell}>Category</th>
            </tr>
          </thead>
          <tbody>
            {shortcuts.map((shortcut, index) => (
              <tr key={index} className={styles.shortcutRow}>
                <td className={styles.shortcutCell}>
                  <div className={styles.shortcutKeys}>
                    {shortcut.keys.map((key, i) => (
                      <React.Fragment key={i}>
                        {i > 0 && <span>+</span>}
                        <span className={styles.keyBadge}>{key}</span>
                      </React.Fragment>
                    ))}
                  </div>
                </td>
                <td className={styles.shortcutCell}>{shortcut.description}</td>
                <td className={styles.shortcutCell}>{shortcut.category}</td>
              </tr>
            ))}
          </tbody>
        </table>
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
        <JmlAppLayout context={this.props.context}>
          <div className={styles.policyHelp}>
            <div className={styles.contentWrapper}>
              {this.renderArticleDetail()}
            </div>
          </div>
        </JmlAppLayout>
      );
    }

    return (
      <JmlAppLayout context={this.props.context}>
        <div className={styles.policyHelp}>
          <div className={styles.contentWrapper}>
            {/* Hero Section */}
            <div className={styles.heroSection}>
              <div className={styles.heroHeader}>
                <Icon iconName="Help" className={styles.heroIcon} />
                <div>
                  <h1 className={styles.heroTitle}>Policy Help Center</h1>
                  <p className={styles.heroSubtitle}>
                    Find answers, learn features, and get support
                  </p>
                </div>
              </div>
              <SearchBox
                placeholder="Search for help articles, FAQs, or topics..."
                value={searchQuery}
                onChange={(_, value) => this.setState({ searchQuery: value || '' })}
                onSearch={(value) => this.handleSearch(value)}
                onClear={() => this.setState({ searchQuery: '' })}
                styles={{
                  root: { maxWidth: '600px', backgroundColor: '#ffffff', borderRadius: '6px' },
                }}
              />
            </div>

            {/* Tab Navigation */}
            <Pivot
              selectedKey={currentTab}
              onLinkClick={(item) => this.setState({ currentTab: (item?.props.itemKey as HelpTab) || 'home' })}
              styles={{ root: { marginBottom: '24px' } }}
            >
              <PivotItem headerText="Home" itemKey="home" itemIcon="Home" />
              <PivotItem headerText="Articles" itemKey="articles" itemIcon="BookAnswers" />
              <PivotItem headerText="FAQs" itemKey="faqs" itemIcon="Unknown" />
              <PivotItem headerText="Shortcuts" itemKey="shortcuts" itemIcon="KeyboardClassic" />
              <PivotItem headerText="Videos" itemKey="videos" itemIcon="Video" />
              <PivotItem headerText="Support" itemKey="support" itemIcon="Headset" />
            </Pivot>

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
    );
  }
}
