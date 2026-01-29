// @ts-nocheck
import * as React from 'react';
import styles from './MyPolicies.module.scss';
import { IMyPoliciesProps } from './IMyPoliciesProps';
import {
  Spinner,
  SpinnerSize,
  Icon,
  SearchBox,
  PrimaryButton,
  DefaultButton
} from '@fluentui/react';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { PageSubheader } from '../../../components/PageSubheader';

// Assigned policy interface
interface IAssignedPolicy {
  id: number;
  title: string;
  category: string;
  department: string;
  version: string;
  dueDate: Date | null;
  assignedDate: Date;
  status: 'unread' | 'in-progress' | 'completed' | 'overdue';
  priority: 'high' | 'medium' | 'low';
  packName?: string;
  hasQuiz: boolean;
  acknowledgementDate?: Date;
}

interface IMyPoliciesState {
  loading: boolean;
  policies: IAssignedPolicy[];
  activeTab: 'all' | 'unread' | 'dueSoon' | 'completed';
  searchQuery: string;
  viewMode: 'list' | 'card';
  expandedPolicyId: number | null;
  compliancePercent: number;
}

// Mock assigned policies
const mockPolicies: IAssignedPolicy[] = [
  {
    id: 1, title: 'Information Security Policy', category: 'IT Security', department: 'IT',
    version: '3.2', dueDate: new Date(Date.now() + 2 * 86400000), assignedDate: new Date(Date.now() - 7 * 86400000),
    status: 'unread', priority: 'high', hasQuiz: true
  },
  {
    id: 2, title: 'Code of Conduct', category: 'HR', department: 'Human Resources',
    version: '2.0', dueDate: new Date(Date.now() + 5 * 86400000), assignedDate: new Date(Date.now() - 14 * 86400000),
    status: 'unread', priority: 'medium', packName: 'Annual Compliance Pack', hasQuiz: true
  },
  {
    id: 3, title: 'Data Privacy & Protection Policy', category: 'Compliance', department: 'Legal',
    version: '1.5', dueDate: new Date(Date.now() + 1 * 86400000), assignedDate: new Date(Date.now() - 3 * 86400000),
    status: 'overdue', priority: 'high', hasQuiz: false
  },
  {
    id: 4, title: 'Anti-Bribery & Corruption Policy', category: 'Compliance', department: 'Legal',
    version: '1.2', dueDate: new Date(Date.now() + 10 * 86400000), assignedDate: new Date(Date.now() - 21 * 86400000),
    status: 'in-progress', priority: 'medium', packName: 'Annual Compliance Pack', hasQuiz: true
  },
  {
    id: 5, title: 'Acceptable Use Policy', category: 'IT Security', department: 'IT',
    version: '1.8', dueDate: null, assignedDate: new Date(Date.now() - 30 * 86400000),
    status: 'completed', priority: 'low', hasQuiz: false, acknowledgementDate: new Date(Date.now() - 5 * 86400000)
  },
  {
    id: 6, title: 'Work From Home Policy', category: 'HR', department: 'Human Resources',
    version: '2.1', dueDate: null, assignedDate: new Date(Date.now() - 60 * 86400000),
    status: 'completed', priority: 'low', hasQuiz: false, acknowledgementDate: new Date(Date.now() - 45 * 86400000)
  },
  {
    id: 7, title: 'Health & Safety Policy', category: 'Operations', department: 'Facilities',
    version: '4.0', dueDate: new Date(Date.now() + 15 * 86400000), assignedDate: new Date(Date.now() - 5 * 86400000),
    status: 'unread', priority: 'medium', hasQuiz: true
  },
  {
    id: 8, title: 'Whistleblower Protection Policy', category: 'Governance', department: 'Legal',
    version: '1.0', dueDate: new Date(Date.now() + 20 * 86400000), assignedDate: new Date(Date.now() - 10 * 86400000),
    status: 'unread', priority: 'low', hasQuiz: false
  },
];

// Helpers
const getStatusLabel = (status: string): string => {
  const labels: Record<string, string> = {
    'unread': 'Unread', 'in-progress': 'In Progress', 'completed': 'Completed', 'overdue': 'Overdue'
  };
  return labels[status] || status;
};

const getStatusColor = (status: string): { bg: string; text: string } => {
  const colors: Record<string, { bg: string; text: string }> = {
    'unread': { bg: '#ccfbf1', text: '#0d9488' },
    'in-progress': { bg: '#fff4ce', text: '#9d5d00' },
    'completed': { bg: '#dff6dd', text: '#107c10' },
    'overdue': { bg: '#fde7e9', text: '#d13438' },
  };
  return colors[status] || { bg: '#f3f2f1', text: '#605e5c' };
};

const getPriorityColor = (priority: string): string => {
  const colors: Record<string, string> = { high: '#d13438', medium: '#ff8c00', low: '#107c10' };
  return colors[priority] || '#605e5c';
};

const formatDate = (date: Date | null): string => {
  if (!date) return 'No due date';
  return date.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
};

const getDaysUntilDue = (date: Date | null): number | null => {
  if (!date) return null;
  const diff = date.getTime() - Date.now();
  return Math.ceil(diff / 86400000);
};

export default class MyPolicies extends React.Component<IMyPoliciesProps, IMyPoliciesState> {
  constructor(props: IMyPoliciesProps) {
    super(props);
    this.state = {
      loading: true,
      policies: [],
      activeTab: 'all',
      searchQuery: '',
      viewMode: 'list',
      expandedPolicyId: null,
      compliancePercent: 0,
    };
  }

  public componentDidMount(): void {
    setTimeout(() => {
      const completed = mockPolicies.filter(p => p.status === 'completed').length;
      const total = mockPolicies.length;
      this.setState({
        loading: false,
        policies: mockPolicies,
        compliancePercent: total > 0 ? Math.round((completed / total) * 100) : 0,
      });
    }, 600);
  }

  private getFilteredPolicies(): IAssignedPolicy[] {
    const { policies, activeTab, searchQuery } = this.state;
    let filtered = [...policies];

    switch (activeTab) {
      case 'unread':
        filtered = filtered.filter(p => p.status === 'unread' || p.status === 'overdue');
        break;
      case 'dueSoon':
        filtered = filtered.filter(p => {
          const days = getDaysUntilDue(p.dueDate);
          return days !== null && days <= 7 && p.status !== 'completed';
        });
        break;
      case 'completed':
        filtered = filtered.filter(p => p.status === 'completed');
        break;
    }

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      filtered = filtered.filter(p =>
        p.title.toLowerCase().includes(q) ||
        p.category.toLowerCase().includes(q) ||
        p.department.toLowerCase().includes(q)
      );
    }

    return filtered;
  }

  private handlePolicyClick = (policyId: number): void => {
    window.location.href = `/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policyId}`;
  };

  private toggleExpand = (policyId: number): void => {
    this.setState(prev => ({
      expandedPolicyId: prev.expandedPolicyId === policyId ? null : policyId
    }));
  };

  private renderProgressHeader(): React.ReactNode {
    const { policies, compliancePercent } = this.state;
    const total = policies.length;
    const completed = policies.filter(p => p.status === 'completed').length;
    const unread = policies.filter(p => p.status === 'unread' || p.status === 'overdue').length;
    const dueSoon = policies.filter(p => {
      const days = getDaysUntilDue(p.dueDate);
      return days !== null && days <= 7 && p.status !== 'completed';
    }).length;

    const strokeDashoffset = 264 - (264 * compliancePercent / 100);

    return (
      <div className={styles.progressHeader}>
        <div className={styles.progressHeaderLeft}>
          <div className={styles.progressRingContainer}>
            <svg className={styles.progressRingSvg} viewBox="0 0 100 100">
              <circle className={styles.progressRingBg} cx="50" cy="50" r="42" />
              <circle
                className={styles.progressRingFill}
                cx="50" cy="50" r="42"
                style={{ strokeDashoffset }}
              />
            </svg>
            <div className={styles.progressRingText}>
              <div className={styles.progressRingPercent}>{compliancePercent}%</div>
              <div className={styles.progressRingLabel}>Complete</div>
            </div>
          </div>
          <div className={styles.progressHeaderInfo}>
            <h3>My Policy Compliance</h3>
            <p>{completed} of {total} policies acknowledged</p>
          </div>
        </div>
        <div className={styles.miniStats}>
          <div className={`${styles.miniStat} ${unread > 0 ? styles.danger : ''}`}>
            <div className={styles.miniStatNumber}>{unread}</div>
            <div className={styles.miniStatLabel}>Unread</div>
          </div>
          <div className={`${styles.miniStat} ${dueSoon > 0 ? styles.warning : ''}`}>
            <div className={styles.miniStatNumber}>{dueSoon}</div>
            <div className={styles.miniStatLabel}>Due Soon</div>
          </div>
          <div className={`${styles.miniStat} ${styles.success}`}>
            <div className={styles.miniStatNumber}>{completed}</div>
            <div className={styles.miniStatLabel}>Done</div>
          </div>
        </div>
      </div>
    );
  }

  private renderTabBar(): React.ReactNode {
    const { activeTab, viewMode, policies } = this.state;
    const unreadCount = policies.filter(p => p.status === 'unread' || p.status === 'overdue').length;
    const dueSoonCount = policies.filter(p => {
      const days = getDaysUntilDue(p.dueDate);
      return days !== null && days <= 7 && p.status !== 'completed';
    }).length;

    const tabs = [
      { key: 'all', label: 'All Assigned', count: policies.length },
      { key: 'unread', label: 'Unread', count: unreadCount },
      { key: 'dueSoon', label: 'Due Soon', count: dueSoonCount },
      { key: 'completed', label: 'Completed', count: policies.filter(p => p.status === 'completed').length },
    ];

    return (
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '12px' }}>
        <div className={styles.tabContainer}>
          {tabs.map(tab => (
            <button
              key={tab.key}
              type="button"
              className={`${styles.tabButton} ${activeTab === tab.key ? styles.activeTab : ''}`}
              onClick={() => this.setState({ activeTab: tab.key as any })}
            >
              {tab.label}
              {tab.count > 0 && <span className={styles.tabBadge}>{tab.count}</span>}
            </button>
          ))}
        </div>
        <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
          <SearchBox
            placeholder="Search policies..."
            value={this.state.searchQuery}
            onChange={(_, v) => this.setState({ searchQuery: v || '' })}
            onClear={() => this.setState({ searchQuery: '' })}
            styles={{ root: { width: 220 } }}
          />
          <div className={styles.viewToggle}>
            <button
              type="button"
              className={`${styles.toggleButton} ${viewMode === 'list' ? styles.activeToggle : ''}`}
              onClick={() => this.setState({ viewMode: 'list' })}
              title="List View"
            >
              <Icon iconName="List" />
            </button>
            <button
              type="button"
              className={`${styles.toggleButton} ${viewMode === 'card' ? styles.activeToggle : ''}`}
              onClick={() => this.setState({ viewMode: 'card' })}
              title="Card View"
            >
              <Icon iconName="GridViewMedium" />
            </button>
          </div>
        </div>
      </div>
    );
  }

  private renderListView(policies: IAssignedPolicy[]): React.ReactNode {
    const { expandedPolicyId } = this.state;

    if (policies.length === 0) {
      return this.renderEmptyState();
    }

    return (
      <div className={styles.policyListView}>
        {policies.map(policy => {
          const days = getDaysUntilDue(policy.dueDate);
          const isExpanded = expandedPolicyId === policy.id;
          const statusColor = getStatusColor(policy.status);
          const itemClass = policy.status === 'overdue' ? styles.urgentItem :
            (days !== null && days <= 3 && policy.status !== 'completed') ? styles.dueSoonItem : '';

          return (
            <div key={policy.id}>
              <div className={`${styles.policyListItem} ${itemClass}`}>
                <button
                  type="button"
                  onClick={(e) => { e.stopPropagation(); this.toggleExpand(policy.id); }}
                  style={{
                    background: 'none', border: 'none', cursor: 'pointer', padding: '4px',
                    marginRight: '8px', color: '#605e5c', fontSize: '14px'
                  }}
                >
                  <Icon iconName={isExpanded ? 'ChevronDown' : 'ChevronRight'} />
                </button>

                <div className={styles.policyIcon}>
                  <Icon iconName="Library" />
                </div>
                <div className={styles.policyInfo} onClick={() => this.handlePolicyClick(policy.id)} style={{ cursor: 'pointer' }}>
                  <div className={styles.policyName}>{policy.title}</div>
                  <div className={styles.policyMeta}>
                    <span><Icon iconName="Tag" style={{ fontSize: 11, marginRight: 3 }} />{policy.category}</span>
                    <span>&bull;</span>
                    <span><Icon iconName="Org" style={{ fontSize: 11, marginRight: 3 }} />{policy.department}</span>
                    <span>&bull;</span>
                    <span>v{policy.version}</span>
                    {policy.packName && <><span>&bull;</span><span><Icon iconName="Package" style={{ fontSize: 11, marginRight: 3 }} />{policy.packName}</span></>}
                    {policy.hasQuiz && <><span>&bull;</span><span><Icon iconName="Questionnaire" style={{ fontSize: 11, marginRight: 3 }} />Quiz</span></>}
                  </div>
                  <div className={styles.policyMeta} style={{ marginTop: '2px' }}>
                    <span><Icon iconName="Calendar" style={{ fontSize: 11, marginRight: 3 }} />Assigned: {formatDate(policy.assignedDate)}</span>
                    <span>&bull;</span>
                    <span style={{ color: getPriorityColor(policy.priority), fontWeight: 600 }}>
                      <Icon iconName="Flag" style={{ fontSize: 11, marginRight: 3 }} />{policy.priority.charAt(0).toUpperCase() + policy.priority.slice(1)} Priority
                    </span>
                    {days !== null && days > 0 && policy.status !== 'completed' && (
                      <><span>&bull;</span><span style={{ color: days <= 3 ? '#d13438' : '#605e5c' }}>{days} days remaining</span></>
                    )}
                  </div>
                </div>
                <div className={styles.policyStatus}>
                  <span style={{
                    display: 'inline-block', padding: '4px 12px', borderRadius: '16px',
                    fontSize: '12px', fontWeight: 500,
                    backgroundColor: statusColor.bg, color: statusColor.text
                  }}>
                    {getStatusLabel(policy.status)}
                  </span>
                  <div className={styles.dueDate}>
                    {policy.status === 'completed' && policy.acknowledgementDate
                      ? `Acknowledged ${formatDate(policy.acknowledgementDate)}`
                      : policy.dueDate
                        ? `Due ${formatDate(policy.dueDate)}`
                        : 'No due date'
                    }
                  </div>
                </div>
              </div>

              {isExpanded && (
                <div style={{
                  padding: '16px 20px 16px 76px', background: '#fafafa',
                  borderBottom: '1px solid #edebe9', borderLeft: '4px solid #0d9488'
                }}>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '16px' }}>
                    <div style={{ background: '#fff', borderRadius: '8px', padding: '14px', border: '1px solid #edebe9' }}>
                      <div style={{ fontSize: '12px', color: '#605e5c', marginBottom: '8px', fontWeight: 600 }}>
                        <Icon iconName="Info" style={{ marginRight: '6px' }} />Policy Details
                      </div>
                      <div style={{ fontSize: '13px', color: '#323130' }}>
                        <div>Category: {policy.category}</div>
                        <div>Department: {policy.department}</div>
                        <div>Version: {policy.version}</div>
                      </div>
                    </div>
                    <div style={{ background: '#fff', borderRadius: '8px', padding: '14px', border: '1px solid #edebe9' }}>
                      <div style={{ fontSize: '12px', color: '#605e5c', marginBottom: '8px', fontWeight: 600 }}>
                        <Icon iconName="Calendar" style={{ marginRight: '6px' }} />Timeline
                      </div>
                      <div style={{ fontSize: '13px', color: '#323130' }}>
                        <div>Assigned: {formatDate(policy.assignedDate)}</div>
                        <div>Due: {formatDate(policy.dueDate)}</div>
                        {days !== null && days > 0 && <div style={{ color: days <= 3 ? '#d13438' : '#605e5c' }}>{days} days remaining</div>}
                      </div>
                    </div>
                    <div style={{ background: '#fff', borderRadius: '8px', padding: '14px', border: '1px solid #edebe9' }}>
                      <div style={{ fontSize: '12px', color: '#605e5c', marginBottom: '8px', fontWeight: 600 }}>
                        <Icon iconName="Shield" style={{ marginRight: '6px' }} />Requirements
                      </div>
                      <div style={{ fontSize: '13px', color: '#323130' }}>
                        <div>Priority: <span style={{ color: getPriorityColor(policy.priority), fontWeight: 600 }}>{policy.priority}</span></div>
                        <div>Quiz Required: {policy.hasQuiz ? 'Yes' : 'No'}</div>
                        {policy.packName && <div>Pack: {policy.packName}</div>}
                      </div>
                    </div>
                    <div style={{ background: '#fff', borderRadius: '8px', padding: '14px', border: '1px solid #edebe9', display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center', gap: '8px' }}>
                      <PrimaryButton
                        text={policy.status === 'completed' ? 'View Again' : 'Read Policy'}
                        onClick={() => this.handlePolicyClick(policy.id)}
                        styles={{ root: { width: '100%', backgroundColor: '#0d9488', borderColor: '#0d9488' }, rootHovered: { backgroundColor: '#0f766e', borderColor: '#0f766e' }, rootPressed: { backgroundColor: '#115e59', borderColor: '#115e59' } }}
                      />
                      {policy.status !== 'completed' && (
                        <DefaultButton
                          text="Mark as Read"
                          iconProps={{ iconName: 'Accept' }}
                          styles={{ root: { width: '100%', color: '#0d9488', borderColor: '#0d9488' }, rootHovered: { color: '#0f766e', borderColor: '#0f766e', background: '#ccfbf1' } }}
                        />
                      )}
                    </div>
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>
    );
  }

  private renderCardView(policies: IAssignedPolicy[]): React.ReactNode {
    if (policies.length === 0) {
      return this.renderEmptyState();
    }

    return (
      <div className={styles.policyCardGrid}>
        {policies.map(policy => {
          const days = getDaysUntilDue(policy.dueDate);
          const statusColor = getStatusColor(policy.status);
          const borderColor = policy.status === 'overdue' ? '#d13438' :
            (days !== null && days <= 3 && policy.status !== 'completed') ? '#ff8c00' : '#0d9488';

          return (
            <div
              key={policy.id}
              className={styles.policyCard}
              style={{ borderLeftColor: borderColor, cursor: 'pointer' }}
              onClick={() => this.handlePolicyClick(policy.id)}
            >
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '12px' }}>
                <div className={styles.policyTitle}>{policy.title}</div>
                <span style={{
                  padding: '4px 10px', borderRadius: '16px', fontSize: '11px', fontWeight: 500,
                  backgroundColor: statusColor.bg, color: statusColor.text, flexShrink: 0
                }}>
                  {getStatusLabel(policy.status)}
                </span>
              </div>
              <div className={styles.category} style={{ fontSize: '12px', marginBottom: '8px' }}>
                {policy.category} &bull; {policy.department}
              </div>
              <div style={{ fontSize: '12px', color: '#605e5c', display: 'flex', flexDirection: 'column', gap: '4px' }}>
                <div><Icon iconName="Calendar" className={styles.icon} /> Due: {formatDate(policy.dueDate)}</div>
                <div><Icon iconName="PageList" className={styles.icon} /> Version {policy.version}</div>
                {policy.hasQuiz && <div><Icon iconName="Questionnaire" className={styles.icon} /> Quiz required</div>}
                {policy.packName && <div><Icon iconName="Package" className={styles.icon} /> {policy.packName}</div>}
              </div>
              <div className={styles.actions}>
                <PrimaryButton
                  text={policy.status === 'completed' ? 'View' : 'Read'}
                  styles={{ root: { minWidth: 0, padding: '4px 16px', height: '28px', backgroundColor: '#0d9488', borderColor: '#0d9488' }, rootHovered: { backgroundColor: '#0f766e', borderColor: '#0f766e' }, rootPressed: { backgroundColor: '#115e59', borderColor: '#115e59' } }}
                />
              </div>
            </div>
          );
        })}
      </div>
    );
  }

  private renderEmptyState(): React.ReactNode {
    return (
      <div className={styles.emptyState}>
        <Icon iconName="CompletedSolid" className={styles.emptyIcon} />
        <h3 style={{ margin: '16px 0 8px', color: '#323130' }}>All caught up!</h3>
        <p style={{ color: '#605e5c' }}>No policies matching your current filter. Try a different tab or clear search.</p>
      </div>
    );
  }

  public render(): React.ReactElement<IMyPoliciesProps> {
    const { loading, viewMode } = this.state;
    const filteredPolicies = this.getFilteredPolicies();

    return (
      <JmlAppLayout
        context={this.props.context}
        breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'My Policies' }]}
        activeNavKey="my-policies"
      >
        <div className={styles.myPolicies}>
          <div style={{ maxWidth: '1400px', width: '100%', margin: '0 auto', padding: '24px', boxSizing: 'border-box' }}>
            {loading ? (
              <div style={{ padding: '60px', textAlign: 'center' }}>
                <Spinner size={SpinnerSize.large} label="Loading your policies..." />
              </div>
            ) : (
              <>
                {this.renderProgressHeader()}
                {this.renderTabBar()}
                {viewMode === 'list'
                  ? this.renderListView(filteredPolicies)
                  : this.renderCardView(filteredPolicies)
                }
              </>
            )}
          </div>
        </div>
      </JmlAppLayout>
    );
  }
}
