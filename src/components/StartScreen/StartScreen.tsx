// @ts-nocheck
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { PolicyManagerRole } from '../../services/PolicyRoleService';
import { RecentlyViewedService, IRecentlyViewedDisplay } from '../../services/RecentlyViewedService';
import styles from './StartScreen.module.scss';
import { tc } from '../../utils/themeColors';

export interface IStartScreenProps {
  sp: SPFI;
  userName: string;
  userRole: PolicyManagerRole;
  siteUrl: string;
  onDismiss: () => void;
}

interface IStartScreenState {
  recentPolicies: IRecentlyViewedDisplay[];
  stats: {
    published: number;
    acknowledged: number;
    pending: number;
    overdue: number;
    drafts: number;
    teamCompliance: number;
    pendingApprovals: number;
    totalUsers: number;
  };
  loading: boolean;
  lastViewedPolicy: IRecentlyViewedDisplay | null;
}

export class StartScreen extends React.Component<IStartScreenProps, IStartScreenState> {

  constructor(props: IStartScreenProps) {
    super(props);
    const recent = RecentlyViewedService.getRecentlyViewed();
    this.state = {
      recentPolicies: recent,
      stats: { published: 0, acknowledged: 0, pending: 0, overdue: 0, drafts: 0, teamCompliance: 0, pendingApprovals: 0, totalUsers: 0 },
      loading: true,
      lastViewedPolicy: recent.length > 0 ? recent[0] : null
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadStats();
  }

  private async loadStats(): Promise<void> {
    try {
      const { sp, userRole } = this.props;

      // Load policy counts
      const policies = await sp.web.lists.getByTitle('PM_Policies')
        .items.select('Id', 'PolicyStatus', 'RequiresAcknowledgement')
        .top(500)();

      const published = policies.filter((p: any) => p.PolicyStatus === 'Published').length;
      const drafts = policies.filter((p: any) => p.PolicyStatus === 'Draft').length;

      // Load user's acknowledgements
      let acknowledged = 0;
      let pending = 0;
      let overdue = 0;
      try {
        const acks = await sp.web.lists.getByTitle('PM_PolicyAcknowledgements')
          .items.select('Id', 'AckStatus', 'DueDate')
          .top(500)();
        acknowledged = acks.filter((a: any) => a.AckStatus === 'Acknowledged').length;
        pending = acks.filter((a: any) => a.AckStatus === 'Pending').length;
        overdue = acks.filter((a: any) => {
          if (a.AckStatus !== 'Pending') return false;
          if (!a.DueDate) return false;
          return new Date(a.DueDate) < new Date();
        }).length;
      } catch { /* list may not exist */ }

      // Manager stats
      let teamCompliance = 0;
      let pendingApprovals = 0;
      if (userRole === 'Manager' || userRole === 'Admin') {
        try {
          const approvals = await sp.web.lists.getByTitle('PM_Approvals')
            .items.filter("Status eq 'Pending'")
            .select('Id')
            .top(100)();
          pendingApprovals = approvals.length;
        } catch { /* */ }
        teamCompliance = acknowledged + pending > 0 ? Math.round((acknowledged / (acknowledged + pending)) * 100) : 0;
      }

      // Admin stats
      let totalUsers = 0;
      if (userRole === 'Admin') {
        try {
          const users = await sp.web.lists.getByTitle('PM_UserProfiles')
            .items.select('Id').top(1000)();
          totalUsers = users.length;
        } catch { /* */ }
      }

      this.setState({
        stats: { published, acknowledged, pending, overdue, drafts, teamCompliance, pendingApprovals, totalUsers },
        loading: false
      });
    } catch {
      this.setState({ loading: false });
    }
  }

  private getGreeting(): string {
    const hour = new Date().getHours();
    if (hour < 12) return 'Good morning';
    if (hour < 18) return 'Good afternoon';
    return 'Good evening';
  }

  private getDateString(): string {
    return new Date().toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
  }

  private navigate(path: string): void {
    // Dismiss start screen before navigating so it doesn't show again
    sessionStorage.setItem('pm_start_dismissed', 'true');
    window.location.href = `${this.props.siteUrl}/SitePages/${path}`;
  }

  private getTimeAgo(dateStr: string): string {
    const diff = Date.now() - new Date(dateStr).getTime();
    const mins = Math.floor(diff / 60000);
    if (mins < 1) return 'now';
    if (mins < 60) return `${mins}m`;
    const hours = Math.floor(mins / 60);
    if (hours < 24) return `${hours}h`;
    const days = Math.floor(hours / 24);
    return `${days}d`;
  }

  public render(): React.ReactElement {
    const { userName, userRole, onDismiss } = this.props;
    const { recentPolicies, stats, loading, lastViewedPolicy } = this.state;
    const firstName = userName.split(' ')[0];

    // Role-based action cards
    const allActions = [
      // User actions (everyone)
      { key: 'browse', title: 'Browse Policies', desc: 'Explore all published policies in the Policy Hub with search and filters.', icon: 'ViewAll', color: tc.primary, bg: tc.primaryLighter, page: 'PolicyHub.aspx', minRole: 'User' },
      { key: 'my', title: 'My Policies', desc: 'View your assigned policies, pending acknowledgements, and read history.', icon: 'ClipboardList', color: tc.accent, bg: '#eff6ff', page: 'MyPolicies.aspx', minRole: 'User' },
      { key: 'search', title: 'Search', desc: 'Find policies by name, number, keywords, or category.', icon: 'Search', color: '#7c3aed', bg: '#f5f3ff', page: 'PolicySearch.aspx', minRole: 'User' },
      { key: 'help', title: 'Help Centre', desc: 'Browse articles, FAQs, keyboard shortcuts, and contact support.', icon: 'Help', color: tc.warning, bg: '#fffbeb', page: 'PolicyHelp.aspx', minRole: 'User' },
      // Author actions
      { key: 'new', title: 'New Policy', desc: 'Create a new policy from scratch, template, or document upload.', icon: 'Add', color: '#16a34a', bg: '#f0fdf4', page: 'PolicyBuilder.aspx', minRole: 'Author' },
      { key: 'author', title: 'Policy Author', desc: 'Manage your policies, drafts, approvals, and delegations.', icon: 'EditNote', color: '#0284c7', bg: '#e0f2fe', page: 'PolicyAuthor.aspx', minRole: 'Author' },
      { key: 'quiz', title: 'Quiz Builder', desc: 'Create and manage comprehension quizzes for your policies.', icon: 'Education', color: '#7c3aed', bg: '#f5f3ff', page: 'QuizBuilder.aspx', minRole: 'Author' },
      // Manager actions
      { key: 'approvals', title: 'Approvals', desc: 'Review and approve pending policy submissions from your team.', icon: 'CheckboxComposite', color: tc.primary, bg: tc.primaryLighter, page: 'PolicyManagerView.aspx?tab=approvals', minRole: 'Manager', badge: stats.pendingApprovals },
      { key: 'distribution', title: 'Distribution', desc: 'Create and manage policy distribution campaigns.', icon: 'Send', color: '#0284c7', bg: '#e0f2fe', page: 'PolicyDistribution.aspx', minRole: 'Manager' },
      { key: 'analytics', title: 'Analytics', desc: 'Executive dashboards, compliance metrics, and SLA tracking.', icon: 'BarChartVertical', color: '#6d28d9', bg: '#ede9fe', page: 'PolicyAnalytics.aspx', minRole: 'Manager' },
      // Admin actions
      { key: 'admin', title: 'Admin Centre', desc: 'Configure settings, manage users, templates, and system options.', icon: 'Settings', color: '#475569', bg: '#f1f5f9', page: 'PolicyAdmin.aspx', minRole: 'Admin' },
    ];

    const roleLevel: Record<string, number> = { User: 0, Author: 1, Manager: 2, Admin: 3 };
    const currentLevel = roleLevel[userRole] || 0;
    const visibleActions = allActions.filter(a => currentLevel >= (roleLevel[a.minRole] || 0));

    // Role-based glance stats
    const glanceStats: Array<{ label: string; value: number | string; color?: string }> = [
      { label: 'Published Policies', value: stats.published },
      { label: 'Acknowledged', value: stats.acknowledged, color: '#34d399' },
      { label: 'Pending Ack', value: stats.pending, color: stats.pending > 0 ? '#fbbf24' : undefined },
      { label: 'Overdue', value: stats.overdue, color: stats.overdue > 0 ? '#f87171' : undefined },
    ];
    if (userRole === 'Author' || userRole === 'Admin') {
      glanceStats.push({ label: 'My Drafts', value: stats.drafts, color: '#94a3b8' });
    }
    if (userRole === 'Manager' || userRole === 'Admin') {
      glanceStats.splice(1, 0, { label: 'Team Compliance', value: `${stats.teamCompliance}%`, color: tc.primary });
      glanceStats.push({ label: 'Pending Approvals', value: stats.pendingApprovals, color: stats.pendingApprovals > 0 ? tc.warning : undefined });
    }
    if (userRole === 'Admin') {
      glanceStats.push({ label: 'Total Users', value: stats.totalUsers });
    }

    return (
      <div className={styles.startScreen}>
        {/* Sidebar */}
        <div className={styles.sidebar}>
          <div className={styles.sidebarBrand}>
            <svg viewBox="0 0 32 32" fill="none" width="32" height="32">
              <path d="M16 2L4 8v8c0 7.7 5.1 14.9 12 16.9 6.9-2 12-9.2 12-16.9V8L16 2z" stroke="#fff" strokeWidth="2" fill="rgba(255,255,255,0.15)"/>
              <path d="M12 16l3 3 5-6" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
            </svg>
            <div>
              <div className={styles.brandName}>Policy Manager</div>
              <div className={styles.brandSub}>POLICY GOVERNANCE & COMPLIANCE</div>
            </div>
          </div>

          <div className={styles.greeting}>
            <h2>{this.getGreeting()}, {firstName}</h2>
            <div className={styles.greetingDate}>{this.getDateString()}</div>
            <div className={styles.greetingDesc}>Your policy governance command centre. Start fresh or pick up where you left off.</div>
          </div>

          {/* Pick up where you left off */}
          {lastViewedPolicy && (
            <div
              className={styles.pickupCard}
              role="button" tabIndex={0}
              onClick={() => this.navigate(`PolicyDetails.aspx?policyId=${lastViewedPolicy.policyId}`)}
              onKeyDown={(e) => { if (e.key === 'Enter') this.navigate(`PolicyDetails.aspx?policyId=${lastViewedPolicy.policyId}`); }}
            >
              <div className={styles.pickupIcon}>
                <svg viewBox="0 0 24 24" fill="none" width="16" height="16"><path d="M5 12h14M12 5l7 7-7 7" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>
              </div>
              <div className={styles.pickupText}>
                <div className={styles.pickupLabel}>Pick up where you left off</div>
                <div className={styles.pickupPolicy}>{lastViewedPolicy.title} &mdash; {this.getTimeAgo(lastViewedPolicy.viewedAt)}</div>
              </div>
            </div>
          )}

          {/* Recent Policies */}
          {recentPolicies.length > 0 && (
            <div className={styles.sidebarSection}>
              <div className={styles.sectionTitle}>Recent Policies</div>
              {recentPolicies.slice(0, 5).map((policy, i) => (
                <div
                  key={i}
                  className={styles.recentItem}
                  role="button" tabIndex={0}
                  onClick={() => this.navigate(`PolicyDetails.aspx?policyId=${policy.policyId}`)}
                  onKeyDown={(e) => { if (e.key === 'Enter') this.navigate(`PolicyDetails.aspx?policyId=${policy.policyId}`); }}
                >
                  <span className={`${styles.recentDot} ${styles[policy.category?.toLowerCase().includes('hr') ? 'green' : policy.category?.toLowerCase().includes('it') ? 'blue' : 'green']}`} />
                  <span className={styles.recentName}>{policy.title}</span>
                  <span className={styles.recentTime}>{this.getTimeAgo(policy.viewedAt)}</span>
                </div>
              ))}
            </div>
          )}

          {/* At a Glance */}
          <div className={styles.sidebarSection}>
            <div className={styles.sectionTitle}>At a Glance</div>
            {loading ? (
              <Spinner size={SpinnerSize.small} label="Loading..." styles={{ root: { color: '#fff' }, label: { color: 'rgba(255,255,255,0.7)' } }} />
            ) : (
              glanceStats.map((stat, i) => (
                <div key={i} className={styles.glanceCard}>
                  <span className={styles.glanceLabel}>{stat.label}</span>
                  <span className={styles.glanceValue} style={stat.color ? { color: stat.color } : undefined}>{stat.value}</span>
                </div>
              ))
            )}
          </div>

          {/* Skip button */}
          <div
            className={styles.skipBtn}
            role="button" tabIndex={0}
            onClick={onDismiss}
            onKeyDown={(e) => { if (e.key === 'Enter') onDismiss(); }}
          >
            Skip to Policy Hub &rarr;
          </div>
        </div>

        {/* Main Content */}
        <div className={styles.main}>
          <h1>What would you like to do?</h1>
          <p className={styles.mainDesc}>
            Policy Manager gives you everything you need to create, distribute, and track policy compliance across your organisation.
            {userRole === 'Author' && ' As an Author, you can create and manage policies.'}
            {userRole === 'Manager' && ' As a Manager, you can approve policies and track your team\'s compliance.'}
            {userRole === 'Admin' && ' As an Admin, you have full access to all features and system configuration.'}
          </p>
          <p className={styles.mainHint}>Choose an action below to get started, or use the sidebar to jump back to a recent policy.</p>

          <div className={styles.sectionLabel}>
            Quick Actions
            <span className={styles.roleBadge}>{userRole}</span>
          </div>

          <div className={styles.actionsGrid}>
            {visibleActions.map(action => (
              <div
                key={action.key}
                className={styles.actionCard}
                role="button" tabIndex={0}
                onClick={() => this.navigate(action.page)}
                onKeyDown={(e) => { if (e.key === 'Enter') this.navigate(action.page); }}
              >
                <div className={styles.actionIcon} style={{ backgroundColor: action.bg }}>
                  <Icon iconName={action.icon} styles={{ root: { fontSize: 22, color: action.color } }} />
                </div>
                <div className={styles.actionTitle}>
                  {action.title}
                  {action.badge !== undefined && action.badge > 0 && (
                    <span className={styles.actionBadge}>{action.badge}</span>
                  )}
                </div>
                <div className={styles.actionDesc}>{action.desc}</div>
              </div>
            ))}
          </div>

          {/* Compliance Summary */}
          {!loading && (
            <>
              <div className={styles.sectionLabel}>Your Compliance Summary</div>
              <div className={styles.statsBar}>
                <div className={styles.statCard}>
                  <div className={styles.statIcon} style={{ background: '#f0fdf4' }}>
                    <Icon iconName="CompletedSolid" styles={{ root: { fontSize: 18, color: '#16a34a' } }} />
                  </div>
                  <div>
                    <div className={styles.statValue} style={{ color: '#16a34a' }}>{stats.acknowledged}</div>
                    <div className={styles.statLabel}>Acknowledged</div>
                  </div>
                </div>
                <div className={styles.statCard}>
                  <div className={styles.statIcon} style={{ background: '#fef3c7' }}>
                    <Icon iconName="Clock" styles={{ root: { fontSize: 18, color: tc.warning } }} />
                  </div>
                  <div>
                    <div className={styles.statValue} style={{ color: tc.warning }}>{stats.pending}</div>
                    <div className={styles.statLabel}>Pending</div>
                  </div>
                </div>
                <div className={styles.statCard}>
                  <div className={styles.statIcon} style={{ background: '#fee2e2' }}>
                    <Icon iconName="Warning" styles={{ root: { fontSize: 18, color: tc.danger } }} />
                  </div>
                  <div>
                    <div className={styles.statValue} style={{ color: tc.danger }}>{stats.overdue}</div>
                    <div className={styles.statLabel}>Overdue</div>
                  </div>
                </div>
                <div className={styles.statCard}>
                  <div className={styles.statIcon} style={{ background: tc.primaryLighter }}>
                    <Icon iconName="Shield" styles={{ root: { fontSize: 18, color: tc.primary } }} />
                  </div>
                  <div>
                    <div className={styles.statValue} style={{ color: tc.primary }}>
                      {stats.acknowledged + stats.pending > 0 ? Math.round((stats.acknowledged / (stats.acknowledged + stats.pending)) * 100) : 100}%
                    </div>
                    <div className={styles.statLabel}>Compliance</div>
                  </div>
                </div>
              </div>
            </>
          )}
        </div>
      </div>
    );
  }
}

export default StartScreen;
