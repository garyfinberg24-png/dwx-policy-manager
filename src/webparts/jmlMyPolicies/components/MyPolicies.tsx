// @ts-nocheck
import * as React from 'react';
import styles from './MyPolicies.module.scss';
import { IMyPoliciesProps } from './IMyPoliciesProps';
import {
  Spinner,
  SpinnerSize,
} from '@fluentui/react';
import { PanelType } from '@fluentui/react/lib/Panel';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { ErrorBoundary } from '../../../components/ErrorBoundary/ErrorBoundary';
import { StyledPanel } from '../../../components/StyledPanel';
import { PM_LISTS } from '../../../constants/SharePointListNames';
import { tc } from '../../../utils/themeColors';

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
  isSecure?: boolean;
}

interface IMyPoliciesState {
  loading: boolean;
  policies: IAssignedPolicy[];
  activeTab: 'all' | 'pending' | 'overdue' | 'completed';
  searchQuery: string;
  selectedPolicyId: number | null;
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
    'unread': 'Pending', 'in-progress': 'Pending', 'completed': 'Acknowledged', 'overdue': 'Overdue'
  };
  return labels[status] || status;
};

const getStatusBadgeClass = (status: string): { bg: string; text: string } => {
  const colors: Record<string, { bg: string; text: string }> = {
    'unread': { bg: '#fef3c7', text: tc.warning },
    'in-progress': { bg: '#fef3c7', text: tc.warning },
    'completed': { bg: '#dcfce7', text: '#16a34a' },
    'overdue': { bg: '#fee2e2', text: tc.danger },
  };
  return colors[status] || { bg: '#f1f5f9', text: '#64748b' };
};

const getRiskBadge = (priority: string): { bg: string; text: string; label: string } => {
  const map: Record<string, { bg: string; text: string; label: string }> = {
    high: { bg: '#fee2e2', text: tc.danger, label: 'Critical' },
    medium: { bg: '#fef3c7', text: tc.warning, label: 'Medium' },
    low: { bg: '#f1f5f9', text: '#64748b', label: 'Low' },
  };
  return map[priority] || { bg: '#f1f5f9', text: '#64748b', label: priority };
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

// SVG icon components
const DocumentIcon: React.FC<{ size?: number; color?: string }> = ({ size = 18, color = 'currentColor' }) => (
  <svg viewBox="0 0 24 24" fill="none" width={size} height={size}>
    <path d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
  </svg>
);

const CheckIcon: React.FC<{ size?: number; color?: string }> = ({ size = 14, color = 'currentColor' }) => (
  <svg viewBox="0 0 24 24" fill="none" width={size} height={size}>
    <path d="M9 12l2 2 4-4" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
  </svg>
);

const WarningIcon: React.FC<{ size?: number; color?: string }> = ({ size = 18, color = 'currentColor' }) => (
  <svg viewBox="0 0 24 24" fill="none" width={size} height={size}>
    <path d="M12 9v2m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" stroke={color} strokeWidth="2"/>
  </svg>
);

const ClockIcon: React.FC<{ size?: number; color?: string }> = ({ size = 18, color = 'currentColor' }) => (
  <svg viewBox="0 0 24 24" fill="none" width={size} height={size}>
    <circle cx="12" cy="12" r="10" stroke={color} strokeWidth="2"/>
    <path d="M12 6v6l4 2" stroke={color} strokeWidth="2" strokeLinecap="round"/>
  </svg>
);

const ChevronRightIcon: React.FC<{ size?: number; color?: string }> = ({ size = 14, color = 'currentColor' }) => (
  <svg viewBox="0 0 24 24" fill="none" width={size} height={size}>
    <path d="M9 5l7 7-7 7" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
  </svg>
);

const CloseIcon: React.FC<{ size?: number; color?: string }> = ({ size = 16, color = 'currentColor' }) => (
  <svg viewBox="0 0 24 24" fill="none" width={size} height={size}>
    <path d="M6 18L18 6M6 6l12 12" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
  </svg>
);

const ShieldIcon: React.FC<{ size?: number; color?: string }> = ({ size = 16, color = 'currentColor' }) => (
  <svg viewBox="0 0 24 24" fill="none" width={size} height={size}>
    <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" stroke={color} strokeWidth="2"/>
  </svg>
);

const QuizIcon: React.FC<{ size?: number; color?: string }> = ({ size = 16, color = 'currentColor' }) => (
  <svg viewBox="0 0 24 24" fill="none" width={size} height={size}>
    <path d="M12 14l9-5-9-5-9 5 9 5z" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    <path d="M12 14v7" stroke={color} strokeWidth="2" strokeLinecap="round"/>
    <path d="M21 9v7l-9 5-9-5V9" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
  </svg>
);

const SearchIcon: React.FC<{ size?: number; color?: string }> = ({ size = 14, color = 'currentColor' }) => (
  <svg viewBox="0 0 24 24" fill="none" width={size} height={size}>
    <circle cx="11" cy="11" r="8" stroke={color} strokeWidth="2"/>
    <path d="M21 21l-4.35-4.35" stroke={color} strokeWidth="2" strokeLinecap="round"/>
  </svg>
);

const ExclamationIcon: React.FC<{ size?: number; color?: string }> = ({ size = 18, color = 'currentColor' }) => (
  <svg viewBox="0 0 24 24" fill="none" width={size} height={size}>
    <path d="M12 8v4m0 4h.01" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    <path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
  </svg>
);

export default class MyPolicies extends React.Component<IMyPoliciesProps, IMyPoliciesState> {
  private _isMounted = false;

  constructor(props: IMyPoliciesProps) {
    super(props);
    this.state = {
      loading: true,
      policies: [],
      activeTab: 'all',
      searchQuery: '',
      selectedPolicyId: null,
      compliancePercent: 0,
    };
  }

  public async componentDidMount(): Promise<void> {
    this._isMounted = true;
    try {
      await this.loadFromSharePoint();
    } catch (err) {
      console.warn('Could not load from SharePoint, falling back to mock data:', err);
      if (this._isMounted) this.loadMockData();
    }
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
  }

  private async loadFromSharePoint(): Promise<void> {
    const { sp } = this.props;

    // Get current user Id
    const currentUser = await sp.web.currentUser();
    const userId = currentUser.Id;

    // Query acknowledgements assigned to this user
    const items = await sp.web.lists.getByTitle(PM_LISTS.POLICY_ACKNOWLEDGEMENTS)
      .items
      .filter(`AckUserId eq ${userId}`)
      .select(
        'Id', 'PolicyId', 'PolicyName', 'PolicyNumber', 'PolicyCategory',
        'AckStatus', 'AssignedDate', 'DueDate', 'AcknowledgedDate',
        'QuizRequired', 'IsMandatory', 'PolicyVersionNumber'
      )
      .top(500)();

    const policies: IAssignedPolicy[] = items.map((item: any) => {
      const ackStatus: string = item.AckStatus || 'Not Sent';
      let status: IAssignedPolicy['status'] = 'unread';
      if (ackStatus === 'Acknowledged') {
        status = 'completed';
      } else if (ackStatus === 'In Progress' || ackStatus === 'Opened') {
        status = 'in-progress';
      } else if (ackStatus === 'Overdue') {
        status = 'overdue';
      } else {
        // Check if overdue based on DueDate
        if (item.DueDate && new Date(item.DueDate) < new Date()) {
          status = 'overdue';
        }
      }

      const priority: IAssignedPolicy['priority'] = item.IsMandatory ? 'high' :
        (item.DueDate && getDaysUntilDue(new Date(item.DueDate)) !== null &&
         (getDaysUntilDue(new Date(item.DueDate)) as number) <= 7) ? 'medium' : 'low';

      return {
        id: item.PolicyId || item.Id,
        title: item.PolicyName || `Policy ${item.PolicyNumber || item.Id}`,
        category: item.PolicyCategory || 'General',
        department: item.PolicyCategory || 'General',
        version: item.PolicyVersionNumber || '1.0',
        dueDate: item.DueDate ? new Date(item.DueDate) : null,
        assignedDate: item.AssignedDate ? new Date(item.AssignedDate) : new Date(),
        status,
        priority,
        hasQuiz: !!item.QuizRequired,
        acknowledgementDate: item.AcknowledgedDate ? new Date(item.AcknowledgedDate) : undefined
      };
    });

    // Cross-reference with PM_Policies to identify secure policies
    try {
      const policyIds = policies.map(p => p.id).filter(Boolean);
      if (policyIds.length > 0) {
        const chunks = [];
        for (let i = 0; i < policyIds.length; i += 20) chunks.push(policyIds.slice(i, i + 20));
        const securePolicyIds = new Set<number>();
        for (const chunk of chunks) {
          try {
            const filter = chunk.map(id => `Id eq ${id}`).join(' or ');
            const policyItems = await sp.web.lists.getByTitle(PM_LISTS.POLICIES)
              .items.filter(filter).select('Id', 'Visibility').top(chunk.length)();
            for (const pi of policyItems) {
              if (pi.Visibility === 'SecurityGroup' || pi.Visibility === 'Custom') {
                securePolicyIds.add(pi.Id);
              }
            }
          } catch { /* non-blocking */ }
        }
        policies.forEach(p => { if (securePolicyIds.has(p.id)) p.isSecure = true; });
      }
    } catch { /* visibility check non-blocking */ }

    const completed = policies.filter(p => p.status === 'completed').length;
    const total = policies.length;
    if (this._isMounted) {
      this.setState({
        loading: false,
        policies,
        compliancePercent: total > 0 ? Math.round((completed / total) * 100) : 0,
      });
    }
  }

  private loadMockData(): void {
    const completed = mockPolicies.filter(p => p.status === 'completed').length;
    const total = mockPolicies.length;
    this.setState({
      loading: false,
      policies: mockPolicies,
      compliancePercent: total > 0 ? Math.round((completed / total) * 100) : 0,
    });
  }

  private getFilteredPolicies(): IAssignedPolicy[] {
    const { policies, activeTab, searchQuery } = this.state;
    let filtered = [...policies];

    switch (activeTab) {
      case 'pending':
        filtered = filtered.filter(p => p.status === 'unread' || p.status === 'in-progress');
        break;
      case 'overdue':
        filtered = filtered.filter(p => p.status === 'overdue');
        break;
      case 'completed':
        // Only show acknowledged policies in Completed filter
        filtered = filtered.filter(p => p.status === 'completed' && p.acknowledgementDate);
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

    // Default sort: Overdue first, then Pending/Due, then Completed
    const statusOrder: Record<string, number> = { 'overdue': 0, 'unread': 1, 'in-progress': 2, 'completed': 3 };
    filtered.sort((a, b) => {
      const orderA = statusOrder[a.status] ?? 2;
      const orderB = statusOrder[b.status] ?? 2;
      if (orderA !== orderB) return orderA - orderB;
      // Within same status group, sort by due date ascending (soonest first)
      const dateA = a.dueDate ? new Date(a.dueDate).getTime() : Infinity;
      const dateB = b.dueDate ? new Date(b.dueDate).getTime() : Infinity;
      return dateA - dateB;
    });

    return filtered;
  }

  private handlePolicyClick = (policyId: number): void => {
    window.location.href = `/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policyId}`;
  };

  private selectPolicy = (policyId: number): void => {
    this.setState(prev => ({
      selectedPolicyId: prev.selectedPolicyId === policyId ? null : policyId
    }));
  };

  private getKpiCounts(): { assigned: number; acknowledged: number; pending: number; overdue: number } {
    const { policies } = this.state;
    return {
      assigned: policies.length,
      acknowledged: policies.filter(p => p.status === 'completed').length,
      pending: policies.filter(p => p.status === 'unread' || p.status === 'in-progress').length,
      overdue: policies.filter(p => p.status === 'overdue').length,
    };
  }

  private getPolicyIconColors(status: string): { bg: string; color: string } {
    switch (status) {
      case 'completed': return { bg: '#dcfce7', color: tc.success };
      case 'overdue': return { bg: '#fee2e2', color: tc.danger };
      default: return { bg: '#fef3c7', color: tc.warning };
    }
  }

  private getDueText(policy: IAssignedPolicy): { text: string; color: string; bold: boolean } {
    if (policy.status === 'completed' && policy.acknowledgementDate) {
      return { text: `Completed ${formatDate(policy.acknowledgementDate)}`, color: '#94a3b8', bold: false };
    }
    const days = getDaysUntilDue(policy.dueDate);
    if (days === null) return { text: 'No due date', color: '#94a3b8', bold: false };
    if (days < 0) return { text: `${Math.abs(days)} days overdue`, color: tc.danger, bold: true };
    if (days <= 3) return { text: `Due in ${days} days`, color: tc.warning, bold: true };
    if (days <= 7) return { text: `Due in ${days} days`, color: tc.warning, bold: true };
    return { text: `Due ${formatDate(policy.dueDate)}`, color: '#94a3b8', bold: false };
  }

  private getProgressSteps(policy: IAssignedPolicy): Array<{ label: string; state: 'done' | 'current' | 'pending' }> {
    const steps: Array<{ label: string; state: 'done' | 'current' | 'pending' }> = [];

    if (policy.status === 'completed') {
      steps.push({ label: 'Assigned', state: 'done' });
      steps.push({ label: 'Read', state: 'done' });
      steps.push({ label: 'Quiz', state: policy.hasQuiz ? 'done' : 'done' });
      steps.push({ label: 'Acknowledge', state: 'done' });
    } else if (policy.status === 'in-progress') {
      steps.push({ label: 'Assigned', state: 'done' });
      steps.push({ label: 'Read', state: 'current' });
      steps.push({ label: 'Quiz', state: 'pending' });
      steps.push({ label: 'Acknowledge', state: 'pending' });
    } else {
      // unread or overdue
      steps.push({ label: 'Assigned', state: 'done' });
      steps.push({ label: 'Read', state: 'current' });
      steps.push({ label: 'Quiz', state: 'pending' });
      steps.push({ label: 'Acknowledge', state: 'pending' });
    }

    // Remove quiz step if not required
    if (!policy.hasQuiz) {
      steps = steps.filter(s => s.label !== 'Quiz');
    }

    // Add Complete/Certificate step
    steps.push({ label: 'Complete', state: policy.status === 'completed' ? 'done' : 'pending' });

    return steps;
  }

  private renderHeroBanner(): React.ReactNode {
    const kpi = this.getKpiCounts();
    const complianceRate = kpi.assigned > 0 ? Math.round((kpi.acknowledged / kpi.assigned) * 100) : 100;
    const circumference = 2 * Math.PI * 17; // r=17
    const offset = circumference - (complianceRate / 100) * circumference;

    // Greeting based on time of day
    const hour = new Date().getHours();
    const greeting = hour < 12 ? 'Good morning' : hour < 18 ? 'Good afternoon' : 'Good evening';
    const userName = this.props.context?.pageContext?.user?.displayName?.split(' ')[0] || 'there';

    return (
      <div style={{
        background: tc.headerBg,
        padding: '16px 40px', position: 'relative', overflow: 'hidden'
      }}>
        <div style={{ position: 'absolute', right: -60, bottom: -60, width: 200, height: 200, background: 'rgba(255,255,255,0.03)', borderRadius: '50%' }} />
        <div style={{ maxWidth: 1400, margin: '0 auto', display: 'grid', gridTemplateColumns: '1fr auto 1fr', alignItems: 'center', gap: 24, position: 'relative', zIndex: 1 }}>

          {/* Left: Ring + Greeting */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
            {/* Compliance ring */}
            <div style={{ textAlign: 'center' }}>
              <div style={{ position: 'relative', width: 56, height: 56 }}>
                <svg viewBox="0 0 40 40" width="56" height="56">
                  <circle cx="20" cy="20" r="17" fill="none" stroke="rgba(255,255,255,0.15)" strokeWidth="4" />
                  <circle cx="20" cy="20" r="17" fill="none" stroke="#fff" strokeWidth="4" strokeLinecap="round"
                    strokeDasharray={circumference} strokeDashoffset={offset} transform="rotate(-90 20 20)" />
                </svg>
                <div style={{ position: 'absolute', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', fontSize: 14, fontWeight: 700, color: '#fff' }}>{complianceRate}%</div>
              </div>
              <div style={{ fontSize: 8, textTransform: 'uppercase', letterSpacing: 0.5, color: 'rgba(255,255,255,0.5)', marginTop: 2 }}>Compliance</div>
            </div>
            {/* Greeting */}
            <div>
              <h1 style={{ fontSize: 20, fontWeight: 700, color: '#fff', margin: '0 0 2px 0' }}>{greeting}, {userName}</h1>
              <p style={{ fontSize: 12, color: 'rgba(255,255,255,0.7)', margin: 0, alignSelf: 'flex-end' }}>
                {kpi.pending > 0 || kpi.overdue > 0
                  ? `${kpi.pending} pending${kpi.overdue > 0 ? `, ${kpi.overdue} overdue` : ''}`
                  : 'Fully compliant'}
              </p>
            </div>
          </div>

          {/* Center: Search input */}
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <div style={{ position: 'relative', display: 'flex', alignItems: 'center' }}>
              <div style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', pointerEvents: 'none', display: 'flex', alignItems: 'center' }}>
                <SearchIcon size={13} color="rgba(255,255,255,0.6)" />
              </div>
              <style>{`.pm-hero-search::placeholder { color: rgba(255,255,255,0.6) !important; }`}</style>
              <input
                className="pm-hero-search"
                type="text"
                placeholder="Search my policies..."
                value={this.state.searchQuery}
                onChange={(e) => this.setState({ searchQuery: e.target.value })}
                style={{
                  background: 'rgba(255,255,255,0.12)',
                  border: '1px solid rgba(255,255,255,0.25)',
                  borderRadius: 6,
                  padding: '7px 12px 7px 30px',
                  fontSize: 12,
                  width: 280,
                  outline: 'none',
                  fontFamily: 'inherit',
                  color: '#fff',
                }}
                onFocus={(e) => {
                  e.currentTarget.style.background = 'rgba(255,255,255,0.2)';
                  e.currentTarget.style.borderColor = 'rgba(255,255,255,0.5)';
                }}
                onBlur={(e) => {
                  e.currentTarget.style.background = 'rgba(255,255,255,0.12)';
                  e.currentTarget.style.borderColor = 'rgba(255,255,255,0.25)';
                }}
              />
            </div>
          </div>

          {/* Right: KPI mini cards */}
          <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
            {[
              { label: 'Assigned', value: kpi.assigned, color: '#fff' },
              { label: 'Done', value: kpi.acknowledged, color: '#fff' },
              { label: 'Pending', value: kpi.pending, color: '#fbbf24' },
              { label: 'Overdue', value: kpi.overdue, color: '#f87171' },
            ].map(k => (
              <div key={k.label} style={{ background: 'rgba(255,255,255,0.12)', border: '1px solid rgba(255,255,255,0.15)', borderRadius: 4, padding: '8px 14px', textAlign: 'center', minWidth: 70 }}>
                <div style={{ fontSize: 18, fontWeight: 700, color: k.color, lineHeight: 1.1 }}>{k.value}</div>
                <div style={{ fontSize: 9, textTransform: 'uppercase', letterSpacing: 0.8, color: 'rgba(255,255,255,0.6)', marginTop: 2 }}>{k.label}</div>
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  }

  private renderToolbar(): React.ReactNode {
    const { activeTab, searchQuery, policies } = this.state;
    const kpi = this.getKpiCounts();
    const filteredPolicies = this.getFilteredPolicies();

    const tabs = [
      { key: 'all' as const, label: 'All', count: null },
      { key: 'pending' as const, label: 'Pending', count: kpi.pending, countBg: 'rgba(0,0,0,0.06)', countColor: '#334155' },
      { key: 'overdue' as const, label: 'Overdue', count: kpi.overdue, countBg: '#fee2e2', countColor: tc.danger },
      { key: 'completed' as const, label: 'Completed', count: null },
    ];

    return (
      <div style={{
        padding: '12px 20px',
        borderBottom: '1px solid #f1f5f9',
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        background: '#fafafa',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          {/* Search moved to hero banner center column */}
          <div style={{ display: 'flex', gap: '6px' }}>
            {tabs.map(tab => {
              const isActive = activeTab === tab.key;
              return (
                <button
                  key={tab.key}
                  type="button"
                  onClick={() => this.setState({ activeTab: tab.key })}
                  style={{
                    padding: '6px 14px',
                    borderRadius: '6px',
                    fontSize: '12px',
                    fontWeight: 600,
                    border: `1px solid ${isActive ? tc.primary : '#e2e8f0'}`,
                    background: isActive ? tc.primary : '#fff',
                    color: isActive ? '#fff' : '#334155',
                    cursor: 'pointer',
                    fontFamily: 'inherit',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '4px',
                  }}
                >
                  {tab.label}
                  {tab.count !== null && tab.count > 0 && (
                    <span style={{
                      background: isActive ? 'rgba(255,255,255,0.25)' : (tab.countBg || 'rgba(0,0,0,0.06)'),
                      color: isActive ? '#fff' : (tab.countColor || '#334155'),
                      padding: '1px 6px',
                      borderRadius: '8px',
                      fontSize: '10px',
                      marginLeft: '2px',
                    }}>{tab.count}</span>
                  )}
                </button>
              );
            })}
          </div>
        </div>
        <span style={{ fontSize: '11px', color: '#94a3b8' }}>{filteredPolicies.length} policies</span>
      </div>
    );
  }

  private renderPolicyRow(policy: IAssignedPolicy): React.ReactNode {
    const { selectedPolicyId } = this.state;
    const isSelected = selectedPolicyId === policy.id;
    const iconColors = this.getPolicyIconColors(policy.status);
    const statusBadge = getStatusBadgeClass(policy.status);
    const riskBadge = getRiskBadge(policy.priority);
    const dueInfo = this.getDueText(policy);
    const isOverdueRow = policy.status === 'overdue';

    return (
      <div
        key={policy.id}
        onClick={() => this.selectPolicy(policy.id)}
        role="button"
        tabIndex={0}
        onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this.selectPolicy(policy.id); } }}
        style={{
          display: 'grid',
          gridTemplateColumns: '44px minmax(180px, 1fr) 90px 48px 100px 100px 140px 100px 36px',
          gap: 12,
          padding: isSelected || isOverdueRow ? '10px 20px 10px 17px' : '10px 20px',
          borderBottom: '1px solid #f1f5f9',
          cursor: 'pointer',
          transition: 'all 0.1s',
          background: isSelected ? tc.primaryLighter : 'transparent',
          borderLeft: isSelected ? `3px solid ${tc.primary}` : isOverdueRow ? `3px solid ${tc.danger}` : '3px solid transparent',
          alignItems: 'center',
        }}
        onMouseEnter={(e) => { if (!isSelected) e.currentTarget.style.background = tc.primaryLighter; }}
        onMouseLeave={(e) => { if (!isSelected) e.currentTarget.style.background = 'transparent'; }}
      >
        {/* Status icon */}
        <div style={{
          width: 36, height: 36, borderRadius: 8,
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          background: iconColors.bg, color: iconColors.color,
        }}>
          {policy.status === 'completed' ? (
            <CheckIcon size={18} color={iconColors.color} />
          ) : policy.status === 'overdue' ? (
            <ExclamationIcon size={18} color={iconColors.color} />
          ) : (
            <WarningIcon size={18} color={iconColors.color} />
          )}
        </div>

        {/* Policy Name */}
        <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', minWidth: 0 }}>
          {policy.title}
        </div>

        {/* Policy # */}
        <div style={{ fontFamily: "'Cascadia Code', 'Fira Code', 'Consolas', monospace", fontSize: 11, color: tc.primary, fontWeight: 600, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
          POL-{policy.id.toString().padStart(3, '0')}
        </div>

        {/* Version */}
        <div style={{ fontSize: 11, color: '#94a3b8', whiteSpace: 'nowrap' }}>
          v{policy.version}
        </div>

        {/* Due Date */}
        <div style={{ fontSize: 12, color: '#475569', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
          {policy.dueDate ? formatDate(policy.dueDate) : '—'}
        </div>

        {/* Category */}
        <div style={{ fontSize: 12, color: '#475569', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
          {policy.category || '—'}
        </div>

        {/* Status badges */}
        <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
          <span style={{
            fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4,
            textTransform: 'uppercase' as const, letterSpacing: 0.3,
            background: statusBadge.bg, color: statusBadge.text, whiteSpace: 'nowrap',
          }}>{getStatusLabel(policy.status)}</span>
          {policy.priority === 'high' && (
            <span style={{
              fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4,
              textTransform: 'uppercase' as const, letterSpacing: 0.3,
              background: riskBadge.bg, color: riskBadge.text, whiteSpace: 'nowrap',
            }}>{riskBadge.label}</span>
          )}
          {policy.isSecure && (
            <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase' as const, letterSpacing: 0.3, background: '#fef2f2', color: tc.danger, display: 'inline-flex', alignItems: 'center', gap: 2, whiteSpace: 'nowrap' }}>
              <svg viewBox="0 0 24 24" fill="none" width="9" height="9"><rect x="3" y="11" width="18" height="11" rx="2" stroke="currentColor" strokeWidth="2"/><path d="M7 11V7a5 5 0 0110 0v4" stroke="currentColor" strokeWidth="2"/></svg>
              Secure
            </span>
          )}
        </div>

        {/* Due In */}
        <div style={{
          fontSize: 11, color: dueInfo.color,
          fontWeight: dueInfo.bold ? 600 : 400,
          whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
        }}>
          {dueInfo.text}
        </div>

        {/* Eye icon — Read Policy */}
        <button
          type="button"
          onClick={(e) => { e.stopPropagation(); this.selectPolicy(policy.id); }}
          style={{
            width: 32, height: 32, borderRadius: 6,
            border: `1px solid ${isSelected ? tc.primary : '#e2e8f0'}`,
            background: isSelected ? tc.primaryLighter : '#fff',
            cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center',
            color: isSelected ? tc.primary : '#94a3b8', transition: 'all 0.15s',
          }}
          onMouseEnter={(e) => { e.currentTarget.style.borderColor = tc.primary; e.currentTarget.style.color = tc.primary; e.currentTarget.style.background = tc.primaryLighter; }}
          onMouseLeave={(e) => { if (!isSelected) { e.currentTarget.style.borderColor = '#e2e8f0'; e.currentTarget.style.color = '#94a3b8'; e.currentTarget.style.background = '#fff'; } }}
          title="Read policy"
          aria-label={`Read ${policy.title}`}
        >
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
            <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>
          </svg>
        </button>
      </div>
    );
  }

  private renderPolicyList(policies: IAssignedPolicy[]): React.ReactNode {
    if (policies.length === 0) {
      return this.renderEmptyState();
    }

    return (
      <div style={{
        background: '#fff',
        border: '1px solid #e2e8f0',
        borderRadius: '10px',
        overflow: 'hidden',
      }}>
        {this.renderToolbar()}
        {/* Column header */}
        <div style={{
          display: 'grid',
          gridTemplateColumns: '44px minmax(180px, 1fr) 90px 48px 100px 100px 140px 100px 36px',
          gap: 12,
          padding: '8px 20px',
          background: '#f8fafc',
          borderBottom: '1px solid #e2e8f0',
          fontSize: 10,
          textTransform: 'uppercase' as const,
          letterSpacing: 0.5,
          color: '#64748b',
          fontWeight: 600,
          alignItems: 'center',
        }}>
          <div></div>
          <div>Policy Name</div>
          <div>Policy #</div>
          <div>Ver</div>
          <div>Due Date</div>
          <div>Category</div>
          <div>Status</div>
          <div>Due In</div>
          <div></div>
        </div>
        <div>
          {policies.map(policy => this.renderPolicyRow(policy))}
        </div>
      </div>
    );
  }

  private renderDetailPanel(): React.ReactNode {
    const { selectedPolicyId, policies } = this.state;
    const policy = selectedPolicyId !== null ? policies.find(p => p.id === selectedPolicyId) : null;

    if (!policy) return null;

    const statusBadge = getStatusBadgeClass(policy.status);
    const riskBadge = getRiskBadge(policy.priority);
    const dueInfo = this.getDueText(policy);
    const days = getDaysUntilDue(policy.dueDate);
    const steps = this.getProgressSteps(policy);

    // Status banner config
    let bannerBg = '#fef3c7';
    let bannerBorder = '#fbbf24';
    let bannerLabel = 'Acknowledgement Pending';
    let bannerDesc = '';
    let bannerIconColor = tc.warning;

    if (policy.status === 'completed') {
      bannerBg = '#dcfce7';
      bannerBorder = '#86efac';
      bannerLabel = 'Acknowledged';
      bannerDesc = policy.acknowledgementDate ? `Completed on ${formatDate(policy.acknowledgementDate)}` : 'Policy acknowledged';
      bannerIconColor = tc.success;
    } else if (policy.status === 'overdue') {
      bannerBg = '#fee2e2';
      bannerBorder = '#fca5a5';
      bannerLabel = 'Overdue';
      bannerDesc = days !== null ? `${Math.abs(days)} days overdue` : 'Past due date';
      bannerIconColor = tc.danger;
    } else {
      bannerDesc = policy.dueDate ? `Due in ${days} days -- ${formatDate(policy.dueDate)}` : 'No due date set';
    }

    // Related policies (pick 2 others from the list)
    const relatedPolicies = policies.filter(p => p.id !== policy.id).slice(0, 2);

    return (
      <StyledPanel
        isOpen={selectedPolicyId !== null && policy !== undefined}
        onDismiss={() => this.setState({ selectedPolicyId: null })}
        type={PanelType.medium}
        headerText={policy ? policy.title : ''}
        isLightDismiss
      >
        {policy && (
        <div style={{ padding: '0' }}>
          {/* Policy number subtitle */}
          <div style={{ fontSize: '12px', color: tc.primary, marginBottom: '16px' }}>POL-{policy.id.toString().padStart(3, '0')} | Version {policy.version}</div>
          {/* Status Banner */}
          <div style={{
            display: 'flex',
            alignItems: 'center',
            gap: '10px',
            padding: '12px 16px',
            borderRadius: '8px',
            marginBottom: '16px',
            background: bannerBg,
            border: `1px solid ${bannerBorder}`,
          }}>
            <div style={{ flexShrink: 0 }}>
              {policy.status === 'completed' ? (
                <CheckIcon size={22} color={bannerIconColor} />
              ) : policy.status === 'overdue' ? (
                <ExclamationIcon size={22} color={bannerIconColor} />
              ) : (
                <ClockIcon size={22} color={bannerIconColor} />
              )}
            </div>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: '13px', fontWeight: 600, color: '#0f172a' }}>{bannerLabel}</div>
              <div style={{ fontSize: '11px', color: '#64748b', marginTop: '2px' }}>{bannerDesc}</div>
            </div>
          </div>

          {/* Progress Steps */}
          <div style={{ marginBottom: '24px' }}>
            <div style={{
              fontSize: '12px',
              fontWeight: 700,
              textTransform: 'uppercase',
              letterSpacing: '0.8px',
              color: '#64748b',
              marginBottom: '12px',
              paddingBottom: '8px',
              borderBottom: '1px solid #f1f5f9',
            }}>Your Progress</div>
            <div style={{ display: 'flex', gap: 0, margin: '16px 0' }}>
              {steps.map((step, idx) => {
                const dotBg = step.state === 'done' ? tc.success : step.state === 'current' ? tc.primary : '#e2e8f0';
                const dotColor = step.state === 'pending' ? '#94a3b8' : '#fff';
                const dotShadow = step.state === 'current' ? '0 0 0 3px rgba(13,148,136,0.2)' : 'none';
                const lineBg = step.state === 'done' ? tc.success : '#e2e8f0';

                return (
                  <div key={step.label} style={{ flex: 1, textAlign: 'center', position: 'relative' }}>
                    {idx < steps.length - 1 && (
                      <div style={{
                        position: 'absolute',
                        top: '14px',
                        left: '50%',
                        width: '100%',
                        height: '2px',
                        background: lineBg,
                        zIndex: 0,
                      }} />
                    )}
                    <div style={{
                      width: '28px',
                      height: '28px',
                      borderRadius: '50%',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      margin: '0 auto 6px',
                      fontSize: '11px',
                      fontWeight: 700,
                      background: dotBg,
                      color: dotColor,
                      boxShadow: dotShadow,
                      position: 'relative',
                      zIndex: 1,
                    }}>
                      {step.state === 'done' ? (
                        <CheckIcon size={14} color="#fff" />
                      ) : (
                        idx + 1
                      )}
                    </div>
                    <div style={{ fontSize: '9px', color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.5px' }}>{step.label}</div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Policy Details */}
          <div style={{ marginBottom: '24px' }}>
            <div style={{
              fontSize: '12px',
              fontWeight: 700,
              textTransform: 'uppercase',
              letterSpacing: '0.8px',
              color: '#64748b',
              marginBottom: '12px',
              paddingBottom: '8px',
              borderBottom: '1px solid #f1f5f9',
            }}>Policy Details</div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px' }}>
              <div>
                <div style={{ fontSize: '10px', color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Category</div>
                <div style={{ fontSize: '13px', fontWeight: 600, color: '#0f172a', marginTop: '2px' }}>
                  <span style={{ fontSize: '10px', fontWeight: 700, padding: '3px 8px', borderRadius: '4px', background: tc.primaryLight, color: tc.primary }}>{policy.category}</span>
                </div>
              </div>
              <div>
                <div style={{ fontSize: '10px', color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Risk Level</div>
                <div style={{ fontSize: '13px', fontWeight: 600, color: '#0f172a', marginTop: '2px' }}>
                  <span style={{ fontSize: '10px', fontWeight: 700, padding: '3px 8px', borderRadius: '4px', background: riskBadge.bg, color: riskBadge.text }}>{riskBadge.label}</span>
                </div>
              </div>
              <div>
                <div style={{ fontSize: '10px', color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Department</div>
                <div style={{ fontSize: '13px', fontWeight: 600, color: '#0f172a', marginTop: '2px' }}>{policy.department}</div>
              </div>
              <div>
                <div style={{ fontSize: '10px', color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Version</div>
                <div style={{ fontSize: '13px', fontWeight: 600, color: '#0f172a', marginTop: '2px' }}>{policy.version}</div>
              </div>
            </div>
          </div>

          {/* Timeline */}
          <div style={{ marginBottom: '24px' }}>
            <div style={{
              fontSize: '12px',
              fontWeight: 700,
              textTransform: 'uppercase',
              letterSpacing: '0.8px',
              color: '#64748b',
              marginBottom: '12px',
              paddingBottom: '8px',
              borderBottom: '1px solid #f1f5f9',
            }}>Timeline</div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px' }}>
              <div>
                <div style={{ fontSize: '10px', color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Assigned</div>
                <div style={{ fontSize: '13px', fontWeight: 600, color: '#0f172a', marginTop: '2px' }}>{formatDate(policy.assignedDate)}</div>
              </div>
              <div>
                <div style={{ fontSize: '10px', color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Due Date</div>
                <div style={{ fontSize: '13px', fontWeight: 600, color: dueInfo.color, marginTop: '2px' }}>{formatDate(policy.dueDate)}</div>
              </div>
              <div>
                <div style={{ fontSize: '10px', color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Priority</div>
                <div style={{ fontSize: '13px', fontWeight: 600, color: '#0f172a', marginTop: '2px' }}>
                  <span style={{ fontSize: '10px', fontWeight: 700, padding: '3px 8px', borderRadius: '4px', background: riskBadge.bg, color: riskBadge.text }}>
                    {policy.priority.charAt(0).toUpperCase() + policy.priority.slice(1)}
                  </span>
                </div>
              </div>
              <div>
                <div style={{ fontSize: '10px', color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Days Remaining</div>
                <div style={{ fontSize: '18px', fontWeight: 700, color: dueInfo.color, marginTop: '2px' }}>
                  {days !== null ? (days < 0 ? Math.abs(days) : days) : '--'}
                </div>
              </div>
            </div>
          </div>

          {/* Requirements */}
          <div style={{ marginBottom: '24px' }}>
            <div style={{
              fontSize: '12px',
              fontWeight: 700,
              textTransform: 'uppercase',
              letterSpacing: '0.8px',
              color: '#64748b',
              marginBottom: '12px',
              paddingBottom: '8px',
              borderBottom: '1px solid #f1f5f9',
            }}>Requirements</div>

            {/* Read requirement */}
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', padding: '8px 0', borderBottom: '1px solid #f8fafc' }}>
              <div style={{
                width: '28px', height: '28px', borderRadius: '6px',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                flexShrink: 0, background: tc.primaryLighter,
              }}>
                <DocumentIcon size={14} color={tc.primary} />
              </div>
              <div style={{ flex: 1, fontSize: '12px', color: '#334155' }}>Read the full policy document</div>
              <span style={{
                fontSize: '8px', fontWeight: 700, padding: '3px 8px', borderRadius: '4px', textTransform: 'uppercase',
                background: policy.status === 'in-progress' || policy.status === 'completed' ? '#dcfce7' : '#fef3c7',
                color: policy.status === 'in-progress' || policy.status === 'completed' ? '#16a34a' : tc.warning,
              }}>
                {policy.status === 'in-progress' || policy.status === 'completed' ? 'Done' : 'Required'}
              </span>
            </div>

            {/* Quiz requirement */}
            {policy.hasQuiz && (
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', padding: '8px 0', borderBottom: '1px solid #f8fafc' }}>
                <div style={{
                  width: '28px', height: '28px', borderRadius: '6px',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  flexShrink: 0, background: '#ede9fe',
                }}>
                  <QuizIcon size={14} color="#7c3aed" />
                </div>
                <div style={{ flex: 1, fontSize: '12px', color: '#334155' }}>Complete comprehension quiz (75% to pass)</div>
                <span style={{
                  fontSize: '8px', fontWeight: 700, padding: '3px 8px', borderRadius: '4px', textTransform: 'uppercase',
                  background: policy.status === 'completed' ? '#dcfce7' : '#fef3c7',
                  color: policy.status === 'completed' ? '#16a34a' : tc.warning,
                }}>
                  {policy.status === 'completed' ? 'Done' : 'Required'}
                </span>
              </div>
            )}

            {/* Acknowledgement requirement */}
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', padding: '8px 0' }}>
              <div style={{
                width: '28px', height: '28px', borderRadius: '6px',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                flexShrink: 0, background: '#dbeafe',
              }}>
                <ShieldIcon size={14} color={tc.accent} />
              </div>
              <div style={{ flex: 1, fontSize: '12px', color: '#334155' }}>Sign digital acknowledgement</div>
              <span style={{
                fontSize: '8px', fontWeight: 700, padding: '3px 8px', borderRadius: '4px', textTransform: 'uppercase',
                background: policy.status === 'completed' ? '#dcfce7' : '#f1f5f9',
                color: policy.status === 'completed' ? '#16a34a' : '#64748b',
              }}>
                {policy.status === 'completed' ? 'Done' : 'Pending'}
              </span>
            </div>
          </div>

          {/* Related Policies */}
          {relatedPolicies.length > 0 && (
            <div style={{ marginBottom: '24px' }}>
              <div style={{
                fontSize: '12px',
                fontWeight: 700,
                textTransform: 'uppercase',
                letterSpacing: '0.8px',
                color: '#64748b',
                marginBottom: '12px',
                paddingBottom: '8px',
                borderBottom: '1px solid #f1f5f9',
              }}>Related Policies</div>
              {relatedPolicies.map(rp => (
                <div
                  key={rp.id}
                  style={{ display: 'flex', alignItems: 'center', gap: '10px', padding: '8px 0', borderBottom: '1px solid #f8fafc', cursor: 'pointer' }}
                  onClick={() => this.selectPolicy(rp.id)}
                  role="button"
                  tabIndex={0}
                  onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this.selectPolicy(rp.id); } }}
                >
                  <div style={{
                    width: '28px', height: '28px', borderRadius: '6px',
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    flexShrink: 0, background: tc.primaryLighter,
                  }}>
                    <DocumentIcon size={14} color={tc.primary} />
                  </div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: '12px', color: '#334155', fontWeight: 600 }}>{rp.title}</div>
                    <div style={{ fontSize: '10px', color: '#94a3b8' }}>POL-{rp.id.toString().padStart(3, '0')} | {getStatusLabel(rp.status)}</div>
                  </div>
                </div>
              ))}
            </div>
          )}

          {/* Primary Action buttons */}
          <div style={{ display: 'flex', gap: '8px', marginTop: '16px' }}>
            <button type="button" onClick={() => this.handlePolicyClick(policy.id)}
              style={{ flex: 1, padding: '10px', borderRadius: '6px', fontSize: '13px', fontWeight: 600, cursor: 'pointer', border: `1px solid ${tc.primary}`, fontFamily: 'inherit', textAlign: 'center', background: tc.primary, color: '#fff' }}
              onMouseEnter={(e) => { e.currentTarget.style.background = tc.primaryDark; }}
              onMouseLeave={(e) => { e.currentTarget.style.background = tc.primary; }}>
              {policy.status === 'completed' ? 'View Policy' : 'Read Policy'}
            </button>
            {policy.hasQuiz && policy.status !== 'completed' && (
              <button type="button" onClick={() => this.handlePolicyClick(policy.id)}
                style={{ flex: 1, padding: '10px', borderRadius: '6px', fontSize: '13px', fontWeight: 600, cursor: 'pointer', border: `1px solid ${tc.primary}`, fontFamily: 'inherit', textAlign: 'center', background: '#fff', color: tc.primary }}
                onMouseEnter={(e) => { e.currentTarget.style.background = tc.primaryLighter; }}
                onMouseLeave={(e) => { e.currentTarget.style.background = '#fff'; }}>
                Start Quiz
              </button>
            )}
          </div>

          {/* Secondary actions — commented out for future release
          {policy.status !== 'completed' && (
            <div style={{ display: 'flex', gap: '8px', marginTop: '8px' }}>
              <button type="button" onClick={() => { window.location.href = `${this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager'}/SitePages/PolicyDetails.aspx?policyId=${policy.id}&mode=browse`; }}
                style={{ flex: 1, padding: '8px', borderRadius: '4px', fontSize: '11px', fontWeight: 600, cursor: 'pointer', border: '1px solid #e2e8f0', fontFamily: 'inherit', textAlign: 'center', background: '#fff', color: '#64748b' }}>
                Mark as Read
              </button>
              <button type="button" disabled style={{ flex: 1, padding: '8px', borderRadius: '4px', fontSize: '11px', fontWeight: 600, cursor: 'not-allowed', border: '1px solid #e2e8f0', fontFamily: 'inherit', textAlign: 'center', background: '#f8fafc', color: '#94a3b8', opacity: 0.7 }}>
                Snooze Reminder
              </button>
              <button type="button" disabled style={{ flex: 1, padding: '8px', borderRadius: '4px', fontSize: '11px', fontWeight: 600, cursor: 'not-allowed', border: '1px solid #e2e8f0', fontFamily: 'inherit', textAlign: 'center', background: '#f8fafc', color: '#94a3b8', opacity: 0.7 }}>
                Request Extension
              </button>
            </div>
          )}
          */}
        </div>
        )}
      </StyledPanel>
    );
  }

  private renderEmptyState(): React.ReactNode {
    return (
      <div style={{
        background: '#fff',
        border: '1px solid #e2e8f0',
        borderRadius: '10px',
        overflow: 'hidden',
      }}>
        {this.renderToolbar()}
        <div style={{
          padding: '60px 20px',
          textAlign: 'center',
        }}>
          <div style={{ marginBottom: '16px', display: 'flex', justifyContent: 'center' }}>
            <CheckIcon size={48} color="#059669" />
          </div>
          <h3 style={{ margin: '16px 0 8px', color: '#323130', fontSize: '16px', fontWeight: 600 }}>All caught up!</h3>
          <p style={{ color: '#605e5c', fontSize: '13px' }}>No policies matching your current filter. Try a different tab or clear search.</p>
        </div>
      </div>
    );
  }

  public render(): React.ReactElement<IMyPoliciesProps> {
    const { loading, selectedPolicyId } = this.state;
    const filteredPolicies = this.getFilteredPolicies();

    return (
      <ErrorBoundary fallbackMessage="An error occurred in My Policies. Please try again.">
      <JmlAppLayout
        context={this.props.context}
        sp={this.props.sp}
        breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'My Policies' }]}
        activeNavKey="my-policies"
      >
        <div className={styles.myPolicies} style={{ width: '100%', height: '100%' }}>
          {loading ? (
            <div style={{ padding: '60px', textAlign: 'center', width: '100%' }}>
              <Spinner size={SpinnerSize.large} label="Loading your policies..." />
            </div>
          ) : (
            <div style={{ width: '100%', height: '100%', minHeight: 'calc(100vh - 180px)' }}>
              {/* Hero Banner — single row: ring + greeting + search + KPIs */}
              {this.renderHeroBanner()}

              {/* Main content area */}
              <div style={{ display: 'flex', width: '100%', flex: 1 }}>
                <div style={{ flex: 1, overflowY: 'auto', padding: '24px 40px' }}>
                  <div style={{ maxWidth: '1400px' }}>
                    {/* Policy List */}
                    {this.renderPolicyList(filteredPolicies)}
                  </div>
                </div>

                {/* Detail Panel (slide-in from right) */}
                {this.renderDetailPanel()}
              </div>
            </div>
          )}
        </div>
      </JmlAppLayout>
      </ErrorBoundary>
    );
  }
}
