import * as React from 'react';
import styles from './PolicyDistribution.module.scss';
import { IPolicyDistributionProps } from './IPolicyDistributionProps';
import { JmlAppLayout } from '../../../components/JmlAppLayout/JmlAppLayout';
import {
  PrimaryButton,
  DefaultButton,
  IconButton,
  SearchBox,
  Panel,
  PanelType,
  TextField,
  Dropdown,
  IDropdownOption,
  DatePicker,
  Toggle,
  Spinner,
  SpinnerSize,
  Icon,
  Label,
  MessageBar,
  MessageBarType,
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

// ============================================================================
// INTERFACES
// ============================================================================

interface IDistributionCampaign {
  id: number;
  campaignName: string;
  // Content — either a single policy or a policy pack
  contentType: 'Policy' | 'Policy Pack';
  policyTitle: string;
  policyId: number;
  policyPackName?: string;
  policyPackId?: number;
  // Targeting — scope + specific users/groups
  scope: string;
  targetUsers: string[];   // individual user names
  targetGroups: string[];  // group/department names
  status: 'Draft' | 'Scheduled' | 'Active' | 'Completed' | 'Paused';
  scheduledDate?: Date;
  distributedDate?: Date;
  dueDate?: Date;
  targetCount: number;
  totalSent: number;
  totalDelivered: number;
  totalOpened: number;
  totalAcknowledged: number;
  totalOverdue: number;
  totalExempted: number;
  totalFailed: number;
  escalationEnabled: boolean;
  reminderSchedule: string;
  isActive: boolean;
  completedDate?: Date;
  createdDate: Date;
  createdBy: string;
}

interface IRecipient {
  id: number;
  name: string;
  email: string;
  department: string;
  status: 'Pending' | 'Sent' | 'Delivered' | 'Opened' | 'Acknowledged' | 'Overdue' | 'Exempted' | 'Failed';
  sentDate?: Date;
  openedDate?: Date;
  acknowledgedDate?: Date;
}

interface IPolicyDistributionState {
  loading: boolean;
  campaigns: IDistributionCampaign[];
  filteredCampaigns: IDistributionCampaign[];
  searchQuery: string;
  activeFilter: string;
  selectedCampaign: IDistributionCampaign | null;
  recipients: IRecipient[];
  showCreatePanel: boolean;
  editingCampaign: IDistributionCampaign | null;
  // Form fields
  formCampaignName: string;
  formContentType: 'Policy' | 'Policy Pack';
  formPolicyId: string;
  formPolicyPackId: string;
  formScope: string;
  formTargetUsers: string;
  formTargetGroups: string;
  formScheduledDate: Date | undefined;
  formDueDate: Date | undefined;
  formEscalation: boolean;
  formReminder: string;
  // Messages
  successMessage: string;
  errorMessage: string;
}

// ============================================================================
// MOCK DATA
// ============================================================================

const MOCK_CAMPAIGNS: IDistributionCampaign[] = [
  {
    id: 1,
    campaignName: 'Q1 2026 — IT Security Policy Update',
    contentType: 'Policy',
    policyTitle: 'Information Security Policy v3.2',
    policyId: 101,
    scope: 'All Employees',
    targetUsers: [],
    targetGroups: ['All Employees'],
    status: 'Active',
    scheduledDate: new Date(2026, 0, 15),
    distributedDate: new Date(2026, 0, 15),
    dueDate: new Date(2026, 1, 15),
    targetCount: 342,
    totalSent: 342,
    totalDelivered: 338,
    totalOpened: 290,
    totalAcknowledged: 245,
    totalOverdue: 52,
    totalExempted: 8,
    totalFailed: 4,
    escalationEnabled: true,
    reminderSchedule: '7,14,21',
    isActive: true,
    createdDate: new Date(2026, 0, 10),
    createdBy: 'Gary Finberg',
  },
  {
    id: 2,
    campaignName: 'GDPR Annual Refresher',
    contentType: 'Policy',
    policyTitle: 'Data Privacy & GDPR Compliance',
    policyId: 102,
    scope: 'Department',
    targetUsers: [],
    targetGroups: ['Legal', 'Compliance'],
    status: 'Completed',
    scheduledDate: new Date(2025, 11, 1),
    distributedDate: new Date(2025, 11, 1),
    dueDate: new Date(2025, 11, 31),
    targetCount: 85,
    totalSent: 85,
    totalDelivered: 85,
    totalOpened: 85,
    totalAcknowledged: 82,
    totalOverdue: 0,
    totalExempted: 3,
    totalFailed: 0,
    escalationEnabled: true,
    reminderSchedule: '7,14',
    isActive: false,
    completedDate: new Date(2025, 11, 29),
    createdDate: new Date(2025, 10, 25),
    createdBy: 'Sarah Mitchell',
  },
  {
    id: 3,
    campaignName: 'New Hire — Health & Safety Onboarding',
    contentType: 'Policy Pack',
    policyTitle: 'New Hire Onboarding Pack',
    policyId: 0,
    policyPackName: 'New Hire Onboarding Pack',
    policyPackId: 10,
    scope: 'New Hires Only',
    targetUsers: [],
    targetGroups: ['New Hires'],
    status: 'Active',
    scheduledDate: new Date(2026, 0, 1),
    distributedDate: new Date(2026, 0, 1),
    dueDate: new Date(2026, 0, 31),
    targetCount: 18,
    totalSent: 18,
    totalDelivered: 18,
    totalOpened: 14,
    totalAcknowledged: 10,
    totalOverdue: 4,
    totalExempted: 0,
    totalFailed: 0,
    escalationEnabled: false,
    reminderSchedule: '3,7',
    isActive: true,
    createdDate: new Date(2025, 11, 28),
    createdBy: 'Gary Finberg',
  },
  {
    id: 4,
    campaignName: 'Code of Conduct 2026',
    contentType: 'Policy',
    policyTitle: 'Code of Conduct & Ethics',
    policyId: 104,
    scope: 'All Employees',
    targetUsers: [],
    targetGroups: ['All Employees'],
    status: 'Scheduled',
    scheduledDate: new Date(2026, 1, 1),
    dueDate: new Date(2026, 2, 1),
    targetCount: 342,
    totalSent: 0,
    totalDelivered: 0,
    totalOpened: 0,
    totalAcknowledged: 0,
    totalOverdue: 0,
    totalExempted: 0,
    totalFailed: 0,
    escalationEnabled: true,
    reminderSchedule: '7,14,21',
    isActive: true,
    createdDate: new Date(2026, 0, 20),
    createdBy: 'Gary Finberg',
  },
  {
    id: 5,
    campaignName: 'Finance Team — Anti-Fraud Policy',
    contentType: 'Policy',
    policyTitle: 'Anti-Fraud & Financial Controls',
    policyId: 105,
    scope: 'Role',
    targetUsers: ['Alice Johnson', 'Bob Williams'],
    targetGroups: ['Finance'],
    status: 'Draft',
    dueDate: new Date(2026, 2, 15),
    targetCount: 45,
    totalSent: 0,
    totalDelivered: 0,
    totalOpened: 0,
    totalAcknowledged: 0,
    totalOverdue: 0,
    totalExempted: 0,
    totalFailed: 0,
    escalationEnabled: false,
    reminderSchedule: '7',
    isActive: false,
    createdDate: new Date(2026, 0, 25),
    createdBy: 'Sarah Mitchell',
  },
];

const MOCK_RECIPIENTS: IRecipient[] = [
  { id: 1, name: 'Alice Johnson', email: 'alice.johnson@company.com', department: 'Engineering', status: 'Acknowledged', sentDate: new Date(2026, 0, 15), openedDate: new Date(2026, 0, 16), acknowledgedDate: new Date(2026, 0, 17) },
  { id: 2, name: 'Bob Williams', email: 'bob.williams@company.com', department: 'Marketing', status: 'Opened', sentDate: new Date(2026, 0, 15), openedDate: new Date(2026, 0, 18) },
  { id: 3, name: 'Carol Davis', email: 'carol.davis@company.com', department: 'Finance', status: 'Overdue', sentDate: new Date(2026, 0, 15) },
  { id: 4, name: 'David Chen', email: 'david.chen@company.com', department: 'HR', status: 'Acknowledged', sentDate: new Date(2026, 0, 15), openedDate: new Date(2026, 0, 15), acknowledgedDate: new Date(2026, 0, 16) },
  { id: 5, name: 'Emily Brown', email: 'emily.brown@company.com', department: 'Engineering', status: 'Sent', sentDate: new Date(2026, 0, 15) },
  { id: 6, name: 'Frank Garcia', email: 'frank.garcia@company.com', department: 'Operations', status: 'Exempted', sentDate: new Date(2026, 0, 15) },
  { id: 7, name: 'Grace Lee', email: 'grace.lee@company.com', department: 'Legal', status: 'Acknowledged', sentDate: new Date(2026, 0, 15), openedDate: new Date(2026, 0, 16), acknowledgedDate: new Date(2026, 0, 20) },
  { id: 8, name: 'Henry Wilson', email: 'henry.wilson@company.com', department: 'Engineering', status: 'Failed', sentDate: new Date(2026, 0, 15) },
];

const SCOPE_OPTIONS: IDropdownOption[] = [
  { key: 'All Employees', text: 'All Employees' },
  { key: 'Department', text: 'Department' },
  { key: 'Location', text: 'Location' },
  { key: 'Role', text: 'Role' },
  { key: 'Custom', text: 'Custom' },
  { key: 'New Hires Only', text: 'New Hires Only' },
];

const REMINDER_OPTIONS: IDropdownOption[] = [
  { key: '3', text: 'Every 3 days' },
  { key: '7', text: 'Every 7 days' },
  { key: '7,14', text: '7 and 14 days' },
  { key: '7,14,21', text: '7, 14, and 21 days' },
  { key: '3,7,14', text: '3, 7, and 14 days' },
  { key: 'custom', text: 'Custom schedule' },
];

const POLICY_PACK_OPTIONS: IDropdownOption[] = [
  { key: '10', text: 'New Hire Onboarding Pack' },
  { key: '11', text: 'Annual Compliance Refresher Pack' },
  { key: '12', text: 'IT Security Essentials Pack' },
  { key: '13', text: 'Manager Leadership Pack' },
  { key: '14', text: 'GDPR & Data Privacy Pack' },
  { key: '15', text: 'Health & Safety Essentials Pack' },
  { key: '16', text: 'Finance & Anti-Fraud Pack' },
];

const POLICY_OPTIONS: IDropdownOption[] = [
  { key: '101', text: 'Information Security Policy v3.2' },
  { key: '102', text: 'Data Privacy & GDPR Compliance' },
  { key: '103', text: 'Acceptable Use Policy' },
  { key: '104', text: 'Code of Conduct & Ethics' },
  { key: '105', text: 'Anti-Fraud & Financial Controls' },
  { key: '106', text: 'Health & Safety at Work' },
  { key: '107', text: 'Remote Working Policy' },
  { key: '108', text: 'Anti-Bribery & Corruption' },
  { key: '109', text: 'Environmental Sustainability Policy' },
  { key: '110', text: 'Whistleblowing Policy' },
];

const FILTER_TABS = ['All', 'Active', 'Scheduled', 'Completed', 'Draft', 'Paused'];

// ============================================================================
// COMPONENT
// ============================================================================

export default class PolicyDistribution extends React.Component<IPolicyDistributionProps, IPolicyDistributionState> {
  constructor(props: IPolicyDistributionProps) {
    super(props);
    this.state = {
      loading: false,
      campaigns: MOCK_CAMPAIGNS,
      filteredCampaigns: MOCK_CAMPAIGNS,
      searchQuery: '',
      activeFilter: 'All',
      selectedCampaign: null,
      recipients: MOCK_RECIPIENTS,
      showCreatePanel: false,
      editingCampaign: null,
      formCampaignName: '',
      formContentType: 'Policy',
      formPolicyId: '',
      formPolicyPackId: '',
      formScope: 'All Employees',
      formTargetUsers: '',
      formTargetGroups: '',
      formScheduledDate: undefined,
      formDueDate: undefined,
      formEscalation: true,
      formReminder: '7,14',
      successMessage: '',
      errorMessage: '',
    };
  }

  // ──────────── Filtering ────────────

  private applyFilters = (search?: string, filter?: string): void => {
    const searchQuery = search !== undefined ? search : this.state.searchQuery;
    const activeFilter = filter !== undefined ? filter : this.state.activeFilter;
    const { campaigns } = this.state;

    let filtered = [...campaigns];

    if (activeFilter !== 'All') {
      filtered = filtered.filter(c => c.status === activeFilter);
    }

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      filtered = filtered.filter(c =>
        c.campaignName.toLowerCase().includes(q) ||
        c.policyTitle.toLowerCase().includes(q) ||
        c.scope.toLowerCase().includes(q)
      );
    }

    this.setState({ filteredCampaigns: filtered, searchQuery, activeFilter });
  };

  // ──────────── CRUD ────────────

  private openCreatePanel = (): void => {
    this.setState({
      showCreatePanel: true,
      editingCampaign: null,
      formCampaignName: '',
      formContentType: 'Policy',
      formPolicyId: '',
      formPolicyPackId: '',
      formScope: 'All Employees',
      formTargetUsers: '',
      formTargetGroups: '',
      formScheduledDate: undefined,
      formDueDate: undefined,
      formEscalation: true,
      formReminder: '7,14',
    });
  };

  private openEditPanel = (campaign: IDistributionCampaign): void => {
    this.setState({
      showCreatePanel: true,
      editingCampaign: campaign,
      formCampaignName: campaign.campaignName,
      formContentType: campaign.contentType,
      formPolicyId: campaign.policyId ? campaign.policyId.toString() : '',
      formPolicyPackId: campaign.policyPackId ? campaign.policyPackId.toString() : '',
      formScope: campaign.scope,
      formTargetUsers: campaign.targetUsers.join(', '),
      formTargetGroups: campaign.targetGroups.join(', '),
      formScheduledDate: campaign.scheduledDate,
      formDueDate: campaign.dueDate,
      formEscalation: campaign.escalationEnabled,
      formReminder: campaign.reminderSchedule,
    });
  };

  private saveCampaign = (): void => {
    const { editingCampaign, campaigns, formCampaignName, formContentType, formPolicyId, formPolicyPackId, formScope, formTargetUsers, formTargetGroups, formScheduledDate, formDueDate, formEscalation, formReminder } = this.state;

    if (!formCampaignName.trim()) {
      this.setState({ errorMessage: 'Campaign name is required.' });
      return;
    }

    const targetUsers = formTargetUsers ? formTargetUsers.split(',').map(s => s.trim()).filter(Boolean) : [];
    const targetGroups = formTargetGroups ? formTargetGroups.split(',').map(s => s.trim()).filter(Boolean) : [];

    if (editingCampaign) {
      const updated = campaigns.map(c =>
        c.id === editingCampaign.id
          ? {
              ...c,
              campaignName: formCampaignName,
              contentType: formContentType,
              policyId: parseInt(formPolicyId) || c.policyId,
              policyPackId: formContentType === 'Policy Pack' ? (parseInt(formPolicyPackId) || c.policyPackId) : undefined,
              policyPackName: formContentType === 'Policy Pack' ? (c.policyPackName || 'Selected Pack') : undefined,
              scope: formScope,
              targetUsers,
              targetGroups,
              scheduledDate: formScheduledDate,
              dueDate: formDueDate,
              escalationEnabled: formEscalation,
              reminderSchedule: formReminder,
            }
          : c
      );
      this.setState({ campaigns: updated, showCreatePanel: false, successMessage: 'Campaign updated successfully.' }, () => this.applyFilters());
    } else {
      const newCampaign: IDistributionCampaign = {
        id: Date.now(),
        campaignName: formCampaignName,
        contentType: formContentType,
        policyTitle: formContentType === 'Policy' ? 'Selected Policy' : 'Selected Policy Pack',
        policyId: parseInt(formPolicyId) || 0,
        policyPackId: formContentType === 'Policy Pack' ? (parseInt(formPolicyPackId) || 0) : undefined,
        policyPackName: formContentType === 'Policy Pack' ? 'Selected Pack' : undefined,
        scope: formScope,
        targetUsers,
        targetGroups,
        status: formScheduledDate ? 'Scheduled' : 'Draft',
        scheduledDate: formScheduledDate,
        dueDate: formDueDate,
        targetCount: formScope === 'All Employees' ? 342 : (targetUsers.length + targetGroups.length * 25),
        totalSent: 0,
        totalDelivered: 0,
        totalOpened: 0,
        totalAcknowledged: 0,
        totalOverdue: 0,
        totalExempted: 0,
        totalFailed: 0,
        escalationEnabled: formEscalation,
        reminderSchedule: formReminder,
        isActive: false,
        createdDate: new Date(),
        createdBy: 'Current User',
      };
      this.setState({ campaigns: [...campaigns, newCampaign], showCreatePanel: false, successMessage: 'Campaign created successfully.' }, () => this.applyFilters());
    }

    setTimeout(() => this.setState({ successMessage: '' }), 3000);
  };

  private deleteCampaign = (id: number): void => {
    const updated = this.state.campaigns.filter(c => c.id !== id);
    this.setState({ campaigns: updated, successMessage: 'Campaign deleted.' }, () => this.applyFilters());
    setTimeout(() => this.setState({ successMessage: '' }), 3000);
  };

  private selectCampaign = (campaign: IDistributionCampaign): void => {
    this.setState({ selectedCampaign: campaign });
  };

  private clearSelection = (): void => {
    this.setState({ selectedCampaign: null });
  };

  // ──────────── Helpers ────────────

  private formatDate = (date?: Date): string => {
    if (!date) return '—';
    return date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  };

  private getStatusStyle = (status: string): string => {
    switch (status) {
      case 'Active': return styles.statusActive;
      case 'Scheduled': return styles.statusScheduled;
      case 'Completed': return styles.statusCompleted;
      case 'Draft': return styles.statusDraft;
      case 'Paused': return styles.statusPaused;
      default: return styles.statusDraft;
    }
  };

  private getRecipientStatusStyle = (status: string): string => {
    switch (status) {
      case 'Acknowledged': return styles.statusActive;
      case 'Opened': return styles.statusScheduled;
      case 'Sent': case 'Delivered': case 'Pending': return styles.statusDraft;
      case 'Overdue': return styles.statusPaused;
      case 'Failed': return styles.statusPaused;
      case 'Exempted': return styles.statusCompleted;
      default: return styles.statusDraft;
    }
  };

  private getAckPercentage = (campaign: IDistributionCampaign): number => {
    if (campaign.targetCount === 0) return 0;
    return Math.round((campaign.totalAcknowledged / campaign.targetCount) * 100);
  };

  // ──────────── KPI Section ────────────

  private getKPIs = (): { label: string; value: string | number; className?: string }[] => {
    const { campaigns } = this.state;
    const active = campaigns.filter(c => c.status === 'Active');
    const totalSent = campaigns.reduce((sum, c) => sum + c.totalSent, 0);
    const totalAck = campaigns.reduce((sum, c) => sum + c.totalAcknowledged, 0);
    const totalOverdue = active.reduce((sum, c) => sum + c.totalOverdue, 0);
    const totalFailed = campaigns.reduce((sum, c) => sum + c.totalFailed, 0);
    const ackRate = totalSent > 0 ? Math.round((totalAck / totalSent) * 100) : 0;

    return [
      { label: 'Total Campaigns', value: campaigns.length },
      { label: 'Active', value: active.length, className: styles.kpiSuccess },
      { label: 'Total Sent', value: totalSent.toLocaleString() },
      { label: 'Acknowledged', value: totalAck.toLocaleString(), className: styles.kpiSuccess },
      { label: 'Ack Rate', value: `${ackRate}%`, className: styles.kpiSuccess },
      { label: 'Overdue', value: totalOverdue, className: totalOverdue > 0 ? styles.kpiAccent : undefined },
      { label: 'Failed', value: totalFailed, className: totalFailed > 0 ? styles.kpiDanger : undefined },
    ];
  };

  // ──────────── RENDER: KPI Dashboard ────────────

  private renderKPIs(): React.ReactElement {
    const kpis = this.getKPIs();
    return (
      <div className={styles.kpiSection}>
        <div className={styles.kpiGrid}>
          {kpis.map((kpi, idx) => (
            <div key={idx} className={`${styles.kpiCard} ${kpi.className || ''}`}>
              <div className={styles.kpiValue}>{kpi.value}</div>
              <div className={styles.kpiLabel}>{kpi.label}</div>
            </div>
          ))}
        </div>
      </div>
    );
  }

  // ──────────── RENDER: Toolbar ────────────

  private renderToolbar(): React.ReactElement {
    const { searchQuery, activeFilter } = this.state;
    return (
      <div className={styles.toolbar}>
        <div className={styles.toolbarLeft}>
          <SearchBox
            placeholder="Search campaigns..."
            value={searchQuery}
            onChange={(_, val) => this.applyFilters(val || '', undefined)}
            styles={{ root: { maxWidth: 280, flex: 1 } }}
          />
          {FILTER_TABS.map(tab => (
            <button
              key={tab}
              className={`${styles.filterChip} ${activeFilter === tab ? styles.filterChipActive : ''}`}
              onClick={() => this.applyFilters(undefined, tab)}
            >
              {tab}
            </button>
          ))}
        </div>
        <div className={styles.toolbarRight}>
          <PrimaryButton
            text="New Campaign"
            iconProps={{ iconName: 'Add' }}
            onClick={this.openCreatePanel}
            styles={{ root: { borderRadius: 6, background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' }, label: { fontWeight: 400 } }}
          />
        </div>
      </div>
    );
  }

  // ──────────── RENDER: Campaign Table ────────────

  private renderCampaignTable(): React.ReactElement {
    const { filteredCampaigns } = this.state;

    if (filteredCampaigns.length === 0) {
      return (
        <div className={styles.emptyState}>
          <div className={styles.emptyIcon}><Icon iconName="MailForward" /></div>
          <div className={styles.emptyTitle}>No campaigns found</div>
          <div className={styles.emptyText}>Create your first distribution campaign to start distributing policies to your team.</div>
        </div>
      );
    }

    return (
      <div className={styles.campaignList}>
        <table className={styles.campaignTable}>
          <thead>
            <tr>
              <th>Campaign</th>
              <th>Type</th>
              <th>Content</th>
              <th>Scope</th>
              <th>Status</th>
              <th>Progress</th>
              <th>Due Date</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {filteredCampaigns.map(campaign => {
              const pct = this.getAckPercentage(campaign);
              return (
                <tr key={campaign.id}>
                  <td>
                    <span className={styles.campaignName} onClick={() => this.selectCampaign(campaign)}>
                      <Icon iconName="MailForward" style={{ fontSize: 14, color: '#0d9488' }} />
                      {campaign.campaignName}
                    </span>
                  </td>
                  <td>
                    <span className={`${styles.statusBadge} ${campaign.contentType === 'Policy Pack' ? styles.statusCompleted : styles.statusDraft}`}>
                      <Icon iconName={campaign.contentType === 'Policy Pack' ? 'Package' : 'Document'} style={{ fontSize: 11 }} />
                      {campaign.contentType}
                    </span>
                  </td>
                  <td>{campaign.contentType === 'Policy Pack' ? campaign.policyPackName : campaign.policyTitle}</td>
                  <td>{campaign.scope}{campaign.targetGroups.length > 0 && campaign.scope !== 'All Employees' ? ` (${campaign.targetGroups.join(', ')})` : ''}</td>
                  <td>
                    <span className={`${styles.statusBadge} ${this.getStatusStyle(campaign.status)}`}>
                      {campaign.status}
                    </span>
                  </td>
                  <td>
                    <div className={styles.progressBarContainer}>
                      <div className={styles.progressTrack}>
                        <div className={styles.progressFill} style={{ width: `${pct}%` }} />
                      </div>
                      <span className={styles.progressLabel}>{pct}%</span>
                    </div>
                  </td>
                  <td>{this.formatDate(campaign.dueDate)}</td>
                  <td>
                    <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={() => this.openEditPanel(campaign)} />
                    <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" onClick={() => this.deleteCampaign(campaign.id)} styles={{ root: { color: '#ef4444' }, rootHovered: { color: '#dc2626' } }} />
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  }

  // ──────────── RENDER: Detail View ────────────

  private renderDetailView(): React.ReactElement {
    const { selectedCampaign, recipients } = this.state;
    if (!selectedCampaign) return <></>;

    const c = selectedCampaign;
    const pct = this.getAckPercentage(c);

    return (
      <div className={styles.detailView}>
        <div className={styles.backButton} onClick={this.clearSelection}>
          <Icon iconName="ChevronLeft" /> Back to campaigns
        </div>

        <div className={styles.detailHeader}>
          <div>
            <h2 className={styles.detailTitle}>{c.campaignName}</h2>
            <div className={styles.detailMeta}>
              <span className={styles.detailMetaItem}>
                <Icon iconName={c.contentType === 'Policy Pack' ? 'Package' : 'Document'} />
                {c.contentType === 'Policy Pack' ? c.policyPackName : c.policyTitle}
              </span>
              <span className={`${styles.statusBadge} ${c.contentType === 'Policy Pack' ? styles.statusCompleted : styles.statusDraft}`} style={{ fontSize: 11, padding: '2px 8px' }}>
                {c.contentType}
              </span>
              <span className={styles.detailMetaItem}><Icon iconName="People" /> {c.scope}{c.targetGroups.length > 0 ? ` — ${c.targetGroups.join(', ')}` : ''}</span>
              <span className={styles.detailMetaItem}><Icon iconName="Calendar" /> Due: {this.formatDate(c.dueDate)}</span>
              <span className={`${styles.statusBadge} ${this.getStatusStyle(c.status)}`}>{c.status}</span>
            </div>
          </div>
          <div className={styles.detailActions}>
            {c.status === 'Draft' && (
              <PrimaryButton text="Send Now" iconProps={{ iconName: 'Send' }} styles={{ root: { borderRadius: 6, background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }} />
            )}
            {c.status === 'Active' && (
              <DefaultButton text="Pause" iconProps={{ iconName: 'Pause' }} styles={{ root: { borderRadius: 6 } }} />
            )}
            <DefaultButton text="Send Reminder" iconProps={{ iconName: 'Ringer' }} styles={{ root: { borderRadius: 6 } }} />
            <IconButton iconProps={{ iconName: 'Edit' }} title="Edit Campaign" onClick={() => this.openEditPanel(c)} />
          </div>
        </div>

        {/* Stats Grid */}
        <div className={styles.detailStatsGrid}>
          {[
            { label: 'Target', value: c.targetCount },
            { label: 'Sent', value: c.totalSent },
            { label: 'Delivered', value: c.totalDelivered },
            { label: 'Opened', value: c.totalOpened },
            { label: 'Acknowledged', value: c.totalAcknowledged },
            { label: 'Overdue', value: c.totalOverdue },
            { label: 'Exempted', value: c.totalExempted },
            { label: 'Failed', value: c.totalFailed },
          ].map((stat, idx) => (
            <div key={idx} className={styles.detailStatCard}>
              <div className={styles.detailStatValue}>{stat.value}</div>
              <div className={styles.detailStatLabel}>{stat.label}</div>
            </div>
          ))}
        </div>

        {/* Overall Progress */}
        <div style={{ marginBottom: 24 }}>
          <div className={styles.sectionTitle}><Icon iconName="ProgressRingDots" /> Overall Acknowledgement Progress</div>
          <div className={styles.progressBarContainer} style={{ maxWidth: 400 }}>
            <div className={styles.progressTrack} style={{ height: 12 }}>
              <div className={styles.progressFill} style={{ width: `${pct}%`, height: 12 }} />
            </div>
            <span className={styles.progressLabel} style={{ fontSize: 16 }}>{pct}%</span>
          </div>
        </div>

        {/* Recipients Table */}
        <div className={styles.recipientSection}>
          <div className={styles.sectionTitle}><Icon iconName="People" /> Recipients ({recipients.length})</div>
          <table className={styles.recipientTable}>
            <thead>
              <tr>
                <th>Name</th>
                <th>Email</th>
                <th>Department</th>
                <th>Status</th>
                <th>Sent</th>
                <th>Opened</th>
                <th>Acknowledged</th>
              </tr>
            </thead>
            <tbody>
              {recipients.map(r => (
                <tr key={r.id}>
                  <td style={{ fontWeight: 500 }}>{r.name}</td>
                  <td>{r.email}</td>
                  <td>{r.department}</td>
                  <td>
                    <span className={`${styles.statusBadge} ${this.getRecipientStatusStyle(r.status)}`}>
                      {r.status}
                    </span>
                  </td>
                  <td>{this.formatDate(r.sentDate)}</td>
                  <td>{this.formatDate(r.openedDate)}</td>
                  <td>{this.formatDate(r.acknowledgedDate)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Timeline */}
        <div className={styles.timeline}>
          <div className={styles.sectionTitle}><Icon iconName="Timeline" /> Activity Timeline</div>
          {[
            { text: `Campaign created by ${c.createdBy}`, date: c.createdDate },
            ...(c.scheduledDate ? [{ text: 'Campaign scheduled for distribution', date: c.scheduledDate }] : []),
            ...(c.distributedDate ? [{ text: `Distributed to ${c.totalSent} recipients`, date: c.distributedDate }] : []),
            ...(c.completedDate ? [{ text: 'Campaign completed — all acknowledgements received', date: c.completedDate }] : []),
          ].sort((a, b) => (b.date?.getTime() || 0) - (a.date?.getTime() || 0)).map((item, idx) => (
            <div key={idx} className={styles.timelineItem}>
              <div className={styles.timelineDot} />
              <div className={styles.timelineContent}>
                <div className={styles.timelineText}>{item.text}</div>
                <div className={styles.timelineDate}>{this.formatDate(item.date)}</div>
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  }

  // ──────────── RENDER: Create/Edit Panel ────────────

  private renderCampaignPanel(): React.ReactElement {
    const { showCreatePanel, editingCampaign, formCampaignName, formContentType, formPolicyId, formPolicyPackId, formScope, formTargetUsers, formTargetGroups, formScheduledDate, formDueDate, formEscalation, formReminder, errorMessage } = this.state;

    return (
      <Panel
        isOpen={showCreatePanel}
        onDismiss={() => this.setState({ showCreatePanel: false, errorMessage: '' })}
        type={PanelType.medium}
        headerText={editingCampaign ? 'Edit Distribution Campaign' : 'New Distribution Campaign'}
        onRenderFooterContent={() => (
          <div style={{ display: 'flex', gap: 8 }}>
            <PrimaryButton text={editingCampaign ? 'Update' : 'Create Campaign'} onClick={this.saveCampaign} styles={{ root: { borderRadius: 6, background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }} />
            <DefaultButton text="Cancel" onClick={() => this.setState({ showCreatePanel: false, errorMessage: '' })} styles={{ root: { borderRadius: 6 } }} />
          </div>
        )}
        isFooterAtBottom={true}
      >
        <div style={{ padding: '8px 0' }}>
          {errorMessage && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ errorMessage: '' })} styles={{ root: { marginBottom: 16 } }}>
              {errorMessage}
            </MessageBar>
          )}

          {/* Campaign Details */}
          <div className={styles.formSection}>
            <div className={styles.formSectionTitle}>Campaign Details</div>
            <div className={styles.formField}>
              <TextField
                label="Campaign Name"
                required
                value={formCampaignName}
                onChange={(_, val) => this.setState({ formCampaignName: val || '' })}
                placeholder="e.g., Q1 2026 — IT Security Policy Update"
              />
            </div>
          </div>

          {/* Content Selection — Policy or Policy Pack */}
          <div className={styles.formSection}>
            <div className={styles.formSectionTitle}>Content to Distribute</div>
            <div className={styles.formField}>
              <Dropdown
                label="Content Type"
                selectedKey={formContentType}
                options={[
                  { key: 'Policy', text: 'Single Policy' },
                  { key: 'Policy Pack', text: 'Policy Pack' },
                ]}
                onChange={(_, opt) => opt && this.setState({ formContentType: opt.key as 'Policy' | 'Policy Pack' })}
              />
            </div>
            {formContentType === 'Policy' ? (
              <div className={styles.formField}>
                <Dropdown
                  label="Policy"
                  selectedKey={formPolicyId || undefined}
                  options={POLICY_OPTIONS}
                  onChange={(_, opt) => opt && this.setState({ formPolicyId: opt.key as string })}
                  placeholder="Select a policy to distribute"
                />
              </div>
            ) : (
              <div className={styles.formField}>
                <Dropdown
                  label="Policy Pack"
                  selectedKey={formPolicyPackId || undefined}
                  options={POLICY_PACK_OPTIONS}
                  onChange={(_, opt) => opt && this.setState({ formPolicyPackId: opt.key as string })}
                  placeholder="Select a policy pack to distribute"
                />
              </div>
            )}
          </div>

          {/* Targeting — users and/or groups */}
          <div className={styles.formSection}>
            <div className={styles.formSectionTitle}>Target Recipients</div>
            <div className={styles.formField}>
              <Dropdown
                label="Distribution Scope"
                selectedKey={formScope}
                options={SCOPE_OPTIONS}
                onChange={(_, opt) => opt && this.setState({ formScope: opt.key as string })}
              />
            </div>
            <div className={styles.formField}>
              <Label>Target Users</Label>
              <PeoplePicker
                context={this.props.context as any}
                personSelectionLimit={50}
                groupName=""
                showtooltip={true}
                defaultSelectedUsers={formTargetUsers ? formTargetUsers.split(',').map(u => u.trim()).filter(Boolean) : []}
                onChange={(items: any[]) => {
                  const users = items.map(item => item.secondaryText || item.text || '').filter(Boolean);
                  this.setState({ formTargetUsers: users.join(', ') });
                }}
                principalTypes={[PrincipalType.User]}
                resolveDelay={500}
                placeholder="Search for users in Entra ID..."
              />
              <div style={{ fontSize: 12, color: '#605e5c', marginTop: 4 }}>Search and select individual users from your organisation directory</div>
            </div>
            <div className={styles.formField}>
              <Label>Target Groups / Departments</Label>
              <PeoplePicker
                context={this.props.context as any}
                personSelectionLimit={20}
                groupName=""
                showtooltip={true}
                defaultSelectedUsers={formTargetGroups ? formTargetGroups.split(',').map(g => g.trim()).filter(Boolean) : []}
                onChange={(items: any[]) => {
                  const groups = items.map(item => item.text || '').filter(Boolean);
                  this.setState({ formTargetGroups: groups.join(', ') });
                }}
                principalTypes={[PrincipalType.SharePointGroup, PrincipalType.SecurityGroup, PrincipalType.DistributionList]}
                resolveDelay={500}
                placeholder="Search for groups in Entra ID..."
              />
              <div style={{ fontSize: 12, color: '#605e5c', marginTop: 4 }}>Search and select security groups, distribution lists, or SharePoint groups</div>
            </div>
          </div>

          {/* Scheduling */}
          <div className={styles.formSection}>
            <div className={styles.formSectionTitle}>Schedule & Due Date</div>
            <div className={styles.formRow}>
              <div className={styles.formField}>
                <DatePicker
                  label="Scheduled Date"
                  value={formScheduledDate}
                  onSelectDate={(date) => this.setState({ formScheduledDate: date || undefined })}
                  placeholder="Select date..."
                />
              </div>
              <div className={styles.formField}>
                <DatePicker
                  label="Due Date"
                  value={formDueDate}
                  onSelectDate={(date) => this.setState({ formDueDate: date || undefined })}
                  placeholder="Select due date..."
                />
              </div>
            </div>
          </div>

          {/* Reminders & Escalation */}
          <div className={styles.formSection}>
            <div className={styles.formSectionTitle}>Reminders & Escalation</div>
            <div className={styles.formField}>
              <Dropdown
                label="Reminder Schedule"
                selectedKey={formReminder}
                options={REMINDER_OPTIONS}
                onChange={(_, opt) => opt && this.setState({ formReminder: opt.key as string })}
              />
            </div>
            <div className={styles.formField}>
              <Toggle
                label="Enable Escalation"
                checked={formEscalation}
                onChange={(_, val) => this.setState({ formEscalation: !!val })}
                onText="Yes — escalate to manager after due date"
                offText="No escalation"
              />
            </div>
          </div>

          {/* Preview */}
          <div className={styles.previewBox}>
            <div className={styles.previewTitle}>Campaign Preview</div>
            <div className={styles.previewStat}>
              <span>Campaign Name</span>
              <span style={{ fontWeight: 600 }}>{formCampaignName || '—'}</span>
            </div>
            <div className={styles.previewStat}>
              <span>Content Type</span>
              <span style={{ fontWeight: 600 }}>{formContentType}</span>
            </div>
            <div className={styles.previewStat}>
              <span>{formContentType === 'Policy' ? 'Policy ID' : 'Pack ID'}</span>
              <span style={{ fontWeight: 600 }}>{(formContentType === 'Policy' ? formPolicyId : formPolicyPackId) || '—'}</span>
            </div>
            <div className={styles.previewStat}>
              <span>Scope</span>
              <span style={{ fontWeight: 600 }}>{formScope}</span>
            </div>
            {formTargetUsers && (
              <div className={styles.previewStat}>
                <span>Target Users</span>
                <span style={{ fontWeight: 600 }}>{formTargetUsers.split(',').filter(Boolean).length} user(s)</span>
              </div>
            )}
            {formTargetGroups && (
              <div className={styles.previewStat}>
                <span>Target Groups</span>
                <span style={{ fontWeight: 600 }}>{formTargetGroups.split(',').filter(Boolean).length} group(s)</span>
              </div>
            )}
            <div className={styles.previewStat}>
              <span>Scheduled</span>
              <span style={{ fontWeight: 600 }}>{formScheduledDate ? this.formatDate(formScheduledDate) : 'Not scheduled'}</span>
            </div>
            <div className={styles.previewStat}>
              <span>Due Date</span>
              <span style={{ fontWeight: 600 }}>{formDueDate ? this.formatDate(formDueDate) : 'Not set'}</span>
            </div>
            <div className={styles.previewStat}>
              <span>Reminders</span>
              <span style={{ fontWeight: 600 }}>{formReminder} days</span>
            </div>
            <div className={styles.previewStat}>
              <span>Escalation</span>
              <span style={{ fontWeight: 600 }}>{formEscalation ? 'Enabled' : 'Disabled'}</span>
            </div>
          </div>
        </div>
      </Panel>
    );
  }

  // ──────────── RENDER: Main ────────────

  public render(): React.ReactElement<IPolicyDistributionProps> {
    const { loading, selectedCampaign, successMessage } = this.state;

    return (
      <JmlAppLayout
        context={this.props.context}
        pageTitle="Policy Distribution & Tracking"
        breadcrumbs={[
          { text: 'Policy Manager', url: '/sites/PolicyManager' },
          ...(selectedCampaign ? [{ text: 'Distribution', url: '#' }, { text: selectedCampaign.campaignName }] : [{ text: 'Distribution & Tracking' }]),
        ]}
        activeNavKey="distribution"
      >
        <div className={styles.policyDistribution}>
          {successMessage && (
            <MessageBar messageBarType={MessageBarType.success} onDismiss={() => this.setState({ successMessage: '' })} styles={{ root: { margin: '0 40px', marginTop: 16 } }}>
              {successMessage}
            </MessageBar>
          )}

          {loading ? (
            <div style={{ padding: 60, textAlign: 'center' }}>
              <Spinner size={SpinnerSize.large} label="Loading distribution campaigns..." />
            </div>
          ) : selectedCampaign ? (
            this.renderDetailView()
          ) : (
            <>
              {this.renderKPIs()}
              {this.renderToolbar()}
              {this.renderCampaignTable()}
            </>
          )}

          {this.renderCampaignPanel()}
        </div>
      </JmlAppLayout>
    );
  }
}
