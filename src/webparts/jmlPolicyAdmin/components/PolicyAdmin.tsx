// @ts-nocheck
/* eslint-disable */
import * as React from 'react';
import { IPolicyAdminProps } from './IPolicyAdminProps';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  DefaultButton,
  PrimaryButton,
  IconButton,
  Icon,
  CommandBar,
  ICommandBarItemProps,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  Selection,
  Dialog,
  DialogType,
  DialogFooter,
  TextField,
  Dropdown,
  IDropdownOption,
  Pivot,
  PivotItem,
  Panel,
  PanelType,
  Label
} from '@fluentui/react';
import { injectPortalStyles } from '../../../utils/injectPortalStyles';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { PolicyService } from '../../../services/PolicyService';
import { createDialogManager } from '../../../hooks/useDialog';
import {
  IPolicy,
  IPolicyPublishRequest,
  IPolicyDistribution,
  PolicyStatus,
  DistributionScope
} from '../../../models/IPolicy';
import styles from './PolicyAdmin.module.scss';

export interface IPolicyAdminState {
  loading: boolean;
  error: string | null;
  policies: IPolicy[];
  selectedPolicies: IPolicy[];
  selectedView: string;
  showPublishPanel: boolean;
  showArchiveDialog: boolean;
  selectedPolicy: IPolicy | null;
  publishTargetUserIds: string;
  publishTargetEmails: string;
  publishTargetDepartments: string;
  publishTargetRoles: string;
  publishDueDate: string;
  archiveReason: string;
  submitting: boolean;
}

export default class PolicyAdmin extends React.Component<IPolicyAdminProps, IPolicyAdminState> {
  private policyService: PolicyService;
  private selection: Selection;
  private dialogManager = createDialogManager();

  constructor(props: IPolicyAdminProps) {
    super(props);
    this.state = {
      loading: true,
      error: null,
      policies: [],
      selectedPolicies: [],
      selectedView: 'pending',
      showPublishPanel: false,
      showArchiveDialog: false,
      selectedPolicy: null,
      publishTargetUserIds: '',
      publishTargetEmails: '',
      publishTargetDepartments: '',
      publishTargetRoles: '',
      publishDueDate: '',
      archiveReason: '',
      submitting: false
    };

    this.policyService = new PolicyService(props.sp);

    this.selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectedPolicies: this.selection.getSelection() as IPolicy[]
        });
      }
    });
  }

  public async componentDidMount(): Promise<void> {
    injectPortalStyles();
    await this.loadPolicies();
  }

  private async loadPolicies(): Promise<void> {
    try {
      this.setState({ loading: true, error: null });
      await this.policyService.initialize();

      const allPolicies = await this.policyService.getAllPolicies();
      this.setState({ policies: allPolicies, loading: false });
    } catch (error) {
      console.error('Failed to load policies:', error);
      this.setState({
        error: 'Failed to load policies. Please try again later.',
        loading: false
      });
    }
  }

  private handleApprove = async (policy: IPolicy): Promise<void> => {
    try {
      this.setState({ submitting: true });
      await this.policyService.approvePolicy(policy.Id, 'Approved by admin');
      await this.loadPolicies();
      this.setState({ submitting: false });
      await this.dialogManager.showAlert('Policy approved successfully!', { variant: 'success' });
    } catch (error) {
      console.error('Failed to approve policy:', error);
      await this.dialogManager.showAlert('Failed to approve policy. Please try again.', { variant: 'error' });
      this.setState({ submitting: false });
    }
  };

  private handleReject = async (policy: IPolicy): Promise<void> => {
    const reason = await this.dialogManager.showPrompt('Please provide a reason for rejecting this policy:', {
      title: 'Reject Policy',
      label: 'Rejection Reason',
      placeholder: 'Enter rejection reason...',
      required: true,
      multiline: true,
      rows: 3
    });
    if (!reason) return;

    try {
      this.setState({ submitting: true });
      await this.policyService.rejectPolicy(policy.Id, reason);
      await this.loadPolicies();
      this.setState({ submitting: false });
      await this.dialogManager.showAlert('Policy rejected.', { variant: 'info' });
    } catch (error) {
      console.error('Failed to reject policy:', error);
      await this.dialogManager.showAlert('Failed to reject policy. Please try again.', { variant: 'error' });
      this.setState({ submitting: false });
    }
  };

  private handlePublish = (policy: IPolicy): void => {
    const dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + 14); // Default 14 days

    this.setState({
      selectedPolicy: policy,
      showPublishPanel: true,
      publishTargetUserIds: '',
      publishTargetEmails: '',
      publishTargetDepartments: '',
      publishTargetRoles: '',
      publishDueDate: dueDate.toISOString().split('T')[0]
    });
  };

  private handleSubmitPublish = async (): Promise<void> => {
    const {
      selectedPolicy,
      publishTargetUserIds,
      publishTargetEmails,
      publishTargetDepartments,
      publishTargetRoles,
      publishDueDate
    } = this.state;

    if (!selectedPolicy) return;

    // Parse inputs
    const userIds = publishTargetUserIds
      ? publishTargetUserIds.split(',').map(id => parseInt(id.trim(), 10)).filter(id => !isNaN(id))
      : undefined;

    const emails = publishTargetEmails
      ? publishTargetEmails.split(',').map(email => email.trim()).filter(e => e)
      : undefined;

    const departments = publishTargetDepartments
      ? publishTargetDepartments.split(',').map(d => d.trim()).filter(d => d)
      : undefined;

    const roles = publishTargetRoles
      ? publishTargetRoles.split(',').map(r => r.trim()).filter(r => r)
      : undefined;

    try {
      this.setState({ submitting: true });

      const request: IPolicyPublishRequest = {
        policyId: selectedPolicy.Id,
        distributionScope: userIds || emails ? DistributionScope.Custom : DistributionScope.AllEmployees,
        targetUserIds: userIds,
        targetEmails: emails,
        targetDepartments: departments,
        targetRoles: roles,
        dueDate: publishDueDate ? new Date(publishDueDate) : undefined,
        sendNotifications: true
      };

      await this.policyService.publishPolicy(request);

      this.setState({
        showPublishPanel: false,
        submitting: false
      });

      await this.loadPolicies();
      await this.dialogManager.showAlert('Policy published successfully!', { variant: 'success' });
    } catch (error) {
      console.error('Failed to publish policy:', error);
      this.setState({
        error: 'Failed to publish policy. Please try again.',
        submitting: false
      });
    }
  };

  private handleArchive = (policy: IPolicy): void => {
    this.setState({
      selectedPolicy: policy,
      showArchiveDialog: true,
      archiveReason: ''
    });
  };

  private handleSubmitArchive = async (): Promise<void> => {
    const { selectedPolicy, archiveReason } = this.state;
    if (!selectedPolicy) return;

    try {
      this.setState({ submitting: true });
      await this.policyService.archivePolicy(selectedPolicy.Id, archiveReason);
      this.setState({ showArchiveDialog: false, submitting: false });
      await this.loadPolicies();
      await this.dialogManager.showAlert('Policy archived successfully!', { variant: 'success' });
    } catch (error) {
      console.error('Failed to archive policy:', error);
      await this.dialogManager.showAlert('Failed to archive policy. Please try again.', { variant: 'error' });
      this.setState({ submitting: false });
    }
  };

  private handleBulkApprove = async (): Promise<void> => {
    const { selectedPolicies } = this.state;
    const eligiblePolicies = selectedPolicies.filter(p => p.Status === PolicyStatus.InReview);

    if (eligiblePolicies.length === 0) {
      await this.dialogManager.showAlert('No policies in "In Review" status selected.', { variant: 'warning' });
      return;
    }

    const confirmed = await this.dialogManager.showConfirm(
      `Are you sure you want to approve ${eligiblePolicies.length} policies?`,
      { title: 'Bulk Approve', confirmText: 'Approve All', cancelText: 'Cancel' }
    );
    if (!confirmed) return;

    try {
      this.setState({ submitting: true });
      let successCount = 0;
      let failCount = 0;

      for (const policy of eligiblePolicies) {
        try {
          await this.policyService.approvePolicy(policy.Id, 'Bulk approved by admin');
          successCount++;
        } catch {
          failCount++;
        }
      }

      await this.loadPolicies();
      this.setState({ submitting: false, selectedPolicies: [] });
      this.selection.setAllSelected(false);

      if (failCount > 0) {
        await this.dialogManager.showAlert(
          `${successCount} policies approved, ${failCount} failed.`,
          { variant: 'warning' }
        );
      } else {
        await this.dialogManager.showAlert(`${successCount} policies approved!`, { variant: 'success' });
      }
    } catch (error) {
      console.error('Failed to bulk approve:', error);
      await this.dialogManager.showAlert('Bulk approve operation failed. Please try again.', { variant: 'error' });
      this.setState({ submitting: false });
    }
  };

  private handleBulkReject = async (): Promise<void> => {
    const { selectedPolicies } = this.state;
    const eligiblePolicies = selectedPolicies.filter(p => p.Status === PolicyStatus.InReview);

    if (eligiblePolicies.length === 0) {
      await this.dialogManager.showAlert('No policies in "In Review" status selected.', { variant: 'warning' });
      return;
    }

    const reason = await this.dialogManager.showPrompt(
      `Provide a reason for rejecting ${eligiblePolicies.length} policies:`,
      {
        title: 'Bulk Reject',
        label: 'Rejection Reason',
        placeholder: 'Enter rejection reason...',
        required: true,
        multiline: true,
        rows: 3
      }
    );
    if (!reason) return;

    try {
      this.setState({ submitting: true });
      let successCount = 0;
      let failCount = 0;

      for (const policy of eligiblePolicies) {
        try {
          await this.policyService.rejectPolicy(policy.Id, reason);
          successCount++;
        } catch {
          failCount++;
        }
      }

      await this.loadPolicies();
      this.setState({ submitting: false, selectedPolicies: [] });
      this.selection.setAllSelected(false);

      if (failCount > 0) {
        await this.dialogManager.showAlert(
          `${successCount} policies rejected, ${failCount} failed.`,
          { variant: 'warning' }
        );
      } else {
        await this.dialogManager.showAlert(`${successCount} policies rejected.`, { variant: 'info' });
      }
    } catch (error) {
      console.error('Failed to bulk reject:', error);
      await this.dialogManager.showAlert('Bulk reject operation failed. Please try again.', { variant: 'error' });
      this.setState({ submitting: false });
    }
  };

  private handleBulkArchive = async (): Promise<void> => {
    const { selectedPolicies } = this.state;
    const eligiblePolicies = selectedPolicies.filter(p => p.Status === PolicyStatus.Published);

    if (eligiblePolicies.length === 0) {
      await this.dialogManager.showAlert('No policies in "Published" status selected.', { variant: 'warning' });
      return;
    }

    const reason = await this.dialogManager.showPrompt(
      `Provide a reason for archiving ${eligiblePolicies.length} policies:`,
      {
        title: 'Bulk Archive',
        label: 'Archive Reason',
        placeholder: 'Enter archive reason...',
        required: true,
        multiline: true,
        rows: 3
      }
    );
    if (!reason) return;

    try {
      this.setState({ submitting: true });
      let successCount = 0;
      let failCount = 0;

      for (const policy of eligiblePolicies) {
        try {
          await this.policyService.archivePolicy(policy.Id, reason);
          successCount++;
        } catch {
          failCount++;
        }
      }

      await this.loadPolicies();
      this.setState({ submitting: false, selectedPolicies: [] });
      this.selection.setAllSelected(false);

      if (failCount > 0) {
        await this.dialogManager.showAlert(
          `${successCount} policies archived, ${failCount} failed.`,
          { variant: 'warning' }
        );
      } else {
        await this.dialogManager.showAlert(`${successCount} policies archived.`, { variant: 'success' });
      }
    } catch (error) {
      console.error('Failed to bulk archive:', error);
      await this.dialogManager.showAlert('Bulk archive operation failed. Please try again.', { variant: 'error' });
      this.setState({ submitting: false });
    }
  };

  private handleBulkPublish = async (): Promise<void> => {
    const { selectedPolicies } = this.state;
    const eligiblePolicies = selectedPolicies.filter(p => p.Status === PolicyStatus.Approved);

    if (eligiblePolicies.length === 0) {
      await this.dialogManager.showAlert('No policies in "Approved" status selected.', { variant: 'warning' });
      return;
    }

    const confirmed = await this.dialogManager.showConfirm(
      `Are you sure you want to publish ${eligiblePolicies.length} policies to all employees?`,
      { title: 'Bulk Publish', confirmText: 'Publish All', cancelText: 'Cancel' }
    );
    if (!confirmed) return;

    try {
      this.setState({ submitting: true });
      let successCount = 0;
      let failCount = 0;

      const dueDate = new Date();
      dueDate.setDate(dueDate.getDate() + 14); // Default 14 days

      for (const policy of eligiblePolicies) {
        try {
          await this.policyService.publishPolicy({
            policyId: policy.Id,
            distributionScope: DistributionScope.AllEmployees,
            dueDate,
            sendNotifications: true
          });
          successCount++;
        } catch {
          failCount++;
        }
      }

      await this.loadPolicies();
      this.setState({ submitting: false, selectedPolicies: [] });
      this.selection.setAllSelected(false);

      if (failCount > 0) {
        await this.dialogManager.showAlert(
          `${successCount} policies published, ${failCount} failed.`,
          { variant: 'warning' }
        );
      } else {
        await this.dialogManager.showAlert(`${successCount} policies published!`, { variant: 'success' });
      }
    } catch (error) {
      console.error('Failed to bulk publish:', error);
      await this.dialogManager.showAlert('Bulk publish operation failed. Please try again.', { variant: 'error' });
      this.setState({ submitting: false });
    }
  };

  private renderModuleNav(): JSX.Element {
    // Policy Builder tabs: Browse, Admin (this), Create Policy, Policy Packs, Quiz Builder
    const navItems = [
      { key: 'browse', text: 'Browse Policies', icon: 'Library', url: '/SitePages/PolicyHub.aspx', isActive: false },
      { key: 'admin', text: 'Policy Admin', icon: 'Settings', url: '', isActive: true },
      { key: 'author', text: 'Create Policy', icon: 'Edit', url: '/SitePages/PolicyAuthor.aspx', isActive: false },
      { key: 'packs', text: 'Policy Packs', icon: 'Package', url: '/SitePages/PolicyPackManager.aspx', isActive: false },
      { key: 'quiz', text: 'Quiz Builder', icon: 'Questionnaire', url: '/SitePages/QuizBuilder.aspx', isActive: false }
    ];

    return (
      <div className={styles.moduleNav}>
        <Stack horizontal tokens={{ childrenGap: 4 }} wrap>
          {navItems.map(item => (
            <DefaultButton
              key={item.key}
              text={item.text}
              iconProps={{ iconName: item.icon }}
              className={item.isActive ? styles.moduleNavActive : styles.moduleNavButton}
              onClick={() => item.url && (window.location.href = item.url)}
              disabled={item.isActive}
            />
          ))}
        </Stack>
      </div>
    );
  }

  private renderCommandBar(): JSX.Element {
    const { selectedPolicies, submitting, selectedView } = this.state;
    const { enableBulkOperations } = this.props;

    const items: ICommandBarItemProps[] = [
      {
        key: 'refresh',
        text: 'Refresh',
        iconProps: { iconName: 'Refresh' },
        onClick: () => { void this.loadPolicies(); }
      }
    ];

    if (enableBulkOperations && selectedPolicies.length > 0) {
      // Count eligible policies for each operation
      const inReviewCount = selectedPolicies.filter(p => p.Status === PolicyStatus.InReview).length;
      const approvedCount = selectedPolicies.filter(p => p.Status === PolicyStatus.Approved).length;
      const publishedCount = selectedPolicies.filter(p => p.Status === PolicyStatus.Published).length;

      // Show relevant bulk operations based on current view and selection
      if (inReviewCount > 0 && (selectedView === 'pending' || selectedView === 'all')) {
        items.push({
          key: 'bulkApprove',
          text: `Approve (${inReviewCount})`,
          iconProps: { iconName: 'CheckMark' },
          onClick: () => { void this.handleBulkApprove(); },
          disabled: submitting
        });
        items.push({
          key: 'bulkReject',
          text: `Reject (${inReviewCount})`,
          iconProps: { iconName: 'Cancel' },
          onClick: () => { void this.handleBulkReject(); },
          disabled: submitting
        });
      }

      if (approvedCount > 0 && (selectedView === 'approved' || selectedView === 'all')) {
        items.push({
          key: 'bulkPublish',
          text: `Publish (${approvedCount})`,
          iconProps: { iconName: 'PublishContent' },
          onClick: () => { void this.handleBulkPublish(); },
          disabled: submitting
        });
      }

      if (publishedCount > 0 && (selectedView === 'published' || selectedView === 'all')) {
        items.push({
          key: 'bulkArchive',
          text: `Archive (${publishedCount})`,
          iconProps: { iconName: 'Archive' },
          onClick: () => { void this.handleBulkArchive(); },
          disabled: submitting
        });
      }
    }

    // Show selection count in far items
    const farItems: ICommandBarItemProps[] = [];
    if (enableBulkOperations && selectedPolicies.length > 0) {
      farItems.push({
        key: 'selectionCount',
        text: `${selectedPolicies.length} selected`,
        iconProps: { iconName: 'CheckboxComposite' },
        onClick: () => {
          this.selection.setAllSelected(false);
          this.setState({ selectedPolicies: [] });
        }
      });
    }

    return <CommandBar items={items} farItems={farItems} />;
  }

  private renderPoliciesList(filterStatus: PolicyStatus[]): JSX.Element {
    const { policies } = this.state;

    const filteredPolicies = policies.filter(p => filterStatus.includes(p.Status));

    const columns: IColumn[] = [
      {
        key: 'policyNumber',
        name: 'Number',
        fieldName: 'PolicyNumber',
        minWidth: 100,
        maxWidth: 120,
        isResizable: true
      },
      {
        key: 'policyName',
        name: 'Name',
        fieldName: 'PolicyName',
        minWidth: 200,
        maxWidth: 300,
        isResizable: true
      },
      {
        key: 'category',
        name: 'Category',
        fieldName: 'PolicyCategory',
        minWidth: 120,
        maxWidth: 150,
        isResizable: true
      },
      {
        key: 'status',
        name: 'Status',
        fieldName: 'Status',
        minWidth: 100,
        maxWidth: 120,
        isResizable: true,
        onRender: (item: IPolicy) => {
          const statusColor = item.Status === PolicyStatus.Published ? '#107C10' :
                             item.Status === PolicyStatus.Draft ? '#605E5C' :
                             item.Status === PolicyStatus.InReview ? '#FFA500' :
                             item.Status === PolicyStatus.Approved ? '#0078D4' : '#D13438';
          return (
            <div
              className={styles.statusBadge}
              style={{ backgroundColor: statusColor }}
            >
              {item.Status}
            </div>
          );
        }
      },
      {
        key: 'version',
        name: 'Version',
        fieldName: 'VersionNumber',
        minWidth: 80,
        maxWidth: 100,
        isResizable: true
      },
      {
        key: 'actions',
        name: 'Actions',
        minWidth: 200,
        maxWidth: 250,
        onRender: (item: IPolicy) => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            {item.Status === PolicyStatus.InReview && (
              <>
                <DefaultButton
                  text="Approve"
                  iconProps={{ iconName: 'CheckMark' }}
                  onClick={() => this.handleApprove(item)}
                />
                <DefaultButton
                  text="Reject"
                  iconProps={{ iconName: 'Cancel' }}
                  onClick={() => this.handleReject(item)}
                />
              </>
            )}
            {item.Status === PolicyStatus.Approved && (
              <PrimaryButton
                text="Publish"
                iconProps={{ iconName: 'PublishContent' }}
                onClick={() => this.handlePublish(item)}
              />
            )}
            {item.Status === PolicyStatus.Published && (
              <DefaultButton
                text="Archive"
                iconProps={{ iconName: 'Archive' }}
                onClick={() => this.handleArchive(item)}
              />
            )}
            <IconButton
              iconProps={{ iconName: 'View' }}
              title="View"
              onClick={() => window.open(`?policyId=${item.Id}`, '_blank')}
            />
          </Stack>
        )
      }
    ];

    return (
      <div className={styles.policiesList}>
        {filteredPolicies.length === 0 ? (
          <MessageBar messageBarType={MessageBarType.info}>
            No policies found in this category.
          </MessageBar>
        ) : (
          <DetailsList
            items={filteredPolicies}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selection={this.selection}
            selectionMode={this.props.enableBulkOperations ? SelectionMode.multiple : SelectionMode.none}
            isHeaderVisible={true}
          />
        )}
      </div>
    );
  }

  private renderPublishPanel(): JSX.Element {
    const {
      showPublishPanel,
      selectedPolicy,
      publishTargetUserIds,
      publishTargetEmails,
      publishTargetDepartments,
      publishTargetRoles,
      publishDueDate,
      submitting
    } = this.state;

    const panelStyles = {
      main: {
        fontFamily: '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif'
      },
      content: {
        padding: '24px'
      }
    };

    const footerStyles: React.CSSProperties = {
      display: 'flex',
      justifyContent: 'flex-end',
      gap: '8px',
      padding: '16px 24px',
      borderTop: '1px solid #edebe9'
    };

    return (
      <Panel
        isOpen={showPublishPanel}
        onDismiss={() => this.setState({ showPublishPanel: false })}
        type={PanelType.medium}
        headerText={`Publish: ${selectedPolicy?.PolicyName}`}
        isFooterAtBottom={true}
        onRenderFooterContent={() => (
          <div style={footerStyles}>
            <DefaultButton
              text="Cancel"
              onClick={() => this.setState({ showPublishPanel: false })}
              disabled={submitting}
            />
            <PrimaryButton
              text="Publish"
              onClick={this.handleSubmitPublish}
              disabled={submitting}
            />
          </div>
        )}
        styles={panelStyles}
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <MessageBar messageBarType={MessageBarType.info}>
            Specify target users by providing user IDs, email addresses, departments, or roles.
          </MessageBar>

          <TextField
            label="User IDs (comma-separated)"
            multiline
            rows={2}
            value={publishTargetUserIds}
            onChange={(e, value) => this.setState({ publishTargetUserIds: value || '' })}
            placeholder="e.g., 123, 456, 789"
          />

          <TextField
            label="Email Addresses (comma-separated)"
            multiline
            rows={2}
            value={publishTargetEmails}
            onChange={(e, value) => this.setState({ publishTargetEmails: value || '' })}
            placeholder="e.g., john@contoso.com, jane@contoso.com"
          />

          <TextField
            label="Departments (comma-separated)"
            value={publishTargetDepartments}
            onChange={(e, value) => this.setState({ publishTargetDepartments: value || '' })}
            placeholder="e.g., IT, HR, Sales"
          />

          <TextField
            label="Roles (comma-separated)"
            value={publishTargetRoles}
            onChange={(e, value) => this.setState({ publishTargetRoles: value || '' })}
            placeholder="e.g., Manager, Developer, Analyst"
          />

          <TextField
            label="Acknowledgement Due Date"
            type="date"
            value={publishDueDate}
            onChange={(e, value) => this.setState({ publishDueDate: value || '' })}
          />
        </Stack>
      </Panel>
    );
  }

  private renderArchiveDialog(): JSX.Element {
    const {
      showArchiveDialog,
      selectedPolicy,
      archiveReason,
      submitting
    } = this.state;

    return (
      <Dialog
        hidden={!showArchiveDialog}
        onDismiss={() => this.setState({ showArchiveDialog: false })}
        dialogContentProps={{
          type: DialogType.normal,
          title: `Archive: ${selectedPolicy?.PolicyName}`,
          subText: 'Provide a reason for archiving this policy'
        }}
      >
        <TextField
          label="Reason"
          multiline
          rows={4}
          value={archiveReason}
          onChange={(e, value) => this.setState({ archiveReason: value || '' })}
          placeholder="Enter reason for archiving..."
          required
        />

        <DialogFooter>
          <PrimaryButton
            text="Archive"
            onClick={this.handleSubmitArchive}
            disabled={submitting || !archiveReason}
          />
          <DefaultButton
            text="Cancel"
            onClick={() => this.setState({ showArchiveDialog: false })}
            disabled={submitting}
          />
        </DialogFooter>
      </Dialog>
    );
  }

  public render(): React.ReactElement<IPolicyAdminProps> {
    const { loading, error, selectedView } = this.state;

    return (
      <JmlAppLayout
        context={this.props.context}
        pageTitle="Policy Administration"
        pageDescription="Review, approve and manage policy lifecycle"
        pageIcon="Admin"
        breadcrumbs={[{ text: 'JML Portal', url: '/sites/JML' }, { text: 'Policy Administration' }]}
        activeNavKey="policies"
        showQuickLinks={true}
        showSearch={true}
        showNotifications={true}
        compactFooter={true}
      >
        <section className={styles.policyAdmin}>
          <Stack tokens={{ childrenGap: 24 }}>
            {this.renderModuleNav()}
            {this.renderCommandBar()}

            {loading && (
              <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
                <Spinner size={SpinnerSize.large} label="Loading policies..." />
              </Stack>
            )}

            {error && (
              <MessageBar messageBarType={MessageBarType.error} isMultiline>
                {error}
              </MessageBar>
            )}

            {!loading && !error && (
              <Pivot
                selectedKey={selectedView}
                onLinkClick={(item) => this.setState({ selectedView: item?.props.itemKey || 'pending' })}
              >
                <PivotItem headerText="Pending Review" itemKey="pending" itemIcon="DocumentReply">
                  {this.renderPoliciesList([PolicyStatus.InReview])}
                </PivotItem>
                <PivotItem headerText="Approved" itemKey="approved" itemIcon="Completed">
                  {this.renderPoliciesList([PolicyStatus.Approved])}
                </PivotItem>
                <PivotItem headerText="Published" itemKey="published" itemIcon="PublishContent">
                  {this.renderPoliciesList([PolicyStatus.Published])}
                </PivotItem>
                <PivotItem headerText="Archived" itemKey="archived" itemIcon="Archive">
                  {this.renderPoliciesList([PolicyStatus.Archived])}
                </PivotItem>
                <PivotItem headerText="All Policies" itemKey="all" itemIcon="BulletedList">
                  {this.renderPoliciesList([PolicyStatus.Draft, PolicyStatus.InReview, PolicyStatus.Approved, PolicyStatus.Published, PolicyStatus.Archived, PolicyStatus.Expired])}
                </PivotItem>
              </Pivot>
            )}
          </Stack>

          {this.renderPublishPanel()}
          {this.renderArchiveDialog()}
          <this.dialogManager.DialogComponent />
        </section>
      </JmlAppLayout>
    );
  }
}
