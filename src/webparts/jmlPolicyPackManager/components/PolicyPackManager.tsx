// @ts-nocheck
/* eslint-disable */
import * as React from 'react';
import { IPolicyPackManagerProps } from './IPolicyPackManagerProps';
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
  TextField,
  Dropdown,
  IDropdownOption,
  Checkbox,
  Label,
  Panel,
  PanelType
} from '@fluentui/react';
import { injectPortalStyles } from '../../../utils/injectPortalStyles';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { PageSubheader } from '../../../components/PageSubheader';
import { PolicyPackService } from '../../../services/PolicyPackService';
import { PolicyService } from '../../../services/PolicyService';
import { createDialogManager } from '../../../hooks/useDialog';
import {
  IPolicyPack,
  IPolicy,
  PolicyStatus,
  ICreatePolicyPackRequest,
  IAssignPolicyPackRequest,
  IPolicyPackDeploymentResult
} from '../../../models/IPolicy';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import styles from './PolicyPackManager.module.scss';

export interface IPolicyPackManagerState {
  loading: boolean;
  error: string | null;
  policyPacks: IPolicyPack[];
  allPolicies: IPolicy[];
  showCreatePanel: boolean;
  showAssignPanel: boolean;
  selectedPack: IPolicyPack | null;
  editingPackId: number | null;
  newPackName: string;
  newPackDescription: string;
  newPackType: string;
  selectedPolicyIds: number[];
  targetProcessType: string;
  assignmentTargetUserIds: string;
  assignmentTargetEmails: string;
  assignmentDepartments: string;
  assignmentRoles: string;
  isSequential: boolean;
  sendWelcomeEmail: boolean;
  sendTeamsNotification: boolean;
  submitting: boolean;
  deploymentResult: IPolicyPackDeploymentResult | null;
  policySearchFilter: string;
  recentPoliciesExpanded: boolean;
  deliveryOptionsExpanded: boolean;
}

export default class PolicyPackManager extends React.Component<IPolicyPackManagerProps, IPolicyPackManagerState> {
  private packService: PolicyPackService;
  private policyService: PolicyService;
  private dialogManager = createDialogManager();

  constructor(props: IPolicyPackManagerProps) {
    super(props);
    this.state = {
      loading: true,
      error: null,
      policyPacks: [],
      allPolicies: [],
      showCreatePanel: false,
      showAssignPanel: false,
      selectedPack: null,
      editingPackId: null,
      newPackName: '',
      newPackDescription: '',
      newPackType: 'Onboarding',
      selectedPolicyIds: [],
      targetProcessType: '',
      assignmentTargetUserIds: '',
      assignmentTargetEmails: '',
      assignmentDepartments: '',
      assignmentRoles: '',
      isSequential: false,
      sendWelcomeEmail: true,
      sendTeamsNotification: true,
      submitting: false,
      deploymentResult: null,
      policySearchFilter: '',
      recentPoliciesExpanded: true,
      deliveryOptionsExpanded: false
    };
    this.packService = new PolicyPackService(props.sp);
    this.policyService = new PolicyService(props.sp);
  }

  public async componentDidMount(): Promise<void> {
    injectPortalStyles();
    await this.loadData();
  }

  private async loadData(): Promise<void> {
    try {
      this.setState({ loading: true, error: null });
      await this.packService.initialize();
      await this.policyService.initialize();

      const packs = await this.packService.getPolicyPacks();
      const allPolicies = await this.policyService.getAllPolicies();
      const policies = allPolicies.filter((p: IPolicy) => p.PolicyStatus === PolicyStatus.Published);

      this.setState({
        policyPacks: packs,
        allPolicies: policies,
        loading: false
      });
    } catch (error) {
      console.error('Failed to load data:', error);
      this.setState({
        error: 'Failed to load policy packs. Please try again later.',
        loading: false
      });
    }
  }

  private handleCreatePack = (): void => {
    this.setState({
      showCreatePanel: true,
      editingPackId: null,
      newPackName: '',
      newPackDescription: '',
      newPackType: 'Onboarding',
      selectedPolicyIds: [],
      isSequential: false,
      sendWelcomeEmail: true,
      sendTeamsNotification: true
    });
  };

  private handleEditPack = (pack: IPolicyPack): void => {
    this.setState({
      showCreatePanel: true,
      editingPackId: pack.Id,
      newPackName: pack.PackName,
      newPackDescription: pack.PackDescription || '',
      newPackType: pack.PackType || 'Onboarding',
      selectedPolicyIds: pack.PolicyIds || [],
      isSequential: pack.IsSequential || false,
      sendWelcomeEmail: pack.SendWelcomeEmail ?? true,
      sendTeamsNotification: pack.SendTeamsNotification ?? true
    });
  };

  private handleSubmitCreate = async (): Promise<void> => {
    const {
      editingPackId,
      newPackName,
      newPackDescription,
      newPackType,
      selectedPolicyIds,
      targetProcessType,
      isSequential,
      sendWelcomeEmail,
      sendTeamsNotification
    } = this.state;

    if (!newPackName || selectedPolicyIds.length === 0) {
      await this.dialogManager.showAlert('Pack name and at least one policy are required.', { variant: 'warning' });
      return;
    }

    try {
      this.setState({ submitting: true });

      const request: ICreatePolicyPackRequest = {
        packName: newPackName,
        packDescription: newPackDescription,
        packType: newPackType as any,
        policyIds: selectedPolicyIds,
        targetProcessType: targetProcessType as any,
        isSequential,
        sendWelcomeEmail,
        sendTeamsNotification
      };

      if (editingPackId) {
        // Update existing pack
        await this.packService.updatePolicyPack(editingPackId, request);
      } else {
        // Create new pack
        await this.packService.createPolicyPack(request);
      }

      this.setState({
        showCreatePanel: false,
        editingPackId: null,
        submitting: false
      });

      await this.loadData();
      await this.dialogManager.showAlert(
        editingPackId ? 'Policy pack updated successfully!' : 'Policy pack created successfully!',
        { variant: 'success' }
      );
    } catch (error) {
      console.error('Failed to save policy pack:', error);
      this.setState({
        error: editingPackId ? 'Failed to update policy pack. Please try again.' : 'Failed to create policy pack. Please try again.',
        submitting: false
      });
    }
  };

  private handleAssignPack = (pack: IPolicyPack): void => {
    this.setState({
      selectedPack: pack,
      showAssignPanel: true,
      assignmentTargetUserIds: '',
      assignmentTargetEmails: '',
      assignmentDepartments: '',
      assignmentRoles: '',
      deploymentResult: null
    });
  };

  private handleSubmitAssignment = async (): Promise<void> => {
    const {
      selectedPack,
      assignmentTargetUserIds,
      assignmentTargetEmails,
      assignmentDepartments,
      assignmentRoles
    } = this.state;

    if (!selectedPack) return;

    // Parse inputs
    const userIds = assignmentTargetUserIds
      ? assignmentTargetUserIds.split(',').map(id => parseInt(id.trim(), 10)).filter(id => !isNaN(id))
      : undefined;

    const emails = assignmentTargetEmails
      ? assignmentTargetEmails.split(',').map(email => email.trim()).filter(e => e)
      : undefined;

    const departments = assignmentDepartments
      ? assignmentDepartments.split(',').map(d => d.trim()).filter(d => d)
      : undefined;

    const roles = assignmentRoles
      ? assignmentRoles.split(',').map(r => r.trim()).filter(r => r)
      : undefined;

    if (!userIds && !emails && !departments && !roles) {
      await this.dialogManager.showAlert('Please specify at least one target: user IDs, emails, departments, or roles.', { variant: 'warning' });
      return;
    }

    try {
      this.setState({ submitting: true });

      const request: IAssignPolicyPackRequest = {
        packId: selectedPack.Id,
        targetUserIds: userIds,
        targetEmails: emails,
        targetDepartments: departments,
        targetRoles: roles
      };

      const result = await this.packService.assignPolicyPack(request);

      this.setState({
        deploymentResult: result,
        submitting: false
      });

      const successCount = Array.isArray(result.successfulAssignments)
        ? result.successfulAssignments.length
        : result.successfulAssignments;
      const failedCount = Array.isArray(result.failedAssignments)
        ? result.failedAssignments.length
        : result.failedAssignments;

      if (failedCount === 0) {
        await this.dialogManager.showAlert(`Successfully assigned policy pack to ${successCount} users!`, { variant: 'success' });
      } else {
        await this.dialogManager.showAlert(
          `Assigned to ${successCount} users. ${failedCount} assignments failed.`,
          { variant: 'warning' }
        );
      }
    } catch (error) {
      console.error('Failed to assign policy pack:', error);
      this.setState({
        error: 'Failed to assign policy pack. Please try again.',
        submitting: false
      });
    }
  };

  private handleDeletePack = async (packId: number): Promise<void> => {
    const confirmed = await this.dialogManager.showConfirm(
      'Are you sure you want to delete this policy pack?',
      { title: 'Delete Policy Pack', confirmText: 'Delete', cancelText: 'Cancel', isDanger: true }
    );
    if (!confirmed) return;

    try {
      await this.packService.deletePolicyPack(packId);
      await this.loadData();
      await this.dialogManager.showAlert('Policy pack deleted successfully!', { variant: 'success' });
    } catch (error) {
      console.error('Failed to delete policy pack:', error);
      await this.dialogManager.showAlert('Failed to delete policy pack. Please try again.', { variant: 'error' });
    }
  };

  private renderModuleNav(): JSX.Element {
    // Policy Builder tabs: Browse, Admin, Create Policy, Policy Packs (this), Quiz Builder
    const navItems = [
      { key: 'browse', text: 'Browse Policies', icon: 'Library', url: '/SitePages/PolicyHub.aspx', isActive: false },
      { key: 'admin', text: 'Policy Admin', icon: 'Settings', url: '/SitePages/PolicyAdmin.aspx', isActive: false },
      { key: 'author', text: 'Create Policy', icon: 'Edit', url: '/SitePages/PolicyAuthor.aspx', isActive: false },
      { key: 'packs', text: 'Policy Packs', icon: 'Package', url: '', isActive: true },
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
    const items: ICommandBarItemProps[] = [
      {
        key: 'create',
        text: 'Create Pack',
        iconProps: { iconName: 'Add' },
        onClick: this.handleCreatePack
      },
      {
        key: 'refresh',
        text: 'Refresh',
        iconProps: { iconName: 'Refresh' },
        onClick: () => this.loadData()
      }
    ];

    return <CommandBar items={items} />;
  }

  private renderPolicyPacksList(): JSX.Element {
    const { policyPacks } = this.state;

    if (policyPacks.length === 0) {
      return (
        <div className={styles.emptyState}>
          <Stack tokens={{ childrenGap: 16 }} horizontalAlign="center">
            <Icon iconName="BulletedList" className={styles.emptyIcon} />
            <Text variant="xLarge">No Policy Packs</Text>
            <Text variant="medium" className={styles.subText}>
              Create your first policy pack to get started
            </Text>
            <PrimaryButton
              text="Create Policy Pack"
              iconProps={{ iconName: 'Add' }}
              onClick={this.handleCreatePack}
            />
          </Stack>
        </div>
      );
    }

    return (
      <div className={styles.packsGrid}>
        {policyPacks.map((pack: IPolicyPack) => (
          <div key={pack.Id} className={styles.packCard}>
            <Stack tokens={{ childrenGap: 12 }}>
              <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                <Text variant="large" className={styles.packTitle}>
                  {pack.PackName}
                </Text>
                <div className={styles.packTypeBadge}>
                  {pack.PackType}
                </div>
              </Stack>

              {pack.PackDescription && (
                <Text variant="small" className={styles.description}>
                  {pack.PackDescription}
                </Text>
              )}

              <Stack horizontal tokens={{ childrenGap: 16 }}>
                <Stack tokens={{ childrenGap: 4 }}>
                  <Text variant="small" className={styles.statLabel}>Policies</Text>
                  <Text variant="medium" className={styles.statValue}>
                    {pack.PolicyIds ? pack.PolicyIds.length : 0}
                  </Text>
                </Stack>
                <Stack tokens={{ childrenGap: 4 }}>
                  <Text variant="small" className={styles.statLabel}>Type</Text>
                  <Text variant="medium" className={styles.statValue}>
                    {pack.IsSequential ? 'Sequential' : 'Parallel'}
                  </Text>
                </Stack>
              </Stack>

              <Stack horizontal tokens={{ childrenGap: 8 }}>
                {pack.SendWelcomeEmail && (
                  <div className={styles.featureBadge}>
                    <Icon iconName="Mail" /> Email
                  </div>
                )}
                {pack.SendTeamsNotification && (
                  <div className={styles.featureBadge}>
                    <Icon iconName="TeamsLogo" /> Teams
                  </div>
                )}
              </Stack>

              <Stack horizontal tokens={{ childrenGap: 8 }} className={styles.cardActions}>
                <DefaultButton
                  text="Assign"
                  iconProps={{ iconName: 'PeopleAdd' }}
                  onClick={() => this.handleAssignPack(pack)}
                />
                <DefaultButton
                  text="Edit"
                  iconProps={{ iconName: 'Edit' }}
                  onClick={() => this.handleEditPack(pack)}
                />
                <IconButton
                  iconProps={{ iconName: 'Delete' }}
                  title="Delete"
                  onClick={() => this.handleDeletePack(pack.Id)}
                />
              </Stack>
            </Stack>
          </div>
        ))}
      </div>
    );
  }

  private renderCreatePanel(): JSX.Element {
    const {
      showCreatePanel,
      editingPackId,
      newPackName,
      newPackDescription,
      newPackType,
      selectedPolicyIds,
      allPolicies,
      isSequential,
      sendWelcomeEmail,
      sendTeamsNotification,
      submitting
    } = this.state;

    const isEditing = editingPackId !== null;

    const packTypeOptions: IDropdownOption[] = [
      { key: 'Onboarding', text: 'Onboarding' },
      { key: 'Department', text: 'Department' },
      { key: 'Role', text: 'Role' },
      { key: 'Location', text: 'Location' },
      { key: 'Custom', text: 'Custom' }
    ];

    return (
      <Panel
        isOpen={showCreatePanel}
        onDismiss={() => this.setState({ showCreatePanel: false, editingPackId: null })}
        type={PanelType.medium}
        headerText={isEditing ? 'Edit Policy Pack' : 'Create Policy Pack'}
        closeButtonAriaLabel="Close"
        isFooterAtBottom={true}
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 0', borderTop: '1px solid #edebe9' }}>
            <PrimaryButton
              text={isEditing ? 'Save Changes' : 'Create Pack'}
              onClick={this.handleSubmitCreate}
              disabled={submitting || !newPackName || selectedPolicyIds.length === 0}
            />
            <DefaultButton
              text="Cancel"
              onClick={() => this.setState({ showCreatePanel: false, editingPackId: null })}
              disabled={submitting}
            />
          </Stack>
        )}
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <Text variant="small" style={{ color: '#605e5c' }}>
            Bundle policies together for easy assignment
          </Text>

          <TextField
            label="Pack Name"
            required
            value={newPackName}
            onChange={(_, value) => this.setState({ newPackName: value || '' })}
            placeholder="e.g., New Hire Onboarding Pack"
          />

          <TextField
            label="Description"
            multiline
            rows={3}
            value={newPackDescription}
            onChange={(_, value) => this.setState({ newPackDescription: value || '' })}
            placeholder="Describe the purpose of this policy pack..."
          />

          <Dropdown
            label="Pack Type"
            selectedKey={newPackType}
            options={packTypeOptions}
            onChange={(_, option) => this.setState({ newPackType: option?.key as string })}
          />

          <div>
            <Label required>Select Policies</Label>

            {/* Recently Created Policies — Collapsible */}
            {(() => {
              const recentPolicies = [...allPolicies]
                .sort((a, b) => {
                  const dateA = a.Created ? new Date(a.Created).getTime() : 0;
                  const dateB = b.Created ? new Date(b.Created).getTime() : 0;
                  return dateB - dateA;
                })
                .slice(0, 5);
              return recentPolicies.length > 0 ? (
                <div style={{ marginBottom: 12 }}>
                  <div
                    onClick={() => this.setState({ recentPoliciesExpanded: !this.state.recentPoliciesExpanded })}
                    style={{
                      display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer',
                      padding: '6px 0', marginBottom: this.state.recentPoliciesExpanded ? 6 : 0
                    }}
                  >
                    <Icon
                      iconName={this.state.recentPoliciesExpanded ? 'ChevronDown' : 'ChevronRight'}
                      style={{ fontSize: 12, color: '#605e5c', transition: 'transform 0.15s' }}
                    />
                    <Text variant="small" style={{ color: '#605e5c', fontWeight: 600 }}>
                      Recently Created
                    </Text>
                    <Text variant="small" style={{ color: '#a0aec0', fontSize: 11 }}>
                      ({recentPolicies.length})
                    </Text>
                  </div>
                  {this.state.recentPoliciesExpanded && (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                      {recentPolicies.map((policy: IPolicy) => {
                        const isSelected = selectedPolicyIds.includes(policy.Id);
                        return (
                          <div
                            key={policy.Id}
                            onClick={() => {
                              if (isSelected) {
                                this.setState({ selectedPolicyIds: selectedPolicyIds.filter(id => id !== policy.Id) });
                              } else {
                                this.setState({ selectedPolicyIds: [...selectedPolicyIds, policy.Id] });
                              }
                            }}
                            style={{
                              display: 'flex', alignItems: 'center', gap: 8, padding: '6px 10px',
                              borderRadius: 4, cursor: 'pointer', fontSize: 13,
                              background: isSelected ? '#e6f7f5' : '#f8f9fa',
                              border: isSelected ? '1px solid #0d9488' : '1px solid #e2e8f0',
                            }}
                          >
                            <Icon iconName={isSelected ? 'CheckboxComposite' : 'Checkbox'} style={{ color: isSelected ? '#0d9488' : '#8a8886', fontSize: 14 }} />
                            <span style={{ flex: 1, color: '#323130' }}>{policy.PolicyNumber} - {policy.PolicyName}</span>
                            {policy.Created && (
                              <span style={{ color: '#a0aec0', fontSize: 11 }}>
                                {new Date(policy.Created).toLocaleDateString()}
                              </span>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              ) : null;
            })()}

            {/* Filter-as-you-type search */}
            <TextField
              placeholder="Search policies by name or number..."
              iconProps={{ iconName: 'Search' }}
              value={this.state.policySearchFilter}
              onChange={(_, value) => this.setState({ policySearchFilter: value || '' })}
              styles={{ root: { marginBottom: 8 } }}
            />

            {/* Search results — only shown when user types a filter */}
            {this.state.policySearchFilter.trim().length > 0 && (() => {
              const search = this.state.policySearchFilter.toLowerCase();
              const filtered = allPolicies.filter((policy: IPolicy) =>
                (policy.PolicyName || '').toLowerCase().includes(search) ||
                (policy.PolicyNumber || '').toLowerCase().includes(search) ||
                (policy.Category || '').toLowerCase().includes(search)
              );
              return (
                <div style={{
                  maxHeight: 300, overflowY: 'auto',
                  border: '1px solid #e2e8f0', borderRadius: 4, padding: 4,
                  background: '#fff', marginBottom: 8
                }}>
                  {filtered.length > 0 ? (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                      {filtered.map((policy: IPolicy) => {
                        const isSelected = selectedPolicyIds.includes(policy.Id);
                        return (
                          <div
                            key={policy.Id}
                            onClick={() => {
                              if (isSelected) {
                                this.setState({ selectedPolicyIds: selectedPolicyIds.filter(id => id !== policy.Id) });
                              } else {
                                this.setState({ selectedPolicyIds: [...selectedPolicyIds, policy.Id] });
                              }
                            }}
                            style={{
                              display: 'flex', alignItems: 'center', gap: 8, padding: '6px 10px',
                              borderRadius: 4, cursor: 'pointer', fontSize: 13,
                              background: isSelected ? '#e6f7f5' : '#f8f9fa',
                              border: isSelected ? '1px solid #0d9488' : '1px solid #e2e8f0',
                            }}
                          >
                            <Icon iconName={isSelected ? 'CheckboxComposite' : 'Checkbox'} style={{ color: isSelected ? '#0d9488' : '#8a8886', fontSize: 14 }} />
                            <span style={{ flex: 1, color: '#323130' }}>{policy.PolicyNumber} - {policy.PolicyName}</span>
                            {policy.Category && (
                              <span style={{
                                color: '#64748b', fontSize: 11, padding: '1px 6px',
                                background: '#f1f5f9', borderRadius: 3
                              }}>
                                {policy.Category}
                              </span>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  ) : (
                    <Text variant="small" style={{ color: '#a0aec0', fontStyle: 'italic', padding: 8, display: 'block', textAlign: 'center' }}>
                      No policies match "{this.state.policySearchFilter}"
                    </Text>
                  )}
                </div>
              );
            })()}

            {/* Selected count + chips */}
            <Text variant="small" className={styles.subText}>
              {selectedPolicyIds.length} policies selected
            </Text>
            {selectedPolicyIds.length > 0 && (
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 6 }}>
                {selectedPolicyIds.map(id => {
                  const policy = allPolicies.find((p: IPolicy) => p.Id === id);
                  if (!policy) return null;
                  return (
                    <div key={id} style={{
                      display: 'inline-flex', alignItems: 'center', gap: 4,
                      background: '#e6f7f5', border: '1px solid #0d9488', borderRadius: 12,
                      padding: '2px 8px 2px 10px', fontSize: 12, color: '#0f766e'
                    }}>
                      <span>{policy.PolicyNumber}</span>
                      <IconButton
                        iconProps={{ iconName: 'Cancel', style: { fontSize: 10 } }}
                        styles={{ root: { width: 18, height: 18, color: '#0f766e' }, rootHovered: { color: '#b91c1c', background: 'transparent' } }}
                        onClick={() => this.setState({ selectedPolicyIds: selectedPolicyIds.filter(pid => pid !== id) })}
                        title="Remove"
                      />
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* ── Delivery & Notification Options ── */}
          <div style={{
            border: '1px solid #e2e8f0', borderRadius: 8,
            background: '#fafbfc', overflow: 'hidden'
          }}>
            <div
              onClick={() => this.setState({ deliveryOptionsExpanded: !this.state.deliveryOptionsExpanded })}
              style={{
                display: 'flex', alignItems: 'center', gap: 8,
                padding: '12px 14px', cursor: 'pointer',
                borderBottom: this.state.deliveryOptionsExpanded ? '1px solid #e2e8f0' : 'none',
                background: '#fff'
              }}
            >
              <Icon
                iconName={this.state.deliveryOptionsExpanded ? 'ChevronDown' : 'ChevronRight'}
                style={{ fontSize: 12, color: '#605e5c' }}
              />
              <Icon iconName="Settings" style={{ fontSize: 14, color: '#0d9488' }} />
              <Text variant="medium" style={{ fontWeight: 600, color: '#323130', flex: 1 }}>
                Delivery &amp; Notification Options
              </Text>
              <Text variant="small" style={{ color: '#a0aec0' }}>
                {isSequential ? 'Sequential' : 'Any order'} &middot; {sendWelcomeEmail ? 'Email' : ''}{sendWelcomeEmail && sendTeamsNotification ? ' + ' : ''}{sendTeamsNotification ? 'Teams' : ''}
              </Text>
            </div>
            {this.state.deliveryOptionsExpanded && (
              <div style={{ padding: '16px 14px', display: 'flex', flexDirection: 'column', gap: 16 }}>

                <Checkbox
                  label="Sequential Acknowledgement (users must complete policies in order)"
                  checked={isSequential}
                  onChange={(_, checked) => this.setState({ isSequential: checked || false })}
                />

                <div style={{ borderTop: '1px solid #e2e8f0', paddingTop: 14 }}>
                  <Text variant="small" style={{ fontWeight: 600, color: '#605e5c', display: 'block', marginBottom: 10 }}>
                    Notifications
                  </Text>
                  <Stack tokens={{ childrenGap: 10 }}>
                    <Checkbox
                      label="Send Welcome Email"
                      checked={sendWelcomeEmail}
                      onChange={(_, checked) => this.setState({ sendWelcomeEmail: checked || false })}
                    />
                    <Checkbox
                      label="Send Teams Notification"
                      checked={sendTeamsNotification}
                      onChange={(_, checked) => this.setState({ sendTeamsNotification: checked || false })}
                    />
                  </Stack>
                </div>

                <div style={{ borderTop: '1px solid #e2e8f0', paddingTop: 14 }}>
                  <Text variant="small" style={{ fontWeight: 600, color: '#605e5c', display: 'block', marginBottom: 4 }}>
                    Approvers
                  </Text>
                  <Text variant="small" style={{ color: '#a0aec0', display: 'block', marginBottom: 10 }}>
                    Add one or more approvers who must sign off before this pack is distributed
                  </Text>
                  <PeoplePicker
                    context={this.props.context as any}
                    personSelectionLimit={10}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={500}
                    placeholder="Search for approvers in Entra ID..."
                    groupName=""
                    showHiddenInUI={false}
                    onChange={(items: any[]) => {
                      const approvers = items.map(item => item.secondaryText || item.text || '').filter(Boolean);
                      console.log('Selected approvers:', approvers);
                    }}
                  />
                </div>

              </div>
            )}
          </div>
        </Stack>
      </Panel>
    );
  }

  private renderAssignPanel(): JSX.Element {
    const {
      showAssignPanel,
      selectedPack,
      assignmentTargetUserIds,
      assignmentTargetEmails,
      assignmentDepartments,
      assignmentRoles,
      submitting,
      deploymentResult
    } = this.state;

    return (
      <Panel
        isOpen={showAssignPanel}
        onDismiss={() => this.setState({ showAssignPanel: false })}
        type={PanelType.medium}
        headerText={`Assign: ${selectedPack?.PackName}`}
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <MessageBar messageBarType={MessageBarType.info}>
            Specify target users by providing user IDs, email addresses, departments, or roles.
            You can use multiple targeting methods.
          </MessageBar>

          <TextField
            label="User IDs (comma-separated)"
            multiline
            rows={2}
            value={assignmentTargetUserIds}
            onChange={(e, value) => this.setState({ assignmentTargetUserIds: value || '' })}
            placeholder="e.g., 123, 456, 789"
          />

          <TextField
            label="Email Addresses (comma-separated)"
            multiline
            rows={2}
            value={assignmentTargetEmails}
            onChange={(e, value) => this.setState({ assignmentTargetEmails: value || '' })}
            placeholder="e.g., john@contoso.com, jane@contoso.com"
          />

          <TextField
            label="Departments (comma-separated)"
            value={assignmentDepartments}
            onChange={(e, value) => this.setState({ assignmentDepartments: value || '' })}
            placeholder="e.g., IT, HR, Sales"
          />

          <TextField
            label="Roles (comma-separated)"
            value={assignmentRoles}
            onChange={(e, value) => this.setState({ assignmentRoles: value || '' })}
            placeholder="e.g., Manager, Developer, Analyst"
          />

          {deploymentResult && (
            <MessageBar
              messageBarType={
                (Array.isArray(deploymentResult.failedAssignments) ? deploymentResult.failedAssignments.length : deploymentResult.failedAssignments) === 0
                  ? MessageBarType.success
                  : MessageBarType.warning
              }
            >
              <Stack tokens={{ childrenGap: 8 }}>
                <Text>
                  Successfully assigned to {Array.isArray(deploymentResult.successfulAssignments) ? deploymentResult.successfulAssignments.length : deploymentResult.successfulAssignments} users
                </Text>
                {(Array.isArray(deploymentResult.failedAssignments) ? deploymentResult.failedAssignments.length : deploymentResult.failedAssignments) > 0 && (
                  <Text>
                    {Array.isArray(deploymentResult.failedAssignments) ? deploymentResult.failedAssignments.length : deploymentResult.failedAssignments} assignments failed
                  </Text>
                )}
                {deploymentResult.emailsSent > 0 && (
                  <Text>{deploymentResult.emailsSent} emails sent</Text>
                )}
                {deploymentResult.teamsNotificationsSent > 0 && (
                  <Text>{deploymentResult.teamsNotificationsSent} Teams notifications sent</Text>
                )}
              </Stack>
            </MessageBar>
          )}

          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text="Assign Pack"
              onClick={this.handleSubmitAssignment}
              disabled={submitting}
            />
            <DefaultButton
              text="Close"
              onClick={() => this.setState({ showAssignPanel: false })}
              disabled={submitting}
            />
          </Stack>
        </Stack>
      </Panel>
    );
  }

  public render(): React.ReactElement<IPolicyPackManagerProps> {
    const { loading, error } = this.state;

    return (
      <JmlAppLayout
        context={this.props.context}
        pageTitle="Policy Builder"
        pageDescription="Create and manage policy packs with bundled documents and templates"
        pageIcon="Edit"
        breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Policy Builder' }]}
        activeNavKey="policies"
        showQuickLinks={true}
        showSearch={true}
        showNotifications={true}
        compactFooter={true}
      >
        <section className={styles.policyPackManager}>
          <Stack tokens={{ childrenGap: 24 }}>
            {/* Module nav removed - now in global header */}
            {this.renderCommandBar()}

            {loading && (
              <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
                <Spinner size={SpinnerSize.large} label="Loading policy packs..." />
              </Stack>
            )}

            {error && (
              <MessageBar messageBarType={MessageBarType.error} isMultiline>
                {error}
              </MessageBar>
            )}

            {!loading && !error && this.renderPolicyPacksList()}
          </Stack>

          {this.renderCreatePanel()}
          {this.renderAssignPanel()}
          <this.dialogManager.DialogComponent />
        </section>
      </JmlAppLayout>
    );
  }
}
