// @ts-nocheck
/* eslint-disable */
import * as React from 'react';
import { IPolicyAuthorProps } from './IPolicyAuthorProps';
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
  TextField,
  Dropdown,
  IDropdownOption,
  Checkbox,
  Label,
  CommandBar,
  ICommandBarItemProps,
  Dialog,
  DialogType,
  DialogFooter,
  Panel,
  PanelType
} from '@fluentui/react';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { injectPortalStyles } from '../../../utils/injectPortalStyles';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { PolicyService } from '../../../services/PolicyService';
import { createDialogManager } from '../../../hooks/useDialog';
import {
  IPolicy,
  PolicyCategory,
  PolicyStatus,
  ComplianceRisk,
  ReadTimeframe
} from '../../../models/IPolicy';
import styles from './PolicyAuthor.module.scss';

export interface IPolicyAuthorState {
  loading: boolean;
  error: string | null;
  saving: boolean;
  policyId: number | null;
  policyNumber: string;
  policyName: string;
  policyCategory: string;
  policySummary: string;
  policyContent: string;
  keyPoints: string[];
  newKeyPoint: string;
  complianceRisk: string;
  readTimeframe: string;
  readTimeframeDays: number;
  requiresAcknowledgement: boolean;
  requiresQuiz: boolean;
  effectiveDate: string;
  expiryDate: string;
  showTemplatePanel: boolean;
  autoSaveEnabled: boolean;
  lastSaved: Date | null;
}

export default class PolicyAuthor extends React.Component<IPolicyAuthorProps, IPolicyAuthorState> {
  private policyService: PolicyService;
  private autoSaveTimer: NodeJS.Timeout | null = null;
  private dialogManager = createDialogManager();

  constructor(props: IPolicyAuthorProps) {
    super(props);

    const urlParams = new URLSearchParams(window.location.search);
    const policyId = urlParams.get('editPolicyId');

    this.state = {
      loading: !!policyId,
      error: null,
      saving: false,
      policyId: policyId ? parseInt(policyId, 10) : null,
      policyNumber: '',
      policyName: '',
      policyCategory: '',
      policySummary: '',
      policyContent: '',
      keyPoints: [],
      newKeyPoint: '',
      complianceRisk: 'Medium',
      readTimeframe: ReadTimeframe.Week1,
      readTimeframeDays: 7,
      requiresAcknowledgement: true,
      requiresQuiz: false,
      effectiveDate: new Date().toISOString().split('T')[0],
      expiryDate: '',
      showTemplatePanel: false,
      autoSaveEnabled: props.enableAutoSave,
      lastSaved: null
    };

    this.policyService = new PolicyService(props.sp);
  }

  public async componentDidMount(): Promise<void> {
    injectPortalStyles();
    await this.policyService.initialize();

    if (this.state.policyId) {
      await this.loadPolicy(this.state.policyId);
    }

    if (this.props.enableAutoSave) {
      this.startAutoSave();
    }
  }

  public componentWillUnmount(): void {
    this.stopAutoSave();
  }

  private startAutoSave(): void {
    this.autoSaveTimer = setInterval(() => {
      this.handleAutoSave();
    }, 60000); // Auto-save every 60 seconds
  }

  private stopAutoSave(): void {
    if (this.autoSaveTimer) {
      clearInterval(this.autoSaveTimer);
      this.autoSaveTimer = null;
    }
  }

  private async loadPolicy(policyId: number): Promise<void> {
    try {
      this.setState({ loading: true, error: null });
      const policy = await this.policyService.getPolicyById(policyId);

      this.setState({
        policyNumber: policy.PolicyNumber,
        policyName: policy.PolicyName,
        policyCategory: policy.PolicyCategory,
        policySummary: policy.PolicySummary || '',
        policyContent: policy.PolicyContent || '',
        keyPoints: policy.KeyPoints || [],
        complianceRisk: policy.ComplianceRisk || 'Medium',
        readTimeframe: policy.ReadTimeframe || ReadTimeframe.Week1,
        readTimeframeDays: policy.ReadTimeframeDays || 7,
        requiresAcknowledgement: policy.RequiresAcknowledgement,
        requiresQuiz: policy.RequiresQuiz || false,
        effectiveDate: (typeof policy.EffectiveDate === 'string' ? policy.EffectiveDate : policy.EffectiveDate.toISOString()).split('T')[0],
        expiryDate: policy.ExpiryDate ? (typeof policy.ExpiryDate === 'string' ? policy.ExpiryDate : policy.ExpiryDate.toISOString()).split('T')[0] : '',
        loading: false
      });
    } catch (error) {
      console.error('Failed to load policy:', error);
      this.setState({
        error: 'Failed to load policy. Please try again.',
        loading: false
      });
    }
  }

  private handleAutoSave = async (): Promise<void> => {
    const { policyId, policyName, autoSaveEnabled } = this.state;

    if (!autoSaveEnabled || !policyId || !policyName) return;

    await this.handleSaveDraft(true);
  };

  private handleSaveDraft = async (isAutoSave: boolean = false): Promise<void> => {
    const {
      policyId,
      policyNumber,
      policyName,
      policyCategory,
      policySummary,
      policyContent,
      keyPoints,
      complianceRisk,
      readTimeframe,
      readTimeframeDays,
      requiresAcknowledgement,
      requiresQuiz,
      effectiveDate,
      expiryDate
    } = this.state;

    if (!policyName || !policyCategory) {
      if (!isAutoSave) {
        void this.dialogManager.showAlert('Policy name and category are required.', { variant: 'warning' });
      }
      return;
    }

    try {
      this.setState({ saving: true, error: null });

      const policyData: Partial<IPolicy> = {
        PolicyNumber: policyNumber || `POL-${Date.now()}`,
        PolicyName: policyName,
        PolicyCategory: policyCategory as PolicyCategory,
        PolicySummary: policySummary,
        PolicyContent: policyContent,
        KeyPoints: keyPoints,
        ComplianceRisk: complianceRisk as ComplianceRisk,
        ReadTimeframe: readTimeframe as ReadTimeframe,
        ReadTimeframeDays: readTimeframeDays,
        RequiresAcknowledgement: requiresAcknowledgement,
        RequiresQuiz: requiresQuiz,
        EffectiveDate: new Date(effectiveDate),
        ExpiryDate: expiryDate ? new Date(expiryDate) : undefined,
        Status: PolicyStatus.Draft
      };

      if (policyId) {
        await this.policyService.updatePolicy(policyId, policyData);
      } else {
        const newPolicy = await this.policyService.createPolicy(policyData);
        this.setState({ policyId: newPolicy.Id, policyNumber: newPolicy.PolicyNumber });
      }

      this.setState({
        saving: false,
        lastSaved: new Date()
      });

      if (!isAutoSave) {
        void this.dialogManager.showAlert('Draft saved successfully!', { variant: 'success' });
      }
    } catch (error) {
      console.error('Failed to save draft:', error);
      if (!isAutoSave) {
        this.setState({
          error: 'Failed to save draft. Please try again.',
          saving: false
        });
      }
    }
  };

  private handleSubmitForReview = async (): Promise<void> => {
    const { policyId } = this.state;

    if (!policyId) {
      await this.dialogManager.showAlert('Please save as draft first.', { variant: 'warning' });
      return;
    }

    // In a real implementation, you would select reviewers
    const reviewerIds = []; // Get from user selection

    try {
      this.setState({ saving: true });
      await this.policyService.submitForReview(policyId, reviewerIds);
      this.setState({ saving: false });
      await this.dialogManager.showAlert('Policy submitted for review successfully!', { variant: 'success' });
      // Optionally redirect to policy list
    } catch (error) {
      console.error('Failed to submit for review:', error);
      this.setState({
        error: 'Failed to submit for review. Please try again.',
        saving: false
      });
    }
  };

  private handleAddKeyPoint = (): void => {
    const { keyPoints, newKeyPoint } = this.state;
    if (newKeyPoint.trim()) {
      this.setState({
        keyPoints: [...keyPoints, newKeyPoint.trim()],
        newKeyPoint: ''
      });
    }
  };

  private handleRemoveKeyPoint = (index: number): void => {
    const { keyPoints } = this.state;
    this.setState({
      keyPoints: keyPoints.filter((_, i) => i !== index)
    });
  };

  private renderModuleNav(): JSX.Element {
    // Policy Builder tabs: Browse, Admin, Create Policy (this), Policy Packs, Quiz Builder
    const navItems = [
      { key: 'browse', text: 'Browse Policies', icon: 'Library', url: '/SitePages/PolicyHub.aspx', isActive: false },
      { key: 'admin', text: 'Policy Admin', icon: 'Settings', url: '/SitePages/PolicyAdmin.aspx', isActive: false },
      { key: 'author', text: 'Create Policy', icon: 'Edit', url: '', isActive: true },
      { key: 'packs', text: 'Policy Packs', icon: 'Package', url: '/SitePages/PolicyPacks.aspx', isActive: false },
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
    const { saving, lastSaved } = this.state;

    const items: ICommandBarItemProps[] = [
      {
        key: 'save',
        text: 'Save Draft',
        iconProps: { iconName: 'Save' },
        onClick: () => { this.handleSaveDraft(); },
        disabled: saving
      },
      {
        key: 'submit',
        text: 'Submit for Review',
        iconProps: { iconName: 'Send' },
        onClick: () => { this.handleSubmitForReview(); },
        disabled: saving
      }
    ];

    const farItems: ICommandBarItemProps[] = [];

    if (lastSaved) {
      farItems.push({
        key: 'lastSaved',
        text: `Last saved: ${lastSaved.toLocaleTimeString()}`,
        iconProps: { iconName: 'Recent' },
        disabled: true
      });
    }

    return <CommandBar items={items} farItems={farItems} />;
  }

  private renderBasicInfo(): JSX.Element {
    const {
      policyNumber,
      policyName,
      policyCategory,
      policySummary,
      complianceRisk,
      readTimeframe,
      readTimeframeDays
    } = this.state;

    const categoryOptions: IDropdownOption[] = Object.values(PolicyCategory).map(cat => ({
      key: cat,
      text: cat
    }));

    const riskOptions: IDropdownOption[] = Object.values(ComplianceRisk).map(risk => ({
      key: risk,
      text: risk
    }));

    const timeframeOptions: IDropdownOption[] = Object.values(ReadTimeframe).map(tf => ({
      key: tf,
      text: tf
    }));

    return (
      <div className={styles.section}>
        <Stack tokens={{ childrenGap: 16 }}>
          <TextField
            label="Policy Number"
            value={policyNumber}
            onChange={(e, value) => this.setState({ policyNumber: value || '' })}
            placeholder="Auto-generated if left blank"
          />

          <TextField
            label="Policy Name"
            required
            value={policyName}
            onChange={(e, value) => this.setState({ policyName: value || '' })}
            placeholder="Enter policy name"
          />

          <Dropdown
            label="Category"
            required
            selectedKey={policyCategory}
            options={categoryOptions}
            onChange={(e, option) => this.setState({ policyCategory: option?.key as string })}
          />

          <TextField
            label="Summary"
            multiline
            rows={3}
            value={policySummary}
            onChange={(e, value) => this.setState({ policySummary: value || '' })}
            placeholder="Brief summary of the policy (2-3 sentences)"
          />

          <Dropdown
            label="Compliance Risk"
            selectedKey={complianceRisk}
            options={riskOptions}
            onChange={(e, option) => this.setState({ complianceRisk: option?.key as string })}
          />

          <Dropdown
            label="Read Timeframe"
            selectedKey={readTimeframe}
            options={timeframeOptions}
            onChange={(e, option) => {
              const selected = option?.key as string;
              this.setState({
                readTimeframe: selected,
                readTimeframeDays: selected === ReadTimeframe.Custom ? readTimeframeDays : 7
              });
            }}
          />

          {readTimeframe === ReadTimeframe.Custom && (
            <TextField
              label="Custom Days"
              type="number"
              value={readTimeframeDays.toString()}
              onChange={(e, value) => this.setState({ readTimeframeDays: parseInt(value || '7', 10) })}
            />
          )}
        </Stack>
      </div>
    );
  }

  private renderContentEditor(): JSX.Element {
    const { policyContent } = this.state;

    return (
      <div className={styles.section}>
        <Text variant="xLarge" className={styles.sectionTitle}>
          Policy Content
        </Text>

        <div className={styles.richTextEditor}>
          <RichText
            value={policyContent}
            onChange={(text) => { this.setState({ policyContent: text }); return text; }}
            placeholder="Enter the detailed policy content..."
          />
        </div>
      </div>
    );
  }

  private renderKeyPoints(): JSX.Element {
    const { keyPoints, newKeyPoint } = this.state;

    return (
      <div className={styles.section}>
        <Text variant="xLarge" className={styles.sectionTitle}>
          Key Points
        </Text>

        <Stack tokens={{ childrenGap: 12 }}>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <TextField
              placeholder="Add a key point"
              value={newKeyPoint}
              onChange={(e, value) => this.setState({ newKeyPoint: value || '' })}
              styles={{ root: { flex: 1 } }}
            />
            <PrimaryButton
              text="Add"
              iconProps={{ iconName: 'Add' }}
              onClick={this.handleAddKeyPoint}
              disabled={!newKeyPoint.trim()}
            />
          </Stack>

          {keyPoints.length > 0 && (
            <div className={styles.keyPointsList}>
              {keyPoints.map((point: string, index: number) => (
                <div key={index} className={styles.keyPointItem}>
                  <Text>{point}</Text>
                  <IconButton
                    iconProps={{ iconName: 'Delete' }}
                    onClick={() => this.handleRemoveKeyPoint(index)}
                  />
                </div>
              ))}
            </div>
          )}
        </Stack>
      </div>
    );
  }

  private renderSettings(): JSX.Element {
    const {
      requiresAcknowledgement,
      requiresQuiz,
      effectiveDate,
      expiryDate
    } = this.state;

    return (
      <div className={styles.section}>
        <Text variant="xLarge" className={styles.sectionTitle}>
          Settings
        </Text>

        <Stack tokens={{ childrenGap: 16 }}>
          <Checkbox
            label="Requires Acknowledgement"
            checked={requiresAcknowledgement}
            onChange={(e, checked) => this.setState({ requiresAcknowledgement: checked || false })}
          />

          <Checkbox
            label="Requires Quiz"
            checked={requiresQuiz}
            onChange={(e, checked) => this.setState({ requiresQuiz: checked || false })}
          />

          <TextField
            label="Effective Date"
            type="date"
            value={effectiveDate}
            onChange={(e, value) => this.setState({ effectiveDate: value || '' })}
          />

          <TextField
            label="Expiry Date (Optional)"
            type="date"
            value={expiryDate}
            onChange={(e, value) => this.setState({ expiryDate: value || '' })}
          />
        </Stack>
      </div>
    );
  }

  public render(): React.ReactElement<IPolicyAuthorProps> {
    const { loading, error, saving } = this.state;

    return (
      <JmlAppLayout
        context={this.props.context}
        pageTitle="Policy Builder"
        pageDescription="Create, edit and manage policy documents and templates"
        pageIcon="Edit"
        breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Policy Builder' }]}
        activeNavKey="policies"
        showQuickLinks={true}
        showSearch={true}
        showNotifications={true}
        compactFooter={true}
      >
        <section className={styles.policyAuthor}>
          <Stack tokens={{ childrenGap: 24 }}>
            {/* Module nav removed - now in global header */}
            {this.renderCommandBar()}

            {loading && (
              <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
                <Spinner size={SpinnerSize.large} label="Loading policy..." />
              </Stack>
            )}

            {error && (
              <MessageBar messageBarType={MessageBarType.error} isMultiline>
                {error}
              </MessageBar>
            )}

            {saving && (
              <MessageBar messageBarType={MessageBarType.info}>
                Saving...
              </MessageBar>
            )}

            {!loading && (
              <div className={styles.editorContainer}>
                {this.renderBasicInfo()}
                {this.renderContentEditor()}
                {this.renderKeyPoints()}
                {this.renderSettings()}
              </div>
            )}
          </Stack>
          <this.dialogManager.DialogComponent />
        </section>
      </JmlAppLayout>
    );
  }
}
