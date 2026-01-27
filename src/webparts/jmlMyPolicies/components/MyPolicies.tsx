// @ts-nocheck
/* eslint-disable */
import * as React from 'react';
import { IMyPoliciesProps } from './IMyPoliciesProps';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  DefaultButton,
  PrimaryButton,
  ProgressIndicator,
  Icon,
  Panel,
  PanelType,
  Checkbox,
  TextField,
  IconButton
} from '@fluentui/react';
import { injectPortalStyles } from '../../../utils/injectPortalStyles';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { PolicyPackService } from '../../../services/PolicyPackService';
import { QuizTaker } from '../../../components/QuizTaker';
import { GamificationService } from '../../../services/GamificationService';
import {
  PolicyCertificateService,
  IAcknowledgementCertificateData
} from '../../../services/PolicyCertificateService';
import {
  IPersonalPolicyView,
  IPolicyAcknowledgement,
  IPolicyPackProgress
} from '../../../models/IPolicy';
import styles from './MyPolicies.module.scss';

// Wizard steps
type WizardStep = 'read' | 'quiz' | 'acknowledge' | 'success';

// Tab types
type TabType = 'all' | 'urgent' | 'dueSoon' | 'completed' | 'policyPacks';

// View types
type ViewType = 'list' | 'card';

export interface IMyPoliciesState {
  loading: boolean;
  error: string | null;
  personalView: IPersonalPolicyView | null;
  refreshing: boolean;

  // Enhanced UI state
  activeTab: TabType;
  viewType: ViewType;

  // Wizard state
  showWizard: boolean;
  selectedPolicy: IPolicyAcknowledgement | null;
  wizardStep: WizardStep;
  readProgress: number;
  readTimeSeconds: number;

  // Quiz state
  showQuiz: boolean;
  quizPassed: boolean;
  quizScore: number;

  // Acknowledgement state
  acknowledged: boolean;
  userRating: number;
  userComments: string;

  // Success state
  pointsEarned: number;
}

export default class MyPolicies extends React.Component<IMyPoliciesProps, IMyPoliciesState> {
  private policyPackService: PolicyPackService;
  private gamificationService: GamificationService;
  private policyCertificateService: PolicyCertificateService;
  private readTimer: any;
  private readStartTime: Date | null = null;

  constructor(props: IMyPoliciesProps) {
    super(props);
    this.state = {
      loading: true,
      error: null,
      personalView: null,
      refreshing: false,

      // Enhanced UI state
      activeTab: 'all',
      viewType: 'list',

      // Wizard state
      showWizard: false,
      selectedPolicy: null,
      wizardStep: 'read',
      readProgress: 0,
      readTimeSeconds: 0,

      // Quiz state
      showQuiz: false,
      quizPassed: false,
      quizScore: 0,

      // Acknowledgement state
      acknowledged: false,
      userRating: 0,
      userComments: '',

      // Success state
      pointsEarned: 0
    };
    this.policyPackService = new PolicyPackService(props.sp);
    this.gamificationService = new GamificationService(props.sp);
    this.policyCertificateService = new PolicyCertificateService();
  }

  public async componentDidMount(): Promise<void> {
    injectPortalStyles();
    await this.loadPersonalView();
  }

  public componentWillUnmount(): void {
    if (this.readTimer) {
      clearInterval(this.readTimer);
    }
  }

  private async loadPersonalView(): Promise<void> {
    try {
      this.setState({ loading: true, error: null });
      await this.policyPackService.initialize();
      const personalView = await this.policyPackService.getPersonalPolicyView();
      this.setState({ personalView, loading: false });
    } catch (error) {
      console.error('Failed to load personal policy view:', error);
      const errorMessage = error instanceof Error ? error.message : String(error);

      // Check for specific error types
      if (errorMessage.includes('404') || errorMessage.includes('does not exist') || errorMessage.includes('not been provisioned')) {
        this.setState({
          error: errorMessage.includes('not been provisioned')
            ? errorMessage
            : 'Policy lists have not been provisioned. Please contact your administrator to run the provisioning script.',
          loading: false
        });
      } else if (errorMessage.includes('Invalid user ID') || errorMessage.includes('user profile')) {
        this.setState({
          error: errorMessage,
          loading: false
        });
      } else if (errorMessage.includes('Unable to access')) {
        // More specific access error from the service
        this.setState({
          error: errorMessage,
          loading: false
        });
      } else {
        // Generic error with debugging info
        this.setState({
          error: `Failed to load your policies. ${errorMessage || 'Please try again later.'}`,
          loading: false
        });
      }
    }
  }

  private handleRefresh = async (): Promise<void> => {
    this.setState({ refreshing: true });
    await this.loadPersonalView();
    this.setState({ refreshing: false });
  };

  private handleTabChange = (tab: TabType): void => {
    this.setState({ activeTab: tab });
  };

  private handleViewToggle = (view: ViewType): void => {
    this.setState({ viewType: view });
  };

  private handlePolicyClick = (policy: IPolicyAcknowledgement): void => {
    this.setState({
      showWizard: true,
      selectedPolicy: policy,
      wizardStep: 'read',
      readProgress: 0,
      readTimeSeconds: 0,
      acknowledged: false,
      userRating: 0,
      userComments: '',
      showQuiz: false,
      quizPassed: false,
      quizScore: 0
    });

    // Start read timer
    this.readStartTime = new Date();
    this.readTimer = setInterval(() => {
      this.setState(prev => ({
        readTimeSeconds: prev.readTimeSeconds + 1,
        readProgress: Math.min(prev.readProgress + 2, 100) // Simulate reading progress
      }));
    }, 1000);
  };

  private handleCloseWizard = (): void => {
    if (this.readTimer) {
      clearInterval(this.readTimer);
    }
    this.setState({
      showWizard: false,
      selectedPolicy: null,
      wizardStep: 'read'
    });
  };

  private handleNextStep = (): void => {
    const { selectedPolicy, wizardStep } = this.state;

    if (wizardStep === 'read') {
      // Check if quiz is required
      if (selectedPolicy?.QuizRequired) {
        this.setState({ wizardStep: 'quiz', showQuiz: true });
      } else {
        this.setState({ wizardStep: 'acknowledge' });
      }
    } else if (wizardStep === 'quiz') {
      this.setState({ wizardStep: 'acknowledge', showQuiz: false });
    } else if (wizardStep === 'acknowledge') {
      this.handleSubmitAcknowledgement();
    }
  };

  private handleQuizComplete = async (result: any): Promise<void> => {
    this.setState({
      quizPassed: result.passed,
      quizScore: result.percentage,
      showQuiz: false
    });

    // Record quiz completion for gamification
    try {
      const userId = this.props.context.pageContext.legacyPageContext?.userId || 0;
      const { selectedPolicy } = this.state;
      if (userId > 0 && selectedPolicy) {
        await this.gamificationService.recordQuizCompletion(
          userId,
          selectedPolicy.QuizId || 1, // Use quiz ID from policy, or default
          result.passed,
          result.percentage
        );
      }
    } catch (gamificationError) {
      console.warn('Failed to record quiz completion points:', gamificationError);
    }

    if (result.passed) {
      this.setState({ wizardStep: 'acknowledge' });
    }
  };

  private handleSubmitAcknowledgement = async (): Promise<void> => {
    const { selectedPolicy, userRating, userComments, readTimeSeconds } = this.state;

    if (!selectedPolicy) return;

    try {
      // Submit acknowledgement
      await this.policyPackService.acknowledgePolicy(selectedPolicy.Id, {
        acknowledgedDate: new Date(),
        readDuration: readTimeSeconds,
        comments: userComments
      });

      // Award gamification points
      // Policy acknowledgement awards 15 points
      const POLICY_ACKNOWLEDGEMENT_POINTS = 15;
      try {
        // Get user ID from context (using userId from the pageContext)
        const userId = this.props.context.pageContext.legacyPageContext?.userId || 0;
        if (userId > 0) {
          await this.gamificationService.recordPolicyAcknowledgement(
            userId,
            selectedPolicy.PolicyId
          );
        }
      } catch (gamificationError) {
        // Log but don't fail the acknowledgement if gamification fails
        console.warn('Failed to record gamification points:', gamificationError);
      }

      // Update rating if provided
      if (userRating > 0) {
        await this.policyPackService.ratePolicyAcknowledgement(selectedPolicy.Id, userRating);
      }

      this.setState({
        wizardStep: 'success',
        pointsEarned: POLICY_ACKNOWLEDGEMENT_POINTS
      });

      // Refresh data
      setTimeout(() => {
        this.loadPersonalView();
      }, 2000);
    } catch (error) {
      console.error('Failed to submit acknowledgement:', error);
      this.setState({
        error: 'Failed to submit acknowledgement. Please try again.'
      });
    }
  };

  private handleDownloadCertificate = async (): Promise<void> => {
    const { selectedPolicy, quizScore, quizPassed, readTimeSeconds } = this.state;
    const { context } = this.props;

    if (!selectedPolicy) return;

    try {
      // Build certificate data
      const certificateData: IAcknowledgementCertificateData = {
        certificateId: this.policyCertificateService.generateCertificateId('POL'),
        employeeName: context.pageContext.user.displayName,
        employeeEmail: context.pageContext.user.email,
        employeeDepartment: '', // Would come from user profile in real implementation
        policyNumber: selectedPolicy.PolicyNumber || `POL-${selectedPolicy.PolicyId}`,
        policyName: selectedPolicy.PolicyName || 'Policy',
        policyCategory: selectedPolicy.PolicyCategory,
        policyVersion: selectedPolicy.PolicyVersionNumber?.toString() || '1.0',
        acknowledgedDate: new Date(),
        acknowledgementMethod: 'Digital',
        quizScore: selectedPolicy.QuizRequired ? quizScore : undefined,
        quizPassed: selectedPolicy.QuizRequired ? quizPassed : undefined
      };

      // Generate PDF certificate
      const result = await this.policyCertificateService.generateAcknowledgementCertificate(
        certificateData,
        {
          companyName: 'JML Employee Lifecycle Management',
          showSignature: true,
          signatureName: 'Policy Compliance Team',
          signatureTitle: 'Compliance Officer',
          primaryColor: '#0078d4',
          accentColor: '#107c10'
        }
      );

      if (result.success && result.blob) {
        // Download the PDF
        this.policyCertificateService.downloadCertificate(result);
      } else {
        console.error('Failed to generate certificate:', result.error);
        this.setState({ error: 'Failed to generate certificate. Please try again.' });
      }
    } catch (error) {
      console.error('Error generating certificate:', error);
      this.setState({ error: 'Failed to generate certificate. Please try again.' });
    }
  };

  private getFilteredPolicies(): IPolicyAcknowledgement[] {
    const { personalView, activeTab } = this.state;
    if (!personalView) return [];

    switch (activeTab) {
      case 'urgent':
        return personalView.urgentPolicies || [];
      case 'dueSoon':
        return personalView.dueSoon || [];
      case 'completed':
        return personalView.overduePolicies?.filter(p => p.Status === 'Acknowledged') || [];
      case 'all':
      default:
        return [
          ...(personalView.urgentPolicies || []),
          ...(personalView.dueSoon || []),
          ...(personalView.newPolicies || [])
        ];
    }
  }

  private getTabCounts(): Record<TabType, number> {
    const { personalView } = this.state;
    if (!personalView) {
      return { all: 0, urgent: 0, dueSoon: 0, completed: 0, policyPacks: 0 };
    }

    return {
      all: (personalView.urgentPolicies?.length || 0) +
           (personalView.dueSoon?.length || 0) +
           (personalView.newPolicies?.length || 0),
      urgent: personalView.urgentPolicies?.length || 0,
      dueSoon: personalView.dueSoon?.length || 0,
      completed: personalView.completed || 0,
      policyPacks: personalView.activePolicyPacks?.length || 0
    };
  }

  private renderTabs(): JSX.Element {
    const { activeTab } = this.state;
    const counts = this.getTabCounts();

    const tabs: { key: TabType; label: string; icon: string }[] = [
      { key: 'all', label: 'All Policies', icon: 'DocumentSet' },
      { key: 'urgent', label: 'Urgent', icon: 'Warning' },
      { key: 'dueSoon', label: 'Due Soon', icon: 'Clock' },
      { key: 'completed', label: 'Completed', icon: 'CheckMark' },
      { key: 'policyPacks', label: 'Policy Packs', icon: 'BulletedList' }
    ];

    return (
      <div className={styles.tabContainer}>
        <Stack horizontal tokens={{ childrenGap: 4 }}>
          {tabs.map(tab => (
            <button
              key={tab.key}
              className={`${styles.tabButton} ${activeTab === tab.key ? styles.activeTab : ''}`}
              onClick={() => this.handleTabChange(tab.key)}
            >
              <Icon iconName={tab.icon} style={{ marginRight: 8 }} />
              {tab.label}
              {counts[tab.key] > 0 && tab.key !== 'completed' && (
                <span className={styles.tabBadge}>{counts[tab.key]}</span>
              )}
            </button>
          ))}
        </Stack>
      </div>
    );
  }

  private renderViewToggle(): JSX.Element {
    const { viewType } = this.state;

    return (
      <div className={styles.viewToggle}>
        <button
          className={`${styles.toggleButton} ${viewType === 'list' ? styles.activeToggle : ''}`}
          onClick={() => this.handleViewToggle('list')}
        >
          <Icon iconName="List" />
        </button>
        <button
          className={`${styles.toggleButton} ${viewType === 'card' ? styles.activeToggle : ''}`}
          onClick={() => this.handleViewToggle('card')}
        >
          <Icon iconName="GridViewMedium" />
        </button>
      </div>
    );
  }

  private renderPolicyList(): JSX.Element {
    const policies = this.getFilteredPolicies();
    const { viewType, activeTab } = this.state;

    if (policies.length === 0) {
      return this.renderEmptyState();
    }

    if (viewType === 'card') {
      return (
        <div className={styles.policyCardGrid}>
          {policies.map(policy => this.renderPolicyCard(policy))}
        </div>
      );
    }

    return (
      <div className={styles.policyListView}>
        {policies.map(policy => this.renderPolicyListItem(policy))}
      </div>
    );
  }

  private renderPolicyListItem(policy: IPolicyAcknowledgement): JSX.Element {
    const isUrgent = policy.DueDate && new Date(policy.DueDate) <= new Date(Date.now() + 24 * 60 * 60 * 1000);
    const isDueSoon = policy.DueDate && new Date(policy.DueDate) <= new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);

    return (
      <div
        key={policy.Id}
        className={`${styles.policyListItem} ${isUrgent ? styles.urgentItem : isDueSoon ? styles.dueSoonItem : ''}`}
        onClick={() => this.handlePolicyClick(policy)}
      >
        <div className={styles.policyIcon}>
          <Icon iconName="DocumentSet" />
        </div>
        <div className={styles.policyInfo}>
          <div className={styles.policyName}>
            {policy.PolicyNumber} - {policy.PolicyName}
          </div>
          <div className={styles.policyMeta}>
            {policy.PolicyCategory} {policy.QuizRequired && '• Quiz Required'}
          </div>
        </div>
        <div className={styles.policyStatus}>
          {policy.DueDate && (
            <div className={styles.dueDate}>
              Due: {new Date(policy.DueDate).toLocaleDateString()}
            </div>
          )}
          <span className={`${styles.statusBadge} ${isUrgent ? styles.urgent : isDueSoon ? styles.dueSoon : ''}`}>
            {isUrgent ? 'Urgent' : isDueSoon ? 'Due Soon' : 'Pending'}
          </span>
        </div>
      </div>
    );
  }

  private renderPolicyCard(policy: IPolicyAcknowledgement): JSX.Element {
    const isUrgent = policy.DueDate && new Date(policy.DueDate) <= new Date(Date.now() + 24 * 60 * 60 * 1000);

    return (
      <div
        key={policy.Id}
        className={`${styles.policyCard} ${isUrgent ? styles.urgentCard : ''}`}
        onClick={() => this.handlePolicyClick(policy)}
      >
        <Stack tokens={{ childrenGap: 8 }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="large" className={styles.policyTitle}>
              {policy.PolicyNumber} - {policy.PolicyName}
            </Text>
            {isUrgent && <Icon iconName="Clock" className={styles.clockIcon} />}
          </Stack>
          <Text variant="small" className={styles.category}>
            {policy.PolicyCategory}
            {policy.QuizRequired && ' • Quiz Required'}
          </Text>
          {policy.DueDate && (
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
              <Icon iconName="Calendar" className={styles.icon} />
              <Text variant="small">
                Due: {new Date(policy.DueDate).toLocaleDateString()}
              </Text>
            </Stack>
          )}
          <Stack horizontal tokens={{ childrenGap: 8 }} className={styles.actions}>
            <PrimaryButton
              text="Read & Acknowledge"
              iconProps={{ iconName: 'ReadingMode' }}
              onClick={(e) => {
                e.stopPropagation();
                this.handlePolicyClick(policy);
              }}
            />
          </Stack>
        </Stack>
      </div>
    );
  }

  private renderProgressHeader(): JSX.Element | null {
    const { personalView } = this.state;
    const { showComplianceScore } = this.props;

    if (!showComplianceScore || !personalView) return null;

    const score = personalView.complianceScore;
    const total = (personalView.urgentPolicies?.length || 0) +
                  (personalView.dueSoon?.length || 0) +
                  (personalView.newPolicies?.length || 0) +
                  (personalView.completed || 0);
    const overdue = personalView.urgentPolicies?.length || 0;
    const pending = personalView.dueSoon?.length || 0;
    const completed = personalView.completed || 0;

    // SVG progress ring calculations
    const radius = 42;
    const circumference = 2 * Math.PI * radius;
    const strokeDashoffset = circumference - (score / 100) * circumference;

    // Status message based on score
    const statusMessage = score >= 90
      ? 'Excellent compliance!'
      : score >= 70
        ? 'Good progress - keep it up!'
        : 'Action needed on pending policies';

    return (
      <div className={styles.progressHeader}>
        <div className={styles.progressHeaderLeft}>
          {/* SVG Progress Ring */}
          <div className={styles.progressRingContainer}>
            <svg className={styles.progressRingSvg} viewBox="0 0 100 100">
              <circle
                className={styles.progressRingBg}
                cx="50"
                cy="50"
                r={radius}
              />
              <circle
                className={styles.progressRingFill}
                cx="50"
                cy="50"
                r={radius}
                style={{ strokeDashoffset }}
              />
            </svg>
            <div className={styles.progressRingText}>
              <div className={styles.progressRingPercent}>{score}%</div>
              <div className={styles.progressRingLabel}>Complete</div>
            </div>
          </div>

          {/* Header Info */}
          <div className={styles.progressHeaderInfo}>
            <h3>Policy Compliance</h3>
            <p>{statusMessage}</p>
          </div>
        </div>

        {/* Mini Stat Cards */}
        <div className={styles.miniStats}>
          <div className={styles.miniStat}>
            <div className={styles.miniStatNumber}>{total}</div>
            <div className={styles.miniStatLabel}>Total</div>
          </div>
          <div className={`${styles.miniStat} ${styles.warning}`}>
            <div className={styles.miniStatNumber}>{pending}</div>
            <div className={styles.miniStatLabel}>Pending</div>
          </div>
          <div className={`${styles.miniStat} ${styles.danger}`}>
            <div className={styles.miniStatNumber}>{overdue}</div>
            <div className={styles.miniStatLabel}>Overdue</div>
          </div>
          <div className={`${styles.miniStat} ${styles.success}`}>
            <div className={styles.miniStatNumber}>{completed}</div>
            <div className={styles.miniStatLabel}>Completed</div>
          </div>
        </div>
      </div>
    );
  }

  private renderPolicyPacks(): JSX.Element | null {
    const { personalView, activeTab } = this.state;
    const { showPolicyPacks } = this.props;

    if (activeTab !== 'policyPacks' || !showPolicyPacks || !personalView || personalView.activePolicyPacks.length === 0) {
      return null;
    }

    return (
      <div className={styles.policyPacksSection}>
        <Stack tokens={{ childrenGap: 12 }}>
          <Text variant="xLarge" className={styles.sectionTitle}>
            <Icon iconName="BulletedList" className={styles.packIcon} />
            Active Policy Packs
          </Text>
          {personalView.activePolicyPacks.map((pack: IPolicyPackProgress) => (
            <div key={pack.assignmentId} className={styles.packCard}>
              <Stack tokens={{ childrenGap: 12 }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                  <Text variant="large" className={styles.packTitle}>
                    {pack.packName}
                  </Text>
                  <Text variant="medium" style={{ fontWeight: 600 }}>
                    {pack.progressPercentage}%
                  </Text>
                </Stack>
                <ProgressIndicator
                  percentComplete={pack.progressPercentage / 100}
                  barHeight={8}
                />
                <Stack horizontal tokens={{ childrenGap: 16 }}>
                  <Stack tokens={{ childrenGap: 4 }}>
                    <Text variant="small" className={styles.packStat}>
                      {pack.acknowledgements?.filter(a => a.Status === 'Acknowledged').length || 0} / {pack.acknowledgements?.length || 0}
                    </Text>
                    <Text variant="xSmall" className={styles.subText}>Completed</Text>
                  </Stack>
                  <Stack tokens={{ childrenGap: 4 }}>
                    <Text variant="small" className={styles.packStat}>
                      {pack.acknowledgements?.filter(a => a.Status === 'Overdue').length || 0}
                    </Text>
                    <Text variant="xSmall" className={styles.subText}>Overdue</Text>
                  </Stack>
                  {pack.estimatedCompletionDate && (
                    <Stack tokens={{ childrenGap: 4 }}>
                      <Text variant="small" className={styles.packStat}>
                        {new Date(pack.estimatedCompletionDate).toLocaleDateString()}
                      </Text>
                      <Text variant="xSmall" className={styles.subText}>Est. Completion</Text>
                    </Stack>
                  )}
                </Stack>
                {!pack.isOnTrack && (
                  <MessageBar messageBarType={MessageBarType.warning}>
                    This policy pack is behind schedule. Please review urgent items.
                  </MessageBar>
                )}
              </Stack>
            </div>
          ))}
        </Stack>
      </div>
    );
  }

  private renderEmptyState(): JSX.Element {
    return (
      <div className={styles.emptyState}>
        <Stack tokens={{ childrenGap: 16 }} horizontalAlign="center">
          <Icon iconName="CompletedSolid" className={styles.emptyIcon} />
          <Text variant="xLarge">All Caught Up!</Text>
          <Text variant="medium" className={styles.subText}>
            You have no pending policy acknowledgements at this time.
          </Text>
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // WIZARD RENDERING
  // ============================================================================

  private renderWizardStepper(): JSX.Element {
    const { wizardStep, selectedPolicy } = this.state;
    const hasQuiz = selectedPolicy?.QuizRequired;

    const steps = [
      { key: 'read', label: 'Read Policy', icon: 'ReadingMode' },
      ...(hasQuiz ? [{ key: 'quiz', label: 'Complete Quiz', icon: 'TestPlan' }] : []),
      { key: 'acknowledge', label: 'Acknowledge', icon: 'CheckMark' },
      { key: 'success', label: 'Complete', icon: 'Trophy' }
    ];

    const currentIndex = steps.findIndex(s => s.key === wizardStep);

    return (
      <div className={styles.wizardStepper}>
        {steps.map((step, index) => (
          <div
            key={step.key}
            className={`${styles.stepItem} ${index === currentIndex ? styles.activeStep : ''} ${index < currentIndex ? styles.completedStep : ''}`}
          >
            <div className={styles.stepCircle}>
              {index < currentIndex ? (
                <Icon iconName="CheckMark" />
              ) : (
                <Icon iconName={step.icon} />
              )}
            </div>
            <div className={styles.stepLabel}>{step.label}</div>
          </div>
        ))}
      </div>
    );
  }

  private renderWizardContent(): JSX.Element {
    const { wizardStep, selectedPolicy, showQuiz } = this.state;

    if (!selectedPolicy) return <></>;

    switch (wizardStep) {
      case 'read':
        return this.renderPolicyReader();

      case 'quiz':
        if (showQuiz) {
          return (
            <QuizTaker
              sp={this.props.sp}
              quizId={1} // TODO: Get actual quiz ID from policy
              policyId={selectedPolicy.PolicyId}
              userId={this.props.context.pageContext.user}
              onComplete={this.handleQuizComplete}
              onCancel={() => this.setState({ showQuiz: false })}
            />
          );
        }
        return this.renderQuizResults();

      case 'acknowledge':
        return this.renderAcknowledgementForm();

      case 'success':
        return this.renderSuccessScreen();

      default:
        return <></>;
    }
  }

  private renderPolicyReader(): JSX.Element {
    const { selectedPolicy, readProgress, readTimeSeconds } = this.state;

    return (
      <div className={styles.policyReader}>
        <div className={styles.readerToolbar}>
          <Stack horizontal tokens={{ childrenGap: 16 }} verticalAlign="center">
            <Text variant="medium" style={{ fontWeight: 600 }}>
              {selectedPolicy?.PolicyNumber} - {selectedPolicy?.PolicyName}
            </Text>
          </Stack>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <IconButton iconProps={{ iconName: 'ZoomIn' }} title="Zoom In" />
            <IconButton iconProps={{ iconName: 'ZoomOut' }} title="Zoom Out" />
            <IconButton iconProps={{ iconName: 'Download' }} title="Download" />
          </Stack>
        </div>

        <div className={styles.readerContent}>
          {/* Policy content - this would load from DocumentURL or HTMLContent */}
          <h1>{selectedPolicy?.PolicyName}</h1>
          <p><strong>Policy Number:</strong> {selectedPolicy?.PolicyNumber}</p>
          <p><strong>Category:</strong> {selectedPolicy?.PolicyCategory}</p>
          <p><strong>Effective Date:</strong> {selectedPolicy?.AssignedDate ? new Date(selectedPolicy.AssignedDate).toLocaleDateString() : 'N/A'}</p>

          <h2>1. Purpose</h2>
          <p>
            This policy document outlines the requirements and guidelines that all employees must follow.
            Please read this document carefully before acknowledging.
          </p>

          <h2>2. Scope</h2>
          <p>
            This policy applies to all employees, contractors, and third-party personnel who have access
            to company systems and data.
          </p>

          <h2>3. Key Requirements</h2>
          <ul>
            <li>All employees must complete this acknowledgement within the specified timeframe</li>
            <li>Compliance with this policy is mandatory</li>
            <li>Violations may result in disciplinary action</li>
          </ul>

          <h2>4. Your Responsibilities</h2>
          <p>
            By acknowledging this policy, you confirm that you have read, understood, and agree to
            comply with all requirements outlined in this document.
          </p>
        </div>

        <div className={styles.readerProgress}>
          <div className={styles.progressLabel}>
            Reading Progress: {Math.floor(readTimeSeconds / 60)}m {readTimeSeconds % 60}s
          </div>
          <ProgressIndicator percentComplete={readProgress / 100} barHeight={4} />
        </div>
      </div>
    );
  }

  private renderQuizResults(): JSX.Element {
    const { quizPassed, quizScore } = this.state;

    return (
      <Stack tokens={{ childrenGap: 16 }} horizontalAlign="center">
        <Icon
          iconName={quizPassed ? "CompletedSolid" : "StatusCircleErrorX"}
          style={{ fontSize: 64, color: quizPassed ? '#107C10' : '#D13438' }}
        />
        <Text variant="xLarge">
          {quizPassed ? 'Quiz Passed!' : 'Quiz Not Passed'}
        </Text>
        <Text variant="large">
          Your Score: {quizScore}%
        </Text>
        {!quizPassed && (
          <DefaultButton
            text="Retake Quiz"
            onClick={() => this.setState({ showQuiz: true })}
          />
        )}
      </Stack>
    );
  }

  private renderAcknowledgementForm(): JSX.Element {
    const { acknowledged, userRating, userComments, selectedPolicy } = this.state;

    return (
      <div className={styles.acknowledgementForm}>
        <Stack tokens={{ childrenGap: 20 }}>
          <Text variant="xLarge" style={{ fontWeight: 600 }}>
            Acknowledge Policy
          </Text>

          <div className={`${styles.acknowledgementCheckbox} ${acknowledged ? styles.checked : ''}`}>
            <Checkbox
              label={`I, ${this.props.context.pageContext.user.displayName}, confirm that I have read, understood, and agree to comply with the ${selectedPolicy?.PolicyName} policy.`}
              checked={acknowledged}
              onChange={(_, checked) => this.setState({ acknowledged: !!checked })}
              styles={{ root: { fontWeight: 500 } }}
            />
          </div>

          <div className={styles.ratingSection}>
            <div className={styles.ratingLabel}>How helpful was this policy? (Optional)</div>
            <div className={styles.starRating}>
              {[1, 2, 3, 4, 5].map(star => (
                <span
                  key={star}
                  className={`${styles.star} ${star <= userRating ? styles.filled : ''}`}
                  onClick={() => this.setState({ userRating: star })}
                >
                  ★
                </span>
              ))}
            </div>
          </div>

          <TextField
            label="Comments (Optional)"
            multiline
            rows={3}
            value={userComments}
            onChange={(_, value) => this.setState({ userComments: value || '' })}
            placeholder="Share any feedback about this policy..."
          />

          <Stack horizontal tokens={{ childrenGap: 16 }} horizontalAlign="end">
            <DefaultButton
              text="Back"
              onClick={() => this.setState({ wizardStep: 'read' })}
            />
            <PrimaryButton
              text="Submit Acknowledgement"
              iconProps={{ iconName: 'CheckMark' }}
              disabled={!acknowledged}
              onClick={this.handleNextStep}
            />
          </Stack>
        </Stack>
      </div>
    );
  }

  private renderSuccessScreen(): JSX.Element {
    const { selectedPolicy, pointsEarned } = this.state;

    return (
      <div className={styles.successScreen}>
        <div className={styles.successIcon}>
          <Icon iconName="CheckMark" />
        </div>
        <div className={styles.successTitle}>Policy Acknowledged!</div>
        <div className={styles.successMessage}>
          Thank you for acknowledging the {selectedPolicy?.PolicyName} policy.
        </div>

        {pointsEarned > 0 && (
          <div className={styles.pointsEarned}>
            🏆 +{pointsEarned} points earned!
          </div>
        )}

        <div className={styles.certificate}>
          <div className={styles.certificateTitle}>Certificate of Acknowledgement</div>
          <div className={styles.certificateDate}>
            {this.props.context.pageContext.user.displayName}<br />
            Acknowledged: {new Date().toLocaleDateString()}
          </div>
        </div>

        <Stack horizontal tokens={{ childrenGap: 16 }} horizontalAlign="center" style={{ marginTop: 24 }}>
          <DefaultButton
            text="Download Certificate"
            iconProps={{ iconName: 'PDF' }}
            onClick={this.handleDownloadCertificate}
          />
          <PrimaryButton
            text="Done"
            onClick={this.handleCloseWizard}
          />
        </Stack>
      </div>
    );
  }

  private renderWizardPanel(): JSX.Element {
    const { showWizard, selectedPolicy, wizardStep } = this.state;

    return (
      <Panel
        isOpen={showWizard}
        onDismiss={this.handleCloseWizard}
        type={PanelType.large}
        headerText=""
        isLightDismiss={false}
        closeButtonAriaLabel="Close"
      >
        <div className={styles.wizardPanel}>
          <div className={styles.wizardHeader}>
            <div className={styles.wizardTitle}>
              {selectedPolicy?.PolicyNumber} - {selectedPolicy?.PolicyName}
            </div>
            <div className={styles.wizardSubtitle}>
              {selectedPolicy?.PolicyCategory}
            </div>
          </div>

          {this.renderWizardStepper()}
          {this.renderWizardContent()}

          {wizardStep === 'read' && (
            <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 16 }} style={{ marginTop: 24 }}>
              <DefaultButton text="Close" onClick={this.handleCloseWizard} />
              <PrimaryButton
                text={this.state.selectedPolicy?.QuizRequired ? "Continue to Quiz" : "Continue to Acknowledge"}
                iconProps={{ iconName: 'Forward' }}
                onClick={this.handleNextStep}
                disabled={this.state.readProgress < 50}
              />
            </Stack>
          )}
        </div>
      </Panel>
    );
  }

  public render(): React.ReactElement<IMyPoliciesProps> {
    const { loading, error, personalView, refreshing, activeTab } = this.state;

    return (
      <JmlAppLayout
        context={this.props.context}
        pageTitle="My Policies"
        pageDescription="View policies assigned to you and track acknowledgements"
        pageIcon="DocumentSet"
        breadcrumbs={[{ text: 'JML Portal', url: '/sites/JML' }, { text: 'My Policies' }]}
        activeNavKey="policies"
        showQuickLinks={true}
        showSearch={true}
        showNotifications={true}
        compactFooter={true}
      >
        <section className={styles.myPolicies}>
          <Stack tokens={{ childrenGap: 24 }}>
            {/* Header with refresh and view toggle */}
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <DefaultButton
                text="Refresh"
                iconProps={{ iconName: 'Refresh' }}
                onClick={this.handleRefresh}
                disabled={loading || refreshing}
              />
              {this.renderViewToggle()}
            </Stack>

            {loading && (
              <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
                <Spinner size={SpinnerSize.large} label="Loading your policies..." />
              </Stack>
            )}

            {error && (
              <MessageBar messageBarType={MessageBarType.error} isMultiline>
                {error}
              </MessageBar>
            )}

            {!loading && !error && personalView && (
              <Stack tokens={{ childrenGap: 24 }}>
                {/* Progress Header with SVG Ring and Mini Stats */}
                {this.renderProgressHeader()}

                {/* Tab Navigation */}
                {this.renderTabs()}

                {/* Policy List or Policy Packs */}
                {activeTab === 'policyPacks' ? (
                  this.renderPolicyPacks()
                ) : (
                  this.renderPolicyList()
                )}
              </Stack>
            )}
          </Stack>

          {/* Wizard Panel */}
          {this.renderWizardPanel()}
        </section>
      </JmlAppLayout>
    );
  }
}
