// @ts-nocheck
import { Icon } from '@fluentui/react/lib/Icon';
/* eslint-disable */
import * as React from 'react';
import { IPolicyDetailsProps } from './IPolicyDetailsProps';
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
  Label,
  TextField,
  Checkbox,
  Rating,
  RatingSize,
  Separator,
  Dialog,
  DialogType,
  DialogFooter,
  Panel,
  PanelType,
  ProgressIndicator
} from '@fluentui/react';
import { injectPortalStyles } from '../../../utils/injectPortalStyles';
import { signalAppReady } from '../../../utils/SharePointOverrides';
import { sanitizeHtml, escapeHtml } from '../../../utils/sanitizeHtml';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { ErrorBoundary } from '../../../components/ErrorBoundary/ErrorBoundary';
import { PageSubheader } from '../../../components/PageSubheader';
import { PolicyService } from '../../../services/PolicyService';
import { PolicySocialService } from '../../../services/PolicySocialService';
import { createDialogManager } from '../../../hooks/useDialog';
import {
  IPolicy,
  IPolicyAcknowledgement,
  IPolicyRating,
  IPolicyComment,
  IPolicyAcknowledgeRequest,
  IPolicyVersion
} from '../../../models/IPolicy';
import { PolicyDocumentComparisonService } from '../../../services/PolicyDocumentComparisonService';
import { StyledPanel } from '../../../components/StyledPanel';
import styles from './PolicyDetails.module.scss';
import { PM_LISTS } from '../../../constants/SharePointListNames';
import { QuizService, IQuizResult } from '../../../services/QuizService';
import { QuizTaker } from '../../../components/QuizTaker/QuizTaker';
import { DwxLinkedRecordService, DwxLinkedRecordsPanel } from '@dwx/core';
import { RecentlyViewedService } from '../../../services/RecentlyViewedService';

// Read flow steps
export type ReadFlowStep = 'reading' | 'quiz' | 'acknowledge' | 'complete';

// Read receipt for audit trail
export interface IReadReceipt {
  Id?: number;
  UserId: number;
  UserEmail: string;
  UserDisplayName: string;
  PolicyId: number;
  PolicyNumber: string;
  PolicyName: string;
  PolicyVersion: string;
  ReadStartTime: Date;
  ReadEndTime: Date;
  ReadDurationSeconds: number;
  QuizRequired: boolean;
  QuizCompleted: boolean;
  QuizScore?: number;
  QuizPassPercentage?: number;
  QuizPassedDate?: Date;
  AcknowledgedDate: Date;
  AcknowledgedTime: string;
  IPAddress: string;
  UserAgent: string;
  DeviceType: string;
  BrowserName: string;
  DigitalSignature: string;
  LegalConfirmationText: string;
  Notes?: string;
  ReceiptNumber: string;
}

// Quiz question model
interface IQuizQuestion {
  id: number;
  question: string;
  options: string[];
  correctIndex: number;
}

export interface IPolicyDetailsState {
  loading: boolean;
  error: string | null;
  policy: IPolicy | null;
  acknowledgement: IPolicyAcknowledgement | null;
  ratings: IPolicyRating[];
  comments: IPolicyComment[];
  policyId: number | null;
  readStartTime: Date | null;
  readDuration: number;
  showAcknowledgeDialog: boolean;
  acknowledgeConfirmation: boolean;
  acknowledgeNotes: string;
  submittingAcknowledgement: boolean;
  showCommentDialog: boolean;
  newComment: string;
  submittingComment: boolean;
  userRating: number;
  reviewTitle: string;
  reviewText: string;
  submittingRating: boolean;
  isFollowing: boolean;
  // Enhanced read flow
  currentFlowStep: ReadFlowStep;
  hasReadPolicy: boolean;
  scrollProgress: number;
  quizRequired: boolean;
  quizCompleted: boolean;
  quizScore: number;
  quizPassed: boolean;
  showAcknowledgePanel: boolean;
  showCongratulationsPanel: boolean;
  readReceipt: IReadReceipt | null;
  showReadReceiptPanel: boolean;
  emailingReceipt: boolean;
  generatingPdf: boolean;
  legalAgreement1: boolean;
  legalAgreement2: boolean;
  legalAgreement3: boolean;
  digitalSignature: string;
  // Horizontal quiz state
  currentQuizQuestion: number;
  quizAnswers: number[];
  quizConfirmed: boolean[];  // which questions have been confirmed (locked in — after 2 attempts or correct)
  quizAttempts: number[];    // how many attempts per question (max 2 before answer revealed)
  quizSubmitted: boolean;
  // Browse mode — read-only viewing from Policy Hub (no wizard/acknowledge flow)
  browseMode: boolean;
  // Live quiz integration
  liveQuizId: number | null;
  currentUserId: number;
  // Fullscreen document viewer
  isFullscreen: boolean;
  // Version history
  showVersionHistoryPanel: boolean;
  versionHistoryLoading: boolean;
  policyVersions: IPolicyVersion[];
  showVersionComparisonPanel: boolean;
  versionComparisonHtml: string;
  versionComparisonLoading: boolean;
}

// Mock quiz questions (will be replaced by QuizTaker integration)
const MOCK_QUIZ_QUESTIONS: IQuizQuestion[] = [
  {
    id: 1,
    question: 'What is the minimum password length required by the policy?',
    options: ['8 characters', '10 characters', '12 characters', '16 characters'],
    correctIndex: 2
  },
  {
    id: 2,
    question: 'How often must passwords be changed?',
    options: ['30 days', '60 days', '90 days', 'Never'],
    correctIndex: 2
  },
  {
    id: 3,
    question: 'Who should you report security incidents to?',
    options: ['Your manager', 'IT Help Desk', 'Security Team', 'All of the above'],
    correctIndex: 2
  },
  {
    id: 4,
    question: 'Is it acceptable to share your login credentials with a colleague?',
    options: ['Yes, if they need access', 'Only in emergencies', 'Never', 'Only with manager approval'],
    correctIndex: 2
  },
  {
    id: 5,
    question: 'What should you do if you suspect a phishing email?',
    options: ['Delete it', 'Forward to Security Team', 'Click to verify', 'Ignore it'],
    correctIndex: 1
  }
];

export default class PolicyDetails extends React.Component<IPolicyDetailsProps, IPolicyDetailsState> {
  private _isMounted = false;
  private policyService: PolicyService;
  private socialService: PolicySocialService;
  private linkedRecordService: DwxLinkedRecordService | null = null;
  private comparisonService: PolicyDocumentComparisonService;
  private readTimer: NodeJS.Timeout | null = null;
  private dialogManager = createDialogManager();
  private documentViewerRef: React.RefObject<HTMLDivElement>;
  private viewerWrapperRef: React.RefObject<HTMLDivElement>;

  constructor(props: IPolicyDetailsProps) {
    super(props);
    this.documentViewerRef = React.createRef();
    this.viewerWrapperRef = React.createRef();
    this.state = {
      loading: true,
      error: null,
      policy: null,
      acknowledgement: null,
      ratings: [],
      comments: [],
      policyId: this.getPolicyIdFromUrl(),
      readStartTime: null,
      readDuration: 0,
      showAcknowledgeDialog: false,
      acknowledgeConfirmation: false,
      acknowledgeNotes: '',
      submittingAcknowledgement: false,
      showCommentDialog: false,
      newComment: '',
      submittingComment: false,
      userRating: 0,
      reviewTitle: '',
      reviewText: '',
      submittingRating: false,
      isFollowing: false,
      // Enhanced read flow
      currentFlowStep: 'reading',
      hasReadPolicy: false,
      scrollProgress: 0,
      quizRequired: false,
      quizCompleted: false,
      quizScore: 0,
      quizPassed: false,
      showAcknowledgePanel: false,
      showCongratulationsPanel: false,
      readReceipt: null,
      showReadReceiptPanel: false,
      emailingReceipt: false,
      generatingPdf: false,
      legalAgreement1: false,
      legalAgreement2: false,
      legalAgreement3: false,
      digitalSignature: '',
      // Horizontal quiz
      currentQuizQuestion: 0,
      quizAnswers: new Array(MOCK_QUIZ_QUESTIONS.length).fill(-1),
      quizConfirmed: new Array(MOCK_QUIZ_QUESTIONS.length).fill(false),
      quizAttempts: new Array(MOCK_QUIZ_QUESTIONS.length).fill(0),
      quizSubmitted: false,
      // Browse mode detection — from Policy Hub browsing
      browseMode: this.getBrowseModeFromUrl(),
      // Review mode — from reviewer email link
      reviewMode: new URLSearchParams(window.location.search).get('mode') === 'review',
      approvalMode: new URLSearchParams(window.location.search).get('mode') === 'approve',
      reviewDecision: '' as string,
      reviewComments: '' as string,
      reviewChecklist: [false, false, false, false, false, false],
      reviewSubmitting: false,
      reviewerItems: [] as any[],
      // Live quiz integration
      liveQuizId: null,
      currentUserId: 0,
      isFullscreen: false,
      // Version history
      showVersionHistoryPanel: false,
      versionHistoryLoading: false,
      policyVersions: [],
      showVersionComparisonPanel: false,
      versionComparisonHtml: '',
      versionComparisonLoading: false
    };
    this.policyService = new PolicyService(props.sp);
    this.socialService = new PolicySocialService(props.sp);
    this.comparisonService = new PolicyDocumentComparisonService(props.sp, props.context.pageContext.web.absoluteUrl);
    if (props.dwxHub) {
      this.linkedRecordService = new DwxLinkedRecordService(props.dwxHub);
    }
  }

  public async componentDidMount(): Promise<void> {
    this._isMounted = true;
    injectPortalStyles();
    await this.loadPolicyDetails();
    this.startReadTracking();
    document.addEventListener('fullscreenchange', this.handleFullscreenChange);
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
    this.stopReadTracking();
    document.removeEventListener('fullscreenchange', this.handleFullscreenChange);
  }

  private handleFullscreenChange = (): void => {
    this.setState({ isFullscreen: !!document.fullscreenElement });
  };

  private toggleFullscreen = async (): Promise<void> => {
    const wrapper = this.viewerWrapperRef.current;
    if (!wrapper) return;

    try {
      if (!document.fullscreenElement) {
        await wrapper.requestFullscreen();
      } else {
        await document.exitFullscreen();
      }
    } catch (err) {
      console.warn('Fullscreen not supported:', err);
    }
  };

  private getPolicyIdFromUrl(): number | null {
    const urlParams = new URLSearchParams(window.location.search);
    const policyId = urlParams.get('policyId');
    return policyId ? parseInt(policyId, 10) : null;
  }

  private getBrowseModeFromUrl(): boolean {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get('mode') === 'browse';
  }

  private async loadPolicyDetails(): Promise<void> {
    const { policyId, browseMode } = this.state;
    if (!policyId) {
      this.setState({ error: 'No policy ID provided', loading: false });
      return;
    }

    try {
      this.setState({ loading: true, error: null });

      // Browse mode (from Policy Hub): ultra-lightweight direct SP query
      // Skip service initialization, audit, social, dashboard, quiz lookups
      if (browseMode) {
        try {
          const item = await this.props.sp.web.lists
            .getByTitle('PM_Policies')
            .items.getById(policyId)
            .select('*')();
          if (this._isMounted) {
            this.setState({ policy: item as any, loading: false });
          }
          RecentlyViewedService.trackView(item.Id, item.PolicyName || item.Title, item.PolicyCategory || '');
        } catch (browseErr) {
          if (this._isMounted) {
            this.setState({ error: `Failed to load policy: ${(browseErr as Error).message}`, loading: false });
          }
        }
        return;
      }

      // Full load for assigned read mode (from My Policies)
      await this.policyService.initialize();

      const policy = await this.policyService.getPolicyById(policyId);

      // Full load for assigned read mode (from My Policies)
      await this.socialService.initialize();

      const currentUser = await this.props.sp.web.currentUser();
      const dashboard = await this.policyService.getUserDashboard(currentUser.Id);
      const acknowledgement = dashboard.pendingAcknowledgements.find(
        (ack: IPolicyAcknowledgement) => ack.PolicyId === policyId
      ) || dashboard.completedAcknowledgements.find(
        (ack: IPolicyAcknowledgement) => ack.PolicyId === policyId
      );

      const ratings = await this.socialService.getPolicyRatings(policyId);
      const comments = await this.socialService.getPolicyComments(policyId);
      const isFollowing = await this.socialService.isFollowingPolicy(policyId);

      // Look up live quiz for this policy (if any)
      let liveQuizId: number | null = null;
      try {
        const quizService = new QuizService(this.props.sp);
        const quizzes = await quizService.getQuizzesByPolicy(policyId);
        const activeQuiz = quizzes.find(q => q.IsActive && q.Status === 'Published');
        if (activeQuiz) {
          liveQuizId = activeQuiz.Id;
        }
      } catch (quizErr) {
        console.warn('Could not look up quiz for policy:', quizErr);
      }

      if (this._isMounted) { this.setState({
        policy,
        acknowledgement,
        ratings,
        comments,
        isFollowing,
        liveQuizId,
        currentUserId: currentUser.Id,
        loading: false
      }); }

      // Track this policy view in Recently Viewed (localStorage)
      RecentlyViewedService.trackView(
        policy.Id,
        policy.PolicyName || policy.Title,
        policy.PolicyCategory || ''
      );

      if (acknowledgement && acknowledgement.AckStatus !== 'Acknowledged') {
        await this.policyService.trackPolicyOpen(acknowledgement.Id);
      }
    } catch (error) {
      console.error('Failed to load policy details, falling back to mock data:', error);
      // Fall back to mock data so the wizard can be tested without live SharePoint lists
      this.loadMockPolicyDetails();
    }
  }

  private loadMockPolicyDetails(): void {
    const { policyId, browseMode } = this.state;
    const mockPolicy: IPolicy = {
      Id: policyId || 1,
      Title: 'Information Security Policy',
      PolicyNumber: 'POL-2024-001',
      PolicyName: 'Information Security Policy',
      PolicyDescription: 'This policy establishes the information security requirements for all employees to protect company assets, data, and systems from unauthorized access, disclosure, and modification.',
      PolicyCategory: 'IT Security',
      PolicyStatus: 'Published',
      PolicyVersion: '2.1',
      EffectiveDate: new Date('2024-01-15'),
      ReviewDate: new Date('2025-01-15'),
      ExpiryDate: new Date('2025-12-31'),
      PolicyOwner: 'Sarah Johnson',
      PolicyOwnerId: 1,
      Department: 'Information Technology',
      RequiresQuiz: true,
      QuizPassingScore: 80,
      AllowRetake: true,
      MaxRetakeAttempts: 3,
      DocumentURL: '/sites/PolicyManager/PolicyDocuments/Information-Security-Policy.pdf',
      Created: new Date('2024-01-10'),
      Modified: new Date('2024-06-15'),
      AuthorId: 1
    } as IPolicy;

    const mockAcknowledgement: IPolicyAcknowledgement = {
      Id: 100,
      Title: 'Ack-001',
      PolicyId: policyId || 1,
      UserId: 1,
      AckStatus: browseMode ? 'Acknowledged' : 'Pending',
      AssignedDate: new Date('2024-06-01'),
      DueDate: new Date('2024-07-01'),
      QuizRequired: true,
      Created: new Date('2024-06-01'),
      Modified: new Date('2024-06-01'),
      AuthorId: 1
    } as IPolicyAcknowledgement;

    this.setState({
      policy: mockPolicy,
      acknowledgement: mockAcknowledgement,
      ratings: [],
      comments: [],
      isFollowing: false,
      loading: false,
      quizRequired: true
    });

    // Track mock policy view in Recently Viewed (localStorage)
    RecentlyViewedService.trackView(
      mockPolicy.Id,
      mockPolicy.PolicyName || mockPolicy.Title,
      mockPolicy.PolicyCategory || ''
    );
  }

  private startReadTracking(): void {
    this.setState({ readStartTime: new Date() });
    this.readTimer = setInterval(() => {
      this.setState((prevState) => ({
        readDuration: prevState.readDuration + 1
      }));
    }, 1000);
  }

  private stopReadTracking(): void {
    if (this.readTimer) {
      clearInterval(this.readTimer);
      this.readTimer = null;
    }
  }

  private formatDuration(seconds: number): string {
    const mins = Math.floor(seconds / 60);
    const secs = seconds % 60;
    return `${mins}m ${secs.toString().padStart(2, '0')}s`;
  }

  // ============================================
  // SOCIAL ACTIONS
  // ============================================

  private handleRate = async (rating: number): Promise<void> => {
    this.setState({ userRating: rating });
  };

  private handleSubmitRating = async (): Promise<void> => {
    const { policy, userRating, reviewTitle, reviewText } = this.state;
    if (!policy) return;

    try {
      this.setState({ submittingRating: true });
      await this.socialService.ratePolicy({
        policyId: policy.Id,
        rating: userRating,
        reviewTitle,
        reviewText
      });
      const ratings = await this.socialService.getPolicyRatings(policy.Id);
      this.setState({ ratings, submittingRating: false, reviewTitle: '', reviewText: '' });
      await this.dialogManager.showAlert('Thank you for your rating!', { variant: 'success' });
    } catch (error) {
      console.error('Failed to submit rating:', error);
      this.setState({ submittingRating: false });
    }
  };

  private handleComment = (): void => {
    this.setState({ showCommentDialog: true });
  };

  private handleSubmitComment = async (): Promise<void> => {
    const { policy, newComment } = this.state;
    if (!policy || !newComment.trim()) return;

    try {
      this.setState({ submittingComment: true });
      await this.socialService.commentOnPolicy({ policyId: policy.Id, commentText: newComment });
      const comments = await this.socialService.getPolicyComments(policy.Id);
      this.setState({ comments, newComment: '', showCommentDialog: false, submittingComment: false });
    } catch (error) {
      console.error('Failed to submit comment:', error);
      this.setState({ submittingComment: false });
    }
  };

  private handleFollow = async (): Promise<void> => {
    const { policy, isFollowing } = this.state;
    if (!policy) return;

    try {
      if (isFollowing) {
        await this.socialService.unfollowPolicy(policy.Id);
      } else {
        await this.socialService.followPolicy({
          policyId: policy.Id,
          notifyOnUpdate: true,
          notifyOnComment: true,
          notifyOnNewVersion: true
        });
      }
      this.setState({ isFollowing: !isFollowing });
    } catch (error) {
      console.error('Failed to follow/unfollow policy:', error);
    }
  };

  private handleShare = async (): Promise<void> => {
    const { policy } = this.state;
    if (!policy) return;
    const url = window.location.href;
    if (navigator.share) {
      try {
        await navigator.share({ title: policy.PolicyName, text: `Check out this policy: ${policy.PolicyNumber}`, url });
      } catch { /* User cancelled */ }
    } else {
      await navigator.clipboard.writeText(url);
      await this.dialogManager.showAlert('Link copied to clipboard!', { variant: 'success' });
    }
  };

  // ============================================
  // READ FLOW METHODS
  // ============================================

  private getDeviceType(): string {
    const ua = navigator.userAgent;
    if (/tablet|ipad|playbook|silk/i.test(ua)) return 'Tablet';
    if (/mobile|android|iphone|ipod/i.test(ua)) return 'Mobile';
    return 'Desktop';
  }

  private getBrowserName(): string {
    const ua = navigator.userAgent;
    if (ua.includes('Chrome')) return 'Chrome';
    if (ua.includes('Firefox')) return 'Firefox';
    if (ua.includes('Safari')) return 'Safari';
    if (ua.includes('Edge')) return 'Edge';
    if (ua.includes('MSIE') || ua.includes('Trident')) return 'Internet Explorer';
    return 'Unknown';
  }

  private generateReceiptNumber(): string {
    const timestamp = Date.now().toString(36).toUpperCase();
    const random = Math.random().toString(36).substring(2, 6).toUpperCase();
    return `RR-${timestamp}-${random}`;
  }

  private handleMarkAsRead = (): void => {
    const { policy } = this.state;
    if (!policy) return;

    if (policy.RequiresQuiz) {
      this.setState({
        hasReadPolicy: true,
        currentFlowStep: 'quiz',
        quizRequired: true
      });
    } else {
      this.setState({
        hasReadPolicy: true,
        currentFlowStep: 'acknowledge',
        showAcknowledgePanel: true
      });
    }
  };

  private handleQuizComplete = (score: number, passed: boolean): void => {
    this.setState({
      quizScore: score,
      quizPassed: passed,
      quizCompleted: true
    });

    if (passed) {
      this.setState({
        currentFlowStep: 'acknowledge',
        showAcknowledgePanel: true
      });
    } else {
      this.setState({
        error: 'You did not pass the quiz. Please review the policy and try again.'
      });
    }
  };

  private handleOpenAcknowledgePanel = (): void => {
    this.setState({ showAcknowledgePanel: true });
  };

  private handleCloseAcknowledgePanel = (): void => {
    this.setState({
      showAcknowledgePanel: false,
      legalAgreement1: false,
      legalAgreement2: false,
      legalAgreement3: false,
      digitalSignature: ''
    });
  };

  private canSubmitAcknowledgement = (): boolean => {
    const { legalAgreement1, legalAgreement2, legalAgreement3, digitalSignature } = this.state;
    return legalAgreement1 && legalAgreement2 && legalAgreement3 && digitalSignature.trim().length >= 3;
  };

  private handleSubmitAcknowledgement = async (): Promise<void> => {
    const {
      policy, acknowledgement, readStartTime, readDuration,
      quizRequired, quizCompleted, quizScore, digitalSignature, acknowledgeNotes
    } = this.state;

    if (!policy) return;
    if (!this.canSubmitAcknowledgement()) {
      this.setState({ error: 'Please complete all acknowledgement requirements.' });
      return;
    }

    // Declare readReceipt in outer scope so catch block can access it
    let readReceipt: IReadReceipt | undefined;

    try {
      this.setState({ submittingAcknowledgement: true, error: null });

      // Get current user — fallback to placeholder if SP call fails
      let currentUser: { Id: number; Email: string; Title: string } = { Id: 0, Email: '', Title: digitalSignature };
      try {
        currentUser = await this.props.sp.web.currentUser();
      } catch (userErr) {
        console.warn('Could not fetch current user (using fallback):', userErr);
      }
      const now = new Date();

      const legalText = `I, ${digitalSignature}, hereby confirm that:
1. I have read and fully understood the policy "${policy.PolicyName}" (${policy.PolicyNumber}).
2. I agree to comply with all requirements and guidelines outlined in this policy.
3. I understand that failure to comply may result in disciplinary action.
4. I acknowledge that this constitutes my electronic signature and consent.`;

      readReceipt = {
        UserId: currentUser.Id,
        UserEmail: currentUser.Email,
        UserDisplayName: currentUser.Title,
        PolicyId: policy.Id,
        PolicyNumber: policy.PolicyNumber,
        PolicyName: policy.PolicyName,
        PolicyVersion: policy.VersionNumber?.toString() || '1.0',
        ReadStartTime: readStartTime || now,
        ReadEndTime: now,
        ReadDurationSeconds: readDuration,
        QuizRequired: quizRequired,
        QuizCompleted: quizCompleted,
        QuizScore: quizScore,
        QuizPassPercentage: policy.QuizPassingScore || 80,
        QuizPassedDate: quizCompleted ? now : undefined,
        AcknowledgedDate: now,
        AcknowledgedTime: now.toLocaleTimeString(),
        IPAddress: 'Captured server-side',
        UserAgent: navigator.userAgent,
        DeviceType: this.getDeviceType(),
        BrowserName: this.getBrowserName(),
        DigitalSignature: digitalSignature,
        LegalConfirmationText: legalText,
        Notes: acknowledgeNotes,
        ReceiptNumber: this.generateReceiptNumber()
      };

      // Try to save to SharePoint — but always advance to complete even if SP calls fail
      try {
        await this.saveReadReceipt(readReceipt);
      } catch (saveErr) {
        console.warn('Could not save read receipt (list may not exist yet):', saveErr);
      }

      try {
        if (acknowledgement && acknowledgement.Id) {
          // Update existing acknowledgement record
          const request: IPolicyAcknowledgeRequest = {
            acknowledgementId: acknowledgement.Id,
            acknowledgedDate: now,
            notes: acknowledgeNotes,
            readDuration: readDuration,
            ipAddress: '',
            userAgent: navigator.userAgent,
            quizScore: quizCompleted ? quizScore : undefined
          };
          await this.policyService.acknowledgePolicy(request);
        } else {
          // No existing record — create a new acknowledgement directly
          try {
            await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_ACKNOWLEDGEMENTS).items.add({
              Title: `${policy.PolicyNumber} - ${currentUser.Title}`,
              PolicyId: policy.Id,
              AckUserId: currentUser.Id,
              UserEmail: currentUser.Email,
              AckStatus: 'Acknowledged',
              AcknowledgedDate: now.toISOString()
            });
          } catch (createErr) {
            console.warn('Could not create acknowledgement record:', createErr);
          }
        }
      } catch (ackErr) {
        console.warn('Could not update acknowledgement record:', ackErr);
      }

      // Always advance to complete — even if SP saves failed
      if (this._isMounted) {
        this.setState({
          readReceipt,
          showAcknowledgePanel: false,
          showCongratulationsPanel: true,
          currentFlowStep: 'complete',
          submittingAcknowledgement: false
        });
      }
    } catch (error: any) {
      console.error('Failed to submit acknowledgement:', error);
      // Still advance to complete — the user has made their declaration
      if (this._isMounted) {
        this.setState({
          readReceipt,
          showAcknowledgePanel: false,
          showCongratulationsPanel: true,
          currentFlowStep: 'complete',
          submittingAcknowledgement: false
        });
      }
    }
  };

  private saveReadReceipt = async (receipt: IReadReceipt): Promise<void> => {
    try {
      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_READ_RECEIPTS).items.add({
        Title: receipt.ReceiptNumber,
        UserId: receipt.UserId,
        UserEmail: receipt.UserEmail,
        UserDisplayName: receipt.UserDisplayName,
        PolicyId: receipt.PolicyId,
        PolicyNumber: receipt.PolicyNumber,
        PolicyName: receipt.PolicyName,
        PolicyVersion: receipt.PolicyVersion,
        ReadStartTime: receipt.ReadStartTime.toISOString(),
        ReadEndTime: receipt.ReadEndTime.toISOString(),
        ReadDurationSeconds: receipt.ReadDurationSeconds,
        QuizRequired: receipt.QuizRequired,
        QuizCompleted: receipt.QuizCompleted,
        QuizScore: receipt.QuizScore,
        QuizPassPercentage: receipt.QuizPassPercentage,
        QuizPassedDate: receipt.QuizPassedDate?.toISOString(),
        AcknowledgedDate: receipt.AcknowledgedDate.toISOString(),
        AcknowledgedTime: receipt.AcknowledgedTime,
        IPAddress: receipt.IPAddress,
        UserAgent: receipt.UserAgent,
        DeviceType: receipt.DeviceType,
        BrowserName: receipt.BrowserName,
        DigitalSignature: receipt.DigitalSignature,
        LegalConfirmationText: receipt.LegalConfirmationText,
        Notes: receipt.Notes,
        ReceiptNumber: receipt.ReceiptNumber
      });
    } catch (error) {
      console.error('Failed to save read receipt:', error);
    }
  };

  private handleEmailReceipt = async (): Promise<void> => {
    const { readReceipt, policy } = this.state;
    if (!readReceipt || !policy) return;

    try {
      this.setState({ emailingReceipt: true });
      const subject = encodeURIComponent(`Policy Acknowledgement Receipt - ${policy.PolicyNumber}`);
      const body = encodeURIComponent(
        `Policy Read Receipt\n\n` +
        `Receipt Number: ${readReceipt.ReceiptNumber}\n` +
        `Policy: ${readReceipt.PolicyNumber} - ${readReceipt.PolicyName}\n` +
        `Version: ${readReceipt.PolicyVersion}\n` +
        `Acknowledged: ${readReceipt.AcknowledgedDate?.toLocaleDateString()} at ${readReceipt.AcknowledgedTime}\n` +
        `Digital Signature: ${readReceipt.DigitalSignature}\n\n` +
        `This email confirms that you have read and acknowledged the above policy.\n` +
        `Please retain this email for your records.`
      );
      window.open(`mailto:${readReceipt.UserEmail}?subject=${subject}&body=${body}`, '_blank');
      this.setState({ emailingReceipt: false });
      await this.dialogManager.showAlert('Your email client has been opened with the receipt details.', { variant: 'info' });
    } catch (error) {
      console.error('Failed to open email client:', error);
      this.setState({ emailingReceipt: false, error: 'Failed to open email client. Please try again.' });
    }
  };

  private generateReceiptEmailHtml(receipt: IReadReceipt): string {
    const { ReportHtmlGenerator } = require('../../../utils/reportHtmlGenerator');
    return ReportHtmlGenerator.generate({
      title: 'Policy Read Receipt',
      subtitle: `Receipt ${receipt.ReceiptNumber}`,
      reportType: 'ACKNOWLEDGEMENT',
      reportId: receipt.ReceiptNumber,
      sections: [
        {
          type: 'summary-card',
          title: 'Acknowledgement Confirmed',
          content: `<p>${receipt.UserDisplayName} has read and acknowledged this policy.</p>`,
          style: 'success',
          data: { badge: 'VERIFIED' }
        },
        {
          type: 'two-column',
          title: 'Employee Details',
          subtitle: 'Policy Details',
          data: {
            left: [
              { label: 'Name', value: receipt.UserDisplayName },
              { label: 'Email', value: receipt.UserEmail },
              { label: 'Read Duration', value: `${Math.floor(receipt.ReadDurationSeconds / 60)} min ${receipt.ReadDurationSeconds % 60} sec` },
              { label: 'Acknowledged', value: `${receipt.AcknowledgedDate.toLocaleDateString()} at ${receipt.AcknowledgedTime}` }
            ],
            right: [
              { label: 'Policy', value: `${receipt.PolicyNumber} - ${receipt.PolicyName}` },
              { label: 'Version', value: receipt.PolicyVersion },
              ...(receipt.QuizRequired ? [{ label: 'Quiz Score', value: `${receipt.QuizScore}%` }] : []),
              { label: 'Receipt Number', value: receipt.ReceiptNumber }
            ]
          }
        },
        {
          type: 'badge-row',
          data: [
            { label: 'Status', value: 'Acknowledged', color: '#059669' },
            { label: 'Policy', value: receipt.PolicyNumber, color: '#0d9488' },
            ...(receipt.QuizRequired ? [{ label: 'Quiz', value: `${receipt.QuizScore}%`, color: receipt.QuizScore >= 75 ? '#059669' : '#dc2626' }] : [])
          ]
        },
        { type: 'divider' },
        {
          type: 'summary-card',
          title: 'Legal Confirmation',
          content: `<p style="white-space: pre-line; font-size: 9pt;">${receipt.LegalConfirmationText}</p>`,
          style: 'muted'
        },
        {
          type: 'text',
          content: '<p style="font-size: 9pt; color: #94a3b8; text-align: center; margin-top: 16px;">This is an automated receipt from Policy Manager. Please retain for your records.</p>',
          style: 'muted'
        },
        {
          type: 'signature',
          data: {
            name: receipt.DigitalSignature,
            role: 'Digital Signature',
            date: `${receipt.AcknowledgedDate.toLocaleDateString()} ${receipt.AcknowledgedTime}`
          }
        }
      ]
    });
  }

  private handleViewReceipt = (): void => {
    this.setState({ showReadReceiptPanel: true });
  };

  private handleGeneratePdf = (): void => {
    const { readReceipt } = this.state;
    if (!readReceipt) return;
    const html = this.generateReceiptEmailHtml(readReceipt);
    const blob = new Blob([html], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const printWindow = window.open(url, '_blank');
    if (printWindow) {
      printWindow.addEventListener('afterprint', () => URL.revokeObjectURL(url));
      printWindow.addEventListener('load', () => printWindow.print());
    } else {
      URL.revokeObjectURL(url);
    }
  };

  private handleCloseCongratulations = (): void => {
    this.setState({ showCongratulationsPanel: false });
    this.loadPolicyDetails();
  };

  // ============================================
  // WIZARD NAVIGATION
  // ============================================

  private getWizardSteps(): Array<{ key: ReadFlowStep; label: string; icon: string }> {
    const { quizRequired, policy } = this.state;
    const steps: Array<{ key: ReadFlowStep; label: string; icon: string }> = [
      { key: 'reading', label: 'Read Policy', icon: 'Read' }
    ];
    if (quizRequired || policy?.RequiresQuiz) {
      steps.push({ key: 'quiz', label: 'Quiz', icon: 'Questionnaire' });
    }
    steps.push({ key: 'acknowledge', label: 'Acknowledge', icon: 'Handwriting' });
    steps.push({ key: 'complete', label: 'Complete', icon: 'CheckMark' });
    return steps;
  }

  private getStepIndex(step: ReadFlowStep): number {
    return this.getWizardSteps().findIndex(s => s.key === step);
  }

  private isStepCompleted(stepKey: ReadFlowStep): boolean {
    const { currentFlowStep, hasReadPolicy, quizPassed } = this.state;
    const currentIndex = this.getStepIndex(currentFlowStep);
    const stepIndex = this.getStepIndex(stepKey);

    if (stepIndex < currentIndex) return true;
    if (stepKey === 'reading' && hasReadPolicy) return true;
    if (stepKey === 'quiz' && quizPassed) return true;
    if (stepKey === 'complete' && currentFlowStep === 'complete') return true;
    return false;
  }

  private handleWizardBack = (): void => {
    const steps = this.getWizardSteps();
    const currentIndex = this.getStepIndex(this.state.currentFlowStep);
    if (currentIndex > 0) {
      this.setState({ currentFlowStep: steps[currentIndex - 1].key });
    }
  };

  private handleWizardNext = (): void => {
    const { currentFlowStep, hasReadPolicy, quizPassed, policy } = this.state;
    const steps = this.getWizardSteps();
    const currentIndex = this.getStepIndex(currentFlowStep);

    if (currentFlowStep === 'reading') {
      if (!hasReadPolicy) {
        this.handleMarkAsRead();
      } else if (policy?.RequiresQuiz) {
        this.setState({ currentFlowStep: 'quiz' });
      } else {
        this.setState({ currentFlowStep: 'acknowledge', showAcknowledgePanel: true });
      }
    } else if (currentFlowStep === 'quiz') {
      if (quizPassed) {
        this.setState({ currentFlowStep: 'acknowledge', showAcknowledgePanel: true });
      }
    } else if (currentFlowStep === 'acknowledge') {
      this.handleOpenAcknowledgePanel();
    }
  };

  private canGoNext(): boolean {
    const { currentFlowStep, scrollProgress, hasReadPolicy, quizPassed, quizRequired, readDuration } = this.state;
    if (currentFlowStep === 'reading') {
      // scrollProgress tracks div scrolling (works for inline/placeholder docs).
      // For iframe-based documents, scroll events don't bubble, so we also
      // allow proceeding after a minimum read time of 30 seconds.
      return scrollProgress >= 95 || hasReadPolicy || readDuration >= 30;
    }
    if (currentFlowStep === 'quiz') return quizPassed;
    if (currentFlowStep === 'acknowledge') return false; // panel handles it
    return false;
  }

  private getNextButtonText(): string {
    const { currentFlowStep, hasReadPolicy } = this.state;
    if (currentFlowStep === 'reading') {
      return hasReadPolicy ? 'Next' : 'I Have Read This Policy';
    }
    if (currentFlowStep === 'quiz') return 'Proceed to Acknowledge';
    if (currentFlowStep === 'acknowledge') return 'Open Acknowledgement';
    return 'Next';
  }

  // ============================================
  // HORIZONTAL QUIZ METHODS
  // ============================================

  private handleQuizSelectAnswer = (questionIndex: number, optionIndex: number): void => {
    if (this.state.quizSubmitted) return;
    if (this.state.quizConfirmed[questionIndex]) return; // Already confirmed — locked
    const newAnswers = [...this.state.quizAnswers];
    newAnswers[questionIndex] = optionIndex;
    this.setState({ quizAnswers: newAnswers });
    // No auto-advance — user must click "Confirm Answer"
  };

  private handleQuizConfirmAnswer = (): void => {
    const { currentQuizQuestion, quizAnswers } = this.state;
    if (quizAnswers[currentQuizQuestion] < 0) return; // No answer selected

    const qi = currentQuizQuestion;
    const q = MOCK_QUIZ_QUESTIONS[qi];
    const isCorrect = quizAnswers[qi] === q.correctIndex;
    const newAttempts = [...this.state.quizAttempts];
    newAttempts[qi] = (newAttempts[qi] || 0) + 1;

    if (isCorrect || newAttempts[qi] >= 2) {
      // Correct on any attempt, or used both attempts — lock the answer and reveal
      const newConfirmed = [...this.state.quizConfirmed];
      newConfirmed[qi] = true;
      this.setState({ quizConfirmed: newConfirmed, quizAttempts: newAttempts });
    } else {
      // First incorrect attempt — let them try again (reset selection, keep attempt count)
      const newAnswers = [...this.state.quizAnswers];
      newAnswers[qi] = -1; // Clear selection so they can pick again
      this.setState({ quizAnswers: newAnswers, quizAttempts: newAttempts });
    }
  };

  private handleQuizNextAfterConfirm = (): void => {
    const { currentQuizQuestion } = this.state;
    if (currentQuizQuestion < MOCK_QUIZ_QUESTIONS.length - 1) {
      this.setState({ currentQuizQuestion: currentQuizQuestion + 1 });
    }
  };

  private handleQuizSubmit = (): void => {
    const { quizAnswers } = this.state;
    let correct = 0;
    MOCK_QUIZ_QUESTIONS.forEach((q, i) => {
      if (quizAnswers[i] === q.correctIndex) correct++;
    });

    const pct = Math.round((correct / MOCK_QUIZ_QUESTIONS.length) * 100);
    const passed = pct >= 80;

    this.setState({
      quizSubmitted: true,
      quizScore: pct,
      quizPassed: passed,
      quizCompleted: true
    });

    if (passed) {
      // Auto-advance to acknowledge after brief pause
      setTimeout(() => {
        this.setState({
          currentFlowStep: 'acknowledge',
          showAcknowledgePanel: true
        });
      }, 2000);
    }
  };

  private handleQuizRetake = (): void => {
    this.setState({
      quizSubmitted: false,
      quizPassed: false,
      quizCompleted: false,
      quizScore: 0,
      quizAnswers: new Array(MOCK_QUIZ_QUESTIONS.length).fill(-1),
      quizConfirmed: new Array(MOCK_QUIZ_QUESTIONS.length).fill(false),
      quizAttempts: new Array(MOCK_QUIZ_QUESTIONS.length).fill(0),
      currentQuizQuestion: 0,
      error: null
    });
  };

  private allQuizAnswered(): boolean {
    return this.state.quizAnswers.every(a => a >= 0);
  }

  // ============================================
  // DOCUMENT VIEWER
  // ============================================

  private getDocumentViewerUrl(documentUrl: string): string {
    if (!documentUrl) return '';
    const ext = documentUrl.split('.').pop()?.toLowerCase() || '';
    const siteUrl = this.props.context.pageContext.web.absoluteUrl;

    if (['docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt'].includes(ext)) {
      // wdHideHeaders hides the Office Online toolbar (Edit Document, Print, Share, etc.)
      return `${siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(documentUrl)}&action=view&wdHideHeaders=1`;
    }
    if (ext === 'pdf') return documentUrl;
    if (['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp'].includes(ext)) return documentUrl;
    return `${siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(documentUrl)}&action=view&wdHideHeaders=1`;
  }

  private getDocumentIcon(documentUrl: string): string {
    const ext = documentUrl?.split('.').pop()?.toLowerCase() || '';
    const iconMap: Record<string, string> = {
      pdf: 'PDF', docx: 'WordDocument', doc: 'WordDocument',
      xlsx: 'ExcelDocument', xls: 'ExcelDocument',
      pptx: 'PowerPointDocument', ppt: 'PowerPointDocument',
      jpg: 'FileImage', jpeg: 'FileImage', png: 'FileImage', gif: 'FileImage',
      svg: 'FileImage', webp: 'FileImage'
    };
    return iconMap[ext] || 'Page';
  }

  private getDocumentTypeLabel(documentUrl: string): string {
    const ext = documentUrl?.split('.').pop()?.toLowerCase() || '';
    const labelMap: Record<string, string> = {
      pdf: 'PDF Document', docx: 'Word Document', doc: 'Word Document',
      xlsx: 'Excel Spreadsheet', xls: 'Excel Spreadsheet',
      pptx: 'PowerPoint Presentation', ppt: 'PowerPoint Presentation',
      jpg: 'Image (JPEG)', jpeg: 'Image (JPEG)', png: 'Image (PNG)',
      gif: 'Image (GIF)', svg: 'Image (SVG)', webp: 'Image (WebP)'
    };
    return labelMap[ext] || 'Document';
  }

  private handleDocumentScroll = (): void => {
    const viewer = this.documentViewerRef.current;
    if (!viewer) return;
    const pct = (viewer.scrollTop / (viewer.scrollHeight - viewer.clientHeight)) * 100;
    const scrollProgress = Math.min(pct, 100);
    this.setState({ scrollProgress });
    if (scrollProgress >= 95 && !this.state.hasReadPolicy) {
      // Don't auto-mark, but enable the Next button
    }
  };

  // ============================================
  // RENDER: WIZARD PROGRESS STEPPER
  // ============================================

  private renderWizardStepper(): JSX.Element {
    const { currentFlowStep } = this.state;
    const steps = this.getWizardSteps();
    const currentIndex = this.getStepIndex(currentFlowStep);

    return (
      <div className={styles.wizardStepper}>
        {steps.map((step, index) => {
          const isActive = step.key === currentFlowStep;
          const isCompleted = this.isStepCompleted(step.key) && !isActive;
          const isFuture = index > currentIndex && !isCompleted;

          return (
            <React.Fragment key={step.key}>
              {index > 0 && (
                <div className={`${styles.stepConnector} ${index <= currentIndex ? styles.done : ''}`} />
              )}
              <div className={`${styles.wizardStep} ${isActive ? styles.active : ''} ${isCompleted ? styles.completed : ''} ${isFuture ? styles.future : ''}`}>
                <div className={styles.stepCircle}>
                  {isCompleted ? (
                    <Icon iconName="CheckMark" styles={{ root: { fontSize: 14 } }} />
                  ) : (
                    <span>{index + 1}</span>
                  )}
                </div>
                <span className={styles.stepLabel}>{step.label}</span>
              </div>
            </React.Fragment>
          );
        })}
      </div>
    );
  }

  // ============================================
  // RENDER: STEP 1 — READ POLICY
  // ============================================

  private renderReadStep(): JSX.Element | null {
    const { policy, acknowledgement, isFollowing, readDuration, scrollProgress } = this.state;
    if (!policy) return null;

    const statusColor = acknowledgement?.AckStatus === 'Acknowledged' ? '#16a34a' :
                         acknowledgement?.AckStatus === 'Overdue' ? '#dc2626' : '#d97706';

    // DocumentURL may be a string or a SharePoint FieldUrlValue object { Url, Description }
    const rawDocUrl = policy.DocumentURL;
    const documentUrl: string | undefined = typeof rawDocUrl === 'string' ? rawDocUrl
      : (rawDocUrl && typeof rawDocUrl === 'object' && (rawDocUrl as { Url?: string }).Url)
        ? (rawDocUrl as { Url: string }).Url
        : undefined;
    const attachments = policy.AttachmentURLs || [];
    const hasDocuments = documentUrl || attachments.length > 0;
    const ext = documentUrl?.split('.').pop()?.toLowerCase() || '';
    const isImage = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp'].includes(ext);
    const viewerUrl = documentUrl ? this.getDocumentViewerUrl(documentUrl) : '';

    // In focused reader mode, the metadata is already in the compact bar above
    const isFocusedMode = !this.state.browseMode && (!this.state.acknowledgement || this.state.acknowledgement.AckStatus !== 'Acknowledged');

    // If PolicyContent has converted HTML, render it natively (no iframe needed)
    const hasConvertedHtml = policy.PolicyContent && policy.PolicyContent.length > 100 && policy.PolicyContent.includes('<');

    return (
      <div className={styles.stepContent}>
        {/* Compact Policy Header — collapsible accordion for maximum reading space */}
        {!isFocusedMode && (
          <div style={{
            background: '#fff',
            border: '1px solid #e2e8f0',
            borderLeft: '4px solid #0d9488',
            borderRadius: 4,
            marginTop: 5,
            marginBottom: 12,
            overflow: 'hidden'
          }}>
            {/* Header Row — always visible */}
            <div
              role="button"
              tabIndex={0}
              onClick={() => this.setState({ _headerExpanded: !(this.state as any)._headerExpanded } as any)}
              onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') this.setState({ _headerExpanded: !(this.state as any)._headerExpanded } as any); }}
              style={{
                display: 'flex', alignItems: 'center', padding: '10px 16px',
                cursor: 'pointer', gap: 12, userSelect: 'none'
              }}
            >
              <Icon iconName={(this.state as any)._headerExpanded ? 'ChevronUp' : 'ChevronDown'}
                styles={{ root: { fontSize: 12, color: '#94a3b8', transition: 'transform 0.2s' } }} />
              <Text style={{ fontWeight: 700, color: '#0f172a', fontSize: 15, flex: 1 }}>
                {policy.PolicyNumber} - {policy.PolicyName}
              </Text>
              <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center" wrap>
                <span className={styles.badgeGreen}>Published</span>
                <span className={styles.badgeTeal}>{policy.PolicyCategory}</span>
                <span className={styles.badgeSlate}>{policy.PolicyNumber}</span>
                <span className={styles.badgeSlate}>{policy.EffectiveDate ? new Date(policy.EffectiveDate).toLocaleDateString() : 'N/A'}</span>
                <span className={styles.badgeSlate}>v{policy.VersionNumber || '1.0'}</span>
                {acknowledgement && (
                  <span className={styles.badgeAmber} style={{ backgroundColor: statusColor === '#16a34a' ? '#dcfce7' : statusColor === '#dc2626' ? '#fee2e2' : '#fef3c7', color: statusColor }}>
                    {acknowledgement.AckStatus || acknowledgement.Status}
                  </span>
                )}
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center" style={{ marginLeft: 8 }}>
                <div className={styles.readTimer}>
                  <Icon iconName="Timer" styles={{ root: { fontSize: 12 } }} />
                  <span>{this.formatDuration(readDuration)}</span>
                </div>
                <IconButton
                  iconProps={{ iconName: 'History' }}
                  title="Version History"
                  ariaLabel="Version History"
                  onClick={(e) => { e.stopPropagation(); this.loadVersionHistory(); }}
                  styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 14, color: '#0d9488' } }}
                />
                <IconButton
                  iconProps={{ iconName: isFollowing ? 'FavoriteStarFill' : 'FavoriteStar' }}
                  title={isFollowing ? 'Unfollow' : 'Follow'}
                  ariaLabel={isFollowing ? 'Unfollow policy' : 'Follow policy'}
                  onClick={(e) => { e.stopPropagation(); this.handleFollow(); }}
                  styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 14, color: isFollowing ? '#f59e0b' : '#94a3b8' } }}
                />
                <IconButton
                  iconProps={{ iconName: 'Share' }}
                  title="Share"
                  ariaLabel="Share policy"
                  onClick={(e) => { e.stopPropagation(); this.handleShare(); }}
                  styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 14, color: '#0d9488' } }}
                />
              </Stack>
            </div>

            {/* Expanded Details — collapsible */}
            {(this.state as any)._headerExpanded && (
              <div style={{ padding: '0 16px 12px', borderTop: '1px solid #f1f5f9' }}>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: 12, paddingTop: 12 }}>
                  <div>
                    <Text variant="tiny" style={{ color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px', display: 'block' }}>Department</Text>
                    <Text variant="small" style={{ color: '#334155', fontWeight: 500 }}>{policy.PolicyCategory || 'General'}</Text>
                  </div>
                  <div>
                    <Text variant="tiny" style={{ color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px', display: 'block' }}>Effective Date</Text>
                    <Text variant="small" style={{ color: '#334155', fontWeight: 500 }}>{policy.EffectiveDate ? new Date(policy.EffectiveDate).toLocaleDateString() : 'N/A'}</Text>
                  </div>
                  <div>
                    <Text variant="tiny" style={{ color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px', display: 'block' }}>Version</Text>
                    <Text variant="small" style={{ color: '#334155', fontWeight: 500 }}>v{policy.VersionNumber || '1.0'}</Text>
                  </div>
                  {acknowledgement?.DueDate && (
                    <div>
                      <Text variant="tiny" style={{ color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px', display: 'block' }}>Ack. Due</Text>
                      <Text variant="small" style={{ color: statusColor, fontWeight: 500 }}>{new Date(acknowledgement.DueDate).toLocaleDateString()}</Text>
                    </div>
                  )}
                  {policy.PolicySummary && (
                    <div style={{ gridColumn: '1 / -1' }}>
                      <Text variant="tiny" style={{ color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px', display: 'block' }}>Summary</Text>
                      <Text variant="small" style={{ color: '#334155' }}>{policy.PolicySummary}</Text>
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        )}

        {/* Native HTML Reader — rendered when PolicyContent has converted HTML */}
        {hasConvertedHtml && (
          <div className={styles.documentViewerWrapper} ref={this.viewerWrapperRef} style={{
            ...(this.state.isFullscreen ? { background: '#fff', padding: 0 } : {}),
            ...(isFocusedMode ? { marginBottom: 0 } : {})
          }}>
            {/* Toolbar — matches iframe viewer chrome */}
            {!isFocusedMode && (
              <div className={styles.viewerToolbar}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Icon iconName={documentUrl ? this.getDocumentIcon(documentUrl) : 'FileHTML'} style={{ fontSize: 16, color: '#0d9488' }} />
                  <Text variant="small" style={{ fontWeight: 600, color: '#334155' }}>
                    {documentUrl ? documentUrl.split('/').pop() : `${policy.PolicyNumber || ''} ${policy.PolicyName || policy.Title}`.trim()}
                  </Text>
                  <Text variant="tiny" style={{ color: '#94a3b8' }}>
                    {documentUrl ? this.getDocumentTypeLabel(documentUrl) : 'HTML Document'}
                  </Text>
                </Stack>
                <Stack horizontal tokens={{ childrenGap: 8 }}>
                  <IconButton
                    iconProps={{ iconName: 'Print' }}
                    title="Print"
                    ariaLabel="Print policy"
                    onClick={() => {
                      const printWindow = window.open('', '_blank');
                      if (printWindow) {
                        printWindow.document.write(`<!DOCTYPE html><html><head><title>${policy.PolicyName || 'Policy'}</title></head><body>${sanitizeHtml(policy.PolicyContent || '')}</body></html>`);
                        printWindow.document.close();
                        printWindow.print();
                      }
                    }}
                    styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 14, color: '#0d9488' } }}
                  />
                  <IconButton
                    iconProps={{ iconName: this.state.isFullscreen ? 'BackToWindow' : 'FullScreen' }}
                    title={this.state.isFullscreen ? 'Exit Fullscreen' : 'Fullscreen'}
                    ariaLabel={this.state.isFullscreen ? 'Exit Fullscreen' : 'Fullscreen'}
                    onClick={this.toggleFullscreen}
                    styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 14, color: '#0d9488' } }}
                  />
                  {documentUrl && (
                    <>
                      <DefaultButton
                        iconProps={{ iconName: 'OpenInNewTab' }}
                        text="Open Original"
                        href={documentUrl}
                        target="_blank"
                        styles={{ root: { height: 28, padding: '0 10px' }, label: { fontSize: 11 } }}
                      />
                      <DefaultButton
                        iconProps={{ iconName: 'Download' }}
                        text="Download"
                        href={documentUrl}
                        styles={{ root: { height: 28, padding: '0 10px' }, label: { fontSize: 11 } }}
                      />
                    </>
                  )}
                </Stack>
              </div>
            )}
            <div
              className={styles.documentViewer}
              ref={this.documentViewerRef}
              onScroll={this.handleDocumentScroll}
              style={{
                padding: isFocusedMode ? '32px 48px 80px' : '24px 32px',
                ...(isFocusedMode ? { height: 'calc(100vh - 140px)' } : {}),
                ...(this.state.isFullscreen ? { height: 'calc(100vh - 80px)' } : {})
              }}
            >
              <div className={styles.scrollProgressBar}>
                <div className={styles.scrollProgressFill} style={{ height: `${scrollProgress}%` }} />
              </div>
              <div dangerouslySetInnerHTML={{ __html: sanitizeHtml(policy.PolicyContent || '') }} />
            </div>
          </div>
        )}

        {/* Document Viewer (iframe fallback — only when no converted HTML available) */}
        {!hasConvertedHtml && hasDocuments && documentUrl && (
          <div className={styles.documentViewerWrapper} ref={this.viewerWrapperRef} style={{
            ...(this.state.isFullscreen ? { background: '#fff', padding: 0 } : {}),
            ...(isFocusedMode ? { marginBottom: 0 } : {})
          }}>
            {/* Toolbar hidden in focused mode — title already in metadata strip */}
            {!isFocusedMode && (
              <div className={styles.viewerToolbar}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Icon iconName={this.getDocumentIcon(documentUrl)} style={{ fontSize: 16, color: '#0d9488' }} />
                  <Text variant="small" style={{ fontWeight: 600, color: '#334155' }}>
                    {documentUrl.split('/').pop()}
                  </Text>
                  <Text variant="tiny" style={{ color: '#94a3b8' }}>
                    {this.getDocumentTypeLabel(documentUrl)}
                  </Text>
                </Stack>
                <Stack horizontal tokens={{ childrenGap: 8 }}>
                  <IconButton
                    iconProps={{ iconName: this.state.isFullscreen ? 'BackToWindow' : 'FullScreen' }}
                    title={this.state.isFullscreen ? 'Exit Fullscreen' : 'Fullscreen'}
                    ariaLabel={this.state.isFullscreen ? 'Exit Fullscreen' : 'Fullscreen'}
                    onClick={this.toggleFullscreen}
                    styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 14, color: '#0d9488' } }}
                  />
                  <DefaultButton
                    iconProps={{ iconName: 'OpenInNewTab' }}
                    text="Open"
                    href={documentUrl}
                    target="_blank"
                    styles={{ root: { height: 28, padding: '0 10px' }, label: { fontSize: 11 } }}
                  />
                  <DefaultButton
                    iconProps={{ iconName: 'Download' }}
                    text="Download"
                    href={documentUrl}
                    styles={{ root: { height: 28, padding: '0 10px' }, label: { fontSize: 11 } }}
                  />
                </Stack>
              </div>
            )}
            <div
              className={styles.documentViewer}
              ref={this.documentViewerRef}
              onScroll={this.handleDocumentScroll}
              style={isFocusedMode ? { height: 600 } : undefined}
            >
              <div className={styles.scrollProgressBar}>
                <div className={styles.scrollProgressFill} style={{ height: `${scrollProgress}%` }} />
              </div>
              {isImage ? (
                <div style={{ textAlign: 'center', padding: 20 }}>
                  <img src={viewerUrl} alt={policy.Title} style={{ maxWidth: '100%', maxHeight: 600, borderRadius: 4 }} />
                </div>
              ) : ext === 'pdf' ? (
                /* PDF — native browser embed, zero chrome */
                <object
                  data={`${documentUrl}#toolbar=0&navpanes=0&scrollbar=1`}
                  type="application/pdf"
                  style={{ width: '100%', height: this.state.isFullscreen ? 'calc(100vh - 80px)' : '100%', border: 'none' }}
                  title={`${policy.Title} PDF Viewer`}
                >
                  {/* Fallback if browser can't render PDF inline */}
                  <iframe src={documentUrl} style={{ width: '100%', height: '100%', border: 'none' }} title={`${policy.Title} PDF Viewer`} />
                </object>
              ) : (
                /* Office docs — Office Online iframe with hidden headers */
                <iframe src={viewerUrl} style={{ width: '100%', height: this.state.isFullscreen ? 'calc(100vh - 80px)' : '100%', border: 'none' }} title={`${policy.Title} Document Viewer`} />
              )}
            </div>
          </div>
        )}

        {/* Placeholder Document Viewer (when no real document URL) */}
        {!hasDocuments && (
          <div className={styles.documentViewerWrapper}>
            <div className={styles.viewerToolbar}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Icon iconName="PDF" style={{ fontSize: 16, color: '#0d9488' }} />
                <Text variant="small" style={{ fontWeight: 600, color: '#334155' }}>
                  {policy.PolicyNumber}-{(policy.PolicyName || policy.Title).replace(/\s+/g, '-')}.pdf
                </Text>
                <Text variant="tiny" style={{ color: '#94a3b8' }}>PDF Document</Text>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <DefaultButton
                  iconProps={{ iconName: 'OpenInNewTab' }}
                  text="Open"
                  disabled={true}
                  styles={{ root: { height: 28, padding: '0 10px' }, label: { fontSize: 11 } }}
                />
                <DefaultButton
                  iconProps={{ iconName: 'Download' }}
                  text="Download"
                  disabled={true}
                  styles={{ root: { height: 28, padding: '0 10px' }, label: { fontSize: 11 } }}
                />
              </Stack>
            </div>
            <div
              className={styles.documentViewer}
              ref={this.documentViewerRef}
              onScroll={this.handleDocumentScroll}
            >
              <div className={styles.scrollProgressBar}>
                <div className={styles.scrollProgressFill} style={{ height: `${scrollProgress}%` }} />
              </div>
              <div style={{ padding: '40px 60px', fontFamily: "'Times New Roman', serif", color: '#1e293b', lineHeight: 1.8 }}>
                <div style={{ textAlign: 'center', marginBottom: 32 }}>
                  <div style={{ fontSize: 11, color: '#64748b', textTransform: 'uppercase', letterSpacing: 2, marginBottom: 8 }}>OFFICIAL POLICY DOCUMENT</div>
                  <h1 style={{ fontSize: 24, fontWeight: 700, color: '#0f172a', margin: '0 0 8px' }}>{policy.PolicyName || policy.Title}</h1>
                  <div style={{ fontSize: 13, color: '#64748b' }}>
                    Policy Number: {policy.PolicyNumber} | Version {policy.VersionNumber || '1.0'} | Effective: {policy.EffectiveDate ? new Date(policy.EffectiveDate).toLocaleDateString() : 'TBD'}
                  </div>
                  <hr style={{ border: 'none', borderTop: '2px solid #0d9488', width: 80, margin: '16px auto' }} />
                </div>

                <h2 style={{ fontSize: 18, color: '#0f766e', marginBottom: 8 }}>1. Purpose</h2>
                <p style={{ marginBottom: 16 }}>
                  {policy.PolicySummary || `This policy establishes the guidelines and requirements for ${(policy.PolicyName || policy.Title).toLowerCase()} across the organisation. All employees are expected to read, understand, and comply with the provisions outlined in this document.`}
                </p>

                <h2 style={{ fontSize: 18, color: '#0f766e', marginBottom: 8 }}>2. Scope</h2>
                <p style={{ marginBottom: 16 }}>
                  This policy applies to all employees, contractors, and third-party personnel within the {policy.Department || policy.PolicyCategory || 'organisation'} department and any associated business units. Compliance is mandatory from the effective date.
                </p>

                <h2 style={{ fontSize: 18, color: '#0f766e', marginBottom: 8 }}>3. Policy Statement</h2>
                <p style={{ marginBottom: 16 }}>
                  The organisation is committed to maintaining the highest standards in {(policy.PolicyCategory || 'operations').toLowerCase()}. All staff members must adhere to the procedures and controls described herein to ensure regulatory compliance and operational excellence.
                </p>

                <h2 style={{ fontSize: 18, color: '#0f766e', marginBottom: 8 }}>4. Responsibilities</h2>
                <ul style={{ marginBottom: 16, paddingLeft: 24 }}>
                  <li style={{ marginBottom: 6 }}><strong>Management:</strong> Ensure policy dissemination and monitor compliance within their teams.</li>
                  <li style={{ marginBottom: 6 }}><strong>Employees:</strong> Read, acknowledge, and follow all policy requirements.</li>
                  <li style={{ marginBottom: 6 }}><strong>Policy Owner:</strong> Review and update the policy according to the scheduled review cycle.</li>
                  <li style={{ marginBottom: 6 }}><strong>Compliance Team:</strong> Conduct audits and report on adherence metrics.</li>
                </ul>

                <h2 style={{ fontSize: 18, color: '#0f766e', marginBottom: 8 }}>5. Compliance & Enforcement</h2>
                <p style={{ marginBottom: 16 }}>
                  Non-compliance with this policy may result in disciplinary action up to and including termination of employment. All violations must be reported to the compliance team within 48 hours of discovery.
                </p>

                <h2 style={{ fontSize: 18, color: '#0f766e', marginBottom: 8 }}>6. Review Schedule</h2>
                <p style={{ marginBottom: 16 }}>
                  This policy will be reviewed {policy.ReviewFrequency || 'annually'} or sooner if regulatory changes require updates. The next scheduled review is {policy.NextReviewDate ? new Date(policy.NextReviewDate).toLocaleDateString() : 'as per the annual review cycle'}.
                </p>

                <div style={{ marginTop: 32, padding: 16, background: '#f0fdfa', borderRadius: 6, borderLeft: '3px solid #0d9488' }}>
                  <div style={{ fontSize: 12, color: '#0f766e', fontWeight: 600, fontFamily: 'system-ui, sans-serif' }}>
                    Document Classification: Internal | Owner: {policy.PolicyOwner || 'Policy Administrator'} | Category: {policy.PolicyCategory || 'General'}
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Attachments — hidden in focused mode */}
        {!isFocusedMode && attachments.length > 0 && (
          <div className={styles.wizardCard}>
            <Text variant="medium" style={{ fontWeight: 600, marginBottom: 8, display: 'block' }}>
              Attachments ({attachments.length})
            </Text>
            <Stack tokens={{ childrenGap: 6 }}>
              {attachments.map((url: string, index: number) => {
                const fileName = url.split('/').pop() || `Attachment ${index + 1}`;
                return (
                  <Stack key={index} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} className={styles.attachmentRow}>
                    <Icon iconName={this.getDocumentIcon(url)} style={{ fontSize: 16, color: '#0d9488' }} />
                    <a href={url} target="_blank" rel="noopener noreferrer" style={{ color: '#0d9488', textDecoration: 'none', flex: 1 }}>{fileName}</a>
                    <Text variant="tiny" style={{ color: '#94a3b8' }}>{this.getDocumentTypeLabel(url)}</Text>
                  </Stack>
                );
              })}
            </Stack>
          </div>
        )}

        {/* Document Read Footer — anchored to bottom of browser, covers PM footer */}
        {!this.state.isFullscreen && !isFocusedMode && (hasConvertedHtml || hasDocuments) && (
          <div style={{
            position: 'fixed',
            bottom: 0,
            left: 0,
            right: 0,
            padding: '8px 24px',
            background: scrollProgress >= 95 || readDuration >= 30 ? '#dcfce7' : '#fef3c7',
            borderTop: `2px solid ${scrollProgress >= 95 || readDuration >= 30 ? '#16a34a' : '#d97706'}`,
            zIndex: 1000,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            transition: 'background 0.3s, border-color 0.3s'
          }}>
            {/* Back button */}
            <DefaultButton
              text="Back"
              iconProps={{ iconName: 'ChevronLeft' }}
              onClick={() => { window.location.href = '/sites/PolicyManager/SitePages/MyPolicies.aspx'; }}
              styles={{ root: { borderRadius: 4, minWidth: 80 } }}
            />

            {/* Status text */}
            <span style={{ fontSize: 13, textAlign: 'center', flex: 1 }}>
              {scrollProgress >= 95 || readDuration >= 30 ? (
                <span style={{ color: '#16a34a', fontWeight: 600 }}>
                  <Icon iconName="CheckMark" /> Document read complete — you may now proceed
                </span>
              ) : (
                <span style={{ color: '#92400e' }}>
                  Please review the document before proceeding ({Math.max(0, 30 - readDuration)}s remaining)
                </span>
              )}
            </span>

            {/* Next button */}
            <PrimaryButton
              text={scrollProgress >= 95 || readDuration >= 30 ? (policy.RequiresQuiz ? 'Proceed to Quiz' : 'Proceed to Acknowledge') : `${Math.max(0, 30 - readDuration)}s`}
              iconProps={{ iconName: 'ChevronRight' }}
              disabled={scrollProgress < 95 && readDuration < 30}
              onClick={() => this.handleMarkAsRead()}
              styles={{
                root: { borderRadius: 4, minWidth: 80, background: scrollProgress >= 95 || readDuration >= 30 ? '#0d9488' : '#94a3b8', borderColor: scrollProgress >= 95 || readDuration >= 30 ? '#0d9488' : '#94a3b8' },
                rootHovered: { background: '#0f766e', borderColor: '#0f766e' }
              }}
            />
          </div>
        )}
      </div>
    );
  }

  /**
   * Renders the per-policy documents section (files from PM_PolicySourceDocuments/{PolicyNumber}/)
   */
  private renderPolicyDocumentsSection(policy: IPolicy): JSX.Element {
    const state = this.state as any;
    const policyDocs = state._policyDocuments || [];
    const docsLoading = state._policyDocsLoading || false;
    const docsLoaded = state._policyDocsLoaded || false;
    const docsExpanded = state._policyDocsExpanded || false;

    // Lazy-load documents on first expand
    const handleToggle = async (): Promise<void> => {
      const newExpanded = !docsExpanded;
      this.setState({ _policyDocsExpanded: newExpanded } as any);
      if (newExpanded && !docsLoaded) {
        this.setState({ _policyDocsLoading: true } as any);
        try {
          const docs = await this.policyService.getPolicyDocuments(policy.PolicyNumber);
          this.setState({ _policyDocuments: docs, _policyDocsLoaded: true, _policyDocsLoading: false } as any);
        } catch {
          this.setState({ _policyDocsLoaded: true, _policyDocsLoading: false } as any);
        }
      }
    };

    return (
      <div style={{ marginTop: 16, border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden' }}>
        <Stack
          horizontal
          verticalAlign="center"
          tokens={{ childrenGap: 8 }}
          onClick={handleToggle}
          style={{ padding: '12px 16px', cursor: 'pointer', backgroundColor: '#f8fafc' }}
        >
          <Icon iconName={docsExpanded ? 'ChevronDown' : 'ChevronRight'} style={{ fontSize: 12 }} />
          <Icon iconName="FolderOpen" style={{ fontSize: 16, color: '#0d9488' }} />
          <Text style={{ fontWeight: 600, flex: 1 }}>Policy Documents</Text>
          {docsLoaded && <Text style={{ color: '#94a3b8', fontSize: 12 }}>{policyDocs.length} file{policyDocs.length !== 1 ? 's' : ''}</Text>}
        </Stack>
        {docsExpanded && (
          <div style={{ padding: '0 16px 12px' }}>
            {docsLoading ? (
              <Spinner size={SpinnerSize.small} label="Loading documents..." />
            ) : policyDocs.length === 0 ? (
              <Text style={{ color: '#605e5c', fontSize: 13 }}>No documents uploaded for this policy.</Text>
            ) : (
              <Stack tokens={{ childrenGap: 6 }}>
                {policyDocs.map((doc: any, i: number) => (
                  <Stack key={i} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ padding: '4px 0' }}>
                    <Icon iconName={this.getDocumentIcon(doc.Name)} style={{ fontSize: 14, color: '#0d9488' }} />
                    <a
                      href={doc.ServerRelativeUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      style={{ color: '#0d9488', textDecoration: 'none', fontSize: 13, flex: 1 }}
                    >
                      {doc.Name}
                    </a>
                    <Text style={{ color: '#94a3b8', fontSize: 11 }}>
                      {doc.Length ? `${Math.round(doc.Length / 1024)} KB` : ''}
                    </Text>
                  </Stack>
                ))}
              </Stack>
            )}
          </div>
        )}
      </div>
    );
  }

  // ============================================
  // RENDER: STEP 2 — QUIZ (HORIZONTAL)
  // ============================================

  private renderQuizStep(): JSX.Element | null {
    const { policy, currentQuizQuestion, quizAnswers, quizSubmitted, quizScore, quizPassed, liveQuizId, currentUserId } = this.state;
    if (!policy) return null;

    // Use live QuizTaker if a real quiz exists for this policy
    if (liveQuizId && currentUserId) {
      return (
        <div className={styles.stepContent}>
          <div className={styles.quizBanner}>
            <Icon iconName="Questionnaire" styles={{ root: { fontSize: 18 } }} />
            <span>This policy requires a comprehension quiz. You must score at least <strong>{policy.QuizPassingScore || 80}%</strong> to proceed.</span>
          </div>
          <QuizTaker
            sp={this.props.sp}
            quizId={liveQuizId}
            policyId={policy.Id || 0}
            userId={{ Id: currentUserId, Title: '', EMail: '' }}
            onComplete={(result: IQuizResult) => {
              this.handleQuizComplete(result.percentage, result.passed);
            }}
            onCancel={() => {
              // Go back to reading step
              this.setState({ currentFlowStep: 'reading' });
            }}
          />
        </div>
      );
    }

    // Fallback: mock quiz (for policies without a live quiz)
    const questions = MOCK_QUIZ_QUESTIONS;
    const allConfirmed = this.state.quizConfirmed.every(c => c);
    const qi = currentQuizQuestion;
    const q = questions[qi];
    const isConfirmed = this.state.quizConfirmed[qi];
    const attempts = this.state.quizAttempts[qi] || 0;
    const isCorrect = isConfirmed && quizAnswers[qi] === q.correctIndex;
    const isIncorrect = isConfirmed && quizAnswers[qi] !== q.correctIndex;
    const isFirstWrongAttempt = !isConfirmed && attempts === 1; // Tried once, got it wrong, can try again
    const passingScore = policy.QuizPassingScore || 80;

    // Explanations per question (for feedback)
    const explanations = [
      'The policy requires a minimum of 12 characters for all passwords to ensure strong security.',
      'Passwords must be changed every 90 days as per the security rotation schedule.',
      'All security incidents must be reported directly to the Security Team for immediate triage.',
      'Sharing login credentials is never permitted, regardless of circumstances. Each user must use their own credentials.',
      'Suspected phishing emails should be forwarded to the Security Team for analysis. Never click links in suspicious emails.'
    ];

    return (
      <div className={styles.stepContent}>
        {/* Quiz Banner */}
        <div className={styles.quizBanner}>
          <Icon iconName="Questionnaire" styles={{ root: { fontSize: 18 } }} />
          <span>This policy requires a comprehension quiz. You must score at least <strong>{passingScore}%</strong> to proceed.</span>
        </div>

        {!quizSubmitted ? (
          <div style={{ maxWidth: 680, margin: '0 auto' }}>
            {/* Progress bar */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 16, marginBottom: 20 }}>
              <div style={{ flex: 1, height: 6, background: '#e2e8f0', borderRadius: 3, overflow: 'hidden' }}>
                <div style={{ height: '100%', background: '#0d9488', borderRadius: 3, width: `${((qi + 1) / questions.length) * 100}%`, transition: 'width 0.5s' }} />
              </div>
              <span style={{ fontSize: 12, color: '#64748b', whiteSpace: 'nowrap' }}>{qi + 1} / {questions.length}</span>
            </div>

            {/* Question dots — green ✓ / red ✗ / current / unanswered */}
            <div style={{ display: 'flex', justifyContent: 'center', gap: 8, marginBottom: 24 }}>
              {questions.map((_, i) => {
                const confirmed = this.state.quizConfirmed[i];
                const att = this.state.quizAttempts[i] || 0;
                const correct = confirmed && quizAnswers[i] === questions[i].correctIndex;
                const incorrect = confirmed && quizAnswers[i] !== questions[i].correctIndex;
                const tryingAgain = !confirmed && att === 1; // Amber — had one wrong, trying again
                const isCurrent = i === qi && !quizSubmitted;
                return (
                  <div
                    key={i}
                    role="button" tabIndex={0}
                    onClick={() => this.setState({ currentQuizQuestion: i })}
                    onKeyDown={(e) => { if (e.key === 'Enter') this.setState({ currentQuizQuestion: i }); }}
                    style={{
                      width: 32, height: 32, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center',
                      fontSize: 12, fontWeight: 600, cursor: 'pointer', transition: 'all 0.2s',
                      background: correct ? '#059669' : incorrect ? '#dc2626' : tryingAgain ? '#d97706' : isCurrent ? '#f0fdfa' : '#fff',
                      border: `2px solid ${correct ? '#059669' : incorrect ? '#dc2626' : tryingAgain ? '#d97706' : isCurrent ? '#0d9488' : '#e2e8f0'}`,
                      color: correct || incorrect || tryingAgain ? '#fff' : isCurrent ? '#0d9488' : '#94a3b8',
                    }}
                  >
                    {correct ? '✓' : incorrect ? '✗' : tryingAgain ? '!' : i + 1}
                  </div>
                );
              })}
            </div>

            {/* Feedback banner (shown after confirm) */}
            {isCorrect && (
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '12px 16px', borderRadius: 8, marginBottom: 20, background: '#f0fdf4', border: '1px solid #bbf7d0', color: '#166534', fontSize: 13 }}>
                <span style={{ fontSize: 18 }}>✓</span>
                <div><strong style={{ display: 'block', marginBottom: 2 }}>Correct!</strong><span style={{ fontSize: 12, opacity: 0.8 }}>{explanations[qi]}</span></div>
              </div>
            )}
            {isFirstWrongAttempt && (
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '12px 16px', borderRadius: 8, marginBottom: 20, background: '#fef3c7', border: '1px solid #fbbf24', color: '#92400e', fontSize: 13 }}>
                <span style={{ fontSize: 18 }}>⚠</span>
                <div><strong style={{ display: 'block', marginBottom: 2 }}>Not quite — try again!</strong><span style={{ fontSize: 12, opacity: 0.8 }}>You have one more attempt. Read the question carefully and select a different answer.</span></div>
              </div>
            )}
            {isIncorrect && (
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '12px 16px', borderRadius: 8, marginBottom: 20, background: '#fef2f2', border: '1px solid #fecaca', color: '#991b1b', fontSize: 13 }}>
                <span style={{ fontSize: 18 }}>✗</span>
                <div><strong style={{ display: 'block', marginBottom: 2 }}>Incorrect — the correct answer is shown below</strong><span style={{ fontSize: 12, opacity: 0.8 }}>{explanations[qi]}</span></div>
              </div>
            )}

            {/* Question card */}
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 32, marginBottom: 20 }}>
              <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', marginBottom: 8 }}>Question {qi + 1} of {questions.length}</div>
              <div style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', marginBottom: 28, lineHeight: 1.4 }}>{q.question}</div>
              {q.options.map((opt, oi) => {
                const isSelected = quizAnswers[qi] === oi;
                const isCorrectOption = isConfirmed && oi === q.correctIndex;
                const isWrongSelection = isConfirmed && isSelected && oi !== q.correctIndex;
                return (
                  <div
                    key={oi}
                    role="button" tabIndex={0}
                    onClick={() => this.handleQuizSelectAnswer(qi, oi)}
                    onKeyDown={(e) => { if (e.key === 'Enter') this.handleQuizSelectAnswer(qi, oi); }}
                    style={{
                      display: 'flex', alignItems: 'center', gap: 14, padding: '14px 18px', marginBottom: 8,
                      border: `2px solid ${isWrongSelection ? '#dc2626' : isCorrectOption ? '#059669' : isSelected ? '#0d9488' : '#e2e8f0'}`,
                      borderStyle: isCorrectOption && !isSelected ? 'dashed' : 'solid',
                      borderRadius: 8, cursor: isConfirmed ? 'default' : 'pointer', transition: 'all 0.2s', fontSize: 14,
                      background: isWrongSelection ? '#fef2f2' : isCorrectOption ? '#f0fdf4' : isSelected ? '#f0fdfa' : '#fff',
                    }}
                  >
                    <div style={{
                      width: 22, height: 22, borderRadius: '50%', flexShrink: 0,
                      display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700,
                      border: `2px solid ${isWrongSelection ? '#dc2626' : isCorrectOption ? '#059669' : isSelected ? '#0d9488' : '#cbd5e1'}`,
                      background: isWrongSelection ? '#dc2626' : isCorrectOption ? '#059669' : isSelected ? '#0d9488' : 'transparent',
                      color: (isWrongSelection || isCorrectOption || isSelected) ? '#fff' : 'transparent',
                    }}>
                      {isWrongSelection ? '✗' : isCorrectOption ? '✓' : isSelected ? '•' : ''}
                    </div>
                    <span>{opt}</span>
                  </div>
                );
              })}
            </div>

            {/* Navigation */}
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <DefaultButton
                text="← Previous"
                disabled={qi === 0}
                onClick={() => this.setState({ currentQuizQuestion: qi - 1 })}
                styles={{ root: { borderRadius: 6 } }}
              />
              <span style={{ fontSize: 13, color: '#64748b' }}>Question {qi + 1} of {questions.length}</span>
              {!isConfirmed ? (
                <PrimaryButton
                  text={isFirstWrongAttempt ? 'Confirm 2nd Attempt →' : 'Confirm Answer →'}
                  disabled={quizAnswers[qi] < 0}
                  onClick={this.handleQuizConfirmAnswer}
                  styles={{ root: { borderRadius: 6, background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e' } }}
                />
              ) : qi < questions.length - 1 ? (
                <PrimaryButton
                  text="Next Question →"
                  onClick={this.handleQuizNextAfterConfirm}
                  styles={{ root: { borderRadius: 6, background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e' } }}
                />
              ) : allConfirmed ? (
                <PrimaryButton
                  text="View Results"
                  iconProps={{ iconName: 'Accept' }}
                  onClick={this.handleQuizSubmit}
                  styles={{ root: { borderRadius: 6, background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e' } }}
                />
              ) : (
                <PrimaryButton text="Answer remaining questions" disabled styles={{ root: { borderRadius: 6 } }} />
              )}
            </div>
          </div>
        ) : (
          /* Results screen with answer review */
          <div style={{ maxWidth: 680, margin: '0 auto' }}>
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 32, textAlign: 'center', marginBottom: 24 }}>
              <div style={{ fontSize: 48, fontWeight: 700, color: quizPassed ? '#059669' : '#dc2626' }}>{quizScore}%</div>
              <div style={{ fontSize: 14, color: '#64748b', marginBottom: 16 }}>
                {Math.round(quizScore / (100 / questions.length))}/{questions.length} correct
              </div>
              <div style={{
                display: 'inline-block', padding: '6px 16px', borderRadius: 20, fontSize: 13, fontWeight: 700,
                background: quizPassed ? '#dcfce7' : '#fee2e2', color: quizPassed ? '#166534' : '#991b1b'
              }}>
                {quizPassed ? '✓ Passed — You may proceed to acknowledgement' : `✗ Not passed — ${passingScore}% required`}
              </div>
              <div style={{ display: 'flex', justifyContent: 'center', gap: 32, marginTop: 20 }}>
                <div style={{ textAlign: 'center' }}><div style={{ fontSize: 20, fontWeight: 700, color: '#059669' }}>{questions.filter((qq, i) => quizAnswers[i] === qq.correctIndex).length}</div><div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', marginTop: 2 }}>Correct</div></div>
                <div style={{ textAlign: 'center' }}><div style={{ fontSize: 20, fontWeight: 700, color: '#dc2626' }}>{questions.filter((qq, i) => quizAnswers[i] !== qq.correctIndex).length}</div><div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', marginTop: 2 }}>Incorrect</div></div>
              </div>
            </div>

            {/* Answer review */}
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 24, marginBottom: 24 }}>
              <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', marginBottom: 16 }}>Answer Review</div>
              {questions.map((qq, i) => {
                const correct = quizAnswers[i] === qq.correctIndex;
                return (
                  <div key={i} style={{ display: 'flex', alignItems: 'flex-start', gap: 12, padding: '14px 0', borderBottom: i < questions.length - 1 ? '1px solid #f1f5f9' : 'none' }}>
                    <div style={{
                      width: 24, height: 24, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center',
                      fontSize: 11, fontWeight: 700, flexShrink: 0, marginTop: 2,
                      background: correct ? '#dcfce7' : '#fee2e2', color: correct ? '#059669' : '#dc2626'
                    }}>{correct ? '✓' : '✗'}</div>
                    <div>
                      <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a', marginBottom: 4 }}>Q{i + 1}: {qq.question}</div>
                      {correct ? (
                        <div style={{ fontSize: 12, color: '#059669', fontWeight: 600 }}>Your answer: {qq.options[quizAnswers[i]]} ✓</div>
                      ) : (
                        <>
                          <div style={{ fontSize: 12, color: '#dc2626', textDecoration: 'line-through' }}>Your answer: {qq.options[quizAnswers[i]] || 'Not answered'}</div>
                          <div style={{ fontSize: 12, color: '#059669', fontWeight: 600 }}>Correct: {qq.options[qq.correctIndex]}</div>
                        </>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Actions */}
            <div style={{ display: 'flex', justifyContent: 'center', gap: 12 }}>
              {quizPassed ? (
                <PrimaryButton
                  text="Proceed to Acknowledgement →"
                  onClick={() => this.setState({ currentFlowStep: 'acknowledge', showAcknowledgePanel: true })}
                  styles={{ root: { borderRadius: 6, background: '#0d9488', borderColor: '#0d9488', padding: '12px 32px' }, rootHovered: { background: '#0f766e' } }}
                />
              ) : (
                <DefaultButton
                  text="Retake Quiz"
                  iconProps={{ iconName: 'Refresh' }}
                  onClick={this.handleQuizRetake}
                  styles={{ root: { borderRadius: 6 } }}
                />
              )}
            </div>
          </div>
        )}
      </div>
    );
  }

  // ============================================
  // RENDER: STEP 3 — ACKNOWLEDGE (Inline prompt + Panel)
  // ============================================

  private renderAcknowledgeStep(): JSX.Element {
    return (
      <div className={styles.stepContent}>
        <div className={styles.wizardCard} style={{ textAlign: 'center', padding: 48 }}>
          <Icon iconName="Lock" styles={{ root: { fontSize: 48, color: '#0d9488', marginBottom: 12 } }} />
          <Text variant="xLarge" style={{ fontWeight: 700, display: 'block', marginBottom: 8 }}>Policy Acknowledgement</Text>
          <Text style={{ color: '#64748b', display: 'block', marginBottom: 24 }}>
            The acknowledgement panel will open from the right. Please review the policy details and complete your legal acknowledgement.
          </Text>
          <PrimaryButton
            text="Open Acknowledgement Panel"
            iconProps={{ iconName: 'Handwriting' }}
            onClick={this.handleOpenAcknowledgePanel}
          />
        </div>
      </div>
    );
  }

  // ============================================
  // RENDER: STEP 4 — COMPLETE
  // ============================================

  private renderCompleteStep(): JSX.Element {
    const { policy, readReceipt, readDuration, quizScore, quizRequired, emailingReceipt, generatingPdf, userRating, reviewTitle, reviewText, submittingRating } = this.state;

    return (
      <div className={styles.stepContent}>
        <div className={styles.successContainer}>
          <div className={styles.trophyIcon}>&#127942;</div>
          <Text variant="xxLarge" style={{ fontWeight: 800, color: '#0f766e', display: 'block', marginBottom: 8 }}>
            Congratulations!
          </Text>
          <Text variant="mediumPlus" style={{ color: '#64748b', display: 'block', marginBottom: 28 }}>
            You have successfully completed the policy acknowledgement process.
          </Text>

          {/* Certificate Card */}
          <div className={styles.certificateCard}>
            <div className={styles.certificateTitle}>Certificate of Compliance</div>
            <div className={styles.certificateMainTitle}>{policy?.PolicyName}</div>
            <div className={styles.certificateGrid}>
              <div className={styles.certItem}>
                <div className={styles.certLabel}>Employee</div>
                <div className={styles.certValue}>{readReceipt?.UserDisplayName || 'Current User'}</div>
              </div>
              <div className={styles.certItem}>
                <div className={styles.certLabel}>Date Completed</div>
                <div className={styles.certValue}>{readReceipt?.AcknowledgedDate?.toLocaleDateString() || new Date().toLocaleDateString()}</div>
              </div>
              <div className={styles.certItem}>
                <div className={styles.certLabel}>Receipt Number</div>
                <div className={styles.certValue}>{readReceipt?.ReceiptNumber || 'N/A'}</div>
              </div>
              <div className={styles.certItem}>
                <div className={styles.certLabel}>Read Duration</div>
                <div className={styles.certValue}>{this.formatDuration(readDuration)}</div>
              </div>
              <div className={styles.certItem}>
                <div className={styles.certLabel}>Quiz Score</div>
                <div className={styles.certValue}>{quizRequired ? `${quizScore}%` : 'N/A (no quiz)'}</div>
              </div>
              <div className={styles.certItem}>
                <div className={styles.certLabel}>Policy Number</div>
                <div className={styles.certValue}>{policy?.PolicyNumber}</div>
              </div>
            </div>
            {readReceipt?.DigitalSignature && (
              <>
                <div className={styles.certSignature}>{readReceipt.DigitalSignature}</div>
                <div className={styles.certSigLabel}>Digital Signature</div>
              </>
            )}
          </div>

          {/* Action Buttons */}
          <Stack horizontal tokens={{ childrenGap: 10 }} wrap horizontalAlign="center" style={{ marginBottom: 16 }}>
            <DefaultButton
              text={emailingReceipt ? 'Sending...' : 'Email Receipt'}
              iconProps={{ iconName: 'Mail' }}
              onClick={this.handleEmailReceipt}
              disabled={emailingReceipt}
            />
            <DefaultButton
              text={generatingPdf ? 'Generating...' : 'Download PDF'}
              iconProps={{ iconName: 'PDF' }}
              onClick={this.handleGeneratePdf}
              disabled={generatingPdf}
            />
            <DefaultButton
              text="View Receipt"
              iconProps={{ iconName: 'View' }}
              onClick={this.handleViewReceipt}
            />
          </Stack>

          {/* Rating Section */}
          {this.props.showRatings && (
            <div className={styles.ratingSection}>
              <Text variant="medium" style={{ fontWeight: 600, marginBottom: 8, display: 'block' }}>Rate this policy</Text>
              <Rating
                rating={userRating}
                size={RatingSize.Large}
                onChange={(ev, rating) => this.handleRate(rating || 0)}
              />
              {userRating > 0 && (
                <Stack tokens={{ childrenGap: 8 }} style={{ marginTop: 12 }}>
                  <TextField
                    label="Review Title (Optional)"
                    value={reviewTitle}
                    onChange={(e, value) => this.setState({ reviewTitle: value || '' })}
                  />
                  <TextField
                    label="Review (Optional)"
                    multiline rows={3}
                    value={reviewText}
                    onChange={(e, value) => this.setState({ reviewText: value || '' })}
                  />
                  <DefaultButton text="Submit Rating" onClick={this.handleSubmitRating} disabled={submittingRating} />
                </Stack>
              )}
            </div>
          )}

          {/* Return Button */}
          <div style={{ marginTop: 24 }}>
            <PrimaryButton
              text="Return to My Policies"
              iconProps={{ iconName: 'Back' }}
              onClick={() => window.location.href = '/sites/PolicyManager/SitePages/MyPolicies.aspx'}
            />
          </div>
        </div>
      </div>
    );
  }

  // ============================================
  // VERSION HISTORY
  // ============================================

  private loadVersionHistory = async (): Promise<void> => {
    this.setState({ showVersionHistoryPanel: true, versionHistoryLoading: true });
    try {
      const { policyId } = this.state;
      if (!policyId) return;
      const versions = await this.policyService.getPolicyVersions(policyId);
      this.setState({ policyVersions: versions, versionHistoryLoading: false });
    } catch (error) {
      console.error('Failed to load version history:', error);
      this.setState({ versionHistoryLoading: false });
    }
  }

  private handleCompareWithCurrent = async (versionId: number): Promise<void> => {
    this.setState({ showVersionComparisonPanel: true, versionComparisonLoading: true });
    try {
      const { policyId } = this.state;
      if (!policyId) return;
      const comparison = await this.comparisonService.compareWithVersion(policyId, versionId);
      const sideBySide = await this.comparisonService.getSideBySideView(comparison.sourceVersion?.Id || versionId, comparison.targetVersion?.Id || 0);
      const html = this.comparisonService.generateSideBySideHtml(sideBySide);
      this.setState({ versionComparisonHtml: html, versionComparisonLoading: false });
    } catch (error) {
      console.error('Failed to compare versions:', error);
      // Fallback: show a simple message
      this.setState({
        versionComparisonHtml: '<div style="padding: 24px; color: #605e5c;">Version comparison data is not available for these versions. Ensure both versions have HTML content saved.</div>',
        versionComparisonLoading: false
      });
    }
  }

  private renderVersionHistoryPanel(): JSX.Element {
    const { showVersionHistoryPanel, versionHistoryLoading, policyVersions, policy } = this.state;

    return (
      <StyledPanel
        isOpen={showVersionHistoryPanel}
        onDismiss={() => this.setState({ showVersionHistoryPanel: false })}
        type={PanelType.medium}
        headerText="Version History"
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 16 }} style={{ padding: '16px 0' }}>
          {versionHistoryLoading ? (
            <Spinner size={SpinnerSize.large} label="Loading version history..." />
          ) : policyVersions.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No previous versions found for this policy.
            </MessageBar>
          ) : (
            policyVersions.map((version, index) => (
              <div
                key={version.Id || index}
                style={{
                  padding: 16,
                  border: '1px solid #e2e8f0',
                  borderRadius: 8,
                  backgroundColor: version.IsCurrentVersion ? '#f0fdfa' : '#ffffff'
                }}
              >
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                  <Stack tokens={{ childrenGap: 4 }}>
                    <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                      <Text style={{ fontWeight: 600, fontSize: 16, color: '#0f172a' }}>
                        v{version.VersionNumber}
                      </Text>
                      <span style={{
                        padding: '2px 8px',
                        borderRadius: 4,
                        fontSize: 11,
                        fontWeight: 600,
                        backgroundColor: version.VersionType === 'Major' ? '#dcfce7' : '#e0f2fe',
                        color: version.VersionType === 'Major' ? '#16a34a' : '#0284c7'
                      }}>
                        {version.VersionType}
                      </span>
                      {version.IsCurrentVersion && (
                        <span style={{
                          padding: '2px 8px',
                          borderRadius: 4,
                          fontSize: 11,
                          fontWeight: 600,
                          backgroundColor: '#ccfbf1',
                          color: '#0d9488'
                        }}>
                          Current
                        </span>
                      )}
                    </Stack>
                    <Text style={{ color: '#605e5c', fontSize: 13 }}>
                      {version.ChangeDescription || 'No description'}
                    </Text>
                    <Text style={{ color: '#94a3b8', fontSize: 12 }}>
                      {version.EffectiveDate ? new Date(version.EffectiveDate).toLocaleDateString('en-US', {
                        year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit'
                      }) : 'Unknown date'}
                    </Text>
                  </Stack>
                  {!version.IsCurrentVersion && (
                    <DefaultButton
                      text="Compare with Current"
                      iconProps={{ iconName: 'BranchCompare' }}
                      onClick={() => this.handleCompareWithCurrent(version.Id)}
                      styles={{ root: { fontSize: 12 } }}
                    />
                  )}
                </Stack>
              </div>
            ))
          )}
        </Stack>
      </StyledPanel>
    );
  }

  private renderVersionComparisonPanel(): JSX.Element {
    const { showVersionComparisonPanel, versionComparisonLoading, versionComparisonHtml } = this.state;

    return (
      <StyledPanel
        isOpen={showVersionComparisonPanel}
        onDismiss={() => this.setState({ showVersionComparisonPanel: false, versionComparisonHtml: '' })}
        type={PanelType.extraLarge}
        headerText="Version Comparison"
        closeButtonAriaLabel="Close"
      >
        <div style={{ padding: '16px 0' }}>
          {versionComparisonLoading ? (
            <Spinner size={SpinnerSize.large} label="Generating comparison..." />
          ) : (
            <div
              dangerouslySetInnerHTML={{ __html: sanitizeHtml(versionComparisonHtml || '') }}
              style={{ border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'auto' }}
            />
          )}
        </div>
      </StyledPanel>
    );
  }

  // ============================================
  // RENDER: ACKNOWLEDGE PANEL (fly-in)
  // ============================================

  private renderAcknowledgePanel(): JSX.Element {
    const {
      showAcknowledgePanel, policy, legalAgreement1, legalAgreement2, legalAgreement3,
      digitalSignature, acknowledgeNotes, submittingAcknowledgement, readDuration
    } = this.state;

    return (
      <Panel
        isOpen={showAcknowledgePanel}
        onDismiss={this.handleCloseAcknowledgePanel}
        type={PanelType.medium}
        hasCloseButton={false}
        closeButtonAriaLabel="Close"
        isFooterAtBottom={true}
        onRenderHeader={() => (
          <div style={{
            background: 'linear-gradient(135deg, #0d9488 0%, #0f766e 100%)',
            padding: '16px 24px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            width: '100%'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
              <Icon iconName="Shield" style={{ fontSize: 20, color: '#fff' }} />
              <span style={{ fontSize: 18, fontWeight: 600, color: '#fff' }}>Policy Acknowledgement</span>
            </div>
            <button
              onClick={this.handleCloseAcknowledgePanel}
              style={{ background: 'none', border: 'none', color: 'rgba(255,255,255,0.8)', cursor: 'pointer', fontSize: 16, padding: 4 }}
              aria-label="Close"
            >
              <Icon iconName="ChromeClose" />
            </button>
          </div>
        )}
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 12 }}>
            <PrimaryButton
              text={submittingAcknowledgement ? 'Submitting...' : 'Submit Acknowledgement'}
              iconProps={{ iconName: 'Accept' }}
              onClick={this.handleSubmitAcknowledgement}
              disabled={!this.canSubmitAcknowledgement() || submittingAcknowledgement}
            />
            <DefaultButton
              text="Cancel"
              onClick={this.handleCloseAcknowledgePanel}
              disabled={submittingAcknowledgement}
            />
          </Stack>
        )}
      >
        <Stack tokens={{ childrenGap: 24 }}>
          <div className={styles.legalHeader}>
            <Icon iconName="Shield" style={{ fontSize: 32, color: '#0d9488' }} />
            <Text variant="xLarge" style={{ fontWeight: 600 }}>Acknowledgement Declaration</Text>
          </div>

          <div className={styles.acknowledgePolicyInfo}>
            <Text variant="large" style={{ fontWeight: 600 }}>{policy?.PolicyNumber}</Text>
            <Text variant="large">{policy?.PolicyName}</Text>
            <Text variant="small" style={{ color: '#605e5c' }}>
              Version {policy?.VersionNumber} | Read time: {this.formatDuration(readDuration)}
            </Text>
          </div>

          <Separator />

          <div className={styles.legalAgreements}>
            <Text variant="medium" style={{ fontWeight: 600, marginBottom: 16, display: 'block' }}>
              Please read and confirm each statement below:
            </Text>
            <Stack tokens={{ childrenGap: 16 }}>
              <Checkbox
                label="I confirm that I have read the entire policy document and understand its contents, requirements, and implications."
                checked={legalAgreement1}
                onChange={(e, checked) => this.setState({ legalAgreement1: checked || false })}
                styles={{ root: { alignItems: 'flex-start' }, checkbox: { marginTop: 2 } }}
              />
              <Checkbox
                label="I agree to comply with all requirements, procedures, and guidelines outlined in this policy. I understand that violations may result in disciplinary action, up to and including termination of employment."
                checked={legalAgreement2}
                onChange={(e, checked) => this.setState({ legalAgreement2: checked || false })}
                styles={{ root: { alignItems: 'flex-start' }, checkbox: { marginTop: 2 } }}
              />
              <Checkbox
                label="I understand that this electronic acknowledgement serves as my official record of having read and accepted this policy, and that it will be retained as part of my employment record for audit and compliance purposes."
                checked={legalAgreement3}
                onChange={(e, checked) => this.setState({ legalAgreement3: checked || false })}
                styles={{ root: { alignItems: 'flex-start' }, checkbox: { marginTop: 2 } }}
              />
            </Stack>
          </div>

          <Separator />

          <div className={styles.digitalSignature}>
            <Label required>Digital Signature (Type your full legal name)</Label>
            <TextField
              value={digitalSignature}
              onChange={(e, value) => this.setState({ digitalSignature: value || '' })}
              placeholder="e.g., John Smith"
              styles={{
                field: { fontFamily: "'Brush Script MT', cursive", fontSize: 18, fontStyle: 'italic' }
              }}
            />
            <Text variant="small" style={{ color: '#605e5c', marginTop: 4 }}>
              By typing your name above, you are providing your electronic signature.
            </Text>
          </div>

          <TextField
            label="Additional Notes (Optional)"
            multiline rows={3}
            value={acknowledgeNotes}
            onChange={(e, value) => this.setState({ acknowledgeNotes: value || '' })}
            placeholder="Add any comments or clarifications..."
          />

          <MessageBar messageBarType={MessageBarType.warning}>
            <strong>Important:</strong> Your acknowledgement will be recorded with a timestamp and stored for audit purposes.
            Ensure you have thoroughly read and understood this policy before proceeding.
          </MessageBar>
        </Stack>
      </Panel>
    );
  }

  // ============================================
  // RENDER: READ RECEIPT PANEL
  // ============================================

  private renderReadReceiptPanel(): JSX.Element {
    const { showReadReceiptPanel, readReceipt } = this.state;

    return (
      <StyledPanel
        isOpen={showReadReceiptPanel}
        onDismiss={() => this.setState({ showReadReceiptPanel: false })}
        type={PanelType.medium}
        headerText="Read Receipt Details"
        closeButtonAriaLabel="Close"
      >
        {readReceipt && (
          <Stack tokens={{ childrenGap: 16 }}>
            <div className={styles.receiptHeader}>
              <Icon iconName="DocumentApproval" style={{ fontSize: 32, color: '#16a34a' }} />
              <Text variant="xLarge" style={{ fontWeight: 600 }}>Policy Acknowledgement Receipt</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>{readReceipt.ReceiptNumber}</Text>
            </div>
            <Separator />
            <Stack tokens={{ childrenGap: 12 }}>
              <div className={styles.receiptRow}><Text style={{ fontWeight: 600, minWidth: 140 }}>Employee:</Text><Text>{readReceipt.UserDisplayName}</Text></div>
              <div className={styles.receiptRow}><Text style={{ fontWeight: 600, minWidth: 140 }}>Email:</Text><Text>{readReceipt.UserEmail}</Text></div>
              <div className={styles.receiptRow}><Text style={{ fontWeight: 600, minWidth: 140 }}>Policy:</Text><Text>{readReceipt.PolicyNumber} - {readReceipt.PolicyName}</Text></div>
              <div className={styles.receiptRow}><Text style={{ fontWeight: 600, minWidth: 140 }}>Version:</Text><Text>{readReceipt.PolicyVersion}</Text></div>
              <div className={styles.receiptRow}><Text style={{ fontWeight: 600, minWidth: 140 }}>Read Duration:</Text><Text>{this.formatDuration(readReceipt.ReadDurationSeconds)}</Text></div>
              {readReceipt.QuizRequired && (
                <div className={styles.receiptRow}><Text style={{ fontWeight: 600, minWidth: 140 }}>Quiz Score:</Text><Text>{readReceipt.QuizScore}%</Text></div>
              )}
              <div className={styles.receiptRow}><Text style={{ fontWeight: 600, minWidth: 140 }}>Acknowledged:</Text><Text>{readReceipt.AcknowledgedDate?.toLocaleString()}</Text></div>
              <div className={styles.receiptRow}><Text style={{ fontWeight: 600, minWidth: 140 }}>Digital Signature:</Text><Text style={{ fontStyle: 'italic' }}>{readReceipt.DigitalSignature}</Text></div>
              <div className={styles.receiptRow}><Text style={{ fontWeight: 600, minWidth: 140 }}>Device:</Text><Text>{readReceipt.DeviceType} - {readReceipt.BrowserName}</Text></div>
            </Stack>
            <Separator />
            <div className={styles.legalConfirmation}>
              <Text variant="small" style={{ fontWeight: 600 }}>Legal Confirmation:</Text>
              <Text variant="small" style={{ whiteSpace: 'pre-line', marginTop: 8 }}>{readReceipt.LegalConfirmationText}</Text>
            </div>
            <Stack horizontal tokens={{ childrenGap: 12 }}>
              <DefaultButton text="Email Copy" iconProps={{ iconName: 'Mail' }} onClick={this.handleEmailReceipt} />
              <DefaultButton text="Print/PDF" iconProps={{ iconName: 'PDF' }} onClick={this.handleGeneratePdf} />
            </Stack>
          </Stack>
        )}
      </StyledPanel>
    );
  }

  // ============================================
  // RENDER: STICKY FOOTER
  // ============================================

  private renderWizardFooter(): JSX.Element | null {
    const { currentFlowStep, acknowledgement } = this.state;

    // Don't show footer for completed state or already-acknowledged policies
    if (currentFlowStep === 'complete') return null;
    if (acknowledgement?.AckStatus === 'Acknowledged') return null;

    const steps = this.getWizardSteps();
    const currentIndex = this.getStepIndex(currentFlowStep);
    const totalSteps = steps.length;
    const stepLabels: Record<ReadFlowStep, string> = {
      reading: 'Read the policy document',
      quiz: 'Complete the comprehension quiz',
      acknowledge: 'Acknowledge the policy',
      complete: 'Policy acknowledged successfully'
    };

    return (
      <div className={styles.wizardFooter}>
        <div className={styles.footerInner}>
          <div className={styles.footerLeft}>
            <DefaultButton
              text="Back"
              iconProps={{ iconName: 'ChevronLeft' }}
              onClick={currentIndex === 0
                ? () => { window.location.href = '/sites/PolicyManager/SitePages/MyPolicies.aspx'; }
                : this.handleWizardBack
              }
            />
          </div>
          <Text variant="small" style={{ color: '#64748b' }}>
            Step {currentIndex + 1} of {totalSteps} — {stepLabels[currentFlowStep]}
          </Text>
          <div className={styles.footerRight}>
            <PrimaryButton
              text={this.getNextButtonText()}
              iconProps={{ iconName: 'ChevronRight' }}
              iconPosition="after"
              disabled={!this.canGoNext()}
              onClick={this.handleWizardNext}
            />
          </div>
        </div>
      </div>
    );
  }

  // ============================================
  // REVIEW MODE — Reviewer decision UI
  // ============================================

  private renderReviewMode(): JSX.Element {
    const { policy } = this.state;
    const st = this.state as any;
    const reviewDecision = st.reviewDecision || '';
    const reviewComments = st.reviewComments || '';
    const reviewChecklist: boolean[] = st.reviewChecklist || [false, false, false, false, false, false];
    const reviewSubmitting = st.reviewSubmitting || false;
    const reviewerItems: any[] = st.reviewerItems || [];
    const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';
    const currentUserEmail = this.props.context?.pageContext?.user?.email || '';

    // Load reviewers on first render
    if (!st._reviewersLoaded && policy) {
      this.setState({ _reviewersLoaded: true } as any);
      this.props.sp.web.lists.getByTitle('PM_PolicyReviewers')
        .items.filter(`PolicyId eq ${policy.Id}`)
        .select('Id', 'ReviewerId', 'ReviewerType', 'ReviewStatus', 'ReviewComments', 'ReviewedDate', 'Reviewer/Id', 'Reviewer/Title', 'Reviewer/EMail')
        .expand('Reviewer')
        .top(20)()
        .then((items: any[]) => { if (this._isMounted) this.setState({ reviewerItems: items } as any); })
        .catch(() => { /* list may not exist */ });
    }

    const checklistLabels = [
      'Content accuracy and completeness',
      'Compliance alignment (regulatory requirements)',
      'Formatting and readability',
      'Key points accurately summarise policy',
      'Appropriate risk classification',
      'Target audience correctly defined'
    ];

    const handleSubmitReview = async (): Promise<void> => {
      if (!reviewDecision) return;
      if ((reviewDecision === 'changes' || reviewDecision === 'reject') && !reviewComments.trim()) {
        void this.dialogManager.showAlert('Comments are required for Request Changes and Reject decisions.', { variant: 'warning' });
        return;
      }
      this.setState({ reviewSubmitting: true } as any);
      try {
        const currentUserId = this.props.context?.pageContext?.legacyPageContext?.userId || 0;
        const currentUserName = this.props.context?.pageContext?.user?.displayName || '';

        // Find this reviewer's record
        const myReview = reviewerItems.find((r: any) => r.Reviewer?.EMail === currentUserEmail || r.ReviewerId === currentUserId);

        // Update reviewer status in PM_PolicyReviewers
        if (myReview) {
          const newStatus = reviewDecision === 'approve' ? 'Approved' : reviewDecision === 'changes' ? 'Revision Requested' : 'Rejected';
          await this.props.sp.web.lists.getByTitle('PM_PolicyReviewers')
            .items.getById(myReview.Id).update({
              ReviewStatus: newStatus,
              ReviewComments: reviewComments,
              ReviewedDate: new Date().toISOString()
            });
        }

        // Determine next action based on all reviewer statuses
        const updatedReviewers = reviewerItems.map((r: any) => {
          if (r.Id === myReview?.Id) return { ...r, ReviewStatus: reviewDecision === 'approve' ? 'Approved' : reviewDecision === 'changes' ? 'Revision Requested' : 'Rejected' };
          return r;
        });
        const allApproved = updatedReviewers.every((r: any) => r.ReviewStatus === 'Approved');
        const anyRejected = updatedReviewers.some((r: any) => r.ReviewStatus === 'Rejected' || r.ReviewStatus === 'Revision Requested');

        // Update policy status
        if (reviewDecision === 'reject' || reviewDecision === 'changes') {
          await this.props.sp.web.lists.getByTitle('PM_Policies')
            .items.getById(policy!.Id).update({ PolicyStatus: 'Draft' });
          // Reset ALL reviewer statuses to Pending for resubmission
          try {
            for (const r of reviewerItems) {
              await this.props.sp.web.lists.getByTitle('PM_PolicyReviewers')
                .items.getById(r.Id).update({ ReviewStatus: 'Pending', ReviewComments: '', ReviewedDate: null });
            }
          } catch { /* best-effort */ }
        } else if (allApproved) {
          // Check if there are final approvers still pending
          const hasApprovers = updatedReviewers.some((r: any) => r.ReviewerType === 'Final Approver' || r.ReviewerType === 'Executive Approver');
          const allApproversApproved = updatedReviewers
            .filter((r: any) => r.ReviewerType === 'Final Approver' || r.ReviewerType === 'Executive Approver')
            .every((r: any) => r.ReviewStatus === 'Approved');

          if (hasApprovers && allApproversApproved) {
            // All reviewers AND approvers approved → Approved
            await this.props.sp.web.lists.getByTitle('PM_Policies')
              .items.getById(policy!.Id).update({ PolicyStatus: 'Approved' });
          } else if (hasApprovers && !allApproversApproved) {
            // Reviewers approved but approvers still pending → Pending Approval
            await this.props.sp.web.lists.getByTitle('PM_Policies')
              .items.getById(policy!.Id).update({ PolicyStatus: 'Pending Approval' });
          } else {
            // No separate approvers — all reviewers approved → Approved
            await this.props.sp.web.lists.getByTitle('PM_Policies')
              .items.getById(policy!.Id).update({ PolicyStatus: 'Approved' });
          }
        }

        // Audit log
        try {
          await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
            Title: `${reviewDecision === 'approve' ? 'Approved' : reviewDecision === 'changes' ? 'ChangesRequested' : 'Rejected'} - Policy ${policy!.Id}`,
            PolicyId: policy!.Id,
            EntityType: 'Policy',
            EntityId: policy!.Id,
            AuditAction: reviewDecision === 'approve' ? 'ReviewApproved' : reviewDecision === 'changes' ? 'ChangesRequested' : 'ReviewRejected',
            ActionDescription: reviewComments || `Policy ${reviewDecision === 'approve' ? 'approved' : reviewDecision === 'changes' ? 'changes requested' : 'rejected'} by ${currentUserName}`,
            PerformedByEmail: currentUserEmail,
            ActionDate: new Date().toISOString()
          });
        } catch { /* best-effort */ }

        // Notify policy author
        try {
          const authorEmail = (policy as any)._policyOwnerEmail || (policy as any).PolicyOwner || '';
          if (authorEmail) {
            const decisionLabel = reviewDecision === 'approve' ? 'Approved' : reviewDecision === 'changes' ? 'Changes Requested' : 'Rejected';
            const policyUrl = `${siteUrl}/SitePages/PolicyBuilder.aspx?editPolicyId=${policy!.Id}`;
            const emailHtml = `
              <div style="font-family:'Segoe UI',sans-serif;max-width:600px;margin:0 auto">
                <div style="background:linear-gradient(135deg,${reviewDecision === 'approve' ? '#059669,#047857' : reviewDecision === 'changes' ? '#d97706,#b45309' : '#dc2626,#b91c1c'});padding:24px 32px;border-radius:8px 8px 0 0">
                  <h1 style="color:#fff;margin:0;font-size:20px">Review ${decisionLabel}</h1>
                  <p style="color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:13px">Policy Manager — DWx Digital Workplace</p>
                </div>
                <div style="background:#fff;padding:24px 32px;border:1px solid #e2e8f0;border-top:none">
                  <p style="font-size:14px;color:#475569"><strong>${escapeHtml(currentUserName)}</strong> has ${reviewDecision === 'approve' ? 'approved' : reviewDecision === 'changes' ? 'requested changes to' : 'rejected'} your policy:</p>
                  <div style="background:#f8fafc;border-left:4px solid ${reviewDecision === 'approve' ? '#059669' : reviewDecision === 'changes' ? '#d97706' : '#dc2626'};padding:16px;border-radius:0 4px 4px 0;margin:16px 0">
                    <p style="margin:0;font-weight:600;font-size:15px;color:#0f172a">${escapeHtml(policy!.PolicyName || policy!.Title || '')}</p>
                  </div>
                  ${reviewComments ? `<div style="margin:16px 0;padding:12px;background:#f8fafc;border-radius:4px"><p style="margin:0 0 4px;font-size:11px;color:#94a3b8;font-weight:600">REVIEWER COMMENTS</p><p style="margin:0;font-size:13px;color:#475569">${escapeHtml(reviewComments)}</p></div>` : ''}
                  <p style="margin:24px 0 16px"><a href="${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy!.Id}&mode=${reviewDecision === 'approve' ? 'approve' : 'review'}" style="background:${reviewDecision === 'approve' ? '#059669' : '#0d9488'};color:#fff;padding:10px 24px;border-radius:4px;text-decoration:none;font-weight:600;font-size:14px;display:inline-block">${reviewDecision === 'approve' ? 'Review & Approve Policy' : 'Edit Policy'}</a></p>
                </div>
                <div style="background:#f8fafc;padding:16px 32px;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 8px 8px;text-align:center">
                  <p style="margin:0;font-size:11px;color:#94a3b8">First Digital — DWx Policy Manager</p>
                </div>
              </div>`;
            await this.props.sp.web.lists.getByTitle('PM_NotificationQueue').items.add({
              Title: `Review ${decisionLabel}: ${policy!.PolicyName || policy!.Title}`,
              RecipientEmail: authorEmail,
              RecipientName: '',
              PolicyId: policy!.Id,
              PolicyTitle: policy!.PolicyName || policy!.Title || '',
              NotificationType: reviewDecision === 'approve' ? 'ApprovalApproved' : 'ApprovalRejected',
              Channel: 'Email',
              Message: emailHtml,
              QueueStatus: 'Pending',
              Priority: 'High'
            });
          }
        } catch { /* notification best-effort */ }

        // Show success and redirect
        const msg = reviewDecision === 'approve'
          ? (allApproved ? 'All reviews complete. The policy is now pending final approval.' : 'Your approval has been recorded. Waiting for other reviewers.')
          : reviewDecision === 'changes' ? 'Your change request has been sent to the author.'
          : 'The policy has been rejected and the author has been notified.';
        await this.dialogManager.showAlert(msg, { variant: reviewDecision === 'approve' ? 'success' : 'warning', title: reviewDecision === 'approve' ? 'Review Approved' : reviewDecision === 'changes' ? 'Changes Requested' : 'Policy Rejected' });
        window.location.href = `${siteUrl}/SitePages/PolicyAuthor.aspx`;
      } catch (err) {
        console.error('Review submission failed:', err);
        void this.dialogManager.showAlert('Failed to submit review. Please try again.', { variant: 'error' });
      } finally {
        this.setState({ reviewSubmitting: false } as any);
      }
    };

    return (
      <ErrorBoundary fallbackMessage="An error occurred in Policy Review. Please try again.">
      <JmlAppLayout
        context={this.props.context} sp={this.props.sp}
        pageTitle="Review Policy" pageDescription="" pageIcon="RedEye"
        breadcrumbs={[{ text: 'Policy Manager', url: siteUrl }, { text: 'Review Policy' }]}
        activeNavKey="author"
      >
        {/* Review Banner */}
        <div style={{ background: 'linear-gradient(135deg, #f0fdfa 0%, #ecfdf5 100%)', borderBottom: '2px solid #0d9488', padding: '14px 0' }}>
          <div style={{ maxWidth: 1400, margin: '0 auto', padding: '0 24px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
              <span style={{ display: 'flex', alignItems: 'center', gap: 8, background: '#0d9488', color: '#fff', padding: '6px 16px', borderRadius: 4, fontSize: 12, fontWeight: 600 }}>
                <Icon iconName="RedEye" styles={{ root: { fontSize: 14 } }} /> Review Mode
              </span>
              <span style={{ fontSize: 12, color: '#64748b' }}>Submitted by <strong style={{ color: '#0f172a' }}>{(policy as any).PolicyOwner || 'Author'}</strong></span>
              <span style={{ fontSize: 12, color: '#64748b' }}>{policy!.PolicyNumber}</span>
            </div>
            <div style={{ display: 'flex', gap: 8 }}>
              <span style={{ padding: '4px 12px', borderRadius: 4, fontSize: 11, fontWeight: 600, background: '#dbeafe', color: '#2563eb' }}>In Review</span>
              <span style={{ padding: '4px 12px', borderRadius: 4, fontSize: 11, fontWeight: 600, background: '#fee2e2', color: '#dc2626' }}>{policy!.ComplianceRisk || 'Medium'} Risk</span>
            </div>
          </div>
        </div>

        {/* Main Layout: Content + Review Panel */}
        <div style={{ maxWidth: 1400, margin: '24px auto', padding: '0 24px', display: 'grid', gridTemplateColumns: '1fr 380px', gap: 24 }}>

          {/* Left: Policy Content */}
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
            <div style={{ padding: '24px 32px', borderBottom: '1px solid #e2e8f0' }}>
              <Text style={{ fontSize: 22, fontWeight: 700, color: '#0f172a', display: 'block' }}>{policy!.PolicyName || policy!.Title}</Text>
              <Text style={{ fontSize: 13, color: '#64748b', marginTop: 4, display: 'block' }}>{policy!.PolicyCategory} &bull; v{(policy as any).VersionNumber || '1.0'} &bull; Effective: {policy!.EffectiveDate ? new Date(policy!.EffectiveDate).toLocaleDateString() : 'TBD'}</Text>
            </div>
            <div style={{ padding: 32, minHeight: 400, lineHeight: 1.7, fontSize: 14, color: '#334155' }}
              dangerouslySetInnerHTML={{ __html: policy!.HTMLContent || policy!.PolicyContent || policy!.Description || '<p>No content available.</p>' }}
            />
            {/* Key Points */}
            {(policy as any).InternalNotes && (
              <div style={{ padding: '20px 32px', background: '#f8fafc', borderTop: '1px solid #e2e8f0' }}>
                <Text style={{ fontSize: 12, fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 1, marginBottom: 10, display: 'block' }}>Key Points</Text>
                {(() => { try { return JSON.parse((policy as any).InternalNotes).map((kp: string, i: number) => (
                  <div key={i} style={{ display: 'flex', alignItems: 'flex-start', gap: 8, padding: '6px 0', fontSize: 13, color: '#475569' }}>
                    <span style={{ width: 6, height: 6, borderRadius: '50%', background: '#0d9488', marginTop: 6, flexShrink: 0 }} />
                    {kp}
                  </div>
                )); } catch { return null; } })()}
              </div>
            )}
          </div>

          {/* Right: Review Panel (scrolls with content) */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16, alignSelf: 'start', position: 'sticky', top: 16 }}>

            {/* Decision */}
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
              <div style={{ padding: '16px 20px', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: 8 }}>
                <div style={{ width: 28, height: 28, borderRadius: 6, background: '#f0fdfa', color: '#0d9488', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 14 }}>&#x2714;</div>
                <Text style={{ fontSize: 14, fontWeight: 600, color: '#0f172a' }}>Your Review Decision</Text>
              </div>
              <div style={{ padding: 20 }}>
                <Stack tokens={{ childrenGap: 8 }}>
                  {[
                    { key: 'approve', label: 'Approve', desc: 'Policy meets all requirements', icon: '&#x2714;', bg: '#dcfce7', color: '#059669' },
                    { key: 'changes', label: 'Request Changes', desc: 'Needs revisions before approval', icon: '&#x270F;', bg: '#fef3c7', color: '#d97706' },
                    { key: 'reject', label: 'Reject', desc: 'Does not meet standards', icon: '&#x2716;', bg: '#fee2e2', color: '#dc2626' }
                  ].map(d => (
                    <div key={d.key}
                      role="button" tabIndex={0}
                      onClick={() => this.setState({ reviewDecision: d.key } as any)}
                      onKeyDown={(e) => { if (e.key === 'Enter') this.setState({ reviewDecision: d.key } as any); }}
                      style={{
                        display: 'flex', alignItems: 'center', gap: 10, padding: '14px 16px', borderRadius: 8,
                        border: `2px solid ${reviewDecision === d.key ? '#0d9488' : '#e2e8f0'}`,
                        background: reviewDecision === d.key ? '#f0fdfa' : '#fff', cursor: 'pointer', transition: 'all 0.15s'
                      }}
                    >
                      <div style={{ width: 36, height: 36, borderRadius: 8, background: d.bg, color: d.color, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 18 }} dangerouslySetInnerHTML={{ __html: d.icon }} />
                      <div>
                        <Text style={{ fontWeight: 600, fontSize: 13, color: '#0f172a', display: 'block' }}>{d.label}</Text>
                        <Text style={{ fontSize: 11, color: '#94a3b8' }}>{d.desc}</Text>
                      </div>
                    </div>
                  ))}
                </Stack>
                <TextField
                  label="Comments"
                  multiline rows={4}
                  value={reviewComments}
                  onChange={(_, v) => this.setState({ reviewComments: v || '' } as any)}
                  placeholder="Add your review comments... (required for Request Changes and Reject)"
                  styles={{ root: { marginTop: 16 } }}
                />
                <PrimaryButton
                  text={reviewSubmitting ? 'Submitting...' : reviewDecision === 'approve' ? 'Submit Approval' : reviewDecision === 'changes' ? 'Submit Change Request' : reviewDecision === 'reject' ? 'Submit Rejection' : 'Select a Decision'}
                  disabled={!reviewDecision || reviewSubmitting}
                  onClick={handleSubmitReview}
                  styles={{
                    root: {
                      width: '100%', marginTop: 16, borderRadius: 6, height: 40,
                      background: reviewDecision === 'approve' ? '#0d9488' : reviewDecision === 'changes' ? '#d97706' : reviewDecision === 'reject' ? '#dc2626' : '#94a3b8',
                      borderColor: reviewDecision === 'approve' ? '#0d9488' : reviewDecision === 'changes' ? '#d97706' : reviewDecision === 'reject' ? '#dc2626' : '#94a3b8'
                    },
                    rootHovered: {
                      background: reviewDecision === 'approve' ? '#0f766e' : reviewDecision === 'changes' ? '#b45309' : reviewDecision === 'reject' ? '#b91c1c' : '#64748b',
                      borderColor: reviewDecision === 'approve' ? '#0f766e' : reviewDecision === 'changes' ? '#b45309' : reviewDecision === 'reject' ? '#b91c1c' : '#64748b'
                    }
                  }}
                />
              </div>
            </div>

            {/* Checklist */}
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
              <div style={{ padding: '16px 20px', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: 8 }}>
                <div style={{ width: 28, height: 28, borderRadius: 6, background: '#dbeafe', color: '#2563eb', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 14 }}>&#x2611;</div>
                <Text style={{ fontSize: 14, fontWeight: 600, color: '#0f172a' }}>Review Checklist</Text>
              </div>
              <div style={{ padding: '12px 20px' }}>
                {checklistLabels.map((label, i) => (
                  <label key={i} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 0', borderBottom: i < 5 ? '1px solid #f1f5f9' : 'none', fontSize: 13, cursor: 'pointer' }}>
                    <input type="checkbox" checked={reviewChecklist[i]} onChange={() => {
                      const updated = [...reviewChecklist];
                      updated[i] = !updated[i];
                      this.setState({ reviewChecklist: updated } as any);
                    }} style={{ accentColor: '#0d9488', width: 16, height: 16 }} />
                    <span style={{ color: reviewChecklist[i] ? '#94a3b8' : '#475569', textDecoration: reviewChecklist[i] ? 'line-through' : 'none' }}>{label}</span>
                  </label>
                ))}
              </div>
            </div>

            {/* Review Chain */}
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
              <div style={{ padding: '16px 20px', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: 8 }}>
                <div style={{ width: 28, height: 28, borderRadius: 6, background: '#ede9fe', color: '#7c3aed', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 14 }}>&#x1F465;</div>
                <Text style={{ fontSize: 14, fontWeight: 600, color: '#0f172a' }}>Review Chain</Text>
              </div>
              <div style={{ padding: '12px 20px' }}>
                {reviewerItems.length === 0 ? (
                  <Text style={{ fontSize: 12, color: '#94a3b8', fontStyle: 'italic' }}>No reviewers assigned</Text>
                ) : reviewerItems.map((r: any, i: number) => {
                  const isMe = r.Reviewer?.EMail === currentUserEmail;
                  const initials = (r.Reviewer?.Title || '??').split(' ').map((n: string) => n[0]).join('').substring(0, 2).toUpperCase();
                  const statusColor = r.ReviewStatus === 'Approved' ? { bg: '#dcfce7', color: '#059669' } :
                    r.ReviewStatus === 'Rejected' || r.ReviewStatus === 'Revision Requested' ? { bg: '#fee2e2', color: '#dc2626' } :
                    isMe ? { bg: '#dbeafe', color: '#2563eb' } : { bg: '#fef3c7', color: '#d97706' };
                  return (
                    <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 0', borderBottom: i < reviewerItems.length - 1 ? '1px solid #f1f5f9' : 'none' }}>
                      <div style={{ width: 28, height: 28, borderRadius: '50%', background: statusColor.bg, color: statusColor.color, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 10, fontWeight: 700 }}>{initials}</div>
                      <span style={{ fontSize: 13, fontWeight: 500, color: '#0f172a', flex: 1 }}>{r.Reviewer?.Title || 'Unknown'}</span>
                      <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 4, background: statusColor.bg, color: statusColor.color }}>
                        {isMe ? 'You' : r.ReviewStatus || 'Pending'}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Previous Comments */}
            {reviewerItems.some((r: any) => r.ReviewComments) && (
              <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
                <div style={{ padding: '16px 20px', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: 8 }}>
                  <div style={{ width: 28, height: 28, borderRadius: 6, background: '#fce7f3', color: '#db2777', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 14 }}>&#x1F4AC;</div>
                  <Text style={{ fontSize: 14, fontWeight: 600, color: '#0f172a' }}>Review Comments</Text>
                </div>
                <div style={{ padding: 20 }}>
                  {reviewerItems.filter((r: any) => r.ReviewComments).map((r: any, i: number) => (
                    <div key={i} style={{ padding: 12, background: '#f8fafc', borderRadius: 6, marginBottom: 8 }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 6 }}>
                        <span style={{ fontSize: 12, fontWeight: 600, color: '#0f172a' }}>{r.Reviewer?.Title || 'Reviewer'}</span>
                        <span style={{ fontSize: 10, color: '#94a3b8' }}>{r.ReviewedDate ? new Date(r.ReviewedDate).toLocaleDateString() : ''}</span>
                      </div>
                      <Text style={{ fontSize: 12, color: '#475569', lineHeight: '1.5' }}>{r.ReviewComments}</Text>
                    </div>
                  ))}
                </div>
              </div>
            )}

          </div>
        </div>
        <this.dialogManager.DialogComponent />
      </JmlAppLayout>
      </ErrorBoundary>
    );
  }

  // ============================================
  // APPROVAL MODE — Final Approver decision UI
  // ============================================

  private renderApprovalMode(): JSX.Element {
    const { policy } = this.state;
    const st = this.state as any;
    const reviewDecision = st.reviewDecision || '';
    const reviewComments = st.reviewComments || '';
    const reviewSubmitting = st.reviewSubmitting || false;
    const reviewerItems: any[] = st.reviewerItems || [];
    const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';
    const currentUserEmail = this.props.context?.pageContext?.user?.email || '';

    // Load reviewers on first render
    if (!st._reviewersLoaded && policy) {
      this.setState({ _reviewersLoaded: true } as any);
      this.props.sp.web.lists.getByTitle('PM_PolicyReviewers')
        .items.filter(`PolicyId eq ${policy.Id}`)
        .select('Id', 'ReviewerId', 'ReviewerType', 'ReviewStatus', 'ReviewComments', 'ReviewedDate', 'Reviewer/Id', 'Reviewer/Title', 'Reviewer/EMail')
        .expand('Reviewer')
        .top(20)()
        .then((items: any[]) => { if (this._isMounted) this.setState({ reviewerItems: items } as any); })
        .catch(() => { /* list may not exist */ });
    }

    const handleSubmitApproval = async (): Promise<void> => {
      if (!reviewDecision) return;
      if ((reviewDecision === 'return' || reviewDecision === 'reject') && !reviewComments.trim()) {
        void this.dialogManager.showAlert('Comments are required when returning or rejecting a policy.', { variant: 'warning' });
        return;
      }
      this.setState({ reviewSubmitting: true } as any);
      try {
        const currentUserId = this.props.context?.pageContext?.legacyPageContext?.userId || 0;
        const currentUserName = this.props.context?.pageContext?.user?.displayName || '';

        // Find this approver's record
        const myRecord = reviewerItems.find((r: any) => r.Reviewer?.EMail === currentUserEmail || r.ReviewerId === currentUserId);

        // Update approver status
        if (myRecord) {
          const newStatus = reviewDecision === 'approve' ? 'Approved' : reviewDecision === 'return' ? 'Revision Requested' : 'Rejected';
          await this.props.sp.web.lists.getByTitle('PM_PolicyReviewers')
            .items.getById(myRecord.Id).update({
              ReviewStatus: newStatus,
              ReviewComments: reviewComments,
              ReviewedDate: new Date().toISOString()
            });
        }

        // Update policy status
        if (reviewDecision === 'approve') {
          // Check if all approvers have approved
          const allItems = reviewerItems.map((r: any) => r.Id === myRecord?.Id ? { ...r, ReviewStatus: 'Approved' } : r);
          const allApproversApproved = allItems
            .filter((r: any) => r.ReviewerType === 'Final Approver' || r.ReviewerType === 'Executive Approver')
            .every((r: any) => r.ReviewStatus === 'Approved');

          if (allApproversApproved) {
            await this.props.sp.web.lists.getByTitle('PM_Policies')
              .items.getById(policy!.Id).update({ PolicyStatus: 'Approved' });
          }
        } else {
          // Return or Reject → back to Draft
          await this.props.sp.web.lists.getByTitle('PM_Policies')
            .items.getById(policy!.Id).update({ PolicyStatus: 'Draft' });
          // Reset all reviewer/approver statuses
          try {
            for (const r of reviewerItems) {
              await this.props.sp.web.lists.getByTitle('PM_PolicyReviewers')
                .items.getById(r.Id).update({ ReviewStatus: 'Pending', ReviewComments: '', ReviewedDate: null });
            }
          } catch { /* best-effort */ }
        }

        // Audit log
        try {
          await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
            Title: `${reviewDecision === 'approve' ? 'ApprovalGranted' : reviewDecision === 'return' ? 'ReturnedToAuthor' : 'ApprovalRejected'} - Policy ${policy!.Id}`,
            PolicyId: policy!.Id, EntityType: 'Policy', EntityId: policy!.Id,
            AuditAction: reviewDecision === 'approve' ? 'ApprovalGranted' : reviewDecision === 'return' ? 'ReturnedToAuthor' : 'ApprovalRejected',
            ActionDescription: reviewComments || `Policy ${reviewDecision === 'approve' ? 'approved' : 'returned'} by ${currentUserName}`,
            PerformedByEmail: currentUserEmail, ActionDate: new Date().toISOString()
          });
        } catch { /* best-effort */ }

        // Notify author
        try {
          const authorEmail = (policy as any)._policyOwnerEmail || (policy as any).PolicyOwner || '';
          if (authorEmail) {
            const decisionLabel = reviewDecision === 'approve' ? 'Approved' : reviewDecision === 'return' ? 'Returned for Changes' : 'Rejected';
            const headerColor = reviewDecision === 'approve' ? '#059669,#047857' : reviewDecision === 'return' ? '#d97706,#b45309' : '#dc2626,#b91c1c';
            await this.props.sp.web.lists.getByTitle('PM_NotificationQueue').items.add({
              Title: `Approval ${decisionLabel}: ${escapeHtml(policy!.PolicyName || policy!.Title || '')}`,
              RecipientEmail: authorEmail, RecipientName: '',
              PolicyId: policy!.Id, PolicyTitle: policy!.PolicyName || policy!.Title || '',
              NotificationType: reviewDecision === 'approve' ? 'ApprovalApproved' : 'ApprovalRejected',
              Channel: 'Email',
              Message: `<div style="font-family:'Segoe UI',sans-serif;max-width:600px;margin:0 auto"><div style="background:linear-gradient(135deg,${headerColor});padding:24px 32px;border-radius:8px 8px 0 0"><h1 style="color:#fff;margin:0;font-size:20px">Approval ${decisionLabel}</h1><p style="color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:13px">Policy Manager — Final Approval</p></div><div style="background:#fff;padding:24px 32px;border:1px solid #e2e8f0;border-top:none"><p style="font-size:14px;color:#475569"><strong>${escapeHtml(currentUserName)}</strong> has ${reviewDecision === 'approve' ? 'approved' : reviewDecision === 'return' ? 'returned' : 'rejected'} your policy:</p><div style="background:#f8fafc;border-left:4px solid ${reviewDecision === 'approve' ? '#059669' : '#d97706'};padding:16px;border-radius:0 4px 4px 0;margin:16px 0"><p style="margin:0;font-weight:600;font-size:15px;color:#0f172a">${escapeHtml(policy!.PolicyName || policy!.Title || '')}</p></div>${reviewComments ? `<div style="margin:16px 0;padding:12px;background:#f8fafc;border-radius:4px"><p style="margin:0 0 4px;font-size:11px;color:#94a3b8;font-weight:600">APPROVER COMMENTS</p><p style="margin:0;font-size:13px;color:#475569">${escapeHtml(reviewComments)}</p></div>` : ''}<p style="margin:24px 0 16px"><a href="${siteUrl}/SitePages/${reviewDecision === 'approve' ? 'PolicyDetails' : 'PolicyBuilder'}.aspx?${reviewDecision === 'approve' ? 'policyId' : 'editPolicyId'}=${policy!.Id}" style="background:${reviewDecision === 'approve' ? '#059669' : '#0d9488'};color:#fff;padding:10px 24px;border-radius:4px;text-decoration:none;font-weight:600;font-size:14px;display:inline-block">${reviewDecision === 'approve' ? 'View Approved Policy' : 'Edit Policy'}</a></p></div><div style="background:#f8fafc;padding:16px 32px;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 8px 8px;text-align:center"><p style="margin:0;font-size:11px;color:#94a3b8">First Digital — DWx Policy Manager</p></div></div>`,
              QueueStatus: 'Pending', Priority: 'High'
            });
          }
        } catch { /* notification best-effort */ }

        const msg = reviewDecision === 'approve'
          ? 'Policy approved! The author can now publish it.'
          : reviewDecision === 'return' ? 'Policy returned to the author for changes.'
          : 'Policy rejected. The author has been notified.';
        await this.dialogManager.showAlert(msg, { variant: reviewDecision === 'approve' ? 'success' : 'warning', title: reviewDecision === 'approve' ? 'Approval Granted' : 'Policy Returned' });
        window.location.href = `${siteUrl}/SitePages/PolicyManagerView.aspx?tab=approvals`;
      } catch (err) {
        console.error('Approval submission failed:', err);
        void this.dialogManager.showAlert('Failed to submit approval. Please try again.', { variant: 'error' });
      } finally {
        this.setState({ reviewSubmitting: false } as any);
      }
    };

    return (
      <ErrorBoundary fallbackMessage="An error occurred in Policy Approval. Please try again.">
      <JmlAppLayout
        context={this.props.context} sp={this.props.sp}
        pageTitle="Approve Policy" pageDescription="" pageIcon="CheckboxComposite"
        breadcrumbs={[{ text: 'Policy Manager', url: siteUrl }, { text: 'Approve Policy' }]}
        activeNavKey="manager"
      >
        {/* Approval Banner — Green/Emerald */}
        <div style={{ background: 'linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%)', borderBottom: '2px solid #059669', padding: '14px 0' }}>
          <div style={{ maxWidth: 1400, margin: '0 auto', padding: '0 24px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
              <span style={{ display: 'flex', alignItems: 'center', gap: 8, background: '#059669', color: '#fff', padding: '6px 16px', borderRadius: 4, fontSize: 12, fontWeight: 600 }}>
                <Icon iconName="CheckboxComposite" styles={{ root: { fontSize: 14 } }} /> Approval Mode
              </span>
              <span style={{ fontSize: 12, color: '#064e3b' }}>Final approval by <strong style={{ color: '#0f172a' }}>{this.props.context?.pageContext?.user?.displayName || 'Approver'}</strong></span>
              <span style={{ fontSize: 12, color: '#064e3b' }}>{policy!.PolicyNumber}</span>
            </div>
            <div style={{ display: 'flex', gap: 8 }}>
              <span style={{ padding: '4px 12px', borderRadius: 4, fontSize: 11, fontWeight: 600, background: '#fef3c7', color: '#d97706' }}>Pending Approval</span>
              <span style={{ padding: '4px 12px', borderRadius: 4, fontSize: 11, fontWeight: 600, background: '#fee2e2', color: '#dc2626' }}>{policy!.ComplianceRisk || 'Medium'} Risk</span>
            </div>
          </div>
        </div>

        {/* Main Layout */}
        <div style={{ maxWidth: 1400, margin: '24px auto', padding: '0 24px', display: 'grid', gridTemplateColumns: '1fr 380px', gap: 24 }}>

          {/* Left: Policy Content */}
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
            <div style={{ padding: '24px 32px', borderBottom: '1px solid #e2e8f0' }}>
              <Text style={{ fontSize: 22, fontWeight: 700, color: '#0f172a', display: 'block' }}>{policy!.PolicyName || policy!.Title}</Text>
              <Text style={{ fontSize: 13, color: '#64748b', marginTop: 4, display: 'block' }}>{policy!.PolicyCategory} &bull; v{(policy as any).VersionNumber || '1.0'}</Text>
            </div>
            <div style={{ padding: 32, minHeight: 400, lineHeight: 1.7, fontSize: 14, color: '#334155' }}
              dangerouslySetInnerHTML={{ __html: sanitizeHtml(policy!.HTMLContent || policy!.PolicyContent || policy!.Description || '<p>No content available.</p>') }}
            />
            {(policy as any).InternalNotes && (
              <div style={{ padding: '20px 32px', background: '#f8fafc', borderTop: '1px solid #e2e8f0' }}>
                <Text style={{ fontSize: 12, fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 1, marginBottom: 10, display: 'block' }}>Key Points</Text>
                {(() => { try { return JSON.parse((policy as any).InternalNotes).map((kp: string, i: number) => (
                  <div key={i} style={{ display: 'flex', alignItems: 'flex-start', gap: 8, padding: '6px 0', fontSize: 13, color: '#475569' }}>
                    <span style={{ width: 6, height: 6, borderRadius: '50%', background: '#059669', marginTop: 6, flexShrink: 0 }} />
                    {kp}
                  </div>
                )); } catch { return null; } })()}
              </div>
            )}
          </div>

          {/* Right: Approval Panel */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16, alignSelf: 'start', position: 'sticky', top: 16 }}>

            {/* Decision */}
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
              <div style={{ padding: '16px 20px', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: 8, background: '#ecfdf5' }}>
                <div style={{ width: 28, height: 28, borderRadius: 6, background: '#059669', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 14 }}>&#x2714;</div>
                <Text style={{ fontSize: 14, fontWeight: 600, color: '#064e3b' }}>Your Approval Decision</Text>
              </div>
              <div style={{ padding: 20 }}>
                <Stack tokens={{ childrenGap: 8 }}>
                  {[
                    { key: 'approve', label: 'Approve for Publication', desc: 'Policy is ready to be published', icon: '&#x2714;', bg: '#dcfce7', color: '#059669' },
                    { key: 'return', label: 'Return to Author', desc: 'Needs revisions before approval', icon: '&#x21A9;', bg: '#fef3c7', color: '#d97706' },
                    { key: 'reject', label: 'Reject', desc: 'Does not meet governance standards', icon: '&#x2716;', bg: '#fee2e2', color: '#dc2626' }
                  ].map(d => (
                    <div key={d.key}
                      role="button" tabIndex={0}
                      onClick={() => this.setState({ reviewDecision: d.key } as any)}
                      onKeyDown={(e) => { if (e.key === 'Enter') this.setState({ reviewDecision: d.key } as any); }}
                      style={{
                        display: 'flex', alignItems: 'center', gap: 10, padding: '14px 16px', borderRadius: 8,
                        border: `2px solid ${reviewDecision === d.key ? '#059669' : '#e2e8f0'}`,
                        background: reviewDecision === d.key ? '#ecfdf5' : '#fff', cursor: 'pointer', transition: 'all 0.15s'
                      }}
                    >
                      <div style={{ width: 36, height: 36, borderRadius: 8, background: d.bg, color: d.color, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 18 }} dangerouslySetInnerHTML={{ __html: d.icon }} />
                      <div>
                        <Text style={{ fontWeight: 600, fontSize: 13, color: '#0f172a', display: 'block' }}>{d.label}</Text>
                        <Text style={{ fontSize: 11, color: '#94a3b8' }}>{d.desc}</Text>
                      </div>
                    </div>
                  ))}
                </Stack>
                <TextField
                  label="Approver Comments"
                  multiline rows={4}
                  value={reviewComments}
                  onChange={(_, v) => this.setState({ reviewComments: v || '' } as any)}
                  placeholder="Add comments... (required for Return and Reject)"
                  styles={{ root: { marginTop: 16 } }}
                />
                <PrimaryButton
                  text={reviewSubmitting ? 'Submitting...' : reviewDecision === 'approve' ? 'Grant Approval' : reviewDecision === 'return' ? 'Return to Author' : reviewDecision === 'reject' ? 'Reject Policy' : 'Select a Decision'}
                  disabled={!reviewDecision || reviewSubmitting}
                  onClick={handleSubmitApproval}
                  styles={{
                    root: {
                      width: '100%', marginTop: 16, borderRadius: 6, height: 40,
                      background: reviewDecision === 'approve' ? '#059669' : reviewDecision === 'return' ? '#d97706' : reviewDecision === 'reject' ? '#dc2626' : '#94a3b8',
                      borderColor: reviewDecision === 'approve' ? '#059669' : reviewDecision === 'return' ? '#d97706' : reviewDecision === 'reject' ? '#dc2626' : '#94a3b8'
                    },
                    rootHovered: {
                      background: reviewDecision === 'approve' ? '#047857' : reviewDecision === 'return' ? '#b45309' : reviewDecision === 'reject' ? '#b91c1c' : '#64748b',
                      borderColor: reviewDecision === 'approve' ? '#047857' : reviewDecision === 'return' ? '#b45309' : reviewDecision === 'reject' ? '#b91c1c' : '#64748b'
                    }
                  }}
                />
              </div>
            </div>

            {/* Review Chain + Previous Decisions */}
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
              <div style={{ padding: '16px 20px', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: 8 }}>
                <div style={{ width: 28, height: 28, borderRadius: 6, background: '#ede9fe', color: '#7c3aed', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 14 }}>&#x1F465;</div>
                <Text style={{ fontSize: 14, fontWeight: 600, color: '#0f172a' }}>Review & Approval Chain</Text>
              </div>
              <div style={{ padding: '12px 20px' }}>
                {reviewerItems.length === 0 ? (
                  <Text style={{ fontSize: 12, color: '#94a3b8', fontStyle: 'italic' }}>Loading...</Text>
                ) : reviewerItems.map((r: any, i: number) => {
                  const isMe = r.Reviewer?.EMail === currentUserEmail;
                  const initials = (r.Reviewer?.Title || '??').split(' ').map((n: string) => n[0]).join('').substring(0, 2).toUpperCase();
                  const isApprover = r.ReviewerType === 'Final Approver' || r.ReviewerType === 'Executive Approver';
                  const statusColor = r.ReviewStatus === 'Approved' ? { bg: '#dcfce7', color: '#059669' } :
                    r.ReviewStatus === 'Rejected' || r.ReviewStatus === 'Revision Requested' ? { bg: '#fee2e2', color: '#dc2626' } :
                    isMe ? { bg: '#dbeafe', color: '#2563eb' } : { bg: '#fef3c7', color: '#d97706' };
                  return (
                    <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 0', borderBottom: i < reviewerItems.length - 1 ? '1px solid #f1f5f9' : 'none' }}>
                      <div style={{ width: 28, height: 28, borderRadius: '50%', background: statusColor.bg, color: statusColor.color, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 10, fontWeight: 700 }}>{initials}</div>
                      <div style={{ flex: 1 }}>
                        <span style={{ fontSize: 13, fontWeight: 500, color: '#0f172a' }}>{r.Reviewer?.Title || 'Unknown'}</span>
                        <span style={{ fontSize: 10, color: '#94a3b8', marginLeft: 6 }}>{isApprover ? 'Approver' : 'Reviewer'}</span>
                      </div>
                      <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 4, background: statusColor.bg, color: statusColor.color }}>
                        {isMe ? 'You' : r.ReviewStatus || 'Pending'}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Previous Comments */}
            {reviewerItems.some((r: any) => r.ReviewComments) && (
              <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
                <div style={{ padding: '16px 20px', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: 8 }}>
                  <div style={{ width: 28, height: 28, borderRadius: 6, background: '#fce7f3', color: '#db2777', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 14 }}>&#x1F4AC;</div>
                  <Text style={{ fontSize: 14, fontWeight: 600, color: '#0f172a' }}>Reviewer Comments</Text>
                </div>
                <div style={{ padding: 20 }}>
                  {reviewerItems.filter((r: any) => r.ReviewComments).map((r: any, i: number) => (
                    <div key={i} style={{ padding: 12, background: '#f8fafc', borderRadius: 6, marginBottom: 8 }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 6 }}>
                        <div>
                          <span style={{ fontSize: 12, fontWeight: 600, color: '#0f172a' }}>{r.Reviewer?.Title || 'Reviewer'}</span>
                          <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 6px', borderRadius: 3, marginLeft: 6, background: r.ReviewStatus === 'Approved' ? '#dcfce7' : '#fee2e2', color: r.ReviewStatus === 'Approved' ? '#059669' : '#dc2626' }}>{r.ReviewStatus}</span>
                        </div>
                        <span style={{ fontSize: 10, color: '#94a3b8' }}>{r.ReviewedDate ? new Date(r.ReviewedDate).toLocaleDateString() : ''}</span>
                      </div>
                      <Text style={{ fontSize: 12, color: '#475569', lineHeight: '1.5' }}>{r.ReviewComments}</Text>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Policy Info */}
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
              <div style={{ padding: '16px 20px', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: 8 }}>
                <div style={{ width: 28, height: 28, borderRadius: 6, background: '#f1f5f9', color: '#64748b', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 14 }}>&#x2139;</div>
                <Text style={{ fontSize: 14, fontWeight: 600, color: '#0f172a' }}>Policy Details</Text>
              </div>
              <div style={{ padding: '12px 20px' }}>
                {[
                  { label: 'Policy Number', value: policy!.PolicyNumber },
                  { label: 'Category', value: policy!.PolicyCategory },
                  { label: 'Risk Level', value: policy!.ComplianceRisk },
                  { label: 'Effective Date', value: policy!.EffectiveDate ? new Date(policy!.EffectiveDate).toLocaleDateString() : 'TBD' },
                  { label: 'Acknowledgement', value: policy!.RequiresAcknowledgement ? 'Required' : 'No' },
                  { label: 'Owner', value: (policy as any).PolicyOwner || '' }
                ].map((item, i) => (
                  <div key={i} style={{ display: 'flex', justifyContent: 'space-between', padding: '6px 0', fontSize: 12, borderBottom: '1px solid #f8fafc' }}>
                    <span style={{ color: '#64748b' }}>{item.label}</span>
                    <span style={{ color: '#0f172a', fontWeight: 500 }}>{item.value || '-'}</span>
                  </div>
                ))}
              </div>
            </div>

          </div>
        </div>
        <this.dialogManager.DialogComponent />
      </JmlAppLayout>
      </ErrorBoundary>
    );
  }

  // ============================================
  // SIMPLE READER — Browse mode (from Policy Hub)
  // No wizard steps, no quiz, no acknowledgement.
  // Just a clean document viewer with back button.
  // ============================================

  private renderSimpleReader(): JSX.Element {
    // Signal SharePoint that the app is ready — hides the SP loading skeleton
    // This is normally done by JmlAppLayout, but the simple reader bypasses it
    try { signalAppReady(); } catch { /* ignore */ }

    const { policy, loading, error } = this.state;

    if (loading) {
      return (
        <ErrorBoundary fallbackMessage="An error occurred loading policy details. Please try again.">
          <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <Spinner size={SpinnerSize.large} label="Loading policy..." />
          </div>
        </ErrorBoundary>
      );
    }

    if (error || !policy) {
      return (
        <ErrorBoundary fallbackMessage="An error occurred loading policy details. Please try again.">
          <div style={{ minHeight: '100vh', padding: 40 }}>
            <MessageBar messageBarType={MessageBarType.error}>{error || 'Policy not found'}</MessageBar>
          </div>
        </ErrorBoundary>
      );
    }

    // Resolve document URL
    const rawDocUrl = policy.DocumentURL;
    const documentUrl: string | undefined = typeof rawDocUrl === 'string' ? rawDocUrl
      : (rawDocUrl && typeof rawDocUrl === 'object' && (rawDocUrl as { Url?: string }).Url)
        ? (rawDocUrl as { Url: string }).Url
        : undefined;
    // Policy body content — check all possible content fields:
    // 1. HTMLContent: from doc conversion pipeline (Word/PPT/Excel → styled HTML)
    // 2. PolicyContent: from rich text editor in Policy Builder
    // 3. Description: basic text fallback
    const bodyHtml = policy.HTMLContent || policy.PolicyContent || policy.Description || '';
    const hasContent = bodyHtml.length > 10;
    const isPdf = ext === 'pdf';
    const viewerUrl = documentUrl ? this.getDocumentViewerUrl(documentUrl) : '';
    const ext = documentUrl?.split('.').pop()?.toLowerCase() || '';

    // Category badge color
    const catColors: Record<string, { bg: string; color: string }> = {
      'Compliance': { bg: '#fef3c7', color: '#92400e' }, 'HR Policies': { bg: '#ccfbf1', color: '#0d9488' },
      'IT & Security': { bg: '#dbeafe', color: '#2563eb' }, 'Health & Safety': { bg: '#fef3c7', color: '#d97706' },
      'Governance': { bg: '#ede9fe', color: '#7c3aed' }, 'Ethics': { bg: '#dcfce7', color: '#059669' },
      'Financial': { bg: '#ede9fe', color: '#7c3aed' }, 'Data Privacy': { bg: '#dbeafe', color: '#0284c7' },
    };
    const cat = catColors[policy.PolicyCategory || ''] || { bg: '#f0f9ff', color: '#0369a1' };

    const riskColors: Record<string, { bg: string; color: string }> = {
      'Critical': { bg: '#fee2e2', color: '#dc2626' }, 'High': { bg: '#fef3c7', color: '#d97706' },
      'Medium': { bg: '#e0e7ff', color: '#6366f1' }, 'Low': { bg: '#dcfce7', color: '#16a34a' },
    };
    const risk = riskColors[policy.ComplianceRisk || ''] || { bg: '#f1f5f9', color: '#64748b' };

    const badgeStyle = (bg: string, color: string): React.CSSProperties => ({
      fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase',
      letterSpacing: 0.5, background: bg, color, display: 'inline-block'
    });

    return (
      <ErrorBoundary fallbackMessage="An error occurred loading policy details. Please try again.">
      <div style={{ minHeight: '100vh', display: 'flex', flexDirection: 'column', background: '#f8fafc' }}>

        {/* Breadcrumb bar */}
        <div style={{
          background: '#f8fafc', borderBottom: '1px solid #e2e8f0', padding: '10px 40px',
          display: 'flex', alignItems: 'center', gap: 8, fontSize: 12, color: '#64748b'
        }}>
          <a href="/sites/PolicyManager" style={{ color: '#0d9488', textDecoration: 'none', fontWeight: 500 }}>Policy Manager</a>
          <span style={{ color: '#cbd5e1' }}>/</span>
          <a href="/sites/PolicyManager/SitePages/PolicyHub.aspx" style={{ color: '#0d9488', textDecoration: 'none', fontWeight: 500 }}>Policy Hub</a>
          <span style={{ color: '#cbd5e1' }}>/</span>
          <span style={{ color: '#0f172a', fontWeight: 600 }}>{policy.PolicyName}</span>
        </div>

        {/* Policy header */}
        <div style={{ background: '#fff', borderBottom: '1px solid #e2e8f0', padding: '24px 40px' }}>
          <div style={{ maxWidth: 1400, margin: '0 auto' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
              <div>
                <h1 style={{ fontSize: 24, fontWeight: 700, color: '#0f172a', marginBottom: 6 }}>{policy.PolicyName}</h1>
                <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
                  <span style={{ fontSize: 12, color: '#0d9488', fontWeight: 600 }}>{policy.PolicyNumber}</span>
                  <span style={badgeStyle(cat.bg, cat.color)}>{policy.PolicyCategory}</span>
                  {policy.ComplianceRisk && <span style={badgeStyle(risk.bg, risk.color)}>{policy.ComplianceRisk}</span>}
                  <span style={badgeStyle('#dcfce7', '#16a34a')}>Published</span>
                  <span style={badgeStyle('#f1f5f9', '#64748b')}>v{policy.VersionNumber || policy.PolicyVersion || '1.0'}</span>
                  {policy.PublishedDate && <span style={{ fontSize: 11, color: '#94a3b8' }}>Published {new Date(policy.PublishedDate as any).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}</span>}
                  {policy.EstimatedReadTimeMinutes && <span style={{ fontSize: 11, color: '#94a3b8' }}>&middot; {policy.EstimatedReadTimeMinutes} min read</span>}
                </div>
              </div>
              <button
                onClick={() => { window.location.href = '/sites/PolicyManager/SitePages/PolicyHub.aspx'; }}
                style={{
                  display: 'flex', alignItems: 'center', gap: 6, padding: '8px 16px', borderRadius: 6,
                  fontSize: 13, fontWeight: 600, color: '#0d9488', background: '#fff', border: '1px solid #e2e8f0',
                  cursor: 'pointer', fontFamily: 'inherit', flexShrink: 0
                }}
              >
                <svg viewBox="0 0 24 24" fill="none" width="14" height="14"><path d="M19 12H5M12 19l-7-7 7-7" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>
                Back to Policy Hub
              </button>
            </div>
          </div>
        </div>

        {/* Document toolbar */}
        <div style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', padding: '8px 40px' }}>
          <div style={{ maxWidth: 1400, margin: '0 auto', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <svg viewBox="0 0 24 24" fill="none" width="16" height="16"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" stroke="#dc2626" strokeWidth="2"/><path d="M14 2v6h6" stroke="#dc2626" strokeWidth="2"/></svg>
              <span style={{ fontSize: 12, color: '#334155', fontWeight: 500 }}>{policy.PolicyNumber}-v{policy.VersionNumber || '1.0'}</span>
              <span style={{ fontSize: 10, color: '#94a3b8', background: '#f1f5f9', padding: '2px 6px', borderRadius: 3 }}>
                {ext.toUpperCase() || 'HTML'}
              </span>
            </div>
            <div style={{ display: 'flex', gap: 4 }}>
              {documentUrl && (
                <button
                  onClick={() => window.open(documentUrl, '_blank')}
                  style={{ display: 'flex', alignItems: 'center', gap: 5, padding: '6px 12px', borderRadius: 4, fontSize: 11, fontWeight: 500, color: '#64748b', background: '#fff', border: '1px solid #e2e8f0', cursor: 'pointer', fontFamily: 'inherit' }}
                >
                  <svg viewBox="0 0 24 24" fill="none" width="12" height="12"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>
                  Download
                </button>
              )}
              <button
                onClick={() => window.print()}
                style={{ display: 'flex', alignItems: 'center', gap: 5, padding: '6px 12px', borderRadius: 4, fontSize: 11, fontWeight: 500, color: '#64748b', background: '#fff', border: '1px solid #e2e8f0', cursor: 'pointer', fontFamily: 'inherit' }}
              >
                <svg viewBox="0 0 24 24" fill="none" width="12" height="12"><path d="M6 9V2h12v7M6 18H4a2 2 0 01-2-2v-5a2 2 0 012-2h16a2 2 0 012 2v5a2 2 0 01-2 2h-2M6 14h12v8H6z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>
                Print
              </button>
              <button
                onClick={() => {
                  const el = document.documentElement;
                  if (!document.fullscreenElement) { el.requestFullscreen().catch(() => {}); }
                  else { document.exitFullscreen().catch(() => {}); }
                }}
                style={{ display: 'flex', alignItems: 'center', gap: 5, padding: '6px 12px', borderRadius: 4, fontSize: 11, fontWeight: 500, color: '#fff', background: '#0d9488', border: '1px solid #0d9488', cursor: 'pointer', fontFamily: 'inherit' }}
              >
                <svg viewBox="0 0 24 24" fill="none" width="12" height="12"><path d="M8 3H5a2 2 0 00-2 2v3m18 0V5a2 2 0 00-2-2h-3m0 18h3a2 2 0 002-2v-3M3 16v3a2 2 0 002 2h3" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>
                Fullscreen
              </button>
            </div>
          </div>
        </div>

        {/* Document content */}
        <div style={{ flex: 1, paddingBottom: 80 }}>
          {/* PDF: full-width iframe, no card wrapper */}
          {isPdf && documentUrl ? (
            <div style={{ maxWidth: 1100, margin: '0 auto', padding: '20px 40px 80px' }}>
              <iframe
                src={documentUrl}
                style={{ width: '100%', height: 'calc(100vh - 220px)', border: '1px solid #e2e8f0', borderRadius: 10, background: '#fff' }}
                title={policy.PolicyName}
              />
            </div>
          ) : (
          <div style={{ maxWidth: 900, margin: '0 auto', padding: '40px 40px 80px' }}>
            <div style={{
              background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
              padding: '48px 56px', minHeight: 400, lineHeight: 1.8, fontSize: 14, color: '#334155'
            }}>
              {/* HTML content or fallback */}
              {hasContent ? (
                <div dangerouslySetInnerHTML={{ __html: bodyHtml }} />
              ) : documentUrl ? (
                <iframe
                  src={viewerUrl}
                  style={{ width: '100%', minHeight: 700, border: 'none' }}
                  title={policy.PolicyName}
                />
              ) : (
                <div style={{ textAlign: 'center', padding: 40, color: '#94a3b8' }}>
                  <p style={{ fontSize: 16, fontWeight: 600, marginBottom: 8 }}>No document content available</p>
                  <p style={{ fontSize: 13 }}>This policy has not been converted to HTML yet. Use the Admin Centre to run the document conversion.</p>
                </div>
              )}
            </div>
          </div>
          )}
        </div>

        {/* Bottom bar */}
        <div style={{
          position: 'fixed', bottom: 0, left: 0, right: 0, zIndex: 10,
          background: '#fff', borderTop: '1px solid #e2e8f0', padding: '12px 40px',
          boxShadow: '0 -2px 8px rgba(0,0,0,0.06)'
        }}>
          <div style={{ maxWidth: 1400, margin: '0 auto', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, fontSize: 12, color: '#64748b' }}>
              <div style={{ width: 8, height: 8, borderRadius: '50%', background: '#0d9488' }} />
              <span>Reading: <strong style={{ color: '#0f172a' }}>{policy.PolicyName}</strong></span>
              <span style={{ color: '#94a3b8' }}>&middot; {policy.PolicyNumber} &middot; v{policy.VersionNumber || policy.PolicyVersion || '1.0'}</span>
            </div>
            <button
              onClick={() => { window.location.href = '/sites/PolicyManager/SitePages/PolicyHub.aspx'; }}
              style={{
                display: 'flex', alignItems: 'center', gap: 6, padding: '8px 16px', borderRadius: 6,
                fontSize: 13, fontWeight: 600, color: '#0d9488', background: '#fff', border: '1px solid #e2e8f0',
                cursor: 'pointer', fontFamily: 'inherit'
              }}
            >
              <svg viewBox="0 0 24 24" fill="none" width="14" height="14"><path d="M19 12H5M12 19l-7-7 7-7" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>
              Back to Policy Hub
            </button>
          </div>
        </div>
      </div>
      </ErrorBoundary>
    );
  }

  // ============================================
  // MAIN RENDER
  // ============================================

  public render(): React.ReactElement<IPolicyDetailsProps> {
    const {
      loading, error, policy, currentFlowStep, acknowledgement,
      showCommentDialog, newComment, submittingComment, browseMode
    } = this.state;

    // ─── Browse mode: Simple Reader (from Policy Hub) ─────────────
    if (browseMode) {
      return this.renderSimpleReader();
    }

    // ─── Review mode: Reviewer decision UI ─────────────
    if ((this.state as any).reviewMode && policy && !loading) {
      return this.renderReviewMode();
    }

    // ─── Approval mode: Final approver decision UI ─────────────
    if ((this.state as any).approvalMode && policy && !loading) {
      return this.renderApprovalMode();
    }

    // Determine if this is an active read flow
    // Active flow = not browse mode AND either no acknowledgement yet OR acknowledgement is still pending
    const isActiveFlow = !acknowledgement || acknowledgement.AckStatus !== 'Acknowledged';

    // ─── VARIATION C: Focused Reader (active flow) ─────────────────
    // No nav distractions. Minimal header. Reading progress bar.
    // Floating action card. User follows guided process only.
    if (isActiveFlow && policy && !loading) {
      const flowSteps = this.getWizardSteps();
      const currentStepIndex = flowSteps.findIndex(s => s.key === currentFlowStep);
      const totalSteps = flowSteps.length;
      const stepLabel = flowSteps[currentStepIndex]?.label || 'Read';
      const scrollPct = (this.state as any).scrollProgress || 0;

      return (
        <ErrorBoundary fallbackMessage="An error occurred loading policy details. Please try again.">
        <div style={{ minHeight: '100vh', display: 'flex', flexDirection: 'column', background: '#fff' }}>
          {/* Minimal focused header — no nav items */}
          <div style={{
            background: 'linear-gradient(135deg, #0d9488 0%, #0f766e 100%)',
            padding: '10px 24px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            flexShrink: 0
          }}>
            <a href="/sites/PolicyManager/SitePages/PolicyHub.aspx" style={{ display: 'flex', alignItems: 'center', gap: 10, textDecoration: 'none', color: '#fff' }}>
              <div style={{
                width: 32, height: 32, background: 'rgba(255,255,255,0.15)',
                borderRadius: 6, display: 'flex', alignItems: 'center', justifyContent: 'center'
              }}>
                <svg viewBox="0 0 24 24" fill="none" style={{ width: 18, height: 18 }}>
                  <path d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
              </div>
              <span style={{ fontWeight: 600, fontSize: 15 }}>Policy Manager</span>
            </a>
            <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
              <span style={{ color: 'rgba(255,255,255,0.7)', fontSize: 13 }}>
                Step {currentStepIndex + 1} of {totalSteps} — {stepLabel}
              </span>
              <a href="/sites/PolicyManager/SitePages/MyPolicies.aspx" style={{
                color: 'rgba(255,255,255,0.8)', fontSize: 13, textDecoration: 'none',
                padding: '6px 14px', border: '1px solid rgba(255,255,255,0.3)', borderRadius: 4
              }}>
                Exit
              </a>
            </div>
          </div>

          {/* Reading progress bar */}
          {currentFlowStep === 'reading' && (
            <div style={{ height: 3, background: '#e2e8f0', flexShrink: 0 }}>
              <div style={{ height: '100%', background: '#0d9488', width: `${scrollPct}%`, transition: 'width 0.3s' }} />
            </div>
          )}

          {/* Compact metadata bar */}
          <div style={{
            display: 'flex', alignItems: 'center', justifyContent: 'space-between',
            padding: '10px 32px', background: '#fff', borderBottom: '1px solid #e2e8f0',
            flexShrink: 0, fontSize: 13
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
              <span style={{ fontWeight: 700, fontSize: 16, color: '#0f172a' }}>{policy.PolicyName}</span>
              <span style={{ padding: '3px 10px', borderRadius: 4, fontSize: 11, fontWeight: 600, background: '#ecfdf5', color: '#059669' }}>{policy.PolicyStatus}</span>
              <span style={{ padding: '3px 10px', borderRadius: 4, fontSize: 11, fontWeight: 600, background: '#f0fdfa', color: '#0d9488' }}>{policy.PolicyCategory}</span>
              <span style={{ padding: '3px 10px', borderRadius: 4, fontSize: 11, fontWeight: 600, background: '#f1f5f9', color: '#64748b' }}>{policy.PolicyNumber}</span>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, color: '#64748b', fontSize: 12 }}>
              <span>v{policy.VersionNumber}</span>
              {policy.EffectiveDate && <span>Effective: {new Date(policy.EffectiveDate as any).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}</span>}
            </div>
          </div>

          {error && (
            <MessageBar messageBarType={MessageBarType.error} isMultiline onDismiss={() => this.setState({ error: null })}>
              {error}
            </MessageBar>
          )}

          {/* Main content area */}
          <div style={{ flex: 1, position: 'relative', overflow: 'auto' }}>
            {currentFlowStep === 'reading' && this.renderReadStep()}
            {currentFlowStep === 'quiz' && this.renderQuizStep()}
            {currentFlowStep === 'acknowledge' && this.renderAcknowledgeStep()}
            {currentFlowStep === 'complete' && this.renderCompleteStep()}

            {/* Floating action card (bottom-right) — only during reading step */}
            {currentFlowStep === 'reading' && (
              <div style={{
                position: 'fixed', bottom: 24, right: 24, width: 300,
                background: '#fff', borderRadius: 8, boxShadow: '0 8px 32px rgba(0,0,0,0.15)',
                border: '1px solid #e2e8f0', overflow: 'hidden', zIndex: 100
              }}>
                <div style={{ background: 'linear-gradient(135deg, #0d9488, #0f766e)', padding: '14px 18px', color: '#fff' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                    <span style={{ fontSize: 13, fontWeight: 600 }}>Reading Progress</span>
                    <span style={{ fontSize: 12, opacity: 0.8 }}>Step {currentStepIndex + 1} of {totalSteps}</span>
                  </div>
                  <div style={{ display: 'flex', gap: 4 }}>
                    {flowSteps.map((s, i) => (
                      <div key={s.key} style={{
                        flex: 1, height: 4, borderRadius: 2,
                        background: i <= currentStepIndex ? 'rgba(255,255,255,0.9)' : 'rgba(255,255,255,0.2)'
                      }} />
                    ))}
                  </div>
                </div>
                <div style={{ padding: '16px 18px' }}>
                  {scrollPct >= 95 ? (
                    <div style={{ fontSize: 13, color: '#059669', marginBottom: 12, fontWeight: 500 }}>
                      Document reviewed. Click below to proceed.
                    </div>
                  ) : (
                    <div style={{ fontSize: 13, color: '#64748b', marginBottom: 12 }}>
                      Scroll to the end of the document to continue.
                    </div>
                  )}
                  <button
                    onClick={() => this.handleMarkAsRead()}
                    disabled={scrollPct < 95}
                    style={{
                      width: '100%', padding: '10px 16px', borderRadius: 4, border: 'none',
                      background: scrollPct >= 95 ? '#0d9488' : '#94a3b8',
                      color: '#fff', fontWeight: 600, fontSize: 14,
                      cursor: scrollPct >= 95 ? 'pointer' : 'not-allowed',
                      display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8
                    }}
                  >
                    I Have Read This Policy →
                  </button>
                  {this.state.readDuration > 0 && (
                    <div style={{ fontSize: 12, color: '#94a3b8', textAlign: 'center', marginTop: 8 }}>
                      Reading time: {this.formatDuration(this.state.readDuration)}
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>

          {/* Panels (acknowledgement, receipt, etc.) */}
          {this.renderAcknowledgePanel()}
          {this.renderReadReceiptPanel()}
          {this.renderVersionHistoryPanel()}
          {this.renderVersionComparisonPanel()}

          {/* Wizard footer — shown for ALL steps (including reading) */}
          {this.renderWizardFooter()}
        </div>
        </ErrorBoundary>
      );
    }

    // ─── Standard layout (browse mode / already acknowledged) ────
    return (
      <ErrorBoundary fallbackMessage="An error occurred loading policy details. Please try again.">
      <JmlAppLayout
        context={this.props.context}
        sp={this.props.sp}
        pageTitle={browseMode ? 'Policy Viewer' : 'Policy Details'}
        pageDescription={browseMode ? 'Viewing policy document' : 'View policy content, version history and acknowledgements'}
        pageIcon="Document"
        breadcrumbs={[
          { text: 'Policy Manager', url: '/sites/PolicyManager' },
          { text: browseMode ? 'Browse Policies' : 'My Policies', url: browseMode ? '/sites/PolicyManager/SitePages/PolicyHub.aspx' : '/sites/PolicyManager/SitePages/MyPolicies.aspx' },
          { text: policy ? `${policy.PolicyNumber}` : 'Policy Details' }
        ]}
        activeNavKey="browse"
        showQuickLinks={true}
        showSearch={true}
        showNotifications={true}
        compactFooter={true}
        dwxHub={this.props.dwxHub}
      >
        <section className={styles.policyDetails}>
          {loading && (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Spinner size={SpinnerSize.large} label="Loading policy..." />
            </Stack>
          )}

          {error && (
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline
              onDismiss={() => this.setState({ error: null })}
              dismissButtonAriaLabel="Close"
            >
              {error}
            </MessageBar>
          )}

          {!loading && !error && policy && (
            <>
              {/* Render current flow step */}
              {currentFlowStep === 'reading' && this.renderReadStep()}
              {currentFlowStep === 'quiz' && this.renderQuizStep()}
              {currentFlowStep === 'acknowledge' && this.renderAcknowledgeStep()}
              {currentFlowStep === 'complete' && this.renderCompleteStep()}

              {/* Panels */}
              {this.renderAcknowledgePanel()}
              {this.renderReadReceiptPanel()}
              {this.renderVersionHistoryPanel()}
              {this.renderVersionComparisonPanel()}
            </>
          )}

          {/* Comment Dialog */}
          <Dialog
            hidden={!showCommentDialog}
            onDismiss={() => this.setState({ showCommentDialog: false })}
            dialogContentProps={{ type: DialogType.normal, title: 'Add Comment' }}
          >
            <TextField
              multiline rows={4}
              value={newComment}
              onChange={(e, value) => this.setState({ newComment: value || '' })}
              placeholder="Share your thoughts about this policy..."
            />
            <DialogFooter>
              <PrimaryButton text="Submit" onClick={this.handleSubmitComment} disabled={!newComment.trim() || submittingComment} />
              <DefaultButton text="Cancel" onClick={() => this.setState({ showCommentDialog: false })} disabled={submittingComment} />
            </DialogFooter>
          </Dialog>
          <this.dialogManager.DialogComponent />
        </section>
      </JmlAppLayout>
      </ErrorBoundary>
    );
  }
}
