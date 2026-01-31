// @ts-nocheck
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
  Icon,
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
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { PageSubheader } from '../../../components/PageSubheader';
import { PolicyService } from '../../../services/PolicyService';
import { PolicySocialService } from '../../../services/PolicySocialService';
import { createDialogManager } from '../../../hooks/useDialog';
import {
  IPolicy,
  IPolicyAcknowledgement,
  IPolicyRating,
  IPolicyComment,
  IPolicyAcknowledgeRequest
} from '../../../models/IPolicy';
import styles from './PolicyDetails.module.scss';
import { PM_LISTS } from '../../../constants/SharePointListNames';
import { QuizService, IQuizResult } from '../../../services/QuizService';
import { QuizTaker } from '../../../components/QuizTaker/QuizTaker';

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
  quizSubmitted: boolean;
  // Browse mode — read-only viewing from Policy Hub (no wizard/acknowledge flow)
  browseMode: boolean;
  // Live quiz integration
  liveQuizId: number | null;
  currentUserId: number;
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
  private policyService: PolicyService;
  private socialService: PolicySocialService;
  private readTimer: NodeJS.Timeout | null = null;
  private dialogManager = createDialogManager();
  private documentViewerRef: React.RefObject<HTMLDivElement>;

  constructor(props: IPolicyDetailsProps) {
    super(props);
    this.documentViewerRef = React.createRef();
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
      quizSubmitted: false,
      // Browse mode detection — from Policy Hub browsing
      browseMode: this.getBrowseModeFromUrl(),
      // Live quiz integration
      liveQuizId: null,
      currentUserId: 0
    };
    this.policyService = new PolicyService(props.sp);
    this.socialService = new PolicySocialService(props.sp);
  }

  public async componentDidMount(): Promise<void> {
    injectPortalStyles();
    await this.loadPolicyDetails();
    this.startReadTracking();
  }

  public componentWillUnmount(): void {
    this.stopReadTracking();
  }

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
    const { policyId } = this.state;
    if (!policyId) {
      this.setState({ error: 'No policy ID provided', loading: false });
      return;
    }

    try {
      this.setState({ loading: true, error: null });
      await this.policyService.initialize();
      await this.socialService.initialize();

      const policy = await this.policyService.getPolicyById(policyId);
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

      this.setState({
        policy,
        acknowledgement,
        ratings,
        comments,
        isFollowing,
        liveQuizId,
        currentUserId: currentUser.Id,
        loading: false
      });

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

    if (!policy || !acknowledgement) return;
    if (!this.canSubmitAcknowledgement()) {
      this.setState({ error: 'Please complete all acknowledgement requirements.' });
      return;
    }

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

      const readReceipt: IReadReceipt = {
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

      // Try to save to SharePoint — but still advance to complete if lists aren't provisioned yet
      try {
        await this.saveReadReceipt(readReceipt);
      } catch (saveErr) {
        console.warn('Could not save read receipt (list may not exist yet):', saveErr);
      }

      try {
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
      } catch (ackErr) {
        console.warn('Could not update acknowledgement record (list may not exist yet):', ackErr);
      }

      this.setState({
        readReceipt,
        showAcknowledgePanel: false,
        showCongratulationsPanel: true,
        currentFlowStep: 'complete',
        submittingAcknowledgement: false
      });
    } catch (error) {
      console.error('Failed to submit acknowledgement:', error);
      this.setState({
        error: 'Failed to submit acknowledgement. Please try again.',
        submittingAcknowledgement: false
      });
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
    return `
      <html>
      <body style="font-family: 'Segoe UI', Tahoma, Geneva, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
        <div style="background: linear-gradient(135deg, #0f4c47, #0f766e); padding: 30px; text-align: center; border-radius: 8px 8px 0 0;">
          <h1 style="color: white; margin: 0;">Policy Read Receipt</h1>
          <p style="color: rgba(255,255,255,0.9); margin: 10px 0 0 0;">Policy Manager</p>
        </div>
        <div style="background: white; padding: 30px; border: 1px solid #e1e1e1;">
          <div style="text-align: center; margin-bottom: 20px;">
            <div style="display: inline-block; background: #16a34a; color: white; padding: 8px 16px; border-radius: 20px; font-weight: 600;">
              &#10003; Acknowledged
            </div>
          </div>
          <table style="width: 100%; border-collapse: collapse;">
            <tr><td style="padding: 8px 0; color: #666;">Receipt Number:</td><td style="padding: 8px 0; font-weight: 600;">${receipt.ReceiptNumber}</td></tr>
            <tr><td style="padding: 8px 0; color: #666;">Employee:</td><td style="padding: 8px 0;">${receipt.UserDisplayName}</td></tr>
            <tr><td style="padding: 8px 0; color: #666;">Email:</td><td style="padding: 8px 0;">${receipt.UserEmail}</td></tr>
            <tr><td style="padding: 8px 0; color: #666;">Policy:</td><td style="padding: 8px 0; font-weight: 600;">${receipt.PolicyNumber} - ${receipt.PolicyName}</td></tr>
            <tr><td style="padding: 8px 0; color: #666;">Version:</td><td style="padding: 8px 0;">${receipt.PolicyVersion}</td></tr>
            <tr><td style="padding: 8px 0; color: #666;">Read Duration:</td><td style="padding: 8px 0;">${Math.floor(receipt.ReadDurationSeconds / 60)} min ${receipt.ReadDurationSeconds % 60} sec</td></tr>
            ${receipt.QuizRequired ? `<tr><td style="padding: 8px 0; color: #666;">Quiz Score:</td><td style="padding: 8px 0;">${receipt.QuizScore}%</td></tr>` : ''}
            <tr><td style="padding: 8px 0; color: #666;">Acknowledged:</td><td style="padding: 8px 0;">${receipt.AcknowledgedDate.toLocaleDateString()} at ${receipt.AcknowledgedTime}</td></tr>
            <tr><td style="padding: 8px 0; color: #666;">Digital Signature:</td><td style="padding: 8px 0; font-style: italic;">${receipt.DigitalSignature}</td></tr>
          </table>
          <div style="margin-top: 20px; padding: 15px; background: #f5f5f5; border-radius: 4px; font-size: 12px; color: #666;">
            <p style="margin: 0;"><strong>Legal Confirmation:</strong></p>
            <p style="margin: 10px 0 0 0; white-space: pre-line;">${receipt.LegalConfirmationText}</p>
          </div>
        </div>
        <div style="background: #f5f5f5; padding: 15px; text-align: center; font-size: 12px; color: #666; border-radius: 0 0 8px 8px;">
          <p style="margin: 0;">This is an automated receipt from Policy Manager. Please retain for your records.</p>
        </div>
      </body>
      </html>
    `;
  }

  private handleViewReceipt = (): void => {
    this.setState({ showReadReceiptPanel: true });
  };

  private handleGeneratePdf = (): void => {
    const { readReceipt } = this.state;
    if (!readReceipt) return;
    const printWindow = window.open('', '_blank');
    if (printWindow) {
      printWindow.document.write(this.generateReceiptEmailHtml(readReceipt));
      printWindow.document.close();
      printWindow.print();
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
    const newAnswers = [...this.state.quizAnswers];
    newAnswers[questionIndex] = optionIndex;
    this.setState({ quizAnswers: newAnswers });

    // Auto-advance after short delay
    if (this.state.currentQuizQuestion < MOCK_QUIZ_QUESTIONS.length - 1) {
      setTimeout(() => {
        this.setState(prev => ({
          currentQuizQuestion: prev.currentQuizQuestion + 1
        }));
      }, 400);
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
      return `${siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(documentUrl)}&action=view`;
    }
    if (ext === 'pdf') return documentUrl;
    if (['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp'].includes(ext)) return documentUrl;
    return `${siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(documentUrl)}&action=view`;
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

    return (
      <div className={styles.stepContent}>
        {/* Policy Metadata Card */}
        <div className={styles.wizardCard}>
          <div className={styles.cardHeader}>
            <div className={styles.cardIcon}>
              <Icon iconName="Document" styles={{ root: { fontSize: 18 } }} />
            </div>
            <div className={styles.cardHeaderInfo}>
              <Text variant="large" style={{ fontWeight: 700, color: '#0f172a' }}>
                {policy.PolicyNumber} - {policy.PolicyName}
              </Text>
              <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginTop: 4 }}>
                <span className={styles.badgeGreen}>Published</span>
                <span className={styles.badgeTeal}>{policy.PolicyCategory}</span>
                <span className={styles.badgeSlate}>{policy.PolicyNumber}</span>
                {acknowledgement && (
                  <span className={styles.badgeAmber} style={{ backgroundColor: statusColor === '#16a34a' ? '#dcfce7' : statusColor === '#dc2626' ? '#fee2e2' : '#fef3c7', color: statusColor }}>
                    {acknowledgement.AckStatus || acknowledgement.Status}
                  </span>
                )}
              </Stack>
            </div>
            <Stack horizontal tokens={{ childrenGap: 8 }} className={styles.cardActions}>
              <div className={styles.readTimer}>
                <Icon iconName="Timer" styles={{ root: { fontSize: 14 } }} />
                <span>{this.formatDuration(readDuration)}</span>
              </div>
              <IconButton
                iconProps={{ iconName: isFollowing ? 'FavoriteStarFill' : 'FavoriteStar' }}
                title={isFollowing ? 'Unfollow' : 'Follow'}
                onClick={this.handleFollow}
              />
              <IconButton
                iconProps={{ iconName: 'Share' }}
                title="Share"
                onClick={this.handleShare}
              />
            </Stack>
          </div>

          <div className={styles.policyMeta}>
            <div className={styles.metaItem}>
              <span className={styles.metaLabel}>Department</span>
              <span className={styles.metaValue}>{policy.PolicyCategory || 'General'}</span>
            </div>
            <div className={styles.metaItem}>
              <span className={styles.metaLabel}>Effective Date</span>
              <span className={styles.metaValue}>{policy.EffectiveDate ? new Date(policy.EffectiveDate).toLocaleDateString() : 'N/A'}</span>
            </div>
            <div className={styles.metaItem}>
              <span className={styles.metaLabel}>Version</span>
              <span className={styles.metaValue}>v{policy.VersionNumber || '1.0'}</span>
            </div>
            {acknowledgement?.DueDate && (
              <div className={styles.metaItem}>
                <span className={styles.metaLabel}>Ack. Due</span>
                <span className={styles.metaValue} style={{ color: statusColor }}>{new Date(acknowledgement.DueDate).toLocaleDateString()}</span>
              </div>
            )}
          </div>
        </div>

        {/* Document Viewer */}
        {hasDocuments && documentUrl && (
          <div className={styles.documentViewerWrapper}>
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
            <div
              className={styles.documentViewer}
              ref={this.documentViewerRef}
              onScroll={this.handleDocumentScroll}
            >
              <div className={styles.scrollProgressBar}>
                <div className={styles.scrollProgressFill} style={{ height: `${scrollProgress}%` }} />
              </div>
              {isImage ? (
                <div style={{ textAlign: 'center', padding: 20 }}>
                  <img src={viewerUrl} alt={policy.Title} style={{ maxWidth: '100%', maxHeight: 600, borderRadius: 4 }} />
                </div>
              ) : (
                <iframe src={viewerUrl} style={{ width: '100%', height: '100%', border: 'none' }} title={`${policy.Title} Document Viewer`} />
              )}
            </div>
            <div className={styles.scrollNotice}>
              {scrollProgress >= 95 || readDuration >= 30 ? (
                <span style={{ color: '#16a34a', fontWeight: 600 }}>
                  <Icon iconName="CheckMark" /> Document read complete — you may now proceed
                </span>
              ) : (
                <span>Please review the document before proceeding ({Math.max(0, 30 - readDuration)}s remaining)</span>
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
            <div className={styles.scrollNotice}>
              {scrollProgress >= 95 ? (
                <span style={{ color: '#16a34a', fontWeight: 600 }}>
                  <Icon iconName="CheckMark" /> Document read complete — you may now proceed
                </span>
              ) : (
                <span>Please scroll through the entire document before proceeding</span>
              )}
            </div>
          </div>
        )}

        {/* Policy Content (HTML/text fallback) */}
        {policy.PolicyContent && (
          <div className={styles.wizardCard}>
            <Text variant="large" style={{ fontWeight: 600, color: '#0d9488', marginBottom: 12, display: 'block' }}>
              Policy Overview
            </Text>
            {policy.PolicySummary && <Text style={{ marginBottom: 12, display: 'block' }}>{policy.PolicySummary}</Text>}
            <div dangerouslySetInnerHTML={{ __html: policy.PolicyContent }} />
          </div>
        )}

        {/* Attachments */}
        {attachments.length > 0 && (
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
    const allAnswered = this.allQuizAnswered();

    return (
      <div className={styles.stepContent}>
        {/* Quiz Banner */}
        <div className={styles.quizBanner}>
          <Icon iconName="Questionnaire" styles={{ root: { fontSize: 18 } }} />
          <span>This policy requires a comprehension quiz. You must score at least <strong>{policy.QuizPassingScore || 80}%</strong> to proceed.</span>
        </div>

        <div className={styles.wizardCard}>
          <div className={styles.cardHeader}>
            <div className={styles.cardIcon}>
              <Icon iconName="Questionnaire" styles={{ root: { fontSize: 18 } }} />
            </div>
            <h2 style={{ fontSize: 17, fontWeight: 700, color: '#0f172a', margin: 0 }}>
              {policy.PolicyName} — Comprehension Quiz
            </h2>
          </div>

          {/* Question dots */}
          <div className={styles.questionDots}>
            {questions.map((_, i) => (
              <div
                key={i}
                className={`${styles.qDot} ${quizAnswers[i] >= 0 ? styles.answered : ''} ${i === currentQuizQuestion && !quizSubmitted ? styles.current : ''}`}
                onClick={() => !quizSubmitted && this.setState({ currentQuizQuestion: i })}
              >
                {i + 1}
              </div>
            ))}
          </div>

          {/* Questions viewport */}
          {!quizSubmitted ? (
            <>
              <div className={styles.questionsViewport}>
                {questions.map((q, qi) => (
                  <div
                    key={qi}
                    className={`${styles.questionCard} ${qi === currentQuizQuestion ? styles.visible : ''} ${quizAnswers[qi] >= 0 ? styles.answered : ''}`}
                  >
                    <div className={styles.questionNum}>Question {qi + 1} of {questions.length}</div>
                    <div className={styles.questionText}>{q.question}</div>
                    <div className={styles.optionGroup}>
                      {q.options.map((opt, oi) => (
                        <label
                          key={oi}
                          className={`${styles.optionLabel} ${quizAnswers[qi] === oi ? styles.selected : ''}`}
                          onClick={() => this.handleQuizSelectAnswer(qi, oi)}
                        >
                          <input type="radio" name={`q${qi}`} checked={quizAnswers[qi] === oi} readOnly />
                          <span>{opt}</span>
                        </label>
                      ))}
                    </div>
                  </div>
                ))}
              </div>

              {/* Nav row */}
              <div className={styles.quizNavRow}>
                <DefaultButton
                  text="Previous"
                  iconProps={{ iconName: 'ChevronLeft' }}
                  disabled={currentQuizQuestion === 0}
                  onClick={() => this.setState({ currentQuizQuestion: currentQuizQuestion - 1 })}
                  styles={{ root: { height: 32 }, label: { fontSize: 12 } }}
                />
                <Text variant="small" style={{ color: '#64748b', fontWeight: 600 }}>
                  Question {currentQuizQuestion + 1} of {questions.length}
                </Text>
                <DefaultButton
                  text="Next"
                  iconProps={{ iconName: 'ChevronRight' }}
                  iconPosition="after"
                  disabled={currentQuizQuestion === questions.length - 1}
                  onClick={() => this.setState({ currentQuizQuestion: currentQuizQuestion + 1 })}
                  styles={{ root: { height: 32 }, label: { fontSize: 12 } }}
                />
              </div>

              {/* Submit */}
              <div style={{ textAlign: 'center', marginTop: 20 }}>
                <PrimaryButton
                  text="Submit Quiz"
                  iconProps={{ iconName: 'Accept' }}
                  disabled={!allAnswered}
                  onClick={this.handleQuizSubmit}
                />
              </div>
            </>
          ) : (
            /* Post-submission result */
            <div className={`${styles.quizResult} ${quizPassed ? styles.passed : styles.failed}`}>
              <Text variant="xxLarge" style={{ fontWeight: 800 }}>{quizScore}%</Text>
              <Text variant="medium" style={{ marginTop: 4 }}>
                {Math.round(quizScore / (100 / questions.length))}/{questions.length} correct — {quizPassed ? 'PASSED' : `Required: ${policy.QuizPassingScore || 80}%`}
              </Text>
              {quizPassed && (
                <Text variant="medium" style={{ marginTop: 8, color: '#16a34a' }}>
                  <Icon iconName="CheckMark" /> You may now proceed to acknowledge the policy.
                </Text>
              )}
              {!quizPassed && (
                <DefaultButton
                  text="Retake Quiz"
                  iconProps={{ iconName: 'Refresh' }}
                  onClick={this.handleQuizRetake}
                  styles={{ root: { marginTop: 16 } }}
                />
              )}
            </div>
          )}
        </div>
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
              onClick={() => window.location.href = '/sites/PolicyManager/SitePages/PolicyHub.aspx'}
            />
          </div>
        </div>
      </div>
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
        headerText="Policy Acknowledgement"
        closeButtonAriaLabel="Close"
        isFooterAtBottom={true}
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
      <Panel
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
      </Panel>
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
              disabled={currentIndex === 0}
              onClick={this.handleWizardBack}
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
  // MAIN RENDER
  // ============================================

  public render(): React.ReactElement<IPolicyDetailsProps> {
    const {
      loading, error, policy, currentFlowStep, acknowledgement,
      showCommentDialog, newComment, submittingComment, browseMode
    } = this.state;

    // Determine if this is an active read flow
    // Browse mode (from Policy Hub) always shows read-only view
    // Active flow = not browse mode AND either no acknowledgement yet OR acknowledgement is still pending
    const isActiveFlow = !browseMode && (!acknowledgement || acknowledgement.AckStatus !== 'Acknowledged');

    return (
      <JmlAppLayout
        context={this.props.context}
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
              {/* Wizard Progress Stepper */}
              {isActiveFlow && this.renderWizardStepper()}

              {/* Active wizard flow — step-driven content */}
              {isActiveFlow && currentFlowStep === 'reading' && this.renderReadStep()}
              {isActiveFlow && currentFlowStep === 'quiz' && this.renderQuizStep()}
              {isActiveFlow && currentFlowStep === 'acknowledge' && this.renderAcknowledgeStep()}
              {isActiveFlow && currentFlowStep === 'complete' && this.renderCompleteStep()}

              {/* Non-active flow (browse mode or already acknowledged) — read-only view */}
              {!isActiveFlow && this.renderReadStep()}

              {/* Panels */}
              {this.renderAcknowledgePanel()}
              {this.renderReadReceiptPanel()}

              {/* Sticky Footer */}
              {isActiveFlow && this.renderWizardFooter()}
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
    );
  }
}
