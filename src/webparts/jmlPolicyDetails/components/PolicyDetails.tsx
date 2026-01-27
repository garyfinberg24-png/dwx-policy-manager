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
}

export default class PolicyDetails extends React.Component<IPolicyDetailsProps, IPolicyDetailsState> {
  private policyService: PolicyService;
  private socialService: PolicySocialService;
  private readTimer: NodeJS.Timeout | null = null;
  private dialogManager = createDialogManager();

  constructor(props: IPolicyDetailsProps) {
    super(props);
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
      digitalSignature: ''
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

      // Get current user ID
      const currentUser = await this.props.sp.web.currentUser();

      // Get user's acknowledgement record
      const dashboard = await this.policyService.getUserDashboard(currentUser.Id);
      const acknowledgement = dashboard.pendingAcknowledgements.find(
        (ack: IPolicyAcknowledgement) => ack.PolicyId === policyId
      ) || dashboard.completedAcknowledgements.find(
        (ack: IPolicyAcknowledgement) => ack.PolicyId === policyId
      );

      // Get ratings and comments
      const ratings = await this.socialService.getPolicyRatings(policyId);
      const comments = await this.socialService.getPolicyComments(policyId);
      const isFollowing = await this.socialService.isFollowingPolicy(policyId);

      this.setState({
        policy,
        acknowledgement,
        ratings,
        comments,
        isFollowing,
        loading: false
      });

      // Track policy opened
      if (acknowledgement && acknowledgement.Status !== 'Acknowledged') {
        await this.policyService.trackPolicyOpen(acknowledgement.Id);
      }
    } catch (error) {
      console.error('Failed to load policy details:', error);
      this.setState({
        error: 'Failed to load policy details. Please try again later.',
        loading: false
      });
    }
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

  private handleAcknowledge = (): void => {
    this.setState({ showAcknowledgeDialog: true });
  };

  private handleAcknowledgeSubmit = async (): Promise<void> => {
    const { policy, acknowledgeConfirmation, acknowledgeNotes, readDuration, acknowledgement } = this.state;

    if (!acknowledgeConfirmation) {
      await this.dialogManager.showAlert('You must confirm that you have read and understood this policy.', { variant: 'warning' });
      return;
    }

    if (!acknowledgement) {
      this.setState({ error: 'No acknowledgement record found' });
      return;
    }

    try {
      this.setState({ submittingAcknowledgement: true });

      const request: IPolicyAcknowledgeRequest = {
        acknowledgementId: acknowledgement.Id,
        acknowledgedDate: new Date(),
        notes: acknowledgeNotes,
        readDuration: readDuration,
        ipAddress: '', // Browser doesn't have direct access
        userAgent: navigator.userAgent,
        quizScore: undefined // Would come from quiz component
      };

      await this.policyService.acknowledgePolicy(request);

      this.setState({
        showAcknowledgeDialog: false,
        submittingAcknowledgement: false,
        acknowledgeConfirmation: false,
        acknowledgeNotes: ''
      });

      // Reload to show updated status
      await this.loadPolicyDetails();

      await this.dialogManager.showAlert('Thank you for acknowledging this policy!', { variant: 'success' });
    } catch (error) {
      console.error('Failed to acknowledge policy:', error);
      this.setState({
        error: 'Failed to acknowledge policy. Please try again.',
        submittingAcknowledgement: false
      });
    }
  };

  private handleRate = async (rating: number): Promise<void> => {
    const { policy } = this.state;
    if (!policy) return;

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

      // Reload ratings
      const ratings = await this.socialService.getPolicyRatings(policy.Id);
      this.setState({
        ratings,
        submittingRating: false,
        reviewTitle: '',
        reviewText: ''
      });

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

      await this.socialService.commentOnPolicy({
        policyId: policy.Id,
        commentText: newComment
      });

      // Reload comments
      const comments = await this.socialService.getPolicyComments(policy.Id);
      this.setState({
        comments,
        newComment: '',
        showCommentDialog: false,
        submittingComment: false
      });
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
        await navigator.share({
          title: policy.PolicyName,
          text: `Check out this policy: ${policy.PolicyNumber}`,
          url: url
        });
      } catch (error) {
        // User cancelled share
      }
    } else {
      // Fallback: Copy to clipboard
      await navigator.clipboard.writeText(url);
      await this.dialogManager.showAlert('Link copied to clipboard!', { variant: 'success' });
    }
  };

  // ============================================
  // ENHANCED READ FLOW METHODS
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

    // Check if quiz is required
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
      // Proceed to acknowledgement
      this.setState({
        currentFlowStep: 'acknowledge',
        showAcknowledgePanel: true
      });
    } else {
      // Quiz failed - user needs to retake or policy owner decides action
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

      const currentUser = await this.props.sp.web.currentUser();
      const now = new Date();

      // Legal confirmation text
      const legalText = `I, ${digitalSignature}, hereby confirm that:
1. I have read and fully understood the policy "${policy.PolicyName}" (${policy.PolicyNumber}).
2. I agree to comply with all requirements and guidelines outlined in this policy.
3. I understand that failure to comply may result in disciplinary action.
4. I acknowledge that this constitutes my electronic signature and consent.`;

      // Create read receipt for audit trail
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

      // Save read receipt to SharePoint list
      await this.saveReadReceipt(readReceipt);

      // Update acknowledgement record
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
      await this.props.sp.web.lists.getByTitle('JML_PolicyReadReceipts').items.add({
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
      // Don't throw - we still want to complete the acknowledgement
    }
  };

  private handleEmailReceipt = async (): Promise<void> => {
    const { readReceipt, policy } = this.state;
    if (!readReceipt || !policy) return;

    try {
      this.setState({ emailingReceipt: true });

      // Generate email content
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

      // Open default email client with pre-filled content
      window.open(`mailto:${readReceipt.UserEmail}?subject=${subject}&body=${body}`, '_blank');

      this.setState({ emailingReceipt: false });
      await this.dialogManager.showAlert('Your email client has been opened with the receipt details. Please send the email to save a copy.', { variant: 'info' });
    } catch (error) {
      console.error('Failed to open email client:', error);
      this.setState({
        emailingReceipt: false,
        error: 'Failed to open email client. Please try again.'
      });
    }
  };

  private generateReceiptEmailHtml(receipt: IReadReceipt): string {
    return `
      <html>
      <body style="font-family: 'Segoe UI', Tahoma, Geneva, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
        <div style="background: linear-gradient(135deg, #004578, #0078d4); padding: 30px; text-align: center; border-radius: 8px 8px 0 0;">
          <h1 style="color: white; margin: 0;">Policy Read Receipt</h1>
          <p style="color: rgba(255,255,255,0.9); margin: 10px 0 0 0;">First Digital - JML Portal</p>
        </div>
        <div style="background: white; padding: 30px; border: 1px solid #e1e1e1;">
          <div style="text-align: center; margin-bottom: 20px;">
            <div style="display: inline-block; background: #107c10; color: white; padding: 8px 16px; border-radius: 20px; font-weight: 600;">
              ✓ Acknowledged
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
          <p style="margin: 0;">This is an automated receipt from the JML Portal. Please retain for your records.</p>
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

    // Generate PDF-like printable view
    const printWindow = window.open('', '_blank');
    if (printWindow) {
      printWindow.document.write(this.generateReceiptEmailHtml(readReceipt));
      printWindow.document.close();
      printWindow.print();
    }
  };

  private handleCloseCongratulations = (): void => {
    this.setState({ showCongratulationsPanel: false });
    // Reload to show updated status
    this.loadPolicyDetails();
  };

  private renderPolicyHeader(): JSX.Element | null {
    const { policy, acknowledgement, isFollowing } = this.state;
    if (!policy) return null;

    const statusColor = acknowledgement?.Status === 'Acknowledged' ? '#107C10' :
                       acknowledgement?.Status === 'Overdue' ? '#D13438' : '#FFA500';

    return (
      <div className={styles.policyHeader}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
            <Stack tokens={{ childrenGap: 8 }}>
              <Text variant="xxLarge" className={styles.policyTitle}>
                {policy.PolicyNumber} - {policy.PolicyName}
              </Text>
              <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
                <Text variant="small" className={styles.category}>
                  {policy.PolicyCategory}
                </Text>
                <Text variant="small">Version {policy.VersionNumber}</Text>
                {acknowledgement && (
                  <div className={styles.statusBadge} style={{ backgroundColor: statusColor }}>
                    {acknowledgement.Status}
                  </div>
                )}
              </Stack>
            </Stack>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
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
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 24 }}>
            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="small" className={styles.label}>Effective Date</Text>
              <Text variant="medium">{policy.EffectiveDate ? new Date(policy.EffectiveDate).toLocaleDateString() : 'N/A'}</Text>
            </Stack>
            {policy.ExpiryDate && (
              <Stack tokens={{ childrenGap: 4 }}>
                <Text variant="small" className={styles.label}>Expiry Date</Text>
                <Text variant="medium">{new Date(policy.ExpiryDate).toLocaleDateString()}</Text>
              </Stack>
            )}
            {acknowledgement?.DueDate && (
              <Stack tokens={{ childrenGap: 4 }}>
                <Text variant="small" className={styles.label}>Acknowledgement Due</Text>
                <Text variant="medium" style={{ color: statusColor }}>
                  {new Date(acknowledgement.DueDate).toLocaleDateString()}
                </Text>
              </Stack>
            )}
          </Stack>
        </Stack>
      </div>
    );
  }

  private renderPolicyContent(): JSX.Element | null {
    const { policy } = this.state;
    if (!policy) return null;

    return (
      <div className={styles.policyContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Stack tokens={{ childrenGap: 8 }}>
            <Text variant="large" className={styles.sectionTitle}>Policy Overview</Text>
            <Text>{policy.PolicySummary}</Text>
          </Stack>

          <Stack tokens={{ childrenGap: 8 }}>
            <Text variant="large" className={styles.sectionTitle}>Policy Details</Text>
            <div dangerouslySetInnerHTML={{ __html: policy.PolicyContent || '' }} />
          </Stack>

          {policy.KeyPoints && policy.KeyPoints.length > 0 && (
            <Stack tokens={{ childrenGap: 8 }}>
              <Text variant="large" className={styles.sectionTitle}>Key Points</Text>
              <ul className={styles.keyPoints}>
                {policy.KeyPoints.map((point: string, index: number) => (
                  <li key={index}>{point}</li>
                ))}
              </ul>
            </Stack>
          )}
        </Stack>
      </div>
    );
  }

  private renderAcknowledgement(): JSX.Element | null {
    const { policy, acknowledgement } = this.state;
    if (!policy || !acknowledgement || acknowledgement.Status === 'Acknowledged') return null;

    return (
      <div className={styles.acknowledgementSection}>
        <MessageBar messageBarType={MessageBarType.info}>
          You must acknowledge that you have read and understood this policy.
        </MessageBar>
        <PrimaryButton
          text="Acknowledge Policy"
          iconProps={{ iconName: 'Accept' }}
          onClick={this.handleAcknowledge}
          styles={{ root: { marginTop: 12 } }}
        />
      </div>
    );
  }

  private renderRatings(): JSX.Element | null {
    const { policy, ratings, userRating, reviewTitle, reviewText, submittingRating } = this.state;
    const { showRatings } = this.props;
    if (!showRatings || !policy) return null;

    const averageRating = policy.AverageRating || 0;
    const ratingCount = policy.RatingCount || 0;

    return (
      <div className={styles.ratingsSection}>
        <Text variant="xLarge" className={styles.sectionTitle}>Ratings & Reviews</Text>

        <Stack tokens={{ childrenGap: 16 }}>
          <Stack horizontal tokens={{ childrenGap: 24 }} verticalAlign="center">
            <Stack tokens={{ childrenGap: 4 }} horizontalAlign="center">
              <Text variant="xxLarge" style={{ fontWeight: 600 }}>{averageRating.toFixed(1)}</Text>
              <Rating
                rating={averageRating}
                size={RatingSize.Large}
                readOnly
              />
              <Text variant="small">{ratingCount} reviews</Text>
            </Stack>
          </Stack>

          <Separator />

          <Stack tokens={{ childrenGap: 12 }}>
            <Text variant="large">Rate this policy</Text>
            <Rating
              rating={userRating}
              size={RatingSize.Large}
              onChange={(ev, rating) => this.handleRate(rating || 0)}
            />
            {userRating > 0 && (
              <>
                <TextField
                  label="Review Title (Optional)"
                  value={reviewTitle}
                  onChange={(e, value) => this.setState({ reviewTitle: value || '' })}
                />
                <TextField
                  label="Review (Optional)"
                  multiline
                  rows={3}
                  value={reviewText}
                  onChange={(e, value) => this.setState({ reviewText: value || '' })}
                />
                <DefaultButton
                  text="Submit Rating"
                  onClick={this.handleSubmitRating}
                  disabled={submittingRating}
                />
              </>
            )}
          </Stack>
        </Stack>
      </div>
    );
  }

  private renderComments(): JSX.Element | null {
    const { comments } = this.state;
    const { showComments } = this.props;
    if (!showComments) return null;

    return (
      <div className={styles.commentsSection}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="xLarge" className={styles.sectionTitle}>Comments ({comments.length})</Text>
          <DefaultButton
            text="Add Comment"
            iconProps={{ iconName: 'Comment' }}
            onClick={this.handleComment}
          />
        </Stack>

        <Stack tokens={{ childrenGap: 16 }} styles={{ root: { marginTop: 16 } }}>
          {comments.map((comment: IPolicyComment) => (
            <div key={comment.Id} className={styles.commentCard}>
              <Stack tokens={{ childrenGap: 8 }}>
                <Stack horizontal horizontalAlign="space-between">
                  <Text variant="medium" style={{ fontWeight: 600 }}>{comment.UserEmail}</Text>
                  <Text variant="small">{new Date(comment.CommentDate).toLocaleDateString()}</Text>
                </Stack>
                <Text>{comment.CommentText}</Text>
                <Stack horizontal tokens={{ childrenGap: 16 }}>
                  <IconButton
                    iconProps={{ iconName: 'Like' }}
                    title="Like"
                    onClick={() => this.socialService.likeComment(comment.Id)}
                  />
                  <Text variant="small">{comment.LikeCount} likes</Text>
                </Stack>
              </Stack>
            </div>
          ))}
        </Stack>
      </div>
    );
  }

  // ============================================
  // ENHANCED READ FLOW RENDER METHODS
  // ============================================

  private renderReadFlowProgress(): JSX.Element {
    const { currentFlowStep, quizRequired, hasReadPolicy, quizCompleted } = this.state;

    const steps = [
      { key: 'reading', label: 'Read Policy', icon: 'Read', completed: hasReadPolicy },
      ...(quizRequired ? [{ key: 'quiz', label: 'Complete Quiz', icon: 'Questionnaire', completed: quizCompleted }] : []),
      { key: 'acknowledge', label: 'Acknowledge', icon: 'Handwriting', completed: currentFlowStep === 'complete' },
      { key: 'complete', label: 'Complete', icon: 'CheckMark', completed: currentFlowStep === 'complete' }
    ];

    const currentIndex = steps.findIndex(s => s.key === currentFlowStep);
    const progress = currentFlowStep === 'complete' ? 1 : (currentIndex / (steps.length - 1));

    return (
      <div className={styles.readFlowProgress}>
        <ProgressIndicator
          label="Policy Reading Progress"
          description={`Step ${currentIndex + 1} of ${steps.length}: ${steps[currentIndex]?.label || 'Complete'}`}
          percentComplete={progress}
        />
        <Stack horizontal tokens={{ childrenGap: 16 }} horizontalAlign="center" styles={{ root: { marginTop: 16 } }}>
          {steps.map((step, index) => (
            <Stack key={step.key} horizontalAlign="center" tokens={{ childrenGap: 4 }}>
              <div
                className={`${styles.flowStepIcon} ${step.completed ? styles.completed : ''} ${currentFlowStep === step.key ? styles.current : ''}`}
              >
                <Icon iconName={step.completed ? 'CheckMark' : step.icon} />
              </div>
              <Text variant="small" style={{ fontWeight: currentFlowStep === step.key ? 600 : 400 }}>
                {step.label}
              </Text>
            </Stack>
          ))}
        </Stack>
      </div>
    );
  }

  private renderQuizSection(): JSX.Element | null {
    const { currentFlowStep, policy, quizCompleted, quizScore, quizPassed } = this.state;
    if (currentFlowStep !== 'quiz' || !policy?.RequiresQuiz) return null;

    return (
      <div className={styles.quizSection}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Text variant="xLarge" className={styles.sectionTitle}>
            <Icon iconName="Questionnaire" style={{ marginRight: 8 }} />
            Policy Comprehension Quiz
          </Text>

          <MessageBar messageBarType={MessageBarType.info}>
            You must complete this quiz and achieve a passing score of {policy.QuizPassingScore || 80}% before you can acknowledge this policy.
          </MessageBar>

          {!quizCompleted ? (
            <Stack tokens={{ childrenGap: 16 }}>
              <Text>
                This quiz will test your understanding of the key concepts in the policy.
                Please read the policy carefully before attempting the quiz.
              </Text>

              {/* Quiz would be embedded here - for now, a placeholder with demo functionality */}
              <div className={styles.quizPlaceholder}>
                <Icon iconName="Questionnaire" style={{ fontSize: 48, color: '#0078d4' }} />
                <Text variant="large" style={{ marginTop: 16 }}>Quiz Builder Integration</Text>
                <Text style={{ color: '#605e5c', marginTop: 8 }}>
                  The quiz for this policy will be displayed here.
                </Text>

                {/* Demo buttons for testing the flow */}
                <Stack horizontal tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 24 } }}>
                  <PrimaryButton
                    text="Complete Quiz (Pass)"
                    iconProps={{ iconName: 'Accept' }}
                    onClick={() => this.handleQuizComplete(85, true)}
                  />
                  <DefaultButton
                    text="Complete Quiz (Fail)"
                    iconProps={{ iconName: 'Cancel' }}
                    onClick={() => this.handleQuizComplete(50, false)}
                  />
                </Stack>
              </div>
            </Stack>
          ) : (
            <div className={`${styles.quizResult} ${quizPassed ? styles.passed : styles.failed}`}>
              <Icon iconName={quizPassed ? 'CheckMark' : 'Cancel'} style={{ fontSize: 32 }} />
              <Text variant="xLarge" style={{ marginTop: 8 }}>
                {quizPassed ? 'Quiz Passed!' : 'Quiz Not Passed'}
              </Text>
              <Text variant="large">Your Score: {quizScore}%</Text>
              {!quizPassed && (
                <DefaultButton
                  text="Retake Quiz"
                  iconProps={{ iconName: 'Refresh' }}
                  onClick={() => this.setState({ quizCompleted: false, quizScore: 0 })}
                  styles={{ root: { marginTop: 16 } }}
                />
              )}
            </div>
          )}
        </Stack>
      </div>
    );
  }

  private renderAcknowledgePanel(): JSX.Element {
    const {
      showAcknowledgePanel, policy, legalAgreement1, legalAgreement2, legalAgreement3,
      digitalSignature, acknowledgeNotes, submittingAcknowledgement, readDuration
    } = this.state;

    const formatDuration = (seconds: number): string => {
      const mins = Math.floor(seconds / 60);
      const secs = seconds % 60;
      return `${mins} minute${mins !== 1 ? 's' : ''} ${secs} second${secs !== 1 ? 's' : ''}`;
    };

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
              text={submittingAcknowledgement ? 'Submitting...' : 'Record Acknowledgement'}
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
          {/* Legal Header */}
          <div className={styles.legalHeader}>
            <Icon iconName="Shield" style={{ fontSize: 32, color: '#0078d4' }} />
            <Text variant="xLarge" style={{ fontWeight: 600 }}>Acknowledgement Declaration</Text>
          </div>

          {/* Policy Info */}
          <div className={styles.acknowledgePolicyInfo}>
            <Text variant="large" style={{ fontWeight: 600 }}>{policy?.PolicyNumber}</Text>
            <Text variant="large">{policy?.PolicyName}</Text>
            <Text variant="small" style={{ color: '#605e5c' }}>
              Version {policy?.VersionNumber} | Read time: {formatDuration(readDuration)}
            </Text>
          </div>

          <Separator />

          {/* Legal Agreements */}
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

          {/* Digital Signature */}
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

          {/* Optional Notes */}
          <TextField
            label="Additional Notes (Optional)"
            multiline
            rows={3}
            value={acknowledgeNotes}
            onChange={(e, value) => this.setState({ acknowledgeNotes: value || '' })}
            placeholder="Add any comments or clarifications..."
          />

          {/* Warning Notice */}
          <MessageBar messageBarType={MessageBarType.warning}>
            <strong>Important:</strong> Your acknowledgement will be recorded with a timestamp and stored for audit purposes.
            Ensure you have thoroughly read and understood this policy before proceeding.
          </MessageBar>
        </Stack>
      </Panel>
    );
  }

  private renderCongratulationsPanel(): JSX.Element {
    const { showCongratulationsPanel, policy, readReceipt, emailingReceipt, generatingPdf } = this.state;

    return (
      <Panel
        isOpen={showCongratulationsPanel}
        onDismiss={this.handleCloseCongratulations}
        type={PanelType.medium}
        headerText=""
        closeButtonAriaLabel="Close"
        styles={{ main: { background: 'linear-gradient(180deg, #f0f9ff 0%, #ffffff 100%)' } }}
      >
        <Stack tokens={{ childrenGap: 24 }} horizontalAlign="center" styles={{ root: { textAlign: 'center', padding: 20 } }}>
          {/* Certificate Image */}
          <div className={styles.certificateContainer}>
            <div className={styles.certificateHeader}>
              <Icon iconName="Certificate" style={{ fontSize: 48, color: '#0078d4' }} />
            </div>
            <div className={styles.certificateBody}>
              <Text variant="xxLarge" style={{ fontWeight: 300, color: '#0078d4' }}>Congratulations!</Text>
              <div className={styles.certificateSeal}>
                <Icon iconName="CheckMark" style={{ fontSize: 32, color: '#107c10' }} />
              </div>
              <Text variant="xLarge" style={{ fontWeight: 600 }}>Policy Acknowledged</Text>
              <Separator styles={{ root: { margin: '16px 0' } }} />
              <Text variant="large">{policy?.PolicyName}</Text>
              <Text variant="medium" style={{ color: '#605e5c' }}>{policy?.PolicyNumber}</Text>
              <div className={styles.certificateDetails}>
                <Text variant="small">Acknowledged by: {readReceipt?.UserDisplayName}</Text>
                <Text variant="small">Date: {readReceipt?.AcknowledgedDate?.toLocaleDateString()}</Text>
                <Text variant="small">Receipt: {readReceipt?.ReceiptNumber}</Text>
              </div>
            </div>
          </div>

          {/* Success Message */}
          <MessageBar messageBarType={MessageBarType.success} styles={{ root: { maxWidth: 400 } }}>
            Your policy acknowledgement has been recorded successfully. A read receipt has been generated for your records.
          </MessageBar>

          {/* Action Buttons */}
          <Stack horizontal tokens={{ childrenGap: 12 }} wrap horizontalAlign="center">
            <DefaultButton
              text={emailingReceipt ? 'Sending...' : 'Email Receipt'}
              iconProps={{ iconName: 'Mail' }}
              onClick={this.handleEmailReceipt}
              disabled={emailingReceipt}
            />
            <DefaultButton
              text={generatingPdf ? 'Generating...' : 'Print/Save PDF'}
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

          {/* Return to Policies */}
          <PrimaryButton
            text="Return to Policy Hub"
            iconProps={{ iconName: 'Back' }}
            onClick={() => window.location.href = '/sites/JML/SitePages/PolicyHub.aspx'}
            styles={{ root: { marginTop: 16 } }}
          />
        </Stack>
      </Panel>
    );
  }

  private renderReadReceiptPanel(): JSX.Element {
    const { showReadReceiptPanel, readReceipt, policy } = this.state;

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
              <Icon iconName="DocumentApproval" style={{ fontSize: 32, color: '#107c10' }} />
              <Text variant="xLarge" style={{ fontWeight: 600 }}>Policy Acknowledgement Receipt</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>{readReceipt.ReceiptNumber}</Text>
            </div>

            <Separator />

            <Stack tokens={{ childrenGap: 12 }}>
              <div className={styles.receiptRow}>
                <Text style={{ fontWeight: 600, minWidth: 140 }}>Employee:</Text>
                <Text>{readReceipt.UserDisplayName}</Text>
              </div>
              <div className={styles.receiptRow}>
                <Text style={{ fontWeight: 600, minWidth: 140 }}>Email:</Text>
                <Text>{readReceipt.UserEmail}</Text>
              </div>
              <div className={styles.receiptRow}>
                <Text style={{ fontWeight: 600, minWidth: 140 }}>Policy:</Text>
                <Text>{readReceipt.PolicyNumber} - {readReceipt.PolicyName}</Text>
              </div>
              <div className={styles.receiptRow}>
                <Text style={{ fontWeight: 600, minWidth: 140 }}>Version:</Text>
                <Text>{readReceipt.PolicyVersion}</Text>
              </div>
              <div className={styles.receiptRow}>
                <Text style={{ fontWeight: 600, minWidth: 140 }}>Read Start:</Text>
                <Text>{readReceipt.ReadStartTime?.toLocaleString()}</Text>
              </div>
              <div className={styles.receiptRow}>
                <Text style={{ fontWeight: 600, minWidth: 140 }}>Read End:</Text>
                <Text>{readReceipt.ReadEndTime?.toLocaleString()}</Text>
              </div>
              <div className={styles.receiptRow}>
                <Text style={{ fontWeight: 600, minWidth: 140 }}>Duration:</Text>
                <Text>{Math.floor(readReceipt.ReadDurationSeconds / 60)}m {readReceipt.ReadDurationSeconds % 60}s</Text>
              </div>
              {readReceipt.QuizRequired && (
                <>
                  <div className={styles.receiptRow}>
                    <Text style={{ fontWeight: 600, minWidth: 140 }}>Quiz Score:</Text>
                    <Text>{readReceipt.QuizScore}%</Text>
                  </div>
                </>
              )}
              <div className={styles.receiptRow}>
                <Text style={{ fontWeight: 600, minWidth: 140 }}>Acknowledged:</Text>
                <Text>{readReceipt.AcknowledgedDate?.toLocaleString()}</Text>
              </div>
              <div className={styles.receiptRow}>
                <Text style={{ fontWeight: 600, minWidth: 140 }}>Digital Signature:</Text>
                <Text style={{ fontStyle: 'italic' }}>{readReceipt.DigitalSignature}</Text>
              </div>
              <div className={styles.receiptRow}>
                <Text style={{ fontWeight: 600, minWidth: 140 }}>Device:</Text>
                <Text>{readReceipt.DeviceType} - {readReceipt.BrowserName}</Text>
              </div>
            </Stack>

            <Separator />

            <div className={styles.legalConfirmation}>
              <Text variant="small" style={{ fontWeight: 600 }}>Legal Confirmation:</Text>
              <Text variant="small" style={{ whiteSpace: 'pre-line', marginTop: 8 }}>
                {readReceipt.LegalConfirmationText}
              </Text>
            </div>

            <Stack horizontal tokens={{ childrenGap: 12 }}>
              <DefaultButton
                text="Email Copy"
                iconProps={{ iconName: 'Mail' }}
                onClick={this.handleEmailReceipt}
              />
              <DefaultButton
                text="Print/PDF"
                iconProps={{ iconName: 'PDF' }}
                onClick={this.handleGeneratePdf}
              />
            </Stack>
          </Stack>
        )}
      </Panel>
    );
  }

  private renderReadFlowActions(): JSX.Element | null {
    const { currentFlowStep, hasReadPolicy, acknowledgement } = this.state;

    // Don't show if already acknowledged
    if (acknowledgement?.Status === 'Acknowledged') return null;
    if (currentFlowStep === 'complete') return null;

    return (
      <div className={styles.readFlowActions}>
        {currentFlowStep === 'reading' && !hasReadPolicy && (
          <Stack tokens={{ childrenGap: 12 }}>
            <MessageBar messageBarType={MessageBarType.info}>
              Please read the entire policy before acknowledging. When you have finished reading, click the button below.
            </MessageBar>
            <PrimaryButton
              text="I Have Read This Policy"
              iconProps={{ iconName: 'CheckMark' }}
              onClick={this.handleMarkAsRead}
            />
          </Stack>
        )}

        {currentFlowStep === 'acknowledge' && (
          <Stack tokens={{ childrenGap: 12 }}>
            <MessageBar messageBarType={MessageBarType.success}>
              Great! You have completed reading the policy. Please proceed to acknowledge.
            </MessageBar>
            <PrimaryButton
              text="Acknowledge Policy"
              iconProps={{ iconName: 'Handwriting' }}
              onClick={this.handleOpenAcknowledgePanel}
            />
          </Stack>
        )}
      </div>
    );
  }

  public render(): React.ReactElement<IPolicyDetailsProps> {
    const {
      loading,
      error,
      policy,
      showAcknowledgeDialog,
      acknowledgeConfirmation,
      acknowledgeNotes,
      submittingAcknowledgement,
      showCommentDialog,
      newComment,
      submittingComment,
      acknowledgement,
      currentFlowStep
    } = this.state;

    return (
      <JmlAppLayout
        context={this.props.context}
        pageTitle="Policy Details"
        pageDescription="View policy content, version history and acknowledgements"
        pageIcon="Document"
        breadcrumbs={[{ text: 'JML Portal', url: '/sites/JML' }, { text: 'Policies', url: '/sites/JML/SitePages/PolicyHub.aspx' }, { text: 'Policy Details' }]}
        activeNavKey="policies"
        showQuickLinks={true}
        showSearch={true}
        showNotifications={true}
        compactFooter={true}
      >
        <section className={styles.policyDetails}>
          <Stack tokens={{ childrenGap: 24 }}>
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

            {!loading && !error && policy && (
              <>
                {/* Read Flow Progress - shows current step in the flow */}
                {acknowledgement && acknowledgement.Status !== 'Acknowledged' && this.renderReadFlowProgress()}

                {this.renderPolicyHeader()}
                {this.renderPolicyContent()}

                {/* Quiz Section - only shown if quiz is required and user has read the policy */}
                {this.renderQuizSection()}

                {/* Read Flow Actions - prompts user to proceed through the flow */}
                {this.renderReadFlowActions()}

                {/* Legacy acknowledgement section for backwards compatibility */}
                {currentFlowStep === 'complete' && this.renderAcknowledgement()}

                {this.renderRatings()}
                {this.renderComments()}

                {/* Panels */}
                {this.renderAcknowledgePanel()}
                {this.renderCongratulationsPanel()}
                {this.renderReadReceiptPanel()}
              </>
            )}
          </Stack>

          <Dialog
            hidden={!showAcknowledgeDialog}
            onDismiss={() => this.setState({ showAcknowledgeDialog: false })}
            dialogContentProps={{
              type: DialogType.normal,
              title: 'Acknowledge Policy',
              subText: 'Please confirm that you have read and understood this policy.'
            }}
          >
            <Stack tokens={{ childrenGap: 16 }}>
              <Checkbox
                label="I confirm that I have read and understood this policy"
                checked={acknowledgeConfirmation}
                onChange={(e, checked) => this.setState({ acknowledgeConfirmation: checked || false })}
              />
              <TextField
                label="Additional Notes (Optional)"
                multiline
                rows={3}
                value={acknowledgeNotes}
                onChange={(e, value) => this.setState({ acknowledgeNotes: value || '' })}
              />
            </Stack>
            <DialogFooter>
              <PrimaryButton
                text="Submit"
                onClick={this.handleAcknowledgeSubmit}
                disabled={!acknowledgeConfirmation || submittingAcknowledgement}
              />
              <DefaultButton
                text="Cancel"
                onClick={() => this.setState({ showAcknowledgeDialog: false })}
                disabled={submittingAcknowledgement}
              />
            </DialogFooter>
          </Dialog>

          <Dialog
            hidden={!showCommentDialog}
            onDismiss={() => this.setState({ showCommentDialog: false })}
            dialogContentProps={{
              type: DialogType.normal,
              title: 'Add Comment'
            }}
          >
            <TextField
              multiline
              rows={4}
              value={newComment}
              onChange={(e, value) => this.setState({ newComment: value || '' })}
              placeholder="Share your thoughts about this policy..."
            />
            <DialogFooter>
              <PrimaryButton
                text="Submit"
                onClick={this.handleSubmitComment}
                disabled={!newComment.trim() || submittingComment}
              />
              <DefaultButton
                text="Cancel"
                onClick={() => this.setState({ showCommentDialog: false })}
                disabled={submittingComment}
              />
            </DialogFooter>
          </Dialog>
          <this.dialogManager.DialogComponent />
        </section>
      </JmlAppLayout>
    );
  }
}
