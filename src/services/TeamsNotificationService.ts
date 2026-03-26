// @ts-nocheck
/**
 * TeamsNotificationService
 *
 * Sends Adaptive Card notifications to users via Microsoft Teams.
 * Uses Microsoft Graph API for proactive messaging.
 *
 * Supports:
 * - Activity Feed notifications (bell icon in Teams)
 * - 1:1 chat Adaptive Cards with action buttons
 * - Channel announcements via Incoming Webhook
 *
 * Phase A: URL-based actions (opens Policy Manager in browser)
 * Phase B: Server-side Bot callbacks (future — acknowledge/approve from Teams)
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';

// ═══════════════════════════════════════════════════════════════
// TYPES
// ═══════════════════════════════════════════════════════════════

export interface ITeamsConfig {
  enabled: boolean;
  channelWebhookUrl: string;
  enableActivityFeed: boolean;
  enableAdaptiveCards: boolean;
  enableChannelPosts: boolean;
  quietHoursStart: number; // 0-23 (e.g., 20 = 8pm)
  quietHoursEnd: number;   // 0-23 (e.g., 7 = 7am)
  respectQuietHours: boolean;
}

export const DEFAULT_TEAMS_CONFIG: ITeamsConfig = {
  enabled: false,
  channelWebhookUrl: '',
  enableActivityFeed: true,
  enableAdaptiveCards: true,
  enableChannelPosts: false,
  quietHoursStart: 20,
  quietHoursEnd: 7,
  respectQuietHours: true
};

export type TeamsCardType =
  | 'policy-published'
  | 'ack-required'
  | 'ack-reminder'
  | 'approval-request'
  | 'approval-result'
  | 'quiz-assigned'
  | 'sla-breach'
  | 'weekly-digest';

export interface ITeamsNotification {
  recipientEmail: string;
  cardType: TeamsCardType;
  data: Record<string, any>;
}

// ═══════════════════════════════════════════════════════════════
// SERVICE
// ═══════════════════════════════════════════════════════════════

export class TeamsNotificationService {
  private context: WebPartContext;
  private config: ITeamsConfig;
  private siteUrl: string;

  constructor(context: WebPartContext, config?: Partial<ITeamsConfig>) {
    this.context = context;
    this.config = { ...DEFAULT_TEAMS_CONFIG, ...config };
    this.siteUrl = context.pageContext?.web?.absoluteUrl || '';
  }

  /**
   * Update service configuration
   */
  public updateConfig(config: Partial<ITeamsConfig>): void {
    this.config = { ...this.config, ...config };
  }

  /**
   * Check if Teams notifications are enabled and within quiet hours
   */
  public isAvailable(): boolean {
    if (!this.config.enabled) return false;
    if (this.config.respectQuietHours) {
      const hour = new Date().getHours();
      const { quietHoursStart, quietHoursEnd } = this.config;
      if (quietHoursStart > quietHoursEnd) {
        // Overnight quiet hours (e.g., 20:00 - 07:00)
        if (hour >= quietHoursStart || hour < quietHoursEnd) return false;
      } else {
        if (hour >= quietHoursStart && hour < quietHoursEnd) return false;
      }
    }
    return true;
  }

  // ═══════════════════════════════════════════════════════════════
  // SEND METHODS
  // ═══════════════════════════════════════════════════════════════

  /**
   * Send an Adaptive Card to a user via Teams 1:1 chat
   */
  public async sendAdaptiveCard(notification: ITeamsNotification): Promise<boolean> {
    if (!this.isAvailable() || !this.config.enableAdaptiveCards) return false;

    try {
      const card = this.buildAdaptiveCard(notification.cardType, notification.data);
      const token = await this.getGraphToken();

      // Get or create 1:1 chat with the user
      const chatId = await this.getOrCreateChat(token, notification.recipientEmail);
      if (!chatId) return false;

      // Send the card as a chat message
      const response = await fetch(`https://graph.microsoft.com/v1.0/chats/${chatId}/messages`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          body: { contentType: 'html', content: '' },
          attachments: [{
            id: this.generateId(),
            contentType: 'application/vnd.microsoft.card.adaptive',
            contentUrl: null,
            content: JSON.stringify(card)
          }]
        })
      });

      return response.ok;
    } catch (err) {
      console.error('[TeamsNotificationService] sendAdaptiveCard failed:', err);
      return false;
    }
  }

  /**
   * Send an Activity Feed notification (Teams bell icon)
   */
  public async sendActivityFeed(
    recipientEmail: string,
    title: string,
    description: string,
    linkUrl?: string
  ): Promise<boolean> {
    if (!this.isAvailable() || !this.config.enableActivityFeed) return false;

    try {
      const token = await this.getGraphToken();

      // Resolve user ID from email
      const userResponse = await fetch(
        `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(recipientEmail)}?$select=id`,
        { headers: { 'Authorization': `Bearer ${token}` } }
      );
      if (!userResponse.ok) return false;
      const userData = await userResponse.json();

      // Send activity notification
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/users/${userData.id}/teamwork/sendActivityNotification`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            topic: {
              source: 'text',
              value: 'Policy Manager',
              webUrl: linkUrl || this.siteUrl
            },
            activityType: 'taskCreated',
            previewText: { content: description },
            templateParameters: [
              { name: 'taskName', value: title }
            ]
          })
        }
      );

      return response.ok;
    } catch (err) {
      console.error('[TeamsNotificationService] sendActivityFeed failed:', err);
      return false;
    }
  }

  /**
   * Send an Adaptive Card to a Teams channel via Incoming Webhook
   */
  public async sendChannelCard(cardType: TeamsCardType, data: Record<string, any>): Promise<boolean> {
    if (!this.isAvailable() || !this.config.enableChannelPosts || !this.config.channelWebhookUrl) return false;

    try {
      const card = this.buildAdaptiveCard(cardType, data);

      const response = await fetch(this.config.channelWebhookUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          type: 'message',
          attachments: [{
            contentType: 'application/vnd.microsoft.card.adaptive',
            contentUrl: null,
            content: card
          }]
        })
      });

      return response.ok;
    } catch (err) {
      console.error('[TeamsNotificationService] sendChannelCard failed:', err);
      return false;
    }
  }

  // ═══════════════════════════════════════════════════════════════
  // ADAPTIVE CARD TEMPLATES
  // ═══════════════════════════════════════════════════════════════

  /**
   * Build an Adaptive Card JSON for the given notification type
   */
  public buildAdaptiveCard(cardType: TeamsCardType, data: Record<string, any>): any {
    switch (cardType) {
      case 'policy-published': return this.cardPolicyPublished(data);
      case 'ack-required': return this.cardAckRequired(data);
      case 'ack-reminder': return this.cardAckReminder(data);
      case 'approval-request': return this.cardApprovalRequest(data);
      case 'approval-result': return this.cardApprovalResult(data);
      case 'quiz-assigned': return this.cardQuizAssigned(data);
      case 'sla-breach': return this.cardSlaBreach(data);
      case 'weekly-digest': return this.cardWeeklyDigest(data);
      default: return this.cardGeneric(data);
    }
  }

  // ── Policy Published ────────────────────────────────────────────
  private cardPolicyPublished(data: Record<string, any>): any {
    return {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.4',
      body: [
        {
          type: 'ColumnSet',
          columns: [
            {
              type: 'Column', width: 'auto',
              items: [{
                type: 'Image',
                url: `${this.siteUrl}/_layouts/15/images/siteIcon.png`,
                size: 'Small', style: 'Person'
              }]
            },
            {
              type: 'Column', width: 'stretch',
              items: [
                { type: 'TextBlock', text: 'Policy Manager', weight: 'Bolder', size: 'Small', color: 'Accent' },
                { type: 'TextBlock', text: 'New Policy Published', spacing: 'None', isSubtle: true, size: 'Small' }
              ]
            }
          ]
        },
        { type: 'TextBlock', text: data.policyTitle || 'Untitled Policy', weight: 'Bolder', size: 'Medium', wrap: true },
        { type: 'TextBlock', text: data.summary || 'A new policy has been published and requires your attention.', wrap: true, size: 'Small' },
        {
          type: 'FactSet',
          facts: [
            { title: 'Category', value: data.category || 'General' },
            { title: 'Risk Level', value: data.riskLevel || 'Medium' },
            { title: 'Effective', value: data.effectiveDate || 'Immediately' },
            ...(data.requiresAck ? [{ title: 'Action Required', value: 'Acknowledgement needed' }] : [])
          ]
        }
      ],
      actions: [
        { type: 'Action.OpenUrl', title: 'Read Policy', url: `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${data.policyId || ''}`, style: 'positive' },
        ...(data.requiresAck ? [{ type: 'Action.OpenUrl', title: 'Acknowledge', url: `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${data.policyId || ''}&action=ack` }] : [])
      ]
    };
  }

  // ── Acknowledgement Required ────────────────────────────────────
  private cardAckRequired(data: Record<string, any>): any {
    return {
      type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4',
      body: [
        {
          type: 'Container',
          style: 'attention',
          items: [
            { type: 'TextBlock', text: 'Action Required', weight: 'Bolder', color: 'Attention', size: 'Small' },
            { type: 'TextBlock', text: 'Policy Acknowledgement Required', weight: 'Bolder', size: 'Medium', wrap: true }
          ]
        },
        { type: 'TextBlock', text: data.policyTitle || '', weight: 'Bolder', size: 'Large', wrap: true, spacing: 'Medium' },
        { type: 'TextBlock', text: data.message || 'Please read and acknowledge this policy.', wrap: true },
        {
          type: 'FactSet',
          facts: [
            { title: 'Deadline', value: data.deadline || 'Not specified' },
            { title: 'Category', value: data.category || 'General' },
            { title: 'Risk Level', value: data.riskLevel || 'Medium' }
          ]
        }
      ],
      actions: [
        { type: 'Action.OpenUrl', title: 'Read & Acknowledge', url: `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${data.policyId || ''}&action=ack`, style: 'positive' },
        { type: 'Action.OpenUrl', title: 'View Policy', url: `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${data.policyId || ''}` }
      ]
    };
  }

  // ── Acknowledgement Reminder ────────────────────────────────────
  private cardAckReminder(data: Record<string, any>): any {
    const isOverdue = data.daysRemaining !== undefined && data.daysRemaining < 0;
    const urgencyColor = isOverdue ? 'Attention' : (data.daysRemaining <= 1 ? 'Warning' : 'Default');
    return {
      type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4',
      body: [
        {
          type: 'Container', style: isOverdue ? 'attention' : 'warning',
          items: [
            { type: 'TextBlock', text: isOverdue ? 'OVERDUE' : 'Reminder', weight: 'Bolder', color: urgencyColor, size: 'Small' },
            { type: 'TextBlock', text: data.policyTitle || '', weight: 'Bolder', size: 'Medium', wrap: true }
          ]
        },
        {
          type: 'TextBlock',
          text: isOverdue
            ? `Your acknowledgement is **${Math.abs(data.daysRemaining)} days overdue**. Please complete immediately.`
            : `You have **${data.daysRemaining} day${data.daysRemaining !== 1 ? 's' : ''}** remaining to acknowledge this policy.`,
          wrap: true
        },
        {
          type: 'FactSet',
          facts: [
            { title: 'Deadline', value: data.deadline || '' },
            { title: 'Status', value: isOverdue ? 'Overdue' : 'Pending' }
          ]
        }
      ],
      actions: [
        { type: 'Action.OpenUrl', title: 'Acknowledge Now', url: `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${data.policyId || ''}&action=ack`, style: 'positive' }
      ]
    };
  }

  // ── Approval Request ────────────────────────────────────────────
  private cardApprovalRequest(data: Record<string, any>): any {
    return {
      type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4',
      body: [
        {
          type: 'Container', style: 'emphasis',
          items: [
            { type: 'TextBlock', text: 'Approval Required', weight: 'Bolder', color: 'Accent', size: 'Small' },
            { type: 'TextBlock', text: data.policyTitle || '', weight: 'Bolder', size: 'Medium', wrap: true }
          ]
        },
        { type: 'TextBlock', text: `**${data.authorName || 'An author'}** has submitted this policy for your approval.`, wrap: true },
        {
          type: 'FactSet',
          facts: [
            { title: 'Submitted by', value: data.authorName || '' },
            { title: 'Level', value: `Level ${data.approvalLevel || 1}` },
            { title: 'Due', value: data.dueDate || 'Not specified' },
            { title: 'Category', value: data.category || 'General' },
            { title: 'Risk', value: data.riskLevel || 'Medium' }
          ]
        },
        ...(data.summary ? [{ type: 'TextBlock', text: data.summary, wrap: true, size: 'Small', isSubtle: true }] : [])
      ],
      actions: [
        { type: 'Action.OpenUrl', title: 'Review & Approve', url: `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${data.policyId || ''}&mode=review`, style: 'positive' },
        { type: 'Action.OpenUrl', title: 'Review & Reject', url: `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${data.policyId || ''}&mode=review`, style: 'destructive' },
        { type: 'Action.OpenUrl', title: 'Review Details', url: `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${data.policyId || ''}&mode=review` }
      ]
    };
  }

  // ── Approval Result ─────────────────────────────────────────────
  private cardApprovalResult(data: Record<string, any>): any {
    const isApproved = data.decision === 'Approved';
    return {
      type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4',
      body: [
        {
          type: 'Container', style: isApproved ? 'good' : 'attention',
          items: [
            { type: 'TextBlock', text: isApproved ? 'Approved' : 'Rejected', weight: 'Bolder', color: isApproved ? 'Good' : 'Attention', size: 'Medium' },
            { type: 'TextBlock', text: data.policyTitle || '', weight: 'Bolder', size: 'Small', wrap: true }
          ]
        },
        { type: 'TextBlock', text: `**${data.approverName || 'An approver'}** has ${isApproved ? 'approved' : 'rejected'} this policy.`, wrap: true },
        ...(data.comments ? [{
          type: 'Container', style: 'emphasis',
          items: [
            { type: 'TextBlock', text: 'Comments:', weight: 'Bolder', size: 'Small' },
            { type: 'TextBlock', text: data.comments, wrap: true, size: 'Small', isSubtle: true }
          ]
        }] : []),
        {
          type: 'FactSet',
          facts: [
            { title: 'Decision', value: data.decision || '' },
            { title: 'Level', value: `Level ${data.approvalLevel || 1}` },
            { title: 'Date', value: data.decisionDate || new Date().toLocaleDateString() }
          ]
        }
      ],
      actions: [
        { type: 'Action.OpenUrl', title: 'View Policy', url: `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${data.policyId || ''}` }
      ]
    };
  }

  // ── Quiz Assigned ───────────────────────────────────────────────
  private cardQuizAssigned(data: Record<string, any>): any {
    return {
      type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4',
      body: [
        {
          type: 'ColumnSet',
          columns: [
            { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: '📝', size: 'ExtraLarge' }] },
            {
              type: 'Column', width: 'stretch',
              items: [
                { type: 'TextBlock', text: 'Quiz Required', weight: 'Bolder', size: 'Small', color: 'Accent' },
                { type: 'TextBlock', text: data.quizTitle || data.policyTitle || '', weight: 'Bolder', size: 'Medium', wrap: true, spacing: 'None' }
              ]
            }
          ]
        },
        { type: 'TextBlock', text: `A comprehension quiz is required for **${data.policyTitle || 'this policy'}**.`, wrap: true },
        {
          type: 'FactSet',
          facts: [
            { title: 'Passing Score', value: `${data.passingScore || 75}%` },
            { title: 'Attempts', value: `${data.maxAttempts || 3} allowed` },
            { title: 'Time Limit', value: data.timeLimit ? `${data.timeLimit} minutes` : 'No limit' }
          ]
        }
      ],
      actions: [
        { type: 'Action.OpenUrl', title: 'Start Quiz', url: `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${data.policyId || ''}&action=quiz`, style: 'positive' },
        { type: 'Action.OpenUrl', title: 'Read Policy First', url: `${this.siteUrl}/SitePages/PolicyDetails.aspx?policyId=${data.policyId || ''}` }
      ]
    };
  }

  // ── SLA Breach ──────────────────────────────────────────────────
  private cardSlaBreach(data: Record<string, any>): any {
    return {
      type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4',
      body: [
        {
          type: 'Container', style: 'attention',
          items: [
            { type: 'TextBlock', text: 'SLA BREACH', weight: 'Bolder', color: 'Attention', size: 'Small' },
            { type: 'TextBlock', text: data.slaType || 'SLA Target Exceeded', weight: 'Bolder', size: 'Medium', wrap: true }
          ]
        },
        { type: 'TextBlock', text: `The **${data.slaType || 'process'}** SLA for **${data.policyTitle || 'a policy'}** has been breached.`, wrap: true },
        {
          type: 'FactSet',
          facts: [
            { title: 'Policy', value: data.policyTitle || '' },
            { title: 'Target', value: `${data.targetDays || '?'} days` },
            { title: 'Actual', value: `${data.actualDays || '?'} days` },
            { title: 'Exceeded by', value: `${(data.actualDays || 0) - (data.targetDays || 0)} days` }
          ]
        }
      ],
      actions: [
        { type: 'Action.OpenUrl', title: 'View Details', url: `${this.siteUrl}/SitePages/PolicyAnalytics.aspx`, style: 'positive' }
      ]
    };
  }

  // ── Weekly Digest ───────────────────────────────────────────────
  private cardWeeklyDigest(data: Record<string, any>): any {
    return {
      type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4',
      body: [
        {
          type: 'Container', style: 'emphasis',
          items: [
            { type: 'TextBlock', text: 'Your Weekly Policy Summary', weight: 'Bolder', size: 'Medium' },
            { type: 'TextBlock', text: `Week of ${new Date().toLocaleDateString()}`, isSubtle: true, size: 'Small', spacing: 'None' }
          ]
        },
        {
          type: 'ColumnSet',
          columns: [
            {
              type: 'Column', width: 'stretch',
              items: [
                { type: 'TextBlock', text: String(data.pendingAck || 0), weight: 'Bolder', size: 'ExtraLarge', horizontalAlignment: 'Center', color: data.pendingAck > 0 ? 'Attention' : 'Default' },
                { type: 'TextBlock', text: 'Pending Ack', size: 'Small', horizontalAlignment: 'Center', isSubtle: true, spacing: 'None' }
              ]
            },
            {
              type: 'Column', width: 'stretch',
              items: [
                { type: 'TextBlock', text: String(data.pendingApprovals || 0), weight: 'Bolder', size: 'ExtraLarge', horizontalAlignment: 'Center', color: data.pendingApprovals > 0 ? 'Warning' : 'Default' },
                { type: 'TextBlock', text: 'Approvals', size: 'Small', horizontalAlignment: 'Center', isSubtle: true, spacing: 'None' }
              ]
            },
            {
              type: 'Column', width: 'stretch',
              items: [
                { type: 'TextBlock', text: String(data.newPolicies || 0), weight: 'Bolder', size: 'ExtraLarge', horizontalAlignment: 'Center', color: 'Accent' },
                { type: 'TextBlock', text: 'New Policies', size: 'Small', horizontalAlignment: 'Center', isSubtle: true, spacing: 'None' }
              ]
            }
          ]
        },
        ...(data.overdueCount > 0 ? [{
          type: 'Container', style: 'attention',
          items: [{ type: 'TextBlock', text: `⚠ You have **${data.overdueCount}** overdue acknowledgement${data.overdueCount !== 1 ? 's' : ''}.`, wrap: true, color: 'Attention' }]
        }] : [])
      ],
      actions: [
        { type: 'Action.OpenUrl', title: 'Open Policy Manager', url: `${this.siteUrl}/SitePages/MyPolicies.aspx`, style: 'positive' },
        { type: 'Action.OpenUrl', title: 'View Policy Hub', url: `${this.siteUrl}/SitePages/PolicyHub.aspx` }
      ]
    };
  }

  // ── Generic fallback ────────────────────────────────────────────
  private cardGeneric(data: Record<string, any>): any {
    return {
      type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4',
      body: [
        { type: 'TextBlock', text: data.title || 'Policy Manager Notification', weight: 'Bolder', size: 'Medium' },
        { type: 'TextBlock', text: data.message || '', wrap: true }
      ],
      actions: [
        { type: 'Action.OpenUrl', title: 'Open Policy Manager', url: this.siteUrl }
      ]
    };
  }

  // ═══════════════════════════════════════════════════════════════
  // HELPERS
  // ═══════════════════════════════════════════════════════════════

  /**
   * Get a Graph API access token via SPFx AAD token provider
   */
  private async getGraphToken(): Promise<string> {
    const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
    return tokenProvider.getToken('https://graph.microsoft.com');
  }

  /**
   * Get or create a 1:1 chat between the app and a user
   */
  private async getOrCreateChat(token: string, recipientEmail: string): Promise<string | null> {
    try {
      // Get current user ID
      const meResponse = await fetch('https://graph.microsoft.com/v1.0/me?$select=id', {
        headers: { 'Authorization': `Bearer ${token}` }
      });
      if (!meResponse.ok) return null;
      const meData = await meResponse.json();

      // Get recipient user ID
      const recipientResponse = await fetch(
        `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(recipientEmail)}?$select=id`,
        { headers: { 'Authorization': `Bearer ${token}` } }
      );
      if (!recipientResponse.ok) return null;
      const recipientData = await recipientResponse.json();

      // Create or get existing 1:1 chat
      const chatResponse = await fetch('https://graph.microsoft.com/v1.0/chats', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          chatType: 'oneOnOne',
          members: [
            {
              '@odata.type': '#microsoft.graph.aadUserConversationMember',
              roles: ['owner'],
              'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${meData.id}')`
            },
            {
              '@odata.type': '#microsoft.graph.aadUserConversationMember',
              roles: ['owner'],
              'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${recipientData.id}')`
            }
          ]
        })
      });

      if (chatResponse.ok) {
        const chatData = await chatResponse.json();
        return chatData.id;
      }
      return null;
    } catch {
      return null;
    }
  }

  private generateId(): string {
    return `card-${Date.now()}-${Math.random().toString(36).substring(2, 8)}`;
  }
}
