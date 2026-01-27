// @ts-nocheck
// Policy Social Service
// Handles social features: rate, comment, share, follow

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import {
  IPolicyRating,
  IPolicyComment,
  IPolicyCommentLike,
  IPolicyShare,
  IPolicyFollower,
  IRatePolicyRequest,
  ICommentPolicyRequest,
  ISharePolicyRequest,
  IFollowPolicyRequest
} from '../models/IPolicy';
import { logger } from './LoggingService';
import { PolicyLists, SocialLists, SystemLists } from '../constants/SharePointListNames';

export class PolicySocialService {
  private sp: SPFI;
  private readonly POLICY_RATINGS_LIST = SocialLists.POLICY_RATINGS;
  private readonly POLICY_COMMENTS_LIST = SocialLists.POLICY_COMMENTS;
  private readonly POLICY_COMMENT_LIKES_LIST = SocialLists.POLICY_COMMENT_LIKES;
  private readonly POLICY_SHARES_LIST = SocialLists.POLICY_SHARES;
  private readonly POLICY_FOLLOWERS_LIST = SocialLists.POLICY_FOLLOWERS;
  private readonly POLICIES_LIST = PolicyLists.POLICIES;
  private readonly NOTIFICATION_QUEUE_LIST = SystemLists.NOTIFICATION_QUEUE;
  private currentUserId: number = 0;
  private currentUserEmail: string = '';
  private currentUserDisplayName: string = '';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Initialize service
   */
  public async initialize(): Promise<void> {
    try {
      const user = await this.sp.web.currentUser();
      this.currentUserId = user.Id;
      this.currentUserEmail = user.Email;
      this.currentUserDisplayName = user.Title || user.Email;
    } catch (error) {
      logger.error('PolicySocialService', 'Failed to initialize:', error);
      throw error;
    }
  }

  // ============================================================================
  // RATINGS
  // ============================================================================

  /**
   * Rate a policy
   */
  public async ratePolicy(request: IRatePolicyRequest): Promise<IPolicyRating> {
    try {
      // Check if user already rated
      const existingRatings = await this.sp.web.lists
        .getByTitle(this.POLICY_RATINGS_LIST)
        .items.filter(`PolicyId eq ${request.policyId} and UserId eq ${this.currentUserId}`)
        .top(1)();

      const ratingData = {
        Title: `Policy ${request.policyId} - Rating by User ${this.currentUserId}`,
        PolicyId: request.policyId,
        UserId: this.currentUserId,
        UserEmail: this.currentUserEmail,
        Rating: request.rating,
        RatingDate: new Date().toISOString(),
        ReviewTitle: request.reviewTitle,
        ReviewText: request.reviewText,
        ReviewHelpfulCount: 0,
        IsVerifiedReader: await this.isVerifiedReader(request.policyId, this.currentUserId)
      };

      let result;
      if (existingRatings.length > 0) {
        // Update existing rating
        await this.sp.web.lists
          .getByTitle(this.POLICY_RATINGS_LIST)
          .items.getById(existingRatings[0].Id)
          .update(ratingData);
        result = { data: { Id: existingRatings[0].Id } };
      } else {
        // Create new rating
        result = await this.sp.web.lists
          .getByTitle(this.POLICY_RATINGS_LIST)
          .items.add(ratingData);
      }

      // Update policy average rating
      await this.updatePolicyAverageRating(request.policyId);

      return await this.getRatingById(result.data.Id);
    } catch (error) {
      logger.error('PolicySocialService', 'Failed to rate policy:', error);
      throw error;
    }
  }

  /**
   * Get policy ratings
   */
  public async getPolicyRatings(policyId: number): Promise<IPolicyRating[]> {
    try {
      const ratings = await this.sp.web.lists
        .getByTitle(this.POLICY_RATINGS_LIST)
        .items.filter(`PolicyId eq ${policyId}`)
        .orderBy('RatingDate', false)
        .top(100)();

      return ratings as IPolicyRating[];
    } catch (error) {
      logger.error('PolicySocialService', 'Failed to get policy ratings:', error);
      throw error;
    }
  }

  /**
   * Get rating by ID
   */
  private async getRatingById(ratingId: number): Promise<IPolicyRating> {
    const rating = await this.sp.web.lists
      .getByTitle(this.POLICY_RATINGS_LIST)
      .items.getById(ratingId)();
    return rating as IPolicyRating;
  }

  /**
   * Update policy average rating
   */
  private async updatePolicyAverageRating(policyId: number): Promise<void> {
    try {
      const ratings = await this.getPolicyRatings(policyId);
      if (ratings.length === 0) return;

      const totalRating = ratings.reduce((sum, r) => sum + r.Rating, 0);
      const averageRating = totalRating / ratings.length;
      const ratingCount = ratings.length;

      await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(policyId)
        .update({
          AverageRating: averageRating,
          RatingCount: ratingCount
        });
    } catch (error) {
      logger.error('PolicySocialService', 'Failed to update average rating:', error);
    }
  }

  // ============================================================================
  // COMMENTS
  // ============================================================================

  /**
   * Add comment to policy
   */
  public async commentOnPolicy(request: ICommentPolicyRequest): Promise<IPolicyComment> {
    try {
      const commentData = {
        Title: `Comment on Policy ${request.policyId}`,
        PolicyId: request.policyId,
        UserId: this.currentUserId,
        UserEmail: this.currentUserEmail,
        CommentText: request.commentText,
        CommentDate: new Date().toISOString(),
        IsEdited: false,
        ParentCommentId: request.parentCommentId,
        ReplyCount: 0,
        LikeCount: 0,
        IsStaffResponse: false,
        IsApproved: true, // Auto-approve for now
        IsDeleted: false
      };

      const result = await this.sp.web.lists
        .getByTitle(this.POLICY_COMMENTS_LIST)
        .items.add(commentData);

      // Update reply count if this is a reply
      if (request.parentCommentId) {
        const parentComment = await this.sp.web.lists
          .getByTitle(this.POLICY_COMMENTS_LIST)
          .items.getById(request.parentCommentId)() as IPolicyComment;

        await this.sp.web.lists
          .getByTitle(this.POLICY_COMMENTS_LIST)
          .items.getById(request.parentCommentId)
          .update({
            ReplyCount: (parentComment.ReplyCount || 0) + 1
          });
      }

      // Notify followers
      await this.notifyFollowersOfComment(request.policyId, result.data.Id);

      return await this.getCommentById(result.data.Id);
    } catch (error) {
      logger.error('PolicySocialService', 'Failed to comment on policy:', error);
      throw error;
    }
  }

  /**
   * Get policy comments
   */
  public async getPolicyComments(policyId: number): Promise<IPolicyComment[]> {
    try {
      const comments = await this.sp.web.lists
        .getByTitle(this.POLICY_COMMENTS_LIST)
        .items.filter(`PolicyId eq ${policyId} and IsDeleted eq false and IsApproved eq true`)
        .orderBy('CommentDate', false)
        .top(500)();

      return comments as IPolicyComment[];
    } catch (error) {
      logger.error('PolicySocialService', 'Failed to get policy comments:', error);
      throw error;
    }
  }

  /**
   * Like a comment
   */
  public async likeComment(commentId: number): Promise<void> {
    try {
      // Check if already liked
      const existing = await this.sp.web.lists
        .getByTitle(this.POLICY_COMMENT_LIKES_LIST)
        .items.filter(`CommentId eq ${commentId} and UserId eq ${this.currentUserId}`)
        .top(1)();

      if (existing.length > 0) {
        // Unlike
        await this.sp.web.lists
          .getByTitle(this.POLICY_COMMENT_LIKES_LIST)
          .items.getById(existing[0].Id)
          .delete();
      } else {
        // Like
        await this.sp.web.lists
          .getByTitle(this.POLICY_COMMENT_LIKES_LIST)
          .items.add({
            Title: `Like on Comment ${commentId}`,
            CommentId: commentId,
            UserId: this.currentUserId,
            LikedDate: new Date().toISOString()
          });
      }

      // Update like count
      const likes = await this.sp.web.lists
        .getByTitle(this.POLICY_COMMENT_LIKES_LIST)
        .items.filter(`CommentId eq ${commentId}`)
        .top(1000)();

      await this.sp.web.lists
        .getByTitle(this.POLICY_COMMENTS_LIST)
        .items.getById(commentId)
        .update({
          LikeCount: likes.length
        });
    } catch (error) {
      logger.error('PolicySocialService', 'Failed to like comment:', error);
      throw error;
    }
  }

  /**
   * Get comment by ID
   */
  private async getCommentById(commentId: number): Promise<IPolicyComment> {
    const comment = await this.sp.web.lists
      .getByTitle(this.POLICY_COMMENTS_LIST)
      .items.getById(commentId)();
    return comment as IPolicyComment;
  }

  // ============================================================================
  // SHARING
  // ============================================================================

  /**
   * Share a policy
   */
  public async sharePolicy(request: ISharePolicyRequest): Promise<IPolicyShare> {
    try {
      const shareData = {
        Title: `Share Policy ${request.policyId}`,
        PolicyId: request.policyId,
        SharedById: this.currentUserId,
        SharedByEmail: this.currentUserEmail,
        ShareMethod: request.shareMethod,
        ShareDate: new Date().toISOString(),
        ShareMessage: request.message,
        SharedWithUserIds: request.recipientUserIds ? JSON.stringify(request.recipientUserIds) : undefined,
        SharedWithEmails: request.recipientEmails ? JSON.stringify(request.recipientEmails) : undefined,
        SharedWithTeamsChannelId: request.teamsChannelId,
        ViewCount: 0
      };

      const result = await this.sp.web.lists
        .getByTitle(this.POLICY_SHARES_LIST)
        .items.add(shareData);

      // Send notifications based on share method
      await this.sendShareNotifications(request, result.data.Id);

      return result.data as IPolicyShare;
    } catch (error) {
      logger.error('PolicySocialService', 'Failed to share policy:', error);
      throw error;
    }
  }

  /**
   * Send share notifications
   */
  private async sendShareNotifications(request: ISharePolicyRequest, shareId: number): Promise<void> {
    try {
      // Get policy details for the notification
      const policy = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(request.policyId)
        .select('Id', 'Title', 'PolicyNumber')();

      const policyTitle = policy.Title || `Policy ${request.policyId}`;
      const policyNumber = policy.PolicyNumber || '';

      // Build recipient list
      const recipients: { email: string; userId?: number }[] = [];

      if (request.recipientEmails && request.recipientEmails.length > 0) {
        request.recipientEmails.forEach(email => {
          recipients.push({ email });
        });
      }

      if (request.recipientUserIds && request.recipientUserIds.length > 0) {
        // Get user emails for the user IDs
        for (const userId of request.recipientUserIds) {
          try {
            const user = await this.sp.web.siteUsers.getById(userId)();
            if (user.Email) {
              recipients.push({ email: user.Email, userId });
            }
          } catch (e) {
            logger.warn('PolicySocialService', `Could not get email for user ${userId}`);
          }
        }
      }

      if (recipients.length === 0) {
        logger.info('PolicySocialService', 'No recipients for share notification');
        return;
      }

      // Queue notifications based on share method
      for (const recipient of recipients) {
        if (request.shareMethod === 'Email' || request.shareMethod === 'Link') {
          // Queue email notification
          await this.sp.web.lists
            .getByTitle(this.NOTIFICATION_QUEUE_LIST)
            .items.add({
              Title: `Policy Shared: ${policyTitle}`,
              NotificationType: 'PolicyShared',
              RecipientEmail: recipient.email,
              RecipientUserId: recipient.userId,
              SenderEmail: this.currentUserEmail,
              SenderUserId: this.currentUserId,
              SenderName: this.currentUserDisplayName,
              PolicyId: request.policyId,
              PolicyTitle: policyTitle,
              PolicyVersion: policyNumber,
              Message: request.message || '',
              Channel: 'Email',
              Priority: 'Normal',
              Status: 'Pending',
              RetryCount: 0,
              MaxRetries: 3,
              RelatedShareId: shareId
            });
        }

        if (request.shareMethod === 'Teams' && request.teamsChannelId) {
          // Queue Teams notification
          await this.sp.web.lists
            .getByTitle(this.NOTIFICATION_QUEUE_LIST)
            .items.add({
              Title: `Policy Shared: ${policyTitle}`,
              NotificationType: 'PolicyShared',
              RecipientEmail: recipient.email,
              RecipientUserId: recipient.userId,
              SenderEmail: this.currentUserEmail,
              SenderUserId: this.currentUserId,
              SenderName: this.currentUserDisplayName,
              PolicyId: request.policyId,
              PolicyTitle: policyTitle,
              PolicyVersion: policyNumber,
              Message: request.message || '',
              Channel: 'Teams',
              Priority: 'Normal',
              Status: 'Pending',
              RetryCount: 0,
              MaxRetries: 3,
              RelatedShareId: shareId,
              TeamsChannelId: request.teamsChannelId
            });
        }
      }

      logger.info('PolicySocialService', `Queued share notifications for ${recipients.length} recipients`);
    } catch (error) {
      logger.error('PolicySocialService', 'Failed to send share notifications:', error);
      // Don't throw - notification failure shouldn't fail the share
    }
  }

  /**
   * Build email body for policy share
   */
  private buildShareEmailBody(policyTitle: string, policyNumber: string, message?: string): string {
    const policyRef = policyNumber ? `(${policyNumber})` : '';
    let body = `<p>${this.currentUserDisplayName} has shared a policy with you:</p>`;
    body += `<h3>${policyTitle} ${policyRef}</h3>`;

    if (message) {
      body += `<p><strong>Message:</strong></p>`;
      body += `<blockquote>${message}</blockquote>`;
    }

    body += `<p>Click below to view the policy:</p>`;
    body += `<p><a href="{{PolicyUrl}}" style="background-color:#0078d4;color:white;padding:10px 20px;text-decoration:none;border-radius:4px;">View Policy</a></p>`;
    body += `<hr>`;
    body += `<p style="font-size:12px;color:#666;">This policy was shared via the JML Policy Management System.</p>`;

    return body;
  }

  /**
   * Build Teams Adaptive Card for policy share
   */
  private buildShareTeamsCard(policyTitle: string, policyNumber: string, message?: string): object {
    const policyRef = policyNumber ? ` (${policyNumber})` : '';
    const card: {
      type: string;
      $schema: string;
      version: string;
      body: Array<{
        type: string;
        text: string;
        size?: string;
        weight?: string;
        wrap?: boolean;
        style?: string;
        color?: string;
      }>;
      actions: Array<{
        type: string;
        title: string;
        url: string;
        style?: string;
      }>;
    } = {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: '📋 Policy Shared',
          size: 'Large',
          weight: 'Bolder'
        },
        {
          type: 'TextBlock',
          text: `${this.currentUserDisplayName} shared a policy with you`,
          wrap: true
        },
        {
          type: 'TextBlock',
          text: `**${policyTitle}${policyRef}**`,
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.OpenUrl',
          title: 'View Policy',
          url: '{{PolicyUrl}}',
          style: 'positive'
        }
      ]
    };

    if (message) {
      card.body.push({
        type: 'TextBlock',
        text: `_"${message}"_`,
        wrap: true,
        style: 'default',
        color: 'accent'
      });
    }

    return card;
  }

  // ============================================================================
  // FOLLOWING
  // ============================================================================

  /**
   * Follow a policy
   */
  public async followPolicy(request: IFollowPolicyRequest): Promise<IPolicyFollower> {
    try {
      // Check if already following
      const existing = await this.sp.web.lists
        .getByTitle(this.POLICY_FOLLOWERS_LIST)
        .items.filter(`PolicyId eq ${request.policyId} and UserId eq ${this.currentUserId}`)
        .top(1)();

      const followerData = {
        Title: `Following Policy ${request.policyId}`,
        PolicyId: request.policyId,
        UserId: this.currentUserId,
        UserEmail: this.currentUserEmail,
        FollowedDate: new Date().toISOString(),
        NotifyOnUpdate: request.notifyOnUpdate,
        NotifyOnComment: request.notifyOnComment,
        NotifyOnNewVersion: request.notifyOnNewVersion,
        EmailNotifications: true,
        TeamsNotifications: true,
        InAppNotifications: true
      };

      let result;
      if (existing.length > 0) {
        // Update existing
        await this.sp.web.lists
          .getByTitle(this.POLICY_FOLLOWERS_LIST)
          .items.getById(existing[0].Id)
          .update(followerData);
        result = { data: { Id: existing[0].Id } };
      } else {
        // Create new
        result = await this.sp.web.lists
          .getByTitle(this.POLICY_FOLLOWERS_LIST)
          .items.add(followerData);
      }

      return result.data as IPolicyFollower;
    } catch (error) {
      logger.error('PolicySocialService', 'Failed to follow policy:', error);
      throw error;
    }
  }

  /**
   * Unfollow a policy
   */
  public async unfollowPolicy(policyId: number): Promise<void> {
    try {
      const followers = await this.sp.web.lists
        .getByTitle(this.POLICY_FOLLOWERS_LIST)
        .items.filter(`PolicyId eq ${policyId} and UserId eq ${this.currentUserId}`)
        .top(1)();

      if (followers.length > 0) {
        await this.sp.web.lists
          .getByTitle(this.POLICY_FOLLOWERS_LIST)
          .items.getById(followers[0].Id)
          .delete();
      }
    } catch (error) {
      logger.error('PolicySocialService', 'Failed to unfollow policy:', error);
      throw error;
    }
  }

  /**
   * Get policy followers
   */
  public async getPolicyFollowers(policyId: number): Promise<IPolicyFollower[]> {
    try {
      const followers = await this.sp.web.lists
        .getByTitle(this.POLICY_FOLLOWERS_LIST)
        .items.filter(`PolicyId eq ${policyId}`)
        .top(1000)();

      return followers as IPolicyFollower[];
    } catch (error) {
      logger.error('PolicySocialService', 'Failed to get policy followers:', error);
      throw error;
    }
  }

  /**
   * Check if user is following a policy
   */
  public async isFollowingPolicy(policyId: number, userId?: number): Promise<boolean> {
    try {
      const targetUserId = userId || this.currentUserId;
      const followers = await this.sp.web.lists
        .getByTitle(this.POLICY_FOLLOWERS_LIST)
        .items.filter(`PolicyId eq ${policyId} and UserId eq ${targetUserId}`)
        .top(1)();

      return followers.length > 0;
    } catch (error) {
      return false;
    }
  }

  /**
   * Notify followers of policy update
   */
  public async notifyFollowersOfUpdate(policyId: number, updateType: 'Update' | 'NewVersion' | 'Comment'): Promise<void> {
    try {
      const followers = await this.getPolicyFollowers(policyId);
      const notifyList = followers.filter(f => {
        if (updateType === 'Update') return f.NotifyOnUpdate;
        if (updateType === 'NewVersion') return f.NotifyOnNewVersion;
        if (updateType === 'Comment') return f.NotifyOnComment;
        return false;
      });

      if (notifyList.length === 0) {
        logger.info('PolicySocialService', `No followers to notify for policy ${policyId} ${updateType}`);
        return;
      }

      // Get policy details
      const policy = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(policyId)
        .select('Id', 'Title', 'PolicyNumber', 'Version')();

      const policyTitle = policy.Title || `Policy ${policyId}`;
      const policyNumber = policy.PolicyNumber || '';
      const version = policy.Version || '1.0';

      // Build notification content based on update type
      const notificationContent = this.buildFollowerNotificationContent(
        updateType,
        policyTitle,
        policyNumber,
        version
      );

      // Map update type to notification type enum value
      const notificationTypeMap: Record<string, string> = {
        'Update': 'PolicyUpdated',
        'NewVersion': 'PolicyUpdated',
        'Comment': 'PolicyComment'
      };

      // Queue notifications for each follower based on their preferences
      for (const follower of notifyList) {
        const baseNotificationData = {
          Title: notificationContent.title,
          NotificationType: notificationTypeMap[updateType] || 'PolicyUpdated',
          RecipientEmail: follower.UserEmail,
          RecipientUserId: follower.UserId,
          PolicyId: policyId,
          PolicyTitle: policyTitle,
          PolicyVersion: version,
          Message: notificationContent.inAppMessage,
          Status: 'Pending',
          Priority: updateType === 'NewVersion' ? 'High' : 'Normal',
          RetryCount: 0,
          MaxRetries: 3,
          RelatedFollowId: follower.Id
        };

        // Email notification
        if (follower.EmailNotifications) {
          await this.sp.web.lists
            .getByTitle(this.NOTIFICATION_QUEUE_LIST)
            .items.add({
              ...baseNotificationData,
              Channel: 'Email'
            });
        }

        // Teams notification
        if (follower.TeamsNotifications) {
          await this.sp.web.lists
            .getByTitle(this.NOTIFICATION_QUEUE_LIST)
            .items.add({
              ...baseNotificationData,
              Channel: 'Teams'
            });
        }

        // In-app notification
        if (follower.InAppNotifications) {
          await this.sp.web.lists
            .getByTitle(this.NOTIFICATION_QUEUE_LIST)
            .items.add({
              ...baseNotificationData,
              Channel: 'InApp'
            });
        }
      }

      logger.info('PolicySocialService', `Queued ${updateType} notifications for ${notifyList.length} followers of policy ${policyId}`);
    } catch (error) {
      logger.error('PolicySocialService', 'Failed to notify followers:', error);
    }
  }

  /**
   * Build notification content for follower updates
   */
  private buildFollowerNotificationContent(
    updateType: 'Update' | 'NewVersion' | 'Comment',
    policyTitle: string,
    policyNumber: string,
    version: string
  ): {
    title: string;
    emailSubject: string;
    emailBody: string;
    teamsTitle: string;
    teamsCard: object;
    inAppTitle: string;
    inAppMessage: string;
  } {
    const policyRef = policyNumber ? ` (${policyNumber})` : '';
    let icon = '📋';
    let actionText = '';
    let description = '';

    switch (updateType) {
      case 'Update':
        icon = '✏️';
        actionText = 'has been updated';
        description = 'The policy content has been modified. Please review the changes.';
        break;
      case 'NewVersion':
        icon = '🆕';
        actionText = `has a new version (v${version})`;
        description = 'A new version of this policy has been published. You may need to re-acknowledge it.';
        break;
      case 'Comment':
        icon = '💬';
        actionText = 'has new comments';
        description = 'New comments have been added to this policy.';
        break;
    }

    const title = `${icon} Policy ${updateType}: ${policyTitle}`;
    const emailSubject = `Policy you follow ${actionText}: ${policyTitle}${policyRef}`;

    let emailBody = `<p>A policy you follow has been updated:</p>`;
    emailBody += `<h3>${policyTitle}${policyRef}</h3>`;
    emailBody += `<p><strong>Update Type:</strong> ${updateType}</p>`;
    emailBody += `<p>${description}</p>`;
    emailBody += `<p><a href="{{PolicyUrl}}" style="background-color:#0078d4;color:white;padding:10px 20px;text-decoration:none;border-radius:4px;">View Policy</a></p>`;
    emailBody += `<hr>`;
    emailBody += `<p style="font-size:12px;color:#666;">You are receiving this because you follow this policy. <a href="{{UnfollowUrl}}">Unfollow</a></p>`;

    const teamsCard = {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: `${icon} Policy ${updateType}`,
          size: 'Large',
          weight: 'Bolder'
        },
        {
          type: 'TextBlock',
          text: `**${policyTitle}${policyRef}**`,
          wrap: true
        },
        {
          type: 'TextBlock',
          text: description,
          wrap: true,
          size: 'Small',
          color: 'Accent'
        }
      ],
      actions: [
        {
          type: 'Action.OpenUrl',
          title: 'View Policy',
          url: '{{PolicyUrl}}',
          style: 'positive'
        }
      ]
    };

    return {
      title,
      emailSubject,
      emailBody,
      teamsTitle: title,
      teamsCard,
      inAppTitle: `${icon} ${policyTitle} ${actionText}`,
      inAppMessage: description
    };
  }

  /**
   * Notify followers of new comment
   */
  private async notifyFollowersOfComment(policyId: number, commentId: number): Promise<void> {
    await this.notifyFollowersOfUpdate(policyId, 'Comment');
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Check if user has acknowledged the policy (verified reader)
   */
  private async isVerifiedReader(policyId: number, userId: number): Promise<boolean> {
    try {
      const acknowledgements = await this.sp.web.lists
        .getByTitle(PolicyLists.POLICY_ACKNOWLEDGEMENTS)
        .items.filter(`PolicyId eq ${policyId} and UserId eq ${userId} and Status eq 'Acknowledged'`)
        .top(1)();

      return acknowledgements.length > 0;
    } catch (error) {
      return false;
    }
  }
}
