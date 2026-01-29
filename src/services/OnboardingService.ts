// @ts-nocheck
// OnboardingService - Manages user onboarding, tutorials, and help system
// Provides interactive tours, contextual help, and What's New announcements

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import {
  IOnboardingTutorial,
  IUserTutorialProgress,
  IContextualHelp,
  IHelpArticle,
  IWhatsNew,
  IUserAnnouncementView,
  IOnboardingChecklist,
  IOnboardingChecklistItem,
  ITutorialStep,
  TutorialStatus,
  TutorialType,
  HelpInteractionType,
  DEFAULT_ONBOARDING_CHECKLIST
} from '../models/IOnboarding';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class OnboardingService {
  private sp: SPFI;
  private currentUserId: string = '';
  private currentUserEmail: string = '';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Initialize service
   */
  public async initialize(): Promise<void> {
    try {
      const currentUser = await this.sp.web.currentUser();
      this.currentUserId = currentUser.Id.toString();
      this.currentUserEmail = currentUser.Email;
    } catch (error) {
      logger.error('OnboardingService', 'Failed to initialize OnboardingService:', error);
    }
  }

  /**
   * Get onboarding checklist for current user
   */
  public async getOnboardingChecklist(): Promise<IOnboardingChecklist> {
    try {
      // Validate user ID
      ValidationUtils.validateUserId(this.currentUserId);

      // Get user progress with secure filter
      const filter = ValidationUtils.buildFilter('UserId', 'eq', this.currentUserId);
      const progress = await this.sp.web.lists
        .getByTitle('PM_UserTutorialProgress')
        .items
        .filter(filter)();

      const completedTutorials = progress
        .filter(p => p.Status === TutorialStatus.Completed)
        .map(p => p.TutorialId);

      // Build checklist
      const items: IOnboardingChecklistItem[] = DEFAULT_ONBOARDING_CHECKLIST.map(item => {
        const completed = this.isChecklistItemCompleted(item, completedTutorials);
        return {
          ...item,
          isCompleted: completed,
          completedDate: completed ? new Date() : undefined
        };
      });

      const completedCount = items.filter(i => i.isCompleted).length;

      return {
        userId: this.currentUserId,
        items,
        completedCount,
        totalCount: items.length,
        percentComplete: Math.round((completedCount / items.length) * 100),
        startedDate: new Date()
      };
    } catch (error) {
      logger.error('OnboardingService', 'Failed to get onboarding checklist:', error);
      return {
        userId: this.currentUserId,
        items: DEFAULT_ONBOARDING_CHECKLIST,
        completedCount: 0,
        totalCount: DEFAULT_ONBOARDING_CHECKLIST.length,
        percentComplete: 0,
        startedDate: new Date()
      };
    }
  }

  /**
   * Get available tutorials for user
   */
  public async getTutorials(type?: TutorialType): Promise<IOnboardingTutorial[]> {
    try {
      let filter = 'IsActive eq true';
      if (type) {
        // Validate enum and build secure filter
        ValidationUtils.validateEnum(type, TutorialType, 'TutorialType');
        filter += ` and ${ValidationUtils.buildFilter('TutorialType', 'eq', type)}`;
      }

      const tutorials = await this.sp.web.lists
        .getByTitle('PM_OnboardingTutorials')
        .items
        .filter(filter)
        .orderBy('Priority')();

      return tutorials.map(t => this.parseTutorial(t));
    } catch (error) {
      logger.error('OnboardingService', 'Failed to get tutorials:', error);
      return [];
    }
  }

  /**
   * Get tutorial by ID
   */
  public async getTutorial(tutorialId: number): Promise<IOnboardingTutorial | null> {
    try {
      const tutorial = await this.sp.web.lists
        .getByTitle('PM_OnboardingTutorials')
        .items
        .getById(tutorialId)();

      return this.parseTutorial(tutorial);
    } catch (error) {
      logger.error('OnboardingService', 'Failed to get tutorial:', error);
      return null;
    }
  }

  /**
   * Start tutorial
   */
  public async startTutorial(tutorialId: number): Promise<IUserTutorialProgress> {
    try {
      // Validate inputs
      const validTutorialId = ValidationUtils.validateInteger(tutorialId, 'tutorialId', 1);
      ValidationUtils.validateUserId(this.currentUserId);

      // Check if already in progress with secure filter
      const userFilter = ValidationUtils.buildFilter('UserId', 'eq', this.currentUserId);
      const tutFilter = ValidationUtils.buildFilter('TutorialId', 'eq', validTutorialId);
      const filter = `${userFilter} and ${tutFilter}`;

      const existing = await this.sp.web.lists
        .getByTitle('PM_UserTutorialProgress')
        .items
        .filter(filter)
        .top(1)();

      if (existing.length > 0) {
        // Resume existing
        return this.parseProgress(existing[0]);
      }

      // Get tutorial info
      const tutorial = await this.getTutorial(tutorialId);
      if (!tutorial) {
        throw new Error('Tutorial not found');
      }

      // Create new progress record
      const progressData = {
        Title: `${tutorial.Title} - ${this.currentUserEmail}`,
        UserId: this.currentUserId,
        UserEmail: this.currentUserEmail,
        TutorialId: tutorialId,
        TutorialType: tutorial.TutorialType,
        Status: TutorialStatus.InProgress,
        CurrentStep: 0,
        CompletedSteps: JSON.stringify([]),
        StartedDate: new Date().toISOString()
      };

      const result = await this.sp.web.lists
        .getByTitle('PM_UserTutorialProgress')
        .items
        .add(progressData);

      // Log interaction
      await this.logHelpInteraction(HelpInteractionType.TutorialStarted, {
        tutorialId
      });

      return this.parseProgress({ ...progressData, Id: result.data.Id });
    } catch (error) {
      logger.error('OnboardingService', 'Failed to start tutorial:', error);
      throw error;
    }
  }

  /**
   * Update tutorial progress
   */
  public async updateTutorialProgress(
    progressId: number,
    currentStep: number,
    completedSteps: number[]
  ): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('PM_UserTutorialProgress')
        .items
        .getById(progressId)
        .update({
          CurrentStep: currentStep,
          CompletedSteps: JSON.stringify(completedSteps)
        });
    } catch (error) {
      logger.error('OnboardingService', 'Failed to update tutorial progress:', error);
      throw error;
    }
  }

  /**
   * Complete tutorial
   */
  public async completeTutorial(progressId: number, timeSpent: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('PM_UserTutorialProgress')
        .items
        .getById(progressId)
        .update({
          Status: TutorialStatus.Completed,
          CompletedDate: new Date().toISOString(),
          TimeSpent: timeSpent
        });

      // Log interaction
      await this.logHelpInteraction(HelpInteractionType.TutorialCompleted, {
        progressId,
        timeSpent
      });
    } catch (error) {
      logger.error('OnboardingService', 'Failed to complete tutorial:', error);
      throw error;
    }
  }

  /**
   * Skip tutorial
   */
  public async skipTutorial(progressId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('PM_UserTutorialProgress')
        .items
        .getById(progressId)
        .update({
          Status: TutorialStatus.Skipped
        });
    } catch (error) {
      logger.error('OnboardingService', 'Failed to skip tutorial:', error);
      throw error;
    }
  }

  /**
   * Get contextual help for page
   */
  public async getContextualHelp(pageUrl: string): Promise<IContextualHelp[]> {
    try {
      // Validate URL to prevent injection
      if (!pageUrl || typeof pageUrl !== 'string') {
        return [];
      }

      // Sanitize pageUrl for OData query
      const sanitizedUrl = ValidationUtils.sanitizeForOData(pageUrl);

      const helps = await this.sp.web.lists
        .getByTitle('PM_ContextualHelp')
        .items
        .filter(`IsActive eq true and substringof('${sanitizedUrl}', PageUrl)`)
        .orderBy('Priority')();

      return helps;
    } catch (error) {
      logger.error('OnboardingService', 'Failed to get contextual help:', error);
      return [];
    }
  }

  /**
   * Search help articles
   */
  public async searchHelpArticles(query: string): Promise<IHelpArticle[]> {
    try {
      // Validate and sanitize search query
      if (!query || typeof query !== 'string' || query.trim().length === 0) {
        return [];
      }

      // Limit query length and sanitize for OData
      const sanitizedQuery = ValidationUtils.sanitizeForOData(query.substring(0, 100));

      const articles = await this.sp.web.lists
        .getByTitle('PM_HelpArticles')
        .items
        .filter(`IsPublished eq true and (substringof('${sanitizedQuery}', Title) or substringof('${sanitizedQuery}', Content))`)
        .orderBy('ViewCount', false)
        .top(10)();

      // Log search with sanitized query
      await this.logHelpInteraction(HelpInteractionType.SearchPerformed, {
        query: ValidationUtils.sanitizeHtml(query),
        resultsCount: articles.length
      });

      return articles;
    } catch (error) {
      logger.error('OnboardingService', 'Failed to search help articles:', error);
      return [];
    }
  }

  /**
   * Get help article by ID
   */
  public async getHelpArticle(articleId: number): Promise<IHelpArticle | null> {
    try {
      const article = await this.sp.web.lists
        .getByTitle('PM_HelpArticles')
        .items
        .getById(articleId)();

      // Increment view count
      await this.sp.web.lists
        .getByTitle('PM_HelpArticles')
        .items
        .getById(articleId)
        .update({
          ViewCount: (article.ViewCount || 0) + 1
        });

      // Log interaction
      await this.logHelpInteraction(HelpInteractionType.ArticleViewed, {
        articleId
      });

      return article;
    } catch (error) {
      logger.error('OnboardingService', 'Failed to get help article:', error);
      return null;
    }
  }

  /**
   * Get What's New announcements
   */
  public async getWhatsNew(): Promise<IWhatsNew[]> {
    try {
      const now = new Date().toISOString();

      const announcements = await this.sp.web.lists
        .getByTitle('PM_WhatsNew')
        .items
        .filter(`IsActive eq true and (ShowUntil eq null or ShowUntil ge datetime'${now}')`)
        .orderBy('ReleaseDate', false)
        .top(5)();

      return announcements;
    } catch (error) {
      logger.error('OnboardingService', 'Failed to get What\'s New:', error);
      return [];
    }
  }

  /**
   * Get unread What's New for user
   */
  public async getUnreadAnnouncements(): Promise<IWhatsNew[]> {
    try {
      // Get all active announcements
      const announcements = await this.getWhatsNew();

      // Validate user ID
      ValidationUtils.validateUserId(this.currentUserId);

      // Get user views with secure filter
      const filter = `${ValidationUtils.buildFilter('UserId', 'eq', this.currentUserId)} and Dismissed eq true`;
      const views = await this.sp.web.lists
        .getByTitle('PM_UserAnnouncementViews')
        .items
        .filter(filter)();

      const viewedIds = views.map(v => v.AnnouncementId);

      // Filter out viewed/dismissed
      return announcements.filter(a => !viewedIds.includes(a.Id));
    } catch (error) {
      logger.error('OnboardingService', 'Failed to get unread announcements:', error);
      return [];
    }
  }

  /**
   * Mark announcement as viewed
   */
  public async markAnnouncementViewed(announcementId: number): Promise<void> {
    try {
      // Validate inputs
      const validAnnouncementId = ValidationUtils.validateInteger(announcementId, 'announcementId', 1);
      ValidationUtils.validateUserId(this.currentUserId);

      // Check if already viewed with secure filter
      const userFilter = ValidationUtils.buildFilter('UserId', 'eq', this.currentUserId);
      const annFilter = ValidationUtils.buildFilter('AnnouncementId', 'eq', validAnnouncementId);
      const filter = `${userFilter} and ${annFilter}`;

      const existing = await this.sp.web.lists
        .getByTitle('PM_UserAnnouncementViews')
        .items
        .filter(filter)
        .top(1)();

      if (existing.length === 0) {
        await this.sp.web.lists
          .getByTitle('PM_UserAnnouncementViews')
          .items
          .add({
            Title: `${announcementId} - ${this.currentUserEmail}`,
            UserId: this.currentUserId,
            AnnouncementId: announcementId,
            ViewedDate: new Date().toISOString(),
            Dismissed: false
          });
      }
    } catch (error) {
      logger.error('OnboardingService', 'Failed to mark announcement as viewed:', error);
    }
  }

  /**
   * Dismiss announcement
   */
  public async dismissAnnouncement(announcementId: number, rating?: number, comment?: string): Promise<void> {
    try {
      // Validate inputs
      const validAnnouncementId = ValidationUtils.validateInteger(announcementId, 'announcementId', 1);
      ValidationUtils.validateUserId(this.currentUserId);

      if (rating !== undefined) {
        ValidationUtils.validateInteger(rating, 'rating', 1, 5);
      }

      // Sanitize comment to prevent XSS
      const sanitizedComment = comment ? ValidationUtils.sanitizeHtml(comment) : undefined;

      // Build secure filter
      const userFilter = ValidationUtils.buildFilter('UserId', 'eq', this.currentUserId);
      const annFilter = ValidationUtils.buildFilter('AnnouncementId', 'eq', validAnnouncementId);
      const filter = `${userFilter} and ${annFilter}`;

      const existing = await this.sp.web.lists
        .getByTitle('PM_UserAnnouncementViews')
        .items
        .filter(filter)
        .top(1)();

      if (existing.length > 0) {
        await this.sp.web.lists
          .getByTitle('PM_UserAnnouncementViews')
          .items
          .getById(existing[0].Id)
          .update({
            Dismissed: true,
            DismissedDate: new Date().toISOString(),
            FeedbackRating: rating,
            FeedbackComment: sanitizedComment
          });
      } else {
        await this.sp.web.lists
          .getByTitle('PM_UserAnnouncementViews')
          .items
          .add({
            Title: `${validAnnouncementId} - ${this.currentUserEmail}`,
            UserId: this.currentUserId,
            AnnouncementId: validAnnouncementId,
            ViewedDate: new Date().toISOString(),
            Dismissed: true,
            DismissedDate: new Date().toISOString(),
            FeedbackRating: rating,
            FeedbackComment: sanitizedComment
          });
      }
    } catch (error) {
      logger.error('OnboardingService', 'Failed to dismiss announcement:', error);
    }
  }

  /**
   * Check if user should see onboarding
   */
  public async shouldShowOnboarding(): Promise<boolean> {
    try {
      // Validate user ID
      ValidationUtils.validateUserId(this.currentUserId);

      // Check if user has completed any tutorials with secure filter
      const filter = ValidationUtils.buildFilter('UserId', 'eq', this.currentUserId);
      const progress = await this.sp.web.lists
        .getByTitle('PM_UserTutorialProgress')
        .items
        .filter(filter)
        .top(1)();

      return progress.length === 0;
    } catch (error) {
      return true; // Show onboarding if we can't determine
    }
  }

  /**
   * Mark checklist item as completed
   */
  public async markChecklistItemCompleted(itemId: string): Promise<void> {
    try {
      // Store in user preferences or separate list
      await this.logHelpInteraction(HelpInteractionType.TutorialCompleted, {
        checklistItemId: itemId
      });
    } catch (error) {
      logger.error('OnboardingService', 'Failed to mark checklist item completed:', error);
    }
  }

  /**
   * Install sample data
   */
  public async installSampleData(): Promise<void> {
    try {
      // Get sample data template
      const templates = await this.sp.web.lists
        .getByTitle('PM_SampleDataTemplates')
        .items
        .filter('IsActive eq true and TemplateType eq \'Full Demo\'')
        .top(1)();

      if (templates.length === 0) {
        throw new Error('No sample data templates available');
      }

      const template = templates[0];
      const dataSet = JSON.parse(template.DataSet);

      // Install sample processes, tasks, templates, etc.
      // This would create the sample data in the respective lists

      // Log installation
      await this.logHelpInteraction(HelpInteractionType.SampleDataInstalled, {
        templateId: template.Id
      });

      // Update install count
      await this.sp.web.lists
        .getByTitle('PM_SampleDataTemplates')
        .items
        .getById(template.Id)
        .update({
          InstallCount: (template.InstallCount || 0) + 1,
          LastUsed: new Date().toISOString()
        });
    } catch (error) {
      logger.error('OnboardingService', 'Failed to install sample data:', error);
      throw error;
    }
  }

  /**
   * Submit help feedback
   */
  public async submitHelpFeedback(
    articleId: number,
    wasHelpful: boolean,
    comment?: string
  ): Promise<void> {
    try {
      // Update article votes
      const article = await this.sp.web.lists
        .getByTitle('PM_HelpArticles')
        .items
        .getById(articleId)();

      await this.sp.web.lists
        .getByTitle('PM_HelpArticles')
        .items
        .getById(articleId)
        .update({
          HelpfulVotes: wasHelpful ? (article.HelpfulVotes || 0) + 1 : article.HelpfulVotes,
          UnhelpfulVotes: !wasHelpful ? (article.UnhelpfulVotes || 0) + 1 : article.UnhelpfulVotes
        });

      // Log feedback
      await this.logHelpInteraction(HelpInteractionType.FeedbackSubmitted, {
        articleId,
        wasHelpful,
        comment
      });
    } catch (error) {
      logger.error('OnboardingService', 'Failed to submit help feedback:', error);
    }
  }

  /**
   * Parse tutorial from SharePoint item
   */
  private parseTutorial(item: any): IOnboardingTutorial {
    return {
      ...item,
      Steps: item.Steps ? JSON.parse(item.Steps) : []
    };
  }

  /**
   * Parse progress from SharePoint item
   */
  private parseProgress(item: any): IUserTutorialProgress {
    return {
      ...item,
      CompletedSteps: item.CompletedSteps ? JSON.parse(item.CompletedSteps) : [],
      SkippedSteps: item.SkippedSteps ? JSON.parse(item.SkippedSteps) : []
    };
  }

  /**
   * Check if checklist item is completed
   */
  private isChecklistItemCompleted(
    item: IOnboardingChecklistItem,
    completedTutorials: number[]
  ): boolean {
    if (item.tutorialId) {
      return completedTutorials.includes(item.tutorialId);
    }
    // Add other completion criteria here
    return false;
  }

  /**
   * Log help interaction
   */
  private async logHelpInteraction(
    interactionType: HelpInteractionType,
    metadata?: any
  ): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('PM_UserHelpInteractions')
        .items
        .add({
          Title: `${interactionType} - ${new Date().toISOString()}`,
          UserId: this.currentUserId,
          InteractionType: interactionType,
          HelpArticleId: metadata?.articleId,
          TutorialId: metadata?.tutorialId,
          TooltipId: metadata?.tooltipId,
          SearchQuery: metadata?.query,
          WasHelpful: metadata?.wasHelpful,
          FeedbackComment: metadata?.comment,
          TimeSpent: metadata?.timeSpent,
          PageUrl: window.location.href
        });
    } catch (error) {
      logger.error('OnboardingService', 'Failed to log help interaction:', error);
    }
  }
}
