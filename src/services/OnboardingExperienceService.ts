// @ts-nocheck
/**
 * OnboardingExperienceService
 * Provides data and operations for the JML Onboarding Experience module
 * Uses SharePoint lists for production data with mock fallback for workbench
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';

import {
  INavigationItem,
  IQuestStage,
  IQuestProgress,
  IQuestTask,
  ITeamMember,
  IFloorPlan,
  IOfficeLocation,
  ISurvivalItem,
  ICoffeeLesson,
  IQuizQuestion,
  IWelcomeMessage,
  IJargonTerm,
  IInterestGroup,
  IBingoChallenge,
  IBingoBoard,
  IOnboardingProgress,
  IOnboardingBadge,
  ILeaderboardEntry,
  ILeaderboardStats,
  OnboardingSection
} from '../webparts/jmlOnboardingExperience/models/IOnboardingExperience';

// SharePoint List Names
const LIST_NAMES = {
  PROGRESS: 'JML_OnboardingProgress',
  QUEST_STAGES: 'JML_OnboardingQuestStages',
  TEAM_MEMBERS: 'JML_OnboardingTeamMembers',
  FLOOR_PLANS: 'JML_OnboardingFloorPlans',
  SURVIVAL_KIT: 'JML_OnboardingSurvivalKit',
  WELCOME_WALL: 'JML_OnboardingWelcomeWall',
  JARGON: 'JML_OnboardingJargon',
  INTEREST_GROUPS: 'JML_OnboardingInterestGroups',
  BINGO: 'JML_OnboardingBingo',
  BADGES: 'JML_OnboardingBadges',
  COFFEE_LESSONS: 'JML_OnboardingCoffeeLessons'
};

export interface IOnboardingExperienceService {
  // Navigation
  getNavigationItems(): Promise<INavigationItem[]>;

  // Quest
  getQuestStages(): Promise<IQuestStage[]>;
  getQuestProgress(userId: string): Promise<IQuestProgress>;
  completeQuestTask(userId: string, taskId: string): Promise<void>;

  // Who's Who
  getTeamMembers(): Promise<ITeamMember[]>;
  getKeyContacts(userId: string): Promise<ITeamMember[]>;

  // Office Explorer
  getFloorPlans(): Promise<IFloorPlan[]>;

  // Survival Kit
  getSurvivalItems(): Promise<ISurvivalItem[]>;
  completeSurvivalItem(userId: string, itemId: string): Promise<void>;

  // Coffee Academy
  getCoffeeLessons(): Promise<ICoffeeLesson[]>;
  completeCoffeeLesson(userId: string, lessonId: string): Promise<void>;

  // Welcome Wall
  getWelcomeMessages(): Promise<IWelcomeMessage[]>;
  addWelcomeMessage(message: string): Promise<void>;
  likeWelcomeMessage(messageId: string): Promise<void>;

  // Jargon Buster
  getJargonTerms(): Promise<IJargonTerm[]>;

  // Interest Groups
  getInterestGroups(): Promise<IInterestGroup[]>;
  joinInterestGroup(userId: string, groupId: string): Promise<void>;
  leaveInterestGroup(userId: string, groupId: string): Promise<void>;

  // Bingo
  getBingoBoard(userId: string): Promise<IBingoBoard>;
  completeBingoChallenge(userId: string, challengeId: string): Promise<void>;

  // Progress
  getOnboardingProgress(userId: string): Promise<IOnboardingProgress>;
  updateOnboardingProgress(userId: string, progress: Partial<IOnboardingProgress>): Promise<void>;
  addXp(userId: string, xp: number, source: string): Promise<void>;

  // Badges
  getAvailableBadges(): Promise<IOnboardingBadge[]>;
  getUserBadges(userId: string): Promise<IOnboardingBadge[]>;
  awardBadge(userId: string, badgeId: string): Promise<void>;

  // Leaderboard
  getLeaderboard(period: string): Promise<ILeaderboardEntry[]>;
  getLeaderboardStats(): Promise<ILeaderboardStats>;
}

export class OnboardingExperienceService implements IOnboardingExperienceService {
  private readonly sp: SPFI;
  private readonly siteUrl: string;
  private readonly isWorkbench: boolean;
  private readonly userEmail: string;
  private readonly userDisplayName: string;
  private readonly userTitle: string;

  constructor(sp: SPFI, siteUrl: string, userEmail: string = '', userDisplayName: string = '', userTitle: string = '') {
    this.sp = sp;
    this.siteUrl = siteUrl;
    this.isWorkbench = siteUrl.indexOf('workbench') > -1 || siteUrl.indexOf('localhost') > -1;
    this.userEmail = userEmail;
    this.userDisplayName = userDisplayName;
    this.userTitle = userTitle || 'Team Member';
  }

  // ============================================
  // NAVIGATION
  // ============================================

  public async getNavigationItems(): Promise<INavigationItem[]> {
    // Navigation is static - return fixed items
    return [
      { id: 'welcome', title: 'Welcome Hub', icon: 'Home', description: 'Your onboarding home base', completed: false, locked: false, xpReward: 50 },
      { id: 'quest', title: 'Onboarding Quest', icon: 'Trophy', description: 'Complete missions to earn XP and badges', completed: false, locked: false, xpReward: 500 },
      { id: 'whos-who', title: "Who's Who", icon: 'People', description: 'Meet your team and key contacts', completed: false, locked: false, xpReward: 100 },
      { id: 'office-explorer', title: 'Office Explorer', icon: 'MapPin', description: 'Navigate your new workspace', completed: false, locked: false, xpReward: 75 },
      { id: 'survival-kit', title: 'Survival Kit', icon: 'FirstAid', description: 'Essential first day checklist', completed: false, locked: false, xpReward: 150 },
      { id: 'coffee-academy', title: 'Coffee Academy', icon: 'CoffeeScript', description: 'Master the office coffee machine', completed: false, locked: false, xpReward: 50 },
      { id: 'welcome-wall', title: 'Welcome Wall', icon: 'Comment', description: 'Messages from your new colleagues', completed: false, locked: false, xpReward: 25 },
      { id: 'buddy-bot', title: 'Buddy Bot', icon: 'Robot', description: 'Your AI onboarding assistant', completed: false, locked: false, xpReward: 0 },
      { id: 'jargon-buster', title: 'Jargon Buster', icon: 'Dictionary', description: 'Decode company acronyms and terms', completed: false, locked: false, xpReward: 50 },
      { id: 'interest-groups', title: 'Interest Groups', icon: 'Group', description: 'Find your tribe', completed: false, locked: false, xpReward: 75 },
      { id: 'bingo', title: 'First Week Bingo', icon: 'GridViewSmall', description: 'Complete challenges for rewards', completed: false, locked: false, xpReward: 250 },
      { id: 'leaderboard', title: 'Leaderboard', icon: 'Trophy', description: 'See top performers and your ranking', completed: false, locked: false, xpReward: 0 }
    ];
  }

  // ============================================
  // QUEST SYSTEM
  // ============================================

  public async getQuestStages(): Promise<IQuestStage[]> {
    if (this.isWorkbench) {
      return this.getMockQuestStages();
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.QUEST_STAGES)
        .items
        .filter("IsActive eq 1")
        .orderBy("SortOrder", true)
        .select("Id", "Title", "StageID", "StageNumber", "Description", "Icon", "XPReward", "BadgeID", "TasksJSON")();

      return items.map((item: Record<string, unknown>) => {
        let tasks: IQuestTask[] = [];
        try {
          tasks = JSON.parse(item.TasksJSON as string || '[]');
        } catch (e) {
          console.warn('Failed to parse tasks JSON for stage:', item.StageID);
        }

        return {
          id: item.StageID as string,
          title: item.Title as string,
          description: item.Description as string || '',
          icon: item.Icon as string || 'Flag',
          xpReward: item.XPReward as number || 0,
          badgeId: item.BadgeID as string,
          unlocked: (item.StageNumber as number) <= 2,
          completed: false,
          tasks
        };
      });
    } catch (error) {
      console.error('Failed to load quest stages from SharePoint:', error);
      return this.getMockQuestStages();
    }
  }

  public async getQuestProgress(userId: string): Promise<IQuestProgress> {
    if (this.isWorkbench) {
      return this.getDefaultQuestProgress();
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.PROGRESS)
        .items
        .filter(`UserEmail eq '${userId}'`)
        .select("CurrentStage", "CompletedTasks", "TotalXP", "BadgesEarned", "StreakDays", "LastActivity")
        .top(1)();

      if (items.length > 0) {
        const item = items[0];
        return {
          currentStage: item.CurrentStage || 1,
          totalStages: 5,
          completedTasks: this.parseJsonArray(item.CompletedTasks),
          totalXp: item.TotalXP || 0,
          badges: this.parseJsonArray(item.BadgesEarned),
          streakDays: item.StreakDays || 1,
          lastActivityDate: item.LastActivity || new Date().toISOString()
        };
      }
    } catch (error) {
      console.error('Failed to load quest progress:', error);
    }

    return this.getDefaultQuestProgress();
  }

  private getDefaultQuestProgress(): IQuestProgress {
    return {
      currentStage: 1,
      totalStages: 5,
      completedTasks: [],
      totalXp: 0,
      badges: [],
      streakDays: 1,
      lastActivityDate: new Date().toISOString()
    };
  }

  public async completeQuestTask(userId: string, taskId: string): Promise<void> {
    if (this.isWorkbench) {
      console.log(`[Workbench] Completing task ${taskId} for user ${userId}`);
      return;
    }

    try {
      const progress = await this.getQuestProgress(userId);
      if (!progress.completedTasks.includes(taskId)) {
        progress.completedTasks.push(taskId);
        await this.updateProgressField(userId, 'CompletedTasks', JSON.stringify(progress.completedTasks));
      }
    } catch (error) {
      console.error('Failed to complete quest task:', error);
    }
  }

  // ============================================
  // WHO'S WHO
  // ============================================

  public async getTeamMembers(): Promise<ITeamMember[]> {
    if (this.isWorkbench) {
      return this.getMockTeamMembers();
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.TEAM_MEMBERS)
        .items
        .filter("IsActive eq 1")
        .orderBy("IsKeyContact", false)
        .orderBy("FullName", true)
        .select("Id", "FullName", "JobTitle", "Department", "Email", "PhotoURL", "FunFact", "Expertise", "IsKeyContact", "ContactType", "Location", "LinkedInURL")();

      return items.map((item: Record<string, unknown>) => ({
        id: String(item.Id),
        name: item.FullName as string,
        title: item.JobTitle as string || '',
        department: item.Department as string || '',
        email: item.Email as string || '',
        photoUrl: this.extractUrl(item.PhotoURL) || '',
        funFact: item.FunFact as string || '',
        expertise: this.parseExpertise(item.Expertise as string),
        isKeyContact: item.IsKeyContact as boolean || false,
        contactType: item.ContactType as 'manager' | 'buddy' | 'hr' | 'it' | 'team',
        location: item.Location as string || '',
        linkedInUrl: this.extractUrl(item.LinkedInURL)
      }));
    } catch (error) {
      console.error('Failed to load team members from SharePoint:', error);
      return this.getMockTeamMembers();
    }
  }

  public async getKeyContacts(userId: string): Promise<ITeamMember[]> {
    const allMembers = await this.getTeamMembers();
    return allMembers.filter(m => m.isKeyContact);
  }

  // ============================================
  // OFFICE EXPLORER
  // ============================================

  public async getFloorPlans(): Promise<IFloorPlan[]> {
    if (this.isWorkbench) {
      return this.getMockFloorPlans();
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.FLOOR_PLANS)
        .items
        .filter("IsActive eq 1")
        .orderBy("SortOrder", true)
        .select("Id", "Title", "FloorNumber", "FloorName", "ImageURL", "LocationsJSON")();

      return items.map((item: Record<string, unknown>) => {
        let locations: IOfficeLocation[] = [];
        try {
          locations = JSON.parse(item.LocationsJSON as string || '[]');
        } catch (e) {
          console.warn('Failed to parse locations JSON for floor:', item.FloorNumber);
        }

        return {
          floor: item.FloorNumber as number,
          name: item.FloorName as string || item.Title as string,
          imageUrl: this.extractUrl(item.ImageURL) || '',
          locations
        };
      });
    } catch (error) {
      console.error('Failed to load floor plans from SharePoint:', error);
      return this.getMockFloorPlans();
    }
  }

  // ============================================
  // SURVIVAL KIT
  // ============================================

  public async getSurvivalItems(): Promise<ISurvivalItem[]> {
    if (this.isWorkbench) {
      return this.getMockSurvivalItems();
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.SURVIVAL_KIT)
        .items
        .filter("IsActive eq 1")
        .orderBy("SortOrder", true)
        .select("Id", "Title", "ItemID", "Description", "Icon", "Category", "Priority", "ActionURL", "HelpText", "XPReward")();

      return items.map((item: Record<string, unknown>) => ({
        id: item.ItemID as string || String(item.Id),
        title: item.Title as string,
        description: item.Description as string || '',
        icon: item.Icon as string || 'CheckMark',
        category: item.Category as 'day-one' | 'first-week' | 'first-month',
        completed: false,
        priority: item.Priority as 'high' | 'medium' | 'low' || 'medium',
        actionUrl: this.extractUrl(item.ActionURL),
        helpText: item.HelpText as string
      }));
    } catch (error) {
      console.error('Failed to load survival items from SharePoint:', error);
      return this.getMockSurvivalItems();
    }
  }

  public async completeSurvivalItem(userId: string, itemId: string): Promise<void> {
    if (this.isWorkbench) {
      console.log(`[Workbench] Completing survival item ${itemId} for user ${userId}`);
      return;
    }

    try {
      const progress = await this.getOnboardingProgress(userId);
      if (!progress.survivalKitProgress.includes(itemId)) {
        progress.survivalKitProgress.push(itemId);
        await this.updateProgressField(userId, 'SurvivalKitProgress', JSON.stringify(progress.survivalKitProgress));
      }
    } catch (error) {
      console.error('Failed to complete survival item:', error);
    }
  }

  // ============================================
  // COFFEE ACADEMY
  // ============================================

  public async getCoffeeLessons(): Promise<ICoffeeLesson[]> {
    if (this.isWorkbench) {
      return this.getMockCoffeeLessons();
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.COFFEE_LESSONS)
        .items
        .filter("IsActive eq 1")
        .orderBy("SortOrder", true)
        .select("Id", "Title", "LessonID", "Description", "LessonType", "Duration", "XPReward", "Content", "QuizQuestionsJSON", "VideoURL")();

      return items.map((item: Record<string, unknown>) => {
        let quizQuestions: IQuizQuestion[] | undefined;
        if (item.QuizQuestionsJSON) {
          try {
            quizQuestions = JSON.parse(item.QuizQuestionsJSON as string);
          } catch (e) {
            console.warn('Failed to parse quiz questions for lesson:', item.LessonID);
          }
        }

        return {
          id: item.LessonID as string || String(item.Id),
          title: item.Title as string,
          description: item.Description as string || '',
          type: item.LessonType as 'video' | 'interactive' | 'quiz' || 'interactive',
          duration: item.Duration as string || '5 min',
          completed: false,
          xpReward: item.XPReward as number || 10,
          content: item.Content as string,
          quizQuestions
        };
      });
    } catch (error) {
      console.error('Failed to load coffee lessons from SharePoint:', error);
      return this.getMockCoffeeLessons();
    }
  }

  public async completeCoffeeLesson(userId: string, lessonId: string): Promise<void> {
    console.log(`Completing coffee lesson ${lessonId} for user ${userId}`);
  }

  // ============================================
  // WELCOME WALL
  // ============================================

  public async getWelcomeMessages(): Promise<IWelcomeMessage[]> {
    if (this.isWorkbench) {
      return this.getMockWelcomeMessages();
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.WELCOME_WALL)
        .items
        .filter("IsActive eq 1")
        .orderBy("IsPinned", false)
        .orderBy("Created", false)
        .top(50)
        .select("Id", "AuthorName", "AuthorTitle", "AuthorPhotoURL", "Message", "Created", "Likes", "LikedBy", "IsPinned")();

      return items.map((item: Record<string, unknown>) => ({
        id: String(item.Id),
        authorName: item.AuthorName as string,
        authorTitle: item.AuthorTitle as string || '',
        authorPhotoUrl: this.extractUrl(item.AuthorPhotoURL) || '',
        message: item.Message as string,
        timestamp: item.Created as string,
        likes: item.Likes as number || 0,
        hasLiked: false
      }));
    } catch (error) {
      console.error('Failed to load welcome messages from SharePoint:', error);
      return this.getMockWelcomeMessages();
    }
  }

  public async addWelcomeMessage(message: string): Promise<void> {
    if (this.isWorkbench) {
      console.log('[Workbench] Adding welcome message:', message);
      return;
    }

    try {
      await this.sp.web.lists.getByTitle(LIST_NAMES.WELCOME_WALL).items.add({
        Title: `Message from ${this.userDisplayName}`,
        AuthorName: this.userDisplayName,
        AuthorTitle: this.userTitle,
        AuthorEmail: this.userEmail,
        Message: message,
        Likes: 0,
        IsActive: true
      });
    } catch (error) {
      console.error('Failed to add welcome message:', error);
      throw error;
    }
  }

  public async likeWelcomeMessage(messageId: string): Promise<void> {
    if (this.isWorkbench) {
      console.log('[Workbench] Liking message:', messageId);
      return;
    }

    try {
      const item = await this.sp.web.lists
        .getByTitle(LIST_NAMES.WELCOME_WALL)
        .items.getById(parseInt(messageId, 10))
        .select("Likes", "LikedBy")();

      const currentLikes = item.Likes || 0;
      const likedBy = this.parseJsonArray(item.LikedBy);

      if (!likedBy.includes(this.userEmail)) {
        likedBy.push(this.userEmail);
        await this.sp.web.lists
          .getByTitle(LIST_NAMES.WELCOME_WALL)
          .items.getById(parseInt(messageId, 10))
          .update({
            Likes: currentLikes + 1,
            LikedBy: JSON.stringify(likedBy)
          });
      }
    } catch (error) {
      console.error('Failed to like message:', error);
    }
  }

  // ============================================
  // JARGON BUSTER
  // ============================================

  public async getJargonTerms(): Promise<IJargonTerm[]> {
    if (this.isWorkbench) {
      return this.getMockJargonTerms();
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.JARGON)
        .items
        .filter("IsActive eq 1")
        .orderBy("Term", true)
        .select("Id", "Term", "Definition", "Category", "Example", "RelatedTerms", "Pronunciation")();

      return items.map((item: Record<string, unknown>) => ({
        id: String(item.Id),
        term: item.Term as string,
        definition: item.Definition as string,
        category: item.Category as string || 'General',
        example: item.Example as string,
        relatedTerms: this.parseCommaSeparated(item.RelatedTerms as string),
        pronunciation: item.Pronunciation as string
      }));
    } catch (error) {
      console.error('Failed to load jargon terms from SharePoint:', error);
      return this.getMockJargonTerms();
    }
  }

  // ============================================
  // INTEREST GROUPS
  // ============================================

  public async getInterestGroups(): Promise<IInterestGroup[]> {
    if (this.isWorkbench) {
      return this.getMockInterestGroups();
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.INTEREST_GROUPS)
        .items
        .filter("IsActive eq 1")
        .orderBy("MemberCount", false)
        .select("Id", "Title", "GroupName", "Description", "Category", "Icon", "MemberCount", "Members", "NextMeeting", "ContactPerson", "ContactEmail", "Tags")();

      return items.map((item: Record<string, unknown>) => ({
        id: String(item.Id),
        name: item.GroupName as string || item.Title as string,
        description: item.Description as string || '',
        category: item.Category as 'sports' | 'hobbies' | 'professional' | 'social' | 'wellness' | 'volunteer',
        iconUrl: item.Icon as string || 'Group',
        memberCount: item.MemberCount as number || 0,
        nextMeeting: item.NextMeeting as string,
        contactPerson: item.ContactPerson as string || '',
        isJoined: false,
        tags: this.parseCommaSeparated(item.Tags as string)
      }));
    } catch (error) {
      console.error('Failed to load interest groups from SharePoint:', error);
      return this.getMockInterestGroups();
    }
  }

  public async joinInterestGroup(userId: string, groupId: string): Promise<void> {
    if (this.isWorkbench) {
      console.log(`[Workbench] User ${userId} joining group ${groupId}`);
      return;
    }

    try {
      const item = await this.sp.web.lists
        .getByTitle(LIST_NAMES.INTEREST_GROUPS)
        .items.getById(parseInt(groupId, 10))
        .select("MemberCount", "Members")();

      const members = this.parseJsonArray(item.Members);
      if (!members.includes(userId)) {
        members.push(userId);
        await this.sp.web.lists
          .getByTitle(LIST_NAMES.INTEREST_GROUPS)
          .items.getById(parseInt(groupId, 10))
          .update({
            MemberCount: (item.MemberCount || 0) + 1,
            Members: JSON.stringify(members)
          });
      }
    } catch (error) {
      console.error('Failed to join interest group:', error);
    }
  }

  public async leaveInterestGroup(userId: string, groupId: string): Promise<void> {
    if (this.isWorkbench) {
      console.log(`[Workbench] User ${userId} leaving group ${groupId}`);
      return;
    }

    try {
      const item = await this.sp.web.lists
        .getByTitle(LIST_NAMES.INTEREST_GROUPS)
        .items.getById(parseInt(groupId, 10))
        .select("MemberCount", "Members")();

      const members = this.parseJsonArray(item.Members);
      const index = members.indexOf(userId);
      if (index > -1) {
        members.splice(index, 1);
        await this.sp.web.lists
          .getByTitle(LIST_NAMES.INTEREST_GROUPS)
          .items.getById(parseInt(groupId, 10))
          .update({
            MemberCount: Math.max((item.MemberCount || 1) - 1, 0),
            Members: JSON.stringify(members)
          });
      }
    } catch (error) {
      console.error('Failed to leave interest group:', error);
    }
  }

  // ============================================
  // BINGO
  // ============================================

  public async getBingoBoard(userId: string): Promise<IBingoBoard> {
    const challenges = await this.getBingoChallenges();
    const progress = await this.getOnboardingProgress(userId);
    const completedChallenges = this.parseJsonArray(progress.bingoBoard?.challenges?.toString() || '');

    // Mark completed challenges
    challenges.forEach(c => {
      c.completed = completedChallenges.includes(c.id) || c.id === 'b-13'; // Free space is always complete
    });

    return {
      challenges,
      completedLines: this.calculateBingoLines(challenges),
      bingoAchieved: this.calculateBingoLines(challenges) > 0,
      totalXpEarned: challenges.filter(c => c.completed).reduce((sum, c) => sum + c.xpReward, 0)
    };
  }

  private async getBingoChallenges(): Promise<IBingoChallenge[]> {
    if (this.isWorkbench) {
      return this.getMockBingoChallenges();
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.BINGO)
        .items
        .filter("IsActive eq 1")
        .orderBy("GridPosition", true)
        .select("Id", "Title", "ChallengeID", "Description", "Icon", "XPReward", "VerificationMethod", "Category", "GridPosition", "IsFreeSpace")();

      return items.map((item: Record<string, unknown>) => ({
        id: item.ChallengeID as string || String(item.Id),
        title: item.Title as string,
        description: item.Description as string || '',
        icon: item.Icon as string || 'CheckMark',
        completed: item.IsFreeSpace as boolean || false,
        xpReward: item.XPReward as number || 10,
        verificationMethod: item.VerificationMethod as 'self' | 'photo' | 'colleague' | 'auto' || 'self',
        category: item.Category as 'social' | 'explore' | 'learn' | 'fun' || 'fun'
      }));
    } catch (error) {
      console.error('Failed to load bingo challenges from SharePoint:', error);
      return this.getMockBingoChallenges();
    }
  }

  private calculateBingoLines(challenges: IBingoChallenge[]): number {
    // Create 5x5 grid
    const grid: boolean[][] = [];
    for (let i = 0; i < 5; i++) {
      grid[i] = [];
      for (let j = 0; j < 5; j++) {
        const idx = i * 5 + j;
        grid[i][j] = challenges[idx]?.completed || false;
      }
    }

    let lines = 0;

    // Check rows
    for (let i = 0; i < 5; i++) {
      if (grid[i].every(cell => cell)) lines++;
    }

    // Check columns
    for (let j = 0; j < 5; j++) {
      if (grid.every(row => row[j])) lines++;
    }

    // Check diagonals
    if ([0, 1, 2, 3, 4].every(i => grid[i][i])) lines++;
    if ([0, 1, 2, 3, 4].every(i => grid[i][4 - i])) lines++;

    return lines;
  }

  public async completeBingoChallenge(userId: string, challengeId: string): Promise<void> {
    if (this.isWorkbench) {
      console.log(`[Workbench] Completing bingo challenge ${challengeId} for user ${userId}`);
      return;
    }

    try {
      const progress = await this.getOnboardingProgress(userId);
      const completed = this.parseJsonArray(progress.bingoBoard?.challenges?.toString() || '');
      if (!completed.includes(challengeId)) {
        completed.push(challengeId);
        await this.updateProgressField(userId, 'BingoProgress', JSON.stringify(completed));
      }
    } catch (error) {
      console.error('Failed to complete bingo challenge:', error);
    }
  }

  // ============================================
  // PROGRESS
  // ============================================

  public async getOnboardingProgress(userId: string): Promise<IOnboardingProgress> {
    if (this.isWorkbench) {
      return this.getDefaultProgress(userId);
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.PROGRESS)
        .items
        .filter(`UserEmail eq '${userId}'`)
        .top(1)();

      if (items.length > 0) {
        const item = items[0];
        return {
          userId: item.UserEmail,
          startDate: item.StartDate || item.Created,
          questProgress: {
            currentStage: item.CurrentStage || 1,
            totalStages: 5,
            completedTasks: this.parseJsonArray(item.CompletedTasks),
            totalXp: item.TotalXP || 0,
            badges: this.parseJsonArray(item.BadgesEarned),
            streakDays: item.StreakDays || 1,
            lastActivityDate: item.LastActivity || new Date().toISOString()
          },
          completedSections: this.parseJsonArray(item.CompletedSections) as OnboardingSection[],
          survivalKitProgress: this.parseJsonArray(item.SurvivalKitProgress),
          bingoBoard: {
            challenges: [],
            completedLines: 0,
            bingoAchieved: false,
            totalXpEarned: 0
          },
          totalXp: item.TotalXP || 0,
          level: Math.floor((item.TotalXP || 0) / 100) + 1,
          badges: this.parseJsonArray(item.BadgesEarned),
          lastVisited: item.LastActivity || new Date().toISOString(),
          lastSection: item.LastSection as OnboardingSection || undefined,
          preferences: {
            showBuddyBot: item.ShowBuddyBot !== false,
            emailNotifications: item.EmailNotifications !== false,
            dailyReminders: true,
            theme: 'system'
          }
        };
      }

      // Create new progress record
      await this.createProgressRecord(userId);
    } catch (error) {
      console.error('Failed to load onboarding progress:', error);
    }

    return this.getDefaultProgress(userId);
  }

  private async createProgressRecord(userId: string): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(LIST_NAMES.PROGRESS).items.add({
        Title: `Progress: ${userId}`,
        UserEmail: userId,
        StartDate: new Date().toISOString(),
        TotalXP: 0,
        Level: 1,
        CurrentStage: 1,
        StreakDays: 1,
        CompletedTasks: '[]',
        CompletedSections: '[]',
        SurvivalKitProgress: '[]',
        BingoProgress: '[]',
        BadgesEarned: '[]',
        ShowBuddyBot: true,
        EmailNotifications: true
      });
    } catch (error) {
      console.error('Failed to create progress record:', error);
    }
  }

  private getDefaultProgress(userId: string): IOnboardingProgress {
    return {
      userId,
      startDate: new Date().toISOString(),
      questProgress: this.getDefaultQuestProgress(),
      completedSections: [],
      survivalKitProgress: [],
      bingoBoard: { challenges: [], completedLines: 0, bingoAchieved: false, totalXpEarned: 0 },
      totalXp: 0,
      level: 1,
      badges: [],
      lastVisited: new Date().toISOString(),
      preferences: { showBuddyBot: true, emailNotifications: true, dailyReminders: true, theme: 'system' }
    };
  }

  public async updateOnboardingProgress(userId: string, progress: Partial<IOnboardingProgress>): Promise<void> {
    if (this.isWorkbench) {
      console.log('[Workbench] Updating progress:', progress);
      return;
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.PROGRESS)
        .items
        .filter(`UserEmail eq '${userId}'`)
        .top(1)();

      if (items.length > 0) {
        const updateData: Record<string, unknown> = {};
        if (progress.totalXp !== undefined) updateData.TotalXP = progress.totalXp;
        if (progress.level !== undefined) updateData.Level = progress.level;
        if (progress.completedSections) updateData.CompletedSections = JSON.stringify(progress.completedSections);
        if (progress.badges) updateData.BadgesEarned = JSON.stringify(progress.badges);
        if (progress.lastSection !== undefined) updateData.LastSection = progress.lastSection;
        updateData.LastActivity = new Date().toISOString();

        await this.sp.web.lists
          .getByTitle(LIST_NAMES.PROGRESS)
          .items.getById(items[0].Id)
          .update(updateData);
      }
    } catch (error) {
      console.error('Failed to update progress:', error);
    }
  }

  public async addXp(userId: string, xp: number, source: string): Promise<void> {
    console.log(`XP Earned: +${xp} from ${source}`);

    if (this.isWorkbench) return;

    try {
      const progress = await this.getOnboardingProgress(userId);
      const newXp = progress.totalXp + xp;
      const newLevel = Math.floor(newXp / 100) + 1;

      await this.updateOnboardingProgress(userId, {
        totalXp: newXp,
        level: newLevel
      });
    } catch (error) {
      console.error('Failed to add XP:', error);
    }
  }

  private async updateProgressField(userId: string, fieldName: string, value: string): Promise<void> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.PROGRESS)
        .items
        .filter(`UserEmail eq '${userId}'`)
        .top(1)();

      if (items.length > 0) {
        const updateData: Record<string, unknown> = {};
        updateData[fieldName] = value;
        updateData.LastActivity = new Date().toISOString();

        await this.sp.web.lists
          .getByTitle(LIST_NAMES.PROGRESS)
          .items.getById(items[0].Id)
          .update(updateData);
      }
    } catch (error) {
      console.error(`Failed to update ${fieldName}:`, error);
    }
  }

  // ============================================
  // BADGES
  // ============================================

  public async getAvailableBadges(): Promise<IOnboardingBadge[]> {
    if (this.isWorkbench) {
      return this.getMockBadges();
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.BADGES)
        .items
        .filter("IsActive eq 1")
        .orderBy("Rarity", true)
        .select("Id", "BadgeID", "BadgeName", "Description", "IconURL", "Rarity", "Category")();

      return items.map((item: Record<string, unknown>) => ({
        id: item.BadgeID as string || String(item.Id),
        name: item.BadgeName as string,
        description: item.Description as string || '',
        iconUrl: this.extractUrl(item.IconURL) || '',
        rarity: item.Rarity as 'common' | 'uncommon' | 'rare' | 'epic' | 'legendary' || 'common',
        category: item.Category as 'quest' | 'social' | 'explorer' | 'learner' | 'champion' || 'quest'
      }));
    } catch (error) {
      console.error('Failed to load badges from SharePoint:', error);
      return this.getMockBadges();
    }
  }

  public async getUserBadges(userId: string): Promise<IOnboardingBadge[]> {
    const progress = await this.getOnboardingProgress(userId);
    const allBadges = await this.getAvailableBadges();
    return allBadges.filter(b => progress.badges.includes(b.id));
  }

  public async awardBadge(userId: string, badgeId: string): Promise<void> {
    if (this.isWorkbench) {
      console.log(`[Workbench] Awarding badge ${badgeId} to user ${userId}`);
      return;
    }

    try {
      const progress = await this.getOnboardingProgress(userId);
      if (!progress.badges.includes(badgeId)) {
        progress.badges.push(badgeId);
        await this.updateProgressField(userId, 'BadgesEarned', JSON.stringify(progress.badges));
      }
    } catch (error) {
      console.error('Failed to award badge:', error);
    }
  }

  // ============================================
  // LEADERBOARD
  // ============================================

  public async getLeaderboard(period: string = 'all-time'): Promise<ILeaderboardEntry[]> {
    if (this.isWorkbench) {
      return this.getMockLeaderboard();
    }

    try {
      let filterQuery = '';
      const now = new Date();

      if (period === 'this-week') {
        const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
        filterQuery = `StartDate ge datetime'${weekAgo.toISOString()}'`;
      } else if (period === 'this-month') {
        const monthAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
        filterQuery = `StartDate ge datetime'${monthAgo.toISOString()}'`;
      }

      let query = this.sp.web.lists
        .getByTitle(LIST_NAMES.PROGRESS)
        .items
        .orderBy("TotalXP", false)
        .top(50)
        .select("Id", "UserEmail", "UserDisplayName", "Department", "StartDate", "TotalXP", "Level", "BadgesEarned", "StreakDays", "PhotoURL");

      if (filterQuery) {
        query = query.filter(filterQuery);
      }

      const items = await query();

      return items.map((item: Record<string, unknown>, index: number) => ({
        id: String(item.Id),
        rank: index + 1,
        userId: item.UserEmail as string,
        displayName: item.UserDisplayName as string || 'Anonymous User',
        photoUrl: this.extractUrl(item.PhotoURL) || '',
        department: item.Department as string || 'Unknown',
        startDate: item.StartDate as string || new Date().toISOString(),
        totalXp: item.TotalXP as number || 0,
        level: item.Level as number || 1,
        badgeCount: this.parseJsonArray(item.BadgesEarned as string).length,
        streak: item.StreakDays as number || 0,
        isCurrentUser: item.UserEmail === this.userEmail,
        trend: 'same' as const,
        previousRank: undefined
      }));
    } catch (error) {
      console.error('Failed to load leaderboard from SharePoint:', error);
      return this.getMockLeaderboard();
    }
  }

  public async getLeaderboardStats(): Promise<ILeaderboardStats> {
    if (this.isWorkbench) {
      return {
        totalParticipants: 42,
        averageXp: 285,
        topDepartment: 'Engineering',
        weeklyGrowth: 12
      };
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_NAMES.PROGRESS)
        .items
        .select("TotalXP", "Department")();

      const totalParticipants = items.length;
      const totalXp = items.reduce((sum: number, item: Record<string, unknown>) => sum + ((item.TotalXP as number) || 0), 0);
      const averageXp = totalParticipants > 0 ? totalXp / totalParticipants : 0;

      // Calculate top department
      const deptCounts: Record<string, number> = {};
      items.forEach((item: Record<string, unknown>) => {
        const dept = (item.Department as string) || 'Unknown';
        deptCounts[dept] = (deptCounts[dept] || 0) + ((item.TotalXP as number) || 0);
      });
      const topDepartment = Object.entries(deptCounts)
        .sort(([, a], [, b]) => b - a)[0]?.[0] || 'Engineering';

      return {
        totalParticipants,
        averageXp,
        topDepartment,
        weeklyGrowth: 12 // Would need historical data to calculate
      };
    } catch (error) {
      console.error('Failed to load leaderboard stats:', error);
      return {
        totalParticipants: 0,
        averageXp: 0,
        topDepartment: 'Unknown',
        weeklyGrowth: 0
      };
    }
  }

  private getMockLeaderboard(): ILeaderboardEntry[] {
    const mockUsers = [
      { name: 'Alex Thompson', dept: 'Engineering', xp: 850, days: 14 },
      { name: 'Maya Patel', dept: 'Product', xp: 720, days: 10 },
      { name: 'Chris Johnson', dept: 'Engineering', xp: 680, days: 12 },
      { name: 'Sarah Williams', dept: 'Marketing', xp: 550, days: 8 },
      { name: 'David Kim', dept: 'Design', xp: 480, days: 7 },
      { name: 'Emma Davis', dept: 'Sales', xp: 420, days: 6 },
      { name: 'Michael Brown', dept: 'Engineering', xp: 380, days: 5 },
      { name: 'Lisa Anderson', dept: 'HR', xp: 320, days: 4 },
      { name: this.userDisplayName || 'You', dept: 'Engineering', xp: 285, days: 3 },
      { name: 'James Wilson', dept: 'Finance', xp: 240, days: 3 }
    ];

    return mockUsers.map((user, index) => ({
      id: `lb-${index + 1}`,
      rank: index + 1,
      userId: user.name === (this.userDisplayName || 'You') ? this.userEmail : `user${index}@company.com`,
      displayName: user.name,
      photoUrl: '',
      department: user.dept,
      startDate: new Date(Date.now() - user.days * 24 * 60 * 60 * 1000).toISOString(),
      totalXp: user.xp,
      level: Math.floor(user.xp / 100) + 1,
      badgeCount: Math.floor(user.xp / 150),
      streak: Math.min(user.days, 7),
      isCurrentUser: user.name === (this.userDisplayName || 'You'),
      trend: index < 3 ? 'up' : index > 7 ? 'down' : 'same',
      previousRank: index < 3 ? index + 2 : index > 7 ? index - 1 : index + 1
    }));
  }

  // ============================================
  // HELPER METHODS
  // ============================================

  private parseJsonArray(value: string | undefined | null): string[] {
    if (!value) return [];
    try {
      const parsed = JSON.parse(value);
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }

  private parseCommaSeparated(value: string | undefined | null): string[] {
    if (!value) return [];
    return value.split(',').map(s => s.trim()).filter(s => s.length > 0);
  }

  private parseExpertise(value: string | undefined | null): string[] {
    if (!value) return [];
    // Handle both JSON array and comma-separated formats
    try {
      const parsed = JSON.parse(value);
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return this.parseCommaSeparated(value);
    }
  }

  private extractUrl(value: unknown): string | undefined {
    if (!value) return undefined;
    if (typeof value === 'string') return value;
    if (typeof value === 'object' && value !== null) {
      const urlObj = value as { Url?: string; Description?: string };
      return urlObj.Url;
    }
    return undefined;
  }

  // ============================================
  // MOCK DATA METHODS (for workbench testing)
  // ============================================

  private getMockQuestStages(): IQuestStage[] {
    return [
      {
        id: 'stage-1', title: 'Day One Hero', description: 'Complete your first day essentials', icon: 'Sunrise',
        xpReward: 100, badgeId: 'day-one-hero', unlocked: true, completed: false,
        tasks: [
          { id: 'task-1-1', title: 'Log into your computer', description: 'Successfully access your workstation', type: 'action', completed: false, xpReward: 20 },
          { id: 'task-1-2', title: 'Meet your manager', description: 'Have a welcome chat with your direct manager', type: 'social', completed: false, xpReward: 25 },
          { id: 'task-1-3', title: 'Get your ID badge', description: 'Collect your employee ID from reception', type: 'action', completed: false, xpReward: 15 },
          { id: 'task-1-4', title: 'Set up your email', description: 'Configure Outlook and send a test email', type: 'action', completed: false, xpReward: 20 }
        ]
      },
      {
        id: 'stage-2', title: 'Team Explorer', description: 'Get to know your team', icon: 'People',
        xpReward: 150, badgeId: 'team-explorer', unlocked: true, completed: false,
        tasks: [
          { id: 'task-2-1', title: 'Meet 3 team members', description: 'Introduce yourself to at least 3 colleagues', type: 'social', completed: false, xpReward: 30 },
          { id: 'task-2-2', title: 'Find your buddy', description: 'Connect with your assigned onboarding buddy', type: 'social', completed: false, xpReward: 25 }
        ]
      }
    ];
  }

  private getMockTeamMembers(): ITeamMember[] {
    return [
      { id: '1', name: 'Sarah Mitchell', title: 'Department Head', department: 'Engineering', email: 'sarah.mitchell@company.com', photoUrl: '', funFact: 'Has climbed 5 of the world\'s highest peaks!', expertise: ['Leadership', 'Strategy'], isKeyContact: true, contactType: 'manager', location: 'Floor 3, Desk 301' },
      { id: '2', name: 'James Chen', title: 'Your Onboarding Buddy', department: 'Engineering', email: 'james.chen@company.com', photoUrl: '', funFact: 'Makes the best coffee in the office!', expertise: ['React', 'TypeScript'], isKeyContact: true, contactType: 'buddy', location: 'Floor 3, Desk 315' },
      { id: '3', name: 'Emily Watson', title: 'HR Business Partner', department: 'Human Resources', email: 'emily.watson@company.com', photoUrl: '', funFact: 'Completed a marathon on every continent', expertise: ['Benefits', 'Policies'], isKeyContact: true, contactType: 'hr', location: 'Floor 1, HR Hub' },
      { id: '4', name: 'Michael Torres', title: 'IT Support Lead', department: 'IT', email: 'michael.torres@company.com', photoUrl: '', funFact: 'Built his first computer at age 12', expertise: ['Hardware', 'Software'], isKeyContact: true, contactType: 'it', location: 'Floor 2, IT Corner' }
    ];
  }

  private getMockFloorPlans(): IFloorPlan[] {
    return [
      { floor: 1, name: 'Ground Floor - Reception & Services', imageUrl: '', locations: [
        { id: 'reception', name: 'Reception', type: 'reception', floor: 1, coordinates: { x: 50, y: 20 }, description: 'Main entrance and visitor check-in', tips: 'Collect your ID badge here' },
        { id: 'cafeteria', name: 'Cafeteria', type: 'cafeteria', floor: 1, coordinates: { x: 30, y: 50 }, description: 'Main dining area', tips: 'Lunch 11:30am-2pm' }
      ]},
      { floor: 2, name: 'Floor 2 - IT & Meeting Rooms', imageUrl: '', locations: [
        { id: 'it-corner', name: 'IT Support', type: 'desk-area', floor: 2, coordinates: { x: 20, y: 30 }, description: 'IT help desk', tips: 'Walk-up support 9am-5pm' },
        { id: 'kitchen-2', name: 'Floor 2 Kitchen', type: 'kitchen', floor: 2, coordinates: { x: 80, y: 50 }, description: 'Tea, coffee, and snacks', tips: 'Free fruit on Mondays!' }
      ]},
      { floor: 3, name: 'Floor 3 - Engineering & Product', imageUrl: '', locations: [
        { id: 'eng-area', name: 'Engineering Zone', type: 'desk-area', floor: 3, coordinates: { x: 40, y: 40 }, description: 'Your new home!', tips: 'Look for the welcome balloon' },
        { id: 'kitchen-3', name: 'Floor 3 Kitchen', type: 'kitchen', floor: 3, coordinates: { x: 20, y: 70 }, description: 'The legendary coffee machine', tips: 'Ask James for training!' }
      ]}
    ];
  }

  private getMockSurvivalItems(): ISurvivalItem[] {
    return [
      { id: 'sk-1', title: 'Collect your ID badge', description: 'Visit reception to get your employee ID card', icon: 'ContactCard', category: 'day-one', completed: false, priority: 'high', helpText: 'Required for building access' },
      { id: 'sk-2', title: 'Log into your computer', description: 'Use the temporary password from IT', icon: 'Lock', category: 'day-one', completed: false, priority: 'high' },
      { id: 'sk-3', title: 'Set up email and calendar', description: 'Configure Outlook and sync your calendar', icon: 'Mail', category: 'day-one', completed: false, priority: 'high' },
      { id: 'sk-4', title: 'Complete IT security training', description: 'Mandatory security awareness module', icon: 'Shield', category: 'first-week', completed: false, priority: 'high' },
      { id: 'sk-5', title: 'Read employee handbook', description: 'Review key policies and procedures', icon: 'ReadingMode', category: 'first-week', completed: false, priority: 'medium' }
    ];
  }

  private getMockCoffeeLessons(): ICoffeeLesson[] {
    return [
      { id: 'coffee-1', title: 'Meet Your Coffee Machine', description: 'Introduction to our DeLonghi Magnifica', type: 'interactive', duration: '3 min', completed: false, xpReward: 15, content: 'Learn the basics of our professional coffee machine.' },
      { id: 'coffee-2', title: 'The Perfect Espresso', description: 'Master the art of the espresso shot', type: 'interactive', duration: '2 min', completed: false, xpReward: 15, content: 'An espresso is the foundation of all coffee drinks.' },
      { id: 'coffee-3', title: 'Coffee Quiz Challenge', description: 'Test your coffee knowledge!', type: 'quiz', duration: '2 min', completed: false, xpReward: 30, quizQuestions: [
        { id: 'q1', question: 'What is the ideal temperature for steaming milk?', options: ['50-55°C', '60-65°C', '70-75°C', '80-85°C'], correctAnswer: 1, explanation: '60-65°C is the sweet spot.' }
      ]}
    ];
  }

  private getMockWelcomeMessages(): IWelcomeMessage[] {
    return [
      { id: 'wm-1', authorName: 'Sarah Mitchell', authorTitle: 'Department Head', authorPhotoUrl: '', message: 'Welcome to the team! We\'re thrilled to have you on board.', timestamp: new Date(Date.now() - 2 * 60 * 60 * 1000).toISOString(), likes: 12, hasLiked: false },
      { id: 'wm-2', authorName: 'James Chen', authorTitle: 'Your Onboarding Buddy', authorPhotoUrl: '', message: 'Hey! Let\'s grab coffee tomorrow - I\'ll show you around!', timestamp: new Date(Date.now() - 60 * 60 * 1000).toISOString(), likes: 8, hasLiked: false }
    ];
  }

  private getMockJargonTerms(): IJargonTerm[] {
    return [
      { id: 'j-1', term: 'JML', definition: 'Joiner, Mover, Leaver - The employee lifecycle stages.', category: 'HR', example: 'We need to process a JML request.', relatedTerms: ['Onboarding', 'Offboarding'] },
      { id: 'j-2', term: 'EOD', definition: 'End of Day - Close of business, usually 5pm.', category: 'General', example: 'Please send the report by EOD.' },
      { id: 'j-3', term: 'OOO', definition: 'Out of Office - When someone is away from work.', category: 'General', example: 'I\'ll be OOO next week for vacation.' }
    ];
  }

  private getMockInterestGroups(): IInterestGroup[] {
    return [
      { id: 'ig-1', name: 'Running Club', description: 'Weekly group runs for all levels.', category: 'sports', iconUrl: 'Running', memberCount: 45, nextMeeting: 'Every Tuesday 7am', contactPerson: 'Tom Anderson', isJoined: false, tags: ['fitness', 'outdoor'] },
      { id: 'ig-2', name: 'Board Games Lunch', description: 'Casual board games during lunch break.', category: 'hobbies', iconUrl: 'Game', memberCount: 32, nextMeeting: 'Every Thursday 12:30pm', contactPerson: 'David Kim', isJoined: false, tags: ['games', 'social'] },
      { id: 'ig-3', name: 'Tech Talks', description: 'Monthly presentations on tech trends.', category: 'professional', iconUrl: 'Code', memberCount: 78, nextMeeting: 'First Wednesday of month', contactPerson: 'Lisa Park', isJoined: false, tags: ['learning', 'tech'] }
    ];
  }

  private getMockBingoChallenges(): IBingoChallenge[] {
    return [
      { id: 'b-1', title: 'Say hello to 5 people', description: 'Introduce yourself to five new colleagues', icon: 'Chat', completed: false, xpReward: 20, verificationMethod: 'self', category: 'social' },
      { id: 'b-2', title: 'Find the best coffee', description: 'Try all coffee machines', icon: 'CoffeeScript', completed: false, xpReward: 15, verificationMethod: 'self', category: 'explore' },
      { id: 'b-13', title: 'FREE SPACE', description: 'You\'ve already earned this one!', icon: 'CheckMark', completed: true, xpReward: 0, verificationMethod: 'auto', category: 'fun' }
    ];
  }

  private getMockBadges(): IOnboardingBadge[] {
    return [
      { id: 'day-one-hero', name: 'Day One Hero', description: 'Completed all first day tasks', iconUrl: '', rarity: 'common', category: 'quest' },
      { id: 'team-explorer', name: 'Team Explorer', description: 'Met all your key contacts', iconUrl: '', rarity: 'common', category: 'social' },
      { id: 'coffee-connoisseur', name: 'Coffee Connoisseur', description: 'Completed Coffee Academy', iconUrl: '', rarity: 'common', category: 'learner' }
    ];
  }
}

export default OnboardingExperienceService;
