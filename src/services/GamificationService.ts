// @ts-nocheck
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";
import { WebPartContext } from "@microsoft/sp-webpart-base";

// ============================================================================
// GAMIFICATION HUB INTERFACES (New)
// ============================================================================

export interface IGamificationProfile {
  userId: number;
  userEmail: string;
  displayName: string;
  totalPoints: number;
  availablePoints: number;
  lifetimePoints: number;
  pointsThisMonth: number;
  currentLevel: number;
  currentTier: string;
  tierMultiplier: number;
  achievementCount: number;
  badgeCount: number;
  currentStreak: number;
  longestStreak: number;
  streakMultiplier: number;
  pointsToNextLevel: number;
  pointsToNextTier: number;
  leaderboardRank: number;
  department: string;
}

export interface IAchievement {
  id: number;
  code: string;
  name: string;
  description: string;
  category: string;
  rarity: string;
  points: number;
  icon: string;
  isUnlocked: boolean;
  unlockedDate?: string;
  progress?: number;
  progressTarget?: number;
}

export interface ILeaderboardEntry {
  rank: number;
  userId: number;
  userEmail: string;
  displayName: string;
  department: string;
  points: number;
  level: number;
  tier: string;
  achievementCount: number;
  isCurrentUser: boolean;
  rankChange?: number;
}

export interface IChallenge {
  id: number;
  code: string;
  name: string;
  description: string;
  type: string;
  category: string;
  status: string;
  startDate: string;
  endDate: string;
  daysLeft: number;
  goalTarget: number;
  currentProgress: number;
  progressPercent: number;
  rewardPoints: number;
  participants: number;
  isJoined: boolean;
}

export interface IReward {
  id: number;
  code: string;
  name: string;
  description: string;
  category: string;
  pointsCost: number;
  originalCost: number;
  tierDiscount: number;
  icon: string;
  imageUrl?: string;
  stockLevel: number;
  isAvailable: boolean;
  isFeatured: boolean;
}

export interface IActivityFeedItem {
  id: number;
  userId: number;
  userEmail: string;
  displayName: string;
  type: string;
  message: string;
  icon: string;
  timestamp: string;
  relatedItemType?: string;
  relatedItemId?: number;
  likes: number;
}

export interface IStreak {
  id: number;
  type: string;
  currentDays: number;
  longestDays: number;
  lastActivityDate: string;
  multiplier: number;
  nextMilestone: number;
  isActive: boolean;
}

export interface IRecognition {
  id: number;
  giverId: number;
  giverName: string;
  giverEmail: string;
  receiverId: number;
  receiverName: string;
  receiverEmail: string;
  type: string;
  message: string;
  timestamp: string;
  likes: number;
  isPublic: boolean;
}

// ============================================================================
// LEGACY POLICY HUB INTERFACES (For PolicyHub module)
// ============================================================================

export interface ILegacyAchievement {
  Id: number;
  Title: string;
  AchievementName: string;
  AchievementDescription: string;
  AchievementType: string;
  IconName: string;
  PointsReward: number;
  RequirementType: string;
  RequirementValue: number;
  BadgeLevel: string;
  IsActive: boolean;
  DisplayOrder: number;
}

export interface ILegacyUserAchievement {
  Id: number;
  UserId: any;
  AchievementId: number;
  AchievementName: string;
  EarnedDate: string;
  PointsEarned: number;
  BadgeLevel: string;
  IsDisplayed: boolean;
}

export interface ILegacyLeaderboardEntry {
  userId: number;
  userName: string;
  userEmail: string;
  totalPoints: number;
  level: number;
  levelName: string;
  rank: number;
  policiesRead: number;
  quizzesCompleted: number;
  streakDays: number;
}

export interface IUserPoints {
  Id: number;
  UserId: any;
  TotalPoints: number;
  QuizPoints: number;
  ReadingPoints: number;
  AcknowledgementPoints: number;
  BonusPoints: number;
  CurrentLevel: number;
  LevelName: string;
  PoliciesRead: number;
  QuizzesCompleted: number;
  QuizzesPassed: number;
  StreakDays: number;
  LastActivityDate: string;
  Rank?: number;
}

export interface IPointTransaction {
  Id: number;
  UserId: any;
  TransactionType: string;
  Points: number;
  TransactionDate: string;
  ReferenceId?: number;
  ReferenceType?: string;
  Description: string;
  NewBalance: number;
}

export interface IUserProgress {
  userPoints: IUserPoints;
  achievements: ILegacyUserAchievement[];
  recentTransactions: IPointTransaction[];
  nextLevel: {
    level: number;
    levelName: string;
    pointsRequired: number;
    pointsRemaining: number;
    progressPercentage: number;
  };
  stats: {
    totalActivities: number;
    thisWeekPoints: number;
    thisMonthPoints: number;
    averageQuizScore: number;
  };
}

export class GamificationService {
  private sp: SPFI;
  private context: WebPartContext | null = null;
  private currentUserEmail: string = '';
  private currentUserId: number = 0;

  // List names
  private readonly REWARDS_LIST = 'JML_GamificationRewards';
  private readonly REDEMPTIONS_LIST = 'JML_GamificationRedemptions';
  private readonly RECOGNITIONS_LIST = 'JML_GamificationRecognitions';

  // Level thresholds (expanded for Gamification Hub)
  private readonly LEVEL_THRESHOLDS = [
    { level: 1, name: "Newcomer", points: 0 },
    { level: 2, name: "Apprentice", points: 100 },
    { level: 3, name: "Explorer", points: 300 },
    { level: 4, name: "Achiever", points: 750 },
    { level: 5, name: "Expert", points: 1500 },
    { level: 6, name: "Master", points: 3000 },
    { level: 7, name: "Champion", points: 6000 },
    { level: 8, name: "Legend", points: 10000 },
    { level: 9, name: "Elite", points: 20000 },
    { level: 10, name: "Grandmaster", points: 50000 }
  ];

  // Tier thresholds
  private readonly TIER_THRESHOLDS = [
    { tier: "Bronze", points: 0, multiplier: 1.0, discount: 0 },
    { tier: "Silver", points: 2500, multiplier: 1.25, discount: 5 },
    { tier: "Gold", points: 10000, multiplier: 1.5, discount: 10 },
    { tier: "Platinum", points: 25000, multiplier: 2.0, discount: 15 }
  ];

  // Point rewards
  private readonly POINTS = {
    POLICY_READ: 10,
    POLICY_ACKNOWLEDGEMENT: 15,
    QUIZ_PASSED: 50,
    QUIZ_PERFECT: 100,
    DAILY_STREAK: 20,
    WEEKLY_STREAK: 50,
    MONTHLY_STREAK: 200,
    TASK_COMPLETED: 25,
    RECOGNITION_GIVEN: 15,
    RECOGNITION_RECEIVED: 30,
    CHALLENGE_JOINED: 10,
    CHALLENGE_COMPLETED: 100
  };

  // Streak multipliers
  private readonly STREAK_MULTIPLIERS = [
    { days: 7, multiplier: 1.1 },
    { days: 14, multiplier: 1.25 },
    { days: 30, multiplier: 1.5 },
    { days: 60, multiplier: 1.75 },
    { days: 90, multiplier: 2.0 }
  ];

  /**
   * Constructor supporting both SPFI (legacy) and WebPartContext (new)
   */
  constructor(spOrContext: SPFI | WebPartContext) {
    if ('web' in spOrContext) {
      // Legacy: SPFI passed directly
      this.sp = spOrContext as SPFI;
    } else {
      // New: WebPartContext passed
      this.context = spOrContext as WebPartContext;
      this.sp = spfi().using(SPFx(this.context));
      this.currentUserEmail = this.context.pageContext.user.email;
      // Initialize current user ID asynchronously
      this.initCurrentUserId().catch(err => console.error('[GamificationService] Failed to init user ID:', err));
    }
  }

  /**
   * Initialize current user ID
   */
  private async initCurrentUserId(): Promise<void> {
    try {
      const user = await this.sp.web.currentUser();
      this.currentUserId = user.Id;
    } catch (error) {
      console.warn('[GamificationService] Could not get current user ID');
    }
  }

  // ============================================================================
  // GAMIFICATION HUB METHODS (New)
  // ============================================================================

  /**
   * Get current user's gamification profile
   */
  public async getCurrentUserProfile(): Promise<IGamificationProfile> {
    try {
      const userEmail = this.currentUserEmail || await this.getCurrentUserEmail();

      // Try to get existing profile from JML_GamificationProfiles
      const profiles = await this.sp.web.lists
        .getByTitle("JML_GamificationProfiles")
        .items.filter(`UserEmail eq '${userEmail}'`)
        .top(1)();

      if (profiles.length > 0) {
        const profile = profiles[0];
        const tier = this.calculateTier(profile.TotalPoints || 0);
        const level = this.calculateLevel(profile.TotalPoints || 0);
        const streak = this.calculateStreakMultiplier(profile.CurrentStreak || 0);

        return {
          userId: profile.Id,
          userEmail: profile.UserEmail || userEmail,
          displayName: profile.DisplayName || 'User',
          totalPoints: profile.TotalPoints || 0,
          availablePoints: profile.AvailablePoints || 0,
          lifetimePoints: profile.LifetimePoints || 0,
          pointsThisMonth: profile.PointsThisMonth || 0,
          currentLevel: level.level,
          currentTier: tier.tier,
          tierMultiplier: tier.multiplier,
          achievementCount: profile.AchievementCount || 0,
          badgeCount: profile.BadgeCount || 0,
          currentStreak: profile.CurrentStreak || 0,
          longestStreak: profile.LongestStreak || 0,
          streakMultiplier: streak,
          pointsToNextLevel: this.getPointsToNextLevel(profile.TotalPoints || 0),
          pointsToNextTier: this.getPointsToNextTier(profile.TotalPoints || 0),
          leaderboardRank: profile.LeaderboardRank || 0,
          department: profile.Department || ''
        };
      }

      // Return default profile for new users
      return this.getDefaultProfile(userEmail);
    } catch (error) {
      console.error("Failed to get current user profile:", error);
      // Return mock data for workbench/development
      return this.getMockProfile();
    }
  }

  /**
   * Get user achievements (both unlocked and locked)
   */
  public async getUserAchievements(): Promise<IAchievement[]> {
    try {
      const userEmail = this.currentUserEmail || await this.getCurrentUserEmail();

      // Get all achievements
      const allAchievements = await this.sp.web.lists
        .getByTitle("JML_Achievements")
        .items.filter("IsActive eq true")
        .orderBy("DisplayOrder", true)();

      // Get user's unlocked achievements
      const userAchievements = await this.sp.web.lists
        .getByTitle("JML_UserAchievements")
        .items.filter(`UserEmail eq '${userEmail}'`)();

      const unlockedCodes = userAchievements.map((ua: any) => ua.AchievementCode);

      return allAchievements.map((a: any) => ({
        id: a.Id,
        code: a.AchievementCode || '',
        name: a.AchievementName || a.Title || '',
        description: a.AchievementDescription || '',
        category: a.AchievementCategory || 'General',
        rarity: a.Rarity || 'Common',
        points: a.PointsReward || 0,
        icon: a.IconName || 'Trophy',
        isUnlocked: unlockedCodes.includes(a.AchievementCode),
        unlockedDate: userAchievements.find((ua: any) => ua.AchievementCode === a.AchievementCode)?.UnlockedDate,
        progress: a.Progress || 0,
        progressTarget: a.ProgressTarget || 100
      }));
    } catch (error) {
      console.error("Failed to get user achievements:", error);
      return this.getMockAchievements();
    }
  }

  /**
   * Get leaderboard entries
   */
  public async getLeaderboard(top: number = 10): Promise<ILeaderboardEntry[]> {
    try {
      const userEmail = this.currentUserEmail || await this.getCurrentUserEmail();

      const leaderboardItems = await this.sp.web.lists
        .getByTitle("JML_Leaderboard")
        .items.filter("IsCurrent eq true and LeaderboardType eq 'Global'")
        .orderBy("LeaderboardRank", true)
        .top(top)();

      return leaderboardItems.map((item: any) => ({
        rank: item.LeaderboardRank || 0,
        userId: item.Id,
        userEmail: item.UserEmail || '',
        displayName: item.DisplayName || 'User',
        department: item.Department || '',
        points: item.Points || 0,
        level: item.UserLevel || 1,
        tier: item.UserTier || 'Bronze',
        achievementCount: item.AchievementCount || 0,
        isCurrentUser: item.UserEmail === userEmail,
        rankChange: item.RankChange || 0
      }));
    } catch (error) {
      console.error("Failed to get leaderboard:", error);
      return this.getMockLeaderboard(top);
    }
  }

  /**
   * Get active challenges
   */
  public async getActiveChallenges(): Promise<IChallenge[]> {
    try {
      const userEmail = this.currentUserEmail || await this.getCurrentUserEmail();
      const today = new Date().toISOString();

      const challenges = await this.sp.web.lists
        .getByTitle("JML_Challenges")
        .items.filter(`ChallengeStatus eq 'Active' and EndDate ge datetime'${today}'`)
        .orderBy("EndDate", true)();

      // Get user's challenge participation
      const participations = await this.sp.web.lists
        .getByTitle("JML_ChallengeParticipants")
        .items.filter(`UserEmail eq '${userEmail}'`)();

      const joinedCodes = participations.map((p: any) => p.ChallengeCode);

      return challenges.map((c: any) => {
        const endDate = new Date(c.EndDate);
        const now = new Date();
        const daysLeft = Math.ceil((endDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
        const participation = participations.find((p: any) => p.ChallengeCode === c.ChallengeCode);

        return {
          id: c.Id,
          code: c.ChallengeCode || '',
          name: c.ChallengeName || c.Title || '',
          description: c.ChallengeDescription || '',
          type: c.ChallengeType || 'Individual',
          category: c.ChallengeCategory || 'General',
          status: c.ChallengeStatus || 'Active',
          startDate: c.StartDate || '',
          endDate: c.EndDate || '',
          daysLeft: Math.max(0, daysLeft),
          goalTarget: c.GoalTarget || 100,
          currentProgress: participation?.CurrentProgress || 0,
          progressPercent: participation ? Math.min(100, (participation.CurrentProgress / c.GoalTarget) * 100) : 0,
          rewardPoints: c.PointsForCompletion || 0,
          participants: c.TotalParticipants || 0,
          isJoined: joinedCodes.includes(c.ChallengeCode)
        };
      });
    } catch (error) {
      console.error("Failed to get active challenges:", error);
      return this.getMockChallenges();
    }
  }

  /**
   * Get available rewards
   */
  public async getAvailableRewards(): Promise<IReward[]> {
    try {
      const profile = await this.getCurrentUserProfile();
      const tier = this.calculateTier(profile.totalPoints);

      const rewards = await this.sp.web.lists
        .getByTitle("JML_RewardsCatalog")
        .items.filter("IsAvailable eq true and StockLevel gt 0")
        .orderBy("IsFeatured", false)
        .orderBy("PointsCost", true)();

      return rewards.map((r: any) => {
        const discount = tier.discount;
        const originalCost = r.PointsCost || 0;
        const discountedCost = Math.round(originalCost * (1 - discount / 100));

        return {
          id: r.Id,
          code: r.RewardCode || '',
          name: r.RewardName || r.Title || '',
          description: r.RewardDescription || '',
          category: r.RewardCategory || 'General',
          pointsCost: discountedCost,
          originalCost: originalCost,
          tierDiscount: discount,
          icon: r.RewardIcon || 'Gift',
          imageUrl: r.RewardImageUrl?.Url || '',
          stockLevel: r.StockLevel || 0,
          isAvailable: r.IsAvailable !== false,
          isFeatured: r.IsFeatured === true
        };
      });
    } catch (error) {
      console.error("Failed to get available rewards:", error);
      return this.getMockRewards();
    }
  }

  /**
   * Get activity feed
   */
  public async getActivityFeed(top: number = 10): Promise<IActivityFeedItem[]> {
    try {
      const activities = await this.sp.web.lists
        .getByTitle("JML_ActivityFeed")
        .items.filter("IsActive eq true and Visibility eq 'Public'")
        .orderBy("ActivityDate", false)
        .top(top)();

      return activities.map((a: any) => ({
        id: a.Id,
        userId: a.FeedUserId || 0,
        userEmail: a.UserEmail || '',
        displayName: a.DisplayName || 'User',
        type: this.mapActivityType(a.ActivityType),
        message: this.formatActivityMessage(a),
        icon: this.getActivityIcon(a.ActivityType),
        timestamp: a.ActivityDate || new Date().toISOString(),
        relatedItemType: a.RelatedItemType,
        relatedItemId: a.RelatedItemId,
        likes: a.LikesCount || 0
      }));
    } catch (error) {
      console.error("Failed to get activity feed:", error);
      return this.getMockActivityFeed();
    }
  }

  /**
   * Get user streaks
   */
  public async getUserStreaks(): Promise<IStreak[]> {
    try {
      const userEmail = this.currentUserEmail || await this.getCurrentUserEmail();

      const streaks = await this.sp.web.lists
        .getByTitle("JML_Streaks")
        .items.filter(`UserEmail eq '${userEmail}' and IsActive eq true`)();

      return streaks.map((s: any) => ({
        id: s.Id,
        type: s.StreakType || 'Daily Login',
        currentDays: s.CurrentStreakDays || 0,
        longestDays: s.LongestStreakDays || 0,
        lastActivityDate: s.LastActivityDate || '',
        multiplier: this.calculateStreakMultiplier(s.CurrentStreakDays || 0),
        nextMilestone: s.NextMilestone || 7,
        isActive: s.IsActive !== false
      }));
    } catch (error) {
      console.error("Failed to get user streaks:", error);
      return this.getMockStreaks();
    }
  }

  /**
   * Get recent recognitions
   */
  public async getRecentRecognitions(top: number = 5): Promise<IRecognition[]> {
    try {
      const recognitions = await this.sp.web.lists
        .getByTitle("JML_Recognition")
        .items.filter("IsPublic eq true and RecognitionStatus eq 'Active'")
        .orderBy("GivenDate", false)
        .top(top)();

      return recognitions.map((r: any) => ({
        id: r.Id,
        giverId: r.GivenById || 0,
        giverName: r.GivenBy?.Title || 'Someone',
        giverEmail: r.GiverEmail || '',
        receiverId: r.GivenToId || 0,
        receiverName: r.GivenTo?.Title || 'Someone',
        receiverEmail: r.ReceiverEmail || '',
        type: r.RecognitionType || 'Kudos',
        message: r.RecognitionMessage || '',
        timestamp: r.GivenDate || new Date().toISOString(),
        likes: r.LikesCount || 0,
        isPublic: r.IsPublic !== false
      }));
    } catch (error) {
      console.error("Failed to get recent recognitions:", error);
      return this.getMockRecognitions();
    }
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  private async getCurrentUserEmail(): Promise<string> {
    if (this.currentUserEmail) return this.currentUserEmail;
    try {
      const currentUser = await this.sp.web.currentUser();
      this.currentUserEmail = currentUser.Email;
      return this.currentUserEmail;
    } catch {
      return 'user@example.com';
    }
  }

  private calculateTier(points: number): { tier: string; multiplier: number; discount: number } {
    for (let i = this.TIER_THRESHOLDS.length - 1; i >= 0; i--) {
      if (points >= this.TIER_THRESHOLDS[i].points) {
        return this.TIER_THRESHOLDS[i];
      }
    }
    return this.TIER_THRESHOLDS[0];
  }

  private getPointsToNextLevel(points: number): number {
    const currentLevel = this.calculateLevel(points);
    const nextThreshold = this.LEVEL_THRESHOLDS.find(t => t.level === currentLevel.level + 1);
    return nextThreshold ? nextThreshold.points - points : 0;
  }

  private getPointsToNextTier(points: number): number {
    const currentTier = this.calculateTier(points);
    const nextTierIndex = this.TIER_THRESHOLDS.findIndex(t => t.tier === currentTier.tier) + 1;
    if (nextTierIndex < this.TIER_THRESHOLDS.length) {
      return this.TIER_THRESHOLDS[nextTierIndex].points - points;
    }
    return 0;
  }

  private calculateStreakMultiplier(days: number): number {
    for (let i = this.STREAK_MULTIPLIERS.length - 1; i >= 0; i--) {
      if (days >= this.STREAK_MULTIPLIERS[i].days) {
        return this.STREAK_MULTIPLIERS[i].multiplier;
      }
    }
    return 1.0;
  }

  private mapActivityType(type: string): string {
    const typeMap: { [key: string]: string } = {
      'Achievement Unlocked': 'achievement',
      'Level Up': 'levelUp',
      'Recognition Given': 'recognition',
      'Recognition Received': 'recognition',
      'Challenge Completed': 'challenge',
      'Challenge Won': 'challenge',
      'Reward Redeemed': 'reward',
      'Streak Milestone': 'streak'
    };
    return typeMap[type] || 'general';
  }

  private getActivityIcon(type: string): string {
    const iconMap: { [key: string]: string } = {
      'Achievement Unlocked': 'Trophy',
      'Level Up': 'SkypeCircleArrow',
      'Recognition Given': 'Heart',
      'Recognition Received': 'Like',
      'Challenge Completed': 'FlameSolid',
      'Challenge Won': 'Medal',
      'Reward Redeemed': 'GiftboxOpen',
      'Streak Milestone': 'Calories'
    };
    return iconMap[type] || 'Info';
  }

  private formatActivityMessage(activity: any): string {
    const name = `<strong>${activity.DisplayName || 'Someone'}</strong>`;
    switch (activity.ActivityType) {
      case 'Achievement Unlocked':
        return `${name} unlocked the "${activity.RelatedItemTitle || 'achievement'}" achievement`;
      case 'Level Up':
        return `${name} reached Level ${activity.RelatedItemTitle || 'new'}!`;
      case 'Recognition Received':
        return `${name} received recognition from ${activity.SecondUserName || 'a colleague'}`;
      case 'Challenge Won':
        return `${name} won the "${activity.RelatedItemTitle || 'challenge'}" challenge!`;
      default:
        return activity.Headline || `${name} did something awesome`;
    }
  }

  private getDefaultProfile(email: string): IGamificationProfile {
    return {
      userId: 0,
      userEmail: email,
      displayName: 'New User',
      totalPoints: 0,
      availablePoints: 0,
      lifetimePoints: 0,
      pointsThisMonth: 0,
      currentLevel: 1,
      currentTier: 'Bronze',
      tierMultiplier: 1.0,
      achievementCount: 0,
      badgeCount: 0,
      currentStreak: 0,
      longestStreak: 0,
      streakMultiplier: 1.0,
      pointsToNextLevel: 100,
      pointsToNextTier: 2500,
      leaderboardRank: 0,
      department: ''
    };
  }

  // ============================================================================
  // MOCK DATA FOR WORKBENCH TESTING
  // ============================================================================

  private getMockProfile(): IGamificationProfile {
    return {
      userId: 1,
      userEmail: 'user@example.com',
      displayName: 'Demo User',
      totalPoints: 3750,
      availablePoints: 2500,
      lifetimePoints: 5000,
      pointsThisMonth: 450,
      currentLevel: 6,
      currentTier: 'Silver',
      tierMultiplier: 1.25,
      achievementCount: 12,
      badgeCount: 8,
      currentStreak: 14,
      longestStreak: 21,
      streakMultiplier: 1.25,
      pointsToNextLevel: 2250,
      pointsToNextTier: 6250,
      leaderboardRank: 5,
      department: 'Human Resources'
    };
  }

  private getMockAchievements(): IAchievement[] {
    return [
      { id: 1, code: 'FIRST_LOGIN', name: 'Welcome Aboard', description: 'Logged in for the first time', category: 'Onboarding', rarity: 'Common', points: 50, icon: 'Home', isUnlocked: true, unlockedDate: '2024-01-15' },
      { id: 2, code: 'TASK_MASTER_10', name: 'Task Master', description: 'Completed 10 onboarding tasks', category: 'Tasks', rarity: 'Uncommon', points: 100, icon: 'CheckMark', isUnlocked: true, unlockedDate: '2024-01-20' },
      { id: 3, code: 'RECOGNITION_GIVER', name: 'Kudos Champion', description: 'Gave recognition to 5 colleagues', category: 'Social', rarity: 'Uncommon', points: 75, icon: 'Like', isUnlocked: true, unlockedDate: '2024-02-01' },
      { id: 4, code: 'STREAK_7', name: 'Week Warrior', description: 'Maintained a 7-day streak', category: 'Streaks', rarity: 'Rare', points: 150, icon: 'Calories', isUnlocked: true, unlockedDate: '2024-02-10' },
      { id: 5, code: 'LEVEL_10', name: 'Grandmaster', description: 'Reached Level 10', category: 'Milestone', rarity: 'Legendary', points: 500, icon: 'Crown', isUnlocked: false, progress: 60, progressTarget: 100 }
    ];
  }

  private getMockLeaderboard(top: number): ILeaderboardEntry[] {
    const entries: ILeaderboardEntry[] = [
      { rank: 1, userId: 1, userEmail: 'alice@example.com', displayName: 'Alice Johnson', department: 'IT', points: 8500, level: 8, tier: 'Gold', achievementCount: 25, isCurrentUser: false, rankChange: 0 },
      { rank: 2, userId: 2, userEmail: 'bob@example.com', displayName: 'Bob Smith', department: 'HR', points: 7200, level: 7, tier: 'Silver', achievementCount: 20, isCurrentUser: false, rankChange: 2 },
      { rank: 3, userId: 3, userEmail: 'carol@example.com', displayName: 'Carol Williams', department: 'Finance', points: 6800, level: 7, tier: 'Silver', achievementCount: 18, isCurrentUser: false, rankChange: -1 },
      { rank: 4, userId: 4, userEmail: 'david@example.com', displayName: 'David Brown', department: 'Marketing', points: 5500, level: 6, tier: 'Silver', achievementCount: 15, isCurrentUser: false, rankChange: 1 },
      { rank: 5, userId: 5, userEmail: 'user@example.com', displayName: 'Demo User', department: 'HR', points: 3750, level: 6, tier: 'Silver', achievementCount: 12, isCurrentUser: true, rankChange: 0 },
      { rank: 6, userId: 6, userEmail: 'emma@example.com', displayName: 'Emma Davis', department: 'Operations', points: 3200, level: 5, tier: 'Silver', achievementCount: 10, isCurrentUser: false, rankChange: -2 },
      { rank: 7, userId: 7, userEmail: 'frank@example.com', displayName: 'Frank Miller', department: 'Sales', points: 2800, level: 5, tier: 'Silver', achievementCount: 9, isCurrentUser: false, rankChange: 1 },
      { rank: 8, userId: 8, userEmail: 'grace@example.com', displayName: 'Grace Lee', department: 'IT', points: 2400, level: 4, tier: 'Bronze', achievementCount: 8, isCurrentUser: false, rankChange: 0 },
      { rank: 9, userId: 9, userEmail: 'henry@example.com', displayName: 'Henry Wilson', department: 'Finance', points: 2100, level: 4, tier: 'Bronze', achievementCount: 7, isCurrentUser: false, rankChange: 3 },
      { rank: 10, userId: 10, userEmail: 'ivy@example.com', displayName: 'Ivy Taylor', department: 'HR', points: 1800, level: 3, tier: 'Bronze', achievementCount: 6, isCurrentUser: false, rankChange: -1 }
    ];
    return entries.slice(0, top);
  }

  private getMockChallenges(): IChallenge[] {
    return [
      { id: 1, code: 'ONBOARD_SPRINT', name: 'Onboarding Sprint', description: 'Complete all your onboarding tasks within your first month', type: 'Individual', category: 'Task Completion', status: 'Active', startDate: '2024-01-01', endDate: '2024-02-28', daysLeft: 14, goalTarget: 100, currentProgress: 75, progressPercent: 75, rewardPoints: 500, participants: 45, isJoined: true },
      { id: 2, code: 'RECOGNITION_WEEK', name: 'Recognition Week', description: 'Give recognition to 5 colleagues this week', type: 'Individual', category: 'Recognition', status: 'Active', startDate: '2024-02-19', endDate: '2024-02-25', daysLeft: 3, goalTarget: 5, currentProgress: 3, progressPercent: 60, rewardPoints: 200, participants: 120, isJoined: true },
      { id: 3, code: 'TEAM_SPIRIT', name: 'Team Spirit Challenge', description: 'Your department earns the most points this month', type: 'Department', category: 'Team Building', status: 'Active', startDate: '2024-02-01', endDate: '2024-02-29', daysLeft: 7, goalTarget: 10000, currentProgress: 6500, progressPercent: 65, rewardPoints: 1000, participants: 8, isJoined: true }
    ];
  }

  private getMockRewards(): IReward[] {
    return [
      { id: 1, code: 'COFFEE_VOUCHER', name: 'Coffee Voucher', description: 'R50 voucher for the cafeteria', category: 'Food & Beverage', pointsCost: 475, originalCost: 500, tierDiscount: 5, icon: 'CoffeeScript', stockLevel: 50, isAvailable: true, isFeatured: true },
      { id: 2, code: 'MOVIE_TICKET', name: 'Movie Tickets (x2)', description: 'Two cinema tickets at Ster-Kinekor', category: 'Entertainment', pointsCost: 1425, originalCost: 1500, tierDiscount: 5, icon: 'Video', stockLevel: 20, isAvailable: true, isFeatured: true },
      { id: 3, code: 'WOOLWORTHS_100', name: 'Woolworths Gift Card', description: 'R100 Woolworths shopping voucher', category: 'Shopping', pointsCost: 950, originalCost: 1000, tierDiscount: 5, icon: 'ShoppingCart', stockLevel: 30, isAvailable: true, isFeatured: false },
      { id: 4, code: 'EXTRA_LEAVE', name: 'Half Day Leave', description: 'Extra half day of paid leave', category: 'Time Off', pointsCost: 4750, originalCost: 5000, tierDiscount: 5, icon: 'Vacation', stockLevel: 10, isAvailable: true, isFeatured: true }
    ];
  }

  private getMockActivityFeed(): IActivityFeedItem[] {
    return [
      { id: 1, userId: 2, userEmail: 'bob@example.com', displayName: 'Bob Smith', type: 'achievement', message: '<strong>Bob Smith</strong> unlocked the "Task Master" achievement', icon: 'Trophy', timestamp: new Date(Date.now() - 1000 * 60 * 30).toISOString(), likes: 5 },
      { id: 2, userId: 3, userEmail: 'carol@example.com', displayName: 'Carol Williams', type: 'levelUp', message: '<strong>Carol Williams</strong> reached Level 7!', icon: 'SkypeCircleArrow', timestamp: new Date(Date.now() - 1000 * 60 * 60 * 2).toISOString(), likes: 12 },
      { id: 3, userId: 1, userEmail: 'alice@example.com', displayName: 'Alice Johnson', type: 'challenge', message: '<strong>Alice Johnson</strong> won the "February Speed Run" challenge!', icon: 'Medal', timestamp: new Date(Date.now() - 1000 * 60 * 60 * 5).toISOString(), likes: 23 },
      { id: 4, userId: 4, userEmail: 'david@example.com', displayName: 'David Brown', type: 'recognition', message: '<strong>David Brown</strong> received recognition from Emma Davis', icon: 'Like', timestamp: new Date(Date.now() - 1000 * 60 * 60 * 8).toISOString(), likes: 8 },
      { id: 5, userId: 6, userEmail: 'emma@example.com', displayName: 'Emma Davis', type: 'streak', message: '<strong>Emma Davis</strong> reached a 30-day streak!', icon: 'Calories', timestamp: new Date(Date.now() - 1000 * 60 * 60 * 24).toISOString(), likes: 15 }
    ];
  }

  private getMockStreaks(): IStreak[] {
    return [
      { id: 1, type: 'Daily Login', currentDays: 14, longestDays: 21, lastActivityDate: new Date().toISOString(), multiplier: 1.25, nextMilestone: 30, isActive: true },
      { id: 2, type: 'Task Completion', currentDays: 7, longestDays: 12, lastActivityDate: new Date().toISOString(), multiplier: 1.1, nextMilestone: 14, isActive: true }
    ];
  }

  private getMockRecognitions(): IRecognition[] {
    return [
      { id: 1, giverId: 2, giverName: 'Bob Smith', giverEmail: 'bob@example.com', receiverId: 5, receiverName: 'Demo User', receiverEmail: 'user@example.com', type: 'Great Job', message: 'Thanks for helping me with the quarterly report! Your attention to detail really made a difference.', timestamp: new Date(Date.now() - 1000 * 60 * 60 * 4).toISOString(), likes: 3, isPublic: true },
      { id: 2, giverId: 3, giverName: 'Carol Williams', giverEmail: 'carol@example.com', receiverId: 1, receiverName: 'Alice Johnson', receiverEmail: 'alice@example.com', type: 'Team Player', message: 'Alice went above and beyond to help the whole team meet our deadline. True team spirit!', timestamp: new Date(Date.now() - 1000 * 60 * 60 * 12).toISOString(), likes: 8, isPublic: true },
      { id: 3, giverId: 4, giverName: 'David Brown', giverEmail: 'david@example.com', receiverId: 6, receiverName: 'Emma Davis', receiverEmail: 'emma@example.com', type: 'Innovation', message: 'Emma\'s suggestion for streamlining our process saved us hours of work. Brilliant idea!', timestamp: new Date(Date.now() - 1000 * 60 * 60 * 24).toISOString(), likes: 12, isPublic: true }
    ];
  }

  // ============================================================================
  // User Points Management
  // ============================================================================

  /**
   * Get or create user points record
   */
  public async getUserPoints(userId: number): Promise<IUserPoints> {
    try {
      const existingPoints = await this.sp.web.lists
        .getByTitle("JML_UserPoints")
        .items.filter(`UserIdId eq ${userId}`)
        .top(1)
        .expand("UserId")();

      if (existingPoints.length > 0) {
        return existingPoints[0] as IUserPoints;
      }

      // Create new user points record
      const result = await this.sp.web.lists
        .getByTitle("JML_UserPoints")
        .items.add({
          Title: `User ${userId} Points`,
          UserIdId: userId,
          TotalPoints: 0,
          QuizPoints: 0,
          ReadingPoints: 0,
          AcknowledgementPoints: 0,
          BonusPoints: 0,
          CurrentLevel: 1,
          LevelName: "Novice",
          PoliciesRead: 0,
          QuizzesCompleted: 0,
          QuizzesPassed: 0,
          StreakDays: 0,
          LastActivityDate: new Date().toISOString()
        });

      return result.data as IUserPoints;
    } catch (error) {
      console.error("Failed to get user points:", error);
      throw error;
    }
  }

  /**
   * Add points to user
   */
  public async addPoints(
    userId: number,
    points: number,
    transactionType: string,
    referenceId?: number,
    referenceType?: string,
    description?: string
  ): Promise<IUserPoints> {
    try {
      const userPoints = await this.getUserPoints(userId);
      const newTotalPoints = userPoints.TotalPoints + points;

      // Determine category-specific points
      let updates: any = {
        TotalPoints: newTotalPoints,
        LastActivityDate: new Date().toISOString()
      };

      switch (transactionType) {
        case "Policy Read":
          updates.ReadingPoints = userPoints.ReadingPoints + points;
          updates.PoliciesRead = userPoints.PoliciesRead + 1;
          break;
        case "Acknowledgement":
          updates.AcknowledgementPoints = userPoints.AcknowledgementPoints + points;
          break;
        case "Quiz Completed":
          updates.QuizPoints = userPoints.QuizPoints + points;
          updates.QuizzesCompleted = userPoints.QuizzesCompleted + 1;
          if (points >= this.POINTS.QUIZ_PASSED) {
            updates.QuizzesPassed = userPoints.QuizzesPassed + 1;
          }
          break;
        case "Bonus":
        case "Achievement":
          updates.BonusPoints = userPoints.BonusPoints + points;
          break;
      }

      // Check for level up
      const newLevel = this.calculateLevel(newTotalPoints);
      if (newLevel.level > userPoints.CurrentLevel) {
        updates.CurrentLevel = newLevel.level;
        updates.LevelName = newLevel.name;
      }

      // Update streak
      const newStreak = await this.updateStreak(userId, userPoints);
      if (newStreak !== userPoints.StreakDays) {
        updates.StreakDays = newStreak;
      }

      // Update user points
      await this.sp.web.lists
        .getByTitle("JML_UserPoints")
        .items.getById(userPoints.Id)
        .update(updates);

      // Log transaction
      await this.logTransaction(
        userId,
        transactionType,
        points,
        newTotalPoints,
        referenceId,
        referenceType,
        description
      );

      // Check for achievements (legacy Policy Hub)
      await this.checkAndAwardLegacyAchievements(userId, { ...userPoints, ...updates });

      // Return updated points
      return await this.getUserPoints(userId);
    } catch (error) {
      console.error("Failed to add points:", error);
      throw error;
    }
  }

  /**
   * Calculate user level based on points
   */
  private calculateLevel(points: number): { level: number; name: string } {
    for (let i = this.LEVEL_THRESHOLDS.length - 1; i >= 0; i--) {
      if (points >= this.LEVEL_THRESHOLDS[i].points) {
        return {
          level: this.LEVEL_THRESHOLDS[i].level,
          name: this.LEVEL_THRESHOLDS[i].name
        };
      }
    }
    return { level: 1, name: "Novice" };
  }

  /**
   * Get next level information
   */
  public getNextLevel(currentPoints: number): {
    level: number;
    levelName: string;
    pointsRequired: number;
    pointsRemaining: number;
    progressPercentage: number;
  } {
    const currentLevel = this.calculateLevel(currentPoints);
    const nextLevelIndex = this.LEVEL_THRESHOLDS.findIndex(l => l.level === currentLevel.level + 1);

    if (nextLevelIndex === -1) {
      // Max level reached
      return {
        level: currentLevel.level,
        levelName: currentLevel.name,
        pointsRequired: 0,
        pointsRemaining: 0,
        progressPercentage: 100
      };
    }

    const nextLevel = this.LEVEL_THRESHOLDS[nextLevelIndex];
    const currentLevelPoints = this.LEVEL_THRESHOLDS.find(l => l.level === currentLevel.level)?.points || 0;
    const pointsForNextLevel = nextLevel.points - currentLevelPoints;
    const pointsEarned = currentPoints - currentLevelPoints;
    const progressPercentage = Math.min(100, Math.round((pointsEarned / pointsForNextLevel) * 100));

    return {
      level: nextLevel.level,
      levelName: nextLevel.name,
      pointsRequired: nextLevel.points,
      pointsRemaining: nextLevel.points - currentPoints,
      progressPercentage
    };
  }

  /**
   * Update user activity streak
   */
  private async updateStreak(userId: number, userPoints: IUserPoints): Promise<number> {
    try {
      const lastActivity = new Date(userPoints.LastActivityDate);
      const today = new Date();
      today.setHours(0, 0, 0, 0);

      const lastActivityDate = new Date(lastActivity);
      lastActivityDate.setHours(0, 0, 0, 0);

      const daysDifference = Math.floor((today.getTime() - lastActivityDate.getTime()) / (1000 * 60 * 60 * 24));

      if (daysDifference === 0) {
        // Activity today, maintain streak
        return userPoints.StreakDays;
      } else if (daysDifference === 1) {
        // Activity yesterday, increment streak
        const newStreak = userPoints.StreakDays + 1;

        // Award streak bonus
        if (newStreak % 7 === 0) {
          await this.addPoints(userId, this.POINTS.WEEKLY_STREAK, "Bonus", undefined, undefined, "7-day streak bonus");
        } else if (newStreak % 30 === 0) {
          await this.addPoints(userId, this.POINTS.MONTHLY_STREAK, "Bonus", undefined, undefined, "30-day streak bonus");
        } else {
          await this.addPoints(userId, this.POINTS.DAILY_STREAK, "Bonus", undefined, undefined, "Daily streak bonus");
        }

        return newStreak;
      } else {
        // Streak broken
        return 1;
      }
    } catch (error) {
      console.error("Failed to update streak:", error);
      return userPoints.StreakDays;
    }
  }

  // ============================================================================
  // Transactions
  // ============================================================================

  /**
   * Log point transaction
   */
  private async logTransaction(
    userId: number,
    transactionType: string,
    points: number,
    newBalance: number,
    referenceId?: number,
    referenceType?: string,
    description?: string
  ): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle("JML_PointTransactions")
        .items.add({
          Title: `${transactionType} - ${points} points`,
          UserIdId: userId,
          TransactionType: transactionType,
          Points: points,
          TransactionDate: new Date().toISOString(),
          ReferenceId: referenceId,
          ReferenceType: referenceType,
          Description: description || `${transactionType}: ${points} points`,
          NewBalance: newBalance
        });
    } catch (error) {
      console.error("Failed to log transaction:", error);
    }
  }

  /**
   * Get user transactions
   */
  public async getUserTransactions(userId: number, top: number = 20): Promise<IPointTransaction[]> {
    try {
      const transactions = await this.sp.web.lists
        .getByTitle("JML_PointTransactions")
        .items.filter(`UserIdId eq ${userId}`)
        .orderBy("TransactionDate", false)
        .top(top)
        .expand("UserId")();

      return transactions as IPointTransaction[];
    } catch (error) {
      console.error("Failed to get user transactions:", error);
      return [];
    }
  }

  // ============================================================================
  // Policy Hub Legacy Achievements (for PolicyHub module)
  // ============================================================================

  /**
   * Get all achievements (Legacy - Policy Hub)
   */
  public async getLegacyAllAchievements(): Promise<ILegacyAchievement[]> {
    try {
      const achievements = await this.sp.web.lists
        .getByTitle("JML_Achievements")
        .items.filter("IsActive eq true")
        .orderBy("DisplayOrder", true)();

      return achievements as ILegacyAchievement[];
    } catch (error) {
      console.error("Failed to get achievements:", error);
      return [];
    }
  }

  /**
   * Get user achievements by userId (Legacy - Policy Hub)
   */
  public async getLegacyUserAchievements(userId: number): Promise<ILegacyUserAchievement[]> {
    try {
      const achievements = await this.sp.web.lists
        .getByTitle("JML_UserAchievements")
        .items.filter(`UserIdId eq ${userId}`)
        .orderBy("EarnedDate", false)
        .expand("UserId")();

      return achievements as ILegacyUserAchievement[];
    } catch (error) {
      console.error("Failed to get user achievements:", error);
      return [];
    }
  }

  /**
   * Award achievement to user (Legacy - Policy Hub)
   */
  public async awardLegacyAchievement(userId: number, achievementId: number): Promise<void> {
    try {
      // Check if already awarded
      const existing = await this.sp.web.lists
        .getByTitle("JML_UserAchievements")
        .items.filter(`UserIdId eq ${userId} and AchievementId eq ${achievementId}`)();

      if (existing.length > 0) {
        return; // Already awarded
      }

      // Get achievement details
      const achievement = await this.sp.web.lists
        .getByTitle("JML_Achievements")
        .items.getById(achievementId)();

      // Award achievement
      await this.sp.web.lists
        .getByTitle("JML_UserAchievements")
        .items.add({
          Title: achievement.AchievementName,
          UserIdId: userId,
          AchievementId: achievementId,
          AchievementName: achievement.AchievementName,
          EarnedDate: new Date().toISOString(),
          PointsEarned: achievement.PointsReward,
          BadgeLevel: achievement.BadgeLevel,
          IsDisplayed: true
        });

      // Award points
      await this.addPoints(
        userId,
        achievement.PointsReward,
        "Achievement",
        achievementId,
        "Achievement",
        `Earned: ${achievement.AchievementName}`
      );
    } catch (error) {
      console.error("Failed to award achievement:", error);
    }
  }

  /**
   * Check and award achievements based on user progress (Legacy - Policy Hub)
   */
  private async checkAndAwardLegacyAchievements(userId: number, userPoints: IUserPoints): Promise<void> {
    try {
      const achievements = await this.getLegacyAllAchievements();
      const userAchievements = await this.getLegacyUserAchievements(userId);
      const earnedAchievementIds = userAchievements.map(a => a.AchievementId);

      for (const achievement of achievements) {
        if (earnedAchievementIds.includes(achievement.Id)) {
          continue; // Already earned
        }

        let shouldAward = false;

        switch (achievement.AchievementType) {
          case "Reading":
            shouldAward = userPoints.PoliciesRead >= achievement.RequirementValue;
            break;

          case "Quiz":
            shouldAward = userPoints.QuizzesPassed >= achievement.RequirementValue;
            break;

          case "Streak":
            shouldAward = userPoints.StreakDays >= achievement.RequirementValue;
            break;

          case "Milestone":
            shouldAward = userPoints.TotalPoints >= achievement.RequirementValue;
            break;

          case "Completion":
            // This requires custom logic per policy/quiz
            break;
        }

        if (shouldAward) {
          await this.awardLegacyAchievement(userId, achievement.Id);
        }
      }
    } catch (error) {
      console.error("Failed to check achievements:", error);
    }
  }

  // ============================================================================
  // Policy Hub Legacy Leaderboard
  // ============================================================================

  /**
   * Get leaderboard (Legacy - Policy Hub)
   */
  public async getLegacyLeaderboard(top: number = 10): Promise<ILegacyLeaderboardEntry[]> {
    try {
      const allUsers = await this.sp.web.lists
        .getByTitle("JML_UserPoints")
        .items.orderBy("TotalPoints", false)
        .top(top)
        .expand("UserId")();

      const leaderboard: ILegacyLeaderboardEntry[] = allUsers.map((user, index) => ({
        userId: user.UserId.Id,
        userName: user.UserId.Title,
        userEmail: user.UserId.Email,
        totalPoints: user.TotalPoints,
        level: user.CurrentLevel,
        levelName: user.LevelName,
        rank: index + 1,
        policiesRead: user.PoliciesRead,
        quizzesCompleted: user.QuizzesCompleted,
        streakDays: user.StreakDays
      }));

      return leaderboard;
    } catch (error) {
      console.error("Failed to get leaderboard:", error);
      return [];
    }
  }

  /**
   * Get user rank
   */
  public async getUserRank(userId: number): Promise<number> {
    try {
      const userPoints = await this.getUserPoints(userId);
      const higherRanked = await this.sp.web.lists
        .getByTitle("JML_UserPoints")
        .items.filter(`TotalPoints gt ${userPoints.TotalPoints}`)();

      return higherRanked.length + 1;
    } catch (error) {
      console.error("Failed to get user rank:", error);
      return 0;
    }
  }

  // ============================================================================
  // User Progress Dashboard
  // ============================================================================

  /**
   * Get comprehensive user progress
   */
  public async getUserProgress(userId: number): Promise<IUserProgress> {
    try {
      const userPoints = await this.getUserPoints(userId);
      const achievements = await this.getLegacyUserAchievements(userId);
      const recentTransactions = await this.getUserTransactions(userId, 10);
      const nextLevel = this.getNextLevel(userPoints.TotalPoints);

      // Calculate stats
      const oneWeekAgo = new Date();
      oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
      const thisWeekTransactions = recentTransactions.filter(
        t => new Date(t.TransactionDate) >= oneWeekAgo
      );
      const thisWeekPoints = thisWeekTransactions.reduce((sum, t) => sum + t.Points, 0);

      const oneMonthAgo = new Date();
      oneMonthAgo.setMonth(oneMonthAgo.getMonth() - 1);
      const thisMonthTransactions = recentTransactions.filter(
        t => new Date(t.TransactionDate) >= oneMonthAgo
      );
      const thisMonthPoints = thisMonthTransactions.reduce((sum, t) => sum + t.Points, 0);

      const totalActivities = userPoints.PoliciesRead + userPoints.QuizzesCompleted;

      // Calculate average quiz score (would need quiz attempts data)
      const averageQuizScore = userPoints.QuizzesCompleted > 0
        ? (userPoints.QuizzesPassed / userPoints.QuizzesCompleted) * 100
        : 0;

      return {
        userPoints,
        achievements,
        recentTransactions,
        nextLevel,
        stats: {
          totalActivities,
          thisWeekPoints,
          thisMonthPoints,
          averageQuizScore: Math.round(averageQuizScore)
        }
      };
    } catch (error) {
      console.error("Failed to get user progress:", error);
      throw error;
    }
  }

  // ============================================================================
  // Quick Actions (called by other services)
  // ============================================================================

  /**
   * Record policy read
   */
  public async recordPolicyRead(userId: number, policyId: number): Promise<void> {
    await this.addPoints(
      userId,
      this.POINTS.POLICY_READ,
      "Policy Read",
      policyId,
      "Policy",
      "Policy reading completed"
    );
  }

  /**
   * Record policy acknowledgement
   */
  public async recordPolicyAcknowledgement(userId: number, policyId: number): Promise<void> {
    await this.addPoints(
      userId,
      this.POINTS.POLICY_ACKNOWLEDGEMENT,
      "Acknowledgement",
      policyId,
      "Policy",
      "Policy acknowledged"
    );
  }

  /**
   * Record quiz completion
   */
  public async recordQuizCompletion(
    userId: number,
    quizId: number,
    passed: boolean,
    score: number
  ): Promise<void> {
    let points = 0;
    let description = "";

    if (passed) {
      if (score === 100) {
        points = this.POINTS.QUIZ_PERFECT;
        description = "Quiz passed with perfect score!";
      } else {
        points = this.POINTS.QUIZ_PASSED;
        description = "Quiz passed";
      }
    }

    if (points > 0) {
      await this.addPoints(
        userId,
        points,
        "Quiz Completed",
        quizId,
        "Quiz",
        description
      );
    }
  }

  // ============================================================================
  // GAMIFICATION HUB ACTIONS (New)
  // ============================================================================

  /**
   * Redeem a reward using points
   */
  public async redeemReward(rewardId: number): Promise<void> {
    try {
      // Get the reward details
      const rewardItem = await this.sp.web.lists
        .getByTitle(this.REWARDS_LIST)
        .items.getById(rewardId)
        .select("Id", "Title", "PointsCost", "StockLevel")();

      if (!rewardItem) {
        throw new Error("Reward not found");
      }

      if (rewardItem.StockLevel !== undefined && rewardItem.StockLevel <= 0) {
        throw new Error("Reward is out of stock");
      }

      // Get current user's points
      const profile = await this.getCurrentUserProfile();
      if (profile.totalPoints < rewardItem.PointsCost) {
        throw new Error("Insufficient points");
      }

      // Deduct points from user
      await this.addPoints(
        this.currentUserId,
        -rewardItem.PointsCost,
        "Reward Redemption",
        rewardId,
        "Reward",
        `Redeemed: ${rewardItem.Title}`
      );

      // Update stock level if tracked
      if (rewardItem.StockLevel !== undefined) {
        await this.sp.web.lists
          .getByTitle(this.REWARDS_LIST)
          .items.getById(rewardId)
          .update({
            StockLevel: rewardItem.StockLevel - 1
          });
      }

      // Create redemption record
      await this.sp.web.lists
        .getByTitle(this.REDEMPTIONS_LIST)
        .items.add({
          Title: `${rewardItem.Title} - Redemption`,
          UserId: this.currentUserId,
          RewardId: rewardId,
          PointsSpent: rewardItem.PointsCost,
          RedeemedDate: new Date().toISOString(),
          Status: "Pending"
        });

      console.log(`[GamificationService] Redeemed reward ${rewardId} for ${rewardItem.PointsCost} points`);
    } catch (error) {
      console.error("[GamificationService] Failed to redeem reward:", error);
      throw error;
    }
  }

  /**
   * Like a recognition
   */
  public async likeRecognition(recognitionId: number): Promise<void> {
    try {
      // Get current like count
      const recognition = await this.sp.web.lists
        .getByTitle(this.RECOGNITIONS_LIST)
        .items.getById(recognitionId)
        .select("Id", "Likes")();

      // Increment likes
      await this.sp.web.lists
        .getByTitle(this.RECOGNITIONS_LIST)
        .items.getById(recognitionId)
        .update({
          Likes: (recognition.Likes || 0) + 1
        });

      console.log(`[GamificationService] Liked recognition ${recognitionId}`);
    } catch (error) {
      console.error("[GamificationService] Failed to like recognition:", error);
      throw error;
    }
  }
}
