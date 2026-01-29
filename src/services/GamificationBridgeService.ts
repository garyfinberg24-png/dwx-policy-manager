// @ts-nocheck
/**
 * GamificationBridgeService
 * Bridges the Onboarding Experience and Enterprise Gamification systems
 *
 * Purpose:
 * - Unified XP tracking across onboarding and enterprise phases
 * - Achievement mapping between systems
 * - Unified leaderboard with phase-aware filtering
 * - Seamless transition from onboarding to enterprise gamification
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

// ============================================================================
// BRIDGE INTERFACES
// ============================================================================

/**
 * Unified points source - tracks where XP comes from
 */
export type PointSource =
  | 'onboarding-quest'
  | 'onboarding-bingo'
  | 'onboarding-survival'
  | 'onboarding-coffee'
  | 'onboarding-social'
  | 'onboarding-other'
  | 'policy-read'
  | 'quiz-passed'
  | 'task-completed'
  | 'challenge-completed'
  | 'recognition'
  | 'bonus';

/**
 * User lifecycle phase for contextual gamification
 */
export type LifecyclePhase = 'onboarding' | 'active' | 'mover' | 'leaver';

/**
 * Unified gamification record - single ledger entry
 */
export interface IGamificationLedgerEntry {
  id: string;
  userId: string;
  userEmail: string;
  displayName: string;
  points: number;
  source: PointSource;
  sourceDescription: string;
  phase: LifecyclePhase;
  timestamp: string;
  relatedItemId?: string;
  multiplierApplied: number;
}

/**
 * Unified user profile spanning both systems
 */
export interface IUnifiedGamificationProfile {
  // Identity
  userId: string;
  userEmail: string;
  displayName: string;
  department: string;
  photoUrl: string;

  // Lifecycle
  phase: LifecyclePhase;
  startDate: string;
  onboardingCompleted: boolean;
  onboardingCompletedDate?: string;

  // Points - Unified
  totalLifetimePoints: number;
  availablePoints: number;
  onboardingPoints: number;
  enterprisePoints: number;

  // Level & Tier
  currentLevel: number;
  levelName: string;
  currentTier: string;
  tierMultiplier: number;
  pointsToNextLevel: number;
  pointsToNextTier: number;

  // Achievements
  totalBadges: number;
  onboardingBadges: number;
  enterpriseBadges: number;

  // Streaks
  currentStreak: number;
  longestStreak: number;
  streakMultiplier: number;

  // Rankings
  globalRank: number;
  departmentRank: number;
  onboardingCohortRank: number;
}

/**
 * Achievement with cross-system mapping
 */
export interface IUnifiedAchievement {
  id: string;
  code: string;
  name: string;
  description: string;
  category: 'onboarding' | 'policy' | 'social' | 'challenge' | 'milestone';
  rarity: 'common' | 'uncommon' | 'rare' | 'epic' | 'legendary';
  pointsReward: number;
  icon: string;

  // Status
  isUnlocked: boolean;
  unlockedDate?: string;
  progress: number;
  progressTarget: number;

  // Cross-system mapping
  sourceSystem: 'onboarding' | 'enterprise' | 'both';
  linkedAchievements?: string[]; // Achievements this unlocks
  prerequisiteAchievements?: string[]; // Required to unlock this
}

/**
 * Unified leaderboard entry
 */
export interface IUnifiedLeaderboardEntry {
  rank: number;
  userId: string;
  userEmail: string;
  displayName: string;
  photoUrl: string;
  department: string;

  // Phase context
  phase: LifecyclePhase;
  startDate: string;
  daysInPhase: number;

  // Points
  totalPoints: number;
  phasePoints: number;
  level: number;
  tier: string;

  // Achievements
  badgeCount: number;
  streak: number;

  // Status
  isCurrentUser: boolean;
  trend: 'up' | 'down' | 'same';
  previousRank?: number;
}

/**
 * Achievement mapping - links onboarding achievements to enterprise rewards
 */
export interface IAchievementMapping {
  onboardingBadgeId: string;
  onboardingBadgeName: string;
  enterpriseAchievementCode: string;
  enterpriseAchievementName: string;
  bonusPointsOnSync: number;
  tierUpgradeBonus?: string; // e.g., "Silver" to start at Silver tier
}

// ============================================================================
// BRIDGE SERVICE
// ============================================================================

export class GamificationBridgeService {
  private readonly sp: SPFI;
  private readonly siteUrl: string;
  private readonly userEmail: string;
  private readonly userDisplayName: string;
  private readonly isWorkbench: boolean;

  // List names
  private readonly LISTS = {
    LEDGER: 'PM_GamificationLedger',
    PROFILES: 'PM_GamificationProfiles',
    ONBOARDING_PROGRESS: 'PM_OnboardingProgress',
    ACHIEVEMENTS: 'PM_Achievements',
    USER_ACHIEVEMENTS: 'PM_UserAchievements',
    ACHIEVEMENT_MAPPINGS: 'PM_AchievementMappings'
  };

  // Level thresholds (unified across systems)
  private readonly LEVELS = [
    { level: 1, name: 'Newcomer', points: 0 },
    { level: 2, name: 'Apprentice', points: 100 },
    { level: 3, name: 'Explorer', points: 300 },
    { level: 4, name: 'Achiever', points: 750 },
    { level: 5, name: 'Expert', points: 1500 },
    { level: 6, name: 'Master', points: 3000 },
    { level: 7, name: 'Champion', points: 6000 },
    { level: 8, name: 'Legend', points: 10000 },
    { level: 9, name: 'Elite', points: 20000 },
    { level: 10, name: 'Grandmaster', points: 50000 }
  ];

  // Tier thresholds
  private readonly TIERS = [
    { tier: 'Bronze', points: 0, multiplier: 1.0 },
    { tier: 'Silver', points: 2500, multiplier: 1.25 },
    { tier: 'Gold', points: 10000, multiplier: 1.5 },
    { tier: 'Platinum', points: 25000, multiplier: 2.0 }
  ];

  // Streak multipliers
  private readonly STREAK_MULTIPLIERS = [
    { days: 7, multiplier: 1.1 },
    { days: 14, multiplier: 1.25 },
    { days: 30, multiplier: 1.5 },
    { days: 60, multiplier: 1.75 },
    { days: 90, multiplier: 2.0 }
  ];

  // Achievement mappings - links onboarding badges to enterprise achievements
  private readonly ACHIEVEMENT_MAPPINGS: IAchievementMapping[] = [
    {
      onboardingBadgeId: 'day-one-hero',
      onboardingBadgeName: 'Day One Hero',
      enterpriseAchievementCode: 'ONBOARDING_DAY1',
      enterpriseAchievementName: 'First Day Champion',
      bonusPointsOnSync: 100
    },
    {
      onboardingBadgeId: 'quest-master',
      onboardingBadgeName: 'Quest Master',
      enterpriseAchievementCode: 'ONBOARDING_QUEST',
      enterpriseAchievementName: 'Onboarding Graduate',
      bonusPointsOnSync: 250,
      tierUpgradeBonus: 'Silver'
    },
    {
      onboardingBadgeId: 'bingo-champion',
      onboardingBadgeName: 'Bingo Champion',
      enterpriseAchievementCode: 'ONBOARDING_BINGO',
      enterpriseAchievementName: 'First Week Star',
      bonusPointsOnSync: 150
    },
    {
      onboardingBadgeId: 'coffee-connoisseur',
      onboardingBadgeName: 'Coffee Connoisseur',
      enterpriseAchievementCode: 'ONBOARDING_COFFEE',
      enterpriseAchievementName: 'Coffee Academy Graduate',
      bonusPointsOnSync: 50
    },
    {
      onboardingBadgeId: 'team-explorer',
      onboardingBadgeName: 'Team Explorer',
      enterpriseAchievementCode: 'ONBOARDING_TEAM',
      enterpriseAchievementName: 'Team Connector',
      bonusPointsOnSync: 75
    },
    {
      onboardingBadgeId: 'jargon-master',
      onboardingBadgeName: 'Jargon Master',
      enterpriseAchievementCode: 'ONBOARDING_JARGON',
      enterpriseAchievementName: 'Language Expert',
      bonusPointsOnSync: 50
    },
    {
      onboardingBadgeId: 'social-butterfly',
      onboardingBadgeName: 'Social Butterfly',
      enterpriseAchievementCode: 'ONBOARDING_SOCIAL',
      enterpriseAchievementName: 'Community Champion',
      bonusPointsOnSync: 75
    },
    {
      onboardingBadgeId: 'survival-expert',
      onboardingBadgeName: 'Survival Expert',
      enterpriseAchievementCode: 'ONBOARDING_SURVIVAL',
      enterpriseAchievementName: 'Setup Master',
      bonusPointsOnSync: 100
    }
  ];

  constructor(sp: SPFI, siteUrl: string, userEmail: string = '', userDisplayName: string = '') {
    this.sp = sp;
    this.siteUrl = siteUrl;
    this.userEmail = userEmail;
    this.userDisplayName = userDisplayName;
    this.isWorkbench = siteUrl.includes('workbench') || siteUrl.includes('localhost');
  }

  // ============================================================================
  // UNIFIED PROFILE METHODS
  // ============================================================================

  /**
   * Get unified profile combining onboarding and enterprise data
   */
  public async getUnifiedProfile(): Promise<IUnifiedGamificationProfile> {
    if (this.isWorkbench) {
      return this.getMockUnifiedProfile();
    }

    try {
      // Fetch from both systems in parallel
      const [onboardingProgress, enterpriseProfile] = await Promise.all([
        this.getOnboardingProgress(),
        this.getEnterpriseProfile()
      ]);

      // Calculate unified metrics
      const totalPoints = (onboardingProgress?.totalXp || 0) + (enterpriseProfile?.totalPoints || 0);
      const level = this.calculateLevel(totalPoints);
      const tier = this.calculateTier(totalPoints);
      const streak = this.calculateStreakMultiplier(onboardingProgress?.streakDays || enterpriseProfile?.currentStreak || 0);

      return {
        userId: enterpriseProfile?.userId || String(onboardingProgress?.userId || ''),
        userEmail: this.userEmail,
        displayName: this.userDisplayName || enterpriseProfile?.displayName || 'User',
        department: enterpriseProfile?.department || onboardingProgress?.department || '',
        photoUrl: enterpriseProfile?.photoUrl || '',

        phase: this.determinePhase(onboardingProgress),
        startDate: onboardingProgress?.startDate || new Date().toISOString(),
        onboardingCompleted: onboardingProgress?.completedSections?.length >= 5,
        onboardingCompletedDate: onboardingProgress?.onboardingCompletedDate,

        totalLifetimePoints: totalPoints,
        availablePoints: enterpriseProfile?.availablePoints || totalPoints,
        onboardingPoints: onboardingProgress?.totalXp || 0,
        enterprisePoints: enterpriseProfile?.totalPoints || 0,

        currentLevel: level.level,
        levelName: level.name,
        currentTier: tier.tier,
        tierMultiplier: tier.multiplier,
        pointsToNextLevel: this.getPointsToNextLevel(totalPoints),
        pointsToNextTier: this.getPointsToNextTier(totalPoints),

        totalBadges: (onboardingProgress?.badges?.length || 0) + (enterpriseProfile?.badgeCount || 0),
        onboardingBadges: onboardingProgress?.badges?.length || 0,
        enterpriseBadges: enterpriseProfile?.badgeCount || 0,

        currentStreak: onboardingProgress?.streakDays || enterpriseProfile?.currentStreak || 0,
        longestStreak: Math.max(onboardingProgress?.streakDays || 0, enterpriseProfile?.longestStreak || 0),
        streakMultiplier: streak,

        globalRank: enterpriseProfile?.leaderboardRank || 0,
        departmentRank: 0,
        onboardingCohortRank: onboardingProgress?.rank || 0
      };
    } catch (error) {
      console.error('Failed to get unified profile:', error);
      return this.getMockUnifiedProfile();
    }
  }

  // ============================================================================
  // XP SYNC METHODS
  // ============================================================================

  /**
   * Record XP gain - writes to unified ledger and syncs to both systems
   */
  public async recordXpGain(
    points: number,
    source: PointSource,
    description: string,
    relatedItemId?: string
  ): Promise<void> {
    if (this.isWorkbench) {
      console.log(`[Bridge] Mock XP recorded: +${points} from ${source} - ${description}`);
      return;
    }

    try {
      const phase = await this.getCurrentPhase();
      const streakMultiplier = await this.getCurrentStreakMultiplier();
      const adjustedPoints = Math.round(points * streakMultiplier);

      // 1. Write to unified ledger
      await this.sp.web.lists.getByTitle(this.LISTS.LEDGER).items.add({
        UserEmail: this.userEmail,
        DisplayName: this.userDisplayName,
        Points: adjustedPoints,
        Source: source,
        SourceDescription: description,
        Phase: phase,
        RelatedItemId: relatedItemId || '',
        MultiplierApplied: streakMultiplier
      });

      // 2. Update onboarding progress (if onboarding source)
      if (source.startsWith('onboarding-')) {
        await this.updateOnboardingXp(adjustedPoints);
      }

      // 3. Update enterprise profile (always, for unified tracking)
      await this.updateEnterprisePoints(adjustedPoints);

      console.log(`[Bridge] XP synced: +${adjustedPoints} (${points} x ${streakMultiplier}) from ${source}`);
    } catch (error) {
      console.error('Failed to record XP gain:', error);
    }
  }

  /**
   * Sync all onboarding XP to enterprise system
   * Call this when onboarding is completed
   */
  public async syncOnboardingToEnterprise(): Promise<{ pointsSynced: number; achievementsSynced: number }> {
    if (this.isWorkbench) {
      return { pointsSynced: 500, achievementsSynced: 3 };
    }

    try {
      const onboardingProgress = await this.getOnboardingProgress();
      if (!onboardingProgress) {
        return { pointsSynced: 0, achievementsSynced: 0 };
      }

      let totalPointsSynced = 0;
      let achievementsSynced = 0;

      // 1. Sync base onboarding XP
      const onboardingXp = onboardingProgress.totalXp || 0;
      if (onboardingXp > 0) {
        await this.updateEnterprisePoints(onboardingXp);
        totalPointsSynced += onboardingXp;
      }

      // 2. Sync badges → achievements with bonus points
      const badges = onboardingProgress.badges || [];
      for (const badgeId of badges) {
        const mapping = this.ACHIEVEMENT_MAPPINGS.find(m => m.onboardingBadgeId === badgeId);
        if (mapping) {
          const synced = await this.syncAchievement(mapping);
          if (synced) {
            achievementsSynced++;
            totalPointsSynced += mapping.bonusPointsOnSync;
          }
        }
      }

      // 3. Mark onboarding as completed in enterprise
      await this.markOnboardingCompleted();

      console.log(`[Bridge] Onboarding sync complete: ${totalPointsSynced} points, ${achievementsSynced} achievements`);
      return { pointsSynced: totalPointsSynced, achievementsSynced };
    } catch (error) {
      console.error('Failed to sync onboarding to enterprise:', error);
      return { pointsSynced: 0, achievementsSynced: 0 };
    }
  }

  // ============================================================================
  // ACHIEVEMENT METHODS
  // ============================================================================

  /**
   * Get unified achievements from both systems
   */
  public async getUnifiedAchievements(): Promise<IUnifiedAchievement[]> {
    if (this.isWorkbench) {
      return this.getMockUnifiedAchievements();
    }

    try {
      // Get onboarding badges
      const onboardingBadges = await this.getOnboardingBadges();

      // Get enterprise achievements
      const enterpriseAchievements = await this.getEnterpriseAchievements();

      // Merge and deduplicate
      const unified: IUnifiedAchievement[] = [];

      // Add onboarding badges
      for (const badge of onboardingBadges) {
        const mapping = this.ACHIEVEMENT_MAPPINGS.find(m => m.onboardingBadgeId === badge.id);
        unified.push({
          id: badge.id,
          code: badge.id,
          name: badge.name,
          description: badge.description,
          category: 'onboarding',
          rarity: badge.rarity as IUnifiedAchievement['rarity'],
          pointsReward: mapping?.bonusPointsOnSync || 50,
          icon: badge.iconUrl || 'Badge',
          isUnlocked: !!badge.earnedDate,
          unlockedDate: badge.earnedDate,
          progress: badge.earnedDate ? 100 : 0,
          progressTarget: 100,
          sourceSystem: 'onboarding',
          linkedAchievements: mapping ? [mapping.enterpriseAchievementCode] : undefined
        });
      }

      // Add enterprise achievements (that aren't linked from onboarding)
      for (const achievement of enterpriseAchievements) {
        const isLinked = this.ACHIEVEMENT_MAPPINGS.some(
          m => m.enterpriseAchievementCode === achievement.code
        );
        if (!isLinked) {
          unified.push({
            id: String(achievement.id),
            code: achievement.code,
            name: achievement.name,
            description: achievement.description,
            category: achievement.category as IUnifiedAchievement['category'],
            rarity: achievement.rarity as IUnifiedAchievement['rarity'],
            pointsReward: achievement.points,
            icon: achievement.icon,
            isUnlocked: achievement.isUnlocked,
            unlockedDate: achievement.unlockedDate,
            progress: achievement.progress || 0,
            progressTarget: achievement.progressTarget || 100,
            sourceSystem: 'enterprise'
          });
        }
      }

      return unified.sort((a, b) => {
        // Sort: unlocked first, then by rarity
        if (a.isUnlocked !== b.isUnlocked) return a.isUnlocked ? -1 : 1;
        const rarityOrder = { legendary: 0, epic: 1, rare: 2, uncommon: 3, common: 4 };
        return rarityOrder[a.rarity] - rarityOrder[b.rarity];
      });
    } catch (error) {
      console.error('Failed to get unified achievements:', error);
      return this.getMockUnifiedAchievements();
    }
  }

  // ============================================================================
  // UNIFIED LEADERBOARD
  // ============================================================================

  /**
   * Get unified leaderboard with phase filtering
   */
  public async getUnifiedLeaderboard(
    filter: 'all' | 'onboarding' | 'department' | 'cohort' = 'all',
    top: number = 20
  ): Promise<IUnifiedLeaderboardEntry[]> {
    if (this.isWorkbench) {
      return this.getMockUnifiedLeaderboard(filter, top);
    }

    try {
      // Fetch from both systems
      const [onboardingLeaderboard, enterpriseLeaderboard] = await Promise.all([
        this.getOnboardingLeaderboard(),
        this.getEnterpriseLeaderboard(top)
      ]);

      // Merge and unify
      const userMap = new Map<string, IUnifiedLeaderboardEntry>();

      // Add onboarding entries
      for (const entry of onboardingLeaderboard) {
        userMap.set(entry.userId, {
          rank: entry.rank,
          userId: entry.userId,
          userEmail: entry.userId,
          displayName: entry.displayName,
          photoUrl: entry.photoUrl || '',
          department: entry.department,
          phase: 'onboarding',
          startDate: entry.startDate,
          daysInPhase: this.daysSince(entry.startDate),
          totalPoints: entry.totalXp,
          phasePoints: entry.totalXp,
          level: entry.level,
          tier: this.calculateTier(entry.totalXp).tier,
          badgeCount: entry.badgeCount,
          streak: entry.streak,
          isCurrentUser: entry.isCurrentUser,
          trend: entry.trend,
          previousRank: entry.previousRank
        });
      }

      // Merge enterprise entries (add points to existing or create new)
      for (const entry of enterpriseLeaderboard) {
        const existing = userMap.get(entry.userEmail);
        if (existing) {
          // User exists in both - combine points
          existing.totalPoints += entry.points;
          existing.phase = 'active'; // They've moved to enterprise
        } else {
          // Enterprise-only user
          userMap.set(entry.userEmail, {
            rank: entry.rank,
            userId: String(entry.userId),
            userEmail: entry.userEmail,
            displayName: entry.displayName,
            photoUrl: '',
            department: entry.department,
            phase: 'active',
            startDate: '',
            daysInPhase: 0,
            totalPoints: entry.points,
            phasePoints: entry.points,
            level: entry.level,
            tier: entry.tier,
            badgeCount: entry.achievementCount,
            streak: 0,
            isCurrentUser: entry.isCurrentUser,
            trend: entry.rankChange && entry.rankChange > 0 ? 'up' : entry.rankChange && entry.rankChange < 0 ? 'down' : 'same',
            previousRank: entry.rank - (entry.rankChange || 0)
          });
        }
      }

      // Convert to array and filter
      let entries = Array.from(userMap.values());

      if (filter === 'onboarding') {
        entries = entries.filter(e => e.phase === 'onboarding');
      } else if (filter === 'department') {
        const userDept = entries.find(e => e.isCurrentUser)?.department;
        if (userDept) {
          entries = entries.filter(e => e.department === userDept);
        }
      }

      // Sort by total points and assign ranks
      entries.sort((a, b) => b.totalPoints - a.totalPoints);
      entries.forEach((entry, index) => {
        entry.rank = index + 1;
      });

      return entries.slice(0, top);
    } catch (error) {
      console.error('Failed to get unified leaderboard:', error);
      return this.getMockUnifiedLeaderboard(filter, top);
    }
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  private calculateLevel(points: number): { level: number; name: string } {
    for (let i = this.LEVELS.length - 1; i >= 0; i--) {
      if (points >= this.LEVELS[i].points) {
        return this.LEVELS[i];
      }
    }
    return this.LEVELS[0];
  }

  private calculateTier(points: number): { tier: string; multiplier: number } {
    for (let i = this.TIERS.length - 1; i >= 0; i--) {
      if (points >= this.TIERS[i].points) {
        return this.TIERS[i];
      }
    }
    return this.TIERS[0];
  }

  private calculateStreakMultiplier(days: number): number {
    for (let i = this.STREAK_MULTIPLIERS.length - 1; i >= 0; i--) {
      if (days >= this.STREAK_MULTIPLIERS[i].days) {
        return this.STREAK_MULTIPLIERS[i].multiplier;
      }
    }
    return 1.0;
  }

  private getPointsToNextLevel(currentPoints: number): number {
    const currentLevel = this.calculateLevel(currentPoints);
    const nextLevelIndex = this.LEVELS.findIndex(l => l.level === currentLevel.level) + 1;
    if (nextLevelIndex >= this.LEVELS.length) return 0;
    return this.LEVELS[nextLevelIndex].points - currentPoints;
  }

  private getPointsToNextTier(currentPoints: number): number {
    const currentTier = this.calculateTier(currentPoints);
    const nextTierIndex = this.TIERS.findIndex(t => t.tier === currentTier.tier) + 1;
    if (nextTierIndex >= this.TIERS.length) return 0;
    return this.TIERS[nextTierIndex].points - currentPoints;
  }

  private daysSince(dateString: string): number {
    if (!dateString) return 0;
    const date = new Date(dateString);
    const now = new Date();
    return Math.floor((now.getTime() - date.getTime()) / (1000 * 60 * 60 * 24));
  }

  private determinePhase(progress: IOnboardingProgressData | null): LifecyclePhase {
    if (!progress) return 'onboarding';
    const daysOnboarding = this.daysSince(progress.startDate);
    if (daysOnboarding > 30 || (progress.completedSections?.length || 0) >= 8) {
      return 'active';
    }
    return 'onboarding';
  }

  // ============================================================================
  // DATA ACCESS METHODS (Private)
  // ============================================================================

  private async getOnboardingProgress(): Promise<IOnboardingProgressData | null> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.LISTS.ONBOARDING_PROGRESS)
        .items.filter(`UserEmail eq '${this.userEmail}'`)
        .top(1)();

      if (items.length === 0) return null;

      const item = items[0];
      return {
        userId: item.Id,
        totalXp: item.TotalXP || 0,
        startDate: item.StartDate || '',
        streakDays: item.StreakDays || 0,
        badges: this.parseJson(item.BadgesEarned) || [],
        completedSections: this.parseJson(item.CompletedSections) || [],
        department: item.Department || '',
        rank: item.LeaderboardRank || 0,
        onboardingCompletedDate: item.OnboardingCompletedDate
      };
    } catch (error) {
      console.error('Failed to get onboarding progress:', error);
      return null;
    }
  }

  private async getEnterpriseProfile(): Promise<IEnterpriseProfileData | null> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.LISTS.PROFILES)
        .items.filter(`UserEmail eq '${this.userEmail}'`)
        .top(1)();

      if (items.length === 0) return null;

      const item = items[0];
      return {
        userId: String(item.Id),
        displayName: item.DisplayName || '',
        department: item.Department || '',
        photoUrl: item.PhotoURL || '',
        totalPoints: item.TotalPoints || 0,
        availablePoints: item.AvailablePoints || 0,
        badgeCount: item.BadgeCount || 0,
        currentStreak: item.CurrentStreak || 0,
        longestStreak: item.LongestStreak || 0,
        leaderboardRank: item.LeaderboardRank || 0
      };
    } catch (error) {
      console.error('Failed to get enterprise profile:', error);
      return null;
    }
  }

  private async updateOnboardingXp(points: number): Promise<void> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.LISTS.ONBOARDING_PROGRESS)
        .items.filter(`UserEmail eq '${this.userEmail}'`)
        .top(1)();

      if (items.length > 0) {
        const currentXp = items[0].TotalXP || 0;
        await this.sp.web.lists
          .getByTitle(this.LISTS.ONBOARDING_PROGRESS)
          .items.getById(items[0].Id)
          .update({ TotalXP: currentXp + points });
      }
    } catch (error) {
      console.error('Failed to update onboarding XP:', error);
    }
  }

  private async updateEnterprisePoints(points: number): Promise<void> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.LISTS.PROFILES)
        .items.filter(`UserEmail eq '${this.userEmail}'`)
        .top(1)();

      if (items.length > 0) {
        const current = items[0];
        await this.sp.web.lists
          .getByTitle(this.LISTS.PROFILES)
          .items.getById(current.Id)
          .update({
            TotalPoints: (current.TotalPoints || 0) + points,
            AvailablePoints: (current.AvailablePoints || 0) + points,
            LifetimePoints: (current.LifetimePoints || 0) + points
          });
      } else {
        // Create new profile
        await this.sp.web.lists.getByTitle(this.LISTS.PROFILES).items.add({
          UserEmail: this.userEmail,
          DisplayName: this.userDisplayName,
          TotalPoints: points,
          AvailablePoints: points,
          LifetimePoints: points
        });
      }
    } catch (error) {
      console.error('Failed to update enterprise points:', error);
    }
  }

  private async getCurrentPhase(): Promise<LifecyclePhase> {
    const progress = await this.getOnboardingProgress();
    return this.determinePhase(progress);
  }

  private async getCurrentStreakMultiplier(): Promise<number> {
    const progress = await this.getOnboardingProgress();
    return this.calculateStreakMultiplier(progress?.streakDays || 0);
  }

  private async syncAchievement(mapping: IAchievementMapping): Promise<boolean> {
    try {
      // Check if already synced
      const existing = await this.sp.web.lists
        .getByTitle(this.LISTS.USER_ACHIEVEMENTS)
        .items.filter(`UserEmail eq '${this.userEmail}' and AchievementCode eq '${mapping.enterpriseAchievementCode}'`)
        .top(1)();

      if (existing.length > 0) return false;

      // Create enterprise achievement
      await this.sp.web.lists.getByTitle(this.LISTS.USER_ACHIEVEMENTS).items.add({
        UserEmail: this.userEmail,
        AchievementCode: mapping.enterpriseAchievementCode,
        AchievementName: mapping.enterpriseAchievementName,
        UnlockedDate: new Date().toISOString(),
        SourceBadge: mapping.onboardingBadgeId
      });

      // Award bonus points
      if (mapping.bonusPointsOnSync > 0) {
        await this.updateEnterprisePoints(mapping.bonusPointsOnSync);
      }

      return true;
    } catch (error) {
      console.error('Failed to sync achievement:', error);
      return false;
    }
  }

  private async markOnboardingCompleted(): Promise<void> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.LISTS.ONBOARDING_PROGRESS)
        .items.filter(`UserEmail eq '${this.userEmail}'`)
        .top(1)();

      if (items.length > 0) {
        await this.sp.web.lists
          .getByTitle(this.LISTS.ONBOARDING_PROGRESS)
          .items.getById(items[0].Id)
          .update({ OnboardingCompletedDate: new Date().toISOString() });
      }
    } catch (error) {
      console.error('Failed to mark onboarding completed:', error);
    }
  }

  private async getOnboardingBadges(): Promise<IOnboardingBadgeData[]> {
    try {
      const progress = await this.getOnboardingProgress();
      const badgeIds = progress?.badges || [];

      // Get badge details
      const badges: IOnboardingBadgeData[] = [];
      for (const id of badgeIds) {
        badges.push({
          id,
          name: this.ACHIEVEMENT_MAPPINGS.find(m => m.onboardingBadgeId === id)?.onboardingBadgeName || id,
          description: '',
          rarity: 'common',
          iconUrl: '',
          earnedDate: new Date().toISOString()
        });
      }
      return badges;
    } catch (error) {
      console.error('Failed to get onboarding badges:', error);
      return [];
    }
  }

  private async getEnterpriseAchievements(): Promise<IEnterpriseAchievementData[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.LISTS.ACHIEVEMENTS)
        .items.filter('IsActive eq true')();

      const userItems = await this.sp.web.lists
        .getByTitle(this.LISTS.USER_ACHIEVEMENTS)
        .items.filter(`UserEmail eq '${this.userEmail}'`)();

      const unlockedCodes = userItems.map((u: Record<string, unknown>) => u.AchievementCode);

      return items.map((a: Record<string, unknown>) => ({
        id: a.Id as number,
        code: (a.AchievementCode as string) || '',
        name: (a.AchievementName as string) || (a.Title as string) || '',
        description: (a.AchievementDescription as string) || '',
        category: (a.AchievementCategory as string) || 'General',
        rarity: (a.Rarity as string) || 'Common',
        points: (a.PointsReward as number) || 0,
        icon: (a.IconName as string) || 'Trophy',
        isUnlocked: unlockedCodes.includes(a.AchievementCode),
        unlockedDate: userItems.find((u: Record<string, unknown>) => u.AchievementCode === a.AchievementCode)?.UnlockedDate as string,
        progress: (a.Progress as number) || 0,
        progressTarget: (a.ProgressTarget as number) || 100
      }));
    } catch (error) {
      console.error('Failed to get enterprise achievements:', error);
      return [];
    }
  }

  private async getOnboardingLeaderboard(): Promise<IOnboardingLeaderboardData[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.LISTS.ONBOARDING_PROGRESS)
        .items.orderBy('TotalXP', false)
        .top(50)();

      return items.map((item: Record<string, unknown>, index: number) => ({
        rank: index + 1,
        userId: (item.UserEmail as string) || '',
        displayName: (item.UserDisplayName as string) || 'User',
        photoUrl: (item.PhotoURL as string) || '',
        department: (item.Department as string) || '',
        startDate: (item.StartDate as string) || '',
        totalXp: (item.TotalXP as number) || 0,
        level: Math.floor(((item.TotalXP as number) || 0) / 100) + 1,
        badgeCount: (this.parseJson(item.BadgesEarned as string) || []).length,
        streak: (item.StreakDays as number) || 0,
        isCurrentUser: item.UserEmail === this.userEmail,
        trend: 'same' as const,
        previousRank: undefined
      }));
    } catch (error) {
      console.error('Failed to get onboarding leaderboard:', error);
      return [];
    }
  }

  private async getEnterpriseLeaderboard(top: number): Promise<IEnterpriseLeaderboardData[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle('PM_Leaderboard')
        .items.filter("IsCurrent eq true")
        .orderBy('LeaderboardRank', true)
        .top(top)();

      return items.map((item: Record<string, unknown>) => ({
        rank: (item.LeaderboardRank as number) || 0,
        userId: item.Id as number,
        userEmail: (item.UserEmail as string) || '',
        displayName: (item.DisplayName as string) || 'User',
        department: (item.Department as string) || '',
        points: (item.Points as number) || 0,
        level: (item.UserLevel as number) || 1,
        tier: (item.UserTier as string) || 'Bronze',
        achievementCount: (item.AchievementCount as number) || 0,
        isCurrentUser: item.UserEmail === this.userEmail,
        rankChange: (item.RankChange as number) || 0
      }));
    } catch (error) {
      console.error('Failed to get enterprise leaderboard:', error);
      return [];
    }
  }

  private parseJson(value: string | undefined | null): string[] {
    if (!value) return [];
    try {
      const parsed = JSON.parse(value);
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }

  // ============================================================================
  // MOCK DATA FOR WORKBENCH
  // ============================================================================

  private getMockUnifiedProfile(): IUnifiedGamificationProfile {
    return {
      userId: 'mock-user-1',
      userEmail: this.userEmail || 'user@company.com',
      displayName: this.userDisplayName || 'Mock User',
      department: 'Engineering',
      photoUrl: '',
      phase: 'onboarding',
      startDate: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString(),
      onboardingCompleted: false,
      totalLifetimePoints: 485,
      availablePoints: 485,
      onboardingPoints: 285,
      enterprisePoints: 200,
      currentLevel: 4,
      levelName: 'Achiever',
      currentTier: 'Bronze',
      tierMultiplier: 1.0,
      pointsToNextLevel: 265,
      pointsToNextTier: 2015,
      totalBadges: 5,
      onboardingBadges: 3,
      enterpriseBadges: 2,
      currentStreak: 7,
      longestStreak: 7,
      streakMultiplier: 1.1,
      globalRank: 42,
      departmentRank: 8,
      onboardingCohortRank: 5
    };
  }

  private getMockUnifiedAchievements(): IUnifiedAchievement[] {
    return [
      {
        id: 'day-one-hero',
        code: 'day-one-hero',
        name: 'Day One Hero',
        description: 'Completed all first day tasks',
        category: 'onboarding',
        rarity: 'common',
        pointsReward: 100,
        icon: 'Badge',
        isUnlocked: true,
        unlockedDate: new Date(Date.now() - 5 * 24 * 60 * 60 * 1000).toISOString(),
        progress: 100,
        progressTarget: 100,
        sourceSystem: 'onboarding',
        linkedAchievements: ['ONBOARDING_DAY1']
      },
      {
        id: 'coffee-connoisseur',
        code: 'coffee-connoisseur',
        name: 'Coffee Connoisseur',
        description: 'Completed Coffee Academy',
        category: 'onboarding',
        rarity: 'common',
        pointsReward: 50,
        icon: 'CoffeeScript',
        isUnlocked: true,
        unlockedDate: new Date(Date.now() - 3 * 24 * 60 * 60 * 1000).toISOString(),
        progress: 100,
        progressTarget: 100,
        sourceSystem: 'onboarding',
        linkedAchievements: ['ONBOARDING_COFFEE']
      },
      {
        id: 'quest-master',
        code: 'quest-master',
        name: 'Quest Master',
        description: 'Complete all onboarding quest stages',
        category: 'onboarding',
        rarity: 'rare',
        pointsReward: 250,
        icon: 'Trophy',
        isUnlocked: false,
        progress: 60,
        progressTarget: 100,
        sourceSystem: 'onboarding',
        linkedAchievements: ['ONBOARDING_QUEST']
      },
      {
        id: 'policy-pioneer',
        code: 'POLICY_PIONEER',
        name: 'Policy Pioneer',
        description: 'Read your first 10 policies',
        category: 'policy',
        rarity: 'uncommon',
        pointsReward: 75,
        icon: 'ReadingMode',
        isUnlocked: false,
        progress: 3,
        progressTarget: 10,
        sourceSystem: 'enterprise'
      }
    ];
  }

  private getMockUnifiedLeaderboard(filter: string, top: number): IUnifiedLeaderboardEntry[] {
    const mockUsers = [
      { name: 'Alex Thompson', dept: 'Engineering', points: 1250, phase: 'active' as const, days: 45 },
      { name: 'Maya Patel', dept: 'Product', points: 920, phase: 'active' as const, days: 30 },
      { name: 'Chris Johnson', dept: 'Engineering', points: 780, phase: 'onboarding' as const, days: 14 },
      { name: 'Sarah Williams', dept: 'Marketing', points: 650, phase: 'onboarding' as const, days: 10 },
      { name: this.userDisplayName || 'You', dept: 'Engineering', points: 485, phase: 'onboarding' as const, days: 7 },
      { name: 'David Kim', dept: 'Design', points: 420, phase: 'onboarding' as const, days: 5 },
      { name: 'Emma Davis', dept: 'Sales', points: 380, phase: 'active' as const, days: 60 },
      { name: 'Michael Brown', dept: 'Engineering', points: 320, phase: 'onboarding' as const, days: 3 },
      { name: 'Lisa Anderson', dept: 'HR', points: 280, phase: 'active' as const, days: 90 },
      { name: 'James Wilson', dept: 'Finance', points: 240, phase: 'onboarding' as const, days: 2 }
    ];

    let filtered = mockUsers;
    if (filter === 'onboarding') {
      filtered = mockUsers.filter(u => u.phase === 'onboarding');
    }

    return filtered.slice(0, top).map((user, index) => ({
      rank: index + 1,
      userId: `user-${index}`,
      userEmail: `${user.name.toLowerCase().replace(' ', '.')}@company.com`,
      displayName: user.name,
      photoUrl: '',
      department: user.dept,
      phase: user.phase,
      startDate: new Date(Date.now() - user.days * 24 * 60 * 60 * 1000).toISOString(),
      daysInPhase: user.days,
      totalPoints: user.points,
      phasePoints: user.points,
      level: this.calculateLevel(user.points).level,
      tier: this.calculateTier(user.points).tier,
      badgeCount: Math.floor(user.points / 150),
      streak: Math.min(user.days, 14),
      isCurrentUser: user.name === (this.userDisplayName || 'You'),
      trend: index < 3 ? 'up' as const : index > 7 ? 'down' as const : 'same' as const,
      previousRank: index < 3 ? index + 2 : index > 7 ? index : index + 1
    }));
  }
}

// ============================================================================
// INTERNAL DATA INTERFACES
// ============================================================================

interface IOnboardingProgressData {
  userId: number;
  totalXp: number;
  startDate: string;
  streakDays: number;
  badges: string[];
  completedSections: string[];
  department: string;
  rank: number;
  onboardingCompletedDate?: string;
}

interface IEnterpriseProfileData {
  userId: string;
  displayName: string;
  department: string;
  photoUrl: string;
  totalPoints: number;
  availablePoints: number;
  badgeCount: number;
  currentStreak: number;
  longestStreak: number;
  leaderboardRank: number;
}

interface IOnboardingBadgeData {
  id: string;
  name: string;
  description: string;
  rarity: string;
  iconUrl: string;
  earnedDate?: string;
}

interface IEnterpriseAchievementData {
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
  progress: number;
  progressTarget: number;
}

interface IOnboardingLeaderboardData {
  rank: number;
  userId: string;
  displayName: string;
  photoUrl: string;
  department: string;
  startDate: string;
  totalXp: number;
  level: number;
  badgeCount: number;
  streak: number;
  isCurrentUser: boolean;
  trend: 'up' | 'down' | 'same';
  previousRank?: number;
}

interface IEnterpriseLeaderboardData {
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
  rankChange: number;
}

export default GamificationBridgeService;
