// @ts-nocheck
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { WebPartContext } from "@microsoft/sp-webpart-base";

// Admin Interfaces
export interface IAdminReward {
  id: number;
  code: string;
  name: string;
  description: string;
  category: string;
  pointsCost: number;
  icon: string;
  imageUrl?: string;
  stockLevel: number;
  isAvailable: boolean;
  isFeatured: boolean;
}

export interface IAdminChallenge {
  id: number;
  code: string;
  name: string;
  description: string;
  type: string;
  category: string;
  status: string;
  startDate: string;
  endDate: string;
  goalTarget: number;
  rewardPoints: number;
  participants: number;
}

export interface IAdminAchievement {
  id: number;
  code: string;
  name: string;
  description: string;
  category: string;
  rarity: string;
  points: number;
  icon: string;
  requirementType: string;
  requirementValue: number;
  isActive: boolean;
}

export interface IRedemptionRequest {
  id: number;
  rewardId: number;
  rewardName: string;
  rewardCode: string;
  userId: number;
  userName: string;
  userEmail: string;
  pointsCost: number;
  requestDate: string;
  status: string;
  notes?: string;
}

export interface ITierConfig {
  id: number;
  name: string;
  displayName: string;
  minPoints: number;
  multiplier: number;
  discount: number;
  color: string;
  icon: string;
}

export interface IPointRule {
  id: number;
  code: string;
  name: string;
  description: string;
  points: number;
  isActive: boolean;
  category: string;
}

export interface IAdminStats {
  totalRewards: number;
  activeChallenges: number;
  pendingRedemptions: number;
  totalUsers: number;
}

export class GamificationAdminService {
  private sp: SPFI;
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
    this.sp = spfi().using(SPFx(context));
  }

  public async getStats(): Promise<IAdminStats> {
    try {
      const [rewards, challenges, redemptions, profiles] = await Promise.all([
        this.sp.web.lists.getByTitle("PM_RewardsCatalog").items.select("Id")(),
        this.sp.web.lists.getByTitle("PM_Challenges").items.filter("ChallengeStatus eq 'Active'").select("Id")(),
        this.sp.web.lists.getByTitle("PM_RewardRedemptions").items.filter("RedemptionStatus eq 'Pending'").select("Id")(),
        this.sp.web.lists.getByTitle("PM_GamificationProfiles").items.select("Id")()
      ]);
      return { totalRewards: rewards.length, activeChallenges: challenges.length, pendingRedemptions: redemptions.length, totalUsers: profiles.length };
    } catch { return { totalRewards: 12, activeChallenges: 5, pendingRedemptions: 8, totalUsers: 156 }; }
  }

  public async getRewards(): Promise<IAdminReward[]> {
    try {
      const items = await this.sp.web.lists.getByTitle("PM_RewardsCatalog").items.orderBy("Title", true)();
      return items.map((item: any) => ({
        id: item.Id, code: item.RewardCode || '', name: item.RewardName || item.Title || '', description: item.RewardDescription || '',
        category: item.RewardCategory || 'General', pointsCost: item.PointsCost || 0, icon: item.RewardIcon || 'Gift',
        imageUrl: item.RewardImageUrl?.Url || '', stockLevel: item.StockLevel || 0, isAvailable: item.IsAvailable !== false, isFeatured: item.IsFeatured === true
      }));
    } catch { return this.getMockRewards(); }
  }

  public async createReward(reward: IAdminReward): Promise<IAdminReward> {
    const result = await this.sp.web.lists.getByTitle("PM_RewardsCatalog").items.add({
      Title: reward.name, RewardCode: reward.code || `RWD_${Date.now()}`, RewardName: reward.name, RewardDescription: reward.description,
      RewardCategory: reward.category, PointsCost: reward.pointsCost, RewardIcon: reward.icon, StockLevel: reward.stockLevel, IsAvailable: reward.isAvailable, IsFeatured: reward.isFeatured
    });
    return { ...reward, id: result.data.Id };
  }

  public async updateReward(reward: IAdminReward): Promise<void> {
    await this.sp.web.lists.getByTitle("PM_RewardsCatalog").items.getById(reward.id).update({
      Title: reward.name, RewardCode: reward.code, RewardName: reward.name, RewardDescription: reward.description,
      RewardCategory: reward.category, PointsCost: reward.pointsCost, RewardIcon: reward.icon, StockLevel: reward.stockLevel, IsAvailable: reward.isAvailable, IsFeatured: reward.isFeatured
    });
  }

  public async deleteReward(id: number): Promise<void> {
    await this.sp.web.lists.getByTitle("PM_RewardsCatalog").items.getById(id).delete();
  }

  public async getChallenges(): Promise<IAdminChallenge[]> {
    try {
      const items = await this.sp.web.lists.getByTitle("PM_Challenges").items.orderBy("Created", false)();
      return items.map((item: any) => ({
        id: item.Id, code: item.ChallengeCode || '', name: item.ChallengeName || item.Title || '', description: item.ChallengeDescription || '',
        type: item.ChallengeType || 'Individual', category: item.ChallengeCategory || 'General', status: item.ChallengeStatus || 'Draft',
        startDate: item.StartDate || '', endDate: item.EndDate || '', goalTarget: item.GoalTarget || 100, rewardPoints: item.PointsForCompletion || 0, participants: item.TotalParticipants || 0
      }));
    } catch { return this.getMockChallenges(); }
  }

  public async createChallenge(challenge: IAdminChallenge): Promise<IAdminChallenge> {
    const result = await this.sp.web.lists.getByTitle("PM_Challenges").items.add({
      Title: challenge.name, ChallengeCode: challenge.code || `CHL_${Date.now()}`, ChallengeName: challenge.name, ChallengeDescription: challenge.description,
      ChallengeType: challenge.type, ChallengeCategory: challenge.category, ChallengeStatus: challenge.status, StartDate: challenge.startDate, EndDate: challenge.endDate,
      GoalTarget: challenge.goalTarget, PointsForCompletion: challenge.rewardPoints, TotalParticipants: 0
    });
    return { ...challenge, id: result.data.Id };
  }

  public async updateChallenge(challenge: IAdminChallenge): Promise<void> {
    await this.sp.web.lists.getByTitle("PM_Challenges").items.getById(challenge.id).update({
      Title: challenge.name, ChallengeCode: challenge.code, ChallengeName: challenge.name, ChallengeDescription: challenge.description,
      ChallengeType: challenge.type, ChallengeCategory: challenge.category, ChallengeStatus: challenge.status, StartDate: challenge.startDate, EndDate: challenge.endDate,
      GoalTarget: challenge.goalTarget, PointsForCompletion: challenge.rewardPoints
    });
  }

  public async deleteChallenge(id: number): Promise<void> {
    await this.sp.web.lists.getByTitle("PM_Challenges").items.getById(id).delete();
  }

  public async getAchievements(): Promise<IAdminAchievement[]> {
    try {
      const items = await this.sp.web.lists.getByTitle("PM_Achievements").items.orderBy("DisplayOrder", true)();
      return items.map((item: any) => ({
        id: item.Id, code: item.AchievementCode || '', name: item.AchievementName || item.Title || '', description: item.AchievementDescription || '',
        category: item.AchievementCategory || 'General', rarity: item.Rarity || 'Common', points: item.PointsReward || 0, icon: item.IconName || 'Medal',
        requirementType: item.RequirementType || 'Count', requirementValue: item.RequirementValue || 1, isActive: item.IsActive !== false
      }));
    } catch { return this.getMockAchievements(); }
  }

  public async createAchievement(achievement: IAdminAchievement): Promise<IAdminAchievement> {
    const result = await this.sp.web.lists.getByTitle("PM_Achievements").items.add({
      Title: achievement.name, AchievementCode: achievement.code || `ACH_${Date.now()}`, AchievementName: achievement.name, AchievementDescription: achievement.description,
      AchievementCategory: achievement.category, Rarity: achievement.rarity, PointsReward: achievement.points, IconName: achievement.icon,
      RequirementType: achievement.requirementType, RequirementValue: achievement.requirementValue, IsActive: achievement.isActive, DisplayOrder: 100
    });
    return { ...achievement, id: result.data.Id };
  }

  public async updateAchievement(achievement: IAdminAchievement): Promise<void> {
    await this.sp.web.lists.getByTitle("PM_Achievements").items.getById(achievement.id).update({
      Title: achievement.name, AchievementCode: achievement.code, AchievementName: achievement.name, AchievementDescription: achievement.description,
      AchievementCategory: achievement.category, Rarity: achievement.rarity, PointsReward: achievement.points, IconName: achievement.icon,
      RequirementType: achievement.requirementType, RequirementValue: achievement.requirementValue, IsActive: achievement.isActive
    });
  }

  public async deleteAchievement(id: number): Promise<void> {
    await this.sp.web.lists.getByTitle("PM_Achievements").items.getById(id).delete();
  }

  public async getPendingRedemptions(): Promise<IRedemptionRequest[]> {
    try {
      const items = await this.sp.web.lists.getByTitle("PM_RewardRedemptions").items.filter("RedemptionStatus eq 'Pending' or RedemptionStatus eq 'Approved'").orderBy("RedemptionDate", false).expand("RedeemedBy")();
      return items.map((item: any) => ({
        id: item.Id, rewardId: item.RewardItemId || 0, rewardName: item.RewardName || '', rewardCode: item.RewardCode || '',
        userId: item.RedeemedById || 0, userName: item.RedeemedBy?.Title || 'Unknown', userEmail: item.RedeemedBy?.EMail || item.UserEmail || '',
        pointsCost: item.PointsSpent || 0, requestDate: item.RedemptionDate || item.Created || '', status: item.RedemptionStatus || 'Pending', notes: item.AdminNotes || ''
      }));
    } catch { return this.getMockRedemptions(); }
  }

  public async processRedemption(id: number, action: 'approve' | 'reject' | 'fulfill'): Promise<void> {
    const statusMap = { approve: 'Approved', reject: 'Rejected', fulfill: 'Fulfilled' };
    await this.sp.web.lists.getByTitle("PM_RewardRedemptions").items.getById(id).update({
      RedemptionStatus: statusMap[action], ProcessedDate: new Date().toISOString(), ProcessedBy: this.context.pageContext.user.displayName
    });
  }

  public async getTiers(): Promise<ITierConfig[]> {
    try {
      const items = await this.sp.web.lists.getByTitle("PM_LoyaltyTiers").items.orderBy("MinimumPoints", true)();
      return items.map((item: any) => ({
        id: item.Id, name: item.TierName || item.Title || '', displayName: item.TierDisplayName || item.TierName || item.Title || '',
        minPoints: item.MinimumPoints || 0, multiplier: item.PointsMultiplier || 1, discount: item.RewardDiscount || 0, color: item.TierColor || '#808080', icon: item.TierIcon || 'FavoriteStarFill'
      }));
    } catch { return this.getMockTiers(); }
  }

  public async updateTier(tier: ITierConfig): Promise<void> {
    await this.sp.web.lists.getByTitle("PM_LoyaltyTiers").items.getById(tier.id).update({
      TierName: tier.name, TierDisplayName: tier.displayName, MinimumPoints: tier.minPoints, PointsMultiplier: tier.multiplier, RewardDiscount: tier.discount, TierColor: tier.color, TierIcon: tier.icon
    });
  }

  public async getPointRules(): Promise<IPointRule[]> {
    try {
      const items = await this.sp.web.lists.getByTitle("PM_PointRules").items.orderBy("RuleCategory", true)();
      return items.map((item: any) => ({
        id: item.Id, code: item.RuleCode || '', name: item.RuleName || item.Title || '', description: item.RuleDescription || '',
        points: item.PointsAwarded || 0, isActive: item.IsActive !== false, category: item.RuleCategory || 'General'
      }));
    } catch { return this.getMockPointRules(); }
  }

  public async updatePointRule(rule: IPointRule): Promise<void> {
    await this.sp.web.lists.getByTitle("PM_PointRules").items.getById(rule.id).update({ RuleName: rule.name, RuleDescription: rule.description, PointsAwarded: rule.points, IsActive: rule.isActive });
  }

  private getMockRewards(): IAdminReward[] {
    return [
      { id: 1, code: 'COFFEE_50', name: 'Coffee Voucher', description: 'R50 voucher for the cafeteria', category: 'Food & Beverage', pointsCost: 500, icon: 'CoffeeScript', stockLevel: 50, isAvailable: true, isFeatured: true },
      { id: 2, code: 'MOVIE_2X', name: 'Movie Tickets (x2)', description: 'Two cinema tickets', category: 'Entertainment', pointsCost: 1500, icon: 'Video', stockLevel: 20, isAvailable: true, isFeatured: true },
      { id: 3, code: 'WOOLWORTHS_100', name: 'Woolworths Gift Card', description: 'R100 shopping voucher', category: 'Shopping', pointsCost: 1000, icon: 'ShoppingCart', stockLevel: 30, isAvailable: true, isFeatured: false },
      { id: 4, code: 'HALF_DAY', name: 'Half Day Leave', description: 'Extra half day of paid leave', category: 'Time Off', pointsCost: 5000, icon: 'Vacation', stockLevel: 10, isAvailable: true, isFeatured: true }
    ];
  }

  private getMockChallenges(): IAdminChallenge[] {
    return [
      { id: 1, code: 'ONBOARD_SPRINT', name: 'Onboarding Sprint', description: 'Complete all onboarding tasks', type: 'Individual', category: 'Task Completion', status: 'Active', startDate: '2024-01-01', endDate: '2024-12-31', goalTarget: 100, rewardPoints: 500, participants: 45 },
      { id: 2, code: 'RECOGNITION_WEEK', name: 'Recognition Week', description: 'Give recognition to 5 colleagues', type: 'Individual', category: 'Recognition', status: 'Active', startDate: '2024-02-01', endDate: '2024-02-28', goalTarget: 5, rewardPoints: 200, participants: 120 }
    ];
  }

  private getMockAchievements(): IAdminAchievement[] {
    return [
      { id: 1, code: 'FIRST_LOGIN', name: 'Welcome Aboard', description: 'Logged in for the first time', category: 'Onboarding', rarity: 'Common', points: 50, icon: 'Home', requirementType: 'Count', requirementValue: 1, isActive: true },
      { id: 2, code: 'TASK_MASTER_10', name: 'Task Master', description: 'Completed 10 tasks', category: 'Tasks', rarity: 'Uncommon', points: 100, icon: 'CheckMark', requirementType: 'Count', requirementValue: 10, isActive: true },
      { id: 3, code: 'STREAK_7', name: 'Week Warrior', description: 'Maintained a 7-day streak', category: 'Streaks', rarity: 'Rare', points: 150, icon: 'Calories', requirementType: 'Streak', requirementValue: 7, isActive: true }
    ];
  }

  private getMockRedemptions(): IRedemptionRequest[] {
    return [
      { id: 1, rewardId: 1, rewardName: 'Coffee Voucher', rewardCode: 'COFFEE_50', userId: 5, userName: 'John Doe', userEmail: 'john@company.com', pointsCost: 500, requestDate: '2024-02-20', status: 'Pending' },
      { id: 2, rewardId: 2, rewardName: 'Movie Tickets', rewardCode: 'MOVIE_2X', userId: 8, userName: 'Jane Smith', userEmail: 'jane@company.com', pointsCost: 1500, requestDate: '2024-02-19', status: 'Approved' }
    ];
  }

  private getMockTiers(): ITierConfig[] {
    return [
      { id: 1, name: 'Bronze', displayName: 'Bronze Member', minPoints: 0, multiplier: 1.0, discount: 0, color: '#CD7F32', icon: 'FavoriteStarFill' },
      { id: 2, name: 'Silver', displayName: 'Silver Member', minPoints: 2500, multiplier: 1.25, discount: 5, color: '#C0C0C0', icon: 'FavoriteStarFill' },
      { id: 3, name: 'Gold', displayName: 'Gold Member', minPoints: 10000, multiplier: 1.5, discount: 10, color: '#FFD700', icon: 'FavoriteStarFill' },
      { id: 4, name: 'Platinum', displayName: 'Platinum Elite', minPoints: 25000, multiplier: 2.0, discount: 15, color: '#E5E4E2', icon: 'CrownSolid' }
    ];
  }

  private getMockPointRules(): IPointRule[] {
    return [
      { id: 1, code: 'TASK_COMPLETE', name: 'Task Completed', description: 'Points for completing a task', points: 25, isActive: true, category: 'Tasks' },
      { id: 2, code: 'RECOGNITION_GIVEN', name: 'Recognition Given', description: 'Points for giving recognition', points: 15, isActive: true, category: 'Social' },
      { id: 3, code: 'DAILY_LOGIN', name: 'Daily Login', description: 'Points for daily login', points: 10, isActive: true, category: 'Engagement' },
      { id: 4, code: 'CHALLENGE_COMPLETE', name: 'Challenge Completed', description: 'Points for completing challenges', points: 100, isActive: true, category: 'Challenges' }
    ];
  }
}
