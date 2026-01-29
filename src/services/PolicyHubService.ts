// @ts-nocheck
// Policy Hub Service
// Advanced policy document center with rich metadata, filtering, and read timeframe tracking

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/search';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import {
  IPolicy,
  IPolicyDocumentMetadata,
  IPolicyHubFilter,
  IPolicyHubSortOptions,
  IPolicyHubSearchResult,
  IPolicyHubFacets,
  IPolicyHubDashboard,
  IPolicyDocumentSearchRequest,
  IReadTimeframeCompliance,
  IReadTimeframeMetric,
  ReadTimeframe,
  PolicyStatus,
  AcknowledgementStatus,
  IPolicyAcknowledgement
} from '../models/IPolicy';
import { logger } from './LoggingService';
import { PolicyLists, PolicyWorkflowLists } from '../constants/SharePointListNames';

export class PolicyHubService {
  private sp: SPFI;
  private readonly POLICIES_LIST = PolicyLists.POLICIES;
  private readonly POLICY_DOCUMENTS_LIST = PolicyLists.POLICY_DOCUMENTS;
  private readonly POLICY_ACKNOWLEDGEMENTS_LIST = PolicyLists.POLICY_ACKNOWLEDGEMENTS;
  private currentUserId: number = 0;

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
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to initialize:', error);
      throw error;
    }
  }

  // ============================================================================
  // POLICY HUB SEARCH & FILTERING
  // ============================================================================

  /**
   * Search policies and documents with advanced filtering
   */
  public async searchPolicyHub(request: IPolicyDocumentSearchRequest): Promise<IPolicyHubSearchResult> {
    try {
      // Build base query
      let policyQuery = this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.select('*');

      // Apply filters
      const filterConditions = this.buildFilterConditions(request.filters);
      if (filterConditions.length > 0) {
        policyQuery = policyQuery.filter(filterConditions.join(' and '));
      }

      // Apply sorting
      if (request.sort) {
        policyQuery = this.applySorting(policyQuery, request.sort);
      }

      // Get total count before pagination
      const allPolicies = await policyQuery.top(5000)();
      const totalCount = allPolicies.length;

      // Apply pagination
      const startIndex = (request.page - 1) * request.pageSize;
      const paginatedPolicies = allPolicies.slice(startIndex, startIndex + request.pageSize);

      // Search text filtering (client-side for now)
      let filteredPolicies = paginatedPolicies;
      if (request.searchText) {
        filteredPolicies = this.applyTextSearch(paginatedPolicies, request.searchText, request.filters?.searchFields);
      }

      // Get documents if requested
      let documents: IPolicyDocumentMetadata[] = [];
      if (request.includeDocuments) {
        documents = await this.getRelatedDocuments(filteredPolicies.map(p => p.Id));
      }

      // Generate facets if requested
      let facets: IPolicyHubFacets = {
        categories: [],
        departments: [],
        complianceRisks: [],
        statuses: [],
        documentTypes: [],
        tags: [],
        readTimeframes: []
      };

      if (request.includeFacets) {
        facets = this.generateFacets(allPolicies, documents);
      }

      return {
        policies: filteredPolicies as IPolicy[],
        documents,
        totalCount,
        filteredCount: filteredPolicies.length,
        facets
      };
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to search policy hub:', error);
      throw error;
    }
  }

  /**
   * Build filter conditions from filter object
   */
  private buildFilterConditions(filters?: IPolicyHubFilter): string[] {
    const conditions: string[] = [];

    if (!filters) return conditions;

    // Status filters
    if (filters.statuses && filters.statuses.length > 0) {
      const statusFilter = filters.statuses.map(s => `PolicyStatus eq '${s}'`).join(' or ');
      conditions.push(`(${statusFilter})`);
    }

    if (filters.isActive !== undefined) {
      conditions.push(`IsActive eq ${filters.isActive}`);
    }

    if (filters.isMandatory !== undefined) {
      conditions.push(`IsMandatory eq ${filters.isMandatory}`);
    }

    if (filters.isFeatured !== undefined) {
      conditions.push(`IsFeatured eq ${filters.isFeatured}`);
    }

    // Category filters
    if (filters.policyCategories && filters.policyCategories.length > 0) {
      const catFilter = filters.policyCategories.map(c => `PolicyCategory eq '${c}'`).join(' or ');
      conditions.push(`(${catFilter})`);
    }

    // Compliance risk filters
    if (filters.complianceRisks && filters.complianceRisks.length > 0) {
      const riskFilter = filters.complianceRisks.map(r => `ComplianceRisk eq '${r}'`).join(' or ');
      conditions.push(`(${riskFilter})`);
    }

    // Date filters
    if (filters.effectiveDateFrom) {
      conditions.push(`EffectiveDate ge datetime'${filters.effectiveDateFrom.toISOString()}'`);
    }
    if (filters.effectiveDateTo) {
      conditions.push(`EffectiveDate le datetime'${filters.effectiveDateTo.toISOString()}'`);
    }
    if (filters.publishedDateFrom) {
      conditions.push(`PublishedDate ge datetime'${filters.publishedDateFrom.toISOString()}'`);
    }
    if (filters.publishedDateTo) {
      conditions.push(`PublishedDate le datetime'${filters.publishedDateTo.toISOString()}'`);
    }

    // Acknowledgement filters
    if (filters.requiresAcknowledgement !== undefined) {
      conditions.push(`RequiresAcknowledgement eq ${filters.requiresAcknowledgement}`);
    }
    if (filters.requiresQuiz !== undefined) {
      conditions.push(`RequiresQuiz eq ${filters.requiresQuiz}`);
    }

    // Read timeframe filters
    if (filters.readTimeframes && filters.readTimeframes.length > 0) {
      const timeframeFilter = filters.readTimeframes.map(t => `ReadTimeframe eq '${t}'`).join(' or ');
      conditions.push(`(${timeframeFilter})`);
    }

    return conditions;
  }

  /**
   * Apply sorting to query
   */
  private applySorting(query: any, sort: IPolicyHubSortOptions): any {
    const fieldMap: Record<string, string> = {
      title: 'Title',
      policyNumber: 'PolicyNumber',
      effectiveDate: 'EffectiveDate',
      publishedDate: 'PublishedDate',
      category: 'PolicyCategory',
      complianceRisk: 'ComplianceRisk',
      viewCount: 'TotalDistributed'
    };

    const spField = fieldMap[sort.field] || 'Title';
    return query.orderBy(spField, sort.direction === 'desc');
  }

  /**
   * Apply text search (client-side)
   */
  private applyTextSearch(policies: any[], searchText: string, searchFields?: string[]): any[] {
    const lowerSearch = searchText.toLowerCase();
    const fields = searchFields || ['title', 'description', 'keywords'];

    return policies.filter(policy => {
      if (fields.includes('title') && policy.Title?.toLowerCase().includes(lowerSearch)) {
        return true;
      }
      if (fields.includes('description') && policy.Description?.toLowerCase().includes(lowerSearch)) {
        return true;
      }
      if (fields.includes('keywords')) {
        const keywords = policy.Keywords ? JSON.parse(policy.Keywords) : [];
        if (keywords.some((k: string) => k.toLowerCase().includes(lowerSearch))) {
          return true;
        }
      }
      return false;
    });
  }

  /**
   * Get related documents for policies
   */
  private async getRelatedDocuments(policyIds: number[]): Promise<IPolicyDocumentMetadata[]> {
    try {
      if (policyIds.length === 0) return [];

      const filter = policyIds.map(id => `PolicyId eq ${id}`).join(' or ');
      const documents = await this.sp.web.lists
        .getByTitle(this.POLICY_DOCUMENTS_LIST)
        .items.filter(`(${filter}) and IsActive eq true`)
        .top(1000)();

      return documents as IPolicyDocumentMetadata[];
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to get related documents:', error);
      return [];
    }
  }

  /**
   * Generate facets for filtering
   */
  private generateFacets(policies: any[], documents: IPolicyDocumentMetadata[]): IPolicyHubFacets {
    // Count occurrences for each facet
    const categoryCounts = new Map<string, number>();
    const departmentCounts = new Map<string, number>();
    const riskCounts = new Map<string, number>();
    const statusCounts = new Map<string, number>();
    const docTypeCounts = new Map<string, number>();
    const tagCounts = new Map<string, number>();
    const timeframeCounts = new Map<string, number>();

    // Process policies
    policies.forEach(policy => {
      // Categories
      if (policy.PolicyCategory) {
        categoryCounts.set(policy.PolicyCategory, (categoryCounts.get(policy.PolicyCategory) || 0) + 1);
      }

      // Compliance risks
      if (policy.ComplianceRisk) {
        riskCounts.set(policy.ComplianceRisk, (riskCounts.get(policy.ComplianceRisk) || 0) + 1);
      }

      // Status
      if (policy.PolicyStatus) {
        statusCounts.set(policy.PolicyStatus, (statusCounts.get(policy.PolicyStatus) || 0) + 1);
      }

      // Read timeframes
      if (policy.ReadTimeframe) {
        timeframeCounts.set(policy.ReadTimeframe, (timeframeCounts.get(policy.ReadTimeframe) || 0) + 1);
      }

      // Tags
      if (policy.Tags) {
        const tags = JSON.parse(policy.Tags);
        tags.forEach((tag: string) => {
          tagCounts.set(tag, (tagCounts.get(tag) || 0) + 1);
        });
      }
    });

    // Process documents
    documents.forEach(doc => {
      if (doc.DocumentType) {
        docTypeCounts.set(doc.DocumentType, (docTypeCounts.get(doc.DocumentType) || 0) + 1);
      }
    });

    // Convert to facet arrays
    return {
      categories: Array.from(categoryCounts.entries()).map(([name, count]) => ({ name, count })),
      departments: Array.from(departmentCounts.entries()).map(([name, count]) => ({ name, count })),
      complianceRisks: Array.from(riskCounts.entries()).map(([name, count]) => ({ name, count })),
      statuses: Array.from(statusCounts.entries()).map(([name, count]) => ({ name, count })),
      documentTypes: Array.from(docTypeCounts.entries()).map(([name, count]) => ({ name, count })),
      tags: Array.from(tagCounts.entries()).map(([name, count]) => ({ name, count })).slice(0, 20), // Top 20 tags
      readTimeframes: Array.from(timeframeCounts.entries()).map(([name, count]) => ({ name, count }))
    };
  }

  // ============================================================================
  // READ TIMEFRAME COMPLIANCE TRACKING
  // ============================================================================

  /**
   * Get read timeframe compliance for a policy
   */
  public async getReadTimeframeCompliance(policyId: number): Promise<IReadTimeframeCompliance> {
    try {
      // Get policy
      const policy = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(policyId)() as IPolicy;

      if (!policy.ReadTimeframe) {
        throw new Error('Policy does not have read timeframe configured');
      }

      // Calculate days for timeframe
      const timeframeDays = this.calculateTimeframeDays(policy.ReadTimeframe, policy.ReadTimeframeDays);

      // Get all acknowledgements for this policy
      const acknowledgements = await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.filter(`PolicyId eq ${policyId}`)
        .top(5000)() as IPolicyAcknowledgement[];

      // Calculate metrics
      const metrics = this.calculateReadTimeframeMetrics(acknowledgements, timeframeDays);

      // Calculate breakdown by department/role
      const byDepartment = this.calculateMetricsByDimension(acknowledgements, timeframeDays, 'UserDepartment');
      const byRole = this.calculateMetricsByDimension(acknowledgements, timeframeDays, 'UserRole');
      const byLocation = this.calculateMetricsByDimension(acknowledgements, timeframeDays, 'UserLocation');

      return {
        policyId,
        policyName: policy.PolicyName,
        readTimeframe: policy.ReadTimeframe,
        readTimeframeDays: timeframeDays,
        ...metrics,
        byDepartment,
        byRole,
        byLocation
      };
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to get read timeframe compliance:', error);
      throw error;
    }
  }

  /**
   * Calculate days for a read timeframe
   */
  private calculateTimeframeDays(timeframe: ReadTimeframe, customDays?: number): number {
    const timeframeMap: Record<ReadTimeframe, number> = {
      [ReadTimeframe.Immediate]: 0,
      [ReadTimeframe.Day1]: 1,
      [ReadTimeframe.Day3]: 3,
      [ReadTimeframe.Week1]: 7,
      [ReadTimeframe.Week2]: 14,
      [ReadTimeframe.Month1]: 30,
      [ReadTimeframe.Month3]: 90,
      [ReadTimeframe.Month6]: 180,
      [ReadTimeframe.Custom]: customDays || 30
    };

    return timeframeMap[timeframe] || 30;
  }

  /**
   * Calculate read timeframe metrics from acknowledgements
   */
  private calculateReadTimeframeMetrics(
    acknowledgements: IPolicyAcknowledgement[],
    timeframeDays: number
  ): Omit<IReadTimeframeCompliance, 'policyId' | 'policyName' | 'readTimeframe' | 'readTimeframeDays' | 'byDepartment' | 'byRole' | 'byLocation'> {
    const now = new Date();
    let readOnTime = 0;
    let readLate = 0;
    let notYetRead = 0;
    let overdue = 0;

    const readTimes: number[] = [];

    acknowledgements.forEach(ack => {
      const assignedDate = new Date(ack.AssignedDate);
      const deadlineDate = new Date(assignedDate.getTime() + timeframeDays * 24 * 60 * 60 * 1000);

      if (ack.AckStatus === AcknowledgementStatus.Acknowledged && ack.AcknowledgedDate) {
        const acknowledgedDate = new Date(ack.AcknowledgedDate);
        const daysToRead = (acknowledgedDate.getTime() - assignedDate.getTime()) / (1000 * 60 * 60 * 24);
        readTimes.push(daysToRead);

        if (acknowledgedDate <= deadlineDate) {
          readOnTime++;
        } else {
          readLate++;
        }
      } else {
        // Not yet read
        if (now > deadlineDate) {
          overdue++;
        } else {
          notYetRead++;
        }
      }
    });

    const totalAssigned = acknowledgements.length;
    const onTimePercentage = totalAssigned > 0 ? (readOnTime / totalAssigned) * 100 : 100;
    const latePercentage = totalAssigned > 0 ? (readLate / totalAssigned) * 100 : 0;
    const complianceRate = totalAssigned > 0 ? ((readOnTime + readLate) / totalAssigned) * 100 : 100;

    // Calculate time statistics
    const averageTimeToRead = readTimes.length > 0 ? readTimes.reduce((sum, time) => sum + time, 0) / readTimes.length : 0;
    const sortedTimes = [...readTimes].sort((a, b) => a - b);
    const medianTimeToRead = sortedTimes.length > 0 ? sortedTimes[Math.floor(sortedTimes.length / 2)] : 0;
    const fastestRead = sortedTimes.length > 0 ? sortedTimes[0] : 0;
    const slowestRead = sortedTimes.length > 0 ? sortedTimes[sortedTimes.length - 1] : 0;

    return {
      totalAssigned,
      readOnTime,
      readLate,
      notYetRead,
      overdue,
      onTimePercentage,
      latePercentage,
      complianceRate,
      averageTimeToRead,
      medianTimeToRead,
      fastestRead,
      slowestRead
    };
  }

  /**
   * Calculate metrics by dimension (department, role, location)
   */
  private calculateMetricsByDimension(
    acknowledgements: IPolicyAcknowledgement[],
    timeframeDays: number,
    dimension: 'UserDepartment' | 'UserRole' | 'UserLocation'
  ): Map<string, IReadTimeframeMetric> {
    const metricsByDimension = new Map<string, IPolicyAcknowledgement[]>();

    // Group acknowledgements by dimension
    acknowledgements.forEach(ack => {
      const dimValue = (ack as any)[dimension] || 'Unknown';
      if (!metricsByDimension.has(dimValue)) {
        metricsByDimension.set(dimValue, []);
      }
      metricsByDimension.get(dimValue)!.push(ack);
    });

    // Calculate metrics for each group
    const result = new Map<string, IReadTimeframeMetric>();
    metricsByDimension.forEach((acks, dimValue) => {
      const now = new Date();
      let readOnTime = 0;
      let readLate = 0;
      let notYetRead = 0;
      let overdue = 0;

      acks.forEach(ack => {
        const assignedDate = new Date(ack.AssignedDate);
        const deadlineDate = new Date(assignedDate.getTime() + timeframeDays * 24 * 60 * 60 * 1000);

        if (ack.AckStatus === AcknowledgementStatus.Acknowledged && ack.AcknowledgedDate) {
          const acknowledgedDate = new Date(ack.AcknowledgedDate);
          if (acknowledgedDate <= deadlineDate) {
            readOnTime++;
          } else {
            readLate++;
          }
        } else {
          if (now > deadlineDate) {
            overdue++;
          } else {
            notYetRead++;
          }
        }
      });

      const assigned = acks.length;
      const complianceRate = assigned > 0 ? ((readOnTime + readLate) / assigned) * 100 : 100;

      result.set(dimValue, {
        assigned,
        readOnTime,
        readLate,
        notYetRead,
        overdue,
        complianceRate
      });
    });

    return result;
  }

  /**
   * Get overall read timeframe compliance across all policies
   */
  public async getOverallReadTimeframeCompliance(): Promise<IReadTimeframeCompliance[]> {
    try {
      // Get all policies with read timeframes
      const policies = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.filter('ReadTimeframe ne null and IsActive eq true')
        .top(1000)() as IPolicy[];

      const complianceReports = await Promise.all(
        policies.map(policy => this.getReadTimeframeCompliance(policy.Id!))
      );

      return complianceReports;
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to get overall read timeframe compliance:', error);
      throw error;
    }
  }

  // ============================================================================
  // POLICY HUB DASHBOARD
  // ============================================================================

  /**
   * Get policy hub dashboard data
   */
  public async getPolicyHubDashboard(userId?: number): Promise<IPolicyHubDashboard> {
    try {
      const targetUserId = userId || this.currentUserId;

      // Get all active policies
      const allPolicies = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.filter('IsActive eq true')
        .top(5000)() as IPolicy[];

      const activePolicies = allPolicies.filter(p => p.PolicyStatus === PolicyStatus.Published);

      // Group by category
      const categoryCounts = new Map<string, number>();
      allPolicies.forEach(p => {
        if (p.PolicyCategory) {
          categoryCounts.set(p.PolicyCategory, (categoryCounts.get(p.PolicyCategory) || 0) + 1);
        }
      });

      // Group by compliance risk
      const riskCounts = new Map<string, number>();
      allPolicies.forEach(p => {
        if (p.ComplianceRisk) {
          riskCounts.set(p.ComplianceRisk, (riskCounts.get(p.ComplianceRisk) || 0) + 1);
        }
      });

      // Get featured and popular policies
      const featuredPolicies = allPolicies.filter((p: any) => p.IsFeatured).slice(0, 5);
      const mostViewedPolicies = [...allPolicies].sort((a, b) => (b.TotalDistributed || 0) - (a.TotalDistributed || 0)).slice(0, 5);

      // Get recently published (last 30 days)
      const thirtyDaysAgo = new Date();
      thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
      const recentlyPublished = allPolicies
        .filter(p => p.PublishedDate && new Date(p.PublishedDate) >= thirtyDaysAgo)
        .slice(0, 5);

      // Get recently updated
      const recentlyUpdated = [...allPolicies]
        .sort((a, b) => {
          const dateA = a.Modified ? new Date(a.Modified).getTime() : 0;
          const dateB = b.Modified ? new Date(b.Modified).getTime() : 0;
          return dateB - dateA;
        })
        .slice(0, 5);

      // Get read timeframe compliance
      const complianceByTimeframe = await this.getComplianceByTimeframe();

      // Calculate overall read timeframe compliance
      const overallCompliance = complianceByTimeframe.length > 0
        ? complianceByTimeframe.reduce((sum, item) => sum + item.complianceRate, 0) / complianceByTimeframe.length
        : 100;

      // Get critical policies that are overdue
      const criticalPoliciesOverdue = await this.getCriticalOverduePolicies();

      // Get user-specific data
      let myPendingPolicies: IPolicy[] | undefined;
      let myOverduePolicies: IPolicy[] | undefined;
      if (targetUserId) {
        const userAcks = await this.sp.web.lists
          .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
          .items.filter(`AckUserId eq ${targetUserId} and AckStatus ne '${AcknowledgementStatus.Acknowledged}'`)
          .top(100)() as IPolicyAcknowledgement[];

        const pendingPolicyIds = userAcks.filter(a => a.AckStatus === AcknowledgementStatus.Sent || a.AckStatus === AcknowledgementStatus.Opened).map(a => a.PolicyId);
        const overduePolicyIds = userAcks.filter(a => a.AckStatus === AcknowledgementStatus.Overdue).map(a => a.PolicyId);

        myPendingPolicies = allPolicies.filter(p => pendingPolicyIds.includes(p.Id!));
        myOverduePolicies = allPolicies.filter(p => overduePolicyIds.includes(p.Id!));
      }

      return {
        totalPolicies: allPolicies.length,
        activePolicies: activePolicies.length,
        policiesByCategory: Array.from(categoryCounts.entries()).map(([category, count]) => ({ category, count })),
        policiesByComplianceRisk: Array.from(riskCounts.entries()).map(([risk, count]) => ({ risk, count })),
        featuredPolicies,
        mostViewedPolicies,
        recentlyPublished,
        recentlyUpdated,
        complianceByTimeframe,
        overallReadTimeframeCompliance: overallCompliance,
        criticalPoliciesOverdue,
        myPendingPolicies,
        myOverduePolicies
      };
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to get policy hub dashboard:', error);
      throw error;
    }
  }

  /**
   * Get compliance by timeframe
   */
  private async getComplianceByTimeframe(): Promise<{ timeframe: string; complianceRate: number }[]> {
    try {
      const complianceReports = await this.getOverallReadTimeframeCompliance();

      const timeframeMap = new Map<string, { total: number; compliant: number }>();

      complianceReports.forEach(report => {
        const tf = report.readTimeframe;
        if (!timeframeMap.has(tf)) {
          timeframeMap.set(tf, { total: 0, compliant: 0 });
        }
        const data = timeframeMap.get(tf)!;
        data.total += report.totalAssigned;
        data.compliant += report.readOnTime;
      });

      return Array.from(timeframeMap.entries()).map(([timeframe, data]) => ({
        timeframe,
        complianceRate: data.total > 0 ? (data.compliant / data.total) * 100 : 100
      }));
    } catch (error) {
      return [];
    }
  }

  /**
   * Get critical policies that are overdue
   */
  private async getCriticalOverduePolicies(): Promise<IPolicy[]> {
    try {
      const criticalPolicies = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.filter(`ComplianceRisk eq 'Critical' and IsActive eq true`)
        .top(100)() as IPolicy[];

      const overduePolicies: IPolicy[] = [];

      for (const policy of criticalPolicies) {
        const acks = await this.sp.web.lists
          .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
          .items.filter(`PolicyId eq ${policy.Id} and AckStatus eq '${AcknowledgementStatus.Overdue}'`)
          .top(1)();

        if (acks.length > 0) {
          overduePolicies.push(policy);
        }
      }

      return overduePolicies;
    } catch (error) {
      return [];
    }
  }

  /**
   * Track document view
   */
  public async trackDocumentView(documentId: number): Promise<void> {
    try {
      const doc = await this.sp.web.lists
        .getByTitle(this.POLICY_DOCUMENTS_LIST)
        .items.getById(documentId)() as IPolicyDocumentMetadata;

      await this.sp.web.lists
        .getByTitle(this.POLICY_DOCUMENTS_LIST)
        .items.getById(documentId)
        .update({
          ViewCount: (doc.ViewCount || 0) + 1,
          LastViewedDate: new Date().toISOString()
        });
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to track document view:', error);
    }
  }

  /**
   * Track document download
   */
  public async trackDocumentDownload(documentId: number): Promise<void> {
    try {
      const doc = await this.sp.web.lists
        .getByTitle(this.POLICY_DOCUMENTS_LIST)
        .items.getById(documentId)() as IPolicyDocumentMetadata;

      await this.sp.web.lists
        .getByTitle(this.POLICY_DOCUMENTS_LIST)
        .items.getById(documentId)
        .update({
          DownloadCount: (doc.DownloadCount || 0) + 1
        });
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to track document download:', error);
    }
  }

  // ============================================================================
  // ROLE-BASED POLICY MANAGEMENT
  // ============================================================================

  /**
   * Get user policy dashboard (My Policies view)
   */
  public async getUserPolicyDashboard(userId: number): Promise<{
    pendingPolicies: IPolicy[];
    completedPolicies: IPolicy[];
    overduePolicies: IPolicy[];
  }> {
    try {
      // Get all acknowledgements for this user
      const acknowledgements = await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.filter(`AckUserId eq ${userId}`)
        .top(500)() as IPolicyAcknowledgement[];

      // Get all related policies
      const policyIds = Array.from(new Set(acknowledgements.map(a => a.PolicyId)));

      if (policyIds.length === 0) {
        return { pendingPolicies: [], completedPolicies: [], overduePolicies: [] };
      }

      const policyFilter = policyIds.map(id => `Id eq ${id}`).join(' or ');
      const policies = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.filter(`(${policyFilter})`)
        .top(500)() as IPolicy[];

      // Create policy lookup map
      const policyMap = new Map<number, IPolicy>();
      policies.forEach(p => policyMap.set(p.Id!, p));

      // Categorize based on acknowledgement status
      const pendingPolicies: IPolicy[] = [];
      const completedPolicies: IPolicy[] = [];
      const overduePolicies: IPolicy[] = [];

      acknowledgements.forEach(ack => {
        const policy = policyMap.get(ack.PolicyId);
        if (!policy) return;

        // Add deadline info to policy
        const policyWithDeadline = {
          ...policy,
          ReadDeadline: ack.DueDate
        };

        switch (ack.AckStatus) {
          case AcknowledgementStatus.Acknowledged:
            completedPolicies.push(policyWithDeadline);
            break;
          case AcknowledgementStatus.Overdue:
            overduePolicies.push(policyWithDeadline);
            break;
          default:
            pendingPolicies.push(policyWithDeadline);
            break;
        }
      });

      return { pendingPolicies, completedPolicies, overduePolicies };
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to get user policy dashboard:', error);
      return { pendingPolicies: [], completedPolicies: [], overduePolicies: [] };
    }
  }

  /**
   * Get policies authored by a specific user
   */
  public async getAuthoredPolicies(userId: number): Promise<IPolicy[]> {
    try {
      const policies = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.filter(`AuthorId eq ${userId}`)
        .orderBy('Modified', false)
        .top(500)() as IPolicy[];

      return policies;
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to get authored policies:', error);
      return [];
    }
  }

  /**
   * Get delegation requests for a manager
   */
  public async getDelegationRequests(managerId: number): Promise<any[]> {
    try {
      const requests = await this.sp.web.lists
        .getByTitle(PolicyWorkflowLists.DELEGATIONS)
        .items.filter(`RequestedById eq ${managerId}`)
        .orderBy('Created', false)
        .top(100)();

      return requests;
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to get delegation requests:', error);
      return [];
    }
  }

  /**
   * Get available authors for delegation
   */
  public async getAvailableAuthors(): Promise<Array<{ id: number; name: string; email: string }>> {
    try {
      // Get users in Policy Authors group
      const groups = await this.sp.web.siteGroups();
      const authorGroup = groups.find(g =>
        g.Title.toLowerCase().includes('policy author') ||
        g.Title.toLowerCase().includes('content author')
      );

      if (!authorGroup) {
        return [];
      }

      const users = await this.sp.web.siteGroups.getById(authorGroup.Id).users();

      return users.map(u => ({
        id: u.Id,
        name: u.Title,
        email: u.Email
      }));
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to get available authors:', error);
      return [];
    }
  }

  /**
   * Get policies pending approval
   */
  public async getPendingApprovals(approverId: number): Promise<IPolicy[]> {
    try {
      const policies = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.filter(`PolicyStatus eq 'Under Review'`)
        .orderBy('Modified', false)
        .top(100)() as IPolicy[];

      return policies;
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to get pending approvals:', error);
      return [];
    }
  }

  /**
   * Create a delegation request
   */
  public async createDelegationRequest(request: any): Promise<number> {
    try {
      // Get requestor name
      const currentUser = await this.sp.web.currentUser();

      const result = await this.sp.web.lists
        .getByTitle(PolicyWorkflowLists.DELEGATIONS)
        .items.add({
          Title: request.RequestTitle,
          RequestTitle: request.RequestTitle,
          RequestDescription: request.RequestDescription,
          PolicyCategory: request.PolicyCategory,
          PolicyTopic: request.PolicyTopic,
          Priority: request.Priority,
          RequestedById: request.RequestedById,
          RequestedByName: currentUser.Title,
          AssignedToId: request.AssignedToId,
          AssignedToName: request.AssignedToName,
          DueDate: request.DueDate?.toISOString(),
          Status: request.Status,
          Notes: request.Notes
        });

      return result.data.Id;
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to create delegation request:', error);
      throw error;
    }
  }

  /**
   * Approve a policy
   */
  public async approvePolicy(policyId: number, approverId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(policyId)
        .update({
          PolicyStatus: PolicyStatus.Published,
          ApprovedById: approverId,
          ApprovedDate: new Date().toISOString(),
          PublishedDate: new Date().toISOString()
        });
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to approve policy:', error);
      throw error;
    }
  }

  /**
   * Reject a policy
   */
  public async rejectPolicy(policyId: number, approverId: number, reason: string): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(policyId)
        .update({
          PolicyStatus: PolicyStatus.Draft,
          RejectedById: approverId,
          RejectedDate: new Date().toISOString(),
          RejectionReason: reason
        });
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to reject policy:', error);
      throw error;
    }
  }

  /**
   * Get policy analytics data
   */
  public async getPolicyAnalytics(): Promise<{
    totalPolicies: number;
    publishedPolicies: number;
    draftPolicies: number;
    expiringPolicies: number;
    overallComplianceRate: number;
    policiesByCategory: Array<{ category: string; count: number }>;
    recentAcknowledgements: IPolicyAcknowledgement[];
    complianceByDepartment: Array<{ department: string; rate: number }>;
  }> {
    try {
      // Get all policies
      const policies = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.top(5000)() as IPolicy[];

      // Calculate metrics
      const totalPolicies = policies.length;
      const publishedPolicies = policies.filter(p => p.PolicyStatus === PolicyStatus.Published).length;
      const draftPolicies = policies.filter(p => p.PolicyStatus === PolicyStatus.Draft).length;

      // Get expiring policies (within 30 days)
      const thirtyDaysFromNow = new Date();
      thirtyDaysFromNow.setDate(thirtyDaysFromNow.getDate() + 30);
      const expiringPolicies = policies.filter(p =>
        p.ExpiryDate && new Date(p.ExpiryDate) <= thirtyDaysFromNow
      ).length;

      // Get policies by category
      const categoryCounts = new Map<string, number>();
      policies.forEach(p => {
        if (p.PolicyCategory) {
          categoryCounts.set(p.PolicyCategory, (categoryCounts.get(p.PolicyCategory) || 0) + 1);
        }
      });
      const policiesByCategory = Array.from(categoryCounts.entries())
        .map(([category, count]) => ({ category, count }))
        .sort((a, b) => b.count - a.count);

      // Get recent acknowledgements
      const recentAcknowledgements = await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.filter(`AckStatus eq '${AcknowledgementStatus.Acknowledged}'`)
        .orderBy('AcknowledgedDate', false)
        .top(20)() as IPolicyAcknowledgement[];

      // Get all acknowledgements for compliance calculation
      const allAcknowledgements = await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.top(5000)() as IPolicyAcknowledgement[];

      // Calculate overall compliance rate
      const totalAcks = allAcknowledgements.length;
      const completedAcks = allAcknowledgements.filter(a => a.AckStatus === AcknowledgementStatus.Acknowledged).length;
      const overallComplianceRate = totalAcks > 0 ? (completedAcks / totalAcks) * 100 : 100;

      // Calculate compliance by department
      const deptMap = new Map<string, { total: number; completed: number }>();
      allAcknowledgements.forEach(ack => {
        const dept = (ack as any).UserDepartment || 'Unknown';
        if (!deptMap.has(dept)) {
          deptMap.set(dept, { total: 0, completed: 0 });
        }
        const data = deptMap.get(dept)!;
        data.total++;
        if (ack.AckStatus === AcknowledgementStatus.Acknowledged) {
          data.completed++;
        }
      });

      const complianceByDepartment = Array.from(deptMap.entries())
        .map(([department, data]) => ({
          department,
          rate: data.total > 0 ? (data.completed / data.total) * 100 : 100
        }))
        .sort((a, b) => b.rate - a.rate);

      return {
        totalPolicies,
        publishedPolicies,
        draftPolicies,
        expiringPolicies,
        overallComplianceRate,
        policiesByCategory,
        recentAcknowledgements,
        complianceByDepartment
      };
    } catch (error) {
      logger.error('PolicyHubService', 'Failed to get policy analytics:', error);
      throw error;
    }
  }
}
