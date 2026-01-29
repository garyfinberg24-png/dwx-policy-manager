// @ts-nocheck
/**
 * TermStoreService
 * Service for fetching terms from SharePoint Term Store (Managed Metadata Service)
 *
 * This service provides methods to retrieve taxonomy terms for use in
 * dropdowns and multi-select fields throughout the JML solution.
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/taxonomy';
import { ITermInfo, ITermSetInfo, ITermGroupInfo } from '@pnp/sp/taxonomy';

/**
 * Interface for a simple term option (used in dropdowns)
 */
export interface ITermOption {
  key: string;
  text: string;
  termId: string;
  labels?: string[];
  sortOrder?: number;
  isDeprecated?: boolean;
  childTerms?: ITermOption[];
}

/**
 * Interface for term set configuration
 */
export interface ITermSetConfig {
  groupName: string;
  termSetName: string;
}

/**
 * Default JML Term Store Configuration
 */
export const PM_TERM_GROUP = 'JML Managed Metadata';

export const CV_TERM_SETS: Record<string, ITermSetConfig> = {
  SKILLS: { groupName: PM_TERM_GROUP, termSetName: 'CV Skills' },
  DEPARTMENTS: { groupName: PM_TERM_GROUP, termSetName: 'CV Departments' },
  POSITIONS: { groupName: PM_TERM_GROUP, termSetName: 'CV Positions' },
  EXPERIENCE_LEVELS: { groupName: PM_TERM_GROUP, termSetName: 'CV Experience Levels' },
  EDUCATION_LEVELS: { groupName: PM_TERM_GROUP, termSetName: 'CV Education Levels' },
  SOURCES: { groupName: PM_TERM_GROUP, termSetName: 'CV Sources' }
};

/**
 * TermStoreService class for interacting with SharePoint Term Store
 */
export class TermStoreService {
  private sp: SPFI;
  private termCache: Map<string, ITermOption[]> = new Map();
  private cacheExpiry: Map<string, number> = new Map();
  private readonly CACHE_DURATION_MS = 5 * 60 * 1000; // 5 minutes

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Get the cache key for a term set
   */
  private getCacheKey(groupName: string, termSetName: string): string {
    return `${groupName}|${termSetName}`;
  }

  /**
   * Check if cached data is still valid
   */
  private isCacheValid(cacheKey: string): boolean {
    const expiry = this.cacheExpiry.get(cacheKey);
    if (!expiry) return false;
    return Date.now() < expiry;
  }

  /**
   * Get terms from a specific term set
   * @param groupName - The term group name (e.g., "JML Managed Metadata")
   * @param termSetName - The term set name (e.g., "CV Skills")
   * @param includeDeprecated - Whether to include deprecated terms (default: false)
   */
  public async getTerms(
    groupName: string,
    termSetName: string,
    includeDeprecated: boolean = false
  ): Promise<ITermOption[]> {
    const cacheKey = this.getCacheKey(groupName, termSetName);

    // Return cached data if valid
    if (this.isCacheValid(cacheKey)) {
      const cached = this.termCache.get(cacheKey);
      if (cached) {
        return includeDeprecated
          ? cached
          : cached.filter(t => !t.isDeprecated);
      }
    }

    try {
      // Get the term store
      const termStore = this.sp.termStore;

      // Get all term groups
      const groups: ITermGroupInfo[] = await termStore.groups();

      // Find the target group
      const targetGroup = groups.find(
        g => g.name.toLowerCase() === groupName.toLowerCase()
      );

      if (!targetGroup) {
        console.warn(`[TermStoreService] Term group not found: ${groupName}`);
        return [];
      }

      // Get term sets in the group
      const termSets: ITermSetInfo[] = await termStore.groups
        .getById(targetGroup.id)
        .sets();

      // Find the target term set
      const targetTermSet = termSets.find(
        ts => ts.localizedNames.some(ln => ln.name.toLowerCase() === termSetName.toLowerCase())
      );

      if (!targetTermSet) {
        console.warn(`[TermStoreService] Term set not found: ${termSetName}`);
        return [];
      }

      // Get all terms in the term set
      const terms: ITermInfo[] = await termStore.groups
        .getById(targetGroup.id)
        .sets.getById(targetTermSet.id)
        .getAllChildrenAsOrderedTree();

      // Convert to ITermOption format
      const termOptions = this.convertTermsToOptions(terms);

      // Cache the results
      this.termCache.set(cacheKey, termOptions);
      this.cacheExpiry.set(cacheKey, Date.now() + this.CACHE_DURATION_MS);

      return includeDeprecated
        ? termOptions
        : termOptions.filter(t => !t.isDeprecated);

    } catch (error) {
      console.error(`[TermStoreService] Error fetching terms for ${termSetName}:`, error);
      return [];
    }
  }

  /**
   * Convert PnP term info to ITermOption format
   */
  private convertTermsToOptions(terms: ITermInfo[]): ITermOption[] {
    const convertTerm = (term: ITermInfo): ITermOption => {
      const defaultLabel = term.labels?.find(l => l.isDefault)?.name || term.labels?.[0]?.name || '';

      return {
        key: defaultLabel,
        text: defaultLabel,
        termId: term.id,
        labels: term.labels?.map(l => l.name),
        isDeprecated: term.isDeprecated || false,
        childTerms: term.children?.map(convertTerm)
      };
    };

    return terms.map(convertTerm);
  }

  /**
   * Get CV Skills terms
   */
  public async getCVSkills(): Promise<ITermOption[]> {
    return this.getTerms(CV_TERM_SETS.SKILLS.groupName, CV_TERM_SETS.SKILLS.termSetName);
  }

  /**
   * Get CV Departments terms
   */
  public async getCVDepartments(): Promise<ITermOption[]> {
    return this.getTerms(CV_TERM_SETS.DEPARTMENTS.groupName, CV_TERM_SETS.DEPARTMENTS.termSetName);
  }

  /**
   * Get CV Positions terms
   */
  public async getCVPositions(): Promise<ITermOption[]> {
    return this.getTerms(CV_TERM_SETS.POSITIONS.groupName, CV_TERM_SETS.POSITIONS.termSetName);
  }

  /**
   * Get CV Experience Levels terms
   */
  public async getCVExperienceLevels(): Promise<ITermOption[]> {
    return this.getTerms(CV_TERM_SETS.EXPERIENCE_LEVELS.groupName, CV_TERM_SETS.EXPERIENCE_LEVELS.termSetName);
  }

  /**
   * Get CV Education Levels terms
   */
  public async getCVEducationLevels(): Promise<ITermOption[]> {
    return this.getTerms(CV_TERM_SETS.EDUCATION_LEVELS.groupName, CV_TERM_SETS.EDUCATION_LEVELS.termSetName);
  }

  /**
   * Get CV Sources terms
   */
  public async getCVSources(): Promise<ITermOption[]> {
    return this.getTerms(CV_TERM_SETS.SOURCES.groupName, CV_TERM_SETS.SOURCES.termSetName);
  }

  /**
   * Load all CV-related term sets at once
   * Useful for initial component load
   */
  public async loadAllCVTerms(): Promise<{
    skills: ITermOption[];
    departments: ITermOption[];
    positions: ITermOption[];
    experienceLevels: ITermOption[];
    educationLevels: ITermOption[];
    sources: ITermOption[];
  }> {
    const [
      skills,
      departments,
      positions,
      experienceLevels,
      educationLevels,
      sources
    ] = await Promise.all([
      this.getCVSkills(),
      this.getCVDepartments(),
      this.getCVPositions(),
      this.getCVExperienceLevels(),
      this.getCVEducationLevels(),
      this.getCVSources()
    ]);

    return {
      skills,
      departments,
      positions,
      experienceLevels,
      educationLevels,
      sources
    };
  }

  /**
   * Clear the term cache
   * Call this when terms might have been updated in the Term Store
   */
  public clearCache(): void {
    this.termCache.clear();
    this.cacheExpiry.clear();
  }

  /**
   * Clear cache for a specific term set
   */
  public clearTermSetCache(groupName: string, termSetName: string): void {
    const cacheKey = this.getCacheKey(groupName, termSetName);
    this.termCache.delete(cacheKey);
    this.cacheExpiry.delete(cacheKey);
  }
}

export default TermStoreService;
