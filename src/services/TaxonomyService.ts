// @ts-nocheck
/**
 * TaxonomyService
 *
 * Service for interacting with SharePoint Managed Metadata (Term Store).
 * Provides methods for retrieving terms, term sets, and managing tag assignments.
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/taxonomy';
import { ITermStore, ITermInfo, ITermSetInfo, ITermGroupInfo } from '@pnp/sp/taxonomy';
import {
  IJmlTerm,
  IJmlTermSet,
  IJmlTermGroup,
  IJmlTermStoreConfig,
  JmlTermSetType,
  IJmlTagSelection
} from '../models/ITaxonomy';

// Default JML Term Store configuration
// These IDs should be updated after running the provisioning script
const DEFAULT_CONFIG: IJmlTermStoreConfig = {
  termGroupName: 'JML Metadata',
  termGroupId: '', // Set after provisioning
  termSets: {
    [JmlTermSetType.BusinessDomains]: '',
    [JmlTermSetType.LifecycleStages]: '',
    [JmlTermSetType.ComplianceRegulatory]: '',
    [JmlTermSetType.PriorityRisk]: '',
    [JmlTermSetType.ContentClassification]: '',
    [JmlTermSetType.DocumentTypes]: '',
    [JmlTermSetType.Audience]: ''
  }
};

export class TaxonomyService {
  private sp: SPFI;
  private config: IJmlTermStoreConfig;
  private termCache: Map<string, IJmlTerm[]> = new Map();
  private cacheExpiry: Map<string, number> = new Map();
  private readonly CACHE_DURATION_MS = 5 * 60 * 1000; // 5 minutes

  constructor(sp: SPFI, config?: Partial<IJmlTermStoreConfig>) {
    this.sp = sp;
    this.config = { ...DEFAULT_CONFIG, ...config };
  }

  /**
   * Set or update the term store configuration
   */
  public setConfig(config: Partial<IJmlTermStoreConfig>): void {
    this.config = { ...this.config, ...config };
    this.clearCache();
  }

  /**
   * Get the current configuration
   */
  public getConfig(): IJmlTermStoreConfig {
    return { ...this.config };
  }

  /**
   * Clear the term cache
   */
  public clearCache(): void {
    this.termCache.clear();
    this.cacheExpiry.clear();
  }

  /**
   * Check if cache is valid for a given key
   */
  private isCacheValid(key: string): boolean {
    const expiry = this.cacheExpiry.get(key);
    return expiry !== undefined && Date.now() < expiry;
  }

  /**
   * Get the default term store
   */
  public async getTermStore(): Promise<ITermStore> {
    return this.sp.termStore;
  }

  /**
   * Get all term groups
   */
  public async getTermGroups(): Promise<IJmlTermGroup[]> {
    try {
      const groups = await this.sp.termStore.groups();
      return groups.map((g: ITermGroupInfo) => this.mapTermGroup(g));
    } catch (error) {
      console.error('Error fetching term groups:', error);
      throw new Error('Failed to retrieve term groups from Term Store');
    }
  }

  /**
   * Get the JML term group
   */
  public async getJmlTermGroup(): Promise<IJmlTermGroup | null> {
    try {
      if (this.config.termGroupId) {
        const group = await this.sp.termStore.groups.getById(this.config.termGroupId)();
        return this.mapTermGroup(group);
      } else if (this.config.termGroupName) {
        const groups = await this.sp.termStore.groups();
        const jmlGroup = groups.find((g: ITermGroupInfo) => g.name === this.config.termGroupName);
        if (jmlGroup) {
          // Update config with found ID
          this.config.termGroupId = jmlGroup.id;
          return this.mapTermGroup(jmlGroup);
        }
      }
      return null;
    } catch (error) {
      console.error('Error fetching JML term group:', error);
      return null;
    }
  }

  /**
   * Get all term sets in the JML term group
   */
  public async getTermSets(): Promise<IJmlTermSet[]> {
    try {
      const group = await this.getJmlTermGroup();
      if (!group) {
        throw new Error('JML term group not found');
      }

      const termSets = await this.sp.termStore.groups.getById(group.id).sets();
      return termSets.map((ts: ITermSetInfo) => this.mapTermSet(ts, group.id));
    } catch (error) {
      console.error('Error fetching term sets:', error);
      throw new Error('Failed to retrieve term sets');
    }
  }

  /**
   * Get a specific term set by type
   */
  public async getTermSetByType(termSetType: JmlTermSetType): Promise<IJmlTermSet | null> {
    try {
      const termSetId = this.config.termSets[termSetType];
      if (!termSetId) {
        // Try to find by name
        const termSets = await this.getTermSets();
        const termSetName = this.getTermSetNameByType(termSetType);
        const found = termSets.find(ts => ts.name === termSetName);
        if (found) {
          // Update config
          this.config.termSets[termSetType] = found.id;
          return found;
        }
        return null;
      }

      const termSet = await this.sp.termStore.sets.getById(termSetId)();
      return this.mapTermSet(termSet, this.config.termGroupId);
    } catch (error) {
      console.error(`Error fetching term set ${termSetType}:`, error);
      return null;
    }
  }

  /**
   * Get terms from a specific term set
   */
  public async getTermsBySetType(
    termSetType: JmlTermSetType,
    includeDeprecated: boolean = false
  ): Promise<IJmlTerm[]> {
    const cacheKey = `terms_${termSetType}_${includeDeprecated}`;

    // Check cache
    if (this.isCacheValid(cacheKey)) {
      return this.termCache.get(cacheKey) || [];
    }

    try {
      const termSet = await this.getTermSetByType(termSetType);
      if (!termSet) {
        return [];
      }

      const terms = await this.sp.termStore.sets.getById(termSet.id).terms();
      const mappedTerms = await this.mapTermsHierarchy(terms, termSet.id, includeDeprecated);

      // Cache results
      this.termCache.set(cacheKey, mappedTerms);
      this.cacheExpiry.set(cacheKey, Date.now() + this.CACHE_DURATION_MS);

      return mappedTerms;
    } catch (error) {
      console.error(`Error fetching terms for ${termSetType}:`, error);
      return [];
    }
  }

  /**
   * Get a flat list of all terms (including children) from a term set
   */
  public async getFlatTermsBySetType(
    termSetType: JmlTermSetType,
    includeDeprecated: boolean = false
  ): Promise<IJmlTerm[]> {
    const hierarchicalTerms = await this.getTermsBySetType(termSetType, includeDeprecated);
    return this.flattenTerms(hierarchicalTerms);
  }

  /**
   * Get a specific term by ID
   */
  public async getTermById(termId: string, termSetType: JmlTermSetType): Promise<IJmlTerm | null> {
    try {
      const terms = await this.getFlatTermsBySetType(termSetType);
      return terms.find(t => t.id === termId) || null;
    } catch (error) {
      console.error(`Error fetching term ${termId}:`, error);
      return null;
    }
  }

  /**
   * Search terms by name
   */
  public async searchTerms(
    searchText: string,
    termSetType?: JmlTermSetType,
    maxResults: number = 20
  ): Promise<IJmlTerm[]> {
    try {
      const searchLower = searchText.toLowerCase();
      let allTerms: IJmlTerm[] = [];

      if (termSetType) {
        allTerms = await this.getFlatTermsBySetType(termSetType);
      } else {
        // Search across all term sets
        for (const type of Object.values(JmlTermSetType)) {
          const terms = await this.getFlatTermsBySetType(type as JmlTermSetType);
          allTerms = allTerms.concat(terms);
        }
      }

      // Filter by search text
      const filtered = allTerms.filter(t =>
        t.name.toLowerCase().includes(searchLower) ||
        t.labels?.some(l => l.name.toLowerCase().includes(searchLower))
      );

      return filtered.slice(0, maxResults);
    } catch (error) {
      console.error('Error searching terms:', error);
      return [];
    }
  }

  /**
   * Convert terms to tag selections (for use in components)
   */
  public termsToTagSelections(
    terms: IJmlTerm[],
    termSetType: JmlTermSetType,
    termSetId: string
  ): IJmlTagSelection[] {
    return terms.map(term => ({
      termId: term.id,
      termName: term.name,
      termSetId: termSetId,
      termSetType: termSetType,
      path: term.path
    }));
  }

  /**
   * Get term set name by type
   */
  private getTermSetNameByType(termSetType: JmlTermSetType): string {
    const nameMap: Record<JmlTermSetType, string> = {
      [JmlTermSetType.BusinessDomains]: 'Business Domains',
      [JmlTermSetType.LifecycleStages]: 'JML Lifecycle Stages',
      [JmlTermSetType.ComplianceRegulatory]: 'Compliance & Regulatory',
      [JmlTermSetType.PriorityRisk]: 'Priority & Risk',
      [JmlTermSetType.ContentClassification]: 'Content Classification',
      [JmlTermSetType.DocumentTypes]: 'Document Types',
      [JmlTermSetType.Audience]: 'Audience'
    };
    return nameMap[termSetType];
  }

  /**
   * Map PnPjs term group to IJmlTermGroup
   */
  private mapTermGroup(group: ITermGroupInfo): IJmlTermGroup {
    return {
      id: group.id,
      name: group.name,
      description: group.description || undefined
    };
  }

  /**
   * Map PnPjs term set to IJmlTermSet
   */
  private mapTermSet(termSet: ITermSetInfo, groupId: string): IJmlTermSet {
    return {
      id: termSet.id,
      name: termSet.localizedNames?.[0]?.name || termSet.id,
      description: termSet.description || undefined,
      groupId: groupId,
      isOpenForTermCreation: (termSet as any).isOpenForTermCreation || false
    };
  }

  /**
   * Map PnPjs terms to IJmlTerm with hierarchy
   */
  private async mapTermsHierarchy(
    terms: ITermInfo[],
    termSetId: string,
    includeDeprecated: boolean
  ): Promise<IJmlTerm[]> {
    // Create a map for quick lookup
    const termMap = new Map<string, IJmlTerm>();
    const rootTerms: IJmlTerm[] = [];

    // First pass: create all term objects
    for (const term of terms) {
      if (!includeDeprecated && term.isDeprecated) {
        continue;
      }

      const mappedTerm: IJmlTerm = {
        id: term.id,
        name: term.labels?.[0]?.name || term.id,
        description: term.descriptions?.[0]?.description,
        termSetId: termSetId,
        isAvailableForTagging: (term as any).isAvailableForTagging !== false,
        isDeprecated: term.isDeprecated || false,
        labels: term.labels?.map(l => ({
          name: l.name,
          isDefault: l.isDefault,
          languageTag: l.languageTag
        })),
        children: [],
        level: 0
      };

      termMap.set(term.id, mappedTerm);
    }

    // Second pass: build hierarchy
    for (const term of terms) {
      if (!includeDeprecated && term.isDeprecated) {
        continue;
      }

      const mappedTerm = termMap.get(term.id);
      if (!mappedTerm) continue;

      // Check if this term has a parent
      const parentId = (term as any).parent?.id;
      if (parentId && termMap.has(parentId)) {
        const parent = termMap.get(parentId)!;
        parent.children = parent.children || [];
        parent.children.push(mappedTerm);
        mappedTerm.parentId = parentId;
        mappedTerm.level = (parent.level || 0) + 1;
      } else {
        rootTerms.push(mappedTerm);
      }
    }

    // Build paths
    this.buildTermPaths(rootTerms, '');

    // Sort by name
    this.sortTerms(rootTerms);

    return rootTerms;
  }

  /**
   * Build term paths recursively
   */
  private buildTermPaths(terms: IJmlTerm[], parentPath: string): void {
    for (const term of terms) {
      term.path = parentPath ? `${parentPath};${term.name}` : term.name;
      if (term.children && term.children.length > 0) {
        this.buildTermPaths(term.children, term.path);
      }
    }
  }

  /**
   * Sort terms alphabetically (recursively)
   */
  private sortTerms(terms: IJmlTerm[]): void {
    terms.sort((a, b) => a.name.localeCompare(b.name));
    for (const term of terms) {
      if (term.children && term.children.length > 0) {
        this.sortTerms(term.children);
      }
    }
  }

  /**
   * Flatten hierarchical terms into a flat list
   */
  private flattenTerms(terms: IJmlTerm[]): IJmlTerm[] {
    const flat: IJmlTerm[] = [];

    const addTerm = (term: IJmlTerm): void => {
      flat.push(term);
      if (term.children) {
        term.children.forEach(addTerm);
      }
    };

    terms.forEach(addTerm);
    return flat;
  }

  /**
   * Initialize the service by loading term set IDs from the term store
   * Call this once on application startup
   */
  public async initialize(): Promise<void> {
    try {
      const group = await this.getJmlTermGroup();
      if (!group) {
        console.warn('JML Metadata term group not found. Please run the provisioning script.');
        return;
      }

      const termSets = await this.sp.termStore.groups.getById(group.id).sets();

      // Map term sets to config
      for (const ts of termSets) {
        const name = ts.localizedNames?.[0]?.name || '';
        switch (name) {
          case 'Business Domains':
            this.config.termSets[JmlTermSetType.BusinessDomains] = ts.id;
            break;
          case 'JML Lifecycle Stages':
            this.config.termSets[JmlTermSetType.LifecycleStages] = ts.id;
            break;
          case 'Compliance & Regulatory':
            this.config.termSets[JmlTermSetType.ComplianceRegulatory] = ts.id;
            break;
          case 'Priority & Risk':
            this.config.termSets[JmlTermSetType.PriorityRisk] = ts.id;
            break;
          case 'Content Classification':
            this.config.termSets[JmlTermSetType.ContentClassification] = ts.id;
            break;
          case 'Document Types':
            this.config.termSets[JmlTermSetType.DocumentTypes] = ts.id;
            break;
          case 'Audience':
            this.config.termSets[JmlTermSetType.Audience] = ts.id;
            break;
        }
      }

      console.log('TaxonomyService initialized with config:', this.config);
    } catch (error) {
      console.error('Failed to initialize TaxonomyService:', error);
    }
  }
}

// Export a factory function for creating the service
export function createTaxonomyService(sp: SPFI, config?: Partial<IJmlTermStoreConfig>): TaxonomyService {
  return new TaxonomyService(sp, config);
}
